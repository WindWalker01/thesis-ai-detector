#!/usr/bin/env python3
"""
AI-text probability analysis for long documents (theses, papers, reports).

Model: followsci/bert-ai-text-detector  (BERT-base-uncased; label 0 = Human,
label 1 = AI-generated; transformer limit = 512 tokens per input).

Long documents are analyzed paragraph-by-paragraph: each paragraph is
tokenized with truncation=True / max_length=512, scored individually, and
aggregated into a document-level AI probability summary. Short or non-prose
blocks (titles, page numbers, brief citations, URLs) are filtered out before
scoring so they cannot skew the statistics.

Usage:
    python ai_detector.py thesis.docx
    python ai_detector.py thesis.pdf --threshold 70 --csv results.csv
    python ai_detector.py thesis.docx --report report.html --csv results.csv
    python ai_detector.py thesis.txt --batch-size 8 --top 15
"""

from __future__ import annotations

import argparse
import csv
import html
import re
import statistics
import sys
import time
from dataclasses import dataclass, field
from pathlib import Path

MODEL_NAME = "followsci/bert-ai-text-detector"
MAX_LENGTH = 512                      # transformer hard limit (tokens)
DEFAULT_THRESHOLD = 70.0              # "spike" threshold in percent
DEFAULT_MIN_WORDS = 8                 # blocks below this are filtered
SUPPORTED_SUFFIXES = {".docx", ".pdf", ".txt", ".md", ".text"}


# --------------------------------------------------------------------------- #
# Console-safe output: legacy Windows code pages (cp1252/OEM...) cannot encode
# every symbol, and stdout may be redirected to files. Never crash while printing.
# --------------------------------------------------------------------------- #
def _configure_streams() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")
        except (AttributeError, ValueError, OSError):
            pass


def _supports(char: str) -> bool:
    encoding = getattr(sys.stdout, "encoding", None) or "utf-8"
    try:
        char.encode(encoding)
    except (UnicodeEncodeError, LookupError):
        return False
    return True


WARN_SYMBOL = "⚠" if _supports("⚠") else "!!"
DASH = "—" if _supports("—") else "-"
QUOTE_OPEN = "«" if _supports("«") else '"'
QUOTE_CLOSE = "»" if _supports("»") else '"'
ELLIPSIS = "…" if _supports("…") else "..."

# Label-name maps used to locate the "AI" logit index in the model config.
AI_LABELS = {
    "ai", "ai text", "ai generated", "ai-generated", "ai generated text",
    "ai-generated text", "machine", "machine generated", "machine-generated",
    "generated", "llm", "artificial",
}
HUMAN_LABELS = {
    "human", "human text", "human written", "human-written", "original",
    "real", "student", "person", "organic",
}


# --------------------------------------------------------------------------- #
# Document readers (structure-preserving, error tolerant)
# --------------------------------------------------------------------------- #
def read_docx(path: Path, include_tables: bool = False) -> str:
    """Return the body of a .docx as raw text, paragraphs separated by blank lines."""
    try:
        import docx  # python-docx
    except ImportError:
        raise SystemExit(
            "Missing dependency 'python-docx'. Install with: pip install python-docx"
        )
    try:
        document = docx.Document(str(path))
    except Exception as exc:
        raise SystemExit(f"Could not open DOCX file '{path}': {exc}") from exc

    paragraphs = [p.text for p in document.paragraphs if p.text and p.text.strip()]
    if include_tables:
        for table in document.tables:
            for row in table.rows:
                cells = [c.text.strip() for c in row.cells if c.text and c.text.strip()]
                if cells:
                    paragraphs.append(" | ".join(cells))
    return "\n\n".join(paragraphs)


def read_pdf(path: Path) -> str:
    """Return page texts of a .pdf joined with single newlines (page-boundary aware)."""
    try:
        from pypdf import PdfReader
    except ImportError:
        try:
            from PyPDF2 import PdfReader  # type: ignore[no-redef]
        except ImportError:
            raise SystemExit(
                "Missing dependency 'pypdf' (or legacy 'PyPDF2'). "
                "Install with: pip install pypdf"
            )
    try:
        reader = PdfReader(str(path))
    except Exception as exc:
        raise SystemExit(f"Could not open PDF file '{path}': {exc}") from exc

    pages: list[str] = []
    for number, page in enumerate(reader.pages, start=1):
        try:
            pages.append(page.extract_text() or "")
        except Exception as exc:
            print(f"  ! Skipping unreadable PDF page {number}: {exc}", file=sys.stderr)
            pages.append("")
    return "\n".join(pages)


def read_plain_text(path: Path) -> str:
    """Decode a .txt/.md with a safe encoding fallback chain."""
    data = path.read_bytes()
    for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    raise SystemExit(
        f"Could not decode '{path}' with any of: utf-8, cp1252, latin-1."
    )


def read_document(path: Path, include_tables: bool = False) -> str:
    suffix = path.suffix.lower()
    if suffix not in SUPPORTED_SUFFIXES:
        raise SystemExit(
            f"Unsupported file type '{suffix}'. "
            f"Supported: {', '.join(sorted(SUPPORTED_SUFFIXES))}"
        )
    if not path.is_file():
        raise SystemExit(f"File not found: {path}")
    if suffix == ".docx":
        return read_docx(path, include_tables=include_tables)
    if suffix == ".pdf":
        return read_pdf(path)
    return read_plain_text(path)


# --------------------------------------------------------------------------- #
# Paragraph segmentation
# --------------------------------------------------------------------------- #
BLANK_LINE_RE = re.compile(r"\n\s*\n+")
HYPHEN_BREAK_RE = re.compile(r"(?<=[A-Za-z])-\r?\n(?=[a-z])")
TERMINAL_PUNCT_RE = re.compile(r"[.!?…:;\"'”’)\]]$")


def _starts_new_paragraph(previous_line: str, next_line: str) -> bool:
    """Heuristic: a physical line ends a paragraph only if it ends with terminal
    punctuation AND the next line starts a fresh sentence/heading. Lines that end
    mid-sentence are treated as hard wraps and joined back together."""
    previous = previous_line.strip()
    if not previous or not TERMINAL_PUNCT_RE.search(previous):
        return False
    first = next_line.lstrip()[:1]
    return first.isupper() or first.isdigit() or first in "\"'“(["


def split_paragraphs(raw: str) -> list[str]:
    """Turn raw document text into clean paragraphs.

    Handles: blank-line separated paragraphs, hard-wrapped text (PDF/plain text),
    line-end hyphenation, page-break/form-feed characters, and soft line breaks.
    """
    text = raw.replace("\x0c", "\n\n")          # form feed = page break
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = HYPHEN_BREAK_RE.sub("", text)         # "exam-\nple" -> "example"

    chunks = BLANK_LINE_RE.split(text)

    # Typical width of wrapped lines, measured only on lines that have a next
    # line inside the same chunk (real wrap candidates). Extracted PDF text has
    # no blank lines between blocks, so a line much shorter than the typical
    # wrap width (heading, page number, reference marker) ends the previous
    # block even when it lacks terminal punctuation.
    candidates = [
        len(stripped)
        for chunk in chunks
        for stripped in (ln.strip() for ln in chunk.split("\n")[:-1])
        if stripped
    ]
    short_line_limit = int(0.5 * statistics.median(candidates)) if candidates else -1

    paragraphs: list[str] = []
    for chunk in chunks:
        current: list[str] = []
        for line in chunk.split("\n"):
            line = re.sub(r"[ \t]+", " ", line).strip()
            if not line:
                continue
            if current and (
                _starts_new_paragraph(current[-1], line)
                or len(current[-1]) <= short_line_limit
            ):
                paragraphs.append(" ".join(current))
                current = []
            current.append(line)
        if current:
            paragraphs.append(" ".join(current))

    cleaned = [re.sub(r"\s+", " ", p).strip() for p in paragraphs]
    return [p for p in cleaned if p]


# --------------------------------------------------------------------------- #
# Filtering (protects the statistics from short / non-prose blocks)
# --------------------------------------------------------------------------- #
@dataclass
class Block:
    position: int                       # original paragraph number (1-based)
    text: str
    words: int
    part: int = 0                       # sub-part index when --split-long is used
    tokens: int = 0
    ai_pct: float = 0.0
    truncated: bool = False
    filtered_reason: str | None = None

    @property
    def label(self) -> str:
        return f"#{self.position}" if self.part == 0 else f"#{self.position}.{self.part}"


URL_RE = re.compile(r"^(?:https?://|www\.)\S+$", re.I)
YEAR_IN_PARENS_RE = re.compile(r"\((?:19|20)\d{2}[a-z]?\)")
AUTHOR_DATE_START_RE = re.compile(
    r"^(?:\[\d{1,3}\]|\d{1,2}\.|[IVXLivxl]+\))?\s*"          # optional ref marker
    r"[A-Z][\w'’\u00C0-\u00FF-]+,\s+[A-Z]\.?(?:\s+[A-Z]\.)?"  # Author, F. [M.]
    r"(?:(?:\s*,\s*|\s+)(?:&\s*)?[A-Z][\w'’\u00C0-\u00FF-]+,?\s+[A-Z]\.?(?:\s+[A-Z]\.)?)*"
    r"\s*[\(\[]?(?:19|20)\d{2}"                                # (2020)
)
REF_MARKER_RE = re.compile(
    r"https?://|doi\s*:|arxiv|retrieved\s+from|\bvol\.\s*\d|pp\.\s*\d|"
    r"\bisbn\b|\bissn\b|\bet\s+al\b|\bproceedings\b|\bjournal\b|"
    r"\buniversity\s+press\b",
    re.I,
)
JOURNAL_VOLUME_RE = re.compile(r"\b\d{1,4}\s*\(\d{1,3}\s*\d*\)")


def looks_like_reference(text: str) -> bool:
    """Heuristic detector for bibliography entries (APA/IEEE-ish)."""
    head = text[:200]
    if not YEAR_IN_PARENS_RE.search(head):
        return False
    if not (AUTHOR_DATE_START_RE.match(text)
            or re.match(r"^(?:\[\d{1,3}\]|\d{1,2}\.)\s+\S", text)):
        return False
    if REF_MARKER_RE.search(text) or JOURNAL_VOLUME_RE.search(text):
        return True
    return len(text.split()) <= 30          # brief citation lines


def filter_reason(text: str, min_words: int, keep_references: bool) -> str | None:
    """Return a reason string if the block should be excluded from scoring."""
    words = text.split()
    count = len(words)
    if URL_RE.match(text):
        return "url"
    if count < min_words:
        return "short"
    alpha_words = sum(1 for w in words if re.search(r"[A-Za-z\u00C0-\u024F]", w))
    if alpha_words < max(3, int(0.4 * count)):
        return "non-prose"
    letters = sum(1 for c in text if c.isalpha() or c.isspace())
    if letters / len(text) < 0.45:
        return "symbol-heavy"
    if not keep_references and looks_like_reference(text):
        return "reference"
    return None


# --------------------------------------------------------------------------- #
# Model loading & scoring
# --------------------------------------------------------------------------- #
def resolve_device(choice: str) -> str:
    import torch
    if choice == "auto":
        return "cuda" if torch.cuda.is_available() else "cpu"
    if choice == "cuda" and not torch.cuda.is_available():
        raise SystemExit(
            "CUDA requested (--device cuda) but no usable GPU was detected. "
            "Use --device cpu or --device auto instead."
        )
    return choice


def detect_ai_label_index(model) -> int:
    """Locate the logit index that corresponds to the AI class.

    The model card specifies label 0 = Human, label 1 = AI-generated; we verify
    that against id2label when available and fall back to index 1.
    """
    id2label = getattr(model.config, "id2label", None) or {}
    ai_index: int | None = None
    human_index: int | None = None
    for index, label in id2label.items():
        normalized = str(label).strip().lower().replace("_", " ")
        if normalized in AI_LABELS:
            ai_index = int(index)
        elif normalized in HUMAN_LABELS:
            human_index = int(index)
    if ai_index is not None and ai_index != human_index:
        return ai_index
    return 1  # model card: index 1 = AI-generated


def load_model(device: str, model_name: str = MODEL_NAME):
    try:
        import torch
        from transformers import AutoModelForSequenceClassification, AutoTokenizer
    except ImportError:
        raise SystemExit(
            "Missing dependencies 'torch' and/or 'transformers'. "
            "Install with: pip install torch transformers"
        )
    import logging
    logging.getLogger("transformers").setLevel(logging.ERROR)  # quiet >512-token note
    print(f"  loading {model_name} ...")
    try:
        tokenizer = AutoTokenizer.from_pretrained(model_name)
        model = AutoModelForSequenceClassification.from_pretrained(model_name)
    except Exception as exc:
        raise SystemExit(
            f"Could not load model '{model_name}': {exc}\n"
            "Check your internet connection (first run downloads ~440 MB, "
            "later runs use the Hugging Face cache), or point --model at a "
            "local folder containing the model files."
        ) from exc
    model.eval()
    model.to(device)
    return tokenizer, model, detect_ai_label_index(model)


def _snippet(text: str, limit: int = 110) -> str:
    flat = re.sub(r"\s+", " ", text).strip()
    if len(flat) <= limit:
        return flat
    return flat[: limit - len(ELLIPSIS)] + ELLIPSIS


def split_long_blocks(blocks: list[Block], tokenizer) -> list[Block]:
    """Optional (--split-long): window >512-token paragraphs on sentence
    boundaries instead of truncating them."""
    expanded: list[Block] = []
    for block in blocks:
        if block.tokens <= MAX_LENGTH:
            expanded.append(block)
            continue
        sentences = re.split(r"(?<=[.!?…])\s+", block.text)
        parts: list[str] = []
        current: list[str] = []
        for sentence in sentences:
            candidate = " ".join(current + [sentence])
            if current and len(tokenizer.encode(candidate, add_special_tokens=True)) > MAX_LENGTH - 2:
                parts.append(" ".join(current))
                current = [sentence]
            else:
                current.append(sentence)
        if current:
            parts.append(" ".join(current))
        for number, part in enumerate(parts, start=1):
            expanded.append(Block(
                position=block.position,
                text=part,
                words=len(part.split()),
                part=number,
            ))
    return expanded


def score_blocks(
    blocks: list[Block],
    tokenizer,
    model,
    device: str,
    ai_index: int,
    batch_size: int,
) -> None:
    """Score every block; strict truncation to MAX_LENGTH tokens per input."""
    import torch

    # Token accounting + truncation detection (strict max_length=512 below).
    for block in blocks:
        block.tokens = len(tokenizer.encode(block.text, add_special_tokens=True))
        block.truncated = block.tokens > MAX_LENGTH

    total = len(blocks)
    for start in range(0, total, batch_size):
        batch = blocks[start:start + batch_size]
        encoded = tokenizer(
            [b.text for b in batch],
            return_tensors="pt",
            truncation=True,           # <-- strict transformer limit
            max_length=MAX_LENGTH,
            padding=True,
        )
        encoded = {key: value.to(device) for key, value in encoded.items()}
        with torch.no_grad():
            logits = model(**encoded).logits
        probabilities = torch.softmax(logits, dim=-1)[:, ai_index].tolist()
        for block, probability in zip(batch, probabilities):
            block.ai_pct = float(probability) * 100.0

        done = start + len(batch)
        if done % 50 == 0 or done == total:
            sys.stdout.write(f"\r  scoring blocks: {done}/{total}")
            sys.stdout.flush()
    sys.stdout.write("\r" + " " * 40 + "\r")
    sys.stdout.flush()


# --------------------------------------------------------------------------- #
# Reporting
# --------------------------------------------------------------------------- #
@dataclass
class SummaryStats:
    """Document-level statistics shared by the console summary and HTML report."""
    mean: float
    weighted: float
    median: float
    p90: float
    over_50: int
    flagged: list[Block]
    truncated: list[Block]
    total_words: int
    reasons: dict[str, int]

    @property
    def reason_text(self) -> str:
        parts = ", ".join(
            f"{count} {name}"
            for name, count in sorted(self.reasons.items(), key=lambda item: -item[1])
        )
        return parts or "none"


def _compute_stats(
    analyzed: list[Block], filtered: list[Block], threshold: float
) -> SummaryStats:
    """Single source of truth for every statistic the script reports."""
    scores = [b.ai_pct for b in analyzed]
    total_words = sum(b.words for b in analyzed)
    mean = statistics.fmean(scores) if scores else 0.0
    weighted = (
        sum(score * block.words for score, block in zip(scores, analyzed))
        / max(total_words, 1)
    )
    median = statistics.median(scores) if scores else 0.0
    ordered = sorted(scores)
    p90 = ordered[min(len(ordered) - 1, int(0.9 * len(ordered)))] if ordered else 0.0
    over_50 = sum(1 for score in scores if score > 50.0)
    flagged = [b for b in analyzed if b.ai_pct > threshold]
    truncated = [b for b in analyzed if b.truncated]
    reasons: dict[str, int] = {}
    for block in filtered:
        key = block.filtered_reason or "?"
        reasons[key] = reasons.get(key, 0) + 1
    return SummaryStats(
        mean=mean,
        weighted=weighted,
        median=median,
        p90=p90,
        over_50=over_50,
        flagged=flagged,
        truncated=truncated,
        total_words=total_words,
        reasons=reasons,
    )


def print_summary(
    analyzed: list[Block],
    filtered: list[Block],
    extracted_count: int,
    args: argparse.Namespace,
    device: str,
    elapsed: float,
    path: Path,
) -> None:
    stats = _compute_stats(analyzed, filtered, args.threshold)
    total_words_scored = stats.total_words
    mean, weighted, median, p90 = stats.mean, stats.weighted, stats.median, stats.p90
    over_50 = stats.over_50
    flagged = stats.flagged
    truncated = stats.truncated
    reason_text = stats.reason_text

    line = "=" * 78
    print(line)
    print(f" AI-DETECTION SUMMARY {DASH} {path.name}")
    print(line)
    print(f" Model:              {args.model}")
    print(f" Device:             {device}   |  Analysis time: {elapsed:.1f}s")
    print(f" Paragraphs found:   {extracted_count}")
    print(f" Analyzed:           {len(analyzed)} blocks / {total_words_scored:,} words")
    print(f" Filtered out:       {len(filtered)} blocks ({reason_text})")
    print(f" Truncated >512 tok: {len(truncated)}"
          + ("   [rerun with --split-long to score full text]" if truncated else ""))
    print("-" * 78)
    print(f" Mean AI probability (per block):    {mean:5.1f} %")
    print(f" Word-weighted AI probability:       {weighted:5.1f} %")
    print(f" Median / 90th percentile:           {median:5.1f} % / {p90:.1f} %")
    print(f" Blocks above 50%:                   {over_50:5d}  "
          f"({100.0 * over_50 / max(len(analyzed), 1):.1f} %)")
    print(f" Blocks above {args.threshold:.0f}%:                 {len(flagged):5d}  "
          f"({100.0 * len(flagged) / max(len(analyzed), 1):.1f} %)")
    print("-" * 78)

    if flagged:
        print(f" {WARN_SYMBOL} HIGH-RISK PARAGRAPHS (AI > {args.threshold:.0f}%):")
        for block in sorted(flagged, key=lambda b: -b.ai_pct):
            print(f"   {block.label:>7}   AI {block.ai_pct:5.1f} %   "
                  f"{block.words:4d} words   {QUOTE_OPEN}{_snippet(block.text)}{QUOTE_CLOSE}")
    else:
        print(f" No paragraph exceeded the {args.threshold:.0f}% threshold.")

    if args.top > 0 and analyzed:
        print("-" * 78)
        print(f" TOP {args.top} PARAGRAPHS BY AI PROBABILITY:")
        for block in sorted(analyzed, key=lambda b: -b.ai_pct)[: args.top]:
            marker = f" {WARN_SYMBOL}" if block.ai_pct > args.threshold else "  "
            print(f"{marker} {block.label:>7}   AI {block.ai_pct:5.1f} %   "
                  f"{block.words:4d} words   {QUOTE_OPEN}{_snippet(block.text)}{QUOTE_CLOSE}")

    print(line)
    print(" Note: paragraph-level classifier (~0.8% false-positive rate on its")
    print(" test set). Use as supporting evidence, never as a sole verdict.")
    print(line)


def write_csv(path: Path, analyzed: list[Block], filtered: list[Block],
              threshold: float) -> None:
    with path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.writer(handle)
        writer.writerow([
            "label", "status", "ai_probability_pct", "words", "tokens",
            "truncated", "above_threshold", "filtered_reason", "snippet",
        ])
        for block in analyzed:
            writer.writerow([
                block.label, "analyzed", f"{block.ai_pct:.2f}", block.words,
                block.tokens, "yes" if block.truncated else "no",
                "yes" if block.ai_pct > threshold else "no", "",
                _snippet(block.text, 200),
            ])
        for block in filtered:
            writer.writerow([
                block.label, "filtered", "", block.words, "", "", "",
                block.filtered_reason or "", _snippet(block.text, 200),
            ])


# --------------------------------------------------------------------------- #
# HTML report (human-readable: every paragraph in FULL, never a snippet)
# --------------------------------------------------------------------------- #
def write_report(
    path: Path,
    analyzed: list[Block],
    filtered: list[Block],
    extracted_count: int,
    stats: SummaryStats,
    args: argparse.Namespace,
    device: str,
    elapsed: float,
    source: Path,
) -> None:
    """Write a self-contained HTML report with every paragraph in full.

    The CSV export stays machine-oriented (snippet column); this report is for
    humans: color-coded risk badges, complete paragraph text, and exactly the
    same statistics as the console summary. No external assets - the file
    opens in any browser and prints/exports to PDF cleanly.
    """
    threshold = args.threshold
    now = time.strftime("%Y-%m-%d %H:%M")

    def badge_class(pct: float) -> str:
        if pct > threshold:
            return "high"
        if pct > 50.0:
            return "mid"
        return "low"

    def block_card(block: Block) -> str:
        """One analyzed paragraph with its COMPLETE text."""
        cls = badge_class(block.ai_pct)
        tags = ""
        if block.truncated:
            tags += '<span class="tag">TRUNCATED &gt;512 tokens</span>'
        if block.part:
            tags += f'<span class="tag part">part {block.part}</span>'
        return (
            f'<article class="block {cls}">'
            f'<div class="bhead">'
            f'<span class="label">{block.label}</span>'
            f'<span class="badge {cls}">AI {block.ai_pct:.1f}%</span>'
            f"<span>{block.words} words</span>"
            f"<span>{block.tokens} tokens</span>{tags}"
            "</div>"
            f'<p class="btext">{html.escape(block.text)}</p>'
            "</article>"
        )

    def filtered_card(block: Block) -> str:
        """One excluded block with its COMPLETE text and the filter reason."""
        reason = html.escape(block.filtered_reason or "?")
        return (
            '<article class="block">'
            f'<div class="bhead"><span class="label">{block.label}</span>'
            f'<span class="tag reason">{reason}</span>'
            f"<span>{block.words} words</span></div>"
            f'<p class="btext small">{html.escape(block.text)}</p>'
            "</article>"
        )

    css = """
* { box-sizing: border-box; }
body { margin: 0; font-family: 'Segoe UI', system-ui, Arial, sans-serif;
       background: #f2f4f7; color: #1b1f24; line-height: 1.55; }
.wrap { max-width: 1000px; margin: 0 auto; padding: 24px 18px 64px; }
header.report { background: linear-gradient(135deg, #182230, #243447); color: #fff;
  border-radius: 12px; padding: 26px 30px; box-shadow: 0 2px 8px rgba(0,0,0,.15); }
header.report h1 { margin: 0 0 4px; font-size: 26px; }
header.report .docname { opacity: .85; font-size: 15px; margin-bottom: 16px; }
.meta { display: grid; grid-template-columns: repeat(auto-fit, minmax(270px, 1fr));
  gap: 6px 26px; font-size: 13.5px; }
.meta span { display: inline-block; min-width: 110px; opacity: .7; }
.toc { display: flex; flex-wrap: wrap; gap: 8px; margin: 16px 2px 0; }
.toc a { background: #fff; border: 1px solid #d6dbe2; border-radius: 999px;
  padding: 5px 14px; font-size: 13.5px; color: #243447; text-decoration: none; }
.toc a:hover { background: #eef2f7; }
section { background: #fff; border: 1px solid #e0e4ea; border-radius: 12px;
  padding: 20px 26px 24px; margin-top: 18px; }
section h2 { margin: 0 0 14px; font-size: 19px; color: #182230; }
section h2.danger { color: #b3261e; }
.cards { display: grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr));
  gap: 12px; }
.card { background: #f7f9fc; border: 1px solid #e3e8ef; border-radius: 10px;
  padding: 12px 16px; }
.card .v { font-size: 27px; font-weight: 700; }
.card .k { font-size: 12.5px; color: #5b6470; }
table.detail { border-collapse: collapse; margin-top: 16px; width: 100%;
  font-size: 14.5px; }
table.detail th, table.detail td { text-align: left; padding: 7px 10px;
  border-bottom: 1px solid #eceff3; vertical-align: top; }
table.detail th { font-weight: 600; color: #46505c; width: 270px; }
.disclaimer { background: #fff8e1; border: 1px solid #ffe082; border-radius: 8px;
  padding: 10px 14px; font-size: 13.5px; margin-top: 16px; }
article.block { border: 1px solid #e5e8ee; border-left: 5px solid #9aa4b0;
  border-radius: 8px; padding: 10px 16px 12px; margin: 12px 0; background: #fbfcfd; }
article.block.high { border-left-color: #c62828; background: #fff5f5; }
article.block.mid { border-left-color: #ef6c00; background: #fffaf2; }
article.block.low { border-left-color: #2e7d32; background: #f5fbf5; }
.bhead { display: flex; flex-wrap: wrap; gap: 8px 16px; align-items: center;
  font-size: 13px; color: #57606b; }
.label { font-weight: 700; color: #182230; font-size: 14px; }
.badge { color: #fff; border-radius: 999px; padding: 2px 12px; font-weight: 700;
  font-size: 13px; }
.badge.high { background: #c62828; }
.badge.mid { background: #ef6c00; }
.badge.low { background: #2e7d32; }
.tag { background: #6a1b9a; color: #fff; font-size: 11px; padding: 2px 8px;
  border-radius: 4px; font-weight: 600; }
.tag.part { background: #455a64; }
.tag.reason { background: #546e7a; }
.btext { margin: 8px 0 0; text-align: justify; font-size: 15px; }
.btext.small { font-size: 13.5px; color: #3c4450; }
.hint { font-size: 13.5px; color: #5b6470; margin-top: -6px; }
details { margin-top: 12px; }
summary { cursor: pointer; font-weight: 600; font-size: 14.5px; }
@media print {
  body { background: #fff; }
  .toc { display: none; }
  section { border: none; page-break-inside: avoid; }
  article.block { page-break-inside: avoid; }
}
"""

    top_link = f'<a href="#top">Top {args.top}</a>' if args.top > 0 else ""

    parts: list[str] = []
    parts.append(f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>AI Text Detection Report &mdash; {html.escape(source.name)}</title>
<style>{css}</style>
</head>
<body>
<div class="wrap">
<header class="report">
  <h1>AI Text Detection Report</h1>
  <div class="docname">{html.escape(source.name)}</div>
  <div class="meta">
    <div><span>Document</span>{html.escape(str(source.resolve()))}</div>
    <div><span>Generated</span>{now}</div>
    <div><span>Model</span>{html.escape(args.model)}</div>
    <div><span>Device</span>{html.escape(device)}</div>
    <div><span>Analysis time</span>{elapsed:.1f}s</div>
    <div><span>Parameters</span>threshold {threshold:.0f}%, min-words {args.min_words}, max_length {MAX_LENGTH}</div>
  </div>
</header>
<nav class="toc">
  <a href="#summary">Summary</a>
  <a href="#highrisk">High-risk ({len(stats.flagged)})</a>
  {top_link}
  <a href="#analyzed">All analyzed ({len(analyzed)})</a>
  <a href="#filtered">Filtered ({len(filtered)})</a>
</nav>
""")

    over_50_pct = 100.0 * stats.over_50 / max(len(analyzed), 1)
    flagged_pct = 100.0 * len(stats.flagged) / max(len(analyzed), 1)
    truncated_note = (
        " &mdash; rerun with <code>--split-long</code> to score the full text"
        if stats.truncated
        else ""
    )

    parts.append(f"""
<section id="summary">
  <h2>Summary</h2>
  <div class="cards">
    <div class="card"><div class="v">{stats.mean:.1f}%</div><div class="k">Mean AI probability</div></div>
    <div class="card"><div class="v">{stats.weighted:.1f}%</div><div class="k">Word-weighted AI probability</div></div>
    <div class="card"><div class="v">{stats.median:.1f}%</div><div class="k">Median</div></div>
    <div class="card"><div class="v">{stats.p90:.1f}%</div><div class="k">90th percentile</div></div>
  </div>
  <table class="detail">
    <tr><th>Paragraphs extracted</th><td>{extracted_count}</td></tr>
    <tr><th>Analyzed</th><td>{len(analyzed)} blocks / {stats.total_words:,} words</td></tr>
    <tr><th>Filtered out</th><td>{len(filtered)} blocks ({stats.reason_text})</td></tr>
    <tr><th>Truncated &gt;512 tokens</th><td>{len(stats.truncated)}{truncated_note}</td></tr>
    <tr><th>Blocks above 50%</th><td>{stats.over_50} ({over_50_pct:.1f}%)</td></tr>
    <tr><th>Blocks above {threshold:.0f}%</th><td>{len(stats.flagged)} ({flagged_pct:.1f}%)</td></tr>
  </table>
  <p class="disclaimer"><b>Note:</b> paragraph-level classifier (~0.8% false-positive
  rate on its test set). Use as supporting evidence, never as a sole verdict.</p>
</section>
""")

    if stats.flagged:
        flagged_html = "".join(
            block_card(block)
            for block in sorted(stats.flagged, key=lambda b: -b.ai_pct)
        )
        parts.append(f"""<section id="highrisk">
  <h2 class="danger">High-risk paragraphs (AI &gt; {threshold:.0f}%)</h2>
{flagged_html}
</section>
""")
    else:
        parts.append(f"""<section id="highrisk">
  <h2>High-risk paragraphs (AI &gt; {threshold:.0f}%)</h2>
  <p class="hint">No paragraph exceeded the {threshold:.0f}% threshold.</p>
</section>
""")

    if args.top > 0 and analyzed:
        top_html = "".join(
            block_card(block)
            for block in sorted(analyzed, key=lambda b: -b.ai_pct)[: args.top]
        )
        parts.append(f"""<section id="top">
  <h2>Top {args.top} paragraphs by AI probability</h2>
{top_html}
</section>
""")

    analyzed_html = "".join(block_card(block) for block in analyzed)
    parts.append(f"""<section id="analyzed">
  <h2>All analyzed paragraphs ({len(analyzed)})</h2>
  <p class="hint">Complete paragraph text in document order - nothing is shortened.</p>
{analyzed_html}
</section>
""")

    if filtered:
        filtered_html = "".join(filtered_card(block) for block in filtered)
        filtered_body = (
            "  <details>\n"
            "    <summary>Show all filtered blocks with full text</summary>\n"
            f"{filtered_html}\n"
            "  </details>\n"
        )
    else:
        filtered_body = '  <p class="hint">None.</p>\n'
    parts.append(f"""<section id="filtered">
  <h2>Filtered blocks ({len(filtered)})</h2>
  <p class="hint">Excluded before scoring so they cannot skew the statistics.</p>
{filtered_body}</section>
""")

    parts.append("""</div>
</body>
</html>
""")
    path.write_text("".join(parts), encoding="utf-8")

# --------------------------------------------------------------------------- #
# CLI / entry point
# --------------------------------------------------------------------------- #
def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Paragraph-level AI-text detection for long documents "
                    f"using {MODEL_NAME}.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
        epilog="Example: python ai_detector.py thesis.docx --threshold 70 --csv out.csv",
    )
    parser.add_argument("file", help="path to the document (.docx, .pdf, .txt, .md)")
    parser.add_argument("--model", default=MODEL_NAME, metavar="ID_OR_DIR",
                        help="Hugging Face model id or local model folder")
    parser.add_argument("--threshold", type=float, default=DEFAULT_THRESHOLD,
                        help="AI %% above which a paragraph is flagged as a spike")
    parser.add_argument("--min-words", type=int, default=DEFAULT_MIN_WORDS,
                        help="drop paragraphs with fewer words than this")
    parser.add_argument("--top", type=int, default=10,
                        help="how many highest-scoring paragraphs to show (0 = hide)")
    parser.add_argument("--batch-size", type=int, default=1,
                        help="paragraphs per forward pass (>1 is faster)")
    parser.add_argument("--device", choices=("auto", "cpu", "cuda"), default="auto",
                        help="inference device")
    parser.add_argument("--csv", metavar="PATH", default=None,
                        help="also write every block score to a CSV file")
    parser.add_argument("--report", metavar="PATH", default=None,
                        help="also write a human-readable HTML report "
                             "(every paragraph in full, no snippets)")
    parser.add_argument("--include-tables", action="store_true",
                        help="DOCX only: include table cell text")
    parser.add_argument("--keep-references", action="store_true",
                        help="do not filter bibliography/reference-style lines")
    parser.add_argument("--split-long", action="store_true",
                        help="split >512-token paragraphs on sentence boundaries "
                             "instead of truncating them")
    args = parser.parse_args(argv)
    if args.batch_size < 1:
        parser.error("--batch-size must be >= 1")
    if args.min_words < 1:
        parser.error("--min-words must be >= 1")
    if not 0.0 <= args.threshold <= 100.0:
        parser.error("--threshold must be between 0 and 100")
    return args


def main(argv: list[str] | None = None) -> int:
    _configure_streams()
    args = parse_args(argv)
    source = Path(args.file).expanduser()
    started = time.perf_counter()

    print(f" Reading document: {source}")
    raw = read_document(source, include_tables=args.include_tables)
    paragraphs = split_paragraphs(raw)
    if not paragraphs:
        print(" No text found in the document (is it a scanned/image PDF?).",
              file=sys.stderr)
        return 2

    analyzed: list[Block] = []
    filtered: list[Block] = []
    for position, text in enumerate(paragraphs, start=1):
        block = Block(position=position, text=text, words=len(text.split()))
        reason = filter_reason(text, args.min_words, args.keep_references)
        if reason is None:
            analyzed.append(block)
        else:
            block.filtered_reason = reason
            filtered.append(block)
    if not analyzed:
        print(" No analyzable prose after filtering; try a lower --min-words.",
              file=sys.stderr)
        return 2
    print(f" Paragraphs: {len(paragraphs)} extracted | "
          f"{len(analyzed)} analyzable | {len(filtered)} filtered")

    device = resolve_device(args.device)
    tokenizer, model, ai_index = load_model(device, args.model)
    if ai_index != 1:
        print(f"  note: AI class found at logit index {ai_index} "
              f"(labels: {model.config.id2label})")

    if args.split_long:
        for block in analyzed:  # pre-compute token lengths for windowing
            block.tokens = len(tokenizer.encode(block.text, add_special_tokens=True))
        scoring_targets: list[Block] = split_long_blocks(analyzed, tokenizer)
    else:
        scoring_targets = analyzed

    print(f" Scoring {len(scoring_targets)} block(s) on {device} "
          f"(batch size {args.batch_size}, max_length={MAX_LENGTH}) ...")
    score_blocks(scoring_targets, tokenizer, model, device, ai_index, args.batch_size)

    elapsed = time.perf_counter() - started
    device_label = device
    if device == "cuda":
        import torch
        device_label = f"cuda ({torch.cuda.get_device_name(0)})"

    print()
    print_summary(analyzed, filtered, len(paragraphs), args, device_label,
                  elapsed, source)
    if args.csv:
        csv_path = Path(args.csv).expanduser()
        write_csv(csv_path, analyzed, filtered, args.threshold)
        print(f" Block-level results written to: {csv_path.resolve()}")
    if args.report:
        report_path = Path(args.report).expanduser()
        write_report(
            report_path,
            analyzed,
            filtered,
            len(paragraphs),
            _compute_stats(analyzed, filtered, args.threshold),
            args,
            device_label,
            elapsed,
            source,
        )
        print(f" HTML report written to: {report_path.resolve()}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print("\n Interrupted.", file=sys.stderr)
        raise SystemExit(130)
