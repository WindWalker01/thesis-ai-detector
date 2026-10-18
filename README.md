# Thesis AI Detector

**What is this?** A small program that reads a long document (your thesis, essay, or report) and estimates, paragraph by paragraph, how likely each part looks like it was written by AI. Everything runs on **your own computer** -- after the one-time setup you don't need the internet.

**What you get at the end:** a `report.html` file you double-click to open in your web browser, with every paragraph shown in full and colour-coded (green = looks human, yellow = uncertain, red = looks AI-generated), plus a headline score for the whole document.

> IMPORTANT, please read: no AI detector is perfect. This tool is about **98-99% accurate per paragraph on academic text** (roughly 1 in 100 human paragraphs may be wrongly flagged). **Never use this result alone to accuse someone of cheating.** Treat it as a starting point for a human conversation, not proof.

---

## 1. What you need

- A **Windows** PC or a **Mac**.
- Your document: a **Word file (`.docx`) works best**. PDF (`.pdf`) and plain text (`.txt`, `.md`) also work.
- **Python 3.10 or newer** (free). How to check:
  - Windows: open the Start menu, type `cmd`, open it, type `python --version` and press Enter.
  - Mac: open **Terminal** (Spotlight search for "Terminal"), type `python3 --version` and press Enter.
  - If you see a version number like `Python 3.12.4`, you're good. If you see an error, install it first:
    1. Go to **https://www.python.org/downloads/**
    2. Download and install Python.
    3. **Windows, very important:** on the first install screen, tick the box **"Add python.exe to PATH"** before clicking Install.
- About **500 MB free space** (downloaded once) and **5-15 minutes** for the first setup. Analysing one thesis takes about **1-5 minutes**.

---

## 2. The easiest way: double-click to run (recommended)

You don't need to type commands. This project includes a launcher that sets everything up and runs the check for you.

**Step 1 -- Get the files.**
Download this project folder onto your computer and open it in File Explorer (Windows) or Finder (Mac). You should see files like `Run.bat`, `run.sh`, `ai_detector.py`, and `requirements.txt`.

**Step 2 -- Put your thesis next to the launcher.**
Copy your thesis file (e.g. `thesis.docx`) into this same folder. If you're nervous, practice first with the included `sample_thesis.docx` -- it works exactly the same way.

**Step 3 -- Run it.**

| Your computer | What to do |
|---|---|
| **Windows** | **Double-click `Run.bat`**. Even easier: **drag your thesis file and drop it onto `Run.bat`**. Otherwise it will ask you to type the file name (e.g. `thesis.docx`) and press Enter. |
| **Mac** | Open **Terminal** in this folder (right-click the folder, then *New Terminal at Folder*, or drag the folder onto the Terminal icon), then type `bash run.sh thesis.docx` and press Enter. To be asked for the file name instead, just run `bash run.sh`. First time only: if macOS blocks it, right-click `run.sh`, choose *Open* once to approve it. |

**What happens next (all automatic):**
1. First run only: it creates a private folder called `.venv` (this keeps this project's libraries separate so nothing else on your PC breaks) and downloads the needed libraries. This takes a few minutes -- the window may look frozen, that's normal.
2. It analyses your document and shows progress in the window.
3. When it says **DONE!**, close the window and **double-click `report.html`** to see your results in the browser. `results.csv` (opens in Excel) lists every paragraph with its score.

> Tip: run it once on `sample_thesis.docx` first so you know what a finished run looks like before using your real thesis.

---

- **The headline number** at the top = the average AI-likelihood across all paragraphs. Low (e.g. under 30%) means the document overall reads as human. High (e.g. over 70%) means most of it reads as AI-generated.
- **Each paragraph** has a coloured badge: Green (low) = human, Yellow (middle) = uncertain, Red (above 70%) = looks AI-generated ("spike").
- **How to interpret spikes:** a few red paragraphs inside an otherwise green document usually means local AI help. A mostly red document means widespread generation.
- **What was skipped:** very short bits (titles, page numbers), web links, and bibliography lines are auto-skipped so they don't distort the score.

If you want spreadsheet rows, open **results.csv** in Excel: one row per paragraph with word count and score.

---

## 4. If something goes wrong

| What you see | What to do |
|---|---|
| `Python was not found` | Install Python from python.org. On Windows tick "Add python.exe to PATH". Close and reopen the terminal, try again. |
| `File not found: thesis.docx` | Copy the file next to Run.bat / run.sh, or type the full path (Windows: C:\Users\You\Documents\thesis.docx, Mac: /Users/You/Documents/thesis.docx). |
| `No text found (scanned/image PDF?)` | Your PDF is photos of pages, not real text. If you can't select text with the mouse, it's scanned. Open it in Word or an OCR tool, save as .docx, retry. |
| Download stalls / weak internet | Use the built-in offline model copy: ask a tech-savvy friend to run `python ai_detector.py thesis.docx --model model_local --report report.html --csv results.csv`. After one good online run, later runs work offline. |
| It takes very long | Normal: 1-5 min for a full thesis. Keep the window open and wait. |
| The window opens then closes immediately | Old launcher bug, fixed. Download the latest `Run.bat`. Workaround for any launcher: open a terminal in the folder first, then run `.\Run.bat` so the error stays visible. |
| Strange `?` or `!!` characters | Old terminal encoding; results are unaffected. On Windows type `chcp 65001` and rerun. |

Still stuck? Copy the entire window text and send it to whoever gave you this project.

---

## 5. Typing the commands yourself (optional)

The launcher does exactly this for you. Only read on if you prefer the terminal.

**Windows (inside the project folder):**

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python ai_detector.py thesis.docx --report report.html --csv results.csv
```

**Mac (Terminal, inside the project folder):**

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python ai_detector.py thesis.docx --report report.html --csv results.csv
```

Try the samples: `python ai_detector.py sample_thesis.docx --report report.html --csv out.csv`.

---

## Appendix: for tech-savvy users (everyone else can ignore this)

**How it works:** Hugging Face model `followsci/bert-ai-text-detector` (BERT-base, label 0 = Human, label 1 = AI, 512-token limit). Long documents are scored paragraph by paragraph and aggregated into mean, word-weighted mean, median, p90, and over-threshold counts. Short blocks, URLs, symbol-heavy lines, and bibliography lines are filtered before scoring.

| Flag | Default | Purpose |
|---|---|---|
| `--threshold` | `70` | AI % above which a paragraph is flagged. |
| `--min-words` | `8` | Drop shorter blocks (titles, page numbers). |
| `--top` | `10` | Top paragraphs to print (`0` = hide). |
| `--batch-size` | `1` | Paragraphs per pass; 8-16 is faster. |
| `--device` | `auto` | `auto` / `cpu` / `cuda`. |
| `--model` | HF id | Model id or local folder (e.g. `--model model_local`) for offline runs. |
| `--csv PATH` | off | Write block scores to CSV. |
| `--report PATH` | off | Write the HTML report. |
| `--split-long` | off | Sentence-window long paragraphs instead of truncating. |
| `--keep-references` | off | Keep bibliography lines. |
| `--include-tables` | off | DOCX: also score tables. |

CPU torch works out of the box (21k-word thesis: ~30 s to a few min). For NVIDIA GPUs pick a CUDA build at https://pytorch.org/start; `CUDA requested but not usable` means rerun with `--device cpu`. First online run downloads ~418 MB to the HF cache; later runs are offline. DOCX recommended; PDFs need selectable text (scanned PDFs have none).
