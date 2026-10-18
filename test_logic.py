"""Quick logic test: paragraph splitting + filtering on sample_thesis.txt."""
from pathlib import Path

import ai_detector as a

raw = a.read_plain_text(Path(__file__).with_name("sample_thesis.txt"))
paras = a.split_paragraphs(raw)
print(f"paragraphs extracted: {len(paras)}\n")
counts: dict[str, int] = {}
for i, text in enumerate(paras, 1):
    reason = a.filter_reason(text, 8, False) or "SCORE"
    counts[reason] = counts.get(reason, 0) + 1
    words = len(text.split())
    print(f"{i:2d} | {reason:10s} | {words:3d}w | {text[:72]}")
print("\nverdict counts:", counts)

# targeted assertions
assert any(a.filter_reason(t, 8, False) == "reference" for t in paras), "refs not caught"
assert any(a.filter_reason(t, 8, False) == "short" for t in paras), "short not caught"
assert any(a.filter_reason(t, 8, False) == "url" for t in paras), "url not caught"
assert any(a.filter_reason(t, 8, False) == "non-prose" for t in paras), "digits not caught"
# hard-wrapped paragraph must be rejoined into ONE block (not 4 line fragments)
wrapped = [t for t in paras if t.startswith("Because the sampling campaign")]
assert len(wrapped) == 1 and len(wrapped[0].split()) > 50, "hard-wrap join failed"
# no prose paragraph should be dropped
assert not any(t.startswith("The intertidal") and a.filter_reason(t, 8, False)
               for t in paras), "prose wrongly filtered"
print("\nALL LOGIC TESTS PASSED")
