"""Generate sample_thesis.pdf (selectable-text) to exercise the pypdf reader."""
from pathlib import Path

from fpdf import FPDF

BLOCKS = [
    "Chapter 2",
    "Literature Review",
    "Smith, J. A., & Doe, R. B. (2019). Methodologies for sample studies. "
    "Journal of Sample Research, 12(3), 45-67.",
    "47 33 12 88 5 91 300 42",
    "The literature on estuarine salinity monitoring can be divided into three "
    "broad strands: ship-based transect campaigns, fixed-station logger networks, "
    "and remote-sensing proxies. Each strand carries its own sampling bias, and "
    "the thesis argues throughout that reconciling them matters more than "
    "extending any single one. Transect campaigns offer spatial coverage but "
    "rarely persist beyond a funding cycle, whereas logger networks provide "
    "unparalleled temporal depth at exactly the points where a harbour authority "
    "happened to grant permission for an instrument to be bolted to a pier.",
    "Furthermore, it is important to note that the comprehensive framework "
    "presented in this chapter underscores the significance of integrating "
    "multidisciplinary perspectives. In conclusion, the findings demonstrate a "
    "robust correlation between the variables examined, thereby highlighting the "
    "critical role of structured methodologies. Moreover, it is worth noting "
    "that stakeholders should leverage these insights to facilitate "
    "data-driven decision-making across all organizational levels.",
    "By contrast, the remote-sensing literature has mostly concerned itself with "
    "coarse pixel products that smooth away the very gradients this thesis "
    "treats as signal rather than noise, so its contribution here is limited to "
    "contextualising the aerial photographs flown after the 1998 storm.",
    "References",
    "Doe, R. B. (2021). Tidal salinity records revisited. Coastal Science "
    "Letters, 8(2), 101-119.",
]

pdf = FPDF()
pdf.set_auto_page_break(auto=True, margin=15)
pdf.add_page()
pdf.set_font("Helvetica", size=11)
for block in BLOCKS:
    pdf.multi_cell(0, 6, block)
    pdf.ln(4)
out = Path(__file__).with_name("sample_thesis.pdf")
pdf.output(str(out))
print(f"wrote {out}")
