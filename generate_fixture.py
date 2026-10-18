"""Generate sample_thesis.docx from the same content as sample_thesis.txt."""
from pathlib import Path

import docx

HUMAN_1 = (
    "The intertidal zone of the northern peninsula has been monitored irregularly "
    "since the early 1990s, yet no study has attempted to reconcile the conflicting "
    "salinity records kept by the regional port authority with those logged by the "
    "university field station. This thesis addresses that gap. Fieldwork was carried "
    "out over fourteen months, during which water samples were drawn at six fixed "
    "stations every fortnight and analysed by a coupled ion-chromatography workflow "
    "that we validated against certified reference material before each campaign."
)
HUMAN_2 = (
    "Because the sampling campaign spanned two hydrological years, we were careful to "
    "distinguish seasonal drift from genuine regime shifts, and every figure reports "
    "confidence intervals derived from block bootstrapping rather than from the naive "
    "standard error, a choice whose consequences are discussed in detail in section 4.3."
)
AI_TEXT = (
    "Furthermore, it is important to note that the comprehensive framework presented "
    "in this chapter underscores the significance of integrating multidisciplinary "
    "perspectives. In conclusion, the findings demonstrate a robust correlation "
    "between the variables examined, thereby highlighting the critical role of "
    "structured methodologies. Moreover, it is worth noting that stakeholders should "
    "leverage these insights to facilitate data-driven decision-making across all "
    "organizational levels."
)
LONG = (
    "The following appendix paragraph deliberately exceeds the transformer window so "
    "that the truncation warning can be verified end to end. It repeats a discussion "
    "of laboratory quality assurance procedures at length, covering calibration drift, "
    "blank contamination, column ageing, detector noise, sample carry-over, "
    "autosampler rinsing, temperature control of the mobile phase, buffer preparation "
    "logs, analyst training records, chain-of-custody forms, retention time "
    "alignment, integration parameters, audit trail exports, instrument maintenance "
    "windows, spare part inventories, software version pinning, reference standard "
    "certificates, proficiency testing rounds, measurement uncertainty budgets, and "
    "internal quality control charts. In addition, it elaborates on how each of these "
    "elements interacts with the others during a fourteen-month campaign: calibration "
    "drift is countered by bracketing standards, blank contamination is tracked "
    "through triplicate reagent blanks, column ageing is recorded alongside "
    "back-pressure traces, detector noise is quantified during overnight vacuum runs, "
    "sample carry-over is measured after high-concentration samples, and management "
    "review closes the cycle each quarter with a documented decision record. Beyond "
    "the checklist above, the campaign log records the weather conditions prevailing "
    "during each sampling round, the identity of the analyst on duty, the lot number "
    "of every reagent bottle opened, and the exact minute at which each vial was "
    "capped and placed in the cooled transport box. Deviations from the standard "
    "operating procedure are annotated in the margin of the field sheets and "
    "transcribed into the electronic record before the data are locked, so that any "
    "later query can be answered by pointing at the original page rather than at a "
    "reconstructed narrative. The audit trail itself is exported weekly to "
    "write-once media, hashed, and countersigned by a second analyst who did not "
    "participate in the measurements, which makes silent alteration of the record "
    "effectively impossible without detection."
)
CLOSING = (
    "The final paragraph returns to ordinary prose after the long appendix block, "
    "offering a short closing reflection on how the peninsula's monitoring practices "
    "might be modernised with automated sondes without losing the interpretive value "
    "that local knowledge brings to the record."
)

paragraphs = [
    "Chapter 1",
    "Introduction",
    "Smith, J. A., & Doe, R. B. (2019). Methodologies for sample studies. "
    "Journal of Sample Research, 12(3), 45-67.",
    "https://example.com/resources/2019",
    "47 33 12 88 5 91 300 42",
    HUMAN_1,
    HUMAN_2,
    AI_TEXT,
    LONG,
    "References",
    "Doe, R. B. (2021). Tidal salinity records revisited. Coastal Science "
    "Letters, 8(2), 101-119.",
    CLOSING,
]

document = docx.Document()
for text in paragraphs:
    document.add_paragraph(text)
out = Path(__file__).with_name("sample_thesis.docx")
document.save(out)
print(f"wrote {out}")
