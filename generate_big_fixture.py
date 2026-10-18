"""Generate a full-scale synthetic thesis (~21,227 words) for end-to-end testing.

Produces big_thesis.docx (primary) and big_thesis.txt, mixing human-style prose,
formulaic AI-style prose, and noise blocks (headings, citations, URLs) in the
same proportions a real thesis would contain.
"""
import random
from pathlib import Path

import docx

TARGET_WORDS = 21_227
random.seed(42)

HUMAN_SENTENCES = [
    "The sampling campaign began in late March and continued without interruption until the following May.",
    "We chose block bootstrapping over the naive standard error because the residuals showed clear serial correlation.",
    "Harbour masters interviewed in June recalled the 1998 storm surge more vividly than any measurement in the archive.",
    "Salinity readings at station three diverged from the port authority's ledger by roughly two parts per thousand.",
    "The ion-chromatography workflow was validated against certified reference material before each fortnightly campaign.",
    "Two long-time ferry operators confirmed that the channel markers were relocated after the dredging programme.",
    "Field notes transcribed that evening record a pump failure which briefly interrupted the depth profiler series.",
    "Confidence intervals in every figure of this chapter are derived from ten thousand block bootstrap replicates.",
    "The university field station kept its loggers at fixed heights, whereas the port authority adjusted theirs seasonally.",
    "Data locks were applied weekly, and any correction afterwards required a countersigned deviation form.",
    "Section 4.3 shows that the apparent trend survives when the two discrepant stations are removed entirely.",
    "The peninsular geology suggests a fractured bedrock aquifer that would explain the delayed freshwater response.",
    "During the second hydrological year, an unusually dry autumn compressed the estuarine mixing zone seaward.",
    "We refrain from stronger causal language because the design is observational and unreplicated at catchment scale.",
    "Instrument drift was quantified with overnight vacuum runs and found to stay inside manufacturer tolerances.",
    "The interviews were coded independently by two researchers, with disagreements resolved through discussion.",
    "Archival charts from the 1970s were digitised and georeferenced before being compared with modern transects.",
    "One vial from station five was re-analysed after the audit flagged an integration parameter change.",
    "The crew's local knowledge proved indispensable when fog forced two sampling rounds to be rescheduled.",
    "Our results accord with the earlier fjord study, though the magnitude of the seasonal signal is smaller here.",
]

AI_SENTENCES = [
    "Furthermore, it is important to note that the comprehensive framework underscores the significance of integration.",
    "In conclusion, the findings demonstrate a robust correlation between the variables examined herein.",
    "Moreover, it is worth noting that stakeholders should leverage these insights to facilitate decision-making.",
    "Overall, the analysis reinforces the notion that continuous monitoring is essential for sustaining performance.",
    "It should also be emphasized that a multifaceted approach greatly enhances the overall quality of outcomes.",
    "Additionally, the results highlight the critical role of structured methodologies in contemporary research.",
    "Consequently, these insights provide a valuable foundation for informing future strategic initiatives.",
    "In summary, the evidence presented in this section collectively underscores the importance of adaptability.",
    "Taken together, these observations illustrate the profound impact of systematic evaluation on success.",
    "Primarily, this chapter aims to elucidate the key factors that contribute to effective governance.",
    "Importantly, the data reveals a significant relationship that warrants further scholarly attention.",
    "Ultimately, fostering collaboration remains paramount for achieving sustainable, long-term objectives.",
    "Notably, the comprehensive overview offered here serves to contextualize the broader discourse.",
    "Therefore, adopting a holistic perspective is imperative for navigating complex organizational challenges.",
    "Broadly speaking, the interplay between policy and practice continues to shape sector-wide outcomes.",
    "Undeniably, such transformative strategies empower institutions to harness emerging opportunities effectively.",
]

NOISE_BLOCKS = [
    "Chapter 4",
    "Results and Discussion",
    "Table 3.1: Descriptive statistics of sampling stations",
    "https://doi.org/10.1000/example",
    "47 33 12 88 5 91 300 42 16",
    "Figure 5.2 — Monthly salinity profiles",
    "Appendix B",
    "Pereira, L. M., & Santos, R. (2018). Estuarine mixing revisited. Coastal Dynamics Journal, 14(2), 88-104.",
    "Acknowledgements",
    "12.4",
]


def build_paragraph(human: bool) -> str:
    pool = HUMAN_SENTENCES if human else AI_SENTENCES
    count = random.randint(5, 9)
    return " ".join(random.choice(pool) for _ in range(count))


paragraphs: list[str] = []
words = 0
index = 0
while words < TARGET_WORDS:
    if index % 9 == 8:                       # sprinkle noise like a real thesis
        block = random.choice(NOISE_BLOCKS)
    else:
        block = build_paragraph(human=(index % 3 != 0))
    paragraphs.append(block)
    words += len(block.split())
    index += 1

document = docx.Document()
for text in paragraphs:
    document.add_paragraph(text)
docx_path = Path(__file__).with_name("big_thesis.docx")
document.save(docx_path)

txt_path = Path(__file__).with_name("big_thesis.txt")
txt_path.write_text("\n\n".join(paragraphs), encoding="utf-8")

noise = sum(1 for i in range(len(paragraphs)) if i % 9 == 8)
print(f"wrote {docx_path} and {txt_path}")
print(f"paragraphs: {len(paragraphs)} (noise blocks: {noise}) | words: {words}")
