# PLEADS SCORE ENGINE V2

Separate Streamlit score-processing app:
Upload judge PDFs → auto-detect → integrity check → split per team → merge judges per team → verification → publication gate → central recap CSV.

The parser reads the scoring schema from the PDF, so it is not locked to only the 3-aspect LOC form. It also handles the 6-aspect presentation form in the supplied template.

OCR for image-only/scanned PDFs is not included.
