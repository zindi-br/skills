---
name: pdf
description: >-
  Use this skill whenever the user wants to do anything with PDF files. This
  includes reading or extracting text/tables from PDFs, merging or splitting
  PDFs, rotating pages, adding watermarks, creating new PDFs, filling PDF forms,
  encrypting/decrypting PDFs, extracting images, and OCR on scanned PDFs. If the
  user mentions a .pdf file or asks to produce one, use this skill. Do NOT use
  this skill when the primary deliverable is a spreadsheet — use the xlsx skill
  instead.
---

# PDF Processing

Essential PDF operations with Python libraries and common CLI tools. Prefer the
smallest tool that solves the task.

## Tool chooser

| Task | Best tool |
|---|---|
| Merge / split / rotate / watermark / encrypt | `pypdf` |
| Extract text or tables | `pdfplumber` |
| Create a new PDF (report, printable doc) | `reportlab` |
| Fast CLI merge/split/decrypt | `qpdf` |
| CLI text extract | `pdftotext` (poppler) |
| Extract embedded images | `pdfimages` (poppler) |
| OCR scanned pages | `pdf2image` + `pytesseract` |
| Fill interactive form fields | `pypdf` (`update_page_form_field_values`) |

Do not `pip install` unless an import fails.

## Quick start

```python
from pypdf import PdfReader, PdfWriter

reader = PdfReader("document.pdf")
print(f"Pages: {len(reader.pages)}")
text = "".join(page.extract_text() or "" for page in reader.pages)
```

## pypdf — structure operations

### Merge

```python
from pypdf import PdfWriter, PdfReader

writer = PdfWriter()
for path in ["doc1.pdf", "doc2.pdf", "doc3.pdf"]:
    reader = PdfReader(path)
    for page in reader.pages:
        writer.add_page(page)

with open("merged.pdf", "wb") as f:
    writer.write(f)
```

### Split (one file per page)

```python
reader = PdfReader("input.pdf")
for i, page in enumerate(reader.pages):
    writer = PdfWriter()
    writer.add_page(page)
    with open(f"page_{i+1}.pdf", "wb") as f:
        writer.write(f)
```

### Rotate

```python
reader = PdfReader("input.pdf")
writer = PdfWriter()
page = reader.pages[0]
page.rotate(90)
writer.add_page(page)
with open("rotated.pdf", "wb") as f:
    writer.write(f)
```

### Watermark

```python
from pypdf import PdfReader, PdfWriter

watermark = PdfReader("watermark.pdf").pages[0]
reader = PdfReader("document.pdf")
writer = PdfWriter()
for page in reader.pages:
    page.merge_page(watermark)
    writer.add_page(page)
with open("watermarked.pdf", "wb") as f:
    writer.write(f)
```

### Encrypt / decrypt

```python
reader = PdfReader("input.pdf")
writer = PdfWriter()
for page in reader.pages:
    writer.add_page(page)
writer.encrypt("userpassword", "ownerpassword")
with open("encrypted.pdf", "wb") as f:
    writer.write(f)

# Decrypt: PdfReader("encrypted.pdf", password="userpassword")
```

### Metadata

```python
reader = PdfReader("document.pdf")
meta = reader.metadata
print(meta.title, meta.author, meta.subject, meta.creator)
```

### Fill fillable form fields

```python
from pypdf import PdfReader, PdfWriter

reader = PdfReader("form.pdf")
writer = PdfWriter()
writer.append(reader)
writer.update_page_form_field_values(writer.pages[0], {
    "field_name": "value",
})
with open("form_filled.pdf", "wb") as f:
    writer.write(f)
```

List field names first if unknown: inspect `reader.get_fields()`.

## pdfplumber — text and tables

```python
import pdfplumber
import pandas as pd

with pdfplumber.open("document.pdf") as pdf:
    text = "\n\n".join(p.extract_text() or "" for p in pdf.pages)

    frames = []
    for i, page in enumerate(pdf.pages):
        for j, table in enumerate(page.extract_tables() or []):
            if table and len(table) > 1:
                df = pd.DataFrame(table[1:], columns=table[0])
                df.attrs["source"] = f"page {i+1} table {j+1}"
                frames.append(df)
```

Prefer `pdfplumber` over `pypdf` for layout-aware text and table extraction.
After extracting tables needed as a spreadsheet deliverable, switch to the
**xlsx** skill for formatting and sheet structure.

## reportlab — create PDFs

### Simple canvas

```python
from reportlab.lib.pagesizes import letter, A4
from reportlab.pdfgen import canvas

c = canvas.Canvas("hello.pdf", pagesize=A4)
width, height = A4
c.drawString(72, height - 72, "Hello World")
c.save()
```

### Multi-page flowables

```python
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.lib.pagesizes import A4

doc = SimpleDocTemplate(
    "report.pdf",
    pagesize=A4,
    leftMargin=0.75 * inch,
    rightMargin=0.75 * inch,
    topMargin=0.75 * inch,
    bottomMargin=0.75 * inch,
)
styles = getSampleStyleSheet()
story = [
    Paragraph("Report Title", styles["Title"]),
    Spacer(1, 12),
    Paragraph("Body text here.", styles["Normal"]),
    PageBreak(),
    Paragraph("Page 2", styles["Heading1"]),
]
doc.build(story)
```

### Subscripts / superscripts

Never use Unicode subscript/superscript glyphs in ReportLab — built-in fonts
lack them and they render as black boxes. Use markup in `Paragraph`:

```python
Paragraph("H<sub>2</sub>O", styles["Normal"])
Paragraph("x<super>2</super>", styles["Normal"])
```

### Readable report defaults

- Page size: A4 unless the user asks for Letter
- Margins ≥ 0.5 inch; keep tables inside margins
- Prefer Platypus `Table` + `TableStyle` for tabular data
- Embed assumptions in a footer line (generation date, source)
- Avoid dumping huge unpaginated tables — split or link to an xlsx instead

## CLI tools

```bash
# Text
pdftotext -layout input.pdf output.txt
pdftotext -f 1 -l 5 input.pdf output.txt

# qpdf merge / slice / decrypt
qpdf --empty --pages file1.pdf file2.pdf -- merged.pdf
qpdf input.pdf --pages . 1-5 -- pages1-5.pdf
qpdf --password=SECRET --decrypt encrypted.pdf decrypted.pdf

# Images
pdfimages -j input.pdf output_prefix
```

## OCR (scanned PDFs)

```python
import pytesseract
from pdf2image import convert_from_path

images = convert_from_path("scanned.pdf")
text = "\n\n".join(
    f"Page {i+1}:\n{pytesseract.image_to_string(img)}"
    for i, img in enumerate(images)
)
```

Requires Poppler for `pdf2image` and a Tesseract install. Tell the user if either is missing.

## Security and hygiene

- Validate uploads: magic bytes should be `%PDF` / MIME `application/pdf` before parsing
- Do not strip owner passwords or bypass encryption unless the user owns the file and provides the password
- Strip or avoid writing sensitive data into PDF metadata (`title`, `author`, custom props)
- Sanitize extracted text before echoing into HTML/Markdown (strip control chars)

```python
import re

def sanitize_extracted_text(text: str) -> str:
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]", "", text)
    return text.strip()
```

## Verification checklist

Before delivering a PDF you created or heavily edited:

- [ ] Opens without errors in a normal viewer
- [ ] Page size/orientation match the content
- [ ] Tables fit margins; text is not clipped
- [ ] No black-box glyphs (sub/superscript misuse)
- [ ] Metadata has no unintended secrets
- [ ] File size is reasonable for the content

## Dependencies

`pypdf`, `pdfplumber`, `reportlab`, `pandas` (pip) · optional: `pdf2image`, `pytesseract`, `python-magic` · CLI: `qpdf`, `pdftotext`, `pdfimages` (poppler)
