# Local PDF / DOCX to Markdown Web App

A local web app that converts PDF and DOCX files into structured Markdown with heading detection, paragraph preservation, and inline formatting. Uses char-level font analysis for intelligent structure recognition. Falls back to Tesseract OCR for scanned or image-only PDFs.

## What it does

### Structure detection

- Analyses char-level font metadata (size, weight, family) to detect document hierarchy
- Titles, section headings, and subheadings rendered as `#`, `##`, `###` based on font size ratios
- Paragraph breaks preserved using vertical line spacing analysis
- Bullet lists detected from x-position indentation relative to body text margin
- Word boundaries preserved by detecting x-position gaps between characters

### Text formatting

- Inline **bold** and *italic* spans detected from font metadata
- Bold-via-character-overlap detected (common in email-to-PDF rendering)
- Ligatures, smart quotes, em-dashes, and repeated characters cleaned automatically
- Page boilerplate (headers, footers, page numbers, URLs) filtered by font size

### Tables

- Tables detected and extracted using `pdfplumber` structural analysis
- Duplicate/ghost columns from PDF rendering merged automatically
- Tables formatted as Markdown with header rows and separator lines

### File support

- **PDF** (embedded text): `pdfplumber` with char-level analysis, falls back to `pdfminer.six`
- **PDF** (scanned/image): OCR via `pdf2image` + `Tesseract`
- **DOCX**: HTML-to-Markdown conversion via `mammoth`

### Web interface

- Drag-and-drop or file selection (up to 50MB)
- Raw Markdown or rendered preview toggle
- Copy to clipboard or download as `.md` file
- Conversion method indicator (Embedded text / OCR / DOCX)

## Setup and run

From the `pdf-to-markdown` folder:

```bash
./run.sh
```

Then open:

```text
http://127.0.0.1:5000
```

The startup script installs system dependencies (`tesseract`, `poppler`), creates a Python virtual environment, and finds an available port (5000-5010).

## Stop the server

Press `Ctrl+C` in the terminal where the app is running.

## Requirements

- macOS with Homebrew
- Python 3.9+
- `tesseract` and `poppler` (installed automatically by `run.sh`)

## How it works

The extraction pipeline analyses each PDF page at the character level:

1. **Character grouping** - Groups `page.chars` into lines by y-position, deduplicates overlapping characters (bold-via-overlap detection), and inserts spaces at x-position gaps
2. **Font classification** - Determines the body font (most frequent at >= 8pt), then classifies each line by font size ratio and weight relative to body text
3. **Structure detection** - Lines with size > 1.35x body become titles; bold sans-serif at body size become section headings; standalone bold lines become subheadings; indented lines become bullets
4. **Paragraph detection** - Vertical gaps > 1.8x the median line spacing insert paragraph breaks
5. **Table extraction** - `pdfplumber.find_tables()` detects structured tables, with column deduplication for offset/overlapping column grids
6. **Normalisation** - Ligatures, smart quotes, repeated characters, and page boilerplate stripped; heading and bullet lines preserved as standalone blocks

All detection is pattern-based and content-agnostic - no document-specific rules are embedded.

## Known limitations

- Complex tables with merged cells or irregular layouts may not convert perfectly
- Multi-page tables may show repeated headers (one per PDF page)
- Images and embedded graphics are not extracted
- OCR quality depends on scan quality and page layout
- Non-English text may produce imperfect OCR results unless Tesseract language packs are configured
- DOCX conversion is simplified via HTML-to-Markdown rules and may not preserve every style nuance
