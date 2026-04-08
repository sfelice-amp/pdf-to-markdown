# PDF / DOCX to Markdown

A local tool that converts PDF and DOCX files into structured Markdown with heading detection, paragraph preservation, and inline formatting. Uses font analysis for intelligent structure recognition. Everything runs locally — no uploads, no cloud services.

Two versions are included:

- **`converter.html`** — A single HTML file that runs entirely in the browser. No server, no installs, no terminal. Just open the file.
- **`app.py`** — A Flask server with OCR fallback for scanned/image-only PDFs via Tesseract.

## Quick start

### Browser version (recommended)

Open `converter.html` in Chrome, Safari, or Firefox. That's it.

- Works from `file://` — no server needed
- Drag-and-drop or click to select a PDF or DOCX
- All processing happens in your browser, nothing leaves your machine

### Server version (OCR support)

If you need OCR for scanned PDFs:

```bash
./run.sh
```

Then open `http://127.0.0.1:5000` in your browser.

The startup script installs system dependencies (`tesseract`, `poppler`), creates a Python virtual environment, and finds an available port (5000–5010). Press `Ctrl+C` to stop the server.

Requires macOS with Homebrew and Python 3.9+.

## What it does

### Structure detection

- Analyses font metadata (size, weight, family) to detect document hierarchy
- Titles, section headings, and subheadings rendered as `#`, `##`, `###` based on font size ratios
- Paragraph breaks preserved using vertical line spacing analysis
- Bullet lists detected from indentation relative to body text margin
- Two-column layouts detected and extracted in reading order (browser version)

### Text formatting

- Inline **bold** and *italic* spans detected from font metadata
- Bold detection from font name parsing (e.g. `GTAmerica-Bold`) when explicit metadata is unavailable
- Ligatures, smart quotes, em-dashes, and repeated characters cleaned automatically
- Page boilerplate (headers, footers, page numbers) filtered by font size

### Tables

- Tables detected from positional gaps between text spans
- Duplicate/ghost columns from PDF rendering merged automatically
- Tables formatted as Markdown with header rows and separator lines

### File support

- **PDF** (embedded text): font-aware extraction with heading and paragraph detection
- **PDF** (scanned/image): OCR via Tesseract (server version only)
- **DOCX**: HTML-to-Markdown conversion via mammoth

### Web interface

- Drag-and-drop or file selection (up to 50 MB)
- Raw Markdown or rendered preview toggle
- Copy to clipboard or download as `.md` file
- Conversion method indicator (Embedded text / OCR / DOCX)

## How it works

The extraction pipeline analyses each PDF page at the character or span level:

1. **Text extraction** — Reads text items with position, size, and font metadata from each page
2. **Column detection** — Identifies two-column layouts by finding consistent vertical gutters near the page center, then splits items into left and right columns in reading order
3. **Line grouping** — Groups text items into lines by y-position, inserting spaces at x-position gaps to preserve word boundaries
4. **Font classification** — Determines the body font (most frequent at ≥ 8pt), then classifies each line by font size ratio and weight relative to body text
5. **Structure detection** — Lines with size > 1.35× body become titles; bold sans-serif at body size become section headings; standalone bold lines become subheadings; indented lines become bullets
6. **Paragraph detection** — Vertical gaps > 1.8× the median line spacing insert paragraph breaks
7. **Table extraction** — Lines with large positional gaps between spans are identified as table rows; consecutive table rows are formatted as Markdown tables with clustered column boundaries
8. **Formatting** — Bold and italic spans within body lines wrapped in `**` and `*` markers; ligatures, smart quotes, and boilerplate cleaned

## Known limitations

- Complex tables with merged cells or irregular layouts may not convert perfectly
- Multi-page tables may show repeated headers (one per PDF page)
- Images and embedded graphics are not extracted
- PDFs with custom font encodings (e.g. Apple .SFNS symbol fonts) may produce garbled text — these require OCR (server version)
- OCR quality depends on scan quality and page layout
- Non-English text may produce imperfect OCR results unless Tesseract language packs are configured
- DOCX conversion is simplified via HTML-to-Markdown rules and may not preserve every style nuance
