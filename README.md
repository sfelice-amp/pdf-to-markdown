# Local PDF / DOCX to Markdown Web App

A small local web app that converts PDF and DOCX files into clean Markdown. The app uses embedded PDF text when available and falls back to Tesseract OCR for scanned or image-only PDFs.

## What it does

- Accepts PDF and DOCX uploads in the browser
- Uses `pdfminer.six` to read embedded text from PDFs first
- Falls back to OCR with `pdf2image` + `Tesseract` when needed
- Converts DOCX files using `mammoth`
- Cleans common OCR artifacts like ligatures, smart quotes, page numbers, and repeated header/footer lines
- Shows output in the browser and lets you download a `.md` file

## Setup and run

From the `local-pdf-to-markdown` folder:

```bash
./run.sh
```

Then open:

```text
http://127.0.0.1:5000
```

## Stop the server

Press `Ctrl+C` in the terminal where the app is running.

## Requirements

- macOS
- Homebrew installed
- Python 3.11+ recommended
- `Tesseract` installed via Homebrew
- `Poppler` installed via Homebrew

## Known limitations

- OCR quality depends on scan quality and page layout
- Complex tables are not preserved as Markdown tables
- Images and embedded graphics are ignored
- Non-English text may produce imperfect OCR results unless Tesseract language packs are configured separately
- DOCX conversion is simplified via HTML-to-Markdown rules and may not preserve every style nuance
