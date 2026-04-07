# Local PDF / DOCX to Markdown Web App

A small local web app that converts PDF and DOCX files into clean Markdown. The app uses embedded PDF text when available and falls back to Tesseract OCR for scanned or image-only PDFs. Features a live preview of rendered Markdown for easy review.

## What it does

- Accepts PDF and DOCX uploads via drag-and-drop or file selection
- Validates file types and sizes (up to 50MB)
- Uses `pdfminer.six` to read embedded text from PDFs first
- Falls back to OCR with `pdf2image` + `Tesseract` when needed
- Converts DOCX files using `mammoth`
- Cleans common OCR artifacts like ligatures, smart quotes, page numbers, and repeated header/footer lines
- Detects and formats simple tables as Markdown tables
- Shows conversion progress with loading indicators
- Provides feedback on conversion method (text vs. OCR) with warnings for OCR usage
- Displays output in raw Markdown or rendered preview
- Allows copying to clipboard or downloading as `.md` file

## Setup and run

From the `pdf-to-markdown` folder:

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
- Python 3.9+ 
- `Tesseract` installed via Homebrew (`brew install tesseract`)
- `Poppler` installed via Homebrew (`brew install poppler`)

## Known limitations

- OCR quality depends on scan quality and page layout
- Complex tables with merged cells or irregular layouts may not convert perfectly
- Images and embedded graphics are ignored
- Non-English text may produce imperfect OCR results unless Tesseract language packs are configured separately
- DOCX conversion is simplified via HTML-to-Markdown rules and may not preserve every style nuance
