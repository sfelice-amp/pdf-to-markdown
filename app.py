import os
import re
import tempfile
from pathlib import Path

from flask import Flask, jsonify, render_template, request
from werkzeug.utils import secure_filename
import mammoth
from pdfminer.high_level import extract_text
from pdf2image import convert_from_path
import pytesseract

BASE_DIR = Path(__file__).resolve().parent
TEMPLATES_DIR = BASE_DIR / "templates"
app = Flask(__name__, template_folder=str(TEMPLATES_DIR))
ALLOWED_EXTENSIONS = {"pdf", "docx"}


def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def split_table_line(line: str) -> list[str]:
    cells = re.split(r"\t| {2,}", line.strip())
    return [cell.strip() for cell in cells if cell.strip()]


def is_table_line(line: str) -> bool:
    return len(split_table_line(line)) >= 2 and re.search(r"\t| {2,}", line) is not None


def is_table_block(lines: list[str]) -> bool:
    if len(lines) < 2:
        return False

    rows = [split_table_line(line) for line in lines]
    if any(len(row) < 2 for row in rows):
        return False

    return len({len(row) for row in rows}) == 1


def format_markdown_table(lines: list[str]) -> list[str]:
    rows = [split_table_line(line) for line in lines]
    col_count = len(rows[0])
    header = rows[0] + [""] * (col_count - len(rows[0]))
    separator = ["---"] * col_count

    table_lines = ["| " + " | ".join(header) + " |", "| " + " | ".join(separator) + " |"]
    for row in rows[1:]:
        padded = row + [""] * (col_count - len(row))
        table_lines.append("| " + " | ".join(padded) + " |")

    return table_lines


def normalize_text(text: str) -> str:
    text = text.replace("ﬁ", "fi").replace("ﬂ", "fl").replace("ﬀ", "ff").replace("ﬃ", "ffi").replace("ﬄ", "ffl")
    text = text.replace("“", '"').replace("”", '"').replace("‘", "'" ).replace("’", "'")
    text = text.replace("–", "-").replace("—", "-").replace("…", "...")

    text = re.sub(r"(\w)-\n(\w)", r"\1\2", text)
    text = re.sub(r"\r\n?|\r", "\n", text)
    text = re.sub(r"(?m)^\s*\d+\s*$", "", text)

    lines = []
    previous = None
    for raw_line in text.splitlines():
        line = raw_line.rstrip()
        if line == previous:
            continue
        lines.append(line)
        previous = line

    blocks: list[list[str] | str] = []
    buffer: list[str] = []

    for line in lines:
        if line.strip() == "":
            if buffer:
                blocks.append(buffer)
                buffer = []
            blocks.append("")
        else:
            buffer.append(line)

    if buffer:
        blocks.append(buffer)

    formatted_blocks: list[str] = []
    for block in blocks:
        if block == "":
            formatted_blocks.append("")
            continue

        if is_table_block(block):
            formatted_blocks.extend(format_markdown_table(block))
        else:
            merged = " ".join(line for line in block if line.strip())
            formatted_blocks.append(merged)

    paragraph_text = "\n\n".join(formatted_blocks)
    paragraph_text = re.sub(r"\n{3,}", "\n\n", paragraph_text)

    return paragraph_text.strip()


def pdf_has_text(text: str) -> bool:
    words = re.sub(r"[^A-Za-z0-9]+", "", text)
    return len(words) >= 30


def extract_pdf_text(path: Path) -> str:
    try:
        text = extract_text(str(path)) or ""
    except Exception:
        return ""
    return text


def ocr_pdf(path: Path) -> str:
    pages = convert_from_path(str(path), dpi=300)
    page_texts = []
    for page in pages:
        page_texts.append(pytesseract.image_to_string(page, lang="eng"))
    return "\n\n".join(page_texts)


def html_to_markdown(html: str) -> str:
    html = html.replace("\r", "")
    html = re.sub(r"<(/?)(p|div|section)[^>]*>", "\n\n", html)
    html = re.sub(r"<br\s*/?>", "\n", html)
    html = re.sub(r"<h([1-6])[^>]*>(.*?)</h\1>", lambda m: "#" * int(m.group(1)) + " " + m.group(2).strip() + "\n\n", html, flags=re.S)
    html = re.sub(r"<li[^>]*>(.*?)</li>", r"- \1\n", html, flags=re.S)
    html = re.sub(r"<a[^>]*href=[\'\"]([^\'\"]+)[\'\"][^>]*>(.*?)</a>", r"[\2](\1)", html, flags=re.S)
    html = re.sub(r"<(strong|b)>", "**", html)
    html = re.sub(r"</(strong|b)>", "**", html)
    html = re.sub(r"<(em|i)>", "*", html)
    html = re.sub(r"</(em|i)>", "*", html)
    html = re.sub(r"<[^>]+>", "", html)
    html = re.sub(r"\n{3,}", "\n\n", html)
    return html.strip()


def convert_docx(path: Path) -> str:
    with path.open("rb") as docx_file:
        result = mammoth.convert_to_html(docx_file)
        markdown = html_to_markdown(result.value)
    return markdown


def convert_pdf(path: Path) -> tuple[str, str]:
    text = extract_pdf_text(path)
    if pdf_has_text(text):
        return text, "Embedded text"
    return ocr_pdf(path), "OCR"


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/convert", methods=["POST"])
def convert_file():
    if "file" not in request.files:
        return jsonify(error="No file uploaded."), 400

    uploaded_file = request.files["file"]
    if uploaded_file.filename == "":
        return jsonify(error="No file selected."), 400

    if not allowed_file(uploaded_file.filename):
        return jsonify(error="Unsupported file type. Use PDF or DOCX."), 400

    filename = secure_filename(uploaded_file.filename)
    suffix = Path(filename).suffix.lower()

    with tempfile.TemporaryDirectory() as tempdir:
        temp_path = Path(tempdir) / filename
        uploaded_file.save(temp_path)

        try:
            if suffix == ".pdf":
                raw_text, source_type = convert_pdf(temp_path)
            else:
                raw_text = convert_docx(temp_path)
                source_type = "DOCX"
        except Exception as exc:
            return jsonify(error=f"Conversion failed: {str(exc)}"), 500

    cleaned = normalize_text(raw_text)
    return jsonify(text=cleaned, filename=f"{Path(filename).stem}.md", source_type=source_type)


if __name__ == "__main__":
    import sys

    port = 5000
    if len(sys.argv) > 1:
        try:
            port = int(sys.argv[1])
        except ValueError:
            pass

    app.run(host="127.0.0.1", port=port, debug=True)
