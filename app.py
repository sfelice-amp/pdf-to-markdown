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

app = Flask(__name__, template_folder="templates")
ALLOWED_EXTENSIONS = {"pdf", "docx"}


def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def normalize_text(text: str) -> str:
    text = text.replace("ﬁ", "fi").replace("ﬂ", "fl").replace("ﬀ", "ff").replace("ﬃ", "ffi").replace("ﬄ", "ffl")
    text = text.replace("“", '"').replace("”", '"').replace("‘", "'").replace("’", "'")
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

    paragraph_text = "\n".join(lines)
    paragraph_text = re.sub(r"(?<!\n)\n(?!\n)", " ", paragraph_text)
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


def convert_pdf(path: Path) -> str:
    text = extract_pdf_text(path)
    if pdf_has_text(text):
        return text
    return ocr_pdf(path)


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
                raw_text = convert_pdf(temp_path)
            else:
                raw_text = convert_docx(temp_path)
        except Exception as exc:
            return jsonify(error=f"Conversion failed: {str(exc)}"), 500

    cleaned = normalize_text(raw_text)
    return jsonify(text=cleaned, filename=f"{Path(filename).stem}.md")


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True)
