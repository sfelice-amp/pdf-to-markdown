import os
import re
import tempfile
from pathlib import Path

from flask import Flask, jsonify, render_template, request
from werkzeug.utils import secure_filename
import mammoth
import pdfplumber
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
    # Split on tabs or 2+ spaces, but also handle cases with single spaces if aligned
    cells = re.split(r'\t| {2,}', line.strip())
    # Filter out empty cells and strip
    cells = [cell.strip() for cell in cells if cell.strip()]
    return cells


def is_table_line(line: str) -> bool:
    cells = split_table_line(line)
    return len(cells) >= 2


def is_table_block(lines: list[str]) -> bool:
    if len(lines) < 2:
        return False

    rows = [split_table_line(line) for line in lines]
    # Count column lengths
    col_counts = [len(row) for row in rows if len(row) >= 2]
    if len(col_counts) < len(rows) * 0.5:  # At least 50% of lines must be table-like
        return False

    # Find the most common column count
    from collections import Counter
    most_common = Counter(col_counts).most_common(1)
    if not most_common:
        return False
    target_cols = most_common[0][0]
    # Allow rows with target_cols or target_cols-1 (for merged cells)
    valid_rows = [row for row in rows if len(row) in (target_cols, target_cols - 1)]
    return len(valid_rows) >= len(rows) * 0.8  # At least 80% valid


def format_markdown_table(lines: list[str]) -> list[str]:
    rows = [split_table_line(line) for line in lines]
    # Filter to valid rows
    col_counts = [len(row) for row in rows]
    from collections import Counter
    target_cols = Counter(col_counts).most_common(1)[0][0]
    valid_rows = [row for row in rows if len(row) in (target_cols, target_cols - 1)]
    
    if not valid_rows:
        return lines  # Fallback

    # Pad rows to target_cols
    padded_rows = []
    for row in valid_rows:
        if len(row) == target_cols - 1:
            row.append("")  # Add empty cell for merged
        padded_rows.append(row[:target_cols])  # Truncate if more

    # Assume first row is header if it looks like it (no numbers, short)
    first_row = padded_rows[0]
    is_header = not any(re.search(r'\d', cell) for cell in first_row) and all(len(cell) < 50 for cell in first_row)
    
    table_lines = []
    if is_header:
        header = first_row
        separator = ["---"] * len(header)
        table_lines.append("| " + " | ".join(header) + " |")
        table_lines.append("| " + " | ".join(separator) + " |")
        data_rows = padded_rows[1:]
    else:
        data_rows = padded_rows

    for row in data_rows:
        table_lines.append("| " + " | ".join(row) + " |")

    return table_lines


def normalize_text(text: str) -> str:
    # Clean up repeated characters first (except in tables)
    lines_raw = text.split('\n')
    lines_cleaned = []
    for line in lines_raw:
        # Only deduplicate non-table lines
        if not line.strip().startswith('|'):
            line = deduplicate_repeated_chars(line)
        lines_cleaned.append(line)
    text = '\n'.join(lines_cleaned)
    
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
        # Preserve markdown table lines as-is
        elif line.strip().startswith("|"):
            if buffer:
                blocks.append(buffer)
                buffer = []
            blocks.append(line)
        else:
            buffer.append(line)

    if buffer:
        blocks.append(buffer)

    formatted_blocks: list[str] = []
    for block in blocks:
        if block == "":
            formatted_blocks.append("")
            continue

        # If it's a markdown table line (already extracted), keep it as-is
        if isinstance(block, str) and block.strip().startswith("|"):
            formatted_blocks.append(block)
            continue

        if is_table_block(block):
            formatted_blocks.extend(format_markdown_table(block))
        else:
            merged = " ".join(line for line in block if line.strip())
            formatted_blocks.append(merged)

    paragraph_text = ""
    for i, block in enumerate(formatted_blocks):
        if i == 0:
            paragraph_text = block
        elif block.strip().startswith("|"):
            # Table rows: join with single newline, no blank line before
            paragraph_text += "\n" + block
        else:
            # Regular blocks: join with double newline (blank line)
            paragraph_text += "\n\n" + block
    
    paragraph_text = re.sub(r"\n{3,}", "\n\n", paragraph_text)

    return paragraph_text.strip()


def pdf_has_text(text: str) -> bool:
    words = re.sub(r"[^A-Za-z0-9]+", "", text)
    return len(words) >= 30


def deduplicate_repeated_chars(text: str) -> str:
    """Remove repeated patterns and characters that PDFs sometimes include.
    Handles both character repetition (MM -> M) and substring repetition (To:To: -> To:).
    Applies multiple passes until no more changes."""
    if not text or len(text) < 2:
        return text
    
    # Apply deduplication multiple times until stable
    prev_result = text
    while True:
        result = _deduplicate_once(prev_result)
        if result == prev_result:
            break
        prev_result = result
    return result


def _deduplicate_once(text: str) -> str:
    """Single pass of deduplication."""
    if not text or len(text) < 2:
        return text
    
    # First, try to remove substring repetitions at the START of the string
    # (e.g., "To:To:To:To: rest of string" -> "To: rest of string")
    for pattern_len in range(1, len(text) // 2 + 1):
        if pattern_len > 20:  # Don't check for very long patterns
            break
        pattern = text[:pattern_len]
        # Count how many times pattern repeats from start
        match_count = 0
        pos = 0
        while pos + pattern_len <= len(text) and text[pos:pos+pattern_len] == pattern:
            match_count += 1
            pos += pattern_len
        
        # If pattern repeats at least 2 times at the start
        if match_count >= 2:
            # Return the pattern plus the remainder
            remainder = text[pos:]
            return pattern + remainder
    
    # Then handle character repetition (runs of 3+ identical chars)
    result = []
    i = 0
    while i < len(text):
        char = text[i]
        count = 1
        # Count consecutive identical characters
        while i + count < len(text) and text[i + count] == char:
            count += 1
        # Keep only one copy if count > 2; otherwise keep all
        if count > 2:
            result.append(char)
        else:
            result.extend([char] * count)
        i += count
    return ''.join(result)


def extract_tables_from_pdf(path: Path) -> list[list[list[str]]]:
    """Extract tables from PDF using pdfplumber for better structure preservation."""
    tables = []
    try:
        with pdfplumber.open(str(path)) as pdf:
            for page in pdf.pages:
                page_tables = page.extract_tables()
                if page_tables:
                    tables.extend(page_tables)
    except Exception:
        pass
    return tables


def tables_to_markdown(tables: list[list[list[str]]]) -> str:
    """Convert extracted tables to markdown format."""
    if not tables:
        return ""
    
    markdown_tables = []
    for table in tables:
        if not table or len(table) < 2:
            continue
        
        # Filter out empty rows
        table = [row for row in table if any(cell and str(cell).strip() for cell in row)]
        if len(table) < 2:
            continue
        
        header = table[0]
        # Clean up repeated characters and format header
        clean_header = [deduplicate_repeated_chars(str(cell or "").strip()) for cell in header]
        markdown_rows = ["| " + " | ".join(clean_header) + " |"]
        markdown_rows.append("| " + " | ".join(["---"] * len(header)) + " |")
        
        for row in table[1:]:
            padded_row = row + [""] * (len(header) - len(row))
            clean_row = [deduplicate_repeated_chars(str(cell or "").strip()) for cell in padded_row[:len(header)]]
            markdown_rows.append("| " + " | ".join(clean_row) + " |")
        
        markdown_tables.append("\n".join(markdown_rows))
    
    return "\n\n".join(markdown_tables)


def extract_pdf_text(path: Path) -> str:
    try:
        # First try to extract tables using pdfplumber for better structure
        tables = extract_tables_from_pdf(path)
        text = extract_text(str(path)) or ""
        
        # If tables were found, insert them at a reasonable point in the text
        if tables:
            markdown_tables = tables_to_markdown(tables)
            # Insert tables after the first substantial content block
            lines = text.split('\n\n')
            if len(lines) > 2:
                text = '\n\n'.join(lines[:2]) + '\n\n' + markdown_tables + '\n\n' + '\n\n'.join(lines[2:])
            else:
                text = text + '\n\n' + markdown_tables
    except Exception:
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
