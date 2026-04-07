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


def collapse_empty_columns(rows: list[list[str]]) -> list[list[str]]:
    """Merge adjacent columns that never both have content in the same row.

    pdfplumber often creates offset column grids where headers use one set of
    columns and data rows use adjacent columns. This detects that pattern and
    merges them into single columns.
    """
    if not rows:
        return rows

    num_cols = max(len(row) for row in rows)
    # Pad all rows to same length
    padded = [row + [""] * (num_cols - len(row)) for row in rows]

    # Iteratively merge adjacent column pairs that never conflict
    changed = True
    while changed:
        changed = False
        num_cols = len(padded[0]) if padded else 0
        for col_idx in range(num_cols - 1):
            # Check if columns col_idx and col_idx+1 ever both have content
            conflict = False
            for row in padded:
                a = row[col_idx].strip()
                b = row[col_idx + 1].strip()
                if a and b:
                    conflict = True
                    break
            if not conflict:
                # Merge: take whichever has content (or empty if both empty)
                for row in padded:
                    a = row[col_idx].strip()
                    b = row[col_idx + 1].strip()
                    row[col_idx] = a if a else b
                # Remove the now-redundant column
                for row in padded:
                    del row[col_idx + 1]
                changed = True
                break  # Restart scan since indices shifted

    # Remove any columns that are still entirely empty
    if padded:
        num_cols = len(padded[0])
        keep = [
            c for c in range(num_cols)
            if any(row[c].strip() for row in padded)
        ]
        padded = [[row[c] for c in keep] for row in padded]

    return padded


def clean_cell(cell: str) -> str:
    """Clean a table cell: collapse newlines to spaces, deduplicate, strip."""
    text = str(cell or "").replace("\n", " ").replace("\r", " ")
    text = re.sub(r" {2,}", " ", text).strip()
    return deduplicate_repeated_chars(text)


def rows_to_markdown(rows: list[list[str]]) -> str:
    """Convert cleaned table rows to a markdown table string."""
    if not rows or len(rows) < 2:
        return ""

    header = rows[0]
    num_cols = len(header)
    clean_header = [clean_cell(cell) for cell in header]
    md_lines = ["| " + " | ".join(clean_header) + " |"]
    md_lines.append("| " + " | ".join(["---"] * num_cols) + " |")

    for row in rows[1:]:
        padded = row + [""] * (num_cols - len(row))
        clean = [clean_cell(cell) for cell in padded[:num_cols]]
        md_lines.append("| " + " | ".join(clean) + " |")

    return "\n".join(md_lines)


def is_page_boilerplate(text: str) -> bool:
    """Check if a text segment is just page headers/footers/boilerplate.

    Matches common patterns like page numbers, version stamps, confidentiality
    headers, and repeated document titles that appear on every page.
    """
    stripped = text.strip()
    lines = [l.strip() for l in stripped.splitlines() if l.strip()]
    if not lines:
        return True

    # Check each non-empty line — if ALL lines are boilerplate, skip the segment
    boilerplate_patterns = [
        r'Page\s+\d+\s+of\s+\d+',        # "Page 3 of 47"
        r'^v\d+\.\d+\s*[-–]\s*\d+',       # "v0.2 - 14 Mar 2026"
        r'(?i)^confidential\b',            # "CONFIDENTIAL - ..."
        r'(?i)^draft\b',                   # "DRAFT"
        r'(?i)^copyright\b',              # Copyright notices
    ]
    for line in lines:
        if not any(re.search(p, line) for p in boilerplate_patterns):
            return False
    return True


def extract_page_content(page) -> str:
    """Extract text and tables from a single pdfplumber page, interleaved by position.

    Uses table bounding boxes to extract text from non-table regions,
    then interleaves tables and text segments by their vertical position.
    """
    tables = page.find_tables()

    if not tables:
        text = page.extract_text() or ""
        return text

    # Collect content segments with their vertical position (top of bbox)
    segments = []

    # Get table bounding boxes: (x0, top, x1, bottom)
    table_bboxes = [t.bbox for t in tables]

    # Sort table bboxes by vertical position
    sorted_bboxes = sorted(table_bboxes, key=lambda b: b[1])

    page_top = 0
    page_bottom = page.height
    page_left = 0
    page_right = page.width

    # Extract text from regions between tables
    current_top = page_top
    for bbox in sorted_bboxes:
        table_top = bbox[1]
        if table_top > current_top:
            crop_box = (page_left, current_top, page_right, table_top)
            try:
                cropped = page.crop(crop_box)
                text = cropped.extract_text() or ""
                if text.strip() and not is_page_boilerplate(text):
                    segments.append((current_top, "text", text))
            except Exception:
                pass
        current_top = bbox[3]  # bottom of this table

    # Extract text below the last table
    if current_top < page_bottom:
        crop_box = (page_left, current_top, page_right, page_bottom)
        try:
            cropped = page.crop(crop_box)
            text = cropped.extract_text() or ""
            if text.strip() and not is_page_boilerplate(text):
                segments.append((current_top, "text", text))
        except Exception:
            pass

    # Extract tables, clean columns, and convert to markdown
    for table_obj, bbox in zip(tables, table_bboxes):
        rows = table_obj.extract()
        if not rows or len(rows) < 2:
            continue
        # Filter out empty rows
        rows = [row for row in rows if any(cell and str(cell).strip() for cell in row)]
        if len(rows) < 2:
            continue

        # Clean cell content
        rows = [[str(cell or "").strip() for cell in row] for row in rows]
        # Collapse empty columns
        rows = collapse_empty_columns(rows)

        md = rows_to_markdown(rows)
        if md:
            segments.append((bbox[1], "table", md))

    # Sort by vertical position and join
    segments.sort(key=lambda s: s[0])

    parts = []
    for _, kind, content in segments:
        parts.append(content)

    return "\n\n".join(parts)


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
    # Remove page footer/header boilerplate lines
    text = re.sub(r"(?m)^.*Page\s+\d+\s+of\s+\d+\s*$", "", text)
    text = re.sub(r"(?m)^\s*v\d+\.\d+\s*[-–]\s*\d+.*$", "", text)
    # Strip repeated confidentiality/document headers that prefix paragraphs
    text = re.sub(r"(?m)^(?:CONFIDENTIAL\s*[-–]\s*\S+(?:\s+\S+){0,3}\s+)", "", text)

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


def detect_repetition_factor(text: str) -> int:
    """Detect if text has systematic character repetition (e.g., every char repeated 4x).

    Returns the dominant repetition factor (>= 3), or 0 if no pattern found.
    """
    if len(text) < 10:
        return 0

    # Count run lengths
    from collections import Counter
    run_lengths = Counter()
    i = 0
    while i < len(text):
        char = text[i]
        count = 1
        while i + count < len(text) and text[i + count] == char:
            count += 1
        if count >= 3 and char != ' ':
            run_lengths[count] += 1
        i += count

    if not run_lengths:
        return 0

    # Find the most common run length >= 3
    most_common = run_lengths.most_common(1)[0]
    factor, freq = most_common

    # Only use if it appears frequently enough (at least 3 occurrences)
    if freq >= 3 and factor >= 3:
        return factor
    return 0


def deduplicate_repeated_chars(text: str) -> str:
    """Remove repeated patterns that PDFs sometimes produce in extracted text.

    Handles substring repetition at the start of a string (e.g., "To:To:To:" -> "To:")
    and collapses character runs. If systematic repetition is detected (e.g., every
    character repeated 4x), divides run lengths by the factor to preserve legitimate
    doubles (like "ss" in "Business").
    """
    if not text or len(text) < 2:
        return text

    # Remove substring repetitions at the START of the string
    # (e.g., "To:To:To:To: rest of string" -> "To: rest of string")
    # Start at length 2 — single-char runs are handled by the character collapsing below
    for pattern_len in range(2, min(len(text) // 2 + 1, 21)):
        pattern = text[:pattern_len]
        match_count = 0
        pos = 0
        while pos + pattern_len <= len(text) and text[pos:pos + pattern_len] == pattern:
            match_count += 1
            pos += pattern_len
        if match_count >= 3:
            return pattern + text[pos:]

    # Detect systematic repetition factor
    factor = detect_repetition_factor(text)

    result = []
    i = 0
    while i < len(text):
        char = text[i]
        count = 1
        while i + count < len(text) and text[i + count] == char:
            count += 1

        if factor and count >= factor:
            # Divide by the repetition factor, rounding to nearest
            target = max(1, round(count / factor))
            result.extend([char] * target)
        elif count >= 3 and not char.isalnum() and char != ' ':
            result.append(char)  # 3+ punctuation -> 1
        else:
            result.extend([char] * count)
        i += count
    return ''.join(result)


def extract_pdf_text(path: Path) -> str:
    """Extract text and tables from a PDF using pdfplumber page by page.

    Tables are placed where they appear on each page, with surrounding text
    extracted from non-table regions. Falls back to pdfminer if pdfplumber fails.
    """
    try:
        page_contents = []
        with pdfplumber.open(str(path)) as pdf:
            for page in pdf.pages:
                content = extract_page_content(page)
                if content.strip():
                    page_contents.append(content)

        text = "\n\n".join(page_contents)
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
