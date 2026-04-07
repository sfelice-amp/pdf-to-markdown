from __future__ import annotations

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


# ---------------------------------------------------------------------------
# Font-aware line extraction for heading / paragraph detection
# ---------------------------------------------------------------------------

def _dedup_overlapping_chars(line_chars: list[dict]) -> tuple[list[dict], list[bool]]:
    """Remove overlapping chars (same position) and track which are bold-via-overlap.

    Some PDFs render bold by printing each character multiple times at the same
    x-position. Returns deduplicated chars and a per-char bold flag.
    """
    if not line_chars:
        return [], []

    deduped: list[dict] = []
    bold_flags: list[bool] = []
    i = 0
    while i < len(line_chars):
        c = line_chars[i]
        # Count how many subsequent chars overlap at same x (within 0.5pt)
        overlap_count = 1
        while (i + overlap_count < len(line_chars)
               and abs(line_chars[i + overlap_count]['x0'] - c['x0']) < 0.5
               and line_chars[i + overlap_count]['text'] == c['text']):
            overlap_count += 1
        deduped.append(c)
        bold_flags.append(overlap_count >= 2)
        i += overlap_count

    return deduped, bold_flags


def extract_lines_with_metadata(chars: list[dict], bbox: tuple | None = None) -> list[dict]:
    """Group page.chars into lines with font metadata.

    Each returned dict has keys: text, top, spans, dominant_font, dominant_size,
    is_bold, is_italic, x0 (left position), has_overlap_bold.
    Detects bold-via-font and bold-via-overlapping-chars (common in email PDFs).
    """
    if bbox:
        x0, y0, x1, y1 = bbox
        chars = [c for c in chars if x0 <= c['x0'] <= x1 and y0 <= c['top'] <= y1]

    if not chars:
        return []

    # Group chars into lines by y-position (cluster within 2pt)
    sorted_chars = sorted(chars, key=lambda c: (c['top'], c['x0']))
    lines: list[list[dict]] = []
    current_line: list[dict] = []
    current_y: float | None = None

    for c in sorted_chars:
        if current_y is None or abs(c['top'] - current_y) <= 2.0:
            current_line.append(c)
            if current_y is None:
                current_y = c['top']
        else:
            if current_line:
                lines.append(current_line)
            current_line = [c]
            current_y = c['top']
    if current_line:
        lines.append(current_line)

    result = []
    for line_chars in lines:
        line_chars.sort(key=lambda c: c['x0'])

        # Deduplicate overlapping chars and detect bold-via-overlap
        deduped, bold_flags = _dedup_overlapping_chars(line_chars)

        text = ''.join(c['text'] for c in deduped).strip()
        # Also apply character-run deduplication (e.g. remaining "SSSS" → "S")
        text = deduplicate_repeated_chars(text)
        if not text:
            continue

        # Track how much of the line is bold-via-overlap
        non_space_bold = sum(1 for c, b in zip(deduped, bold_flags)
                            if b and c['text'].strip())
        non_space_total = sum(1 for c in deduped if c['text'].strip())
        overlap_bold_ratio = non_space_bold / non_space_total if non_space_total else 0

        # Build font spans (from deduped chars) with bold-via-overlap info
        spans = []
        current_span_font = None
        current_span_size = None
        current_span_bold_overlap = None
        current_span_text: list[str] = []
        for c, is_overlap_bold in zip(deduped, bold_flags):
            fn = c.get('fontname', '')
            sz = round(c['size'], 1)
            if (fn != current_span_font or sz != current_span_size
                    or is_overlap_bold != current_span_bold_overlap):
                if current_span_text:
                    raw = ''.join(current_span_text)
                    clean = deduplicate_repeated_chars(raw)
                    effective_bold = (current_span_bold_overlap
                                     or 'Bold' in (current_span_font or '')
                                     or 'bold' in (current_span_font or ''))
                    spans.append((current_span_font, current_span_size,
                                  clean, effective_bold))
                current_span_font = fn
                current_span_size = sz
                current_span_bold_overlap = is_overlap_bold
                current_span_text = [c['text']]
            else:
                current_span_text.append(c['text'])
        if current_span_text:
            raw = ''.join(current_span_text)
            clean = deduplicate_repeated_chars(raw)
            effective_bold = (current_span_bold_overlap
                              or 'Bold' in (current_span_font or '')
                              or 'bold' in (current_span_font or ''))
            spans.append((current_span_font, current_span_size,
                          clean, effective_bold))

        # Determine dominant font (by char count, ignoring whitespace)
        from collections import Counter
        font_counts: Counter = Counter()
        for fn, sz, span_text, _ in spans:
            count = len(span_text.replace(' ', ''))
            if count > 0:
                font_counts[(fn, sz)] += count
        if not font_counts:
            continue
        dominant_font, dominant_size = font_counts.most_common(1)[0][0]

        is_bold_font = 'Bold' in dominant_font or 'bold' in dominant_font
        is_bold = is_bold_font or overlap_bold_ratio > 0.5
        is_italic = 'Italic' in dominant_font or 'italic' in dominant_font

        # Record left margin position for bullet detection
        line_x0 = deduped[0]['x0'] if deduped else 0

        result.append({
            'text': text,
            'top': line_chars[0]['top'],
            'x0': line_x0,
            'spans': spans,
            'dominant_font': dominant_font,
            'dominant_size': dominant_size,
            'is_bold': is_bold,
            'is_italic': is_italic,
            'has_overlap_bold': overlap_bold_ratio > 0.5,
        })

    return result


def classify_and_build_markdown(lines: list[dict]) -> str:
    """Classify lines as title/heading/body/boilerplate and build markdown.

    Uses font statistics to determine the body font, then classifies by
    relative size, font family differences, and bold-via-overlap detection.
    Detects paragraph breaks from vertical gaps, and bullet lists from
    x-position indentation.
    """
    if not lines:
        return ""

    from collections import Counter
    import statistics

    # Determine body font: most frequent (font, size) by char count
    font_char_counts: Counter = Counter()
    for line in lines:
        count = len(line['text'].replace(' ', ''))
        font_char_counts[(line['dominant_font'], line['dominant_size'])] += count

    body_font, body_size = font_char_counts.most_common(1)[0][0]
    body_is_serif = 'Serif' in body_font or 'Times' in body_font or 'serif' in body_font

    # Calculate median line gap for body text (for paragraph break detection)
    body_gaps = []
    prev_top = None
    for line in lines:
        if abs(line['dominant_size'] - body_size) < 0.5:
            if prev_top is not None:
                gap = line['top'] - prev_top
                if 0 < gap < 100:  # sanity bound
                    body_gaps.append(gap)
            prev_top = line['top']

    median_gap = statistics.median(body_gaps) if body_gaps else 14.0
    para_break_threshold = median_gap * 1.8

    # Determine left margin for bullet detection.
    # Use the leftmost x0 position that has at least 2 lines at body-text size.
    # This avoids page headers/footers at unusual positions skewing the margin,
    # while correctly identifying the body left edge even when bullet items dominate.
    body_line_x0s: list[float] = []
    for line in lines:
        if abs(line['dominant_size'] - body_size) < 1.0 and line['dominant_size'] > 1.0:
            body_line_x0s.append(round(line['x0'], 0))
    x0_counts: Counter = Counter(body_line_x0s)
    # Take the leftmost x0 with at least 2 occurrences
    qualified = sorted(x for x, cnt in x0_counts.items() if cnt >= 2)
    left_margin = qualified[0] if qualified else (min(body_line_x0s) if body_line_x0s else 0)
    bullet_indent_threshold = 10  # x0 must be >=10pt right of margin

    # Classify each line
    output_parts: list[str] = []
    prev_top = None

    for line in lines:
        size = line['dominant_size']
        font = line['dominant_font']
        text = line['text']
        is_bold = line['is_bold']
        is_sans = 'Sans' in font or 'Helvetica' in font or 'Arial' in font
        line_x0 = line.get('x0', 0)

        # Skip boilerplate (page headers/footers in small font)
        if size < body_size * 0.85:
            prev_top = line['top']
            continue

        # Skip tiny invisible text
        if size <= 1.0:
            prev_top = line['top']
            continue

        # Skip common page header/footer patterns
        stripped_text = text.strip()
        if re.match(r'^https?://', stripped_text):
            prev_top = line['top']
            continue
        if re.match(r'^\[.*preprint\]$', stripped_text, re.IGNORECASE):
            prev_top = line['top']
            continue
        # Skip repeated page headers like "JMIR Preprints Lee et al"
        if re.match(r'^.{0,30}(Preprints?|preprints?).{0,30}et\s+al', stripped_text):
            prev_top = line['top']
            continue
        # Skip page numbers like "1 of 2", "Page 3 of 10"
        if re.match(r'^(?:Page\s+)?\d+\s+of\s+\d+$', stripped_text, re.IGNORECASE):
            prev_top = line['top']
            continue

        # Detect paragraph break from vertical gap
        if prev_top is not None:
            gap = line['top'] - prev_top
            if gap > para_break_threshold:
                output_parts.append('')  # blank line = paragraph break

        # Check if this line is indented (potential bullet)
        is_indented = (line_x0 - left_margin) >= bullet_indent_threshold

        # Classify heading level
        size_ratio = size / body_size if body_size > 0 else 1.0

        if size_ratio > 1.35:
            # Title
            output_parts.append(f'# {text}')
        elif size_ratio > 1.1 or (is_bold and is_sans and body_is_serif):
            # Section heading
            output_parts.append(f'## {text}')
        elif is_indented:
            # Indented line — bullet point (may be bold or regular)
            formatted = _format_inline_spans(line['spans'], body_font, body_size)
            output_parts.append(f'- {formatted}')
        elif is_bold and len(text) < 80:
            # Check if this is a standalone bold line (potential subheading)
            all_bold = all(s[3] for s in line['spans'] if s[2].strip())
            if all_bold and len(text.split()) <= 10:
                output_parts.append(f'### {text}')
            else:
                # Bold body text — wrap in **
                formatted = _format_inline_spans(line['spans'], body_font, body_size)
                output_parts.append(formatted)
        else:
            # Regular body text — apply inline bold/italic
            formatted = _format_inline_spans(line['spans'], body_font, body_size)
            output_parts.append(formatted)

        prev_top = line['top']

    # Post-processing: merge multi-line headings and bullet continuations
    merged = []
    for part in output_parts:
        if not merged:
            merged.append(part)
            continue

        # Merge consecutive heading lines at the same level (multi-line titles)
        if part.startswith('#'):
            match_cur = re.match(r'^(#{1,6})\s+(.*)', part)
            match_prev = re.match(r'^(#{1,6})\s+(.*)', merged[-1])
            if match_cur and match_prev and match_cur.group(1) == match_prev.group(1):
                merged[-1] = f'{match_prev.group(1)} {match_prev.group(2)} {match_cur.group(2)}'
                continue

        # Merge bullet continuation lines: if previous was a bullet and this is
        # also a bullet but doesn't start with a bold label, it's a wrapped line
        if (part.startswith('- ') and merged[-1].startswith('- ')
                and not re.match(r'^- \*\*', part)
                and not re.match(r'^- \d', part)):
            # Continuation of previous bullet
            merged[-1] = f'{merged[-1]} {part[2:]}'
            continue

        merged.append(part)

    return '\n'.join(merged)


def _format_inline_spans(spans: list[tuple], body_font: str, body_size: float) -> str:
    """Format a line's font spans into markdown with **bold** and *italic*.

    Spans are 4-tuples: (fontname, size, text, is_bold).
    """
    parts = []
    for font, size, text, span_bold in spans:
        if not text.strip():
            parts.append(text)
            continue
        is_bold = span_bold
        is_italic = 'Italic' in font or 'italic' in font
        # Only mark as bold/italic if different from body font
        body_is_bold = 'Bold' in body_font or 'bold' in body_font
        body_is_italic = 'Italic' in body_font or 'italic' in body_font

        if is_bold and is_italic and not (body_is_bold and body_is_italic):
            parts.append(f'***{text.strip()}***')
        elif is_bold and not body_is_bold:
            parts.append(f'**{text.strip()}**')
        elif is_italic and not body_is_italic:
            parts.append(f'*{text.strip()}*')
        else:
            parts.append(text)

    return ''.join(parts).strip()


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


def _extract_text_region(page, bbox: tuple | None = None) -> str:
    """Extract text from a page region using char-level font analysis.

    Returns markdown with headings, paragraph breaks, and inline bold/italic.
    Falls back to page.extract_text() if char analysis yields nothing.
    """
    all_chars = page.chars
    if bbox:
        x0, y0, x1, y1 = bbox
        region_chars = [c for c in all_chars
                        if c['x0'] >= x0 - 1 and c['x0'] <= x1 + 1
                        and c['top'] >= y0 - 1 and c['top'] <= y1 + 1]
    else:
        region_chars = all_chars

    lines = extract_lines_with_metadata(region_chars)
    if lines:
        return classify_and_build_markdown(lines)

    # Fallback
    if bbox:
        try:
            cropped = page.crop(bbox)
            return cropped.extract_text() or ""
        except Exception:
            return ""
    return page.extract_text() or ""


def extract_page_content(page) -> str:
    """Extract text and tables from a single pdfplumber page, interleaved by position.

    Uses char-level font analysis for text regions (headings, paragraphs, bold/italic)
    and table detection for structured data. Content is interleaved by vertical position.
    """
    tables = page.find_tables()

    if not tables:
        return _extract_text_region(page)

    # Collect content segments with their vertical position (top of bbox)
    segments = []

    # Get table bounding boxes: (x0, top, x1, bottom)
    table_bboxes = [t.bbox for t in tables]
    sorted_bboxes = sorted(table_bboxes, key=lambda b: b[1])

    page_top = 0
    page_bottom = page.height
    page_left = 0
    page_right = page.width

    # Extract text from regions between tables using char-level analysis
    current_top = page_top
    for bbox in sorted_bboxes:
        table_top = bbox[1]
        if table_top > current_top:
            region_bbox = (page_left, current_top, page_right, table_top)
            text = _extract_text_region(page, region_bbox)
            if text.strip():
                segments.append((current_top, "text", text))
        current_top = bbox[3]  # bottom of this table

    # Extract text below the last table
    if current_top < page_bottom:
        region_bbox = (page_left, current_top, page_right, page_bottom)
        text = _extract_text_region(page, region_bbox)
        if text.strip():
            segments.append((current_top, "text", text))

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
        # Only deduplicate non-table and non-heading lines
        if not line.strip().startswith('|') and not re.match(r'^#{1,6}\s', line.strip()):
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
        stripped = line.strip()
        if stripped == "":
            if buffer:
                blocks.append(buffer)
                buffer = []
            blocks.append("")
        # Preserve markdown table lines as-is
        elif stripped.startswith("|"):
            if buffer:
                blocks.append(buffer)
                buffer = []
            blocks.append(line)
        # Preserve markdown heading lines as standalone blocks
        elif re.match(r'^#{1,6}\s', stripped):
            if buffer:
                blocks.append(buffer)
                buffer = []
            blocks.append(line)
        # Preserve markdown bullet lines as standalone blocks
        elif stripped.startswith('- '):
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

        # If it's a standalone string (table line or heading), keep as-is
        if isinstance(block, str):
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
