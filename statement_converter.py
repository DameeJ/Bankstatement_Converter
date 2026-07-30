"""
Bank Statement PDF -> Excel Converter (v3)
============================================
Key difference from v2: this version does NOT rename or reorder columns.
Whatever labels and order your bank uses in the PDF (TRANS DATE, VALUE DATE,
NARRATION, CHQ NO, DEBIT, CREDIT, BALANCE -- or any other bank's own set)
is exactly what appears in the cleaned Excel output, in the same order,
including columns that are empty for every row (they stay as empty columns,
not deleted).

Internally the script still needs to KNOW which column is "the date column"
or "the debit column" to validate/repair/sort -- it detects that via
keyword matching (HEADER_KEYWORDS below), same as before. That detection
is invisible to the output: it's only used for logic, never for display.

Pipeline:
 1. Detect text-based vs scanned PDF.
 2. Extract every row from the PDF's table(s) (pdfplumber for text-based,
    OCR for scanned), with watermark/noise character filtering.
 3. RAW sheet: export everything except obvious footer junk (address,
    phone numbers, social links, "Download App" etc) -- the account
    summary block (Opening Balance, Total Debit, etc.) is KEPT here.
 4. Find the real transaction table header row and map each column index
    to a semantic role (Date / Description / Debit / Credit / Balance /
    Value Date / Reference) purely for internal logic. The header text
    itself is preserved verbatim for the column name.
 5. Validate & repair rows using the role map (spilled amounts, spilled
    dates, stray word-fragments in a reference/cheque column).
 6. Add a derived "Reference Number" column: pulls the longest run of
    5+ digits out of the description text (transaction reference codes
    banks embed in narration), blank if none found. This is additive --
    delete it if you don't need it.
 7. CLEANED sheet(s): no summary block, starts at the header row, same
    column names/order as the bank statement, plus Reference Number.
 8. Small statements -> one sheet. Large statements -> split by month
    for readability (never affects parsing, see MONTH_SPLIT_THRESHOLD).

Usage:
    python statement_converter.py input.pdf output.xlsx
"""

import sys
import os
import re
import time
import tempfile
import pdfplumber
import pandas as pd
from datetime import datetime
from dateutil import parser as dateutil_parser
from collections import Counter
from openpyxl.utils import get_column_letter

# ----------------------------------------------------------------------
# CONFIG
# ----------------------------------------------------------------------

# Pure display/organisation choice, applied after all parsing -- doesn't
# affect extraction accuracy or speed either way. Change freely.
MONTH_SPLIT_THRESHOLD = 100

# Keyword -> semantic role, used ONLY to detect which column is which for
# internal validation/repair logic. Never used to rename output columns.
HEADER_KEYWORDS = {
    "Date":        ["date", "trans date", "transaction date", "posting date"],
    "Description": ["transaction details", "narration", "description", "particulars", "details"],
    "Reference":   ["reference", "refence", "ref no", "ref", "cheque no", "chq no", "chq"],
    "Value Date":  ["value date", "val date"],
    "Debit":       ["withdrawal", "withdrawals", "debit", "dr"],
    "Credit":      ["lodgement", "lodgements", "deposit", "credit", "cr"],
    "Balance":     ["balance", "running balance", "closing balance"],
}

def _normalize(s):
    return re.sub(r"\s+", "", s.lower())

# (canonical, normalized_keyword, original_keyword) triples, longest
# normalized keyword first, so "valuedate" matches Value Date before the
# shorter "date" keyword can wrongly claim it. original_keyword (with its
# spaces intact) is kept so we can reconstruct a lost space in the display
# header, e.g. "TRANSDATE" -> "TRANS DATE".
_HEADER_MATCH_ORDER = sorted(
    ((canonical, _normalize(kw), kw) for canonical, kws in HEADER_KEYWORDS.items() for kw in kws),
    key=lambda triple: len(triple[1]),
    reverse=True,
)

DATE_FORMATS = ["%d-%b-%Y", "%d-%b-%y", "%d/%m/%Y", "%d-%m-%Y", "%d %b %Y", "%Y-%m-%d"]
AMOUNT_RE = re.compile(r"\(?-?\d{1,3}(?:,\d{3})*(?:\.\d{2})?\)?")
DATE_RE = re.compile(r"\d{1,2}[/\-\s][A-Za-z0-9]{2,4}[/\-\s]\d{2,4}")

# Digit runs of 5+ inside description text, not touching a comma/period
# (which would usually mean it's part of a formatted amount, not a code).
REFERENCE_NUMBER_RE = re.compile(r"(?<![\d.,])\d{5,}(?![\d.,])")

# Rows that look like transactions but aren't (repeated page footers etc.)
NOISE_ROW_PATTERNS = [
    r"^page \d+ of \d+$",
    r"^continued",
    r"^statement generated",
    r"^end of statement$",
]

# Footer/contact junk to strip from the RAW sheet -- NOT the account
# summary block, which is deliberately kept there.
FOOTER_NOISE_PATTERNS = [
    r"download app",
    r"chat with \w+",
    r"our website",
    r"head office",
    r"privacy policy",
    r"\+\d[\d\- ]{7,}\d",          # phone numbers
    r"[\w.\-]+@[\w.\-]+\.\w+",      # emails
    r"^(facebook|twitter|instagram|linkedin|youtube)$",
]


# ----------------------------------------------------------------------
# STEP 1: PDF type detection
# ----------------------------------------------------------------------

def is_text_based(pdf_path, sample_pages=2):
    with pdfplumber.open(pdf_path) as pdf:
        text_found = sum(len((p.extract_text() or "").strip()) for p in pdf.pages[:sample_pages])
    return text_found > 30


# ----------------------------------------------------------------------
# STEP 2: Extraction (text-based + OCR fallback)
# ----------------------------------------------------------------------

def build_watermark_filter(page):
    """
    Filters out characters that are almost certainly a decorative
    watermark (e.g. a diagonal 'VOID' stamp) rather than real content.
    Requires BOTH a minority color AND a dramatically oversized font --
    color alone isn't enough signal, since a legitimate white-on-navy
    header banner is also a "rare color" but must never be stripped.
    """
    chars = page.chars
    if not chars:
        return lambda obj: True

    color_counts = Counter(str(c.get("non_stroking_color")) for c in chars)
    if not color_counts:
        return lambda obj: True

    body_color, _ = color_counts.most_common(1)[0]
    body_sizes = [c["size"] for c in chars if str(c.get("non_stroking_color")) == body_color]
    body_median_size = sorted(body_sizes)[len(body_sizes) // 2] if body_sizes else 10

    total = sum(color_counts.values())
    noise_colors = set()
    for color, count in color_counts.items():
        if color == body_color or count / total >= 0.15:
            continue
        color_sizes = [c["size"] for c in chars if str(c.get("non_stroking_color")) == color]
        median_size = sorted(color_sizes)[len(color_sizes) // 2] if color_sizes else 0
        if median_size > body_median_size * 2.5:
            noise_colors.add(color)

    if not noise_colors:
        return lambda obj: True

    def keep(obj):
        if obj["object_type"] != "char":
            return True
        return str(obj.get("non_stroking_color")) not in noise_colors

    return keep


def extract_raw_tables_text_pdf(pdf_path):
    """Extract every row from every table on every page. Includes the
    account summary block -- filtering that out happens later, and only
    for the cleaned sheet, not this raw extraction."""
    all_rows = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            keep_fn = build_watermark_filter(page)
            filtered_page = page.filter(keep_fn)
            tables = filtered_page.extract_tables()
            for table in tables:
                for row in table:
                    if row is None:
                        continue
                    # Drop None cells (merged-cell spacer artifacts from
                    # pdfplumber), keep real empty strings.
                    cells = [c for c in row if c is not None]
                    cleaned = [(c or "").strip().replace("\n", " ") for c in cells]
                    if any(cleaned):
                        all_rows.append({"page": page_num, "cells": cleaned})
    return all_rows


def extract_raw_tables_scanned_pdf(pdf_path, dpi=300, progress=print):
    import pytesseract
    from pdf2image import convert_from_path

    progress("  Rendering pages to images for OCR (this is the slow step)...")
    images = convert_from_path(pdf_path, dpi=dpi)
    all_rows = []

    for page_num, image in enumerate(images, start=1):
        progress(f"  OCR: page {page_num}/{len(images)}...")
        data = pytesseract.image_to_data(image, output_type=pytesseract.Output.DATAFRAME)
        data = data[data.conf > 30].dropna(subset=["text"])
        data = data[data.text.str.strip() != ""]
        if data.empty:
            continue

        data = data.sort_values(["top", "left"])
        line_height_tolerance = 10
        rows, current_row, current_top = [], [], None
        for _, word in data.iterrows():
            if current_top is None or abs(word["top"] - current_top) <= line_height_tolerance:
                current_row.append(word)
                current_top = word["top"] if current_top is None else current_top
            else:
                rows.append(current_row)
                current_row = [word]
                current_top = word["top"]
        if current_row:
            rows.append(current_row)

        for row_words in rows:
            row_words = sorted(row_words, key=lambda w: w["left"])
            cells, current_cell, last_right = [], [], None
            gap_threshold = 25
            for w in row_words:
                if last_right is not None and (w["left"] - last_right) > gap_threshold:
                    cells.append(" ".join(current_cell))
                    current_cell = []
                current_cell.append(str(w["text"]))
                last_right = w["left"] + w["width"]
            if current_cell:
                cells.append(" ".join(current_cell))
            if any(c.strip() for c in cells):
                all_rows.append({"page": page_num, "cells": cells})

    return all_rows


# ----------------------------------------------------------------------
# STEP 3: RAW sheet (keeps summary, strips footer junk)
# ----------------------------------------------------------------------

def is_footer_noise(cells):
    text = " ".join(cells).lower()
    return any(re.search(pattern, text) for pattern in FOOTER_NOISE_PATTERNS)


def build_raw_export_df(raw_rows):
    kept = [r["cells"] for r in raw_rows if not is_footer_noise(r["cells"])]
    if not kept:
        return pd.DataFrame()
    width = max(len(r) for r in kept)
    padded = [r + [""] * (width - len(r)) for r in kept]
    columns = [f"Col {i+1}" for i in range(width)]
    return pd.DataFrame(padded, columns=columns)


# ----------------------------------------------------------------------
# STEP 4: Header detection + role mapping (roles used for logic only)
# ----------------------------------------------------------------------

def match_header(cell_text):
    text = _normalize(cell_text)
    if not text:
        return None
    for canonical, keyword, _orig in _HEADER_MATCH_ORDER:
        if text == keyword or text.startswith(keyword) or keyword in text:
            return canonical
    return None


def matching_keyword(cell_text):
    """Like match_header but also returns the original (spaced) keyword
    that hit, so we can reinsert a lost space (e.g. 'TRANSDATE' ->
    'TRANS DATE') using the keyword's own word boundaries -- a cosmetic
    fix for an extraction artifact, not a rename of the bank's own label."""
    text = _normalize(cell_text)
    if not text:
        return None, None
    for canonical, keyword, orig in _HEADER_MATCH_ORDER:
        if text == keyword or text.startswith(keyword) or keyword in text:
            return canonical, orig
    return None, None


def fix_header_spacing(raw_text, keyword):
    """Reinsert a space lost during PDF extraction of a wrapped header
    cell, using the matched keyword's word boundaries. Only applies when
    the extracted text has literally zero spaces and the lengths line up
    exactly -- otherwise leaves the bank's text untouched."""
    if not keyword or " " not in keyword or " " in raw_text.strip():
        return raw_text
    parts = keyword.split()
    normalized_raw = re.sub(r"\s+", "", raw_text)
    if sum(len(p) for p in parts) != len(normalized_raw):
        return raw_text
    pieces, idx = [], 0
    for p in parts:
        pieces.append(normalized_raw[idx:idx + len(p)])
        idx += len(p)
    return " ".join(pieces)


def find_header_row(raw_rows):
    for idx, item in enumerate(raw_rows):
        matches = [match_header(c) for c in item["cells"]]
        matches = [m for m in matches if m]
        if "Date" in matches and "Description" in matches and any(
            m in matches for m in ["Debit", "Credit", "Balance"]
        ):
            return idx, item["cells"]
    return None, None


def build_column_schema(header_cells):
    """Returns list of {index, display, role} -- display is the bank's
    own header text (lightly de-artifacted), role is for internal logic
    only and never shown to the user."""
    schema = []
    for i, cell in enumerate(header_cells):
        if not cell.strip():
            continue
        role, keyword = matching_keyword(cell)
        display = fix_header_spacing(cell.strip(), keyword)
        schema.append({"index": i, "display": display, "role": role})
    return schema


def role_column(schema, role):
    for col in schema:
        if col["role"] == role:
            return col["display"]
    return None


def build_transactions_df(raw_rows, header_idx, schema):
    display_cols = [c["display"] for c in schema]
    desc_display = role_column(schema, "Description")

    records = []
    for item in raw_rows[header_idx + 1:]:
        cells = item["cells"]
        row = {d: "" for d in display_cols}
        for i, val in enumerate(cells):
            matching = next((c for c in schema if c["index"] == i), None)
            if matching:
                d = matching["display"]
                row[d] = (row[d] + " " + val).strip() if row[d] else val
            elif desc_display:
                # Cell beyond the known columns -- almost always spillover
                # from the description; attach it there rather than
                # inventing a new column.
                row[desc_display] = (row[desc_display] + " " + val).strip()
        records.append(row)

    return pd.DataFrame(records, columns=display_cols)


# ----------------------------------------------------------------------
# STEP 5: Validate & repair (uses roles to find the right columns)
# ----------------------------------------------------------------------

def parse_date(value):
    value = value.strip()
    if not value:
        return None
    normalized = re.sub(r"\s*-\s*", "-", value)  # "30-May- 2024" -> "30-May-2024"

    # Tier 1: known exact formats (fast, zero false-positive risk)
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(normalized, fmt)
        except ValueError:
            continue

    # Tier 2: dateutil, strict (handles format/spacing variants we haven't
    # explicitly listed -- "18-May 2021", "18.05.2021", "May 18, 2021", etc.)
    # dayfirst=True matches how these statements write numeric dates (DD-MM-YYYY).
    for candidate in (normalized, value):
        try:
            dt = dateutil_parser.parse(candidate, dayfirst=True, fuzzy=False)
            if 1990 <= dt.year <= datetime.now().year + 1:
                return dt
        except (ValueError, OverflowError):
            pass

    # Tier 3: dateutil fuzzy, last resort (pulls a date out of a string with
    # extra junk attached). Bounded by a sane year range so it can't turn
    # random noise into a false date.
    try:
        dt = dateutil_parser.parse(value, dayfirst=True, fuzzy=True)
        if 1990 <= dt.year <= datetime.now().year + 1:
            return dt
    except (ValueError, OverflowError):
        pass

    return None


def looks_like_amount(value):
    value = value.strip()
    if value == "":
        return True
    return bool(AMOUNT_RE.fullmatch(value))


def clean_amount(value):
    value = value.strip().replace(",", "")
    if value == "":
        return None
    negative = value.startswith("(") and value.endswith(")")
    value = value.strip("()")
    try:
        num = float(value)
        return -num if negative else num
    except ValueError:
        return None


def extract_reference_number(description):
    """Pulls the longest 5+ digit run out of the description and returns
    (reference_number, description_with_that_text_removed) -- the code
    moves into its own column instead of appearing in both places."""
    matches = REFERENCE_NUMBER_RE.findall(description)
    if not matches:
        return "", description
    ref = max(matches, key=len)
    cleaned_desc = description.replace(ref, "", 1)
    cleaned_desc = re.sub(r"\s+", " ", cleaned_desc).strip()
    return ref, cleaned_desc


def clean_dataframe(df, schema):
    date_col = role_column(schema, "Date")
    desc_col = role_column(schema, "Description")
    debit_col = role_column(schema, "Debit")
    credit_col = role_column(schema, "Credit")
    ref_col = role_column(schema, "Reference")  # bank's own ref/chq column, if any
    amount_cols = [c for c in [debit_col, credit_col] if c]

    def is_repeated_header(row):
        return match_header(row[date_col]) == "Date" and match_header(row[desc_col]) == "Description"

    def is_noise_row(row):
        desc = row[desc_col].strip().lower()
        return any(re.match(p, desc) for p in NOISE_ROW_PATTERNS) or is_repeated_header(row)

    def has_suspect_reference(row):
        if not ref_col:
            return False
        val = row[ref_col].strip()
        return bool(val) and not any(ch.isdigit() for ch in val) and len(val) <= 6

    def repair_row(row):
        row = row.copy()
        if parse_date(row[date_col]) is None:
            for col in [desc_col] + amount_cols:
                match = DATE_RE.search(row[col])
                if match and parse_date(match.group()):
                    row[date_col] = match.group()
                    row[col] = row[col].replace(match.group(), "").strip()
                    break
        for target in amount_cols:
            if not looks_like_amount(row[target]):
                match = AMOUNT_RE.search(row[target])
                if match:
                    row[target] = match.group()
                else:
                    found = AMOUNT_RE.findall(row[desc_col])
                    if found:
                        candidate = found[-1]
                        row[desc_col] = row[desc_col].replace(candidate, "", 1).strip()
                        row[target] = candidate
        if has_suspect_reference(row):
            row[desc_col] = (row[desc_col].rstrip() + row[ref_col].strip()).strip()
            row[ref_col] = ""
        return row

    cleaned_rows, flags = [], []
    for _, row in df.iterrows():
        if is_noise_row(row):
            continue
        problems = []
        if parse_date(row[date_col]) is None:
            problems.append("bad_date")
        for col in amount_cols:
            if not looks_like_amount(row[col]):
                problems.append(f"bad_{col}")
        if row[desc_col].strip() == "":
            problems.append("empty_description")
        if has_suspect_reference(row):
            problems.append("suspect_reference")

        if problems:
            row = repair_row(row)
            problems = []
            if parse_date(row[date_col]) is None:
                problems.append("bad_date")
            for col in amount_cols:
                if not looks_like_amount(row[col]):
                    problems.append(f"bad_{col}")

        cleaned_rows.append(row)
        flags.append(";".join(problems))

    cleaned_df = pd.DataFrame(cleaned_rows).reset_index(drop=True)
    ref_and_desc = cleaned_df[desc_col].apply(extract_reference_number)
    cleaned_df["Reference Number"] = ref_and_desc.apply(lambda pair: pair[0])
    cleaned_df[desc_col] = ref_and_desc.apply(lambda pair: pair[1])
    cleaned_df["Flag"] = flags

    cleaned_df["_Date_parsed"] = cleaned_df[date_col].apply(parse_date)
    for col in amount_cols:
        cleaned_df[col] = cleaned_df[col].apply(clean_amount)

    still_bad = cleaned_df["_Date_parsed"].isna()
    dropped = cleaned_df[still_bad].copy()
    cleaned_df = cleaned_df[~still_bad].copy()

    cleaned_df["_Month"] = cleaned_df["_Date_parsed"].dt.strftime("%B %Y")
    cleaned_df["_Year"] = cleaned_df["_Date_parsed"].dt.year.astype(str)
    # kind="stable" preserves original row order for same-day transactions
    # (default quicksort does not, and can shuffle the running balance).
    cleaned_df = cleaned_df.sort_values("_Date_parsed", kind="stable").reset_index(drop=True)

    return cleaned_df, dropped, desc_col


# ----------------------------------------------------------------------
# STEP 6-7: Export
# ----------------------------------------------------------------------

def autofit_columns(worksheet, df):
    for i, col in enumerate(df.columns, start=1):
        max_len = df[col].apply(lambda v: len(str(v)) if pd.notna(v) else 0).max() if len(df) else 0
        max_len = max(max_len, len(str(col)))
        worksheet.column_dimensions[get_column_letter(i)].width = min(max_len + 3, 50)


def export_to_excel(raw_export_df, cleaned_df, dropped_df, schema, desc_col, output_path, split_mode):
    display_cols = [c["display"] for c in schema]
    desc_position = display_cols.index(desc_col)
    display_cols.insert(desc_position + 1, "Reference Number")
    display_cols.append("Flag")

    group_col = {"month": "_Month", "year": "_Year"}.get(split_mode)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        if len(raw_export_df):
            raw_export_df.to_excel(writer, sheet_name="RAW_EXTRACTED", index=False)
            autofit_columns(writer.sheets["RAW_EXTRACTED"], raw_export_df)

        if group_col:
            for key, group in cleaned_df.groupby(group_col, sort=False):
                sheet_name = str(key)[:31]
                group_export = group[display_cols]
                group_export.to_excel(writer, sheet_name=sheet_name, index=False)
                autofit_columns(writer.sheets[sheet_name], group_export)

            combined_export = cleaned_df[display_cols]
            combined_export.to_excel(writer, sheet_name="ALL_TRANSACTIONS", index=False)
            autofit_columns(writer.sheets["ALL_TRANSACTIONS"], combined_export)
        else:
            combined_export = cleaned_df[display_cols]
            combined_export.to_excel(writer, sheet_name="TRANSACTIONS", index=False)
            autofit_columns(writer.sheets["TRANSACTIONS"], combined_export)

        if len(dropped_df):
            dropped_cols = [c["display"] for c in schema]
            dropped_export = dropped_df[dropped_cols]
            dropped_export.to_excel(writer, sheet_name="NEEDS_REVIEW", index=False)
            autofit_columns(writer.sheets["NEEDS_REVIEW"], dropped_export)


# ----------------------------------------------------------------------
# MAIN
# ----------------------------------------------------------------------

def convert(pdf_path, output_path, progress=print):
    """progress: callable(str) -- defaults to print() for CLI use; the
    Streamlit app passes in something that updates the UI instead."""
    t0 = time.time()
    progress(f"[1/6] Checking PDF type: {pdf_path}")
    text_based = is_text_based(pdf_path)

    if text_based:
        progress("      -> Text-based PDF. Extracting tables (with watermark filtering)...")
        raw_rows = extract_raw_tables_text_pdf(pdf_path)
    else:
        progress("      -> Scanned/image PDF detected. Running OCR (this takes longer)...")
        raw_rows = extract_raw_tables_scanned_pdf(pdf_path, progress=progress)

    if not raw_rows:
        raise ValueError("No tables/rows could be extracted. This layout may need a custom strategy.")
    progress(f"[2/6] Extracted {len(raw_rows)} raw rows in {time.time()-t0:.1f}s")

    progress("[3/6] Building RAW sheet (summary kept, footer/contact junk removed)...")
    raw_export_df = build_raw_export_df(raw_rows)

    progress("[4/6] Detecting header row (column names/order kept exactly as in the PDF)...")
    header_idx, header_cells = find_header_row(raw_rows)
    if header_idx is None:
        raise ValueError(
            "Could not find a recognizable transaction table header. "
            "This bank's column labels may need to be added to HEADER_KEYWORDS."
        )
    schema = build_column_schema(header_cells)
    progress(f"      -> Header: {[c['display'] for c in schema]}")
    progress(f"      -> Discarded {header_idx} pre-header rows (account summary) from the cleaned sheet only.")

    df = build_transactions_df(raw_rows, header_idx, schema)
    progress(f"[5/6] Validating & repairing {len(df)} transaction rows...")
    cleaned_df, dropped_df, desc_col = clean_dataframe(df, schema)
    n_flagged = (cleaned_df["Flag"] != "").sum()
    progress(f"      -> {len(cleaned_df)} clean transactions ({n_flagged} flagged for review).")
    if len(dropped_df):
        progress(f"      -> {len(dropped_df)} rows dropped (no valid date) -> NEEDS_REVIEW sheet.")

    n_years = cleaned_df["_Year"].nunique() if len(cleaned_df) else 0
    n_months = cleaned_df["_Month"].nunique() if len(cleaned_df) else 0
    if len(cleaned_df) <= MONTH_SPLIT_THRESHOLD:
        split_mode = "none"
    elif n_years <= 1:
        split_mode = "month"
    else:
        split_mode = "year"

    split_description = {
        "none": "single sheet, under threshold",
        "month": f"split into {n_months} monthly sheets (single year, over threshold)",
        "year": f"split into {n_years} yearly sheets (spans multiple years, over threshold)",
    }[split_mode]
    progress(f"[6/6] Exporting ({split_description})...")
    export_to_excel(raw_export_df, cleaned_df, dropped_df, schema, desc_col, output_path, split_mode)

    elapsed = time.time() - t0
    progress(f"Done in {elapsed:.1f}s -> {output_path}")

    return {
        "elapsed_seconds": round(elapsed, 1),
        "n_transactions": len(cleaned_df),
        "n_flagged": int(n_flagged),
        "n_dropped": len(dropped_df),
        "n_months": n_months,
        "n_years": n_years,
        "split_mode": split_mode,
        "columns": [c["display"] for c in schema],
    }


def convert_bytes(pdf_bytes, filename="statement.pdf", progress=print):
    """
    Bytes-in, bytes-out. Writes the uploaded PDF to a temp directory,
    runs the full conversion, reads the resulting Excel file back into
    memory, then lets the temp directory (and everything in it -- the
    original PDF included) get deleted on exit. Nothing persists on disk
    after this function returns, success or failure.
    """
    with tempfile.TemporaryDirectory() as tmpdir:
        input_path = os.path.join(tmpdir, filename)
        output_path = os.path.join(tmpdir, "converted.xlsx")

        with open(input_path, "wb") as f:
            f.write(pdf_bytes)

        stats = convert(input_path, output_path, progress=progress)

        with open(output_path, "rb") as f:
            output_bytes = f.read()

    return output_bytes, stats


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python statement_converter.py input.pdf output.xlsx")
        sys.exit(1)
    try:
        convert(sys.argv[1], sys.argv[2])
    except ValueError as e:
        print(f"Error: {e}")
        sys.exit(1)
