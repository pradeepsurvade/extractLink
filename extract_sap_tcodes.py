"""
SAP Transaction Code Extractor
================================
Scans a folder of Word (.docx) documents, extracts SAP transaction codes,
and outputs an Excel file with:
  Column A: Transaction Code  (distinct - merged cell)
  Column B: Document Name     (multiple rows per code)
  Column C: Source Sentence   (the line where the code was found)

Logic: A transaction code is ONLY extracted when it immediately follows
a known SAP trigger phrase (Run, T-code, Tcode, SAP Transaction, etc.)
exactly as shown in the reference screenshot.

Usage: place .docx files in the input/ folder, run script, get output/SAP_Tcodes_Extract.xlsx
"""

import os
import re
from collections import defaultdict

try:
    from docx import Document
except ImportError:
    raise SystemExit("python-docx not installed. Run: pip install python-docx")

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
except ImportError:
    raise SystemExit("openpyxl not installed. Run: pip install openpyxl")


# ---------------------------------------------------------------------------
# Trigger-based extraction patterns
# Each pattern captures the transaction code in group(1).
# Ordered from most specific to least specific.
# ---------------------------------------------------------------------------
TRIGGER_PATTERNS = [
    # "Run SAP t code CJ88" / "Run SAP t-code CJ88"
    re.compile(r'Run\s+SAP\s+[Tt][-\s]?[Cc]ode\s+([A-Z][A-Z0-9_\-]{1,29})', re.IGNORECASE),

    # "Run the YRTR_ASSET_BALANCES Tcode"  (code comes BEFORE the word Tcode)
    re.compile(r'Run\s+the\s+([A-Z][A-Z0-9_\-]{1,29})\s+[Tt]code', re.IGNORECASE),

    # "Run S_ALR-87013611 SAP Tcode" (code before trailing SAP Tcode)
    re.compile(r'Run\s+([A-Z][A-Z0-9_\-]{1,29})\s+SAP\s+[Tt]code', re.IGNORECASE),

    # "Run S_ALR_87011964" / "Run FAGLB03" — only when followed by SAP context words
    # Postfix: SAP, Tcode, T-code, T code, Transaction (case-insensitive)
    re.compile(r'Run\s+([A-Z][A-Z0-9_\-]{1,29})\s+(?:SAP|[Tt][-\s]?[Cc]ode|[Tt]ransaction)\b', re.IGNORECASE),

    # "T-code ZRTR_GL_PARK_GL" / "T-code (CJ20N)" / "Tcode CJ20N" / "Tcode: FB01"
    re.compile(r'[Tt][-\s]?[Cc]ode[s]?\s*[:\(]?\s*([A-Z][A-Z0-9_\-]{1,29})\)?', re.IGNORECASE),

    # "SAP Transaction CJ20N" / "SAP Transaction (CJ20N)" — SAP directly before Transaction
    re.compile(r'SAP\s+[Tt]ransaction\s*[:\(]?\s*([A-Z][A-Z0-9_\-]{1,29})\)?', re.IGNORECASE),

    # "SAP Asset transaction (S_ALR_87012048)" — SAP + any words + transaction + (CODE)
    re.compile(r'SAP\s+\w[\w\s]{0,40}[Tt]ransaction\s*\(\s*([A-Z][A-Z0-9_\-]{1,29})\s*\)', re.IGNORECASE),

    # "Transaction code FB01" / "Transaction: FB01"
    re.compile(r'[Tt]ransaction\s+[Cc]ode\s*[:\(]?\s*([A-Z][A-Z0-9_\-]{1,29})\)?', re.IGNORECASE),
]

# Words that must never be treated as transaction codes
NOT_A_TCODE = {
    'THE', 'AND', 'FOR', 'RUN', 'SAP', 'CODE', 'CODES', 'WITH', 'FROM', 'INTO',
    'THIS', 'THAT', 'HAVE', 'BEEN', 'WILL', 'ALSO', 'CAN', 'NOT', 'ARE', 'ALL',
    'BUT', 'USE', 'NEW', 'OLD', 'END', 'ADD', 'SET', 'GET', 'PUT', 'OUT', 'OFF',
    'TOP', 'BOX', 'YES', 'NO', 'OK', 'GO', 'DO', 'IF', 'AS', 'AT', 'BE', 'BY',
    'IN', 'IS', 'IT', 'OF', 'ON', 'OR', 'SO', 'TO', 'UP', 'US', 'WE', 'ME',
    'TCODE', 'TCODES', 'TRANSACTION', 'REPORT', 'MODULE', 'SYSTEM', 'TABLE',
    'FIELD', 'VALUE', 'PLEASE', 'CLICK', 'OPEN', 'CLOSE', 'ENTER', 'SELECT',
    'PRESS', 'BUTTON', 'SCREEN', 'WINDOW', 'MENU', 'LIST', 'VIEW', 'NEXT',
    'BACK', 'SAVE', 'EXIT', 'NOTE', 'STEP', 'THEN', 'WHEN', 'ONCE', 'AFTER',
    'BEFORE', 'USING', 'BELOW', 'ABOVE', 'RIGHT', 'LEFT', 'HERE', 'THERE',
}


def is_valid_tcode(candidate: str) -> bool:
    c = candidate.upper().strip().rstrip('.,;)(')
    if not c or c in NOT_A_TCODE:
        return False
    if len(c) < 2 or not c[0].isalpha():
        return False
    # Only alphanumeric + underscore
    if not re.match(r'^[A-Z][A-Z0-9_]+$', c):
        return False
    return True


def extract_tcodes_from_text(text: str) -> list:
    """Return list of (tcode, source_sentence) — trigger-phrase based only."""
    results = []
    text_stripped = text.strip()
    if not text_stripped:
        return results

    seen = set()
    for pattern in TRIGGER_PATTERNS:
        for m in pattern.finditer(text):
            raw = m.group(1).strip().rstrip('.,;)(')
            candidate = raw.upper()
            if is_valid_tcode(candidate) and candidate not in seen:
                seen.add(candidate)
                candidate = candidate.replace('-', '_')
                results.append((candidate, text_stripped))

    return results


def extract_from_docx(filepath: str) -> list:
    """Return list of (tcode, rel_path, source_sentence) for all matches in a docx."""
    doc = Document(filepath)
    matches = []

    for para in doc.paragraphs:
        for tcode, sentence in extract_tcodes_from_text(para.text):
            matches.append((tcode, sentence))

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for tcode, sentence in extract_tcodes_from_text(cell.text):
                    matches.append((tcode, sentence))

    return matches


def scan_folder(folder: str):
    """
    Recursively scan all .docx files in folder and subfolders.
    Returns:
      tcode_data : {tcode: [(rel_path, sentence), ...]}
      others     : [rel_paths with no t-codes found]
    """
    tcode_data = defaultdict(list)
    others = []

    all_docx = []
    for dirpath, _, filenames in os.walk(folder):
        for f in filenames:
            if f.lower().endswith('.docx') and not f.startswith('~$'):
                all_docx.append(os.path.join(dirpath, f))
    all_docx.sort()

    if not all_docx:
        raise SystemExit(
            f"No .docx files found in '{folder}' or any of its subfolders.\n"
            f"Place your Word documents inside the 'input' folder and try again."
        )

    print(f"Found {len(all_docx)} Word document(s) to scan...\n")

    for filepath in all_docx:
        rel_path = os.path.relpath(filepath, folder)
        try:
            matches = extract_from_docx(filepath)
        except Exception as e:
            print(f"  [WARN] Could not read '{rel_path}': {e}")
            others.append(rel_path)
            continue

        if matches:
            for tcode, sentence in matches:
                tcode_data[tcode].append((rel_path, sentence))
            unique = len(set(m[0] for m in matches))
            print(f"  + {rel_path} -> {unique} t-code(s) found")
        else:
            others.append(rel_path)
            print(f"  - {rel_path} -> no t-codes found (-> Others)")

    return tcode_data, others


def build_excel(tcode_data: dict, others: list, output_path: str):
    wb = Workbook()
    ws = wb.active
    ws.title = "SAP T-Codes"

    # Styles
    hdr_font      = Font(name='Arial', bold=True, color='FFFFFF', size=11)
    hdr_fill      = PatternFill('solid', start_color='1F4E79')
    hdr_align     = Alignment(horizontal='center', vertical='center', wrap_text=True)
    tcode_font    = Font(name='Courier New', bold=True, color='1F4E79', size=10)
    others_font   = Font(name='Courier New', bold=True, color='C00000', size=10)
    doc_font      = Font(name='Arial', size=10)
    sentence_font = Font(name='Arial', size=10, italic=True)
    grey_font     = Font(name='Arial', size=10, italic=True, color='808080')
    top_align     = Alignment(horizontal='left', vertical='top',   wrap_text=True)
    mid_align     = Alignment(horizontal='center', vertical='center', wrap_text=True)

    def mk_border(weight='thin'):
        s = Side(style=weight, color='BDD7EE' if weight == 'thin' else '1F4E79')
        return Border(left=s, right=s, top=s, bottom=s)

    thin_border   = mk_border('thin')
    medium_border = mk_border('medium')
    alt_fill      = PatternFill('solid', start_color='EBF3FB')
    base_fill     = PatternFill('solid', start_color='FFFFFF')

    # Header row
    for col, h in enumerate(['Transaction Code', 'Document Name', 'Source Sentence'], 1):
        c = ws.cell(row=1, column=col, value=h)
        c.font = hdr_font; c.fill = hdr_fill
        c.alignment = hdr_align; c.border = medium_border
    ws.row_dimensions[1].height = 32

    row   = 2
    block = 0

    for tcode in sorted(tcode_data.keys()):
        entries   = tcode_data[tcode]
        n         = len(entries)
        start_row = row
        fill      = alt_fill if (block % 2 == 0) else base_fill
        block    += 1

        for rel_path, sentence in entries:
            c = ws.cell(row=row, column=2, value=rel_path)
            c.font = doc_font; c.fill = fill
            c.alignment = top_align; c.border = thin_border

            c = ws.cell(row=row, column=3, value=sentence)
            c.font = sentence_font; c.fill = fill
            c.alignment = top_align; c.border = thin_border

            ws.row_dimensions[row].height = 38
            row += 1

        # Column A — write once, merge if multiple rows
        c = ws.cell(row=start_row, column=1, value=tcode)
        c.font = tcode_font; c.fill = fill
        c.alignment = mid_align; c.border = medium_border

        if n > 1:
            ws.merge_cells(start_row=start_row, start_column=1,
                           end_row=row - 1, end_column=1)
            c = ws.cell(row=start_row, column=1)
            c.font = tcode_font; c.fill = fill
            c.alignment = mid_align; c.border = medium_border

    # Others section
    if others:
        others_fill = PatternFill('solid', start_color='FFE7E7')
        start_row   = row
        n           = len(others)

        for rel_path in sorted(others):
            fill = alt_fill if (row % 2 == 0) else base_fill

            c = ws.cell(row=row, column=2, value=rel_path)
            c.font = doc_font; c.fill = fill
            c.alignment = top_align; c.border = thin_border

            c = ws.cell(row=row, column=3, value='No SAP transaction code detected')
            c.font = grey_font; c.fill = fill
            c.alignment = top_align; c.border = thin_border

            ws.row_dimensions[row].height = 30
            row += 1

        c = ws.cell(row=start_row, column=1, value='Others')
        c.font = others_font; c.fill = others_fill
        c.alignment = mid_align; c.border = medium_border

        if n > 1:
            ws.merge_cells(start_row=start_row, start_column=1,
                           end_row=row - 1, end_column=1)
            c = ws.cell(row=start_row, column=1)
            c.font = others_font; c.fill = others_fill
            c.alignment = mid_align; c.border = medium_border

    ws.column_dimensions['A'].width = 26
    ws.column_dimensions['B'].width = 38
    ws.column_dimensions['C'].width = 82
    ws.freeze_panes = 'B2'

    wb.save(output_path)
    print(f"\nExcel saved to: {output_path}")
    print(f"  {len(tcode_data)} unique transaction code(s) found.")
    print(f"  {len(others)} document(s) with no codes -> 'Others'.")


def main():
    script_dir  = os.path.dirname(os.path.abspath(__file__))
    input_dir   = os.path.join(script_dir, 'input')
    output_dir  = os.path.join(script_dir, 'output')
    output_file = os.path.join(output_dir, 'SAP_Tcodes_Extract.xlsx')

    os.makedirs(input_dir,  exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)

    print(f"Input  folder : {input_dir}")
    print(f"Output folder : {output_dir}\n")

    tcode_data, others = scan_folder(input_dir)
    build_excel(tcode_data, others, output_file)


if __name__ == '__main__':
    main()
