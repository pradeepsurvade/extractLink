import os, zipfile, re
import xml.etree.ElementTree as ET
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

# ─── DOCX header extraction ───────────────────────────────────────────────────

def para_text(p):
    return re.sub(r'  +', ' ', ''.join(t.text for t in p.iter(f'{{{W_NS}}}t') if t.text)).strip()

def extract_header_fields(docx_path):
    """
    Dynamically extract label/value pairs from the docx page header table.
    Scans ALL header XML files (header1.xml, header2.xml, etc.) to find
    the one containing the table — no filename is hardcoded.
    No field names are hardcoded; labels are read verbatim from the document.

    Two cell layouts handled:
      Single-paragraph: "Label: value text"       → split on first colon
      Multi-paragraph:  ["Label: extra", "value"]  → label = colon paragraph, value = rest
    """
    try:
        with zipfile.ZipFile(docx_path, 'r') as z:
            # Find all word/headerN.xml files that contain a table
            header_files = [
                n for n in z.namelist()
                if re.match(r'word/header\d+\.xml$', n)
            ]
            header_xml = None
            for hf in sorted(header_files):
                content = z.read(hf)
                if b'<w:tbl' in content:   # quick check before full parse
                    header_xml = content
                    break
            if header_xml is None:
                return {}
        root = ET.fromstring(header_xml)
    except Exception:
        return {}

    header = {}
    for cell in root.iter(f'{{{W_NS}}}tc'):
        non_empty = [para_text(p) for p in cell.iter(f'{{{W_NS}}}p')]
        non_empty = [p for p in non_empty if p]

        colon_paras = [p for p in non_empty if ':' in p]
        if not colon_paras:
            continue

        lp = colon_paras[0]  # paragraph that carries the colon

        if len(non_empty) == 1:
            # Single paragraph: "Label: value" — split on first colon
            idx = lp.index(':')
            label = lp[:idx].strip()
            value = lp[idx + 1:].strip()
        else:
            # Multiple paragraphs: colon paragraph = label (verbatim), rest = value
            label = lp
            remaining = [p for p in non_empty if p != lp]
            value = ' '.join(remaining).strip()

        if label:
            header[label] = value

    return header

def scan_docx_folder(folder):
    """Scan folder/subfolders for .docx files; return records and ordered header keys."""
    all_header_keys = []
    seen_keys = set()
    raw = []

    for docx_path in sorted(Path(folder).rglob('*.docx')):
        fields = extract_header_fields(str(docx_path))
        for k in fields:
            if k not in seen_keys:
                all_header_keys.append(k)
                seen_keys.add(k)
        raw.append((docx_path.parent.name, docx_path.name, fields))

    records = []
    for folder_name, file_name, fields in raw:
        row = {'_folder': folder_name, '_file': file_name}
        for k in all_header_keys:
            row[k] = fields.get(k, '')
        records.append(row)

    return records, all_header_keys

# ─── Excel formatting helpers ─────────────────────────────────────────────────

DARK_BLUE   = 'FF1F4E79'
MED_BLUE    = 'FF2E75B6'
LIGHT_GREY  = 'FFF2F2F2'
WHITE_FONT  = 'FFFFFFFF'
DARK_FONT   = 'FF1F4E79'
BORDER_DARK = 'FF4472C4'
BORDER_GREY = 'FFBFBFBF'

def thin_border(color):
    s = Side(style='thin', color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def medium_border(color):
    s = Side(style='medium', color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def medium_border_no_top(color):
    s = Side(style='medium', color=color)
    return Border(left=s, right=s, bottom=s)

def apply_main_header(cell, value):
    cell.value = value
    cell.font = Font(bold=True, color=WHITE_FONT, name='Calibri', size=11)
    cell.fill = PatternFill('solid', fgColor=DARK_BLUE)
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    cell.border = thin_border(BORDER_GREY)

def apply_sub_header(cell, value):
    cell.value = value
    cell.font = Font(bold=True, color=WHITE_FONT, name='Calibri', size=11)
    cell.fill = PatternFill('solid', fgColor=MED_BLUE)
    cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    cell.border = thin_border(BORDER_GREY)

def apply_folder_file(cell, value, is_top_row):
    cell.value = value
    cell.font = Font(bold=True, color=DARK_FONT, name='Calibri', size=11)
    cell.fill = PatternFill('solid', fgColor=LIGHT_GREY)
    cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
    cell.border = medium_border(BORDER_DARK) if is_top_row else medium_border_no_top(BORDER_DARK)

def apply_data(cell, value):
    cell.value = value
    cell.font = Font(name='Calibri', size=11)
    cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
    cell.border = thin_border(BORDER_GREY)

# ─── Build Excel ──────────────────────────────────────────────────────────────

def build_excel(records, header_keys, output_path):
    wb = Workbook()
    ws = wb.active
    ws.title = 'All Tables'

    n_header_cols = len(header_keys)

    # Row 1: fixed column headers + merged "Header Content" spanning all dynamic cols
    apply_main_header(ws.cell(1, 1), 'Folder Name')
    apply_main_header(ws.cell(1, 2), 'Word Doc Name')
    apply_main_header(ws.cell(1, 3), 'Header Content')
    if n_header_cols > 1:
        ws.merge_cells(start_row=1, start_column=3, end_row=1, end_column=2 + n_header_cols)
    ws.row_dimensions[1].height = 28

    # 2 rows per record: sub-header (dynamic labels) + data (dynamic values)
    excel_row = 2
    for rec in records:
        # Sub-header: labels taken verbatim from docx
        for ci, key in enumerate(header_keys, start=3):
            apply_sub_header(ws.cell(excel_row, ci), key)
        for col_idx, field in ((1, '_folder'), (2, '_file')):
            apply_folder_file(ws.cell(excel_row, col_idx), rec[field], is_top_row=True)
        ws.row_dimensions[excel_row].height = 18
        excel_row += 1

        # Data: values taken verbatim from docx
        for ci, key in enumerate(header_keys, start=3):
            apply_data(ws.cell(excel_row, ci), rec.get(key, ''))
        for col_idx in (1, 2):
            apply_folder_file(ws.cell(excel_row, col_idx), None, is_top_row=False)
        ws.row_dimensions[excel_row].height = 18
        excel_row += 1

        # Merge folder/file cells across both rows
        ws.merge_cells(start_row=excel_row-2, start_column=1, end_row=excel_row-1, end_column=1)
        ws.merge_cells(start_row=excel_row-2, start_column=2, end_row=excel_row-1, end_column=2)

    # Column widths
    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 42
    for i in range(n_header_cols):
        ws.column_dimensions[get_column_letter(3 + i)].width = 32

    wb.save(output_path)
    print(f'Saved: {output_path}')
    print(f'  Records      : {len(records)}')
    print(f'  Header fields: {header_keys}')
    if records:
        print(f'  Sample values: {[records[0].get(k) for k in header_keys]}')

# ─── Main ─────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    INPUT_FOLDER = '/mnt/user-data/uploads'
    OUTPUT_PATH  = '/mnt/user-data/outputs/header_extract_formatted.xlsx'

    records, header_keys = scan_docx_folder(INPUT_FOLDER)
    if not records:
        print('No .docx files found.')
    else:
        os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
        build_excel(records, header_keys, OUTPUT_PATH)
