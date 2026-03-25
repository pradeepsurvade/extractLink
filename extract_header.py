import os, zipfile, re
import xml.etree.ElementTree as ET
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

# ─── DOCX header extraction ───────────────────────────────────────────────────

def get_cell_text(cell_elem):
    return ' '.join(t.text for t in cell_elem.iter(f'{{{W_NS}}}t') if t.text).strip()

def extract_header_fields(docx_path):
    """
    Dynamically extract all 'Label: Value' pairs from the docx header table.
    No field names are hardcoded — whatever label appears before a colon
    becomes the key, and whatever follows becomes the value.
    """
    try:
        with zipfile.ZipFile(docx_path, 'r') as z:
            if 'word/header1.xml' not in z.namelist():
                return {}
            content = z.read('word/header1.xml')
        root = ET.fromstring(content)
    except Exception:
        return {}

    header = {}
    for row in root.iter(f'{{{W_NS}}}tr'):
        for cell in row.iter(f'{{{W_NS}}}tc'):
            text = get_cell_text(cell)
            if ':' in text:
                label_raw, value_raw = text.split(':', 1)
                label = re.sub(r'\s+', ' ', label_raw).strip()
                value = re.sub(r'\s+', ' ', value_raw).strip()
                if label:
                    # Remove duplicate leading word echoed from label into value
                    # e.g. "SOP No. and Title: SOP SOP_FIN..." → strip first "SOP "
                    first_word = label.split()[0] if label.split() else ''
                    if first_word and value.startswith(first_word + ' '):
                        remainder = value[len(first_word):].strip()
                        if remainder.startswith(first_word):
                            value = remainder
                    header[label] = value
    return header

def scan_docx_folder(folder):
    """Scan folder/subfolders for .docx files; return records and ordered header keys."""
    all_header_keys = []
    seen_keys = set()
    raw = []

    for docx_path in sorted(Path(folder).rglob('*.docx')):
        folder_name = docx_path.parent.name
        file_name = docx_path.name
        fields = extract_header_fields(str(docx_path))
        for k in fields:
            if k not in seen_keys:
                all_header_keys.append(k)
                seen_keys.add(k)
        raw.append((folder_name, file_name, fields))

    records = []
    for folder_name, file_name, fields in raw:
        row = {'Folder Name': folder_name, 'File Name': file_name}
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

    # Row 1: top-level headers
    apply_main_header(ws.cell(1, 1), 'Folder Name')
    apply_main_header(ws.cell(1, 2), 'Word Doc Name')
    apply_main_header(ws.cell(1, 3), 'Header Content')
    if n_header_cols > 1:
        ws.merge_cells(start_row=1, start_column=3, end_row=1, end_column=2 + n_header_cols)
    ws.row_dimensions[1].height = 28

    # Data rows (2 rows per record: sub-header + values)
    excel_row = 2
    for rec in records:
        # Sub-header row: field labels (dynamic - taken directly from the docx)
        for ci, key in enumerate(header_keys, start=3):
            apply_sub_header(ws.cell(excel_row, ci), key)
        for col_idx in (1, 2):
            apply_folder_file(ws.cell(excel_row, col_idx),
                              rec['Folder Name'] if col_idx == 1 else rec['File Name'],
                              is_top_row=True)
        ws.row_dimensions[excel_row].height = 18
        excel_row += 1

        # Value row
        for ci, key in enumerate(header_keys, start=3):
            apply_data(ws.cell(excel_row, ci), rec.get(key, ''))
        for col_idx in (1, 2):
            apply_folder_file(ws.cell(excel_row, col_idx), None, is_top_row=False)
        ws.row_dimensions[excel_row].height = 18
        excel_row += 1

        # Merge Folder Name and File Name cells across the 2 rows
        ws.merge_cells(start_row=excel_row - 2, start_column=1, end_row=excel_row - 1, end_column=1)
        ws.merge_cells(start_row=excel_row - 2, start_column=2, end_row=excel_row - 1, end_column=2)

    # Column widths
    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 42
    for i in range(n_header_cols):
        from openpyxl.utils import get_column_letter; col_letter = get_column_letter(3 + i)
        ws.column_dimensions[col_letter].width = 32

    wb.save(output_path)
    print(f'Saved: {output_path}  ({len(records)} records, {len(header_keys)} header fields: {header_keys})')

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
