"""
================================================================================
Linked Object Scanner for Word Documents
================================================================================
Purpose:
    Scans all .docx files in DOCX_FOLDER for linked/embedded OLE objects
    (e.g. Excel sheets, PDFs) and checks whether each has a hyperlink URL.

Behaviour:
    - Objects WITH a hyperlink URL  → recorded in Excel report (URL shown)
    - Objects WITHOUT a hyperlink   → file extracted and saved to disk,
                                      saved path shown in Excel report

Output structure (created automatically inside OUTPUT_FOLDER):
    Output/
    ├── <DocxFileName>/               ← one subfolder per Word document
    │   ├── embedded_file1.xlsx
    │   ├── embedded_file2.pdf
    │   └── ...
    └── embedded_objects_report.xlsx  ← combined Excel report for all docs

Configuration:
    Edit the CONFIG block below — no command-line arguments needed.
    Run with:  python scan_linked_objects.py
================================================================================
"""

import os
import io
import re
import shutil
import zipfile
import urllib.parse
from datetime import datetime
from collections import Counter  # kept for potential future use
from lxml import etree
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


# ══════════════════════════════════════════════════════════════════════════════
# CONFIG  — edit these values if needed
# ══════════════════════════════════════════════════════════════════════════════
_HERE         = os.path.dirname(os.path.abspath(__file__))
DOCX_FOLDER   = os.path.join(_HERE, "Input")    # input folder (may contain subfolders)
OUTPUT_FOLDER = os.path.join(_HERE, "Output")   # root output folder
REPORT_NAME   = "embedded_objects_report.xlsx"  # Excel report filename
# ══════════════════════════════════════════════════════════════════════════════


# ── XML namespace URIs used inside .docx (Open XML standard) ─────────────────
NS_PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
NS_R   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS_W   = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'


# ── Excel report colour palette (matches original template) ──────────────────
C_HEADER_BG  = "1F4E79"   # dark navy  — header row background
C_HEADER_FG  = "FFFFFF"   # white      — header row text
C_BLUE_ROW   = "DCE6F1"   # light blue — row has a URL
C_YELLOW_ROW = "FFF2CC"   # yellow     — row has no URL (needs attention)
C_EXCEL_TYPE = "1F6B36"   # green      — Excel type badge
C_PDF_TYPE   = "C00000"   # red        — PDF type badge
C_OTHER_TYPE = "595959"   # grey       — other type badge
C_URL_FG     = "0563C1"   # blue       — hyperlink / path text
C_GREY_FG    = "7F7F7F"   # grey       — saved-path text


def _border():
    """Return a uniform thin border used on all data cells."""
    s = Side(style="thin", color="B8CCE4")
    return Border(left=s, right=s, top=s, bottom=s)


def _emf_display_name(emf_data):
    """
    Extract the display filename from an EMF (Enhanced Metafile) icon image.

    When Word embeds an OLE object as an icon, the icon is stored as an EMF
    file. The filename label shown under the icon is embedded in the EMF as
    UTF-16LE strings. Word may split a long name across multiple strings for
    word-wrapped display, but always also stores the full filename as a single
    string immediately before the sentinel value 'IconOnly'.

    Strategy:
      1. Collect all UTF-16LE strings before each 'IconOnly' occurrence.
      2. Among those candidates, prefer the longest one that contains a '.'
         (file extension) and no backslash (not a Windows path).

    Returns the full display filename, or '' if not found.
    """
    import re
    # Extract all UTF-16LE printable strings (3+ characters)
    strings = re.findall(b'(?:[\x20-\x7e]\x00){3,}', emf_data)
    decoded = [s.decode('utf-16-le') for s in strings]

    # Gather every candidate that appears before an 'IconOnly' marker
    candidates = []
    for i, s in enumerate(decoded):
        if s == 'IconOnly':
            # Scan backwards to find the best filename candidate
            for j in range(i - 1, max(i - 6, -1), -1):
                c = decoded[j]
                if c == 'IconOnly':
                    break
                # Must look like a filename: has extension dot, no path separator
                if '.' in c and '\\' not in c and len(c) < 260:
                    candidates.append(c)

    if not candidates:
        return ''

    # Return the longest candidate — Word stores both word-wrapped fragments
    # AND the full filename; the full one is always the longest
    return max(candidates, key=len)


# ══════════════════════════════════════════════════════════════════════════════
# STEP 1 — Parse a single .docx and return all OLE object metadata
# ══════════════════════════════════════════════════════════════════════════════
def parse_docx(docx_path):
    """
    Open a .docx (which is a ZIP archive) and extract metadata for every
    OLE linked/embedded object found in the document body.

    Returns a list of dicts, one per object, containing:
        source_docx   : basename of the source Word file
        object_file   : filename of the embedded object
        object_type   : 'Excel', 'PDF', 'Word', or other
        section       : nearest heading above the object (e.g. 'APPENDIX')
        row_label     : text of the table row containing the object
        hyperlink_url : URL if the object is wrapped in a hyperlink, else ''
        internal_path : path inside the docx zip (e.g. word/embeddings/...)
        ole_target    : raw OLE relationship target (file path or URL)
    """

    # ── Read the raw XML and file list from inside the .docx zip ─────────────
    with zipfile.ZipFile(docx_path) as z:
        doc_xml   = z.read("word/document.xml")             # main document body
        rels_xml  = z.read("word/_rels/document.xml.rels")  # relationship map
        all_names = z.namelist()                             # all files in zip
        # Pre-read all EMF icon files keyed by their relationship ID.
        # Each OLE object icon is an EMF image; the EMF contains the display
        # filename as a UTF-16LE string (extracted later via _emf_display_name).
        emf_data_by_rid = {
            name.split('/')[-1].replace('word/', ''): z.read(name)
            for name in all_names
            if name.startswith('word/media/') and name.lower().endswith('.emf')
        }

    # ── Build a dict of all relationships: {rId: {type, target}} ─────────────
    rels = {}
    for rel in etree.fromstring(rels_xml).findall(f'{{{NS_PKG}}}Relationship'):
        rels[rel.get('Id')] = {
            'type':   rel.get('Type', '').split('/')[-1],
            'target': rel.get('Target', ''),
        }

    # ── Separate hyperlinks for quick lookup: {rId: url} ─────────────────────
    hyperlinks = {
        rid: v['target']
        for rid, v in rels.items()
        if v['type'] == 'hyperlink'
    }

    # ── Parse the document XML into an element tree ───────────────────────────
    tree          = etree.fromstring(doc_xml)
    body          = tree.find(f'{{{NS_W}}}body')
    body_children = list(body) if body is not None else []

    # ── Build a map of ShapeID → display filename from EMF icon images ────────
    # Each OLE object has a VML <v:shape> whose <v:imagedata r:id="rIdXX"/>
    # points to an EMF image. That EMF encodes the filename label shown in Word.
    NS_V = 'urn:schemas-microsoft-com:vml'
    shape_display_names = {}   # {shape_id: display_filename}
    for shape in tree.iter(f'{{{NS_V}}}shape'):
        shape_id  = shape.get('id', '')
        imagedata = shape.find(f'{{{NS_V}}}imagedata')
        if imagedata is None:
            continue
        img_rid = imagedata.get(f'{{{NS_R}}}id', '')
        if not img_rid or img_rid not in rels:
            continue
        img_target = rels[img_rid].get('target', '')          # e.g. media/image4.emf
        img_file   = os.path.basename(img_target)             # e.g. image4.emf
        emf_bytes  = emf_data_by_rid.get(img_file, b'')
        if emf_bytes:
            name = _emf_display_name(emf_bytes)
            if name:
                shape_display_names[shape_id] = name

    def get_text(elem):
        """Concatenate all <w:t> text nodes under an element."""
        return ''.join(t.text or '' for t in elem.iter(f'{{{NS_W}}}t')).strip()

    def find_section(ole_elem):
        """
        Walk backwards through the document body to find the nearest
        heading paragraph above the OLE element.
        Returns the heading text in uppercase (e.g. 'APPENDIX').
        """
        def contains(container, target):
            return any(c is target for c in container.iter())

        # Find which top-level body child contains this OLE element
        idx = next(
            (i for i, c in enumerate(body_children) if contains(c, ole_elem)),
            None
        )
        if idx is None:
            return "DOCUMENT"

        # Scan backwards for a paragraph with a Heading style
        for i in range(idx - 1, -1, -1):
            child = body_children[i]
            if child.tag == f'{{{NS_W}}}p':
                pPr = child.find(f'{{{NS_W}}}pPr')
                if pPr is not None:
                    pStyle = pPr.find(f'{{{NS_W}}}pStyle')
                    if pStyle is not None and 'eading' in pStyle.get(f'{{{NS_W}}}val', ''):
                        txt = get_text(child)
                        if txt:
                            return txt.upper()
        return "DOCUMENT"

    def find_row_label(ole_elem):
        """
        Walk up the element tree to find the enclosing table row (<w:tr>)
        and return the text of its first non-empty cell.
        This gives context about where in a table the object sits.
        """
        parent = ole_elem.getparent()
        while parent is not None:
            if parent.tag.split('}')[-1] == 'tr':
                for tc in parent.findall(f'.//{{{NS_W}}}tc'):
                    txt = get_text(tc)
                    if txt:
                        return txt[:100]
                break
            parent = parent.getparent()
        return ""

    # ── Iterate over every OLEObject element in the document ─────────────────
    results = []
    for ole in tree.iter():
        if ole.tag.split('}')[-1] != 'OLEObject':
            continue

        rid  = ole.get(f'{{{NS_R}}}id') or ''
        prog = ole.get('ProgID', '')   # e.g. 'Excel.Sheet.12', 'AcroExch.Document.DC'

        if not rid or rid not in rels:
            continue

        raw     = rels[rid]['target']
        decoded = urllib.parse.unquote(raw)

        # ── Check whether the OLE target itself is a web URL ─────────────────
        # Some objects link via https:// in their OLE relationship directly
        # rather than through a <w:hyperlink> wrapper element
        ole_is_url = decoded.startswith('http://') or decoded.startswith('https://')

        # Strip file:/// prefix to get a plain local filesystem path
        if decoded.startswith('file:///'):
            decoded = decoded[8:]

        obj_file = os.path.basename(decoded)
        ext      = os.path.splitext(obj_file)[1].lower()

        # ── Determine the object type from ProgID or file extension ──────────
        # .bin files are OLE compound containers — use ProgID to get real type
        if ext == '.bin' or not ext:
            if 'Acro' in prog or 'PDF' in prog:
                obj_type = 'PDF'
            elif 'Excel' in prog:
                obj_type = 'Excel'
            elif 'Word' in prog:
                obj_type = 'Word'
            else:
                obj_type = 'Object'
        elif 'Excel' in prog or ext in ('.xlsx', '.xls', '.xlsm'):
            obj_type = 'Excel'
        elif 'Acro' in prog or ext == '.pdf':
            obj_type = 'PDF'
        elif 'Word' in prog or ext in ('.docx', '.doc'):
            obj_type = 'Word'
        else:
            obj_type = ext.upper().lstrip('.') or 'Object'

        # ── Locate the embedded copy inside the docx zip ─────────────────────
        # External URL objects have no embedded copy — skip entirely.
        internal = None
        if not ole_is_url:
            # Best match: the OLE relationship target already contains the exact
            # embedding path (e.g. 'embeddings/Microsoft_Excel_Worksheet2.xlsx').
            # Normalise to the full zip path (word/embeddings/...) and verify it
            # actually exists in the zip. This is always correct and avoids the
            # bug where multiple same-extension objects all grab the same file.
            rel_target = rels[rid]['target']   # e.g. 'embeddings/Microsoft_Excel_Worksheet2.xlsx'
            candidate  = 'word/' + rel_target.lstrip('/')   # e.g. 'word/embeddings/...'
            if candidate in all_names:
                internal = candidate

            # Fallback: PDFs may be stored as OLE .bin compound container files.
            # Their OLE target may itself end in .bin, or the matching zip entry does.
            if not internal and obj_type == 'PDF':
                for name in all_names:
                    if 'embeddings' in name and name.lower().endswith('.bin'):
                        internal = name
                        break

            # Last resort: match by file extension (covers edge cases where the
            # rel_target path uses an unexpected prefix)
            if not internal:
                for name in all_names:
                    if 'embeddings' in name and name.lower().endswith(ext):
                        internal = name
                        break

        # ── Check whether this OLE element is wrapped in a <w:hyperlink> ─────
        url = ''
        p   = ole.getparent()
        while p is not None:
            if p.tag.split('}')[-1] == 'hyperlink':
                url = hyperlinks.get(p.get(f'{{{NS_R}}}id', ''), '')
                break
            p = p.getparent()

        # If no hyperlink wrapper but the OLE target itself is a URL, use it
        if not url and ole_is_url:
            url = urllib.parse.unquote(raw)

        # ── Resolve the display filename ─────────────────────────────────────
        # Priority: (1) EMF icon label (what Word shows) →
        #           (2) basename from the OLE target path
        shape_id     = ole.get('ShapeID', '')
        display_name = shape_display_names.get(shape_id, '') or obj_file

        results.append({
            'source_docx':   os.path.basename(docx_path),
            'object_file':   display_name,     # human-readable display filename
            'object_type':   obj_type,
            'section':       find_section(ole),
            'row_label':     find_row_label(ole),
            'hyperlink_url': url,
            'internal_path': internal,
            'ole_target':    urllib.parse.unquote(raw),
        })

    return results


def _fix_hidden_workbook(xlsx_path):
    # Word embeds Excel files with several OLE-specific flags that prevent them
    # from opening correctly as standalone files:
    #   1. visibility="hidden"  — hides the workbook window while embedded
    #   2. missing activeTab    — causes Excel to show no active sheet (blank)
    #   3. <oleSize>            — marks the file as an OLE sub-view; Excel
    #                             refuses to open it normally as a standalone file
    # This function removes all three issues so the file opens cleanly.
    try:
        with open(xlsx_path, 'rb') as f:
            original = f.read()

        buf = io.BytesIO(original)
        with zipfile.ZipFile(buf, 'r') as zin:
            if 'xl/workbook.xml' not in zin.namelist():
                return
            wb_xml = zin.read('xl/workbook.xml')

        # Check if any patching is needed
        needs_patch = (
            b'visibility="hidden"' in wb_xml or
            b'activeTab=' not in wb_xml or
            b'<oleSize' in wb_xml
        )
        if not needs_patch:
            return

        patched = wb_xml

        # 1. Remove visibility="hidden" from <workbookView>
        patched = patched.replace(b' visibility="hidden"', b'')
        patched = patched.replace(b'visibility="hidden"', b'')

        # 2. Add activeTab="0" if missing so first sheet is shown on open
        if b'activeTab=' not in patched:
            patched = patched.replace(b'<workbookView', b'<workbookView activeTab="0"')

        # 3. Remove <oleSize .../> — this tag tells Excel the workbook is an
        #    OLE embedded sub-view and prevents it opening as a normal file.
        #    It appears as a self-closing tag, e.g. <oleSize ref="A1:P89"/>
        while b'<oleSize' in patched:
            start = patched.find(b'<oleSize')
            end   = patched.find(b'>', start) + 1
            patched = patched[:start] + patched[end:]

        # Rebuild the xlsx zip with the patched workbook.xml
        out_buf = io.BytesIO()
        buf.seek(0)
        with zipfile.ZipFile(buf, 'r') as zin, \
             zipfile.ZipFile(out_buf, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = patched if item.filename == 'xl/workbook.xml' else zin.read(item.filename)
                zout.writestr(item, data)

        with open(xlsx_path, 'wb') as f:
            f.write(out_buf.getvalue())

    except Exception:
        pass   # non-critical — leave file as-is if patching fails


def _verify_file(path):
    """
    Verify that an extracted file can actually be opened and contains data.
    Returns (ok: bool, message: str).

    Checks performed:
      - File exists and has non-zero size
      - For .xlsx/.xls/.xlsm : valid zip, workbook.xml readable, has data cells
      - For .pdf             : starts with %PDF magic bytes
      - For other types      : file is non-empty
    """
    if not os.path.isfile(path):
        return False, "file not found after extraction"

    size = os.path.getsize(path)
    if size == 0:
        return False, "file is empty (0 bytes)"

    ext = os.path.splitext(path)[1].lower()

    # ── Excel verification ────────────────────────────────────────────────────
    if ext in ('.xlsx', '.xlsm', '.xls'):
        try:
            with zipfile.ZipFile(path) as z:
                # Check zip integrity
                bad = z.testzip()
                if bad:
                    return False, f"corrupt zip entry: {bad}"

                # Must have workbook.xml
                if 'xl/workbook.xml' not in z.namelist():
                    return False, "missing xl/workbook.xml"

                wb_xml = z.read('xl/workbook.xml')

                # Must not still be hidden
                if b'visibility="hidden"' in wb_xml:
                    return False, "workbook still has visibility=hidden"

                # Must have at least one sheet with data cells
                total_cells = 0
                for name in z.namelist():
                    if 'xl/worksheets/sheet' in name:
                        sheet_data  = z.read(name)
                        total_cells += len(re.findall(b'<c ', sheet_data))

                if total_cells == 0:
                    return False, "no data cells found in any sheet"

            return True, f"OK ({size:,} bytes, {total_cells} cells)"

        except zipfile.BadZipFile:
            return False, "not a valid zip/xlsx file"
        except Exception as e:
            return False, f"verification error: {e}"

    # ── PDF verification ──────────────────────────────────────────────────────
    elif ext == '.pdf':
        try:
            with open(path, 'rb') as f:
                header = f.read(8)
            if not header.startswith(b'%PDF'):
                return False, f"invalid PDF header: {header[:8].hex()}"
            return True, f"OK ({size:,} bytes)"
        except Exception as e:
            return False, f"verification error: {e}"

    # ── Generic check ─────────────────────────────────────────────────────────
    else:
        return True, f"OK ({size:,} bytes)"


# ══════════════════════════════════════════════════════════════════════════════
# STEP 2 — Extract embedded files into a per-docx subfolder
# ══════════════════════════════════════════════════════════════════════════════
def extract_files(objects, docx_path):
    """
    For each OLE object that has NO hyperlink URL, extract the embedded file
    from inside the .docx zip and save it to:

        Output/<DocxFileNameWithoutExtension>/<embedded_filename>

    Updates each object dict in-place with:
        saved_path : absolute path of the saved file (or error/empty string)

    Objects that already have a hyperlink URL are skipped (saved_path = '').
    """

    # Create a subfolder named after the Word document (without .docx extension)
    docx_stem   = os.path.splitext(os.path.basename(docx_path))[0]
    doc_out_dir = os.path.join(OUTPUT_FOLDER, docx_stem)
    os.makedirs(doc_out_dir, exist_ok=True)

    with zipfile.ZipFile(docx_path) as z:
        for obj in objects:

            # ── Skip objects that already have a hyperlink URL ────────────────
            if obj['hyperlink_url']:
                obj['saved_path'] = ''
                continue

            # ── Case 1: Object is embedded inside the docx zip ────────────────
            if obj['internal_path']:

                # object_file is already the display name resolved in parse_docx
                # (from the EMF icon label). For .bin containers that had no EMF
                # display name, fall back to deriving from the OLE target path.
                out_name = obj['object_file']
                if obj['internal_path'].endswith('.bin') and out_name.lower().endswith('.bin'):
                    # Still a raw .bin name — try OLE target path as last resort
                    ole_tgt   = urllib.parse.unquote(obj.get('ole_target', ''))
                    candidate = ole_tgt.replace('file:///', '').lstrip('/')
                    base_name = os.path.basename(candidate)
                    if base_name and not base_name.lower().endswith('.bin'):
                        out_name = base_name
                    else:
                        type_ext = '.pdf' if obj['object_type'] == 'PDF' else '.bin'
                        out_name = f"{docx_stem}_embedded{type_ext}"

                # Build the destination path; append a timestamp if a file with
                # the same name already exists (prevents silent overwrites)
                dest = os.path.join(doc_out_dir, out_name)
                if os.path.exists(dest):
                    base, e = os.path.splitext(out_name)
                    dest = os.path.join(
                        doc_out_dir,
                        f"{base}_{datetime.now().strftime('%H%M%S%f')}{e}"
                    )

                try:
                    raw_data = z.read(obj['internal_path'])

                    # PDFs embedded as OLE .bin compound containers:
                    # locate the actual PDF stream by searching for the
                    # %PDF magic bytes and the %%EOF end marker
                    if obj['internal_path'].endswith('.bin') and obj['object_type'] == 'PDF':
                        pdf_start = raw_data.find(b'%PDF')
                        pdf_end   = raw_data.rfind(b'%%EOF')
                        if pdf_start >= 0 and pdf_end >= 0:
                            raw_data = raw_data[pdf_start:pdf_end + 5]

                    with open(dest, 'wb') as f:
                        f.write(raw_data)

                    # Fix Excel files that Word stored as hidden workbooks —
                    # removes visibility="hidden" so the file opens correctly
                    if dest.lower().endswith(('.xlsx', '.xls', '.xlsm')):
                        _fix_hidden_workbook(dest)

                    abs_dest = os.path.abspath(dest)

                    # Verify the extracted file is valid and openable
                    ok, msg = _verify_file(abs_dest)
                    if ok:
                        obj['saved_path']  = abs_dest
                        obj['verify_status'] = f'✔ {msg}'
                    else:
                        obj['saved_path']  = abs_dest
                        obj['verify_status'] = f'✘ VERIFY FAILED: {msg}'

                    obj['object_file'] = out_name  # update display name to match

                except Exception as ex:
                    obj['saved_path']    = f'[error: {ex}]'
                    obj['verify_status'] = f'✘ extraction error: {ex}'

            # ── Case 2: Object is an external file reference (not in zip) ─────
            else:
                ole_target = obj.get('ole_target', '')

                # Convert file:/// URI to a plain local filesystem path
                local_path = ole_target
                if local_path.startswith('file:///'):
                    local_path = '/' + local_path[8:]
                elif local_path.startswith('file://'):
                    local_path = local_path[7:]

                if local_path and os.path.isfile(local_path):
                    # File exists on disk — copy it into the per-docx subfolder
                    dest = os.path.join(doc_out_dir, obj['object_file'])
                    if os.path.exists(dest):
                        base, e = os.path.splitext(obj['object_file'])
                        dest = os.path.join(
                            doc_out_dir,
                            f"{base}_{datetime.now().strftime('%H%M%S%f')}{e}"
                        )
                    try:
                        shutil.copy2(local_path, dest)
                        abs_dest = os.path.abspath(dest)
                        ok, msg  = _verify_file(abs_dest)
                        obj['saved_path']    = abs_dest
                        obj['verify_status'] = f'✔ {msg}' if ok else f'✘ VERIFY FAILED: {msg}'
                    except Exception as ex:
                        obj['saved_path']    = f'[copy error: {ex}]'
                        obj['verify_status'] = f'✘ copy error: {ex}'
                else:
                    # File not found on disk — leave saved_path empty so the
                    # URL column in the report falls back to showing ole_target
                    obj['saved_path'] = ''


# ══════════════════════════════════════════════════════════════════════════════
# STEP 3 — Write the combined Excel report for all documents
# ══════════════════════════════════════════════════════════════════════════════
def write_report(all_objects):
    """
    Build a formatted Excel workbook from scratch and save it to:
        Output/embedded_objects_report.xlsx

    Sheet 'Embedded Objects' columns:
        A: Word Document  — source .docx filename
        B: Word Section   — nearest heading in the document
        C: Row Label      — table row context text
        D: Object Type    — Excel / PDF / Word
        E: URL            — hyperlink URL  OR  relative path from Output folder

    Formatting:
        Header row  : bold, grey background
        Data rows   : white background, black borders
    """

    wb = Workbook()
    ws = wb.active
    ws.title = "Embedded Objects"

    # Styling constants
    GREY_HEADER_BG = "808080"   # grey header background
    WHITE_BG       = "FFFFFF"   # white data rows
    BLACK_BORDER   = "000000"   # black cell borders
    HEADER_FG      = "FFFFFF"   # white header text
    URL_COLOR      = "0563C1"   # blue for URL/path text

    def _cell_border():
        """Thin black border for all data cells."""
        s = Side(style="thin", color=BLACK_BORDER)
        return Border(left=s, right=s, top=s, bottom=s)

    # ── Column definitions: (header label, column width) ─────────────────────
    columns = [
        ("Word Document", 35),   # A
        ("Word Section",  20),   # B
        ("Row Label",     55),   # C
        ("Object Type",   14),   # D
        ("URL",           70),   # E
    ]

    # ── Header row (row 1) — bold, grey background ────────────────────────────
    for ci, (header, width) in enumerate(columns, 1):
        c = ws.cell(row=1, column=ci, value=header)
        c.font      = Font(name="Arial", size=11, bold=True, color=HEADER_FG)
        c.fill      = PatternFill("solid", start_color=GREY_HEADER_BG)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border    = _cell_border()
        ws.column_dimensions[c.column_letter].width = width
    ws.row_dimensions[1].height = 25

    # ── Helper: write a data cell with white background and black border ──────
    def dcell(r, c, val, bold=False, color=None, wrap=False):
        cell = ws.cell(row=r, column=c, value=val)
        cell.font      = Font(name="Arial", size=10, bold=bold, color=color or BLACK_BORDER)
        cell.fill      = PatternFill("solid", start_color=WHITE_BG)
        cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=wrap)
        cell.border    = _cell_border()
        return cell

    # ── Helper: shorten full saved path to relative "..Output/..." form ───────
    def short_path(full_path):
        """
        Convert an absolute saved path to a short relative form starting from
        the Output folder, e.g.:
            /home/user/Output/DocName/file.xlsx  →  ..Output/DocName/file.xlsx
        """
        # Normalise separators
        norm = full_path.replace('\\', '/')
        out  = OUTPUT_FOLDER.replace('\\', '/')
        out_name = os.path.basename(out)   # e.g. "Output"
        # Find '.../<OutputFolderName>/...' in the path and truncate before it
        marker = f'/{out_name}/'
        idx = norm.find(marker)
        if idx >= 0:
            return '..' + norm[idx:]       # e.g. ..Output/DocName/file.xlsx
        return full_path                   # fallback: return as-is

    # ── Data rows (row 2 onwards, one row per embedded object) ────────────────
    for i, obj in enumerate(all_objects):
        r       = 2 + i
        has_url = bool(obj['hyperlink_url'])

        dcell(r, 1, obj['source_docx'])           # A: Word Document
        dcell(r, 2, obj['section'])                # B: Word Section
        dcell(r, 3, obj['row_label'], wrap=True)   # C: Row Label
        dcell(r, 4, obj['object_type'])            # D: Object Type

        # ── Column E: URL or relative saved path ─────────────────────────────
        if has_url:
            # Object has a real hyperlink URL — show it in blue
            dcell(r, 5, obj['hyperlink_url'], color=URL_COLOR, wrap=True)
        else:
            saved    = obj.get('saved_path', '')
            original = obj.get('ole_target', '')

            if saved and not saved.startswith('['):
                # Successfully extracted — show short relative path
                dcell(r, 5, short_path(saved), color=URL_COLOR, wrap=True)
            elif original:
                # Not extracted — show original OLE reference
                dcell(r, 5, original, color=URL_COLOR, wrap=True)
            else:
                dcell(r, 5, 'Path not available', wrap=True)

        ws.row_dimensions[r].height = 40 if len(obj.get('row_label', '')) > 40 else 18

    total = len(all_objects)

    # Freeze header and enable auto-filter on all 5 columns
    ws.freeze_panes    = "A2"
    ws.auto_filter.ref = f"A1:E{1 + total}"

    out_path = os.path.join(OUTPUT_FOLDER, REPORT_NAME)
    wb.save(out_path)
    return out_path


# ══════════════════════════════════════════════════════════════════════════════
# MAIN — orchestrate scanning, extraction and reporting
# ══════════════════════════════════════════════════════════════════════════════
def main():
    """
    Entry point:
      1. Find all .docx files in DOCX_FOLDER
      2. Parse each document for OLE linked/embedded objects
      3. Extract files (those without a URL) into Output/<DocxName>/ subfolders
      4. Write a combined Excel report to Output/embedded_objects_report.xlsx
    """

    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    # ── Discover all .docx files recursively (skip Word temp files ~$...) ──────
    docx_files = sorted([
        os.path.join(root, f)
        for root, _, files in os.walk(DOCX_FOLDER)
        for f in files
        if f.lower().endswith('.docx') and not f.startswith('~')
    ])

    if not docx_files:
        print(f"\n[!] No .docx files found in: {DOCX_FOLDER}")
        return

    print(f"\nFound {len(docx_files)} .docx file(s) in: {DOCX_FOLDER}")
    for f in docx_files:
        print(f"  • {os.path.basename(f)}")

    all_objects = []   # accumulates parsed results from all documents

    # ── Process each Word document one by one ─────────────────────────────────
    for idx, docx_path in enumerate(docx_files, 1):
        docx_name = os.path.basename(docx_path)
        print(f"\n[{idx}/{len(docx_files)}] Processing: {docx_name}")

        # Parse OLE objects from the document
        try:
            objects = parse_docx(docx_path)
            print(f"      {len(objects)} linked object(s) found")
        except Exception as ex:
            print(f"      [!] Failed to parse: {ex}")
            continue

        # Extract embedded files into Output/<DocxName>/ subfolder
        extract_files(objects, docx_path)

        # Log which files were saved (only those without a URL)
        no_link     = [o for o in objects if not o['hyperlink_url']]
        docx_stem   = os.path.splitext(docx_name)[0]
        doc_out_dir = os.path.join(OUTPUT_FOLDER, docx_stem)
        print(f"      {len(no_link)} file(s) without URL → saved to: {doc_out_dir}")
        for o in no_link:
            status = o.get('verify_status', '')
            print(f"        • {o['object_file']}  →  {o.get('saved_path', '')}  [{status}]")

        all_objects.extend(objects)

    # ── Write the combined Excel report covering all documents ────────────────
    if all_objects:
        report = write_report(all_objects)
        print(f"\nExcel report saved → {report}")
    else:
        print("\n[!] No linked objects found across all documents.")

    # ── Print final summary ───────────────────────────────────────────────────
    total    = len(all_objects)
    with_url = sum(1 for o in all_objects if o['hyperlink_url'])
    print(f"\n{'─' * 50}")
    print(f"  Documents scanned : {len(docx_files)}")
    print(f"  Total objects     : {total}")
    print(f"  With URL          : {with_url}")
    print(f"  No URL            : {total - with_url}")
    print(f"{'─' * 50}\n")


if __name__ == "__main__":
    main()
