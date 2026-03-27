"""
Hyperlink Extractor for Word Documents (.docx / .doc)
Single sheet 'Embedded Objects' — visual grouping via borders only (filterable).
URLs: http/https/mailto/ftp only. Links not starting with http are excluded.

.docx -> XML extraction (section, table, row, column context)
.doc  -> Pure-Python OLE CFB reader (stdlib struct only, zero pip installs)
"""

from pathlib import Path
from itertools import groupby
import zipfile, struct, re

from lxml import etree
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Paths ──────────────────────────────────────────────────────────────────
INPUT_FOLDER = Path(__file__).parent / 'input'
OUTPUT_PATH  = Path(__file__).parent / 'output' / 'Hyperlink_extract_formatted.xlsx'

# ── XML namespaces ─────────────────────────────────────────────────────────
NS = {
    'w':   'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r':   'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
}
HEADING_STYLES = {f'heading{i}' for i in range(1, 7)}
SECTION_STYLES = HEADING_STYLES | {'title', 'subtitle'}

# ── OLE CFB constants ──────────────────────────────────────────────────────
OLE_MAGIC = b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'
FREESECT  = 0xFFFFFFFF
ENDOFCHAIN= 0xFFFFFFFE
DIFSECT   = 0xFFFFFFFC
FATSECT   = 0xFFFFFFFD

# Only keep URLs that start with http (http/https)
URL_PAT = re.compile(
    r'(https?://[^\x00-\x08\x0a-\x1f]{5,})'
)

# ══════════════════════════════════════════════════════════════════════════
#  Pure-Python OLE reader  (stdlib struct, no pip)
# ══════════════════════════════════════════════════════════════════════════

def _read_ole_streams(path):
    with open(path, 'rb') as fh:
        data = fh.read()
    if data[:8] != OLE_MAGIC:
        return {}
    sec_size      = 1 << struct.unpack_from('<H', data, 30)[0]
    mini_ss       = 1 << struct.unpack_from('<H', data, 32)[0]
    dir_start     = struct.unpack_from('<I', data, 48)[0]
    mini_cutoff   = struct.unpack_from('<I', data, 56)[0]
    minifat_start = struct.unpack_from('<I', data, 60)[0]
    header_difat  = list(struct.unpack_from('<109I', data, 76))
    fat_sids      = [s for s in header_difat
                     if s not in (FREESECT, ENDOFCHAIN, DIFSECT, FATSECT)]

    def sec_off(sid): return 512 + sid * sec_size

    fat = []
    for sid in fat_sids:
        fat.extend(struct.unpack_from('<' + 'I' * (sec_size // 4), data, sec_off(sid)))

    def read_chain(start, raw=data):
        out, sec, seen = [], start, set()
        while sec not in (ENDOFCHAIN, FREESECT) and sec < len(fat):
            if sec in seen: break
            seen.add(sec)
            out.append(raw[sec_off(sec):sec_off(sec) + sec_size])
            sec = fat[sec]
        return b''.join(out)

    dir_data    = read_chain(dir_start)
    minifat     = []
    if minifat_start not in (ENDOFCHAIN, FREESECT):
        mf      = read_chain(minifat_start)
        minifat = list(struct.unpack_from('<' + 'I' * (len(mf) // 4), mf))

    mini_stream = read_chain(struct.unpack_from('<I', dir_data, 116)[0])

    def read_mini(start):
        out, sec, seen = [], start, set()
        while sec not in (ENDOFCHAIN, FREESECT) and sec < len(minifat):
            if sec in seen: break
            seen.add(sec)
            o = sec * mini_ss
            out.append(mini_stream[o:o + mini_ss])
            sec = minifat[sec]
        return b''.join(out)

    streams = {}
    for i in range(len(dir_data) // 128):
        e   = dir_data[i * 128:(i + 1) * 128]
        nl  = struct.unpack_from('<H', e, 64)[0]
        typ = e[66]
        ss  = struct.unpack_from('<I', e, 116)[0]
        sz  = struct.unpack_from('<I', e, 120)[0]
        if typ == 2 and nl > 0 and ss not in (ENDOFCHAIN, FREESECT):
            name = e[:nl - 2].decode('utf-16-le', errors='replace')
            streams[name.lower()] = (
                read_mini(ss) if sz < mini_cutoff else read_chain(ss)
            )[:sz]
    return streams


def extract_doc_urls(path):
    """Extract http/https URLs from .doc via pure-Python OLE reader."""
    streams = _read_ole_streams(path)
    seen, urls = set(), []
    for sname in ('data', 'worddocument', '1table'):
        raw = streams.get(sname)
        if not raw:
            continue
        try:
            text = raw.decode('utf-16-le', errors='replace')
        except Exception:
            continue
        for tok in re.split(r'[\x00-\x08\x0b\x0c\x0e-\x1f]+', text):
            m = URL_PAT.match(tok.strip())
            if m:
                url = m.group(1).rstrip('\x15\x14\x13').strip()
                if url not in seen:
                    seen.add(url)
                    urls.append(url)
    return urls


# ══════════════════════════════════════════════════════════════════════════
#  DOCX extractor
# ══════════════════════════════════════════════════════════════════════════

def _txt(el):
    return ' '.join(
        (t.text or '').strip()
        for t in el.findall('.//{%s}t' % NS['w'])
        if (t.text or '').strip()
    )

def _load_rels(z):
    rels = {}
    try:
        with z.open('word/_rels/document.xml.rels') as f:
            tree = etree.parse(f)
        for rel in tree.findall('{%s}Relationship' % NS['rel']):
            if 'hyperlink' in rel.get('Type', ''):
                rels[rel.get('Id')] = rel.get('Target', '')
    except KeyError:
        pass
    return rels


def _build_num_map(z, body):
    """
    Build a map {para_element_id -> number_string} for every paragraph
    that carries a w:numPr (auto-numbered by Word).

    Reads word/numbering.xml to get each list's start value, then walks
    all paragraphs in document order counting increments — exactly what
    Word renders on screen.  Only decimal / ordinal formats are supported
    (sufficient for section headings numbered 1, 2, 3…).

    Returns dict keyed by id(para_element).
    """
    # ── Read abstractNum start values from numbering.xml ──────────────
    abs_starts = {}   # abstractNumId -> {ilvl -> int}
    num_to_abs = {}   # numId         -> abstractNumId

    try:
        with z.open('word/numbering.xml') as f:
            num_tree = etree.parse(f)
    except KeyError:
        return {}

    for an in num_tree.findall('{%s}abstractNum' % NS['w']):
        anid = an.get('{%s}abstractNumId' % NS['w'])
        abs_starts[anid] = {}
        for lvl in an.findall('{%s}lvl' % NS['w']):
            ilvl     = lvl.get('{%s}ilvl' % NS['w'])
            start_el = lvl.find('{%s}start' % NS['w'])
            fmt_el   = lvl.find('{%s}numFmt' % NS['w'])
            fmt      = fmt_el.get('{%s}val' % NS['w']) if fmt_el is not None else 'decimal'
            start    = int(start_el.get('{%s}val' % NS['w'])) if start_el is not None else 1
            # Only track decimal/ordinal lists (section numbers)
            if fmt in ('decimal', 'decimalZero', 'ordinal', 'cardinalText'):
                abs_starts[anid][ilvl] = start

    for num in num_tree.findall('{%s}num' % NS['w']):
        nid   = num.get('{%s}numId' % NS['w'])
        anid_el = num.find('{%s}abstractNumId' % NS['w'])
        if anid_el is not None:
            anid = anid_el.get('{%s}val' % NS['w'])
            num_to_abs[nid] = anid
        # Apply startOverride if present
        for ov in num.findall('{%s}lvlOverride' % NS['w']):
            ilvl_ov = ov.get('{%s}ilvl' % NS['w'])
            so = ov.find('{%s}startOverride' % NS['w'])
            if so is not None:
                anid_v = num_to_abs.get(nid)
                if anid_v and anid_v in abs_starts:
                    abs_starts[anid_v][ilvl_ov] = int(so.get('{%s}val' % NS['w']))

    # ── Walk body direct children — same element objects as main extract loop ──
    # Using the element itself as dict key (lxml creates different Python wrapper
    # objects for the same XML node on each traversal, making id() unreliable).
    counters = {}   # (numId, ilvl) -> current int value
    num_map  = {}   # element -> number string

    for para in body:   # direct children only — mirrors main extract_from_docx loop
        if para.tag.split('}')[-1] != 'p':
            continue
        pPr   = para.find('{%s}pPr' % NS['w'])
        numPr = pPr.find('{%s}numPr' % NS['w']) if pPr is not None else None
        if numPr is None:
            continue
        nid_el  = numPr.find('{%s}numId' % NS['w'])
        ilvl_el = numPr.find('{%s}ilvl' % NS['w'])
        nid  = nid_el.get('{%s}val' % NS['w'])  if nid_el  is not None else None
        ilvl = ilvl_el.get('{%s}val' % NS['w']) if ilvl_el is not None else '0'
        if not nid or nid == '0':
            continue
        anid = num_to_abs.get(nid)
        if not anid or ilvl not in abs_starts.get(anid, {}):
            continue

        key = (nid, ilvl)
        counters[key] = counters[key] + 1 if key in counters else abs_starts[anid][ilvl]
        num_map[para] = str(counters[key])   # element as key — always same object

    return num_map

def _pstyle(para):
    ps = para.find('{%s}pPr/{%s}pStyle' % (NS['w'], NS['w']))
    return (ps.get('{%s}val' % NS['w']) or '').lower() if ps is not None else ''

def _is_heading_style(para):
    """True if paragraph has a Heading / Title / Subtitle style."""
    style = _pstyle(para)
    return style in SECTION_STYLES or 'heading' in style

def _is_header_row(row):
    """
    True if a table row is a true header row (column labels), not a data row.
    Checks three signals — any one is sufficient:
      1. w:tblHeader marker in trPr  (explicit Word repeat-header flag)
      2. ALL cells have a non-trivial background shading (header fill colour)
      3. ALL non-empty cells are bold (bold labels = header pattern)
    """
    trPr = row.find('{%s}trPr' % NS['w'])
    if trPr is not None and trPr.find('{%s}tblHeader' % NS['w']) is not None:
        return True
    cells = row.findall('{%s}tc' % NS['w'])
    if not cells:
        return False
    NON_HEADER_FILLS = {None, 'auto', 'ffffff', 'FFFFFF', '000000', '00000000'}
    shadings = [
        (cell.find('.//{%s}shd' % NS['w']) is not None and
         cell.find('.//{%s}shd' % NS['w']).get('{%s}fill' % NS['w']) not in NON_HEADER_FILLS)
        for cell in cells
    ]
    if all(shadings):
        return True
    non_empty = [(c, _txt(c)) for c in cells if _txt(c)]
    if non_empty and all(bool(c.findall('.//{%s}b' % NS['w'])) for c, _ in non_empty):
        return True
    return False


def _hl_urls(el, rels):
    """
    Return all http/https URLs from an element via two mechanisms:

    1. <w:hyperlink r:id="rIdN"> — standard relationship-based hyperlinks.
    2. Runs with <w:rStyle w:val="Hyperlink"> — plain runs styled as hyperlinks
       with the URL as literal text (no relationship). Consecutive Hyperlink
       runs are concatenated to handle URLs split across multiple runs.
    """
    import re as _re
    URL_RE = _re.compile(r'https?://\S+')
    seen = set()
    urls = []

    def _add(url):
        url = url.strip()
        if url.startswith('http') and url not in seen:
            seen.add(url)
            urls.append(url)

    # ── 1. Relationship-based <w:hyperlink> elements ──────────────────────
    for hl in el.findall('.//{%s}hyperlink' % NS['w']):
        rid    = hl.get('{%s}id' % NS['r'])
        anchor = hl.get('{%s}anchor' % NS['w'])
        if anchor and anchor.startswith('_Toc'):
            continue
        if rid and rid in rels:
            _add(rels[rid])

    # ── 2. Runs styled with the "Hyperlink" character style ───────────────
    # Iterate paragraph by paragraph so consecutive runs stay together.
    for para in el.findall('.//{%s}p' % NS['w']):
        # Skip paragraphs that are already fully inside a w:hyperlink element
        # (those URLs were captured above)
        if para.getparent() is not None:
            parent_tag = para.getparent().tag.split('}')[-1]
            if parent_tag == 'hyperlink':
                continue

        runs = para.findall('{%s}r' % NS['w'])
        i = 0
        while i < len(runs):
            run = runs[i]
            rStyle = run.find('{%s}rPr/{%s}rStyle' % (NS['w'], NS['w']))
            is_hl_style = (
                rStyle is not None and
                'hyperlink' in (rStyle.get('{%s}val' % NS['w']) or '').lower()
            )
            if is_hl_style:
                # Concatenate all consecutive Hyperlink-styled runs
                parts = []
                while i < len(runs):
                    r2 = runs[i]
                    rs2 = r2.find('{%s}rPr/{%s}rStyle' % (NS['w'], NS['w']))
                    if rs2 is not None and 'hyperlink' in (rs2.get('{%s}val' % NS['w']) or '').lower():
                        t = r2.find('{%s}t' % NS['w'])
                        parts.append((t.text or '') if t is not None else '')
                        i += 1
                    else:
                        break
                full_text = ''.join(parts).strip()
                if URL_RE.match(full_text):
                    _add(full_text)
            else:
                i += 1

    return urls


def extract_from_docx(docx_path):
    records = []
    with zipfile.ZipFile(docx_path) as z:
        rels    = _load_rels(z)
        with z.open('word/document.xml') as f:
            tree = etree.parse(f)
        body = tree.find('{%s}body' % NS['w'])
        if body is None:
            return records
        num_map = _build_num_map(z, body)   # para id -> resolved number string

    body = tree.find('{%s}body' % NS['w'])
    if body is None:
        return records

    # Section detection — document-agnostic, no hardcoding:
    #
    # Rule 1: Any Heading-styled paragraph always updates the current section.
    # Rule 2: Before the first heading appears, the last non-empty paragraph
    #         seen before each table is used as the section label (covers
    #         "Program Details", "Document Details" etc. which authors write
    #         as plain bold paragraphs outside any heading style).
    # Once a Heading is seen it becomes "sticky" — body text after that point
    # no longer overrides the section label, preventing list items, captions,
    # and URL-only paragraphs from becoming spurious section names.
    children = list(body)
    heading_seen = False
    current_section = ''
    last_nonempty_para = ''

    for idx, child in enumerate(children):
        tag = child.tag.split('}')[-1]

        if tag == 'p':
            content = _txt(child)
            if content:
                last_nonempty_para = content
            if content and _is_heading_style(child):
                # Prepend the auto-number if Word generated one via numPr
                # and it isn't already present at the start of the run text.
                auto_num = num_map.get(child)
                if auto_num:
                    import re as _re
                    # Only prepend if content doesn't already start with that number
                    if not _re.match(r'^' + re.escape(auto_num) + r'[\s\.]', content):
                        content = auto_num + ' ' + content
                current_section = content
                heading_seen = True

        if tag == 'tbl':
            # Before the first heading, use the last non-empty paragraph as section
            if not heading_seen and last_nonempty_para:
                current_section = last_nonempty_para

            rows = child.findall('{%s}tr' % NS['w'])

            # Detect header row: use row 0 only if it is a true label row.
            # Tables like "Program Details" have no header — row 0 is data.
            header_labels = []
            data_start = 0
            if rows and _is_header_row(rows[0]):
                header_labels = [_txt(tc)
                                 for tc in rows[0].findall('{%s}tc' % NS['w'])]
                data_start = 1   # skip header when assigning row_label
            else:
                # No internal header — look back for a preceding single-row
                # shaded title table (e.g. "Program/Project Name | FA Manage Asset Disposal")
                # whose cell values serve as the column labels for this data table.
                for j in range(idx - 1, -1, -1):
                    prev_tag = children[j].tag.split('}')[-1]
                    prev_txt = _txt(children[j])
                    if prev_tag == 'tbl':
                        prev_rows = children[j].findall('{%s}tr' % NS['w'])
                        if prev_rows and _is_header_row(prev_rows[0]):
                            header_labels = [_txt(tc)
                                             for tc in prev_rows[0].findall('{%s}tc' % NS['w'])]
                        break
                    if prev_txt:   # non-empty paragraph stops lookback
                        break

            for row_idx, row in enumerate(rows):
                cells     = row.findall('{%s}tc' % NS['w'])
                row_label = _txt(cells[0]) if cells else ''

                for col_idx, cell in enumerate(cells):
                    for url in _hl_urls(cell, rels):
                        records.append({
                            'section':       current_section,
                            'table_flag':    'Within Table',
                            'row_label':     row_label,
                            'row_label_col': header_labels[0] if header_labels else '',
                            'url_col':       (header_labels[col_idx]
                                             if col_idx < len(header_labels)
                                             else get_column_letter(col_idx + 1)),
                            'url':           url,
                        })

        elif tag == 'p':
            for url in _hl_urls(child, rels):
                records.append({
                    'section':       current_section,
                    'table_flag':    'Outside Table',
                    'row_label':     '',
                    'row_label_col': '',
                    'url_col':       '',
                    'url':           url,
                })

    return records


# ══════════════════════════════════════════════════════════════════════════
#  Excel writer — single sheet, border-only visual grouping (filterable)
# ══════════════════════════════════════════════════════════════════════════

# ── Style constants matching sample exactly ────────────────────────────────
HDR_FONT  = Font(name='Calibri', size=11, bold=True,  color='FFFFFFFF')
HDR_FILL  = PatternFill('solid', start_color='FF4472C4')
HDR_ALIGN = Alignment(horizontal='center', vertical='center', wrap_text=True)

DAT_FONT  = Font(name='Calibri', size=10, color='FF000000')
DAT_ALIGN = Alignment(horizontal='left', vertical='center', wrap_text=True)
WHITE_FILL= PatternFill('solid', start_color='FFFFFFFF')
NO_FILL   = PatternFill(fill_type=None)

THIN_BLACK  = Side(style='thin',   color='FF000000')
NO_BORDER   = Side(border_style=None)

def _border(left=False, right=False, top=False, bottom=False):
    return Border(
        left   = THIN_BLACK if left   else NO_BORDER,
        right  = THIN_BLACK if right  else NO_BORDER,
        top    = THIN_BLACK if top    else NO_BORDER,
        bottom = THIN_BLACK if bottom else NO_BORDER,
    )

HDR_BORDER = Border(
    left=THIN_BLACK, right=THIN_BLACK,
    top=THIN_BLACK,  bottom=THIN_BLACK,
)


def _apply_header(ws):
    ws.row_dimensions[1].height = 25
    headers = ['Folder Name', 'Word Document', 'Word Section', 'Table',
               'Row Label', 'Row Label Column', 'URL Column', 'URL']
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font      = HDR_FONT
        cell.fill      = HDR_FILL
        cell.alignment = HDR_ALIGN
        cell.border    = HDR_BORDER


def _write_group(ws, excel_row, group, folder, fname):
    """
    Write one file-group. Visual merge via borders:
      A & B cols:
        - ALL rows: right=thin (always visible boundary)
        - FIRST row: top=thin
        - LAST row:  bottom=thin  (sample shows last row also no bottom — 
                     but we add bottom on last to close the block cleanly)
        - MIDDLE rows: no top, no bottom  → seamless block appearance
      C D E H cols: full thin border all sides + white fill (Within Table rows)
      F G cols:     no border, no fill (matches sample exactly)
      Outside Table rows: D/E/F/G/H get no border (only C and H carry full border)
    """
    n = len(group)

    for i, rec in enumerate(group):
        r         = excel_row + i
        is_first  = (i == 0)
        is_last   = (i == n - 1)
        ws.row_dimensions[r].height = 18

        is_table  = rec['table_flag'] == 'Within Table'

        # ── Col A: Folder Name ─────────────────────────────────────────
        # Value written on every row so filter captures all rows in group.
        # Non-first rows use white font — value present for filter but invisible,
        # creating the seamless merged-block appearance without actual merges.
        ca = ws.cell(row=r, column=1, value=folder)
        ca.font      = Font(name='Calibri', size=11,
                            color='FF000000' if is_first else 'FFFFFFFF')
        ca.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
        ca.border    = _border(right=True, top=is_first, bottom=is_last)

        # ── Col B: Word Document ───────────────────────────────────────
        cb = ws.cell(row=r, column=2, value=fname)
        cb.font      = Font(name='Calibri', size=11,
                            color='FF000000' if is_first else 'FFFFFFFF')
        cb.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
        cb.border    = _border(right=True, top=is_first, bottom=is_last)

        # ── Col C: Word Section ────────────────────────────────────────
        cc = ws.cell(row=r, column=3, value=rec['section'])
        cc.font      = DAT_FONT
        cc.alignment = DAT_ALIGN
        cc.fill      = WHITE_FILL if is_table else NO_FILL
        cc.border    = _border(left=True, right=True, top=True, bottom=True)

        # ── Col D: Table flag ──────────────────────────────────────────
        cd = ws.cell(row=r, column=4, value=rec['table_flag'])
        cd.font      = DAT_FONT
        cd.alignment = DAT_ALIGN
        cd.fill      = WHITE_FILL if is_table else NO_FILL
        cd.border    = _border(left=True, right=True, top=True, bottom=True) if is_table else _border()

        # ── Col E: Row Label ───────────────────────────────────────────
        ce = ws.cell(row=r, column=5, value=rec['row_label'] or None)
        ce.font      = DAT_FONT
        ce.alignment = DAT_ALIGN
        ce.fill      = WHITE_FILL if is_table else NO_FILL
        ce.border    = _border(left=True, right=True, top=True, bottom=True) if is_table else _border()

        # ── Col F: Row Label Column (no border, no fill — matches sample) ─
        cf = ws.cell(row=r, column=6, value=rec['row_label_col'] or None)
        cf.font      = Font(name='Calibri', size=10, bold=True, color='FF000000')
        cf.alignment = DAT_ALIGN
        cf.fill      = NO_FILL
        cf.border    = _border()

        # ── Col G: URL Column (no border, no fill — matches sample) ───────
        cg = ws.cell(row=r, column=7, value=rec['url_col'] or None)
        cg.font      = Font(name='Calibri', size=10, bold=True, color='FF000000')
        cg.alignment = DAT_ALIGN
        cg.fill      = NO_FILL
        cg.border    = _border()

        # ── Col H: URL ────────────────────────────────────────────────
        ch = ws.cell(row=r, column=8, value=rec['url'])
        ch.font      = DAT_FONT
        ch.alignment = DAT_ALIGN
        ch.fill      = WHITE_FILL if is_table else NO_FILL
        ch.border    = _border(left=True, right=True, top=True, bottom=True) if is_table else _border()

    return excel_row + n


def write_excel(all_records, output_path):
    wb = Workbook()
    ws = wb.active
    ws.title        = 'HyperLink_Report'
    ws.freeze_panes = 'A2'

    # Enable auto-filter on header row so every column is filterable
    ws.auto_filter.ref = 'A1:H1'

    # Column widths from sample
    for col, w in zip('ABCDEFGH',
                      [12.54, 45.18, 32.27, 20.0, 21.09, 18.0, 18.0, 70.0]):
        ws.column_dimensions[col].width = w

    _apply_header(ws)

    # Group by (folder, filename) — keep original discovery order
    seen_keys, groups = [], {}
    for rec in all_records:
        key = (rec['folder'], rec['filename'])
        if key not in groups:
            groups[key] = []
            seen_keys.append(key)
        groups[key].append(rec)

    excel_row = 2
    for key in seen_keys:
        folder, fname = key
        excel_row = _write_group(ws, excel_row, groups[key], folder, fname)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    print(f'Saved -> {output_path}')


# ══════════════════════════════════════════════════════════════════════════
#  Main
# ══════════════════════════════════════════════════════════════════════════

def main():
    all_records = []

    files = (list(INPUT_FOLDER.rglob('*.docx')) +
             list(INPUT_FOLDER.rglob('*.doc')))

    if not files:
        print(f'No .docx/.doc files found under {INPUT_FOLDER}')
        return

    for fpath in files:
        try:
            rel    = fpath.relative_to(INPUT_FOLDER)
            folder = rel.parts[0] if len(rel.parts) > 1 else INPUT_FOLDER.name
        except ValueError:
            folder = fpath.parent.name

        fname = fpath.name
        ext   = fpath.suffix.lower()
        print(f'Processing: {fpath.relative_to(INPUT_FOLDER)}')

        try:
            if ext == '.docx':
                recs = extract_from_docx(fpath)
                for r in recs:
                    r['folder']   = folder
                    r['filename'] = fname
                all_records += recs
                t = sum(1 for r in recs if r['table_flag'] == 'Within Table')
                o = sum(1 for r in recs if r['table_flag'] == 'Outside Table')
                print(f'  -> {t} table links, {o} outside links')

            elif ext == '.doc':
                urls = extract_doc_urls(fpath)
                for url in urls:
                    all_records.append({
                        'folder':        folder,
                        'filename':      fname,
                        'section':       '',
                        'table_flag':    'Outside Table',
                        'row_label':     '',
                        'row_label_col': '',
                        'url_col':       '',
                        'url':           url,
                    })
                print(f'  -> {len(urls)} links [.doc binary]')

        except Exception as exc:
            print(f'  WARNING: {fpath.name}: {exc}')

    print(f'\nTotal records: {len(all_records)}')
    write_excel(all_records, OUTPUT_PATH)


if __name__ == '__main__':
    main()