"""
Extract "Process Risk" tables from all .doc / .docx files in INPUT_FOLDER
(including sub-folders) and write a formatted Excel file to OUTPUT_FOLDER.

Layout
------
A  Folder Name    – fake-merge (value on first row of each file block)
B  Word Doc Name  – same fake-merge trick
C+ Content cols   – common headers derived from all files that have a table;
                    "Table Not Found" (amber) in every column when no table exists

Notes
-----
- Supports both .docx and legacy .doc files.
  Legacy .doc files are converted to .docx via LibreOffice (soffice) before parsing.
- Merged title rows (e.g. "Risk and Issues Log" spanning all columns) are
  automatically detected and skipped; row 1 of the table is used as the header.
- Column headers are unified across all files into one common header row.
- No "Filter Category" column.
- Files with no matching Process Risk section/table are skipped entirely.
- Section detection uses DUAL strategy: checks paragraph heading style AND
  paragraph text. This handles documents where heading styles are missing,
  renamed, or differ between LibreOffice versions.

Usage
-----
1. Place all .doc / .docx files in   <script_dir>/Input/  (sub-folders OK).
2. Run:  python extract_process_risk.py
3. Output written to               <script_dir>/Output/process_risk_output.xlsx
"""

import os
import re
import shutil
import subprocess
import tempfile
from dataclasses import dataclass, field
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ── paths ─────────────────────────────────────────────────────────────────────
_HERE        = os.path.dirname(os.path.abspath(__file__))
INPUT_FOLDER = os.path.join(_HERE, "Input")
OUTPUT_FILE  = os.path.join(_HERE, "Output", "process_risk_output.xlsx")

# ── constants ─────────────────────────────────────────────────────────────────
RISK_KEYWORDS = ("process risk",)
HEADING_TAGS  = ("heading", "title", "toc")

W_T      = qn("w:t")
W_PPR    = qn("w:pPr")
W_PSTYLE = qn("w:pStyle")

# ── styles ────────────────────────────────────────────────────────────────────
def _F(h):          return PatternFill("solid", fgColor=h)
def _Ft(b, c):      return Font(name="Calibri", bold=b, color=c, size=11)
def _A(h, v):       return Alignment(horizontal=h, vertical=v, wrap_text=True)
def _S(s, c):       return Side(style=s, color=c)
def _B(t, b, l, r): return Border(top=t, bottom=b, left=l, right=r)

_NO, _MED, _THIN = Side(style=None), _S("medium", "4472C4"), _S("thin", "BFBFBF")
_TALL = _B(_THIN, _THIN, _THIN, _THIN)

F_HDR = _F("1F4E79")   # dark blue  – global header row
F_FN  = _F("F2F2F2")   # light grey – folder/filename cells
F_VAL = _F("FFFFFF")   # white      – data cells

FT_HDR            = _Ft(True,  "FFFFFF")
FT_FN_V, FT_FN_H = _Ft(True,  "1F4E79"), _Ft(True, "F2F2F2")
FT_VAL            = _Ft(False, "000000")

AL_CC = _A("center", "center")
AL_LC = _A("left",   "center")
AL_LT = _A("left",   "top")

_FN_BDR = {
    (f, l): _B(_MED if f else _NO, _MED if l else _NO, _MED, _THIN)
    for f in (True, False) for l in (True, False)
}

# ── cell writers ──────────────────────────────────────────────────────────────
def _put(c, fill, font, align, border, value=None):
    c.fill = fill; c.font = font; c.alignment = align; c.border = border
    if value is not None:
        c.value = value

def _hdr(c, v):             _put(c, F_HDR, FT_HDR, AL_CC, _TALL, v)
def _val(c, v=""):          _put(c, F_VAL, FT_VAL, AL_LT, _TALL, v)
def _fn(c, v, first, last): _put(c, F_FN,
                                   FT_FN_V if first else FT_FN_H,
                                   AL_LC, _FN_BDR[(first, last)], v)

# ── data model ────────────────────────────────────────────────────────────────
@dataclass
class Block:
    folder:     str
    filename:   str
    has_table:  bool = False
    base_hdrs:  list = field(default_factory=list)
    rows:       list = field(default_factory=list)
    col_list:   list = field(default_factory=list)
    col_widths: list = field(default_factory=list)

# ── docx helpers ──────────────────────────────────────────────────────────────
def _paras(cell) -> list:
    return [p.text.strip() for p in cell.paragraphs if p.text.strip()]

def _text(cell) -> str:
    return "\n".join(_paras(cell))

def _is_merged_title_row(row) -> bool:
    """Return True when every cell contains the same non-empty text (merged header)."""
    texts  = [_text(c).strip() for c in row.cells]
    unique = {t for t in texts if t}
    return len(unique) == 1

def _is_risk_section(body: list, tbl_elem) -> bool:
    """
    Scan upward through all preceding paragraphs.

    Stops at the first paragraph that is EITHER:
      (a) styled as a heading  (style name contains "heading", "title", or "toc"), OR
      (b) a non-empty paragraph whose text contains a risk keyword.

    Returns True only if that stopping paragraph contains a risk keyword.

    This dual strategy handles:
      - Documents with proper Heading styles
      - Documents where LibreOffice converts heading styles differently
        (e.g. "Heading_20_1" vs "Heading1")
      - Documents that use bold Normal paragraphs as visual headings
    """
    try:
        idx = body.index(tbl_elem)
    except ValueError:
        return False

    for elem in reversed(body[:idx]):
        if elem.tag.split("}")[-1] != "p":
            continue

        pPr   = elem.find(W_PPR)
        style = ""
        if pPr is not None:
            ps = pPr.find(W_PSTYLE)
            if ps is not None:
                style = (ps.get(qn("w:val")) or "").lower().replace("-", " ")

        text       = "".join(t.text or "" for t in elem.iter() if t.tag == W_T).lower().strip()
        is_heading = any(h in style for h in HEADING_TAGS)
        has_risk   = any(k in text for k in RISK_KEYWORDS)

        # Stop when we hit a heading-styled paragraph OR any paragraph with risk text
        if is_heading or (text and has_risk):
            return has_risk

        # Plain non-heading paragraph with no risk text → keep scanning upward

    return False

def _extract(table) -> tuple:
    """
    Parse table, skipping any merged title row.
    Returns (base_hdrs, logical_rows).
    """
    if not table.rows:
        return [], []

    rows  = table.rows
    start = 1 if (len(rows) > 1 and _is_merged_title_row(rows[0])) else 0

    if start >= len(rows):
        return [], []

    base_hdrs = [
        _text(c).strip() or f"Col{i+1}"
        for i, c in enumerate(rows[start].cells)
    ]

    ref_hdr = base_hdrs[0] if base_hdrs else None   # "Reference" is always col 0

    logical = []
    for row in rows[start + 1:]:
        rd = {
            h: _text(row.cells[i]).strip() if i < len(row.cells) else ""
            for i, h in enumerate(base_hdrs)
        }
        # Skip rows where the Reference column is blank
        if ref_hdr and not rd.get(ref_hdr, "").strip():
            continue
        logical.append(rd)

    return base_hdrs, logical

# ── .doc → .docx conversion ───────────────────────────────────────────────────
def _soffice_exe() -> str:
    """
    Return the soffice executable name that works on this machine.
    Tries 'soffice' (Linux/macOS) then 'soffice.exe' (Windows), then
    common Windows installation paths.
    """
    import shutil as _shutil, sys as _sys
    for candidate in ("soffice", "soffice.exe"):
        if _shutil.which(candidate):
            return candidate
    # Common Windows LibreOffice install paths
    if _sys.platform == "win32":
        import winreg
        for base in (
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ):
            if Path(base).exists():
                return base
    return "soffice"   # fallback — will fail with a clear error message


def _convert_doc_to_docx(src: Path):
    """
    Convert a legacy .doc to .docx using LibreOffice (headless).
    Returns (Path_to_converted_docx, tmp_dir) or (None, None) on failure.
    Caller must shutil.rmtree(tmp_dir) when done.
    """
    tmp_dir = tempfile.mkdtemp(prefix="doc_convert_")
    exe = _soffice_exe()
    try:
        result = subprocess.run(
            [exe, "--headless", "--convert-to", "docx",
             "--outdir", tmp_dir, str(src)],
            capture_output=True, text=True, timeout=120,
        )
        if result.returncode != 0:
            print(f"  [WARN] soffice failed (exe={exe}): {result.stderr.strip()}")
            shutil.rmtree(tmp_dir, ignore_errors=True)
            return None, None
        converted = next(Path(tmp_dir).glob("*.docx"), None)
        if converted is None:
            print(f"  [WARN] soffice ran but produced no .docx (exe={exe})")
            shutil.rmtree(tmp_dir, ignore_errors=True)
            return None, None
        return converted, tmp_dir
    except FileNotFoundError:
        print(f"  [ERROR] LibreOffice not found (tried: {exe}). "
              f"Install LibreOffice and ensure 'soffice' is on your PATH.")
        shutil.rmtree(tmp_dir, ignore_errors=True)
        return None, None
    except Exception as exc:
        print(f"  [WARN] conversion error: {exc}")
        shutil.rmtree(tmp_dir, ignore_errors=True)
        return None, None


def _open_doc(fpath: Path):
    """Return (Document, tmp_dir_or_None). Caller cleans up tmp_dir."""
    if fpath.suffix.lower() == ".docx":
        return Document(str(fpath)), None
    print("  → Converting .doc → .docx via LibreOffice …")
    converted, tmp_dir = _convert_doc_to_docx(fpath)
    if converted is None:
        return None, None
    return Document(str(converted)), tmp_dir


# ── block collection ──────────────────────────────────────────────────────────
def collect() -> tuple:
    """Return (blocks, skipped) where skipped is a list of (folder, filename, reason)."""
    blocks  = []
    skipped = []   # (folder, filename, reason)

    # Collect all .doc and .docx files via suffix check (not glob pattern)
    # so behaviour is identical on Windows (case-insensitive FS) and Linux.
    # On Windows, rglob("*.doc") incorrectly matches .docx files too.
    # If both file.doc and file.docx exist, prefer .docx (already converted).
    all_files = [
        p for p in sorted(Path(INPUT_FOLDER).rglob("*"))
        if p.suffix.lower() in (".doc", ".docx")
        and not p.name.startswith("~$")
        and p.is_file()
    ]
    stem_map: dict[str, Path] = {}
    for p in all_files:
        key = (str(p.parent), p.stem.lower())
        if key not in stem_map or p.suffix.lower() == ".docx":
            stem_map[key] = p
    found = sorted(stem_map.values(), key=lambda p: (str(p.parent), p.name))

    for fpath in found:
        rel    = fpath.relative_to(INPUT_FOLDER)
        folder = str(rel.parent) if str(rel.parent) != "." else "(root)"
        print(f"Processing: {rel}")

        blk = Block(folder=folder, filename=fpath.name)
        doc, tmp_dir = None, None
        found_section = False   # tracks whether "Process Risk" heading was seen

        try:
            doc, tmp_dir = _open_doc(fpath)
        except Exception as exc:
            print(f"  [ERROR] open failed: {exc}")

        if doc is not None:
            try:
                body = list(doc.element.body)
                for ti, table in enumerate(doc.tables):
                    if not _is_risk_section(body, table._tbl):
                        continue
                    found_section = True   # heading exists
                    b_hdrs, l_rows = _extract(table)
                    if not b_hdrs:
                        continue
                    blk.has_table  = True
                    blk.base_hdrs  = b_hdrs
                    blk.rows       = l_rows
                    blk.col_list   = b_hdrs
                    blk.col_widths = [min(max(len(l) * 1.2, 16), 55) for l in b_hdrs]
                    print(f"  → Table #{ti+1}: {len(l_rows)} row(s), headers: {b_hdrs}")
                    break

                # If no table matched, check whether the section heading exists
                # by scanning HEADING-STYLED paragraphs only (not body text)
                # to avoid false positives from body text mentioning "process risk"
                if not blk.has_table and not found_section:
                    for elem in body:
                        if elem.tag.split("}")[-1] != "p":
                            continue
                        pPr = elem.find(W_PPR); style = ""
                        if pPr is not None:
                            ps = pPr.find(W_PSTYLE)
                            if ps is not None:
                                style = (ps.get(qn("w:val")) or "").lower()
                        is_heading = any(h in style for h in HEADING_TAGS)
                        if not is_heading:
                            continue   # only check actual heading paragraphs
                        text = "".join(
                            t.text or "" for t in elem.iter()
                            if t.tag == W_T
                        ).lower().strip()
                        if any(k in text for k in RISK_KEYWORDS):
                            found_section = True
                            break

            except Exception as exc:
                print(f"  [ERROR] parse failed: {exc}")
            finally:
                if tmp_dir:
                    shutil.rmtree(tmp_dir, ignore_errors=True)

        if not blk.has_table:
            if not found_section:
                reason = "'Process Risk' section not found"
            else:
                reason = "'Process Risk' section found but table is missing"
            print(f"  [INFO] {reason} — adding to skipped list.")
            skipped.append((folder, fpath.name, reason))
            continue

        blocks.append(blk)

    return blocks, skipped

# ── excel writer ──────────────────────────────────────────────────────────────
CS = 3   # content starts at column C (1=A, 2=B, 3=C …)

def write(blocks: list, path: str, skipped: list | None = None) -> None:
    skipped = skipped or []
    if not blocks and not skipped:
        print("[WARN] Nothing to write.")
        return

    all_hdr_sets = [tuple(b.col_list) for b in blocks if b.has_table]
    if all_hdr_sets:
        seen: dict = {}
        for hset in all_hdr_sets:
            for h in hset:
                if h not in seen:
                    seen[h] = True
        common_hdrs = list(seen.keys())
    else:
        common_hdrs = ["Reference", "Risks – What Could Go Wrong?",
                       "Implication", "Control No. per RACF"]

    total_cols = CS - 1 + len(common_hdrs)

    wb = Workbook()
    ws = wb.active
    ws.title = "Process Risk"

    # Row 1 — global header
    _hdr(ws.cell(1, 1), "Folder Name")
    _hdr(ws.cell(1, 2), "Word Doc Name")
    for i, lbl in enumerate(common_hdrs):
        _hdr(ws.cell(1, CS + i), lbl)

    # Fill colour used for A/B cells — reused for invisible interior borders
    _FN_HEX = "F2F2F2"   # must match F_FN fgColor

    cur = 2

    for blk in blocks:
        n_rows = len(blk.rows)
        for ri, lr in enumerate(blk.rows):
            first  = ri == 0
            last   = ri == n_rows - 1
            middle = not first and not last   # interior rows of a 3+ row block

            # ── Cols A & B: visual-merge via cell formatting ─────────────────
            # Real value written on EVERY row  → filter returns all rows.
            # Visual illusion of a merged cell block:
            #   • First row  : bold dark text, thick top + left + right border
            #   • Interior   : text colour = fill colour (invisible), no top/
            #                  bottom border (colour = fill), left/right borders
            #   • Last row   : same as interior + thick bottom border
            #   • Single row : full thick border (first AND last)
            for col, val in [(1, blk.folder), (2, blk.filename)]:
                c = ws.cell(cur, col)
                c.value     = val
                c.fill      = F_FN

                # Text visible only on first row; rest blend into fill
                c.font = Font(
                    name="Calibri", bold=True, size=11,
                    color="1F4E79" if first else _FN_HEX,
                )
                # Vertical alignment: top on first row (text anchors to top of
                # the visual block), center on others (irrelevant — invisible)
                c.alignment = Alignment(
                    horizontal="left",
                    vertical="top" if first else "center",
                    wrap_text=True,
                )
                # Borders:
                #   outer edges of the block  → medium blue
                #   internal horizontal lines → none (colour = fill = invisible)
                #   right edge                → thin grey (separates from content)
                c.border = Border(
                    top    = Side(
                        style="medium" if first else "thin",
                        color="4472C4"  if first else _FN_HEX,
                    ),
                    bottom = Side(
                        style="medium" if last  else "thin",
                        color="4472C4"  if last  else _FN_HEX,
                    ),
                    left   = Side(style="medium", color="4472C4"),
                    right  = Side(style="thin",   color="BFBFBF"),
                )

            for i, h in enumerate(common_hdrs):
                _val(ws.cell(cur, CS + i), lr.get(h, ""))
            cur += 1

    last_data_row = cur - 1

    ws.auto_filter.ref = f"A1:{get_column_letter(total_cols)}{last_data_row}"
    ws.freeze_panes    = "A2"

    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 48
    widths: dict = {}
    for blk in blocks:
        for i, w in enumerate(blk.col_widths):
            widths[i] = max(widths.get(i, 0), w)
    for i, w in widths.items():
        ws.column_dimensions[get_column_letter(CS + i)].width = w

    ws.row_dimensions[1].height = 28
    for r in range(2, last_data_row + 1):
        ws.row_dimensions[r].height = 40

    # ── "Not Found" sheet ────────────────────────────────────────────────────
    if skipped:
        ws_nf = wb.create_sheet("Not Found")

        # Column widths
        ws_nf.column_dimensions["A"].width = 22
        ws_nf.column_dimensions["B"].width = 50
        ws_nf.column_dimensions["C"].width = 45

        # Header row
        nf_hdr_fill = PatternFill("solid", fgColor="C00000")   # dark red
        nf_hdr_font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
        nf_val_fill = PatternFill("solid", fgColor="FFFFFF")
        nf_val_font = Font(name="Calibri", bold=False, color="000000", size=11)
        nf_wrn_fill = PatternFill("solid", fgColor="FFF2CC")   # amber
        nf_wrn_font = Font(name="Calibri", bold=False, color="7F6000", size=11)
        nf_align_hdr = Alignment(horizontal="center", vertical="center",
                                  wrap_text=True)
        nf_align_val = Alignment(horizontal="left",   vertical="center",
                                  wrap_text=True)
        thin  = Side(style="thin",   color="BFBFBF")
        bdr   = Border(top=thin, bottom=thin, left=thin, right=thin)

        for ci, lbl in enumerate(["Folder Name", "Word Doc Name", "Reason"], 1):
            c = ws_nf.cell(1, ci)
            c.value = lbl; c.fill = nf_hdr_fill; c.font = nf_hdr_font
            c.alignment = nf_align_hdr; c.border = bdr
        ws_nf.row_dimensions[1].height = 28

        for ri, (folder, filename, reason) in enumerate(skipped, 2):
            # Choose amber fill for "table missing" (section exists but no table)
            # and plain white for "section not found"
            is_missing_table = "table is missing" in reason
            row_fill = nf_wrn_fill if is_missing_table else nf_val_fill
            row_font_reason = nf_wrn_font if is_missing_table else nf_val_font

            for ci, val in enumerate([folder, filename, reason], 1):
                c = ws_nf.cell(ri, ci)
                c.value = val; c.border = bdr; c.alignment = nf_align_val
                c.fill = row_fill
                c.font = row_font_reason if ci == 3 else nf_val_font

            ws_nf.row_dimensions[ri].height = 36

        ws_nf.auto_filter.ref = f"A1:C{1 + len(skipped)}"
        ws_nf.freeze_panes = "A2"
        print(f"   Not Found sheet: {len(skipped)} file(s) listed")

    # ── Summary sheet ────────────────────────────────────────────────────────
    ws_sum = wb.create_sheet("Summary")

    # Distinct file counts
    distinct_sheet1 = len({blk.filename for blk in blocks})
    distinct_sheet2 = len({row[1] for row in skipped}) if skipped else 0

    # Styles
    s_hdr_fill = PatternFill("solid", fgColor="1F4E79")
    s_hdr_font = Font(name="Calibri", bold=True,  color="FFFFFF", size=11)
    s_lbl_fill = PatternFill("solid", fgColor="2E75B6")
    s_lbl_font = Font(name="Calibri", bold=True,  color="FFFFFF", size=11)
    s_val_fill = PatternFill("solid", fgColor="FFFFFF")
    s_val_font = Font(name="Calibri", bold=True,  color="1F4E79", size=14)
    s_thin     = Side(style="thin",   color="BFBFBF")
    s_med      = Side(style="medium", color="4472C4")
    s_bdr_hdr  = Border(top=s_med,  bottom=s_med,  left=s_med,  right=s_med)
    s_bdr_lbl  = Border(top=s_thin, bottom=s_thin, left=s_med,  right=s_thin)
    s_bdr_val  = Border(top=s_thin, bottom=s_thin, left=s_thin, right=s_med)
    s_al_cc    = Alignment(horizontal="center", vertical="center")
    s_al_lc    = Alignment(horizontal="left",   vertical="center")

    def _s_hdr(c, v):
        c.value=v; c.fill=s_hdr_fill; c.font=s_hdr_font
        c.alignment=s_al_cc; c.border=s_bdr_hdr
    def _s_lbl(c, v):
        c.value=v; c.fill=s_lbl_fill; c.font=s_lbl_font
        c.alignment=s_al_lc; c.border=s_bdr_lbl
    def _s_val(c, v):
        c.value=v; c.fill=s_val_fill; c.font=s_val_font
        c.alignment=s_al_cc; c.border=s_bdr_val

    # Header row
    _s_hdr(ws_sum.cell(1, 1), "Sheet")
    _s_hdr(ws_sum.cell(1, 2), "Description")
    _s_hdr(ws_sum.cell(1, 3), "Distinct File Count")

    # Row 2 — Sheet 1
    _s_lbl(ws_sum.cell(2, 1), "Process Risk")
    _s_lbl(ws_sum.cell(2, 2), "Files with Process Risk table extracted")
    _s_val(ws_sum.cell(2, 3), distinct_sheet1)

    # Row 3 — Sheet 2
    _s_lbl(ws_sum.cell(3, 1), "Not Found")
    _s_lbl(ws_sum.cell(3, 2), "Files where section or table was not found")
    _s_val(ws_sum.cell(3, 3), distinct_sheet2)

    # Row 4 — Total
    tot_fill = PatternFill("solid", fgColor="D6E4F0")
    tot_font_lbl = Font(name="Calibri", bold=True, color="1F4E79", size=11)
    tot_font_val = Font(name="Calibri", bold=True, color="1F4E79", size=14)
    s_bdr_tot = Border(top=s_med, bottom=s_med, left=s_med, right=s_med)
    for ci, val in enumerate(["Total", "All files processed", distinct_sheet1 + distinct_sheet2], 1):
        c = ws_sum.cell(4, ci)
        c.value = val; c.fill = tot_fill; c.border = s_bdr_tot
        c.alignment = s_al_lc if ci < 3 else s_al_cc
        c.font = tot_font_val if ci == 3 else tot_font_lbl

    ws_sum.column_dimensions["A"].width = 18
    ws_sum.column_dimensions["B"].width = 45
    ws_sum.column_dimensions["C"].width = 22
    for r in range(1, 5):
        ws_sum.row_dimensions[r].height = 32

    os.makedirs(os.path.dirname(path), exist_ok=True)
    wb.save(path)
    print(f"\n✅  {cur - 1} rows | {total_cols} cols → {path}")


# ── entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    blocks, skipped = collect()
    write(blocks, OUTPUT_FILE, skipped)