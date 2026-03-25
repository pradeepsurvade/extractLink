"""
Extract "Process Performance" tables from all .docx in INPUT_FOLDER
and write a formatted Excel file to OUTPUT_FOLDER.

Layout
------
A  Folder Name       – fake-merge (value every row, visible only on first)
B  Word Doc Name     – same fake-merge trick
C  Filter Category   – Header Content 1 value repeated on every row;
                       "Table Not Found" when no table exists
D+ Header Content N  – file-specific columns, N/A cells shaded for other files
"""

import os
import re
from dataclasses import dataclass, field
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ── paths ────────────────────────────────────────────────────────────────────
_HERE        = os.path.dirname(os.path.abspath(__file__))
INPUT_FOLDER = os.path.join(_HERE, "Input")
OUTPUT_FILE  = os.path.join(_HERE, "Output", "process_performance_output.xlsx")

PERF_KEYWORDS = ("process performance", "sla", "kpi")
HEADING_TAGS  = ("heading", "title", "toc")
W_T           = qn("w:t")
W_PPR         = qn("w:pPr")
W_PSTYLE      = qn("w:pStyle")

# ── label detection ──────────────────────────────────────────────────────────
# Matches "Up To Six Words: value" — colon must follow a short, dot-free prefix
_LABEL_RE = re.compile(r"^\*{0,2}((?:\S+ ?){1,6}?)\*{0,2}\s*:\s+\S", re.ASCII)

def _has_label(text: str) -> bool:
    m = _LABEL_RE.match(text.lstrip())
    return bool(m) and "." not in m.group(1)

# ── styles (built once at import time) ───────────────────────────────────────
def _F(h):       return PatternFill("solid", fgColor=h)
def _Ft(b, c):   return Font(name="Calibri", bold=b, color=c, size=11)
def _A(h, v):    return Alignment(horizontal=h, vertical=v, wrap_text=True)
def _S(s, c):    return Side(style=s, color=c)
def _B(t,b,l,r): return Border(top=t, bottom=b, left=l, right=r)

_NO, _MED, _THIN = Side(style=None), _S("medium","4472C4"), _S("thin","BFBFBF")
_TALL = _B(_THIN,_THIN,_THIN,_THIN)

F_HDR,F_FN,F_SH,F_VAL,F_NA    = _F("1F4E79"),_F("F2F2F2"),_F("2E75B6"),_F("FFFFFF"),_F("F5F5F5")
FT_HDR                          = _Ft(True,  "FFFFFF")
FT_FN_V, FT_FN_H               = _Ft(True,  "1F4E79"), _Ft(True, "F2F2F2")
FT_SH, FT_VAL, FT_NA           = _Ft(True,  "FFFFFF"), _Ft(False,"000000"), _Ft(False,"D0D0D0")
AL_CC, AL_LC, AL_LT             = _A("center","center"), _A("left","center"), _A("left","top")

# All 4 border variants for fn-cells — keyed by (is_first, is_last)
_FN_BDR = {(f,l): _B(_MED if f else _NO, _MED if l else _NO, _MED, _THIN)
           for f in (True,False) for l in (True,False)}

# ── cell writers ─────────────────────────────────────────────────────────────
def _put(c, fill, font, align, border, value=None):
    c.fill=fill; c.font=font; c.alignment=align; c.border=border
    if value is not None: c.value = value

def _hdr(c, v):              _put(c, F_HDR, FT_HDR, AL_CC, _TALL, v)
def _sh(c, v):               _put(c, F_SH,  FT_SH,  AL_LC, _TALL, v)
def _val(c, v=""):           _put(c, F_VAL, FT_VAL, AL_LT, _TALL, v)
def _na(c):                  _put(c, F_NA,  FT_NA,  AL_CC, _TALL, "")
def _fn(c, v, first, last):  _put(c, F_FN,  FT_FN_V if first else FT_FN_H,
                                   AL_LC, _FN_BDR[(first, last)], v)


# ── data model ────────────────────────────────────────────────────────────────
@dataclass
class Block:
    folder:     str
    filename:   str
    has_table:  bool            = False
    base_hdrs:  list            = field(default_factory=list)
    rows:       list            = field(default_factory=list)  # dicts with "_lp" key
    col_list:   list            = field(default_factory=list)  # final column labels
    col_widths: list            = field(default_factory=list)  # estimated px widths


# ── docx extraction ───────────────────────────────────────────────────────────
def _paras(cell) -> list:
    return [p.text.strip() for p in cell.paragraphs if p.text.strip()]

def _text(cell) -> str:
    return "\n".join(_paras(cell))

def _merge_continuations(raw: list) -> list:
    """Fold unlabelled paragraphs into the preceding labelled slot."""
    out = []
    for p in raw:
        if _has_label(p): out.append(p)
        elif out:         out[-1] += "\n" + p
        else:             out.append(p)
    return out

def _is_perf_section(body: list, tbl_elem) -> bool:
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
        if any(h in style for h in HEADING_TAGS):
            text = "".join(t.text or "" for t in elem.iter() if t.tag == W_T).lower()
            return any(k in text for k in PERF_KEYWORDS)
    return False

def _extract(table) -> tuple:
    """Return (base_hdrs, last_col, logical_rows) — deduped + continuations merged."""
    if not table.rows:
        return [], None, []

    cells_r0 = table.rows[0].cells
    doc_hdrs  = [_text(c) or f"Col{i+1}" for i, c in enumerate(cells_r0)]
    base_hdrs, last_col = doc_hdrs[:-1], doc_hdrs[-1]

    groups: dict = {}
    for row in table.rows[1:]:
        key = tuple(_text(c) for c in row.cells[:-1])
        if key not in groups:
            groups[key] = {"base_vals": list(key), "seen": [], "lp": []}
        seen = groups[key]["seen"]
        for p in _paras(row.cells[-1]):
            if p not in seen:
                seen.append(p)

    logical = []
    for g in groups.values():
        rd = dict(zip(base_hdrs, g["base_vals"]))
        rd["_lp"] = _merge_continuations(g["seen"])
        logical.append(rd)

    return base_hdrs, last_col, logical


# ── block collection ──────────────────────────────────────────────────────────
def collect() -> list:
    blocks = []
    for fpath in sorted(Path(INPUT_FOLDER).rglob("*.docx")):
        rel    = fpath.relative_to(INPUT_FOLDER)
        folder = str(rel.parent) if str(rel.parent) != "." else "(root)"
        print(f"Processing: {rel}")

        blk = Block(folder=folder, filename=fpath.name)
        try:
            doc = Document(str(fpath))
        except Exception as e:
            print(f"  [ERROR] {e}"); blocks.append(blk); continue

        body = list(doc.element.body)   # built once per doc, reused across tables

        for ti, table in enumerate(doc.tables):
            if not _is_perf_section(body, table._tbl):
                continue
            b_hdrs, last_col, l_rows = _extract(table)
            if not b_hdrs:
                continue

            mlp = max((len(r["_lp"]) for r in l_rows), default=0)
            lc_labels = (
                [f"{last_col} {i+1}" for i in range(mlp)] if last_col and mlp > 1
                else [last_col] if last_col
                else []
            )
            col_list   = b_hdrs + lc_labels
            col_widths = [min(max(len(l) * 1.05, 14), 50) for l in col_list]

            blk.has_table  = True
            blk.base_hdrs  = b_hdrs
            blk.rows       = l_rows
            blk.col_list   = col_list
            blk.col_widths = col_widths
            print(f"  → Table #{ti+1}: {len(l_rows)} row(s), {len(col_list)} col(s)")
            break   # stop after first matching table per file

        if not blk.has_table:
            print("  [INFO] No Process Performance table — blank row.")
        blocks.append(blk)

    return blocks


# ── excel writer ──────────────────────────────────────────────────────────────
CS = 4   # Content Start: 1=A 2=B 3=C 4=D

def write(blocks: list, path: str) -> None:
    if not blocks:
        print("[WARN] Nothing to write."); return

    max_cc     = max((len(b.col_list) for b in blocks), default=1)
    total_cols = CS - 1 + max_cc   # A + B + C + content

    wb = Workbook()
    ws = wb.active
    ws.title = "Process Performance"

    # Row 1 — global header
    for ci, lbl in enumerate(
        ["Folder Name", "Word Doc Name", "Filter Category"] +
        [f"Header Content {i+1}" for i in range(max_cc)], 1
    ):
        _hdr(ws.cell(row=1, column=ci), lbl)

    cur = 2
    for blk in blocks:
        filter_cat = blk.col_list[0] if blk.col_list else "Table Not Found"
        n_own      = len(blk.col_list)

        if blk.has_table and blk.rows:
            n_block = 1 + len(blk.rows)
            mlp     = max((len(r["_lp"]) for r in blk.rows), default=0)
            lp_off  = CS + len(blk.base_hdrs)

            # sub-header row
            _fn(ws.cell(cur,1), blk.folder,   True, n_block==1)
            _fn(ws.cell(cur,2), blk.filename, True, n_block==1)
            _sh(ws.cell(cur,3), filter_cat)
            for i, lbl in enumerate(blk.col_list):
                _sh(ws.cell(cur, CS+i), lbl)
            for ci in range(CS+n_own, total_cols+1):
                _na(ws.cell(cur, ci))
            cur += 1

            # data rows
            for ri, lr in enumerate(blk.rows):
                is_last = ri == len(blk.rows) - 1
                _fn(ws.cell(cur,1), blk.folder,   False, is_last)
                _fn(ws.cell(cur,2), blk.filename, False, is_last)
                _sh(ws.cell(cur,3), filter_cat)
                for i, h in enumerate(blk.base_hdrs):
                    _val(ws.cell(cur, CS+i), lr.get(h, ""))
                lp = lr.get("_lp", [])
                for pi in range(mlp):
                    _val(ws.cell(cur, lp_off+pi), lp[pi] if pi < len(lp) else "")
                for ci in range(CS+n_own, total_cols+1):
                    _na(ws.cell(cur, ci))
                cur += 1

        else:
            # no-table: single blank row
            _fn(ws.cell(cur,1), blk.folder,   True, True)
            _fn(ws.cell(cur,2), blk.filename, True, True)
            _sh(ws.cell(cur,3), "Table Not Found")
            for ci in range(CS, total_cols+1):
                _val(ws.cell(cur, ci))
            cur += 1

    ws.auto_filter.ref = f"A1:{get_column_letter(total_cols)}{cur-1}"
    ws.freeze_panes    = "A2"

    # Column widths — A/B/C fixed; content cols = max width across all blocks
    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 46
    ws.column_dimensions["C"].width = 28
    widths: dict = {}
    for blk in blocks:
        for i, w in enumerate(blk.col_widths):
            widths[i] = max(widths.get(i, 0), w)
    for i, w in widths.items():
        ws.column_dimensions[get_column_letter(CS+i)].width = w

    # Row heights
    ws.row_dimensions[1].height = 28
    for r in range(2, cur):
        ws.row_dimensions[r].height = 36

    os.makedirs(os.path.dirname(path), exist_ok=True)
    wb.save(path)
    print(f"\n✅  {cur-1} rows | {total_cols} cols → {path}")


# ── entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    write(collect(), OUTPUT_FILE)
