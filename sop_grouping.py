"""
SOP Document Grouping
======================
Reads Word documents from input/, extracts process steps from
"SOP Activities Overview" and "Detailed Process Steps" sections,
and groups documents with similar process steps.

Output: output/SOP_Groups.xlsx
  Column A: Group Name
  Column B: Document Name
  Column C: Similar Steps Found (with match %)
  Column D: SAP Transaction Codes

Setup (one-time, for best accuracy):
  pip install sentence-transformers
  python -c "from sentence_transformers import SentenceTransformer; SentenceTransformer('all-MiniLM-L6-v2')"

Usage:
  python sop_grouping.py
  python sop_grouping.py --groups 5
"""

import os, re, argparse
import numpy as np
from collections import defaultdict

try:
    from docx import Document
except ImportError:
    raise SystemExit("Run: pip install python-docx")

try:
    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.cluster import AgglomerativeClustering
    from sklearn.metrics.pairwise import cosine_similarity
except ImportError:
    raise SystemExit("Run: pip install scikit-learn numpy")

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
except ImportError:
    raise SystemExit("Run: pip install openpyxl")


# =============================================================================
# CONFIGURATION
# =============================================================================

# Section headings to look for (case-insensitive partial match).
# Based on the SOP template structure.
ACTIVITIES_OVERVIEW_HEADINGS = [
    "SOP Activities Overview",
    "Activities Overview",
    "Process Overview",
]

DETAILED_STEPS_HEADINGS = [
    "Detailed Process Steps",
    "Process Detail",
    "Process Steps",
    "Procedure",
]

# Minimum words needed in extracted steps to treat as a full SOP.
MIN_WORDS = 20

# Embedding model — runs fully offline after one-time download.
# Set to None to skip and always use TF-IDF.
EMBEDDING_MODEL = "all-MiniLM-L6-v2"

# How many matching steps to show per document in the Excel report.
TOP_STEPS = 5

# =============================================================================


# ---------------------------------------------------------------------------
# Extract steps from the two key sections
# ---------------------------------------------------------------------------

def is_heading(para) -> bool:
    style = (para.style.name or "").lower()
    if style.startswith("heading") or style.startswith("title"):
        return True
    runs = [r for r in para.runs if r.text.strip()]
    if runs and all(r.bold for r in runs):
        return True
    t = para.text.strip()
    if len(t.split()) <= 8 and t == t.upper() and len(t) > 3:
        return True
    return False


def text_matches(text: str, headings: list) -> bool:
    t = text.strip().upper()
    for h in headings:
        if h.upper() in t or t in h.upper():
            return True
    return False


def extract_table_after_heading(doc, target_headings: list) -> list:
    """
    Find the heading matching target_headings, then return all text from
    the NEXT table after that heading. Returns list of cell text strings.
    """
    paras = list(doc.paragraphs)
    tables = list(doc.tables)

    # Build a flat ordered list of (type, object) for paragraphs and tables
    # by iterating the document XML body children
    from docx.oxml.ns import qn
    body = doc.element.body
    items = []
    para_idx = 0
    tbl_idx  = 0

    for child in body:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag == 'p':
            if para_idx < len(paras):
                items.append(('para', paras[para_idx]))
                para_idx += 1
        elif tag == 'tbl':
            if tbl_idx < len(tables):
                items.append(('table', tables[tbl_idx]))
                tbl_idx += 1

    # Find heading then grab the next table
    found_heading = False
    for i, (typ, obj) in enumerate(items):
        if typ == 'para' and is_heading(obj) and text_matches(obj.text, target_headings):
            found_heading = True
            # Scan forward for next table
            for j in range(i+1, len(items)):
                t2, obj2 = items[j]
                if t2 == 'para' and is_heading(obj2) and obj2.text.strip():
                    break   # hit another heading — stop
                if t2 == 'table':
                    rows = []
                    for row in obj2.rows:
                        cells = [c.text.strip() for c in row.cells if c.text.strip()]
                        if cells:
                            rows.append(cells)
                    return rows
    return []


def clean_step(text: str) -> str:
    """Strip region noise, extra whitespace, template markers."""
    # Remove template instruction markers (*** ... ***)
    text = re.sub(r'\*+[^*]+\*+', '', text)
    # Remove region/currency noise
    text = re.sub(
        r'\b(germany|german|austria|austrian|uk|united kingdom|india|indian|'
        r'france|french|usa|america|singapore|apac|emea|gbp|usd|eur|inr|sgd|'
        r'vat|gst|tds|mwst|hmrc)\b',
        '', text, flags=re.IGNORECASE
    )
    text = re.sub(r'\s+', ' ', text).strip()
    return text


def extract_steps_from_doc(filepath: str) -> tuple:
    """
    Extract process steps from the document.
    Tries SOP Activities Overview table first, then Detailed Process Steps table.
    Returns (steps_list, source_section, sap_codes).
    """
    try:
        doc = Document(filepath)
    except Exception as e:
        return [], "Error", []

    steps     = []
    source    = None

    # ── Try SOP Activities Overview first ────────────────────────────────────
    # Table structure: Activity ID | Process Description | Objective | Frequency
    rows = extract_table_after_heading(doc, ACTIVITIES_OVERVIEW_HEADINGS)
    if rows and len(rows) > 1:
        # Skip header row, extract "Process Description" column (usually col 1)
        header = [c.lower() for c in rows[0]]
        desc_col = next((i for i, h in enumerate(header)
                        if any(k in h for k in ['description', 'activity', 'process'])), 1)
        for row in rows[1:]:
            if len(row) > desc_col:
                step = clean_step(row[desc_col])
                if 3 <= len(step.split()) <= 50:
                    steps.append(step)
        if steps:
            source = "SOP Activities Overview"

    # ── Try Detailed Process Steps table ─────────────────────────────────────
    # Table structure: Activity ID/Name/Role | Description/Work Instruction | Notes
    detail_rows = extract_table_after_heading(doc, DETAILED_STEPS_HEADINGS)
    if detail_rows and len(detail_rows) > 1:
        header = [c.lower() for c in detail_rows[0]]
        desc_col = next((i for i, h in enumerate(header)
                        if any(k in h for k in ['description', 'detail', 'instruction', 'work'])), 1)
        for row in detail_rows[1:]:
            if len(row) > desc_col:
                step = clean_step(row[desc_col])
                if 3 <= len(step.split()) <= 80:
                    if step not in steps:   # avoid duplicates with overview
                        steps.append(step)
        if steps and not source:
            source = "Detailed Process Steps"
        elif steps:
            source = "Activities Overview + Detailed Steps"

    # ── Fallback: numbered steps in paragraph text ───────────────────────────
    if not steps:
        capturing = False
        all_headings = ACTIVITIES_OVERVIEW_HEADINGS + DETAILED_STEPS_HEADINGS
        for para in doc.paragraphs:
            t = para.text.strip()
            if is_heading(para):
                capturing = text_matches(t, all_headings)
            elif capturing and t:
                # Numbered line: "13.1 Do something" or "1. Do something"
                if re.match(r'^[\d\.]+\s+\w', t):
                    step = clean_step(re.sub(r'^[\d\.]+\s+', '', t))
                    if 3 <= len(step.split()) <= 60:
                        steps.append(step)

        if steps:
            source = "Paragraph steps (fallback)"

    sap_codes = _get_sap_codes(doc)
    return steps, source, sap_codes


# ---------------------------------------------------------------------------
# SAP code extraction
# ---------------------------------------------------------------------------

SAP_PATTERNS = [
    re.compile(r'[Tt][-\s]?[Cc]ode[s]?\s*[:\(]?\s*([A-Z][A-Z0-9_\-]{1,29})\)?'),
    re.compile(r'Run\s+([A-Z][A-Z0-9_\-]{1,29})\s+(?:SAP|[Tt]code|[Tt]ransaction)\b', re.IGNORECASE),
    re.compile(r'Run\s+SAP\s+[Tt]code\s+([A-Z][A-Z0-9_\-]{1,29})', re.IGNORECASE),
    re.compile(r'[Tt]ransaction\s+([A-Z][A-Z0-9_\-]{1,29})\b'),
    re.compile(r'\bOpen\s+[Tt]ransaction\s+([A-Z][A-Z0-9_\-]{1,29})', re.IGNORECASE),
    re.compile(r'\b(S_ALR_[A-Z0-9_]+)\b', re.IGNORECASE),
    re.compile(r'\bGo\s+to\s+(?:the\s+)?transaction\s+([A-Z][A-Z0-9_\-]{1,29})', re.IGNORECASE),
    re.compile(r'\bGo\s+to\s+([A-Z][A-Z0-9_\-]{1,29})\s+in\s+SAP', re.IGNORECASE),
]
NOT_TCODE = {
    "THE","AND","FOR","RUN","SAP","CODE","WITH","FROM","THIS","THAT","WILL",
    "CAN","NOT","ALL","BUT","USE","NEW","OLD","END","ADD","SET","GET","OUT",
    "YES","NO","OK","GO","DO","IN","IS","IT","OF","ON","OR","TO","UP","WE",
    "TCODE","TRANSACTION","REPORT","SYSTEM","TABLE","FIELD","VALUE","MENU",
    "DIFF","DEP","STEP","NOTE","BACK","SAVE","EXIT","VIEW","LIST","OPEN",
}

def _get_sap_codes(doc) -> list:
    codes = set()
    for para in doc.paragraphs:
        for pat in SAP_PATTERNS:
            for m in pat.finditer(para.text):
                c = m.group(1).upper().replace("-","_")
                if c.startswith("T_CODE_"): c = c[7:]
                if (c not in NOT_TCODE and len(c) >= 2
                        and c[0].isalpha() and re.match(r'^[A-Z][A-Z0-9_]+$', c)):
                    codes.add(c)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for pat in SAP_PATTERNS:
                    for m in pat.finditer(cell.text):
                        c = m.group(1).upper().replace("-","_")
                        if c.startswith("T_CODE_"): c = c[7:]
                        if (c not in NOT_TCODE and len(c) >= 2
                                and c[0].isalpha() and re.match(r'^[A-Z][A-Z0-9_]+$', c)):
                            codes.add(c)
    return sorted(codes)


# ---------------------------------------------------------------------------
# Similarity model
# ---------------------------------------------------------------------------

_MODEL = None

def load_model():
    global _MODEL
    if _MODEL is not None:
        return _MODEL
    if EMBEDDING_MODEL is None:
        return None
    try:
        from sentence_transformers import SentenceTransformer
        os.environ["TRANSFORMERS_OFFLINE"] = "1"
        os.environ["HF_DATASETS_OFFLINE"]  = "1"
        _MODEL = SentenceTransformer(EMBEDDING_MODEL, local_files_only=True)
        return _MODEL
    except ImportError:
        return None
    except Exception as e:
        msg = str(e).lower()
        if any(k in msg for k in ("local_files","no such file","not found","snapshot")):
            print(f"\n  [INFO] Model not in local cache. One-time setup:")
            print(f"  [INFO]   pip install sentence-transformers")
            print(f"  [INFO]   python -c \"from sentence_transformers import SentenceTransformer; SentenceTransformer('{EMBEDDING_MODEL}')\"")
            print(f"  [INFO] After that it runs fully offline. Using TF-IDF for now.\n")
        return None


SYNONYMS = {
    "bill":"invoice","bills":"invoices","supplier":"vendor","suppliers":"vendors",
    "creditor":"vendor","wage":"salary","wages":"salaries","entry":"posting",
    "entries":"postings","verify":"validate","check":"review","confirm":"approve",
    "authorise":"approve","authorize":"approve","disposal":"retirement",
    "capitalize":"capitalise","capitalization":"capitalisation",
}

def _expand(text: str) -> str:
    return " ".join(SYNONYMS.get(w.lower(), w.lower()) for w in text.split())


def encode(texts: list, model) -> np.ndarray:
    if model is not None:
        return model.encode(texts, convert_to_numpy=True,
                            show_progress_bar=False, normalize_embeddings=True)
    expanded = [_expand(t) for t in texts]
    vec = TfidfVectorizer(stop_words="english", ngram_range=(1,2),
                          max_features=5000, sublinear_tf=True)
    return vec.fit_transform(expanded).toarray()


def similarity_matrix(texts: list, model) -> np.ndarray:
    vecs = encode(texts, model)
    if model is not None:
        sim = np.dot(vecs, vecs.T)
    else:
        sim = cosine_similarity(vecs)
    return np.clip(sim, 0, 1)


def cluster_docs(sim: np.ndarray, n: int) -> np.ndarray:
    n = max(2, min(n, len(sim) - 1))
    dist = np.clip(1 - sim, 0, None)
    return AgglomerativeClustering(
        n_clusters=n, metric="precomputed", linkage="average"
    ).fit_predict(dist)


def find_matches(doc_steps: list, other_steps: list, model, top_n=TOP_STEPS) -> list:
    """
    For each step in this doc, find best matching step across other docs in cluster.
    Returns [(pct, step_text), ...] sorted by match score descending.
    """
    if not doc_steps:
        return []
    if not other_steps:
        return [(0, s) for s in doc_steps[:top_n]]

    all_other = [s for steps in other_steps for s in steps]
    if not all_other:
        return [(0, s) for s in doc_steps[:top_n]]

    threshold = 0.35 if model else 0.12

    if model is not None:
        e_doc   = model.encode(doc_steps,  convert_to_numpy=True,
                               normalize_embeddings=True, show_progress_bar=False)
        e_other = model.encode(all_other,  convert_to_numpy=True,
                               normalize_embeddings=True, show_progress_bar=False)
        sim = np.dot(e_doc, e_other.T)
    else:
        all_texts = [_expand(s) for s in doc_steps + all_other]
        vec = TfidfVectorizer(stop_words="english", ngram_range=(1,2)).fit_transform(all_texts).toarray()
        sim = cosine_similarity(vec[:len(doc_steps)], vec[len(doc_steps):])

    best_scores = sim.max(axis=1)
    scored = sorted(zip(best_scores, doc_steps), key=lambda x: -x[0])

    results = [(int(s*100), step) for s, step in scored if s >= threshold]
    if not results:
        results = [(0, step) for step in doc_steps[:top_n]]
    return results[:top_n]


# ---------------------------------------------------------------------------
# Group naming
# ---------------------------------------------------------------------------

VERBS  = {"posting","payment","approval","settlement","reconciliation","processing",
          "depreciation","capitalisation","capitalization","retirement","allocation",
          "invoicing","closing","validation","calculation","onboarding","budgeting",
          "reporting","clearance","transfer","accrual","extraction","analysis",
          "capitalization","creation","review","preparation"}
NOUNS  = {"invoice","vendor","asset","payroll","salary","ledger","account","budget",
          "cost","project","employee","payment","journal","balance","capital",
          "depreciation","period","order","expense","bank","data","transaction",
          "request","report","entries","items"}
SKIP   = {"the","and","for","in","on","at","to","of","with","by","from","is","are",
          "was","were","be","have","has","do","does","this","that","it","as","if",
          "not","no","can","will","may","all","any","use","new","old","step","steps",
          "click","select","enter","please","must","user","document","file","system",
          "process","procedure","sap","transaction","tcode","code","run","go","see",
          "note","page","section","based","only","should","whether","also","then",
          "after","before","from","into","such","these","those","each"}

def name_group(all_steps: list, cluster_id: int) -> str:
    text  = " ".join(all_steps).lower()
    words = re.findall(r'\b[a-z]{4,}\b', text)
    freq  = defaultdict(int)
    for w in words:
        if w not in SKIP: freq[w] += 1

    top_verbs = sorted([w for w in freq if w in VERBS],  key=lambda w: -freq[w])
    top_nouns = sorted([w for w in freq if w in NOUNS], key=lambda w: -freq[w])

    if top_verbs and top_nouns:
        return f"{top_nouns[0].title()} {top_verbs[0].title()}"
    elif top_verbs:
        return f"{top_verbs[0].title()} Process"
    elif top_nouns:
        return f"{top_nouns[0].title()} Processing"
    else:
        tops = sorted(freq, key=lambda w: -freq[w])[:2]
        return " ".join(w.title() for w in tops) if tops else f"Group {cluster_id+1}"


# ---------------------------------------------------------------------------
# Excel report
# ---------------------------------------------------------------------------

def build_excel(groups: dict, path: str, engine: str, n_docs: int):
    wb = Workbook()
    ws = wb.active
    ws.title = "SOP Groups"

    HF = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    HB = PatternFill("solid", start_color="1F4E79")
    HA = Alignment(horizontal="center", vertical="center", wrap_text=True)
    GF = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    DF = Font(name="Arial", size=10)
    SF = Font(name="Arial", size=9, italic=True)
    TF = Font(name="Courier New", size=9, color="1F4E79")
    RF = Font(name="Arial", size=9, color="808080", italic=True)
    TA = Alignment(horizontal="left", vertical="top", wrap_text=True)
    MA = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ts  = Side(style="thin",   color="BDD7EE")
    ms  = Side(style="medium", color="1F4E79")
    TB  = Border(left=ts, right=ts, top=ts, bottom=ts)
    MB  = Border(left=ms, right=ms, top=ms, bottom=ms)

    PALETTE = ["1F4E79","2E75B6","70AD47","ED7D31","FFC000",
               "5B9BD5","A9D18E","F4B183","FFD966","9DC3E6"]
    RFILLS  = [PatternFill("solid", start_color=c) for c in
               ["EBF3FB","EFF7E6","FFF2CC","FCE4D6","DAEEF3"]]
    REF_G   = PatternFill("solid", start_color="808080")
    REF_R   = PatternFill("solid", start_color="F2F2F2")

    # Banner
    ws.merge_cells("A1:D1")
    c = ws["A1"]
    n_grp = len([g for g in groups if g != "Reference Documents"])
    c.value = (f"Similarity engine: {engine}  |  "
               f"Documents scanned: {n_docs}  |  "
               f"Process groups: {n_grp}  |  "
               f"Sections: SOP Activities Overview + Detailed Process Steps")
    c.font      = Font(name="Arial", bold=True, color="1F4E79", size=9)
    c.fill      = PatternFill("solid", start_color="DEEAF1")
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    c.border    = MB
    ws.row_dimensions[1].height = 24

    for col, h in enumerate(["Group Name","Document Name",
                              "Similar Process Steps","SAP Transaction Codes"], 1):
        c = ws.cell(row=2, column=col, value=h)
        c.font = HF; c.fill = HB; c.alignment = HA; c.border = MB
    ws.row_dimensions[2].height = 28

    row = 3
    sorted_groups = sorted(groups.items(),
        key=lambda x: ("~" if x[0] == "Reference Documents" else x[0]))

    for g_idx, (gname, docs) in enumerate(sorted_groups):
        is_ref   = (gname == "Reference Documents")
        g_fill   = REF_G if is_ref else PatternFill("solid", start_color=PALETTE[g_idx % len(PALETTE)])
        row_fill = REF_R if is_ref else RFILLS[g_idx % len(RFILLS)]
        start    = row
        n        = len(docs)

        for rel, step_lines, codes in docs:
            c = ws.cell(row=row, column=2, value=rel)
            c.font = DF; c.fill = row_fill; c.alignment = TA; c.border = TB

            steps_text = "\n".join(step_lines) if step_lines else "—"
            c = ws.cell(row=row, column=3, value=steps_text)
            c.font = RF if is_ref else SF
            c.fill = row_fill; c.alignment = TA; c.border = TB

            c = ws.cell(row=row, column=4, value=", ".join(codes) if codes else "—")
            c.font = TF; c.fill = row_fill; c.alignment = TA; c.border = TB

            ws.row_dimensions[row].height = max(45, 14 * len(step_lines))
            row += 1

        c = ws.cell(row=start, column=1, value=gname)
        c.font = GF; c.fill = g_fill; c.alignment = MA; c.border = MB
        if n > 1:
            ws.merge_cells(start_row=start, start_column=1, end_row=row-1, end_column=1)
            c = ws.cell(row=start, column=1)
            c.font = GF; c.fill = g_fill; c.alignment = MA; c.border = MB

    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 45
    ws.column_dimensions["C"].width = 72
    ws.column_dimensions["D"].width = 35
    ws.freeze_panes = "B3"
    wb.save(path)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--groups", "-g", type=int, default=None)
    args = parser.parse_args()

    script_dir  = os.path.dirname(os.path.abspath(__file__))
    input_dir   = os.path.join(script_dir, "input")
    output_dir  = os.path.join(script_dir, "output")
    output_file = os.path.join(output_dir, "SOP_Groups.xlsx")
    os.makedirs(input_dir, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)

    all_files = sorted([
        os.path.join(dp, f)
        for dp, _, fs in os.walk(input_dir)
        for f in fs if f.lower().endswith(".docx") and not f.startswith("~$")
    ])
    if not all_files:
        raise SystemExit(f"No .docx files found in {input_dir}")

    print(f"\nFound {len(all_files)} document(s)\n")

    # ── Extract steps from every document ────────────────────────────────────
    docs = []
    for fp in all_files:
        rel   = os.path.relpath(fp, input_dir)
        steps, source, codes = extract_steps_from_doc(fp)
        text   = " ".join(steps)
        is_ref = len(text.split()) < MIN_WORDS

        if is_ref:
            print(f"  [REF]  {rel}  ({len(steps)} steps — marked Reference Document)")
        else:
            print(f"  [OK]   {rel}  ({len(steps)} steps from '{source}')")

        docs.append({"rel": rel, "steps": steps, "source": source,
                     "codes": codes, "is_ref": is_ref, "text": text})

    active = [d for d in docs if not d["is_ref"]]
    refs   = [d for d in docs if d["is_ref"]]

    if len(active) < 2:
        print(f"\nOnly {len(active)} document(s) with sufficient steps.")
        print("All documents will be grouped as Reference Documents.")
        groups = {"Reference Documents": [(d["rel"],
                   ["Process steps not found or too short"],
                   d["codes"]) for d in docs]}
        build_excel(groups, output_file, "N/A", len(docs))
        print(f"\n✓ Report saved: {output_file}")
        return

    # ── Load model ────────────────────────────────────────────────────────────
    print(f"\nLoading similarity engine...")
    model  = load_model()
    engine = (f"sentence-transformers / {EMBEDDING_MODEL} (offline)"
              if model else "TF-IDF cosine similarity")
    print(f"  Engine: {engine}")

    # ── Cluster ───────────────────────────────────────────────────────────────
    n_clusters = args.groups or max(2, min(8, len(active)-1, round(len(active)/3)))
    print(f"\nClustering {len(active)} documents into {n_clusters} group(s)...\n")

    texts  = [d["text"] for d in active]
    sim    = similarity_matrix(texts, model)
    labels = cluster_docs(sim, n_clusters)

    # ── Build groups ──────────────────────────────────────────────────────────
    cluster_steps = defaultdict(list)
    for i, lbl in enumerate(labels):
        cluster_steps[lbl].extend(active[i]["steps"])

    cluster_names = {lbl: name_group(steps, lbl)
                     for lbl, steps in cluster_steps.items()}

    groups = {}
    for i, d in enumerate(active):
        lbl   = labels[i]
        gname = cluster_names[lbl]

        other = [active[j]["steps"] for j, l in enumerate(labels)
                 if l == lbl and j != i and active[j]["steps"]]

        matched = find_matches(d["steps"], other, model)

        if matched and matched[0][0] > 0:
            lines = [f"• {step}  [{pct}% match]" for pct, step in matched]
        else:
            lines = [f"• {step}" for step in d["steps"][:TOP_STEPS]]

        groups.setdefault(gname, []).append((d["rel"], lines, d["codes"]))
        print(f"  → [{gname}]  {d['rel']}")

    if refs:
        for d in refs:
            groups.setdefault("Reference Documents", []).append(
                (d["rel"],
                 ["No process steps found — check section headings in PROCESS_SECTION_HEADINGS"],
                 d["codes"]))

    # ── Write report ──────────────────────────────────────────────────────────
    build_excel(groups, output_file, engine, len(docs))

    print(f"\n✓ Report saved: {output_file}")
    print(f"\nSummary:")
    for gname, dl in sorted(groups.items(),
                             key=lambda x: "~" if x[0]=="Reference Documents" else x[0]):
        print(f"  {gname}: {len(dl)} document(s)")


if __name__ == "__main__":
    main()
