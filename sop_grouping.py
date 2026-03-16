"""
SOP Document Grouping
======================
Reads Word documents from the input/ folder, extracts the process steps
from a configurable section heading, and groups documents with similar
process steps using semantic similarity.

Output: output/SOP_Groups.xlsx
  Column A: Group Name
  Column B: Document Name
  Column C: Similar Steps Found
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
# CONFIGURATION — edit these to match your documents
# =============================================================================

# The section heading in your Word documents that contains the process steps.
# The script looks for a heading containing this text (case-insensitive).
# Add multiple variants if your documents use different heading names.
PROCESS_SECTION_HEADINGS = [
    "Process Detail",
    "Detailed Process Steps",
    "Process Steps",
    "SOP Activities Overview",
    "Process Activity",
    "Procedure",
]

# Minimum number of words a process section must have to be used for grouping.
# Documents below this are labelled "Reference Document".
MIN_WORDS = 15

# Embedding model for semantic similarity (runs fully offline after download).
# Set to None to always use TF-IDF (no setup needed, less accurate).
EMBEDDING_MODEL = "all-MiniLM-L6-v2"

# Number of similar steps to show in the reason column per document.
TOP_STEPS_TO_SHOW = 5

# =============================================================================


# -----------------------------------------------------------------------------
# Step 1: Extract process steps from each document
# -----------------------------------------------------------------------------

def is_section_heading(para) -> bool:
    """True if paragraph is styled as a heading or is bold all-caps."""
    style = (para.style.name or "").lower()
    if style.startswith("heading") or style.startswith("title"):
        return True
    runs = [r for r in para.runs if r.text.strip()]
    if runs and all(r.bold for r in runs):
        return True
    text = para.text.strip()
    if len(text.split()) <= 8 and text == text.upper() and len(text) > 3:
        return True
    return False


def heading_matches(text: str) -> bool:
    """True if this heading matches one of our PROCESS_SECTION_HEADINGS."""
    t = text.strip().upper()
    for h in PROCESS_SECTION_HEADINGS:
        if h.upper() in t or t in h.upper():
            return True
    return False


def extract_steps(filepath: str) -> tuple:
    """
    Open a .docx and extract paragraphs under the process section heading.
    Returns (steps_list, section_heading_found, full_doc_text).

    Each step is a clean string like "Verify invoice details and vendor number".
    """
    try:
        doc = Document(filepath)
    except Exception as e:
        return [], None, ""

    all_paras = [(is_section_heading(p), p.text.strip())
                 for p in doc.paragraphs if p.text.strip()]

    full_text = " ".join(t for _, t in all_paras)

    # Walk paragraphs: collect text under matching heading
    capturing = False
    heading_found = None
    raw_lines = []

    for is_hdg, text in all_paras:
        if is_hdg:
            if heading_matches(text):
                capturing = True
                heading_found = text
            else:
                if capturing:   # stop at the next unrelated heading
                    break
        elif capturing:
            raw_lines.append(text)

    # Also check tables under that section
    if heading_found:
        in_section = False
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    t = cell.text.strip()
                    if t and not in_section and heading_matches(t):
                        in_section = True
                    elif in_section and t:
                        raw_lines.append(t)

    # Parse raw_lines into individual steps
    steps = []
    for line in raw_lines:
        # Split on numbered patterns: "1.", "2.", "Step 1:", bullets
        parts = re.split(r'(?:^|\n)(?:\d+[.)]\s*|Step\s*\d+\s*[:\-]?\s*|[•\-\*]\s*)', line)
        for part in parts:
            part = part.strip().rstrip(".,;")
            # Remove region noise (country names, currencies)
            part = re.sub(
                r'\b(india|indian|uk|united kingdom|germany|german|france|french|'
                r'usa|america|australia|singapore|apac|emea|gbp|usd|eur|inr|sgd|'
                r'vat|gst|tds|mwst|hmrc)\b',
                '', part, flags=re.IGNORECASE
            )
            part = re.sub(r'\s+', ' ', part).strip()
            if 4 <= len(part.split()) <= 40:
                steps.append(part)

    return steps, heading_found, full_text


# -----------------------------------------------------------------------------
# Step 2: Compute similarity between documents
# -----------------------------------------------------------------------------

def load_model():
    """Load sentence-transformers offline. Returns None if not available."""
    if EMBEDDING_MODEL is None:
        return None
    try:
        from sentence_transformers import SentenceTransformer
        os.environ["TRANSFORMERS_OFFLINE"] = "1"
        os.environ["HF_DATASETS_OFFLINE"]  = "1"
        model = SentenceTransformer(EMBEDDING_MODEL, local_files_only=True)
        return model
    except ImportError:
        return None
    except Exception as e:
        msg = str(e).lower()
        if any(k in msg for k in ("local_files", "no such file", "not found", "snapshot")):
            print(f"\n  [INFO] Model not cached. Run once to download (~90MB):")
            print(f"  [INFO]   pip install sentence-transformers")
            print(f"  [INFO]   python -c \"from sentence_transformers import SentenceTransformer; SentenceTransformer('{EMBEDDING_MODEL}')\"")
            print(f"  [INFO] After that it runs fully offline. Using TF-IDF now.\n")
        return None


def encode_texts(texts: list, model) -> np.ndarray:
    """Encode texts to normalised vectors (embedding or TF-IDF)."""
    if model is not None:
        vecs = model.encode(texts, convert_to_numpy=True,
                            show_progress_bar=False, normalize_embeddings=True)
        return vecs
    else:
        # TF-IDF with synonym expansion
        SYNONYMS = {
            "bill": "invoice", "bills": "invoices", "supplier": "vendor",
            "suppliers": "vendors", "creditor": "vendor", "wage": "salary",
            "wages": "salaries", "entry": "posting", "entries": "postings",
            "verify": "validate", "check": "review", "confirm": "approve",
            "authorise": "approve", "authorize": "approve", "disposal": "retirement",
        }
        def expand(text):
            return " ".join(SYNONYMS.get(w.lower(), w.lower()) for w in text.split())

        expanded = [expand(t) for t in texts]
        vec = TfidfVectorizer(stop_words="english", ngram_range=(1, 2),
                              max_features=5000, sublinear_tf=True)
        return vec.fit_transform(expanded).toarray()


def build_similarity_matrix(doc_step_texts: list, model) -> np.ndarray:
    """
    Build an n×n cosine similarity matrix between documents.
    Each document is represented by its concatenated process steps.
    """
    vecs = encode_texts(doc_step_texts, model)
    if model is not None:
        sim = np.dot(vecs, vecs.T)
    else:
        sim = cosine_similarity(vecs)
    return np.clip(sim, 0, 1)


def cluster(sim_matrix: np.ndarray, n_clusters: int) -> np.ndarray:
    """Agglomerative clustering on distance matrix (1 - similarity)."""
    n = len(sim_matrix)
    n_clusters = max(2, min(n_clusters, n - 1))
    dist = np.clip(1 - sim_matrix, 0, None)
    return AgglomerativeClustering(
        n_clusters=n_clusters, metric="precomputed", linkage="average"
    ).fit_predict(dist)


# -----------------------------------------------------------------------------
# Step 3: Find matching steps between documents in the same cluster
# -----------------------------------------------------------------------------

def find_matching_steps(doc_steps: list, other_steps: list,
                        model, top_n: int = TOP_STEPS_TO_SHOW) -> list:
    """
    For each step in this document, find the most similar step from
    any other document in the same cluster.

    Returns list of (similarity_pct, this_step_text).
    """
    if not doc_steps or not other_steps:
        return [(0, s) for s in doc_steps[:top_n]]

    all_other = [s for steps in other_steps for s in steps]
    if not all_other:
        return [(0, s) for s in doc_steps[:top_n]]

    if model is not None:
        emb_doc   = model.encode(doc_steps,  convert_to_numpy=True,
                                 normalize_embeddings=True, show_progress_bar=False)
        emb_other = model.encode(all_other, convert_to_numpy=True,
                                 normalize_embeddings=True, show_progress_bar=False)
        sim = np.dot(emb_doc, emb_other.T)          # (n_doc, n_other)
        best_scores = sim.max(axis=1)                # best match per doc step
        threshold = 0.40
    else:
        # TF-IDF cosine
        all_texts = doc_steps + all_other
        vecs = TfidfVectorizer(stop_words="english", ngram_range=(1, 2)).fit_transform(all_texts).toarray()
        emb_doc   = vecs[:len(doc_steps)]
        emb_other = vecs[len(doc_steps):]
        sim = cosine_similarity(emb_doc, emb_other)
        best_scores = sim.max(axis=1)
        threshold = 0.15

    scored = sorted(zip(best_scores, doc_steps), key=lambda x: -x[0])
    results = [(int(s * 100), step) for s, step in scored if s >= threshold]
    if not results:
        results = [(0, step) for step in doc_steps[:top_n]]
    return results[:top_n]


# -----------------------------------------------------------------------------
# Step 4: Name the group from its most common step vocabulary
# -----------------------------------------------------------------------------

PROCESS_VERBS = {
    "posting", "payment", "approval", "settlement", "reconciliation",
    "processing", "depreciation", "capitalisation", "retirement", "allocation",
    "invoicing", "closing", "validation", "calculation", "onboarding",
    "budgeting", "reporting", "clearance", "transfer", "accrual",
}
PROCESS_NOUNS = {
    "invoice", "vendor", "asset", "payroll", "salary", "ledger", "account",
    "budget", "cost", "project", "employee", "payment", "journal", "balance",
    "capital", "depreciation", "period", "order", "expense", "bank",
}
EXCLUDE = {
    "the","and","for","in","on","at","to","of","with","by","from","is","are",
    "was","were","be","have","has","do","does","this","that","it","as","if",
    "not","no","can","will","may","all","any","per","etc","use","new","old",
    "step","steps","click","select","enter","open","close","save","screen",
    "please","must","user","document","file","system","process","procedure",
    "sap","transaction","tcode","code","run","go","see","note","menu","page",
}

def name_group(steps_all: list, cluster_id: int) -> str:
    """Generate a process-meaningful group name from all steps in the cluster."""
    text = " ".join(steps_all).lower()
    words = re.findall(r'\b[a-z]{4,}\b', text)
    freq = defaultdict(int)
    for w in words:
        if w not in EXCLUDE:
            freq[w] += 1

    verbs = sorted([w for w in freq if w in PROCESS_VERBS], key=lambda w: -freq[w])
    nouns = sorted([w for w in freq if w in PROCESS_NOUNS], key=lambda w: -freq[w])

    if verbs and nouns:
        return f"{nouns[0].title()} {verbs[0].title()}"
    elif verbs:
        return f"{verbs[0].title()} Process"
    elif nouns:
        return f"{nouns[0].title()} Processing"
    else:
        top = sorted(freq, key=lambda w: -freq[w])[:2]
        return " ".join(w.title() for w in top) if top else f"Group {cluster_id + 1}"


# -----------------------------------------------------------------------------
# Step 5: Extract SAP codes
# -----------------------------------------------------------------------------

SAP_PATTERNS = [
    re.compile(r'[Tt][-\s]?[Cc]ode[s]?\s*[:\(]?\s*([A-Z][A-Z0-9_\-]{1,29})\)?'),
    re.compile(r'Run\s+([A-Z][A-Z0-9_\-]{1,29})\s+(?:SAP|[Tt]code|[Tt]ransaction)\b', re.IGNORECASE),
    re.compile(r'Run\s+SAP\s+[Tt]code\s+([A-Z][A-Z0-9_\-]{1,29})', re.IGNORECASE),
    re.compile(r'[Tt]ransaction\s+([A-Z][A-Z0-9_\-]{1,29})\b'),
    re.compile(r'\bOpen\s+[Tt]ransaction\s+([A-Z][A-Z0-9_\-]{1,29})', re.IGNORECASE),
    re.compile(r'\b(S_ALR_[A-Z0-9_]+)\b', re.IGNORECASE),
]
NOT_TCODE = {
    "THE","AND","FOR","RUN","SAP","CODE","WITH","FROM","THIS","THAT","WILL",
    "CAN","NOT","ALL","BUT","USE","NEW","OLD","END","ADD","SET","GET","OUT",
    "YES","NO","OK","GO","DO","IN","IS","IT","OF","ON","OR","TO","UP","WE",
    "TCODE","TRANSACTION","REPORT","SYSTEM","TABLE","FIELD","VALUE","MENU",
    "DIFF","DEP","STEP","NOTE","BACK","SAVE","EXIT","VIEW","LIST","OPEN",
}

def get_sap_codes(filepath: str) -> list:
    try:
        doc = Document(filepath)
    except Exception:
        return []
    codes = set()
    for para in doc.paragraphs:
        for pat in SAP_PATTERNS:
            for m in pat.finditer(para.text):
                c = m.group(1).upper().replace("-", "_")
                if c.upper().startswith("T_CODE_"):
                    c = c[7:]
                if c not in NOT_TCODE and len(c) >= 2 and c[0].isalpha():
                    if re.match(r'^[A-Z][A-Z0-9_]+$', c):
                        codes.add(c)
    return sorted(codes)


# -----------------------------------------------------------------------------
# Step 6: Build Excel report
# -----------------------------------------------------------------------------

def build_excel(groups: dict, output_path: str, engine: str):
    wb = Workbook()
    ws = wb.active
    ws.title = "SOP Groups"

    # Styles
    H_FONT  = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    H_FILL  = PatternFill("solid", start_color="1F4E79")
    H_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
    G_FONT  = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    D_FONT  = Font(name="Arial", size=10)
    S_FONT  = Font(name="Arial", size=9, italic=True)
    T_FONT  = Font(name="Courier New", size=9, color="1F4E79")
    R_FONT  = Font(name="Arial", size=9, color="808080", italic=True)
    TOP     = Alignment(horizontal="left", vertical="top", wrap_text=True)
    MID     = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin    = Side(style="thin",   color="BDD7EE")
    med     = Side(style="medium", color="1F4E79")
    TB      = Border(left=thin, right=thin, top=thin, bottom=thin)
    MB      = Border(left=med,  right=med,  top=med,  bottom=med)

    COLORS = ["1F4E79","2E75B6","70AD47","ED7D31","FFC000",
              "5B9BD5","A9D18E","F4B183","FFD966","9DC3E6"]
    ROWS   = [PatternFill("solid", start_color=c) for c in
              ["EBF3FB","EFF7E6","FFF2CC","FCE4D6","DAEEF3"]]
    REF_G  = PatternFill("solid", start_color="808080")
    REF_R  = PatternFill("solid", start_color="F2F2F2")

    # Banner
    ws.merge_cells("A1:D1")
    c = ws["A1"]
    c.value = (f"Similarity engine: {engine}  |  "
               f"Process sections: {', '.join(PROCESS_SECTION_HEADINGS[:3])}  |  "
               f"Groups: {len([g for g in groups if g != 'Reference Documents'])}")
    c.font      = Font(name="Arial", bold=True, color="1F4E79", size=9)
    c.fill      = PatternFill("solid", start_color="DEEAF1")
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    c.border    = MB
    ws.row_dimensions[1].height = 22

    for col, h in enumerate(["Group Name","Document Name",
                              "Similar Process Steps","SAP Transaction Codes"], 1):
        c = ws.cell(row=2, column=col, value=h)
        c.font = H_FONT; c.fill = H_FILL
        c.alignment = H_ALIGN; c.border = MB
    ws.row_dimensions[2].height = 28

    row = 3
    sorted_groups = sorted(
        groups.items(),
        key=lambda x: ("~" if x[0] == "Reference Documents" else x[0])
    )

    for g_idx, (group_name, docs) in enumerate(sorted_groups):
        is_ref   = (group_name == "Reference Documents")
        g_fill   = REF_G if is_ref else PatternFill("solid", start_color=COLORS[g_idx % len(COLORS)])
        row_fill = REF_R if is_ref else ROWS[g_idx % len(ROWS)]
        start    = row
        n        = len(docs)

        for rel, step_lines, sap_codes in docs:
            # Column B — document name
            c = ws.cell(row=row, column=2, value=rel)
            c.font = D_FONT; c.fill = row_fill
            c.alignment = TOP; c.border = TB

            # Column C — similar steps
            steps_text = "\n".join(step_lines) if step_lines else "—"
            c = ws.cell(row=row, column=3, value=steps_text)
            c.font = R_FONT if is_ref else S_FONT
            c.fill = row_fill; c.alignment = TOP; c.border = TB

            # Column D — SAP codes
            c = ws.cell(row=row, column=4, value=", ".join(sap_codes) if sap_codes else "—")
            c.font = T_FONT; c.fill = row_fill
            c.alignment = TOP; c.border = TB

            ws.row_dimensions[row].height = max(50, 15 * len(step_lines)) if step_lines else 30
            row += 1

        # Column A — group name, merged
        c = ws.cell(row=start, column=1, value=group_name)
        c.font = G_FONT; c.fill = g_fill
        c.alignment = MID; c.border = MB
        if n > 1:
            ws.merge_cells(start_row=start, start_column=1, end_row=row-1, end_column=1)
            c = ws.cell(row=start, column=1)
            c.font = G_FONT; c.fill = g_fill
            c.alignment = MID; c.border = MB

    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 45
    ws.column_dimensions["C"].width = 70
    ws.column_dimensions["D"].width = 35
    ws.freeze_panes = "B3"
    wb.save(output_path)


# -----------------------------------------------------------------------------
# Main
# -----------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--groups", "-g", type=int, default=None,
                        help="Number of groups (auto if not set)")
    args = parser.parse_args()

    script_dir  = os.path.dirname(os.path.abspath(__file__))
    input_dir   = os.path.join(script_dir, "input")
    output_dir  = os.path.join(script_dir, "output")
    output_file = os.path.join(output_dir, "SOP_Groups.xlsx")
    os.makedirs(input_dir,  exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)

    # ── Collect documents ────────────────────────────────────────────────────
    all_files = sorted([
        os.path.join(dp, f)
        for dp, _, fs in os.walk(input_dir)
        for f in fs if f.lower().endswith(".docx") and not f.startswith("~$")
    ])
    if not all_files:
        raise SystemExit(f"No .docx files found in {input_dir}")

    print(f"\nFound {len(all_files)} documents")
    print(f"Looking for sections: {PROCESS_SECTION_HEADINGS}\n")

    # ── Extract steps ────────────────────────────────────────────────────────
    docs = []       # {rel, steps, heading, sap_codes, is_ref}
    for fp in all_files:
        rel   = os.path.relpath(fp, input_dir)
        steps, heading, _ = extract_steps(fp)
        codes = get_sap_codes(fp)
        text  = " ".join(steps)
        is_ref = len(text.split()) < MIN_WORDS

        if is_ref:
            print(f"  [REF]  {rel}  ({len(steps)} steps — too short, marked Reference)")
        else:
            print(f"  [OK]   {rel}  ({len(steps)} steps from '{heading}')")

        docs.append({
            "rel": rel, "steps": steps, "heading": heading,
            "codes": codes, "is_ref": is_ref, "text": text
        })

    # ── Cluster non-reference docs ───────────────────────────────────────────
    active = [d for d in docs if not d["is_ref"]]
    refs   = [d for d in docs if d["is_ref"]]

    if len(active) < 2:
        raise SystemExit(f"\nNeed at least 2 documents with process steps. Found {len(active)}.")

    n_clusters = args.groups or max(2, min(8, len(active) - 1, round(len(active) / 3)))

    print(f"\nLoading similarity model...")
    model  = load_model()
    engine = f"sentence-transformers / {EMBEDDING_MODEL} (offline)" if model else "TF-IDF cosine similarity"
    print(f"  Engine: {engine}")

    print(f"\nClustering {len(active)} documents into {n_clusters} groups...")
    texts  = [d["text"] for d in active]
    sim    = build_similarity_matrix(texts, model)
    labels = cluster(sim, n_clusters)

    # ── Build groups ─────────────────────────────────────────────────────────
    # Collect all steps per cluster for group naming
    cluster_steps = defaultdict(list)
    for i, lbl in enumerate(labels):
        cluster_steps[lbl].extend(active[i]["steps"])

    cluster_names = {lbl: name_group(steps, lbl)
                     for lbl, steps in cluster_steps.items()}

    groups = {}

    for i, d in enumerate(active):
        lbl        = labels[i]
        group_name = cluster_names[lbl]

        # Get steps from all OTHER docs in same cluster
        other_steps = [active[j]["steps"] for j, l in enumerate(labels)
                       if l == lbl and j != i and active[j]["steps"]]

        matched = find_matching_steps(d["steps"], other_steps, model)

        # Format step lines for Excel
        if matched and matched[0][0] > 0:
            step_lines = [f"• {step}  [{pct}%]" for pct, step in matched]
        else:
            step_lines = [f"• {step}" for step in d["steps"][:TOP_STEPS_TO_SHOW]]

        groups.setdefault(group_name, []).append((d["rel"], step_lines, d["codes"]))
        print(f"  → [{group_name}]  {d['rel']}")

    # Reference documents
    if refs:
        for d in refs:
            step_lines = ["No sufficient process steps found — marked as Reference Document"]
            groups.setdefault("Reference Documents", []).append(
                (d["rel"], step_lines, d["codes"])
            )

    # ── Write Excel ──────────────────────────────────────────────────────────
    build_excel(groups, output_file, engine)

    print(f"\n✓ Report saved: {output_file}")
    print(f"\nSummary:")
    for name, docs_list in sorted(groups.items(), key=lambda x: ("~" if x[0]=="Reference Documents" else x[0])):
        print(f"  {name}: {len(docs_list)} document(s)")


if __name__ == "__main__":
    main()
