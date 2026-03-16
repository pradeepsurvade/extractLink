"""
SOP Document Grouping
======================
Scans all .docx files in the input/ folder (and subfolders), groups them by
content similarity using TF-IDF vectors + KMeans clustering, and produces
an Excel report with:

  Column A: Group Name      (auto-generated from top keywords, merged per group)
  Column B: Document Name   (relative path)
  Column C: Reason          (why this doc was placed in this group)

How it works:
  1. Extract full text from each .docx
  2. Convert to TF-IDF vectors (word frequency weighted by rarity across docs)
  3. KMeans clustering to group similar documents
  4. Top keywords per cluster become the group name
  5. Per-document reason = its top keywords that align with the cluster theme

Usage:
  python sop_grouping.py               (auto-detects number of groups)
  python sop_grouping.py --groups 5    (force 5 groups)
  python sop_grouping.py --groups 5 --min-cluster-size 2
"""

import os
import re
import sys
import math
import argparse
from collections import defaultdict, Counter

try:
    from docx import Document
except ImportError:
    raise SystemExit("python-docx not installed. Run: pip install python-docx")

try:
    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.cluster import KMeans
    from sklearn.metrics import silhouette_score
    import numpy as np
except ImportError:
    raise SystemExit("scikit-learn / numpy not installed. Run: pip install scikit-learn numpy")

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
except ImportError:
    raise SystemExit("openpyxl not installed. Run: pip install openpyxl")


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

# Minimum documents needed to attempt clustering
MIN_DOCS = 2

# Max keywords used in group name
GROUP_NAME_KEYWORDS = 4

# Max keywords shown in per-document reason
DOC_REASON_KEYWORDS = 6

# Words to ignore when generating group names / reasons (domain stopwords)
STOPWORDS = {
    'the', 'a', 'an', 'and', 'or', 'but', 'in', 'on', 'at', 'to', 'for',
    'of', 'with', 'by', 'from', 'is', 'are', 'was', 'were', 'be', 'been',
    'being', 'have', 'has', 'had', 'do', 'does', 'did', 'will', 'would',
    'could', 'should', 'may', 'might', 'shall', 'can', 'this', 'that',
    'these', 'those', 'it', 'its', 'as', 'if', 'then', 'than', 'so',
    'not', 'no', 'nor', 'up', 'out', 'into', 'through', 'during', 'before',
    'after', 'above', 'below', 'between', 'each', 'all', 'both', 'few',
    'more', 'most', 'other', 'some', 'such', 'any', 'per', 'also', 'etc',
    'please', 'must', 'user', 'step', 'steps', 'click', 'select', 'enter',
    'open', 'close', 'following', 'using', 'used', 'use', 'new', 'save',
    'screen', 'button', 'field', 'fields', 'form', 'page', 'section',
    'document', 'documents', 'file', 'files', 'system', 'sap', 'transaction',
    'tcode', 'code', 'process', 'procedure', 'go', 'see', 'note', 'menu',
    've', 'will', 'upon', 'via', 'within', 'without', 'where', 'when',
    'how', 'what', 'which', 'who', 'run', 'based', 'required', 'ensure',
}


# ---------------------------------------------------------------------------
# Text extraction
# ---------------------------------------------------------------------------

def extract_text(filepath: str) -> str:
    """Extract all text from a .docx file as a single string."""
    doc = Document(filepath)
    parts = []

    for para in doc.paragraphs:
        t = para.text.strip()
        if t:
            parts.append(t)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                t = cell.text.strip()
                if t:
                    parts.append(t)

    return ' '.join(parts)


def clean_text(text: str) -> str:
    """Lowercase and remove non-alphabetic characters for analysis."""
    text = text.lower()
    text = re.sub(r'[^a-z\s]', ' ', text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip()


# ---------------------------------------------------------------------------
# Clustering
# ---------------------------------------------------------------------------

def optimal_clusters(n_docs: int, requested: int = None) -> int:
    """Determine a sensible number of clusters."""
    if requested:
        return min(requested, n_docs)
    # Heuristic: n/3 rounded, clamped between 2 and min(8, n-1)
    k = max(2, min(8, n_docs - 1, round(n_docs / 3)))
    return k


def cluster_documents(texts: list, n_clusters: int):
    """
    Vectorise texts with TF-IDF and cluster with KMeans.
    Returns (labels, vectorizer, tfidf_matrix).
    """
    vectorizer = TfidfVectorizer(
        stop_words='english',
        max_features=5000,
        ngram_range=(1, 2),   # unigrams + bigrams for better phrases
        min_df=1,
        sublinear_tf=True,    # dampens very frequent terms
    )

    tfidf_matrix = vectorizer.fit_transform(texts)

    km = KMeans(
        n_clusters=n_clusters,
        n_init=20,            # more initialisations = more stable result
        max_iter=300,
        random_state=42,
    )
    labels = km.fit_predict(tfidf_matrix)

    return labels, vectorizer, tfidf_matrix, km


# ---------------------------------------------------------------------------
# Group naming and reasoning
# ---------------------------------------------------------------------------

def looks_like_tcode(word: str) -> bool:
    """Return True if the word looks like a SAP transaction code or code fragment."""
    import re as _re
    w = word.upper()
    # Patterns: pure digits, S_ALR, Z/Y codes, short alphanum codes like CJ20N
    if _re.match(r'^\d+$', w): return True
    if _re.match(r'^S_ALR', w): return True
    if _re.match(r'^[ZY][A-Z0-9_]{2,}$', w): return True
    if _re.match(r'^[A-Z]{1,3}\d+[A-Z]?$', w): return True
    if '_' in w and len(w) > 6: return True
    return False


def top_cluster_keywords(km, vectorizer, cluster_id: int, n: int = 10) -> list:
    """Return the top n meaningful keywords for a cluster centroid."""
    feature_names = vectorizer.get_feature_names_out()
    centroid = km.cluster_centers_[cluster_id]
    top_indices = centroid.argsort()[::-1]

    keywords = []
    for idx in top_indices:
        word = feature_names[idx]
        # Skip stopwords, short tokens, and SAP code patterns
        if (word.lower() not in STOPWORDS
                and len(word) > 3
                and not looks_like_tcode(word)
                and word.isalpha()):   # skip hyphenated / numeric tokens in group name
            keywords.append(word)
        if len(keywords) >= n:
            break
    return keywords


def make_group_name(keywords: list, cluster_id: int) -> str:
    """Generate a readable group name from top keywords."""
    if not keywords:
        return f"Group {cluster_id + 1}"
    # Capitalise each keyword, join with ' / '
    label_words = [w.replace('_', ' ').title() for w in keywords[:GROUP_NAME_KEYWORDS]]
    return ' / '.join(label_words)


def doc_top_keywords(doc_vector, vectorizer, cluster_keywords: list, n: int = DOC_REASON_KEYWORDS) -> list:
    """
    Return the top keywords for a specific document that overlap with
    or complement the cluster theme.
    """
    feature_names = vectorizer.get_feature_names_out()
    # Get non-zero features for this document
    doc_array = np.asarray(doc_vector.todense()).flatten()
    top_indices = doc_array.argsort()[::-1]

    keywords = []
    for idx in top_indices:
        if doc_array[idx] == 0:
            break
        word = feature_names[idx]
        if word.lower() not in STOPWORDS and len(word) > 2:
            keywords.append(word)
        if len(keywords) >= n:
            break
    return keywords


def make_reason(doc_keywords: list, cluster_keywords: list, group_name: str) -> str:
    """Generate a human-readable reason for the grouping."""
    if not doc_keywords:
        return f"Grouped under '{group_name}' based on overall content similarity."

    shared = [k for k in doc_keywords if k in cluster_keywords]
    unique = [k for k in doc_keywords if k not in cluster_keywords]

    parts = []
    if shared:
        shared_str = ', '.join(w.title() for w in shared[:3])
        parts.append(f"Shares key topics with group: {shared_str}")
    if unique:
        unique_str = ', '.join(w.title() for w in unique[:3])
        parts.append(f"Doc-specific focus: {unique_str}")

    if parts:
        return '. '.join(parts) + '.'
    return f"Content similarity with '{group_name}' group."


# ---------------------------------------------------------------------------
# Excel report
# ---------------------------------------------------------------------------

def build_report(groups: dict, output_path: str):
    """
    groups: {group_name: [(rel_path, reason), ...]}
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "SOP Groups"

    # Styles
    hdr_font    = Font(name='Arial', bold=True, color='FFFFFF', size=11)
    hdr_fill    = PatternFill('solid', start_color='1F4E79')
    hdr_align   = Alignment(horizontal='center', vertical='center', wrap_text=True)

    grp_font    = Font(name='Arial', bold=True, color='FFFFFF', size=10)
    doc_font    = Font(name='Arial', size=10)
    reason_font = Font(name='Arial', size=10, italic=True)
    top_align   = Alignment(horizontal='left', vertical='top', wrap_text=True)
    mid_align   = Alignment(horizontal='center', vertical='center', wrap_text=True)

    thin_s  = Side(style='thin',   color='BDD7EE')
    med_s   = Side(style='medium', color='1F4E79')
    thin_b  = Border(left=thin_s, right=thin_s, top=thin_s, bottom=thin_s)
    med_b   = Border(left=med_s,  right=med_s,  top=med_s,  bottom=med_s)

    # Distinct fill colours per group (cycles through palette)
    PALETTE = [
        '1F4E79', '2E75B6', '70AD47', 'ED7D31', 'FFC000',
        '5B9BD5', 'A9D18E', 'F4B183', 'FFD966', '9DC3E6',
    ]

    alt_fills = [
        PatternFill('solid', start_color='EBF3FB'),
        PatternFill('solid', start_color='EFF7E6'),
        PatternFill('solid', start_color='FFF2CC'),
        PatternFill('solid', start_color='FCE4D6'),
        PatternFill('solid', start_color='DAEEF3'),
    ]
    base_fill = PatternFill('solid', start_color='FFFFFF')

    # Header
    for col, h in enumerate(['Group Name', 'Document Name', 'Reason for Grouping'], 1):
        c = ws.cell(row=1, column=col, value=h)
        c.font = hdr_font; c.fill = hdr_fill
        c.alignment = hdr_align; c.border = med_b
    ws.row_dimensions[1].height = 30

    row = 2
    for g_idx, (group_name, docs) in enumerate(sorted(groups.items())):
        grp_fill_hex = PALETTE[g_idx % len(PALETTE)]
        grp_fill     = PatternFill('solid', start_color=grp_fill_hex)
        row_fill     = alt_fills[g_idx % len(alt_fills)]

        start_row = row
        n = len(docs)

        for rel_path, reason in docs:
            c = ws.cell(row=row, column=2, value=rel_path)
            c.font = doc_font; c.fill = row_fill
            c.alignment = top_align; c.border = thin_b

            c = ws.cell(row=row, column=3, value=reason)
            c.font = reason_font; c.fill = row_fill
            c.alignment = top_align; c.border = thin_b

            ws.row_dimensions[row].height = 45
            row += 1

        # Group name in col A — merged across all doc rows
        c = ws.cell(row=start_row, column=1, value=group_name)
        c.font = grp_font; c.fill = grp_fill
        c.alignment = mid_align; c.border = med_b

        if n > 1:
            ws.merge_cells(start_row=start_row, start_column=1,
                           end_row=row - 1,    end_column=1)
            c = ws.cell(row=start_row, column=1)
            c.font = grp_font; c.fill = grp_fill
            c.alignment = mid_align; c.border = med_b

    ws.column_dimensions['A'].width = 28
    ws.column_dimensions['B'].width = 42
    ws.column_dimensions['C'].width = 75
    ws.freeze_panes = 'B2'

    wb.save(output_path)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description='Group SOP documents by content similarity.'
    )
    parser.add_argument('--groups', '-g', type=int, default=None,
                        help='Number of groups (auto-detected if not specified)')
    args = parser.parse_args()

    script_dir  = os.path.dirname(os.path.abspath(__file__))
    input_dir   = os.path.join(script_dir, 'input')
    output_dir  = os.path.join(script_dir, 'output')
    output_file = os.path.join(output_dir, 'SOP_Groups.xlsx')

    os.makedirs(input_dir,  exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)

    print(f"Input  folder : {input_dir}")
    print(f"Output folder : {output_dir}\n")

    # ── Collect all .docx files ───────────────────────────────────────────────
    all_docx = []
    for dirpath, _, filenames in os.walk(input_dir):
        for f in filenames:
            if f.lower().endswith('.docx') and not f.startswith('~$'):
                all_docx.append(os.path.join(dirpath, f))
    all_docx.sort()

    if not all_docx:
        raise SystemExit(
            f"No .docx files found in '{input_dir}' or any subfolders.\n"
            "Place your Word documents in the input/ folder and try again."
        )

    print(f"Found {len(all_docx)} document(s). Extracting text...\n")

    # ── Extract text ──────────────────────────────────────────────────────────
    rel_paths, raw_texts, clean_texts = [], [], []

    for filepath in all_docx:
        rel = os.path.relpath(filepath, input_dir)
        try:
            raw  = extract_text(filepath)
            cln  = clean_text(raw)
        except Exception as e:
            print(f"  [WARN] Could not read '{rel}': {e}")
            continue

        if len(cln.split()) < 10:
            print(f"  [SKIP] '{rel}' has too little text to cluster.")
            continue

        rel_paths.append(rel)
        raw_texts.append(raw)
        clean_texts.append(cln)
        print(f"  + {rel} ({len(cln.split())} words)")

    n_docs = len(rel_paths)
    if n_docs < MIN_DOCS:
        raise SystemExit(f"\nNeed at least {MIN_DOCS} readable documents to cluster. Found {n_docs}.")

    # ── Cluster ───────────────────────────────────────────────────────────────
    n_clusters = optimal_clusters(n_docs, args.groups)
    print(f"\nClustering {n_docs} document(s) into {n_clusters} group(s)...\n")

    labels, vectorizer, tfidf_matrix, km = cluster_documents(clean_texts, n_clusters)

    # ── Build group name and per-doc reason ───────────────────────────────────
    groups = {}   # group_name -> [(rel_path, reason), ...]

    # Pre-compute cluster keywords
    cluster_kws = {
        i: top_cluster_keywords(km, vectorizer, i, n=15)
        for i in range(n_clusters)
    }
    cluster_names = {
        i: make_group_name(cluster_kws[i], i)
        for i in range(n_clusters)
    }

    for idx, (rel_path, label) in enumerate(zip(rel_paths, labels)):
        group_name    = cluster_names[label]
        c_keywords    = cluster_kws[label]
        doc_vec       = tfidf_matrix[idx]
        doc_keywords  = doc_top_keywords(doc_vec, vectorizer, c_keywords)
        reason        = make_reason(doc_keywords, c_keywords, group_name)

        groups.setdefault(group_name, []).append((rel_path, reason))
        print(f"  [{group_name}]  {rel_path}")

    # ── Write Excel ───────────────────────────────────────────────────────────
    build_report(groups, output_file)

    print(f"\nReport saved to: {output_file}")
    print(f"  {n_clusters} group(s), {n_docs} document(s) grouped.")

    # Print summary
    print("\nGroup summary:")
    for name, docs in sorted(groups.items()):
        print(f"  {name}: {len(docs)} doc(s)")


if __name__ == '__main__':
    main()
