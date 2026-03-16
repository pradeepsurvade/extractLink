"""
SOP Document Grouping
======================
Scans all .docx files in the input/ folder (and subfolders), extracts content
from specific sections, groups documents by similarity of that content using
TF-IDF vectors + KMeans clustering, and produces an Excel report:

  Column A: Group Name      (auto-generated from top keywords, merged per group)
  Column B: Document Name   (relative path)
  Column C: Reason          (why this doc is in this group)

How it works:
  1. For each doc, extract text ONLY from headings matching TARGET_SECTIONS
  2. Convert section text to TF-IDF vectors
  3. KMeans clustering groups docs with similar section content
  4. Top keywords per cluster become the group name

Usage:
  python sop_grouping.py               (auto-detects number of groups)
  python sop_grouping.py --groups 5    (force 5 groups)
"""

import os
import re
import math
import argparse
from collections import defaultdict

try:
    from docx import Document
    from docx.oxml.ns import qn
except ImportError:
    raise SystemExit("python-docx not installed. Run: pip install python-docx")

try:
    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.cluster import KMeans
    import numpy as np
except ImportError:
    raise SystemExit("scikit-learn / numpy not installed. Run: pip install scikit-learn numpy")

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
except ImportError:
    raise SystemExit("openpyxl not installed. Run: pip install openpyxl")


# ---------------------------------------------------------------------------
# TARGET SECTIONS
# Add or remove section heading names here (case-insensitive, partial match).
# Only paragraphs that fall under these headings will be used for grouping.
# If a document has none of these sections, its full text is used as fallback.
# ---------------------------------------------------------------------------
TARGET_SECTIONS = [
    "SOP ACTIVITIES OVERVIEW",
    "DETAILED PROCESS STEPS",
    # Add more section names here, e.g.:
    # "SCOPE",
    # "OBJECTIVES",
    # "ROLES AND RESPONSIBILITIES",
]


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
MIN_DOCS           = 2
GROUP_NAME_KEYWORDS = 4
DOC_REASON_KEYWORDS = 6

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
    'via', 'within', 'without', 'where', 'when', 'how', 'what', 'which',
}


# ---------------------------------------------------------------------------
# Section-aware text extraction
# ---------------------------------------------------------------------------

def is_heading(para) -> bool:
    """Return True if paragraph is a Word heading style."""
    style_name = para.style.name.lower() if para.style else ''
    return (
        style_name.startswith('heading')
        or style_name.startswith('title')
        or para.runs and all(r.bold for r in para.runs if r.text.strip())
    )


def heading_matches_target(heading_text: str) -> bool:
    """Return True if heading text matches any TARGET_SECTIONS (case-insensitive, partial)."""
    h = heading_text.strip().upper()
    for target in TARGET_SECTIONS:
        if target.upper() in h or h in target.upper():
            return True
    return False


def extract_section_text(filepath: str) -> tuple:
    """
    Extract text from TARGET_SECTIONS headings only.
    Returns (section_text, sections_found, full_text).

    section_text   : concatenated text of matched sections (used for clustering)
    sections_found : list of section heading names that were found
    full_text      : complete document text (fallback if no sections found)
    """
    doc = Document(filepath)
    paragraphs = doc.paragraphs

    # Build a flat list of (is_heading, text) for all paragraphs
    para_data = []
    for para in paragraphs:
        text = para.text.strip()
        if not text:
            continue
        para_data.append((is_heading(para), text))

    # Also collect table cell text (appended after paragraph scan)
    table_texts = []
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                t = cell.text.strip()
                if t:
                    table_texts.append(t)

    full_text = ' '.join(t for _, t in para_data) + ' ' + ' '.join(table_texts)

    # Walk paragraphs: when we hit a matching heading, collect until next heading
    section_parts = []
    sections_found = []
    capturing = False

    for is_hdg, text in para_data:
        if is_hdg:
            if heading_matches_target(text):
                capturing = True
                sections_found.append(text)
            else:
                capturing = False   # stop at next non-matching heading
        elif capturing:
            section_parts.append(text)

    section_text = ' '.join(section_parts).strip()

    # If no section headings found via style, try text-pattern matching
    # (handles docs where headings are bold paragraphs or ALL CAPS lines)
    if not section_text:
        capturing = False
        for is_hdg, text in para_data:
            upper = text.upper()
            is_target = any(t.upper() in upper or upper in t.upper()
                            for t in TARGET_SECTIONS)
            # Treat ALL-CAPS short lines as headings too
            looks_like_heading = (
                len(text.split()) <= 8 and text == text.upper() and text.isalpha() is False
            ) or is_hdg

            if looks_like_heading and is_target:
                capturing = True
                sections_found.append(text)
            elif looks_like_heading and not is_target:
                capturing = False
            elif capturing:
                section_parts.append(text)

        section_text = ' '.join(section_parts).strip()

    return section_text, sections_found, full_text


def clean_text(text: str) -> str:
    text = text.lower()
    text = re.sub(r'[^a-z\s]', ' ', text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip()


# ---------------------------------------------------------------------------
# Clustering
# ---------------------------------------------------------------------------

def optimal_clusters(n_docs: int, requested: int = None) -> int:
    if requested:
        return min(requested, n_docs)
    k = max(2, min(8, n_docs - 1, round(n_docs / 3)))
    return k


def cluster_documents(texts: list, n_clusters: int):
    vectorizer = TfidfVectorizer(
        stop_words='english',
        max_features=5000,
        ngram_range=(1, 2),
        min_df=1,
        sublinear_tf=True,
    )
    tfidf_matrix = vectorizer.fit_transform(texts)
    km = KMeans(n_clusters=n_clusters, n_init=20, max_iter=300, random_state=42)
    labels = km.fit_predict(tfidf_matrix)
    return labels, vectorizer, tfidf_matrix, km


# ---------------------------------------------------------------------------
# Group naming and reasoning
# ---------------------------------------------------------------------------

def looks_like_tcode(word: str) -> bool:
    w = word.upper()
    if re.match(r'^\d+$', w): return True
    if re.match(r'^S_ALR', w): return True
    if re.match(r'^[ZY][A-Z0-9_]{2,}$', w): return True
    if re.match(r'^[A-Z]{1,3}\d+[A-Z]?$', w): return True
    if '_' in w and len(w) > 6: return True
    return False


def top_cluster_keywords(km, vectorizer, cluster_id: int, n: int = 15) -> list:
    feature_names = vectorizer.get_feature_names_out()
    centroid = km.cluster_centers_[cluster_id]
    top_indices = centroid.argsort()[::-1]
    keywords = []
    for idx in top_indices:
        word = feature_names[idx]
        if (word.lower() not in STOPWORDS
                and len(word) > 3
                and not looks_like_tcode(word)
                and word.isalpha()):
            keywords.append(word)
        if len(keywords) >= n:
            break
    return keywords


def make_group_name(keywords: list, cluster_id: int) -> str:
    if not keywords:
        return f"Group {cluster_id + 1}"
    label_words = [w.title() for w in keywords[:GROUP_NAME_KEYWORDS]]
    return ' / '.join(label_words)


def doc_top_keywords(doc_vector, vectorizer, n: int = DOC_REASON_KEYWORDS) -> list:
    feature_names = vectorizer.get_feature_names_out()
    doc_array = np.asarray(doc_vector.todense()).flatten()
    top_indices = doc_array.argsort()[::-1]
    keywords = []
    for idx in top_indices:
        if doc_array[idx] == 0:
            break
        word = feature_names[idx]
        if word.lower() not in STOPWORDS and len(word) > 3 and not looks_like_tcode(word):
            keywords.append(word)
        if len(keywords) >= n:
            break
    return keywords


def make_reason(doc_keywords: list, cluster_keywords: list,
                group_name: str, sections_found: list, used_fallback: bool) -> str:
    parts = []

    # Mention which sections were used
    if sections_found:
        section_names = '; '.join(s.title() for s in sections_found[:2])
        parts.append(f"Matched sections: {section_names}")
    elif used_fallback:
        parts.append("No target sections found — grouped on full document content")

    # Shared keywords with cluster
    shared = [k for k in doc_keywords if k in cluster_keywords]
    if shared:
        parts.append("Shared topics: " + ', '.join(w.title() for w in shared[:4]))

    # Doc-specific keywords
    unique = [k for k in doc_keywords if k not in cluster_keywords]
    if unique:
        parts.append("Specific focus: " + ', '.join(w.title() for w in unique[:3]))

    return '. '.join(parts) + '.' if parts else f"Grouped with '{group_name}' by content similarity."


# ---------------------------------------------------------------------------
# Excel report
# ---------------------------------------------------------------------------

def build_report(groups: dict, output_path: str, target_sections: list):
    wb = Workbook()
    ws = wb.active
    ws.title = "SOP Groups"

    hdr_font  = Font(name='Arial', bold=True, color='FFFFFF', size=11)
    hdr_fill  = PatternFill('solid', start_color='1F4E79')
    hdr_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    grp_font  = Font(name='Arial', bold=True, color='FFFFFF', size=10)
    doc_font  = Font(name='Arial', size=10)
    rsn_font  = Font(name='Arial', size=10, italic=True)
    top_align = Alignment(horizontal='left', vertical='top', wrap_text=True)
    mid_align = Alignment(horizontal='center', vertical='center', wrap_text=True)

    thin_s = Side(style='thin',   color='BDD7EE')
    med_s  = Side(style='medium', color='1F4E79')
    thin_b = Border(left=thin_s, right=thin_s, top=thin_s, bottom=thin_s)
    med_b  = Border(left=med_s,  right=med_s,  top=med_s,  bottom=med_s)

    PALETTE = [
        '1F4E79','2E75B6','70AD47','ED7D31','FFC000',
        '5B9BD5','A9D18E','F4B183','FFD966','9DC3E6',
    ]
    ROW_FILLS = [
        PatternFill('solid', start_color='EBF3FB'),
        PatternFill('solid', start_color='EFF7E6'),
        PatternFill('solid', start_color='FFF2CC'),
        PatternFill('solid', start_color='FCE4D6'),
        PatternFill('solid', start_color='DAEEF3'),
    ]

    # Banner showing which sections were targeted
    ws.merge_cells('A1:C1')
    banner = ws['A1']
    banner.value = "Grouped by sections: " + " | ".join(target_sections)
    banner.font      = Font(name='Arial', bold=True, color='1F4E79', size=10)
    banner.fill      = PatternFill('solid', start_color='DEEAF1')
    banner.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    banner.border    = med_b
    ws.row_dimensions[1].height = 20

    # Headers
    for col, h in enumerate(['Group Name', 'Document Name', 'Reason for Grouping'], 1):
        c = ws.cell(row=2, column=col, value=h)
        c.font = hdr_font; c.fill = hdr_fill
        c.alignment = hdr_align; c.border = med_b
    ws.row_dimensions[2].height = 28

    row = 3
    for g_idx, (group_name, docs) in enumerate(sorted(groups.items())):
        grp_fill = PatternFill('solid', start_color=PALETTE[g_idx % len(PALETTE)])
        row_fill = ROW_FILLS[g_idx % len(ROW_FILLS)]
        start_row = row
        n = len(docs)

        for rel_path, reason in docs:
            c = ws.cell(row=row, column=2, value=rel_path)
            c.font = doc_font; c.fill = row_fill
            c.alignment = top_align; c.border = thin_b

            c = ws.cell(row=row, column=3, value=reason)
            c.font = rsn_font; c.fill = row_fill
            c.alignment = top_align; c.border = thin_b

            ws.row_dimensions[row].height = 48
            row += 1

        c = ws.cell(row=start_row, column=1, value=group_name)
        c.font = grp_font; c.fill = grp_fill
        c.alignment = mid_align; c.border = med_b

        if n > 1:
            ws.merge_cells(start_row=start_row, start_column=1,
                           end_row=row - 1, end_column=1)
            c = ws.cell(row=start_row, column=1)
            c.font = grp_font; c.fill = grp_fill
            c.alignment = mid_align; c.border = med_b

    ws.column_dimensions['A'].width = 28
    ws.column_dimensions['B'].width = 45
    ws.column_dimensions['C'].width = 78
    ws.freeze_panes = 'B3'
    wb.save(output_path)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description='Group SOP documents by section content similarity.')
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
    print(f"Output folder : {output_dir}")
    print(f"Target sections: {TARGET_SECTIONS}\n")

    # Collect .docx files
    all_docx = []
    for dirpath, _, filenames in os.walk(input_dir):
        for f in filenames:
            if f.lower().endswith('.docx') and not f.startswith('~$'):
                all_docx.append(os.path.join(dirpath, f))
    all_docx.sort()

    if not all_docx:
        raise SystemExit(f"No .docx files found in '{input_dir}' or any subfolders.")

    print(f"Found {len(all_docx)} document(s). Extracting section text...\n")

    rel_paths, clean_texts, meta = [], [], []

    for filepath in all_docx:
        rel = os.path.relpath(filepath, input_dir)
        try:
            section_text, sections_found, full_text = extract_section_text(filepath)
        except Exception as e:
            print(f"  [WARN] Could not read '{rel}': {e}")
            continue

        used_fallback = False
        if len(section_text.split()) >= 10:
            text_for_cluster = section_text
            print(f"  + {rel}")
            print(f"      Sections found: {sections_found if sections_found else 'matched via text pattern'}")
        else:
            # Fallback to full text if section not found
            text_for_cluster = full_text
            used_fallback = True
            print(f"  ~ {rel} [FALLBACK: target sections not found, using full text]")

        if len(clean_text(text_for_cluster).split()) < 5:
            print(f"  [SKIP] '{rel}' has too little text.")
            continue

        rel_paths.append(rel)
        clean_texts.append(clean_text(text_for_cluster))
        meta.append({
            'rel': rel,
            'sections_found': sections_found,
            'used_fallback': used_fallback,
            'section_text': section_text,
        })

    n_docs = len(rel_paths)
    if n_docs < MIN_DOCS:
        raise SystemExit(f"\nNeed at least {MIN_DOCS} readable documents. Found {n_docs}.")

    # Cluster
    n_clusters = optimal_clusters(n_docs, args.groups)
    print(f"\nClustering {n_docs} document(s) into {n_clusters} group(s)...\n")

    labels, vectorizer, tfidf_matrix, km = cluster_documents(clean_texts, n_clusters)

    # Pre-compute cluster keywords and names
    cluster_kws   = {i: top_cluster_keywords(km, vectorizer, i) for i in range(n_clusters)}
    cluster_names = {i: make_group_name(cluster_kws[i], i)      for i in range(n_clusters)}

    groups = {}
    for idx, label in enumerate(labels):
        m          = meta[idx]
        group_name = cluster_names[label]
        c_kws      = cluster_kws[label]
        doc_kws    = doc_top_keywords(tfidf_matrix[idx], vectorizer)
        reason     = make_reason(doc_kws, c_kws, group_name,
                                 m['sections_found'], m['used_fallback'])
        groups.setdefault(group_name, []).append((m['rel'], reason))
        print(f"  [{group_name}]  {m['rel']}")

    build_report(groups, output_file, TARGET_SECTIONS)

    print(f"\nReport saved to: {output_file}")
    print(f"  {n_clusters} group(s), {n_docs} document(s) grouped.\n")
    print("Group summary:")
    for name, docs in sorted(groups.items()):
        print(f"  {name}: {len(docs)} doc(s)")


if __name__ == '__main__':
    main()
