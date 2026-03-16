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
    from sklearn.cluster import KMeans, AgglomerativeClustering
    from sklearn.metrics.pairwise import cosine_similarity as sklearn_cosine
    import numpy as np
except ImportError:
    raise SystemExit("scikit-learn / numpy not installed. Run: pip install scikit-learn numpy")

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
except ImportError:
    raise SystemExit("openpyxl not installed. Run: pip install openpyxl")


# ---------------------------------------------------------------------------
# SAP T-code extraction (inlined from extract_sap_tcodes logic)
# ---------------------------------------------------------------------------
_TCODE_TRIGGER_PATTERNS = [
    re.compile(r'Run\s+SAP\s+[Tt][-\s]?[Cc]ode\s+([A-Z][A-Z0-9_\-]{1,29})', re.IGNORECASE),
    re.compile(r'Run\s+the\s+([A-Z][A-Z0-9_\-]{1,29})\s+[Tt]code', re.IGNORECASE),
    re.compile(r'Run\s+([A-Z][A-Z0-9_\-]{1,29})\s+SAP\s+[Tt]code', re.IGNORECASE),
    re.compile(r'Run\s+([A-Z][A-Z0-9_\-]{1,29})\s+(?:SAP|[Tt][-\s]?[Cc]ode|[Tt]ransaction)\b', re.IGNORECASE),
    re.compile(r'[Tt][-\s]?[Cc]ode[s]?\s*[:\(]?\s*([A-Z][A-Z0-9_\-]{1,29})\)?', re.IGNORECASE),
    re.compile(r'SAP\s+[Tt]ransaction\s*[:\(]?\s*([A-Z][A-Z0-9_\-]{1,29})\)?', re.IGNORECASE),
    re.compile(r'SAP\s+\w[\w\s]{0,40}[Tt]ransaction\s*\(\s*([A-Z][A-Z0-9_\-]{1,29})\s*\)', re.IGNORECASE),
    re.compile(r'[Tt]ransaction\s+[Cc]ode\s*[:\(]?\s*([A-Z][A-Z0-9_\-]{1,29})\)?', re.IGNORECASE),
    re.compile(r'\bOpen\s+Transaction\s*[\u201c\u201d\u2018\u2019"\'"]?([A-Z][A-Z0-9_\-]{1,29})[\u201c\u201d\u2018\u2019"\'"]?', re.IGNORECASE),
    re.compile(r'\bOpen\s+Transaction\s+([A-Z][A-Z0-9_\-]{1,29})\b', re.IGNORECASE),
    re.compile(r'\bTransaction\s+([A-Z][A-Z0-9_\-]{1,29})\b'),
    re.compile(r'\b(S_ALR_[A-Z0-9_]+)\b', re.IGNORECASE),
]

_NOT_A_TCODE = {
    'THE','AND','FOR','RUN','SAP','CODE','CODES','WITH','FROM','INTO','THIS',
    'THAT','HAVE','BEEN','WILL','ALSO','CAN','NOT','ARE','ALL','BUT','USE',
    'NEW','OLD','END','ADD','SET','GET','PUT','OUT','OFF','TOP','BOX','YES',
    'NO','OK','GO','DO','IF','AS','AT','BE','BY','IN','IS','IT','OF','ON',
    'OR','SO','TO','UP','US','WE','ME','TCODE','TCODES','TRANSACTION',
    'REPORT','MODULE','SYSTEM','TABLE','FIELD','VALUE','PLEASE','CLICK',
    'OPEN','CLOSE','ENTER','SELECT','PRESS','BUTTON','SCREEN','WINDOW',
    'MENU','LIST','VIEW','NEXT','BACK','SAVE','EXIT','NOTE','STEP','THEN',
    'WHEN','ONCE','AFTER','BEFORE','USING','BELOW','ABOVE','RIGHT','LEFT',
    'HERE','THERE','DIFF','DEP',
}

_TABLE_CODE_HEADERS = {
    'process terms','process acronyms','process term','process acronym',
    'transaction code','transaction codes','t-code','tcode','t code',
    'sap transaction','sap tcode','sap t-code','sap code','sap codes',
}

_TABLE_IGNORE = {
    'NA','N/A','N.A','N.A.','NONE','NIL','TBD','TBC','-','--','---',
    'WBS','WNS','N','A','YES','NO','TRUE','FALSE',
}


def _is_valid_tcode(candidate: str) -> bool:
    c = candidate.upper().strip().rstrip('.,;)(')
    if not c or c in _NOT_A_TCODE or len(c) < 2 or not c[0].isalpha():
        return False
    if not re.match(r'^[A-Z][A-Z0-9_]+$', c):
        return False
    prefix = 'T_CODE_'
    if c.startswith(prefix):
        c = c[len(prefix):]
    return bool(c)


def _strip_prefix(code: str) -> str:
    if code.upper().startswith('T_CODE_'):
        return code[7:]
    return code


def _extract_tcodes_from_text(text: str) -> set:
    tcodes = set()
    seen = set()
    for pat in _TCODE_TRIGGER_PATTERNS:
        for m in pat.finditer(text):
            raw = m.group(1).strip().rstrip('.,;)(').upper().replace('-', '_')
            raw = _strip_prefix(raw)
            if _is_valid_tcode(raw) and raw not in seen:
                seen.add(raw)
                tcodes.add(raw)
    return tcodes


def extract_sap_codes(filepath: str) -> list:
    """Extract all SAP transaction codes from a .docx file."""
    doc = Document(filepath)
    tcodes = set()

    for para in doc.paragraphs:
        tcodes |= _extract_tcodes_from_text(para.text)

    for table in doc.tables:
        rows = list(table.rows)
        if not rows:
            continue
        # Check if first row is a SAP-code table header
        sap_cols = set()
        for ci, cell in enumerate(rows[0].cells):
            if cell.text.strip().lower() in _TABLE_CODE_HEADERS:
                sap_cols.add(ci)

        for ri, row in enumerate(rows):
            for ci, cell in enumerate(row.cells):
                t = cell.text.strip()
                if not t:
                    continue
                if ci in sap_cols and ri > 0:
                    cand = t.upper().replace('-', '_')
                    cand = _strip_prefix(cand)
                    if cand not in _TABLE_IGNORE and _is_valid_tcode(cand):
                        tcodes.add(cand)
                else:
                    tcodes |= _extract_tcodes_from_text(t)

    return sorted(tcodes)


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
# GROUPING HINTS (PROMPT)
# Write plain English instructions to control how documents are grouped.
# Each string is one instruction. Uncomment or add lines freely.
#
# FOCUS KEYWORDS — special phrases that change WHAT content is compared:
#   "only sap code"              → compare only the SAP transaction codes per doc
#   "only process steps"         → compare only the Detailed Process Steps section
#   "sap code and process steps" → compare SAP codes + process steps together
#   "only activities overview"   → compare only the SOP Activities Overview section
#   (if no focus keyword, compares all TARGET_SECTIONS content as before)
#
# GROUPING INSTRUCTIONS — plain language to nudge cluster boundaries:
#   "group GL reconciliation and journal posting together"
#   "keep asset depreciation separate from payroll"
#   "treat onboarding and payroll as one HR group"
# ---------------------------------------------------------------------------
GROUPING_HINTS = [
    # "only sap code",
    # "only process steps",
    # "sap code and process steps",
    # "group GL reconciliation and journal posting together",
    # "keep asset depreciation separate from payroll",
]


# ---------------------------------------------------------------------------
# EMBEDDING MODEL — runs 100% OFFLINE after one-time setup
#
# YOUR DATA NEVER GOES TO THE INTERNET:
#   - The model is downloaded ONCE to your local machine cache
#   - All processing happens entirely on your CPU, in local memory
#   - TRANSFORMERS_OFFLINE=1 is set in code to block any network calls
#   - Verified offline: no API key, no cloud, no data transmission
#
# ONE-TIME SETUP (do this once, needs internet only for download):
#   Step 1:  pip install sentence-transformers
#   Step 2:  python -c "from sentence_transformers import SentenceTransformer; SentenceTransformer('all-MiniLM-L6-v2')"
#   After that: run sop_grouping.py normally — fully offline forever
#
# If model is not cached, script automatically falls back to TF-IDF (also offline).
#
# Model options (all open-source, Apache 2.0 / MIT licensed):
#   "all-MiniLM-L6-v2"          — recommended: best balance of accuracy + speed (~90MB)
#   "all-MiniLM-L12-v2"         — more accurate, slightly larger (~120MB)
#   "paraphrase-MiniLM-L3-v2"   — fastest, smallest (~60MB), slightly less accurate
# ---------------------------------------------------------------------------
EMBEDDING_MODEL = "all-MiniLM-L6-v2"


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

def parse_focus_mode(hints: list) -> str:
    """
    Scan GROUPING_HINTS for focus keywords and return a mode string:
      'sap_only'        → use only SAP codes as the comparison vector
      'steps_only'      → use only Detailed Process Steps section text
      'activities_only' → use only SOP Activities Overview section text
      'sap_and_steps'   → use SAP codes + process steps text combined
      'all'             → use all TARGET_SECTIONS text (default)
    """
    combined = ' '.join(hints).lower()

    if 'only sap code' in combined or 'sap code only' in combined:
        return 'sap_only'
    if 'sap code and process steps' in combined or 'process steps and sap code' in combined:
        return 'sap_and_steps'
    if 'only process steps' in combined or 'process steps only' in combined:
        return 'steps_only'
    if 'only activities overview' in combined or 'activities overview only' in combined:
        return 'activities_only'
    return 'all'


def apply_focus_mode(mode: str, section_text: str, full_text: str,
                     sap_codes: list, filepath: str) -> str:
    """
    Return the text string that will be used for TF-IDF clustering,
    based on the active focus mode.
    """
    if mode == 'sap_only':
        # Compare purely on SAP codes — join them as space-separated words
        return ' '.join(sap_codes) if sap_codes else section_text or full_text

    if mode == 'sap_and_steps':
        # SAP codes + process steps section text
        steps_text = _extract_named_section(filepath, 'DETAILED PROCESS STEPS')
        combined = ' '.join(sap_codes) + ' ' + (steps_text or section_text)
        return combined.strip() or full_text

    if mode == 'steps_only':
        steps_text = _extract_named_section(filepath, 'DETAILED PROCESS STEPS')
        return steps_text or section_text or full_text

    if mode == 'activities_only':
        act_text = _extract_named_section(filepath, 'SOP ACTIVITIES OVERVIEW')
        return act_text or section_text or full_text

    # Default: all TARGET_SECTIONS content
    return section_text or full_text


def _extract_named_section(filepath: str, section_name: str) -> str:
    """Extract text from a specific named section heading in a docx."""
    try:
        doc = Document(filepath)
        parts = []
        capturing = False
        for para in doc.paragraphs:
            text = para.text.strip()
            if not text:
                continue
            is_hdg = is_heading(para) or (len(text.split()) <= 8 and text == text.upper())
            if is_hdg:
                if section_name.upper() in text.upper():
                    capturing = True
                else:
                    capturing = False
            elif capturing:
                parts.append(text)
        return ' '.join(parts).strip()
    except Exception:
        return ''


def optimal_clusters(n_docs: int, requested: int = None) -> int:
    if requested:
        return min(requested, n_docs)
    k = max(2, min(8, n_docs - 1, round(n_docs / 3)))
    return k


# ---------------------------------------------------------------------------
# Synonym map — maps variant words to a canonical form so that
# "invoice" and "bill", "vendor" and "supplier", etc. are treated as the same.
# Add more pairs here to improve semantic matching.
# ---------------------------------------------------------------------------
SYNONYM_MAP = {
    # Invoice / payment variants
    'bill': 'invoice',        'bills': 'invoices',
    'billing': 'invoicing',   'receipt': 'invoice',
    'supplier': 'vendor',     'suppliers': 'vendors',
    'creditor': 'vendor',     'creditors': 'vendors',
    # Asset variants
    'disposal': 'retirement', 'write-off': 'retirement',
    'writeoff': 'retirement', 'scrapping': 'retirement',
    'purchase': 'acquisition','bought': 'acquisition',
    # GL / journal variants
    'entry': 'posting',       'entries': 'postings',
    'booking': 'posting',     'bookings': 'postings',
    'journal': 'ledger',      'journals': 'ledgers',
    'general ledger': 'ledger',
    # Payroll variants
    'wage': 'salary',         'wages': 'salaries',
    'compensation': 'salary', 'remuneration': 'salary',
    'headcount': 'employee',  'staff': 'employee',
    'workforce': 'employee',
    # Project variants
    'wbs': 'project',         'work package': 'project',
    'initiative': 'project',  'programme': 'project',
    # Process variants
    'verify': 'validate',     'verification': 'validation',
    'check': 'review',        'checking': 'review',
    'confirm': 'approve',     'confirmation': 'approval',
    'authorise': 'approve',   'authorisation': 'approval',
    'authorize': 'approve',   'authorization': 'approval',
}


def expand_synonyms(text: str) -> str:
    """Replace words/phrases with their canonical synonyms for better matching."""
    # Phrase-level first (two-word phrases)
    for phrase, replacement in SYNONYM_MAP.items():
        if ' ' in phrase:
            text = re.sub(re.escape(phrase), replacement, text, flags=re.IGNORECASE)
    # Word-level
    words = text.lower().split()
    return ' '.join(SYNONYM_MAP.get(w, w) for w in words)


def _try_load_sentence_transformer():
    """
    Load sentence-transformers model in OFFLINE mode.

    YOUR DATA NEVER LEAVES YOUR MACHINE:
      - TRANSFORMERS_OFFLINE=1 tells the library to never make network calls
      - HF_DATASETS_OFFLINE=1 same for datasets
      - local_files_only=True  forces loading from local cache only
      - The model is loaded entirely in local memory and runs on your CPU

    First-time setup (one-time only, done separately before using the script):
      pip install sentence-transformers
            print(f"  [INFO]     python -c 'from sentence_transformers import SentenceTransformer; SentenceTransformer(\"{EMBEDDING_MODEL}\")' ")
    After that, the model is cached locally and this script runs fully offline.
    """
    try:
        from sentence_transformers import SentenceTransformer
        import os

        # Force offline mode — NO network calls ever made during inference
        os.environ['TRANSFORMERS_OFFLINE'] = '1'
        os.environ['HF_DATASETS_OFFLINE']  = '1'

        # local_files_only=True raises an error if model is not cached locally
        # rather than silently trying to download it
        model = SentenceTransformer(EMBEDDING_MODEL, local_files_only=True)
        return model

    except ImportError:
        # sentence-transformers not installed — use TF-IDF fallback
        return None
    except Exception as e:
        if 'local_files_only' in str(e) or 'No such file' in str(e) or 'not found' in str(e).lower():
            print(f"  [INFO] Model '{EMBEDDING_MODEL}' not in local cache.")
            print(f"  [INFO] Run this once to download it (requires internet, one-time only):")
            print(f"  [INFO]     python -c 'from sentence_transformers import SentenceTransformer; SentenceTransformer(\"{EMBEDDING_MODEL}\")' ")
            print(f"  [INFO] After that, the model runs fully offline. Falling back to TF-IDF for now.")
        else:
            print(f"  [WARN] Could not load embedding model: {e}")
        return None


def cluster_documents(texts: list, n_clusters: int):
    """
    Semantic Similarity Analysis — two-tier approach:

    TIER 1 (preferred): sentence-transformers embeddings
      Uses a pre-trained neural language model to encode each text into a
      dense vector that captures meaning, not just word frequency.
      "invoice posting" and "bill payment" → similar vectors (same meaning).
      Install: pip install sentence-transformers
      Model downloads automatically (~90MB) on first run, then cached.

    TIER 2 (fallback): TF-IDF + synonym expansion + cosine similarity
      If sentence-transformers is not installed, falls back to statistical
      word-frequency vectors with synonym normalisation (invoice=bill etc.).
      Good but won't catch semantic equivalences it hasn't seen.

    Both tiers use:
      - Cosine similarity matrix  (angle between vectors = meaning similarity)
      - Agglomerative clustering  (merges most-similar pairs first — far better
                                   than KMeans which assigns to nearest centroid)
      - GROUPING_HINTS injection  (nudges cluster boundaries per your instructions)
    """
    from sklearn.metrics.pairwise import cosine_similarity as cos_sim
    from sklearn.cluster import AgglomerativeClustering

    n_clusters = min(n_clusters, len(texts) - 1)
    n_clusters = max(2, n_clusters)

    # ── Tier 1: sentence-transformers ────────────────────────────────────────
    model = _try_load_sentence_transformer()

    if model is not None:
        print("  [Semantic engine: sentence-transformers / " + EMBEDDING_MODEL + "]")

        # Inject hint context into each text before encoding
        instruction_hints = [h for h in GROUPING_HINTS
                             if not any(k in h.lower() for k in
                                        ['only sap', 'only process', 'only activities',
                                         'sap code and', 'process steps and'])]
        if instruction_hints:
            hint_suffix = ' . ' + ' . '.join(instruction_hints)
            encode_texts = [t + hint_suffix for t in texts]
        else:
            encode_texts = texts

        # Encode all texts to dense embeddings
        embeddings = model.encode(
            encode_texts,
            convert_to_numpy=True,
            show_progress_bar=False,
            normalize_embeddings=True,   # unit vectors → cosine = dot product
        )

        # Cosine similarity (dot product since normalised)
        sim_matrix = np.dot(embeddings, embeddings.T)
        sim_matrix = np.clip(sim_matrix, -1, 1)

        dist_matrix = 1 - sim_matrix
        dist_matrix = np.clip(dist_matrix, 0, None)

        agg = AgglomerativeClustering(
            n_clusters=n_clusters,
            metric='precomputed',
            linkage='average',
        )
        labels = agg.fit_predict(dist_matrix)

        # For keyword extraction we still need TF-IDF
        vectorizer = TfidfVectorizer(
            stop_words='english', max_features=5000,
            ngram_range=(1, 2), min_df=1, sublinear_tf=True,
        )
        expanded = [expand_synonyms(t) for t in texts]
        tfidf_matrix = vectorizer.fit_transform(expanded)

        return labels, vectorizer, tfidf_matrix, None

    # ── Tier 2: TF-IDF + cosine fallback ─────────────────────────────────────
    print("  [Semantic engine: TF-IDF cosine similarity (install sentence-transformers for better accuracy)]")

    expanded = [expand_synonyms(t) for t in texts]

    instruction_hints = [h for h in GROUPING_HINTS
                         if not any(k in h.lower() for k in
                                    ['only sap', 'only process', 'only activities',
                                     'sap code and', 'process steps and'])]
    boosted = expanded[:]
    if instruction_hints:
        hint_block = expand_synonyms(' '.join(instruction_hints) * 3)
        hint_clean = re.sub(r'[^a-z\s]', ' ', hint_block.lower())
        boosted = [t + ' ' + hint_clean for t in expanded]

    vectorizer = TfidfVectorizer(
        stop_words='english', max_features=5000,
        ngram_range=(1, 2), min_df=1, sublinear_tf=True,
    )
    tfidf_matrix = vectorizer.fit_transform(boosted)
    sim_matrix   = cos_sim(tfidf_matrix)
    dist_matrix  = np.clip(1 - sim_matrix, 0, None)

    agg = AgglomerativeClustering(
        n_clusters=n_clusters, metric='precomputed', linkage='average',
    )
    labels = agg.fit_predict(dist_matrix)

    return labels, vectorizer, tfidf_matrix, None


# ---------------------------------------------------------------------------
# Group naming and reasoning
# ---------------------------------------------------------------------------

# Words that look like names, titles, or non-process terms — excluded from group names
NAME_EXCLUSIONS = {
    'manager', 'officer', 'director', 'head', 'lead', 'owner', 'coordinator',
    'analyst', 'specialist', 'controller', 'executive', 'assistant', 'associate',
    'team', 'staff', 'personnel', 'user', 'users', 'member', 'members',
    'john', 'jane', 'smith', 'jones',  # common name fragments
    'january', 'february', 'march', 'april', 'june', 'july', 'august',
    'september', 'october', 'november', 'december', 'monthly', 'weekly', 'daily',
    'report', 'reports', 'form', 'forms', 'email', 'document', 'documents',
    'system', 'systems', 'data', 'information', 'details', 'record', 'records',
}

# Strong process-action verbs — these anchor good group names
PROCESS_VERBS = {
    'posting', 'payment', 'approval', 'settlement', 'reconciliation', 'processing',
    'capitalisation', 'capitalization', 'depreciation', 'retirement', 'allocation',
    'budgeting', 'onboarding', 'invoicing', 'procurement', 'reporting', 'closing',
    'review', 'validation', 'verification', 'calculation', 'execution', 'creation',
    'submission', 'confirmation', 'adjustment', 'transfer', 'accrual', 'clearance',
}

# Strong process-domain nouns — give context to the verb
PROCESS_NOUNS = {
    'invoice', 'invoices', 'vendor', 'vendors', 'asset', 'assets', 'payroll',
    'salary', 'salaries', 'journal', 'journals', 'ledger', 'account', 'accounts',
    'budget', 'cost', 'costs', 'project', 'projects', 'employee', 'employees',
    'payment', 'payments', 'order', 'orders', 'purchase', 'expense', 'expenses',
    'depreciation', 'capital', 'period', 'balance', 'balances', 'bank', 'tax',
    'revenue', 'profit', 'loss', 'fixed', 'current', 'accrual', 'liability',
}


def looks_like_tcode(word: str) -> bool:
    w = word.upper()
    if re.match(r'^\d+$', w): return True
    if re.match(r'^S_ALR', w): return True
    if re.match(r'^[ZY][A-Z0-9_]{2,}$', w): return True
    if re.match(r'^[A-Z]{1,3}\d+[A-Z]?$', w): return True
    if '_' in w and len(w) > 6: return True
    return False


def is_excluded_word(word: str) -> bool:
    """Return True if word should never appear in a group name."""
    w = word.lower()
    return (
        w in STOPWORDS
        or w in NAME_EXCLUSIONS
        or looks_like_tcode(word)
        or not word.isalpha()
        or len(word) <= 3
    )


def top_cluster_keywords(km, vectorizer, cluster_id: int, n: int = 30,
                         labels: list = None, tfidf_matrix=None) -> list:
    """Return top n meaningful process keywords for a cluster.
    Uses mean of member doc vectors (works with or without KMeans).
    """
    feature_names = vectorizer.get_feature_names_out()

    if labels is not None and tfidf_matrix is not None:
        # Compute centroid as mean of all docs in this cluster
        member_indices = [i for i, l in enumerate(labels) if l == cluster_id]
        if not member_indices:
            return []
        centroid = np.asarray(
            tfidf_matrix[member_indices].mean(axis=0)
        ).flatten()
    elif km is not None:
        centroid = km.cluster_centers_[cluster_id]
    else:
        return []

    top_indices = centroid.argsort()[::-1]
    keywords = []
    for idx in top_indices:
        word = feature_names[idx]
        if not is_excluded_word(word):
            keywords.append(word)
        if len(keywords) >= n:
            break
    return keywords


def make_group_name(keywords: list, raw_texts: list, cluster_id: int) -> str:
    """
    Build a process-flow group name from cluster keywords.
    Strategy:
      1. Prefer verb-action words (posting, settlement, reconciliation...)
      2. Pair with domain noun words (invoice, asset, payroll...)
      3. Fall back to top 2-3 keywords if no process terms found
    Name format: "<Verb> & <Verb>" or "<Noun> <Verb> & <Noun> <Verb>"
    """
    if not keywords:
        return f"Process Group {cluster_id + 1}"

    kw_set = set(w.lower() for w in keywords)

    # Find process verbs and nouns present in this cluster
    found_verbs = [w for w in keywords if w.lower() in PROCESS_VERBS]
    found_nouns = [w for w in keywords if w.lower() in PROCESS_NOUNS]

    parts = []

    if found_verbs and found_nouns:
        # Pair top noun + top verb
        noun = found_nouns[0].title()
        verb = found_verbs[0].title()
        parts.append(f"{noun} {verb}")
        # Add second pair if available
        if len(found_nouns) > 1 or len(found_verbs) > 1:
            n2 = found_nouns[1].title() if len(found_nouns) > 1 else found_nouns[0].title()
            v2 = found_verbs[1].title() if len(found_verbs) > 1 else found_verbs[0].title()
            if f"{n2} {v2}" != parts[0]:
                parts.append(f"{n2} {v2}")
    elif found_verbs:
        parts = [w.title() for w in found_verbs[:3]]
    elif found_nouns:
        parts = [w.title() for w in found_nouns[:3]]
    else:
        # Fallback: top non-excluded keywords
        parts = [w.title() for w in keywords[:3]]

    return " & ".join(parts[:2]) if len(parts) >= 2 else parts[0]


def doc_top_keywords(doc_vector, vectorizer, n: int = DOC_REASON_KEYWORDS) -> list:
    feature_names = vectorizer.get_feature_names_out()
    doc_array = np.asarray(doc_vector.todense()).flatten()
    top_indices = doc_array.argsort()[::-1]
    keywords = []
    for idx in top_indices:
        if doc_array[idx] == 0:
            break
        word = feature_names[idx]
        if not is_excluded_word(word):
            keywords.append(word)
        if len(keywords) >= n:
            break
    return keywords


def make_reason(doc_keywords: list, cluster_keywords: list,
                group_name: str, sections_found: list,
                used_fallback: bool, raw_section_text: str) -> str:
    """
    Generate a process-flow focused reason sentence.
    Describes WHAT process activities link this doc to its group.
    """
    parts = []

    # Which sections were used as basis
    if sections_found:
        section_names = ' & '.join(s.title() for s in sections_found[:2])
        parts.append(f"Based on {section_names} content")
    elif used_fallback:
        parts.append("Target sections not found — based on full document content")

    # Process verbs shared with cluster (the core activity match)
    shared_verbs = [k for k in doc_keywords
                    if k.lower() in PROCESS_VERBS and k in cluster_keywords]
    shared_nouns = [k for k in doc_keywords
                    if k.lower() in PROCESS_NOUNS and k in cluster_keywords]

    if shared_verbs:
        parts.append("Shared process activities: " +
                     ', '.join(w.title() for w in shared_verbs[:3]))
    if shared_nouns:
        parts.append("Shared process subjects: " +
                     ', '.join(w.title() for w in shared_nouns[:3]))

    # Doc-specific process focus
    unique_verbs = [k for k in doc_keywords
                    if k.lower() in PROCESS_VERBS and k not in cluster_keywords]
    unique_nouns = [k for k in doc_keywords
                    if k.lower() in PROCESS_NOUNS and k not in cluster_keywords]
    unique = unique_verbs[:2] + unique_nouns[:2]
    if unique:
        parts.append("Additional focus: " + ', '.join(w.title() for w in unique[:3]))

    if len(parts) <= 1:
        # Generic fallback
        all_shared = [k for k in doc_keywords if k in cluster_keywords]
        if all_shared:
            parts.append("Common process terms: " +
                         ', '.join(w.title() for w in all_shared[:4]))

    return '. '.join(parts) + '.' if parts else f"Similar process flow to {group_name} group."
# ---------------------------------------------------------------------------
# Excel report
# ---------------------------------------------------------------------------

def build_report(groups: dict, output_path: str, target_sections: list, grouping_hints: list):
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

    tcode_font = Font(name='Courier New', size=10, color='1F4E79')

    # Banner — sections + hints
    ws.merge_cells('A1:D1')
    banner = ws['A1']
    hint_text = ("  |  Hints: " + " | ".join(grouping_hints)) if grouping_hints else ""
    banner.value = "Grouped by sections: " + " | ".join(target_sections) + hint_text
    banner.font      = Font(name='Arial', bold=True, color='1F4E79', size=10)
    banner.fill      = PatternFill('solid', start_color='DEEAF1')
    banner.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    banner.border    = med_b
    ws.row_dimensions[1].height = 22

    # Headers
    for col, h in enumerate(['Group Name', 'Document Name', 'Reason for Grouping', 'SAP Transaction Codes'], 1):
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

        for rel_path, reason, sap_codes in docs:
            c = ws.cell(row=row, column=2, value=rel_path)
            c.font = doc_font; c.fill = row_fill
            c.alignment = top_align; c.border = thin_b

            c = ws.cell(row=row, column=3, value=reason)
            c.font = rsn_font; c.fill = row_fill
            c.alignment = top_align; c.border = thin_b

            codes_str = ', '.join(sap_codes) if sap_codes else '—'
            c = ws.cell(row=row, column=4, value=codes_str)
            c.font = tcode_font; c.fill = row_fill
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
    ws.column_dimensions['C'].width = 70
    ws.column_dimensions['D'].width = 40
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
    focus_mode = parse_focus_mode(GROUPING_HINTS)
    print(f"Target sections: {TARGET_SECTIONS}")
    print(f"Focus mode     : {focus_mode}")
    if GROUPING_HINTS:
        print(f"Grouping hints : {GROUPING_HINTS}")
    print()

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

        sap_codes    = extract_sap_codes(filepath)
        focus_mode   = parse_focus_mode(GROUPING_HINTS)
        used_fallback = False

        # Apply focus mode to select what content is used for comparison
        focused_text = apply_focus_mode(focus_mode, section_text, full_text,
                                        sap_codes, filepath)

        if focus_mode == 'sap_only':
            text_for_cluster = focused_text
            print(f"  + {rel}  [mode: SAP codes only → {', '.join(sap_codes) or 'none found'}]")
        elif focus_mode in ('steps_only', 'activities_only', 'sap_and_steps'):
            text_for_cluster = focused_text
            print(f"  + {rel}  [mode: {focus_mode}]")
        elif len(section_text.split()) >= 10:
            text_for_cluster = section_text
            print(f"  + {rel}")
            print(f"      Sections found: {sections_found if sections_found else 'matched via text pattern'}")
        else:
            text_for_cluster = full_text
            used_fallback = True
            print(f"  ~ {rel} [FALLBACK: target sections not found, using full text]")

        if len(clean_text(text_for_cluster).split()) < 2:
            print(f"  [SKIP] '{rel}' has too little text for mode '{focus_mode}'.")
            continue

        rel_paths.append(rel)
        clean_texts.append(clean_text(text_for_cluster))
        meta.append({
            'rel': rel,
            'sections_found': sections_found,
            'used_fallback': used_fallback,
            'section_text': section_text,
            'sap_codes': sap_codes,
        })

    n_docs = len(rel_paths)
    if n_docs < MIN_DOCS:
        raise SystemExit(f"\nNeed at least {MIN_DOCS} readable documents. Found {n_docs}.")

    # Cluster
    n_clusters = optimal_clusters(n_docs, args.groups)
    print(f"\nClustering {n_docs} document(s) into {n_clusters} group(s)...\n")

    labels, vectorizer, tfidf_matrix, km = cluster_documents(clean_texts, n_clusters)

    # Pre-compute cluster keywords and names
    cluster_kws   = {i: top_cluster_keywords(km, vectorizer, i, labels=labels, tfidf_matrix=tfidf_matrix) for i in range(n_clusters)}
    cluster_names = {i: make_group_name(cluster_kws[i], [], i)       for i in range(n_clusters)}

    groups = {}
    for idx, label in enumerate(labels):
        m          = meta[idx]
        group_name = cluster_names[label]
        c_kws      = cluster_kws[label]
        doc_kws    = doc_top_keywords(tfidf_matrix[idx], vectorizer)
        reason     = make_reason(doc_kws, c_kws, group_name,
                                 m['sections_found'], m['used_fallback'],
                                 m['section_text'])
        groups.setdefault(group_name, []).append((m['rel'], reason, m['sap_codes']))
        print(f"  [{group_name}]  {m['rel']}")

    build_report(groups, output_file, TARGET_SECTIONS, GROUPING_HINTS)

    print(f"\nReport saved to: {output_file}")
    print(f"  {n_clusters} group(s), {n_docs} document(s) grouped.\n")
    print("Group summary:")
    for name, docs in sorted(groups.items()):
        print(f"  {name}: {len(docs)} doc(s)")


if __name__ == '__main__':
    main()
