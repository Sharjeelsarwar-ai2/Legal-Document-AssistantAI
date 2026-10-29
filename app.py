import io
import html
import re
import hashlib
import tempfile
from pathlib import Path

import faiss
import gdown
import numpy as np
import streamlit as st
from docx import Document
from groq import Groq
from pypdf import PdfReader
from sentence_transformers import SentenceTransformer


# =========================================================
# PAGE CONFIG
# =========================================================
st.set_page_config(
    page_title="LegalAI - Legal Document Assistant",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded",
)


# =========================================================
# MODERN UI
# =========================================================
st.markdown(
    """
    <style>
/* LegalAI — Chambers identity (fresh for this app only)
   Parchment / deep pine / muted brass. Serif display + clean sans.
   Palette:
     parchment  #F6F1E4   paper #FFFDF7   pine ink #16281F
     pine       #1F4433   pine deep #122A1F   sage #7A9B87
     brass      #B08D3E   brass light #D9BC6A   line #DDD2B8
     slate text #3D4A42   muted #7C857D   success #2E7D4F   oxblood #9B3B2E
*/
@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,500;9..144,600;9..144,700&family=Inter:wght@400;500;600;700;800&display=swap');

:root {
    --parchment: #F6F1E4;
    --paper: #FFFDF7;
    --pine: #1F4433;
    --pine-deep: #122A1F;
    --ink: #16281F;
    --slate: #3D4A42;
    --muted: #7C857D;
    --brass: #B08D3E;
    --brass-light: #D9BC6A;
    --sage: #7A9B87;
    --line: #DDD2B8;
    --line-soft: #E7DECA;
    --success: #2E7D4F;
    --oxblood: #9B3B2E;
}

html, body { background: var(--parchment); }
.stApp {
    background:
        radial-gradient(circle at 88% -6%, rgba(176,141,62,.14), transparent 26rem),
        radial-gradient(circle at -8% 30%, rgba(31,68,51,.07), transparent 28rem),
        linear-gradient(rgba(22,40,31,.028) 1px, transparent 1px),
        var(--parchment);
    background-size: auto, auto, 100% 34px, auto;
    color: var(--ink);
    font-family: 'Inter', system-ui, -apple-system, 'Segoe UI', sans-serif;
}
.main .block-container { max-width: 1180px; padding: 1.6rem 3rem 5rem; }
[data-testid="stHeader"] { background: transparent; }
[data-testid="stAppViewContainer"] { background: transparent; }

/* ---------- Top bar ---------- */
.topbar {
    display: flex; align-items: center; justify-content: space-between; gap: 16px;
    padding: 10px 2px 16px; border-bottom: 1px solid var(--line); margin-bottom: 22px;
}
.topbar-brand { display: flex; align-items: center; gap: 11px; }
.topbar-mark {
    width: 38px; height: 38px; border-radius: 10px; display: grid; place-items: center;
    background: var(--pine-deep); color: var(--brass-light); font-family: 'Fraunces', Georgia, serif;
    font-size: 22px; font-weight: 700; box-shadow: 0 6px 16px rgba(18,42,31,.22);
}
.topbar-name { font-family: 'Fraunces', Georgia, serif; font-size: 21px; font-weight: 700; letter-spacing: -.2px; color: var(--ink); line-height: 1; }
.topbar-sub { font-size: 9px; font-weight: 700; letter-spacing: 2.2px; text-transform: uppercase; color: var(--muted); margin-top: 4px; }
.topbar-pills { display: flex; gap: 8px; flex-wrap: wrap; justify-content: flex-end; }
.topbar-pill {
    font-size: 11px; font-weight: 600; color: var(--pine); background: var(--paper);
    border: 1px solid var(--line); border-radius: 999px; padding: 7px 12px; white-space: nowrap;
}
.topbar-pill b { color: var(--brass); }

/* ---------- Sidebar ---------- */
[data-testid="stSidebar"] { background: #EDE4CB; border-right: 1px solid var(--line); }
[data-testid="stSidebar"] > div:first-child { background: #EDE4CB; }
[data-testid="stSidebar"] * { color: var(--ink); }
[data-testid="stSidebar"] .stCaption { color: var(--muted) !important; }
[data-testid="stSidebar"] hr { border-color: var(--line); }
[data-testid="stSidebar"] h3 { color: var(--pine) !important; font-size: 11px !important; letter-spacing: 1.8px; text-transform: uppercase; font-weight: 800; }
.brand { padding: 6px 0 20px; }
.brand-mark { color: var(--brass); font-family: 'Fraunces', Georgia, serif; font-size: 34px; font-weight: 700; line-height: 1; }
.brand-name { color: var(--ink); font-family: 'Fraunces', Georgia, serif; font-size: 26px; font-weight: 700; letter-spacing: -.3px; margin-top: 2px; }
.brand-sub { color: var(--muted); font-size: 9px; margin-top: 6px; letter-spacing: 1.8px; font-weight: 700; text-transform: uppercase; }
[data-testid="stFileUploader"] section {
    background: var(--paper); border: 1.5px dashed var(--brass); border-radius: 14px; padding: 10px;
}
[data-testid="stFileUploader"] small { color: var(--muted) !important; }
[data-testid="stFileUploader"] button { color: var(--pine) !important; border-color: var(--line) !important; }

/* ---------- Hero ---------- */
.hero {
    position: relative; overflow: hidden;
    background: linear-gradient(138deg, #1E4232 0%, #122A1F 58%, #0E2118 100%);
    border: 1px solid #2A4A3A; border-radius: 24px; padding: 44px 40px 38px;
    margin: 0 0 20px; box-shadow: 0 24px 55px rgba(18,42,31,.28);
    outline: 1px solid rgba(217,188,106,.28); outline-offset: -9px;
}
.hero-mark {
    position: absolute; right: 18px; top: -38px; font-family: 'Fraunces', Georgia, serif;
    font-size: 190px; line-height: 1; color: rgba(217,188,106,.10); pointer-events: none; font-weight: 700;
}
.eyebrow { color: var(--brass-light); font-size: 11px; font-weight: 800; letter-spacing: 3.6px; text-transform: uppercase; margin-bottom: 14px; position: relative; }
.hero h1 {
    margin: 0; font-family: 'Fraunces', Georgia, serif; font-weight: 600;
    font-size: clamp(36px, 5.4vw, 58px); line-height: 1.04; letter-spacing: -.5px; color: #F8F3E6; position: relative;
}
.hero h1 em { font-style: italic; color: var(--brass-light); }
.hero p { margin: 16px 0 0; color: #C9D4C8; font-size: 16px; max-width: 660px; line-height: 1.65; position: relative; }
.hero-rule { width: 84px; height: 3px; border: 0; margin: 22px 0 0; border-radius: 10px; background: linear-gradient(90deg, var(--brass-light), var(--brass)); position: relative; }
.hero-chips { display: flex; gap: 8px; flex-wrap: wrap; margin-top: 20px; position: relative; }
.hero-chip { font-size: 12px; font-weight: 600; color: #EDE6D2; background: rgba(255,253,247,.07); border: 1px solid rgba(237,230,210,.22); border-radius: 999px; padding: 7px 12px; }

/* ---------- Cards, sections, metrics ---------- */
.metric-card, .source-card, .answer-card {
    background: var(--paper); border: 1px solid var(--line); border-radius: 18px;
    box-shadow: 0 10px 26px rgba(22,40,31,.07);
}
.metric-card { padding: 18px 20px; min-height: 88px; border-top: 3px solid var(--brass); transition: transform .2s ease, box-shadow .2s ease; }
.metric-card:hover, .source-card:hover { transform: translateY(-2px); box-shadow: 0 16px 34px rgba(22,40,31,.12); }
.metric-title { color: var(--muted); font-size: 10px; font-weight: 800; letter-spacing: 2px; text-transform: uppercase; }
.metric-value { color: var(--ink); font-family: 'Fraunces', Georgia, serif; font-size: 30px; font-weight: 700; margin-top: 6px; line-height: 1; }
.metric-note { color: var(--muted); font-size: 12px; margin-top: 6px; }
.section-label {
    color: var(--pine); font-family: 'Fraunces', Georgia, serif; font-size: 21px; font-weight: 700;
    margin: 34px 0 13px; display: flex; align-items: center; gap: 12px;
}
.section-label:after { content: ""; height: 1px; flex: 1; background: linear-gradient(90deg, var(--brass), transparent); }
.answer-card { padding: 24px 28px; margin-top: 8px; line-height: 1.75; color: var(--ink); }
.answer-card p:last-child { margin-bottom: 0; }
.source-card { padding: 17px 19px; margin: 10px 0; transition: transform .2s ease, box-shadow .2s ease; }
.source-title { color: var(--ink); font-weight: 700; font-size: 14px; }
.source-meta { color: var(--muted); font-size: 11px; margin-top: 6px; }
.source-text { color: var(--slate); line-height: 1.7; font-size: 14px; white-space: pre-wrap; }
.disclaimer {
    background: #FBF3D9; border: 1px solid var(--brass-light); border-left: 4px solid var(--brass);
    color: #5C4A1E; border-radius: 13px; padding: 14px 17px; font-size: 13px; line-height: 1.6; margin-bottom: 20px;
}
.ask-lede { color: var(--slate); font-size: 14px; line-height: 1.6; margin: -4px 0 12px; }
.stButton > button {
    border-radius: 12px; font-weight: 700; border: 1px solid var(--line); min-height: 44px;
    background: var(--paper); color: var(--pine); letter-spacing: .2px; box-shadow: 0 4px 12px rgba(22,40,31,.06);
}
.stButton > button:hover { border-color: var(--brass); color: var(--pine-deep); }
.stButton > button[kind="primary"], .stButton > button[data-testid="stBaseButton-primary"] {
    background: var(--pine) !important; color: #F8F3E6 !important; border: 1px solid var(--pine) !important;
}
.stButton > button[kind="primary"]:hover, .stButton > button[data-testid="stBaseButton-primary"]:hover { background: var(--pine-deep) !important; }
[data-testid="stSidebar"] .stButton > button { background: var(--pine); color: #F8F3E6 !important; border: 0; }
[data-testid="stSidebar"] .stButton > button *, [data-testid="stSidebar"] .stButton > button p { color: #F8F3E6 !important; }
.stButton > button *, .stButton > button p { color: inherit; }
[data-testid="stSidebar"] .stButton > button:hover { background: var(--pine-deep); filter: none; }
.stTextArea textarea, .stTextInput input {
    border-radius: 14px; border: 1px solid var(--line); background: var(--paper); color: var(--ink);
    font-size: 15px; box-shadow: 0 4px 14px rgba(22,40,31,.05);
}
.stTextArea textarea::placeholder, .stTextInput input::placeholder { color: #9AA39B; }
.stTextArea textarea:focus, .stTextInput input:focus { border-color: var(--brass); box-shadow: 0 0 0 3px rgba(176,141,62,.18); }
div[data-testid="stExpander"] { border: 1px solid var(--line); border-radius: 14px; background: var(--paper); }
div[data-testid="stExpander"] summary, div[data-testid="stExpander"] summary p { color: var(--pine) !important; font-weight: 600; }
.stAlert { background: var(--paper); border: 1px solid var(--line); color: var(--ink); border-radius: 12px; }
[data-testid="stMetric"] { background: var(--paper); border: 1px solid var(--line); border-radius: 14px; padding: 12px 14px; }
[data-testid="stMetric"] label { color: var(--muted) !important; }
[data-testid="stMetric"] [data-testid="stMetricValue"] { color: var(--ink) !important; }
div[data-testid="stProgress"] { height: 8px; margin: 10px 0 18px; }
div[data-testid="stProgress"] > div { background: #E7DECA; border-radius: 99px; }
div[data-testid="stProgress"] > div > div { background: linear-gradient(90deg, var(--brass), var(--pine)); border-radius: 99px; }
div[role="dialog"] { background: var(--paper) !important; border: 1px solid var(--line); border-radius: 22px; box-shadow: 0 28px 90px rgba(22,40,31,.30); }
div[role="dialog"] [data-testid="stMarkdownContainer"] { color: var(--ink); }
.preview-kicker { color: var(--brass); font-size: 10px; font-weight: 800; letter-spacing: 2.5px; text-transform: uppercase; }
.preview-title { color: var(--ink); font-family: 'Fraunces', Georgia, serif; font-size: 25px; font-weight: 700; margin: 5px 0 3px; }
.preview-meta { color: var(--muted); font-size: 12px; }
.preview-button button { background: var(--paper) !important; color: var(--pine) !important; border-color: var(--line) !important; }
.preview-button button:hover { background: var(--pine) !important; color: #F8F3E6 !important; }

/* ---------- Basketball upload court ----------
   Visual companion to the REAL Streamlit file uploader above it.
   The arc / net / progress animations are CSS only and are only
   shown in the "shot" state, which Python renders after a real
   file has been processed. No upload is simulated. */
.court { margin: 14px 0 6px; }
.court-head { display: flex; align-items: center; justify-content: space-between; gap: 10px; margin-bottom: 8px; }
.court-title { font-size: 10px; font-weight: 800; letter-spacing: 1.8px; text-transform: uppercase; color: var(--pine); }
.court-count { font-size: 11px; font-weight: 700; color: var(--pine); background: var(--paper); border: 1px solid var(--line); border-radius: 999px; padding: 4px 10px; white-space: nowrap; }
.court-count b { color: var(--brass); }
.court-zone {
    position: relative; overflow: hidden; height: 218px;
    background: var(--paper); border: 1.5px dashed var(--brass); border-radius: 16px;
    box-shadow: 0 8px 20px rgba(22,40,31,.06);
}
.court-zone::before {
    content: ""; position: absolute; left: 0; right: 0; top: 34px; height: 1px;
    background: repeating-linear-gradient(90deg, var(--line-soft) 0 10px, transparent 10px 20px);
}
.court-drop { position: absolute; left: 0; right: 0; top: 10px; text-align: center; font-size: 12px; font-weight: 700; color: var(--slate); }
.court-drop span { display: block; font-size: 11px; font-weight: 500; color: var(--muted); margin-top: 2px; }
.court-hoop { position: absolute; right: 14px; top: 44px; width: 104px; height: 96px; transform-origin: top center; }
.court-board {
    position: absolute; left: 0; top: 0; width: 104px; height: 58px; background: #FDFBF3;
    border: 2px solid var(--pine); border-radius: 6px;
}
.court-board::after { content: ""; position: absolute; left: 50%; top: 30px; width: 34px; height: 22px; transform: translateX(-50%); border: 2px solid var(--brass); border-radius: 2px; }
.court-rim {
    position: absolute; left: 50%; top: 56px; width: 56px; height: 12px; transform: translateX(-50%);
    border: 3px solid var(--oxblood); border-radius: 50%; background: transparent;
}
.court-net {
    position: absolute; left: 50%; top: 66px; width: 48px; height: 30px; transform: translateX(-50%);
    background:
        repeating-linear-gradient(45deg, transparent 0 6px, #B9B2A0 6px 7px),
        repeating-linear-gradient(-45deg, transparent 0 6px, #B9B2A0 6px 7px);
    clip-path: polygon(0 0, 100% 0, 78% 100%, 22% 100%);
    transform-origin: top center;
}
.court-ball {
    position: absolute; left: 16px; bottom: 44px; z-index: 2;
    display: flex; align-items: center; gap: 8px; max-width: 62%;
    background: #fff; border: 1px solid var(--line); border-radius: 10px; padding: 8px 10px;
    box-shadow: 0 8px 18px rgba(22,40,31,.14); font-size: 12px; font-weight: 600; color: var(--ink);
}
.court-ball .doc-icon {
    flex: 0 0 auto; width: 26px; height: 32px; border-radius: 4px 4px 4px 0; position: relative;
    background: var(--pine); color: #F8F3E6; display: grid; place-items: center; font-size: 8px; font-weight: 800; letter-spacing: .4px;
}
.court-ball .doc-icon::after { content: ""; position: absolute; right: 0; top: 0; border-style: solid; border-width: 0 8px 8px 0; border-color: transparent var(--paper) transparent transparent; }
.court-ball .doc-name { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.court-progress { position: absolute; left: 16px; right: 16px; bottom: 26px; height: 8px; background: #E7DECA; border-radius: 99px; overflow: hidden; opacity: 0; }
.court-progress i { display: block; height: 100%; width: 0; background: var(--brass); border-radius: 99px; }
.court-done {
    position: absolute; left: 16px; bottom: 44px; z-index: 3; display: none;
    align-items: center; gap: 7px; background: #EAF3EC; border: 1px solid #BFDCC8; color: var(--success);
    font-size: 12px; font-weight: 700; border-radius: 999px; padding: 7px 12px;
}
.court-status { position: absolute; left: 16px; right: 16px; bottom: 8px; text-align: center; font-size: 11px; color: var(--muted); }
.court-hint { font-size: 11px; line-height: 1.55; color: var(--muted); margin-top: 8px; }
.court[data-state="aim"] .court-ball { outline: 2px solid rgba(176,141,62,.35); }
.court[data-state="shot"] .court-ball { animation: court-arc 1.15s cubic-bezier(.32,.72,.38,1) .15s forwards; }
.court[data-state="shot"] .court-net { animation: court-sway 1.1s ease-in-out 1.05s 1; }
.court[data-state="shot"] .court-progress { opacity: 1; }
.court[data-state="shot"] .court-progress i { animation: court-fill 1.25s ease .95s forwards, court-green .01s ease 2.2s forwards; }
.court[data-state="shot"] .court-done { display: inline-flex; animation: court-done-in .35s ease 2.2s both; }
.court[data-state="done"] .court-ball { display: none; }
.court[data-state="done"] .court-progress { opacity: 1; }
.court[data-state="done"] .court-progress i { width: 100%; background: var(--success); }
.court[data-state="done"] .court-done { display: inline-flex; }
@keyframes court-arc {
    0% { transform: translate(0, 0) rotate(0deg); opacity: 1; }
    42% { transform: translate(52px, -96px) rotate(-7deg); opacity: 1; }
    78% { transform: translate(118px, -58px) rotate(5deg) scale(.94); opacity: 1; }
    100% { transform: translate(138px, -40px) rotate(0deg) scale(.86); opacity: 0; }
}
@keyframes court-sway {
    0%, 100% { transform: translateX(-50%) rotate(0deg); }
    30% { transform: translateX(-50%) rotate(5deg) skewX(4deg); }
    60% { transform: translateX(-50%) rotate(-3deg) skewX(-3deg); }
}
@keyframes court-fill { from { width: 0; } to { width: 100%; } }
@keyframes court-green { to { background: var(--success); } }
@keyframes court-done-in { from { opacity: 0; transform: translateY(4px); } to { opacity: 1; transform: translateY(0); } }

@media (max-width: 800px) {
    .main .block-container { padding: 1.1rem 1rem 4rem; }
    .hero { padding: 30px 22px; }
    .hero h1 { font-size: 36px; }
    .hero-mark { font-size: 130px; }
    .topbar { flex-direction: column; align-items: flex-start; }
    .topbar-pills { justify-content: flex-start; }
    .court-zone { height: 206px; }
    .court-hoop { right: 8px; }
    @keyframes court-arc {
        0% { transform: translate(0, 0) rotate(0deg); opacity: 1; }
        42% { transform: translate(38px, -92px) rotate(-7deg); opacity: 1; }
        78% { transform: translate(88px, -56px) rotate(5deg) scale(.94); opacity: 1; }
        100% { transform: translate(104px, -40px) rotate(0deg) scale(.86); opacity: 0; }
    }
}
@media (prefers-reduced-motion: reduce) {
    *, *::before, *::after { animation-duration: .01ms !important; animation-delay: 0s !important; transition-duration: .01ms !important; scroll-behavior: auto !important; }
    .court[data-state="shot"] .court-ball { display: none; }
    .court[data-state="shot"] .court-progress { opacity: 1; }
    .court[data-state="shot"] .court-progress i { width: 100%; background: var(--success); }
    .court[data-state="shot"] .court-done { display: inline-flex; }
}

    </style>
    """,
    unsafe_allow_html=True,
)

# =========================================================
# SETTINGS
# =========================================================
SUPPORTED_TYPES = ["pdf", "docx", "txt", "md"]

# Simple character-based chunking.
CHUNK_SIZE = 1200
CHUNK_OVERLAP = 200

TOP_K_SEMANTIC = 8
TOP_K_KEYWORD = 8
TOP_K_FINAL = 5

SEMANTIC_WEIGHT = 0.70
KEYWORD_WEIGHT = 0.30

NOT_FOUND_MESSAGE = (
    "The requested information was not found in the provided documents."
)


# =========================================================
# SESSION STATE
# =========================================================
def initialize_state():
    defaults = {
        "documents": [],
        "chunks": [],
        "embeddings": None,
        "faiss_index": None,
        "processed_ids": set(),
        "last_question": "",
        "last_answer": "",
        "last_sources": [],
        "drive_loaded_link": "",
        "basketball_shot": None,
        "basketball_seen": False,
    }

    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


initialize_state()


# =========================================================
# CACHED EMBEDDING MODEL
# =========================================================
@st.cache_resource(show_spinner="Loading Sentence Transformer model...")
def load_embedding_model():
    return SentenceTransformer("all-MiniLM-L6-v2")


# =========================================================
# FILE ID
# =========================================================
def file_hash(file_bytes, filename, source):
    data = (
        source.encode("utf-8")
        + b"|"
        + filename.encode("utf-8")
        + b"|"
        + file_bytes
    )
    return hashlib.sha256(data).hexdigest()


# =========================================================
# DOCUMENT EXTRACTION
# =========================================================
def extract_pdf(file_bytes, filename, source):
    pages = []
    reader = PdfReader(io.BytesIO(file_bytes))

    for page_number, page in enumerate(reader.pages, start=1):
        text = page.extract_text() or ""
        text = text.strip()

        if text:
            pages.append(
                {
                    "filename": filename,
                    "page": page_number,
                    "source": source,
                    "text": text,
                }
            )

    return pages


def extract_docx(file_bytes, filename, source):
    document = Document(io.BytesIO(file_bytes))

    paragraphs = [
        paragraph.text.strip()
        for paragraph in document.paragraphs
        if paragraph.text.strip()
    ]

    text = "\n".join(paragraphs)

    if not text:
        return []

    return [
        {
            "filename": filename,
            "page": None,
            "source": source,
            "text": text,
        }
    ]


def extract_txt(file_bytes, filename, source):
    text = file_bytes.decode("utf-8", errors="ignore").strip()

    if not text:
        return []

    return [
        {
            "filename": filename,
            "page": None,
            "source": source,
            "text": text,
        }
    ]


def extract_md(file_bytes, filename, source):
    text = file_bytes.decode("utf-8", errors="ignore").strip()

    if not text:
        return []

    return [
        {
            "filename": filename,
            "page": None,
            "source": source,
            "text": text,
        }
    ]


def extract_document(file_bytes, filename, source):
    extension = Path(filename).suffix.lower()

    if extension == ".pdf":
        return extract_pdf(file_bytes, filename, source)

    if extension == ".docx":
        return extract_docx(file_bytes, filename, source)

    if extension == ".txt":
        return extract_txt(file_bytes, filename, source)

    if extension == ".md":
        return extract_md(file_bytes, filename, source)

    return []


# =========================================================
# TEXT CHUNKING
# =========================================================
def chunk_text(text, chunk_size=CHUNK_SIZE, overlap=CHUNK_OVERLAP):
    text = re.sub(r"\s+", " ", text).strip()

    if not text:
        return []

    chunks = []
    start = 0

    while start < len(text):
        end = min(start + chunk_size, len(text))
        chunk = text[start:end].strip()

        if chunk:
            chunks.append(chunk)

        if end >= len(text):
            break

        start = max(end - overlap, start + 1)

    return chunks


def create_chunks(extracted_pages, filename, source):
    chunks = []

    for page_data in extracted_pages:
        page_chunks = chunk_text(page_data["text"])

        for index, chunk in enumerate(page_chunks):
            chunks.append(
                {
                    "filename": filename,
                    "page": page_data["page"],
                    "source": source,
                    "text": chunk,
                    "chunk_id": (
                        f"{filename}-{page_data['page']}-{index}"
                    ),
                }
            )

    return chunks


# =========================================================
# EMBEDDINGS + FAISS
# =========================================================
def create_embeddings(texts):
    model = load_embedding_model()

    embeddings = model.encode(
        texts,
        convert_to_numpy=True,
        normalize_embeddings=True,
        show_progress_bar=False,
    )

    return np.asarray(embeddings, dtype="float32")


def rebuild_search_index():
    if not st.session_state.chunks:
        st.session_state.embeddings = None
        st.session_state.faiss_index = None
        return

    texts = [chunk["text"] for chunk in st.session_state.chunks]
    embeddings = create_embeddings(texts)

    dimension = embeddings.shape[1]

    # Inner product works as cosine similarity because
    # Sentence Transformer embeddings are normalized.
    index = faiss.IndexFlatIP(dimension)
    index.add(embeddings)

    st.session_state.embeddings = embeddings
    st.session_state.faiss_index = index


# =========================================================
# KEYWORD SEARCH
# =========================================================
STOP_WORDS = {
    "the", "a", "an", "is", "are", "was", "were",
    "what", "which", "who", "when", "where", "how",
    "why", "does", "do", "did", "of", "to", "in",
    "on", "for", "and", "or", "with", "this", "that",
    "these", "those", "about", "from", "by", "be",
    "as", "it", "its", "their", "there", "under",
    "according", "can", "could", "would", "should",
}


def important_words(text):
    words = re.findall(
        r"[A-Za-z0-9][A-Za-z0-9_-]{1,}",
        text.lower(),
    )

    return [
        word
        for word in words
        if word not in STOP_WORDS
    ]


def keyword_scores(question):
    query_words = set(important_words(question))

    scores = np.zeros(
        len(st.session_state.chunks),
        dtype="float32",
    )

    if not query_words:
        return scores

    for index, chunk in enumerate(st.session_state.chunks):
        chunk_words = set(
            important_words(chunk["text"])
        )

        matches = len(
            query_words.intersection(chunk_words)
        )

        scores[index] = matches / len(query_words)

    return scores


# =========================================================
# HYBRID SEARCH
# =========================================================
def hybrid_search(question, top_k=TOP_K_FINAL):
    if st.session_state.faiss_index is None:
        return []

    model = load_embedding_model()

    question_embedding = model.encode(
        [question],
        convert_to_numpy=True,
        normalize_embeddings=True,
    ).astype("float32")

    semantic_k = min(
        TOP_K_SEMANTIC,
        len(st.session_state.chunks),
    )

    semantic_values, semantic_indices = (
        st.session_state.faiss_index.search(
            question_embedding,
            semantic_k,
        )
    )

    semantic_scores = np.zeros(
        len(st.session_state.chunks),
        dtype="float32",
    )

    for score, index in zip(
        semantic_values[0],
        semantic_indices[0],
    ):
        if index >= 0:
            semantic_scores[index] = float(score)

    keyword_score_array = keyword_scores(question)

    # Normalize semantic scores to 0-1.
    min_score = semantic_scores.min()
    max_score = semantic_scores.max()

    if max_score > min_score:
        semantic_normalized = (
            (semantic_scores - min_score)
            / (max_score - min_score)
        )
    else:
        semantic_normalized = semantic_scores

    hybrid_scores = (
        SEMANTIC_WEIGHT * semantic_normalized
        + KEYWORD_WEIGHT * keyword_score_array
    )

    candidate_count = min(
        max(TOP_K_SEMANTIC, TOP_K_KEYWORD),
        len(st.session_state.chunks),
    )

    candidate_indices = np.argsort(
        hybrid_scores
    )[::-1][:candidate_count]

    results = []

    for index in candidate_indices:
        if hybrid_scores[index] <= 0:
            continue

        result = dict(
            st.session_state.chunks[index]
        )

        result["semantic_score"] = float(
            semantic_normalized[index]
        )
        result["keyword_score"] = float(
            keyword_score_array[index]
        )
        result["hybrid_score"] = float(
            hybrid_scores[index]
        )

        results.append(result)

    return results[:top_k]


# =========================================================
# GROQ
# =========================================================
def get_groq_client():
    api_key = st.secrets.get("GROQ_API_KEY")

    if not api_key:
        raise ValueError(
            "GROQ_API_KEY is missing. "
            "Add it to .streamlit/secrets.toml."
        )

    return Groq(api_key=api_key)


def ask_groq(question, retrieved_chunks):
    context_parts = []

    for number, chunk in enumerate(
        retrieved_chunks,
        start=1,
    ):
        page = (
            chunk["page"]
            if chunk["page"] is not None
            else "Not available"
        )

        context_parts.append(
            f"""SOURCE {number}
Filename: {chunk["filename"]}
Page: {page}
Source: {chunk["source"]}

Document text:
{chunk["text"]}
"""
        )

    context = "\n\n".join(context_parts)

    system_prompt = """You are a legal document analysis assistant.

Answer the user's question using ONLY the document context supplied
in the user message.

Rules:
1. Do not use outside knowledge.
2. Do not invent or assume clauses, facts, dates, parties,
   obligations, penalties, deadlines, section numbers, or
   interpretations.
3. If the answer is not supported by the supplied context, say:
   "The requested information was not found in the provided documents."
4. When useful, mention the relevant filename and page.
5. If the supplied documents contain conflicting information,
   clearly identify the conflict.
6. Keep the answer clear and reasonably concise.
7. This is document analysis, not legal advice.
8. Never present your response as a substitute for a qualified lawyer.
"""

    user_prompt = f"""DOCUMENT CONTEXT:

{context}

USER QUESTION:
{question}

Answer only from the document context above."""

    client = get_groq_client()

    model_name = st.secrets.get(
        "GROQ_MODEL",
        "openai/gpt-oss-120b",
    )

    response = client.chat.completions.create(
        model=model_name,
        messages=[
            {
                "role": "system",
                "content": system_prompt,
            },
            {
                "role": "user",
                "content": user_prompt,
            },
        ],
        temperature=0,
    )

    return response.choices[0].message.content.strip()


# =========================================================
# GOOGLE DRIVE PUBLIC FOLDER
# =========================================================
def extract_drive_folder_id(url):
    match = re.search(
        r"/folders/([a-zA-Z0-9_-]+)",
        url,
    )

    if match:
        return match.group(1)

    # Also allow a folder ID directly.
    if re.fullmatch(
        r"[a-zA-Z0-9_-]{10,}",
        url.strip(),
    ):
        return url.strip()

    return None


def extract_drive_file_id(url):
    match = re.search(
        r"/file/d/([a-zA-Z0-9_-]+)",
        url,
    )

    if match:
        return match.group(1)

    return None


def load_public_drive_folder(url):
    file_id = extract_drive_file_id(url)
    if file_id:
        temp_dir = Path(tempfile.mkdtemp(prefix="legalai_drive_file_"))
        downloaded_path = gdown.download(
            id=file_id,
            output=str(temp_dir / "drive_document"),
            quiet=True,
            fuzzy=True,
        )
        if not downloaded_path:
            return []

        path = Path(downloaded_path)
        extension = path.suffix.lower().replace(".", "")
        if extension not in SUPPORTED_TYPES:
            return []

        return [
            {
                "filename": path.name,
                "bytes": path.read_bytes(),
                "source": "Google Drive",
            }
        ]

    folder_id = extract_drive_folder_id(url)

    if not folder_id:
        raise ValueError(
            "Invalid Google Drive folder link. "
            "Please paste a public Google Drive folder URL."
        )

    temp_dir = tempfile.mkdtemp(
        prefix="legalai_drive_"
    )

    downloaded_files = gdown.download_folder(
        id=folder_id,
        output=temp_dir,
        quiet=True,
        use_cookies=False,
    )

    if not downloaded_files:
        return []

    files = []

    for item in downloaded_files:
        path = Path(item)

        if not path.is_file():
            continue

        extension = path.suffix.lower().replace(".", "")

        if extension not in SUPPORTED_TYPES:
            continue

        try:
            files.append(
                {
                    "filename": path.name,
                    "bytes": path.read_bytes(),
                    "source": "Google Drive",
                }
            )
        except OSError:
            continue

    return files


# =========================================================
# PROCESS DOCUMENTS
# =========================================================
def process_document(
    file_bytes,
    filename,
    source,
    document_id,
):
    extracted = extract_document(
        file_bytes,
        filename,
        source,
    )

    if not extracted:
        return None, []

    chunks = create_chunks(
        extracted,
        filename,
        source,
    )

    total_characters = sum(
        len(page["text"])
        for page in extracted
    )

    extension = Path(filename).suffix.lower()

    document_info = {
        "filename": filename,
        "type": extension.replace(".", "").upper(),
        "pages": (
            len(extracted)
            if extension == ".pdf"
            else None
        ),
        "characters": total_characters,
        "source": source,
        "status": "Processed",
        "full_text": "\n\n".join(page["text"] for page in extracted),
        "id": document_id,
    }

    return document_info, chunks


def add_documents(file_items):
    new_documents = []
    new_chunks = []
    skipped = 0

    for item in file_items:
        filename = item["filename"]
        file_bytes = item["bytes"]
        source = item.get(
            "source",
            "Local Upload",
        )

        document_id = file_hash(
            file_bytes,
            filename,
            source,
        )

        if document_id in st.session_state.processed_ids:
            skipped += 1
            continue

        document_info, chunks = process_document(
            file_bytes,
            filename,
            source,
            document_id,
        )

        if document_info is None:
            continue

        new_documents.append(document_info)
        new_chunks.extend(chunks)

        st.session_state.processed_ids.add(
            document_id
        )

    if new_documents:
        st.session_state.documents.extend(
            new_documents
        )

        st.session_state.chunks.extend(
            new_chunks
        )

        # Embeddings and FAISS are rebuilt only when
        # new documents are actually added.
        rebuild_search_index()

    return (
        len(new_documents),
        len(new_chunks),
        skipped,
    )


# =========================================================
# DOCUMENT PREVIEW DIALOG
# =========================================================
@st.dialog("Document preview", width="large")
def show_document_preview(document):
    st.markdown('<div class="preview-kicker">Full document preview</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="preview-title">{html.escape(document["filename"])}</div>', unsafe_allow_html=True)
    st.markdown(
        f'<div class="preview-meta">{document["type"]}  ·  {document["characters"]:,} characters  ·  {html.escape(document.get("source", "Local Upload"))}</div>',
        unsafe_allow_html=True,
    )
    st.markdown("<br>", unsafe_allow_html=True)
    st.text_area(
        "Document text",
        value=document.get("full_text", "No extracted text is available."),
        height=560,
        disabled=True,
        label_visibility="collapsed",
    )


def find_document(filename, source):
    for document in st.session_state.documents:
        if document["filename"] == filename and document.get("source") == source:
            if document.get("full_text"):
                return document

            # Documents already present in a live session may predate the
            # full_text field. Reconstruct a readable preview from chunks.
            legacy_document = dict(document)
            matching_chunks = [
                chunk["text"]
                for chunk in st.session_state.chunks
                if chunk["filename"] == filename
                and chunk.get("source") == source
            ]
            legacy_document["full_text"] = "\n\n".join(matching_chunks)
            return legacy_document
    return None


# =========================================================
# SIDEBAR
# =========================================================
with st.sidebar:
    st.markdown(
        """
        <div class="brand">
            <div class="brand-mark">§</div>
            <div class="brand-name">LegalAI</div>
            <div class="brand-sub">DOCUMENT INTELLIGENCE WORKSPACE</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("---")

    # -------------------------
    # Local PC uploads
    # -------------------------
    st.markdown("### UPLOAD DOCUMENTS")

    uploaded_files = st.file_uploader(
        "Choose legal documents",
        type=SUPPORTED_TYPES,
        accept_multiple_files=True,
        label_visibility="collapsed",
    )


    if st.button(
        "Process PC Documents",
        use_container_width=True,
    ):
        if not uploaded_files:
            st.warning(
                "Please select at least one document."
            )
        else:
            file_items = [
                {
                    "filename": file.name,
                    "bytes": file.getvalue(),
                    "source": "Local Upload",
                }
                for file in uploaded_files
            ]

            with st.spinner(
                "Extracting, chunking and embedding..."
            ):
                added, chunks_added, skipped = (
                    add_documents(file_items)
                )

            if added:
                st.session_state["basketball_shot"] = {
                    "names": [f["filename"] for f in file_items],
                    "filename": file_items[0]["filename"],
                }
                st.session_state["basketball_seen"] = False

            st.success(
                f"Processed {added} new document(s) "
                f"and created {chunks_added} chunks."
            )

            if skipped:
                st.info(
                    f"{skipped} document(s) "
                    "were already processed."
                )

    # -------------------------
    # Basketball upload court (visual companion to the real uploader above)
    # The animation only ever reflects a real selection / real processing:
    # "aim" = files selected, not yet processed; "shot"/"done" = files
    # actually processed by the Process button below in this session.
    # It never simulates an upload.
    _uploaded_names = [f.name for f in uploaded_files] if uploaded_files else []
    _shot = st.session_state.get("basketball_shot")
    _shot_names = _shot["names"] if isinstance(_shot, dict) else []
    if _uploaded_names and _shot_names and _uploaded_names == _shot_names:
        _court_state = "shot" if not st.session_state.get("basketball_seen") else "done"
        if _court_state == "shot":
            st.session_state["basketball_seen"] = True
        _court_file = _shot.get("filename", _uploaded_names[0]) if isinstance(_shot, dict) else _uploaded_names[0]
        _court_status = "Shot taken — file processed."
    elif _uploaded_names:
        _court_state = "aim"
        if _shot is not None:
            st.session_state["basketball_shot"] = None
            st.session_state["basketball_seen"] = False
        _court_file = _uploaded_names[0]
        _court_status = "Ready — press Process PC Documents to shoot."
    else:
        _court_state = "idle"
        _court_file = ""
        _court_status = "Select a document above to take the shot."
    _court_ext = (
        Path(_court_file).suffix.lstrip(".").upper()[:4]
        if _court_file
        else "DOC"
    )
    _court_total = len(st.session_state.documents)
    st.markdown(
        f"""
        <div class="court" data-state="{_court_state}">
            <div class="court-head">
                <span class="court-title">Upload court</span>
                <span class="court-count">Uploaded <b>{_court_total}</b></span>
            </div>
            <div class="court-zone">
                <div class="court-drop">Drop files here<span>or take the shot</span></div>
                <div class="court-hoop" aria-hidden="true">
                    <div class="court-board"></div>
                    <div class="court-rim"></div>
                    <div class="court-net"></div>
                </div>
                <div class="court-ball" aria-hidden="true">
                    <span class="doc-icon">{html.escape(_court_ext)}</span>
                    <span class="doc-name">{html.escape(_court_file) if _court_file else "document"}</span>
                </div>
                <div class="court-progress" aria-hidden="true"><i></i></div>
                <div class="court-done">✓ Done — file processed</div>
                <div class="court-status">{html.escape(_court_status)}</div>
            </div>
            <div class="court-hint">This animation follows your real selection above. Files are only processed when you press Process PC Documents — nothing here is a simulated upload.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("---")

    # -------------------------
    # Google Drive
    # -------------------------
    st.markdown("### GOOGLE DRIVE")

    drive_link = st.text_input(
        "Public Drive folder link",
        placeholder=(
            "Paste public Google Drive folder link"
        ),
        label_visibility="collapsed",
    )

    st.caption(
        "The folder must be publicly accessible."
    )

    if st.button(
        "Load Drive Folder",
        use_container_width=True,
    ):
        if not drive_link.strip():
            st.warning(
                "Please paste a Google Drive folder link."
            )
        else:
            try:
                with st.spinner(
                    "Loading files from Google Drive..."
                ):
                    drive_files = (
                        load_public_drive_folder(
                            drive_link.strip()
                        )
                    )

                if not drive_files:
                    st.warning(
                        "No supported PDF, DOCX, TXT, "
                        "or MD files were found."
                    )
                else:
                    with st.spinner(
                        "Processing Drive documents..."
                    ):
                        added, chunks_added, skipped = (
                            add_documents(drive_files)
                        )

                    st.session_state.drive_loaded_link = (
                        drive_link.strip()
                    )

                    st.success(
                        f"Loaded {added} new document(s) "
                        f"and created {chunks_added} chunks."
                    )

                    if skipped:
                        st.info(
                            f"{skipped} document(s) "
                            "were already processed."
                        )

            except Exception as error:
                st.error(
                    f"Could not load the Drive folder: {error}"
                )

    st.markdown("---")

    # -------------------------
    # Knowledge base
    # -------------------------
    st.markdown("### DOCUMENT LIBRARY")

    st.metric(
        "Documents",
        len(st.session_state.documents),
    )

    st.metric(
        "Chunks",
        len(st.session_state.chunks),
    )

    if st.button(
        "Clear All Documents",
        use_container_width=True,
    ):
        st.session_state.documents = []
        st.session_state.chunks = []
        st.session_state.embeddings = None
        st.session_state.faiss_index = None
        st.session_state.processed_ids = set()
        st.session_state.last_question = ""
        st.session_state.last_answer = ""
        st.session_state.last_sources = []
        st.session_state.drive_loaded_link = ""
        st.session_state.basketball_shot = None
        st.session_state.basketball_seen = False

        st.rerun()


# =========================================================
# MAIN HEADER
# =========================================================
st.markdown(
    """
    <div class="topbar">
        <div class="topbar-brand">
            <div class="topbar-mark">§</div>
            <div>
                <div class="topbar-name">LegalAI</div>
                <div class="topbar-sub">Document Intelligence Workspace</div>
            </div>
        </div>
        <div class="topbar-pills">
            <span class="topbar-pill"><b>§</b> Grounded RAG</span>
            <span class="topbar-pill">Hybrid Search</span>
            <span class="topbar-pill">Private Workspace</span>
        </div>
    </div>
    <div class="hero">
        <div class="hero-mark" aria-hidden="true">§</div>
        <div class="eyebrow">Private document workspace</div>
        <h1>Ask your<br>documents<em>.</em></h1>
        <p>Upload a legal document, then ask a question. Every answer is drawn only from the documents you provide, with the supporting passages shown underneath.</p>
        <div class="hero-rule"></div>
        <div class="hero-chips">
            <span class="hero-chip">PDF · DOCX · TXT · MD</span>
            <span class="hero-chip">Page-aware sources</span>
            <span class="hero-chip">Google Drive folders</span>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="disclaimer">
        <strong>Important:</strong>
        LegalAI is a document analysis and information retrieval
        tool. It does not provide legal advice or replace a
        qualified legal professional. Always verify important
        information against the original documents.
    </div>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# METRICS
# =========================================================
col1, col2, col3 = st.columns(3)

with col1:
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-title">Documents</div>
            <div class="metric-value">
                {len(st.session_state.documents)}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with col2:
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-title">Searchable Chunks</div>
            <div class="metric-value">
                {len(st.session_state.chunks)}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with col3:
    st.markdown(
        """
        <div class="metric-card">
            <div class="metric-title">Workspace</div>
            <div class="metric-value">Ready</div>
            <div class="metric-note">Documents are processed once per session.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


st.markdown("")


# =========================================================
# LOADED DOCUMENTS
# =========================================================
if st.session_state.documents:
    st.markdown('<div class="section-label">Your document library</div>', unsafe_allow_html=True)

    for document in st.session_state.documents:
        page_text = (
            f"{document['pages']} pages"
            if document["pages"] is not None
            else "Page count not available"
        )
        source_label = document.get("source", "Local Upload")
        drive_document = source_label == "Google Drive"

        # Native Streamlit elements are used here deliberately. They prevent
        # metadata such as <div> tags from ever being displayed as document text.
        with st.container(border=True):
            st.markdown(f"**▪ {document['filename']}**")
            st.caption(
                f"{document['type']}  •  {page_text}  •  "
                f"{document['characters']:,} characters  •  {source_label}"
            )
            st.caption("✓ Processed")
            preview_col, status_col = st.columns([1, 2])
            with preview_col:
                st.markdown('<div class="preview-button">', unsafe_allow_html=True)
                if st.button("VIEW DOCUMENT", key=f"library_preview_{document['id']}", use_container_width=True):
                    preview_document = find_document(
                        document["filename"],
                        document.get("source", "Local Upload"),
                    ) or document
                    show_document_preview(preview_document)
                st.markdown('</div>', unsafe_allow_html=True)
            with status_col:
                if drive_document:
                    st.progress(1.0, text="DOCUMENT READY  ·  100%")

else:
    st.info(
        "Upload legal documents from your PC or load "
        "a public Google Drive folder to begin."
    )


# =========================================================
# QUESTION AREA
# =========================================================
st.markdown('<div class="section-label">Ask your documents</div>', unsafe_allow_html=True)
st.markdown('<p class="ask-lede">Ask in plain language. The answer will use only the retrieved passages from your processed documents.</p>', unsafe_allow_html=True)

question = st.text_area(
    "Ask a question",
    placeholder=(
        "Example: What are the termination conditions "
        "in the agreement?"
    ),
    height=100,
)

ask_button = st.button(
    "ASK THE DOCUMENTS",
    type="primary",
    use_container_width=True,
)


# =========================================================
# ANSWER GENERATION
# =========================================================
if ask_button:
    if not question.strip():
        st.warning("Please enter a question.")

    elif not st.session_state.chunks:
        st.warning(
            "Please process at least one document first."
        )

    else:
        with st.spinner(
            "Searching documents and generating answer..."
        ):
            try:
                retrieved = hybrid_search(
                    question.strip(),
                    TOP_K_FINAL,
                )

                if not retrieved:
                    answer = NOT_FOUND_MESSAGE
                else:
                    answer = ask_groq(
                        question.strip(),
                        retrieved,
                    )

                st.session_state.last_question = (
                    question.strip()
                )

                st.session_state.last_answer = answer

                st.session_state.last_sources = retrieved

            except Exception as error:
                st.error(
                    f"Something went wrong: {error}"
                )


# =========================================================
# DISPLAY ANSWER
# =========================================================
if st.session_state.last_answer:
    st.markdown('<div class="section-label">Analysis</div>', unsafe_allow_html=True)

    with st.container(border=True):
        st.markdown(st.session_state.last_answer)

    st.markdown('<div class="section-label">Supporting passages</div>', unsafe_allow_html=True)

    if not st.session_state.last_sources:
        st.info(
            "No document sources were retrieved."
        )

    for number, source in enumerate(
        st.session_state.last_sources,
        start=1,
    ):
        page = (
            source["page"]
            if source["page"] is not None
            else "Not available"
        )

        with st.expander(
            f"Source {number} — {source['filename']}"
        ):
            st.markdown(
                f"""
                **Document:** {source["filename"]}

                **Page:** {page}

                **Source:** {source["source"]}

                **Relevance:** {source["hybrid_score"]:.3f}
                """
            )

            source_document = find_document(source["filename"], source["source"])
            if source_document is not None:
                st.markdown('<div class="preview-button">', unsafe_allow_html=True)
                if st.button("VIEW FULL DOCUMENT", key=f"source_preview_{number}_{source['filename']}", use_container_width=True):
                    show_document_preview(source_document)
                st.markdown('</div>', unsafe_allow_html=True)

            st.markdown("**Retrieved passage:**")

            # Render retrieved text as text, never as HTML. This prevents
            # document markup or model output from leaking into the page.
            st.markdown(
                f'<div class="source-text">{html.escape(source["text"])}</div>',
                unsafe_allow_html=True,
            )
