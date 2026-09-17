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
    :root {
        --night: #11172b;
        --night-2: #18213b;
        --panel: rgba(35, 45, 73, .72);
        --panel-solid: #202a47;
        --stroke: rgba(150, 166, 205, .25);
        --text: #f5f4f2;
        --soft: #a9b2ca;
        --gold: #e2bd67;
        --red: #b52d35;
    }

    .stApp {
        background:
            radial-gradient(circle at 8% 0%, rgba(71, 89, 145, .20), transparent 31rem),
            radial-gradient(circle at 92% 72%, rgba(145, 42, 49, .10), transparent 30rem),
            linear-gradient(145deg, var(--night) 0%, #151c33 52%, #11172a 100%);
        color: var(--text);
    }
    .main .block-container { max-width: 1160px; padding: 2.5rem 3rem 5rem; }
    [data-testid="stHeader"] { background: transparent; }
    [data-testid="stSidebar"] { background: #11172a; border-right: 1px solid rgba(145,160,195,.16); }
    [data-testid="stSidebar"] * { color: #edf0f6 !important; }
    [data-testid="stSidebar"] .stCaption { color: #8f9ab5 !important; }
    [data-testid="stFileUploader"] section { background: rgba(31,42,70,.65); border: 1px dashed #566482; border-radius: 14px; }
    [data-testid="stFileUploader"] small { color: #aeb8d0 !important; }

    .brand { padding: 4px 0 24px; }
    .brand-mark { color: var(--gold); font-size: 28px; font-weight: 400; line-height: 1; }
    .brand-name { color: #f7f7f4; font-size: 23px; font-weight: 760; letter-spacing: -.5px; }
    .brand-sub { color: #8995b2; font-size: 10px; margin-top: 6px; letter-spacing: 1.45px; }
    [data-testid="stSidebar"] hr { border-color: rgba(145,160,195,.18); }
    [data-testid="stSidebar"] h3 { color: #bec7db !important; font-size: 11px !important; letter-spacing: 1.7px; text-transform: uppercase; }

    .hero {
        position: relative; overflow: hidden; text-align: center;
        background: linear-gradient(140deg, rgba(35,45,76,.86), rgba(18,25,47,.64));
        border: 1px solid var(--stroke); border-radius: 25px; padding: 39px 34px 34px;
        margin: 0 0 21px; box-shadow: 0 22px 65px rgba(0,0,0,.23), inset 0 1px rgba(255,255,255,.05);
    }
    .hero:after { content: ""; position: absolute; width: 230px; height: 230px; border-radius: 50%; right: -90px; top: -120px; background: rgba(226,189,103,.08); filter: blur(6px); }
    .eyebrow { color: #aeb8cf; font-size: 11px; font-weight: 800; letter-spacing: 4px; text-transform: uppercase; margin-bottom: 16px; }
    .hero h1 { margin: 0; font-size: clamp(32px, 5vw, 53px); line-height: 1; font-weight: 850; letter-spacing: -2px; text-transform: uppercase; color: #f5f4f2; }
    .hero h1:first-letter { color: #e9e9e3; }
    .hero p { margin: 17px auto 0; color: #929db8; font-size: 12px; max-width: 650px; line-height: 1.7; letter-spacing: 3px; text-transform: uppercase; }
    .hero-rule { width: 94px; height: 4px; border: 0; margin: 20px auto 0; border-radius: 10px; background: linear-gradient(90deg, var(--gold), #d48b5e, var(--red)); box-shadow: 0 0 18px rgba(226,189,103,.3); }

    .metric-card, .source-card, .answer-card { background: var(--panel); border: 1px solid var(--stroke); border-radius: 17px; box-shadow: 0 14px 35px rgba(0,0,0,.13), inset 0 1px rgba(255,255,255,.035); }
    .metric-card { padding: 17px 19px; min-height: 84px; }
    .metric-title { color: #919db8; font-size: 10px; font-weight: 800; letter-spacing: 2px; text-transform: uppercase; }
    .metric-value { color: #f4f3ed; font-size: 26px; font-weight: 760; margin-top: 7px; }
    .section-label { color: #b5bfd4; font-size: 12px; font-weight: 800; letter-spacing: 3px; text-transform: uppercase; margin: 30px 0 13px; }
    .answer-card { padding: 25px 29px; margin-top: 8px; line-height: 1.75; color: #e7e9ee; }
    .answer-card p:last-child { margin-bottom: 0; }
    .source-card { padding: 17px 19px; margin: 10px 0; }
    .source-title { color: #eef0ef; font-weight: 720; font-size: 14px; }
    .source-meta { color: #929db8; font-size: 11px; margin-top: 6px; }
    .source-text { color: #ccd2df; line-height: 1.7; font-size: 14px; white-space: pre-wrap; }
    .disclaimer { background: rgba(78,64,38,.28); border: 1px solid rgba(226,189,103,.30); color: #d9c897; border-radius: 13px; padding: 13px 16px; font-size: 11px; line-height: 1.6; margin-bottom: 20px; }
    .stButton > button { border-radius: 11px; font-weight: 750; border: 1px solid var(--stroke); min-height: 42px; background: rgba(39,49,78,.75); color: #eef0f6; letter-spacing: .3px; }
    .stButton > button:hover { border-color: var(--gold); color: var(--gold); }
    [data-testid="stSidebar"] .stButton > button { background: linear-gradient(90deg, #d3a953, #e1bd70); color: #172039 !important; border: 0; }
    [data-testid="stSidebar"] .stButton > button:hover { color: #172039 !important; filter: brightness(1.08); }
    .stTextArea textarea, .stTextInput input { border-radius: 13px; border: 1px solid var(--stroke); background: rgba(26,35,61,.75); color: #f0f1f4; }
    .stTextArea textarea::placeholder, .stTextInput input::placeholder { color: #737f9c; }
    .stTextArea textarea:focus, .stTextInput input:focus { border-color: var(--gold); box-shadow: 0 0 0 1px var(--gold); }
    div[data-testid="stExpander"] { border: 1px solid var(--stroke); border-radius: 14px; background: rgba(31,41,67,.56); }
    div[data-testid="stExpander"] summary, div[data-testid="stExpander"] summary p { color: #d8ddea !important; }
    .stAlert { background: rgba(35,45,73,.72); border: 1px solid var(--stroke); color: #dce1eb; }
    [data-testid="stMetric"] { background: rgba(35,45,73,.68); border: 1px solid var(--stroke); border-radius: 14px; padding: 10px; }
    div[data-testid="stProgress"] { height: 8px; margin: 10px 0 18px; }
    div[data-testid="stProgress"] > div { background: rgba(38,48,77,.92); border-radius: 99px; }
    div[data-testid="stProgress"] > div > div { background: linear-gradient(90deg, #e0b655 0%, #edc96e 45%, #d64c43 100%); border-radius: 99px; box-shadow: 0 0 14px rgba(218,167,75,.32); }
    @media (max-width: 800px) { .main .block-container { padding: 1.25rem 1rem 4rem; } .hero { padding: 30px 20px; } .hero h1 { font-size: 34px; } }
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


def load_public_drive_folder(url):
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

            st.success(
                f"Processed {added} new document(s) "
                f"and created {chunks_added} chunks."
            )

            if skipped:
                st.info(
                    f"{skipped} document(s) "
                    "were already processed."
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

        st.rerun()


# =========================================================
# MAIN HEADER
# =========================================================
st.markdown(
    """
    <div class="hero">
        <div class="eyebrow">Private document workspace</div>
        <h1>ASK YOUR DOCUMENTS.</h1>
        <p>SEARCH • UNDERSTAND • ACT WITH CONFIDENCE</p>
        <div class="hero-rule"></div>
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

            st.markdown("**Retrieved text:**")

            # Render retrieved text as text, never as HTML. This prevents
            # document markup or model output from leaking into the page.
            st.markdown(
                f'<div class="source-text">{html.escape(source["text"])}</div>',
                unsafe_allow_html=True,
            )
