LegalAI Document Assistant

LegalAI is a Streamlit document assistant for asking grounded questions about PDF, DOCX, TXT, and Markdown files. It supports local uploads and public Google Drive folders or files, extracts text, preserves available page metadata, creates overlapping chunks, embeds those chunks once, and retrieves relevant passages for each question.

Features

•
PDF, DOCX, TXT, and MD extraction with dedicated extraction functions.

•
Overlapping text chunking with filename, source, page, and chunk identifiers.

•
Sentence Transformers embeddings cached with Streamlit and stored in session state.

•
FAISS vector retrieval combined with important-word keyword scoring.

•
Groq responses constrained to the retrieved document context.

•
Retrieved source passages displayed below each answer.

•
Full-document preview modal from the document library and retrieved sources.

•
Public Google Drive folder and individual-file ingestion through the same pipeline.

•
Documents are processed and embedded once per session using stable file hashes.

Installation

Bash


pip install -r requirements.txt



Create .streamlit/secrets.toml:

Plain Text


GROQ_API_KEY = "your-groq-api-key"
GROQ_MODEL = "openai/gpt-oss-120b"



The API key is read from Streamlit secrets and is never hardcoded in app.py.

Run

Bash


streamlit run app.py



Upload documents from the sidebar, optionally load a public Google Drive source, then ask questions from the main workspace. Always verify important information against the original documents; the assistant is a document-analysis tool and not a substitute for legal advice.

