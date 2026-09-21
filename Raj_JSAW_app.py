"""
Raj JSAW — Streamlit App (JOB-e)
Single-file app, matching TimberLens's own structure. JOB-e is the sole
conversational interface — no per-agent tabs/chats, per explicit design
choice made earlier in this project. Search/Talent/Company Research/Resume
agents are tools JOB-e calls; a dedicated Draft Review panel handles
anything Resume Agent produces (edit / approve & deliver / discard), since
that's a UI action, not something decided through chat once a real UI
exists to do it properly.

Run with: streamlit run Raj_JSAW_app.py
"""

import streamlit as st
import os
import re
import json
import copy
import hashlib
import imaplib
import email as email_lib
import email.header
import smtplib
import asyncio
import tempfile
from datetime import datetime, timedelta
from pathlib import Path
from email.message import EmailMessage
from typing import List, Dict, Any

import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.filters import AutoFilter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableColumn, TableStyleInfo
from dotenv import load_dotenv
from anthropic import Anthropic

import nest_asyncio
nest_asyncio.apply()

st.set_page_config(
    page_title="JOB-e — Raj JSAW Command Center",
    page_icon="💼",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── HEADER — matches TimberLens's real header pattern exactly ──────────────
st.markdown(
    """
    <div style="text-align: center; margin-top: -15px; margin-bottom: 25px;">
        <h1 style='margin: 0; font-size: 40px; padding: 0; font-family: "Playfair Display", serif !important; color: #1E3A8A !important;'>
            JOB-e — Raj JSAW Command Center
        </h1>
        <h6 style='color: #8a7560; font-weight: 500; margin: 0; padding-top: 8px; font-family: "DM Sans", sans-serif; letter-spacing: 0.03em;'>
            Job Search & Application Workflow &middot; Richardson, TX &middot; Multi-Agent System
        </h6>
    </div>
    """,
    unsafe_allow_html=True,
)

# ── THEME — full stylesheet ported directly from TimberLens's real CSS ─────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@500;700&family=DM+Sans:wght@300;400;500&display=swap');

html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }

.stApp {
    background-color: #FFFFFF;
    color: #1E293B;
}

section[data-testid="stSidebar"] {
    background-color: #F1F5F9;
    border-right: 1px solid #CBD5E1;
}
section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] span,
section[data-testid="stSidebar"] label {
    color: #334155 !important;
}

h1, h2, h3 {
    font-family: 'Playfair Display', serif !important;
    color: #1E3A8A !important;
}

.kpi-card {
    background: #E0F2FE;
    border: 1px solid #BAE6FD;
    border-radius: 10px;
    padding: 18px 14px;
    text-align: center;
    margin-bottom: 8px;
}
.kpi-val {
    font-family: 'Playfair Display', serif;
    font-size: 26px;
    color: #0369A1;
    font-weight: 700;
}
.kpi-lbl {
    font-size: 11px;
    color: #0284C7;
    text-transform: uppercase;
    letter-spacing: .1em;
    margin-top: 4px;
}

.status-line {
    font-size: 12px;
    color: #475569;
    font-family: 'DM Sans', monospace;
    padding: 4px 0;
}

div[data-testid="stTextArea"] textarea,
div[data-testid="stTextInput"] input {
    background-color: #FFFFFF !important;
    color: #1E293B !important;
    border: 1px solid #CBD5E1 !important;
    border-radius: 8px !important;
}

.stButton > button {
    background: linear-gradient(135deg, #2563EB, #1D4ED8) !important;
    color: #FFFFFF !important;
    border: none !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    font-family: 'DM Sans', sans-serif !important;
    letter-spacing: .04em !important;
}
.stButton > button:hover {
    box-shadow: 0 4px 14px rgba(37, 99, 235, 0.35) !important;
}

.stSelectbox > div > div,
.stMultiSelect > div > div {
    background-color: #FFFFFF !important;
    border-color: #CBD5E1 !important;
    color: #1E293B !important;
}

.stTabs [data-baseweb="tab"] { color: #64748B !important; font-size: 18px !important; font-weight: 600 !important; }
.stTabs [aria-selected="true"] { color: #2563EB !important; border-bottom-color: #2563EB !important; }
[data-testid="stTabs"] [data-baseweb="tab-list"] { display: flex !important; justify-content: space-between !important; width: 100% !important; }
[data-testid="stTabs"] [data-baseweb="tab"] { flex: 1 !important; text-align: center !important; }

.stDataFrame { background-color: #FFFFFF !important; }

.report-box {
    background: #F0F9FF;
    border-left: 3px solid #0284C7;
    border-radius: 0 10px 10px 0;
    padding: 18px 22px;
    margin: 10px 0;
    color: #0F172A;
}
.user-box {
    background: #F0FDF4;
    border-left: 3px solid #16A34A;
    border-radius: 0 10px 10px 0;
    padding: 12px 18px;
    margin: 8px 0;
    color: #14532D;
}

.status-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 6px; margin-top: 8px; margin-bottom: 8px; }
.status-cell {
    padding: 6px; font-size: 12px !important; font-weight: 500; border-radius: 4px;
    text-align: center; display: flex; align-items: center; justify-content: center;
    border: 1px solid #e0e0e0;
}
.status-pass { background-color: #e6f4ea !important; color: #137333 !important; border-color: #ceead6 !important; }
.status-fail { background-color: #fce8e6 !important; color: #c5221f !important; border-color: #fad2cf !important; }
</style>
""", unsafe_allow_html=True)

# =============================================================================
# CONFIG — ported directly from the notebook's config cell, unchanged logic
# =============================================================================

# ── LOCAL CONFIG — editable via the sidebar's Setup & Paths panel, matching
# TimberLens's own load_config()/save_config() pattern. Saved to the user's
# home folder, not this project folder, so it survives even if the project
# folder itself gets moved/renamed.
_CONFIG_PATH = Path.home() / ".raj_jsaw_config.json"


def load_app_config():
    if _CONFIG_PATH.exists():
        try:
            return json.loads(_CONFIG_PATH.read_text())
        except (json.JSONDecodeError, OSError):
            return {}
    return {}


def save_app_config(cfg: dict):
    _CONFIG_PATH.write_text(json.dumps(cfg, indent=2))


_app_cfg = load_app_config()

BASE_DIR = Path(_app_cfg.get("base_dir", "/Users/rajhomedesktop/Desktop/Raj JSAW"))

CONFIG_DIR = BASE_DIR / "config"
CONTEXT_HUB_DIR = Path(_app_cfg.get("context_hub_dir") or (BASE_DIR / "context_hub"))
STATE_DIR = BASE_DIR / "state"
DRAFTS_DIR = STATE_DIR / "drafts"
SKILLS_DIR = BASE_DIR / "skills"
LOGS_DIR = BASE_DIR / "logs"
JOB_ALERTS_DIR = Path(_app_cfg.get("job_alerts_dir") or (BASE_DIR / "Job Alerts"))
AGENT_OUTPUT_DIR = Path(_app_cfg.get("agent_output_dir") or (BASE_DIR / "Agent output_Updated Cover and resumes"))

for _dir in (CONFIG_DIR, CONTEXT_HUB_DIR, STATE_DIR, DRAFTS_DIR, SKILLS_DIR, LOGS_DIR, JOB_ALERTS_DIR, AGENT_OUTPUT_DIR):
    _dir.mkdir(parents=True, exist_ok=True)

JOB_TRACKER_PATH = JOB_ALERTS_DIR / _app_cfg.get("tracker_filename", "Job Search-SOR.xlsx")
JOB_TRACKER_SHEET = "Job Tracker"
JOB_TRACKER_TABLE = "JobTracker"
APPLICATIONS_SHEET = "Applications & Outcomes"
APPLICATIONS_TABLE = "ApplicationsTracker"

ANTHROPIC_API_KEY_PATH = Path(_app_cfg.get("api_key_path", "/Users/rajhomedesktop/Desktop/Raj JSAW/config/CLAUDE_API_KEY.rtf"))


def _crude_rtf_strip(raw: str) -> str:
    text = re.sub(r"\\'[0-9a-fA-F]{2}", "", raw)
    text = re.sub(r"\\[a-zA-Z]+-?\d* ?", "", text)
    text = text.replace("{", "").replace("}", "")
    text = text.replace("\\\\", "\\").replace("\\{", "{").replace("\\}", "}")
    return text


def _load_api_key_from_rtf(path: Path):
    if not path.exists():
        return None
    raw = path.read_text(encoding="utf-8", errors="ignore")
    try:
        from striprtf.striprtf import rtf_to_text
        text = rtf_to_text(raw)
    except ImportError:
        text = _crude_rtf_strip(raw)
    key_match = re.search(r"sk-ant-[A-Za-z0-9_-]+", text)
    if key_match:
        return key_match.group(0)
    tokens = text.split()
    return tokens[0] if tokens else None


# Direct-paste key (entered in Setup & Paths) takes priority over the .rtf
# file path, which takes priority over an environment variable — this is
# what makes a brand-new machine work: paste the key once in the sidebar,
# nothing to place on disk manually first.
ANTHROPIC_API_KEY = (
    _app_cfg.get("api_key_direct")
    or _load_api_key_from_rtf(ANTHROPIC_API_KEY_PATH)
    or os.getenv("ANTHROPIC_API_KEY")
)

# Gmail credentials: direct entry (Setup & Paths) takes priority over a
# G-Mail.env file, same reasoning — a new machine shouldn't need a
# pre-existing .env file just to get started.
if _app_cfg.get("ceo_email") and _app_cfg.get("email_app_password"):
    CEO_EMAIL = _app_cfg["ceo_email"]
    EMAIL_APP_PASSWORD = _app_cfg["email_app_password"]
else:
    load_dotenv(CONFIG_DIR / "G-Mail.env")
    CEO_EMAIL = os.getenv("CEO_EMAIL")
    EMAIL_APP_PASSWORD = os.getenv("EMAIL_APP_PASSWORD")

_REQUIRED_SECRETS = {
    "ANTHROPIC_API_KEY": ANTHROPIC_API_KEY,
    "CEO_EMAIL": CEO_EMAIL,
    "EMAIL_APP_PASSWORD": EMAIL_APP_PASSWORD,
}


def validate_secrets():
    missing = [k for k, v in _REQUIRED_SECRETS.items() if not v]
    if missing:
        st.error(
            f"Missing required config: {', '.join(missing)}. Check "
            f"`{ANTHROPIC_API_KEY_PATH}` and `{CONFIG_DIR / 'G-Mail.env'}`."
        )
        st.stop()


client = Anthropic(api_key=ANTHROPIC_API_KEY)

MODEL_ROUTING = {
    "search_agent": "claude-haiku-4-5-20251001",
    "talent_agent": "claude-sonnet-5",
    "resume_agent": "claude-sonnet-5",
    "company_research_agent": "claude-sonnet-5",
    "job_e": "claude-sonnet-5",
    "jd_finder": "claude-haiku-4-5-20251001",
}

FITMENT_THRESHOLD_PERCENT = _app_cfg.get("fitment_threshold", 60)
DELIVERY_MODE = "both"  # default the review panel pre-selects — person can override per draft


# =============================================================================
# SHARED ANTHROPIC API HELPER
# =============================================================================

def load_skill(skill_path: Path) -> str:
    """Read a SKILL.md file and strip its YAML frontmatter."""
    text = skill_path.read_text(encoding="utf-8")
    if text.startswith("---"):
        end = text.find("---", 3)
        if end != -1:
            text = text[end + 3:]
    return text.strip()


def call_agent(agent_name: str, skill_path: Path, user_message: str, context_blocks=None, max_tokens: int = 4000) -> str:
    """Call the Anthropic API using a SKILL.md as the system prompt, with prompt caching."""
    if agent_name not in MODEL_ROUTING:
        raise ValueError(f"Unknown agent_name '{agent_name}' — not in MODEL_ROUTING")

    model = MODEL_ROUTING[agent_name]
    system_blocks = [
        {"type": "text", "text": load_skill(skill_path), "cache_control": {"type": "ephemeral"}}
    ]
    for block in context_blocks or []:
        system_blocks.append({"type": "text", "text": block, "cache_control": {"type": "ephemeral"}})

    last_error = None
    for attempt in range(MAX_RETRIES_PER_STAGE + 1):
        try:
            response = client.messages.create(
                model=model,
                max_tokens=max_tokens,
                system=system_blocks,
                messages=[{"role": "user", "content": user_message}],
            )
            return "".join(block.text for block in response.content if block.type == "text")
        except Exception as e:
            last_error = e
            print(f"[{agent_name}] attempt {attempt + 1} failed: {e}")

    raise RuntimeError(f"{agent_name} failed after {MAX_RETRIES_PER_STAGE + 1} attempt(s): {last_error}")


# =============================================================================
# FILE READER — CONTEXT HUB INGESTION
# =============================================================================

SUPPORTED_TEXT_EXTENSIONS = {".md", ".txt", ".markdown", ".rst"}


def read_file_text(path: Path):
    """Extract plain text from any file, regardless of format. Returns None on failure."""
    suffix = path.suffix.lower()

    if suffix == ".docx":
        return _read_docx(path)
    if suffix == ".pdf":
        return _read_pdf(path)
    if suffix == ".pptx":
        return _read_pptx(path)
    if suffix == ".xlsx":
        return _read_xlsx(path)
    if suffix in (".html", ".htm"):
        return _read_html(path)
    if suffix == ".ipynb":
        return _read_ipynb(path)
    if suffix in SUPPORTED_TEXT_EXTENSIONS:
        return path.read_text(encoding="utf-8", errors="ignore")

    try:
        return path.read_text(encoding="utf-8", errors="strict")
    except (UnicodeDecodeError, ValueError):
        print(f"[file_reader] skipping {path.name} — unsupported/binary file type ({suffix or 'no extension'})")
        return None


def _read_docx(path: Path):
    try:
        import docx
    except ImportError:
        print(f"[file_reader] skipping {path.name} — python-docx not installed")
        return None
    try:
        doc = docx.Document(str(path))
        paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if cell.text.strip():
                        paragraphs.append(cell.text)
        return "\n".join(paragraphs)
    except Exception as e:
        print(f"[file_reader] skipping {path.name} — failed to read docx: {e}")
        return None


def _read_pdf(path: Path):
    try:
        from pypdf import PdfReader
    except ImportError:
        print(f"[file_reader] skipping {path.name} — pypdf not installed")
        return None
    try:
        reader = PdfReader(str(path))
        return "\n".join(page.extract_text() or "" for page in reader.pages)
    except Exception as e:
        print(f"[file_reader] skipping {path.name} — failed to read pdf: {e}")
        return None


def _read_pptx(path: Path):
    try:
        from pptx import Presentation
    except ImportError:
        print(f"[file_reader] skipping {path.name} — python-pptx not installed")
        return None
    try:
        prs = Presentation(str(path))
        chunks = []
        for i, slide in enumerate(prs.slides, start=1):
            slide_lines = []
            for shape in slide.shapes:
                if shape.has_text_frame and shape.text_frame.text.strip():
                    slide_lines.append(shape.text_frame.text)
                if shape.has_table:
                    for row in shape.table.rows:
                        for cell in row.cells:
                            if cell.text.strip():
                                slide_lines.append(cell.text)
                if getattr(shape, "has_chart", False):
                    slide_lines.append("[chart — not text-extractable]")
            if slide_lines:
                chunks.append(f"--- slide {i} ---\n" + "\n".join(slide_lines))
        return "\n\n".join(chunks)
    except Exception as e:
        print(f"[file_reader] skipping {path.name} — failed to read pptx: {e}")
        return None


def _read_xlsx(path: Path):
    try:
        import openpyxl
    except ImportError:
        print(f"[file_reader] skipping {path.name} — openpyxl not installed")
        return None
    try:
        wb = openpyxl.load_workbook(str(path), data_only=True, read_only=True)
        chunks = []
        for sheet in wb.worksheets:
            rows_text = []
            for row in sheet.iter_rows(values_only=True):
                cells = [str(c) for c in row if c is not None]
                if cells:
                    rows_text.append(" | ".join(cells))
            if rows_text:
                chunks.append(f"--- sheet: {sheet.title} ---\n" + "\n".join(rows_text))
        return "\n\n".join(chunks)
    except Exception as e:
        print(f"[file_reader] skipping {path.name} — failed to read xlsx: {e}")
        return None


def _read_html(path: Path):
    try:
        from bs4 import BeautifulSoup
    except ImportError:
        print(f"[file_reader] skipping {path.name} — beautifulsoup4 not installed")
        return None
    try:
        raw = path.read_text(encoding="utf-8", errors="ignore")
        soup = BeautifulSoup(raw, "html.parser")
        for tag in soup(["script", "style"]):
            tag.decompose()
        text = soup.get_text(separator="\n")
        lines = [line.strip() for line in text.splitlines()]
        return "\n".join(line for line in lines if line)
    except Exception as e:
        print(f"[file_reader] skipping {path.name} — failed to read html: {e}")
        return None


def _read_ipynb(path: Path):
    """Extracts code + markdown cell source from a Jupyter notebook, in
    order, labeled by cell type — so a notebook in context_hub/ reads as
    readable text, not a JSON blob."""
    try:
        nb_json = json.loads(path.read_text(encoding="utf-8", errors="ignore"))
    except (json.JSONDecodeError, UnicodeDecodeError) as e:
        print(f"[file_reader] skipping {path.name} — failed to parse notebook JSON: {e}")
        return None

    chunks = []
    for i, cell in enumerate(nb_json.get("cells", [])):
        source = cell.get("source", [])
        text = "".join(source) if isinstance(source, list) else str(source)
        if not text.strip():
            continue
        label = "markdown" if cell.get("cell_type") == "markdown" else "code"
        chunks.append(f"--- cell {i} ({label}) ---\n{text}")

    return "\n\n".join(chunks) if chunks else None


# =============================================================================
# SEARCH AGENT
# =============================================================================

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.filters import AutoFilter

SEARCH_SENDERS = {
    "linkedin_alerts": "jobalerts-noreply@linkedin.com",
    "indeed_matches": "donotreply@match.indeed.com",
}

SEARCH_CONFIG_PATH = CONFIG_DIR / "search_config.json"

# Column order as they actually exist in the JobTracker table (B..N).
# Serial # (col A) sits outside the table and is handled separately.
JOB_TRACKER_COLUMNS = [
    "Job Title", "Company", "Location", "Date Posted", "Open or Closed",
    "Hiring Team / Manager", "Hybrid / Remote / On-site", "Pay / CTC / Salary",
    "Actively Recruiting?", "Job Description Summary", "Job URL",
    "Timestamp Captured", "High Match",
    "Match Score", "Strong Points", "Weak Points", "Cover Created", "Resume Updated",
]


def _connect_imap():
    imap = imaplib.IMAP4_SSL("imap.gmail.com")
    imap.login(CEO_EMAIL, EMAIL_APP_PASSWORD)
    imap.select("inbox")
    return imap


def load_search_config() -> dict:
    if not SEARCH_CONFIG_PATH.exists():
        return {"title_filters": [], "exclude_keywords": []}
    return json.loads(SEARCH_CONFIG_PATH.read_text())


def matches_search_config(title: str, config: dict) -> bool:
    """Case-insensitive substring match against title_filters/exclude_keywords.
    Same logic used before as a pre-filter — now applied as a post-hoc label
    instead, via label_high_match() below."""
    title_lower = (title or "").lower()
    excludes = [k.lower() for k in config.get("exclude_keywords", [])]
    if any(k in title_lower for k in excludes):
        return False
    filters = [k.lower() for k in config.get("title_filters", [])]
    if filters and not any(k in title_lower for k in filters):
        return False
    return True


# ---------------------------------------------------------------------------
# Job Tracker Excel — the single source-of-record file, read/appended/labeled
# in place on every run rather than a fresh JSON file each time.
# ---------------------------------------------------------------------------

def open_job_tracker():
    """Opens the JobTracker workbook. Raises clearly if the starter file is
    missing rather than silently creating a blank one — the tracker's
    formatting, dropdowns, and comments were hand-built and shouldn't be
    silently replaced."""
    if not JOB_TRACKER_PATH.exists():
        raise FileNotFoundError(
            f"{JOB_TRACKER_PATH} not found. Expected the starter tracker file "
            f"to already be in 'Job Alerts/' — restore it before running Search Agent."
        )
    wb = openpyxl.load_workbook(JOB_TRACKER_PATH)
    ws = wb[JOB_TRACKER_SHEET]
    return wb, ws


def get_last_captured_timestamp(ws):
    """Max 'Timestamp Captured' across existing rows, or None on a tracker
    with no data rows yet (first-ever run) — falls back to LOOKBACK_DAYS
    in that case."""
    col_idx = 2 + JOB_TRACKER_COLUMNS.index("Timestamp Captured")
    latest = None
    for (val,) in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=col_idx, max_col=col_idx, values_only=True):
        if val and (latest is None or val > latest):
            latest = val
    return latest


def get_existing_urls(ws) -> set:
    col_idx = 2 + JOB_TRACKER_COLUMNS.index("Job URL")
    urls = set()
    for (val,) in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=col_idx, max_col=col_idx, values_only=True):
        if val:
            urls.add(val)
    return urls


def _next_serial(ws) -> int:
    max_serial = 0
    for (val,) in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=1, values_only=True):
        if isinstance(val, (int, float)):
            max_serial = max(max_serial, int(val))
    return max_serial + 1


def _next_empty_row(ws) -> int:
    """Finds the first row (from row 2) with no Job Title. ws.max_row can't
    be trusted for this — the template has data-validation/formatting
    pre-applied down to row 500, so openpyxl reports max_row=499 even on a
    completely empty tracker, which would put new rows at 500 instead of 2."""
    title_col = 2 + JOB_TRACKER_COLUMNS.index("Job Title")
    row_num = 2
    while ws.cell(row=row_num, column=title_col).value is not None:
        row_num += 1
    return row_num


_VALID_WORKPLACE_TYPES = {"Remote", "Hybrid", "On-site"}


def append_posting_row(ws, posting: dict) -> int:
    """Appends one new row for a posting parsed from Gmail. Fields the
    crawler fills in later (Date Posted, Open or Closed, Hiring Team, Pay,
    JD Summary) are left blank at this stage — Search Agent only writes
    what it actually knows from the alert email."""
    row_num = _next_empty_row(ws)
    serial = _next_serial(ws)

    flavors = posting.get("flavors") or []
    if any("actively" in f.lower() for f in flavors):
        recruiting = "Yes"
    elif flavors:
        recruiting = "Unclear"
    else:
        recruiting = None

    workplace = posting.get("workplace_type")
    if workplace not in _VALID_WORKPLACE_TYPES:
        workplace = None  # only write values the dropdown actually accepts

    row_values = [
        posting.get("title"), posting.get("company"), posting.get("location"),
        None, None, None, workplace, None, recruiting, None,
        posting.get("url"), datetime.now(), None,
        None, None, None, None, None,  # Match Score, Strong/Weak Points, Cover Created, Resume Updated
    ]
    ws.cell(row=row_num, column=1, value=serial)
    for offset, val in enumerate(row_values):
        ws.cell(row=row_num, column=2 + offset, value=val)

    # Real clickable hyperlink, not just URL text — per project decision
    url_col = 2 + JOB_TRACKER_COLUMNS.index("Job URL")
    url_cell = ws.cell(row=row_num, column=url_col)
    if posting.get("url"):
        url_cell.hyperlink = posting["url"]
        url_cell.style = "Hyperlink"

    return row_num


def label_high_match(ws, config: dict) -> int:
    """Re-applies title_filters/exclude_keywords across EVERY row in the
    tracker, not just new ones — so editing search_config.json and
    re-running relabels existing rows too, without needing a fresh search."""
    title_col = 2 + JOB_TRACKER_COLUMNS.index("Job Title")
    match_col = 2 + JOB_TRACKER_COLUMNS.index("High Match")
    labeled = 0
    for row_num in range(2, ws.max_row + 1):
        title = ws.cell(row=row_num, column=title_col).value
        if not title:
            continue
        is_match = matches_search_config(title, config)
        ws.cell(row=row_num, column=match_col, value=("High Match" if is_match else "No"))
        labeled += 1
    return labeled


def _resize_table(ws):
    """Grows the JobTracker table's ref to cover exactly the rows with real
    data — same ws.max_row caveat as _next_empty_row() above, so this finds
    the last non-empty row by scanning rather than trusting max_row, which
    would stretch the table down to row 500 regardless of actual content."""
    if JOB_TRACKER_TABLE not in ws.tables:
        return
    last_data_row = _next_empty_row(ws) - 1
    table = ws.tables[JOB_TRACKER_TABLE]
    last_col = get_column_letter(1 + len(JOB_TRACKER_COLUMNS))  # B..S
    new_ref = f"B1:{last_col}{max(last_data_row, 1)}"
    table.ref = new_ref
    table.autoFilter = AutoFilter(ref=new_ref)


# ---------------------------------------------------------------------------
# Applications & Outcomes — separate tab, only touched when the person tells
# JOB-e they actually applied. Never auto-populated from Cover Created — a
# drafted cover letter isn't the same fact as an application actually sent.
# ---------------------------------------------------------------------------

APPLICATIONS_COLUMNS = ["Company", "Role", "Date Applied", "Method", "Outcome", "Outcome Date", "Notes"]


def _applications_next_row(ws) -> int:
    """Checks Company (col 2), not Serial # (col 1) — auto-detected
    application-outcome rows (from Gmail scanning) may not have a confident
    Serial # match to link back to JobTracker, but Company is always
    populated either way, so it's the reliable emptiness signal."""
    row_num = 2
    while ws.cell(row=row_num, column=2).value is not None:
        row_num += 1
    return row_num


def add_application_outcome(serial, company, role, date_applied=None, method=None, outcome=None, outcome_date=None, notes=None):
    """Appends one row to the Applications & Outcomes tab. Called by JOB-e
    only when the person explicitly says they applied — never automatically."""
    wb = openpyxl.load_workbook(JOB_TRACKER_PATH)
    if APPLICATIONS_SHEET not in wb.sheetnames:
        raise RuntimeError(f"{APPLICATIONS_SHEET} sheet not found — tracker may need migration.")
    ws = wb[APPLICATIONS_SHEET]

    row_num = _applications_next_row(ws)
    ws.cell(row=row_num, column=1, value=serial)
    ws.cell(row=row_num, column=2, value=company)
    ws.cell(row=row_num, column=3, value=role)
    ws.cell(row=row_num, column=4, value=date_applied or datetime.now().strftime("%m/%d/%Y"))
    ws.cell(row=row_num, column=5, value=method)
    ws.cell(row=row_num, column=6, value=outcome or "Applied")
    ws.cell(row=row_num, column=7, value=outcome_date)
    ws.cell(row=row_num, column=8, value=notes)

    if APPLICATIONS_TABLE in ws.tables:
        table = ws.tables[APPLICATIONS_TABLE]
        table.ref = f"A1:H{row_num}"
        table.autoFilter = AutoFilter(ref=table.ref)

    wb.save(JOB_TRACKER_PATH)
    print(f"Applications & Outcomes: added {company} / {role}")
    return row_num


def save_job_tracker(wb, ws):
    _resize_table(ws)
    wb.save(JOB_TRACKER_PATH)


# ---------------------------------------------------------------------------
# Gmail parsing — LinkedIn digest (HTML) and Indeed match (plain text)
# ---------------------------------------------------------------------------

def fetch_alert_emails(since_dt):
    """Fetches both plain-text and HTML bodies for alert emails newer than
    since_dt. LinkedIn digests are parsed from HTML (workplace type and
    consistent badge text only exist there — plain text drops both).
    Indeed match emails are parsed from plain text."""
    imap = _connect_imap()
    since = since_dt.strftime("%d-%b-%Y")
    raw_messages = []

    for source, sender in SEARCH_SENDERS.items():
        status, ids = imap.search(None, f'(FROM "{sender}" SINCE "{since}")')
        if status != "OK":
            continue
        for msg_id in ids[0].split():
            status, data = imap.fetch(msg_id, "(RFC822)")
            if status != "OK" or not data or not data[0]:
                continue
            msg = email_lib.message_from_bytes(data[0][1])
            plaintext, html = "", ""
            for part in msg.walk():
                content_type = part.get_content_type()
                if content_type == "text/plain" and not plaintext:
                    payload = part.get_payload(decode=True)
                    if payload:
                        charset = part.get_content_charset() or "utf-8"
                        plaintext = payload.decode(charset, errors="ignore")
                elif content_type == "text/html" and not html:
                    payload = part.get_payload(decode=True)
                    if payload:
                        charset = part.get_content_charset() or "utf-8"
                        html = payload.decode(charset, errors="ignore")
            raw_messages.append({
                "source": source, "subject": str(msg["subject"]),
                "plaintext_body": plaintext, "html_body": html,
            })

    imap.logout()
    return raw_messages


def _clean_linkedin_url(raw_url: str) -> str:
    job_id_match = re.search(r"/jobs/view/(\d+)", raw_url)
    if job_id_match:
        return f"https://www.linkedin.com/jobs/view/{job_id_match.group(1)}/"
    return raw_url.split("&")[0]


def _is_title_link(tag):
    classes = tag.get("class") or []
    return tag.name == "a" and "font-bold" in classes and "text-system-blue-50" in classes


def _is_company_location_line(tag):
    classes = tag.get("class") or []
    return tag.name == "p" and "text-system-gray-100" in classes and "text-xs" in classes


def _is_flavor_badge(tag):
    return tag.name == "p" and tag.get("class") == ["job-card-flavor__detail"]


def parse_linkedin_digest(html: str):
    """Parses a LinkedIn job-alert digest from its HTML body using LinkedIn's
    actual CSS classes (verified against a real 6-job digest pulled from
    Gmail), not guessed markup."""
    from bs4 import BeautifulSoup

    if not html:
        return []

    soup = BeautifulSoup(html, "html.parser")
    elements = soup.find_all(lambda t: _is_title_link(t) or _is_company_location_line(t) or _is_flavor_badge(t))

    cards, current = [], None
    for el in elements:
        if _is_title_link(el):
            if current:
                cards.append(current)
            current = {"title": el.get_text(strip=True), "raw_url": el.get("href", ""), "company_location_raw": None, "flavors": []}
        elif _is_company_location_line(el) and current is not None:
            current["company_location_raw"] = el.get_text(strip=True)
        elif _is_flavor_badge(el) and current is not None:
            current["flavors"].append(el.get_text(strip=True))
    if current:
        cards.append(current)

    postings = []
    for card in cards:
        if not card["raw_url"]:
            continue
        raw = card["company_location_raw"] or ""
        parts = raw.split("\u00b7")
        company = parts[0].strip() if parts else ""
        rest = parts[1].strip() if len(parts) > 1 else ""
        m = re.match(r"^(.*?)\s*\(([^)]+)\)\s*$", rest)
        location, workplace_type = (m.group(1).strip(), m.group(2).strip()) if m else (rest, None)

        postings.append({
            "title": card["title"], "company": company, "location": location,
            "workplace_type": workplace_type, "flavors": card["flavors"],
            "source": "linkedin_alerts", "url": _clean_linkedin_url(card["raw_url"]),
        })

    return postings


def parse_indeed_match(text: str, subject: str):
    paragraphs = [p.strip() for p in re.split(r"\n\s*\n", text) if p.strip()]
    title = company = location = salary = None

    salary_para = next((p for p in paragraphs if re.search(r"^Salary:", p, re.MULTILINE)), None)
    if salary_para:
        lines = [l.strip() for l in salary_para.split("\n") if l.strip()]
        salary_idx = next(i for i, l in enumerate(lines) if l.startswith("Salary:"))
        salary = lines[salary_idx]
        if salary_idx >= 3:
            title, company, location = lines[0], lines[1], lines[2]
    elif len(paragraphs) > 2:
        lines = [l.strip() for l in paragraphs[2].split("\n") if l.strip()]
        if len(lines) >= 3:
            title, company, location = lines[0], lines[1], lines[2]

    if not title:
        return []

    url_match = re.search(r"View job:\s*(\S+)", text)
    url = url_match.group(1) if url_match else ""
    if not url:
        return []

    return [{
        "title": title, "company": company or "", "location": location or "",
        "workplace_type": None, "flavors": [], "source": "indeed_matches",
        "url": url, "salary": salary,
    }]


def parse_postings(raw_messages):
    postings = []
    for msg in raw_messages:
        if msg["source"] == "linkedin_alerts":
            postings.extend(parse_linkedin_digest(msg["html_body"]))
        elif msg["source"] == "indeed_matches":
            postings.extend(parse_indeed_match(msg["plaintext_body"], msg["subject"]))
    return postings


# ---------------------------------------------------------------------------
# Application/outcome email scanning — "data from Gmail" for the
# Applications & Outcomes tab, per the actual spec (not a manual JOB-e
# trigger, which was the original wrong design). Best-effort: application
# confirmation/rejection/interview emails come from thousands of different
# ATS senders with no consistent format, unlike the two known job-alert
# senders — this can't be pattern-matched the way LinkedIn/Indeed parsing
# is, so it's a batched classification pass instead.
# ---------------------------------------------------------------------------

# Proven query pattern — not guessed. Ported from a prior successful manual
# pass over this same inbox (via the Gmail MCP connector in a separate chat)
# that found 113 real application confirmations over 6 months using exactly
# this subject-based approach. Two lessons carried over directly:
# (1) subject-based search reliably catches Workday/Greenhouse/iCIMS/Ashby/
#     SmartRecruiters confirmations — body search was less reliable there;
# (2) distinguishing confirmation vs. rejection vs. interview genuinely
#     needs thread/body content, not just the subject line, so this fetches
#     a short body snippet per candidate now, not headers alone.
_OUTCOME_SUBJECT_TERMS = (
    'subject:application OR subject:applying OR subject:"interest in" OR '
    'subject:interview OR subject:"thank you for applying" OR '
    'subject:"application received" OR subject:"update on your application" OR '
    'subject:"thank you from" OR subject:"thank you for your interest" OR '
    'subject:"thank you for your application"'
)


def _fetch_candidate_outcome_emails(since_dt, max_candidates=80):
    """
    Uses Gmail's X-GM-RAW IMAP extension, which accepts the exact same query
    syntax as Gmail's own search bar — so this reuses a proven working query
    rather than the broad fetch-everything-then-filter approach this had
    before. Grabs a short body snippet per candidate (not just headers) so
    the classification step has enough signal to tell a confirmation from a
    rejection sharing the same thread/subject.

    Fetches headers and body as TWO SEPARATE imap.fetch() calls per message,
    not one combined call requesting both BODY.PEEK sections at once — a
    real bug found via a live diagnostic run: combining them produced a
    multi-part IMAP response where imaplib's data list ordering couldn't be
    relied on, and reading data[0][1] silently returned empty header
    values for every message (the search itself worked — real message IDs
    came back — only the parsing was broken). Two simple, unambiguous
    fetches trade a bit of efficiency for actually working correctly.
    """
    imap = _connect_imap()
    since_days = max((datetime.now() - since_dt).days, 1)
    gmail_query = (
        f"newer_than:{since_days}d ({_OUTCOME_SUBJECT_TERMS}) "
        f"-from:linkedin.com -from:indeed.com"
    )
    # -category:promotions was dropped: confirmed via live diagnostic that Gmail
    # miscategorizes real interview-scheduling automation (e.g. paradox.ai-based
    # "Amethyst from Accenture") as Promotions, silently stripping genuine
    # interview emails before they ever reached classification. Filtering
    # actual promotional noise is now the classification step's job, not a
    # blunt category exclusion at the search level.

    escaped_query = gmail_query.replace("\\", "\\\\").replace('"', '\\"')
    try:
        status, ids = imap.search(None, "X-GM-RAW", f'"{escaped_query}"')
    except Exception as e:
        print(f"[search_agent] X-GM-RAW search failed ({e}) — falling back to a plain SINCE search")
        since = since_dt.strftime("%d-%b-%Y")
        status, ids = imap.search(None, f'(SINCE "{since}")')

    if status != "OK" or not ids or not ids[0]:
        imap.logout()
        return []

    known_alert_senders = set(SEARCH_SENDERS.values())
    candidates = []
    for msg_id in ids[0].split()[-max_candidates:]:
        header_status, header_data = imap.fetch(msg_id, "(BODY.PEEK[HEADER.FIELDS (FROM SUBJECT DATE)])")
        if header_status != "OK" or not header_data or not header_data[0]:
            continue
        header_bytes = header_data[0][1] if isinstance(header_data[0], tuple) else header_data[0]
        msg = email_lib.message_from_bytes(header_bytes)

        def _decode_header_value(raw_value):
            """Real emails routinely MIME-encode Subject/From when they contain
            non-ASCII characters (RFC 2047, e.g. '=?UTF-8?Q?...?='). Found via
            live diagnostic: several genuine interview-email subjects were
            coming through as raw undecoded Q-encoded gibberish, unreadable to
            both a human debugging this and to the classification step."""
            if not raw_value:
                return ""
            try:
                parts = email_lib.header.decode_header(raw_value)
                decoded = "".join(
                    (part.decode(enc or "utf-8", errors="ignore") if isinstance(part, bytes) else part)
                    for part, enc in parts
                )
                return decoded
            except Exception:
                return str(raw_value)

        from_addr = _decode_header_value(msg.get("From", ""))
        subject = _decode_header_value(msg.get("Subject", ""))
        date_str = str(msg.get("Date", ""))

        if not from_addr and not subject:
            print(f"[search_agent] warning: empty headers parsed for message {msg_id} — skipping")
            continue
        if any(sender in from_addr for sender in known_alert_senders):
            continue

        snippet = ""
        body_status, body_data = imap.fetch(msg_id, "(BODY.PEEK[TEXT])")
        if body_status == "OK" and body_data and body_data[0]:
            body_bytes = body_data[0][1] if isinstance(body_data[0], tuple) else body_data[0]
            try:
                body_msg = email_lib.message_from_bytes(body_bytes)
                if body_msg.get_content_maintype() == "multipart" or body_msg.is_multipart():
                    for part in body_msg.walk():
                        if part.get_content_type() == "text/plain":
                            payload = part.get_payload(decode=True)
                            if payload:
                                charset = part.get_content_charset() or "utf-8"
                                snippet = payload.decode(charset, errors="ignore")[:500]
                                break
                else:
                    payload = body_msg.get_payload(decode=True)
                    if payload:
                        charset = body_msg.get_content_charset() or "utf-8"
                        snippet = payload.decode(charset, errors="ignore")[:500]
            except Exception as e:
                print(f"[search_agent] warning: could not decode body for message {msg_id}: {e}")

        candidates.append({"from": from_addr, "subject": subject, "date": date_str, "snippet": snippet})

    imap.logout()
    return candidates


def _classify_outcome_emails(candidates, known_companies):
    """One batched call — classifies which candidate emails are actually
    about a job application (not newsletters, marketing, unrelated mail),
    and extracts company/role/outcome for the ones that are. Conservative:
    only confident, clearly-application-related emails get logged; ambiguous
    ones are left out rather than guessed at."""
    if not candidates:
        return []

    email_list = "\n".join(
        f"{i}. From: {c['from']} | Subject: {c['subject']} | Date: {c['date']}\n"
        f"   Snippet: {c.get('snippet', '')[:300]}"
        for i, c in enumerate(candidates)
    )
    companies_hint = ", ".join(sorted(set(known_companies))[:100])

    user_message = (
        f"Which of these emails are actually about a job application Raj submitted "
        f"— a confirmation it was received, a status update, a rejection, an interview "
        f"invite, or an offer? Ignore newsletters, marketing, job ALERT emails (new "
        f"posting notifications, not application status), and anything unrelated. "
        f"The snippet matters here — a rejection or update often shares the same "
        f"subject as the original confirmation, so read the actual content, not just "
        f"the subject line.\n\n"
        f"Companies already in his job tracker, for matching: {companies_hint}\n\n"
        f"{email_list}\n\n"
        f"Only include emails you're genuinely confident are about an application's "
        f"status — skip anything ambiguous rather than guessing.\n\n"
        f"Respond with one line per relevant email, nothing else:\n"
        f"<row>index|company|role_if_guessable_else_Unknown|Applied_or_Interview_or_Rejected_or_Offer_or_No_Response</row>"
    )
    try:
        response = client.messages.create(
            model=MODEL_ROUTING["search_agent"], max_tokens=1500,
            messages=[{"role": "user", "content": user_message}],
        )
        text = "".join(b.text for b in response.content if b.type == "text")
        results = []
        for m in re.finditer(r"<row>(.*?)</row>", text):
            parts = m.group(1).split("|")
            if len(parts) != 4:
                continue
            idx_str, company, role, outcome = [p.strip() for p in parts]
            try:
                idx = int(idx_str)
            except ValueError:
                continue
            if idx >= len(candidates) or outcome not in ("Applied", "Interview", "Rejected", "Offer", "No Response"):
                continue
            results.append({"candidate": candidates[idx], "company": company, "role": role, "outcome": outcome})
        return results
    except Exception as e:
        print(f"[search_agent] outcome email classification failed: {e}")
        return []


def scan_gmail_for_application_outcomes(lookback_days: int = None):
    """
    Best-effort scan for application confirmation/status/rejection/interview/
    offer emails, appended to the Applications & Outcomes tab. This is
    inherently fuzzier than the LinkedIn/Indeed alert parsing — those come
    from 2 known senders with a fixed template; application-status emails
    come from an unpredictable mix of ATS platforms with no consistent
    format, so this classifies rather than pattern-matches, and skips
    anything it's not confident about rather than guessing.

    lookback_days=None (the default): INCREMENTAL — scans only since the
    last successful scan (tracked in the local config file), not a fixed
    window re-scanned every time. This was a real cost problem before this
    fix: a daily sync was re-fetching and re-classifying the same
    overlapping 14 days of mail every single day, burning real token cost
    on emails already seen. First-ever run (no recorded last-scan time)
    falls back to 14 days.

    Pass an explicit number (e.g. 60, 180) to force a deliberate deeper
    look-back — e.g. catching up after being away for a week, or backfilling
    older history the incremental default would never reach on its own.

    For deep historical backfills, prefer import_historical_applications()
    against an already-compiled tracker if one exists — meaningfully more
    reliable than re-deriving the same data through best-effort live
    classification.
    """
    wb, ws = open_job_tracker()
    known_companies = [
        ws.cell(row=r, column=3).value for r in range(2, ws.max_row + 1)
        if ws.cell(row=r, column=3).value
    ]

    if lookback_days is not None:
        since_dt = datetime.now() - timedelta(days=lookback_days)
        effective_lookback_days = lookback_days
    else:
        last_scan_iso = _app_cfg.get("last_applications_scan")
        since_dt = None
        if last_scan_iso:
            try:
                since_dt = datetime.fromisoformat(last_scan_iso)
            except ValueError:
                since_dt = None
        if since_dt is None:
            since_dt = datetime.now() - timedelta(days=14)  # first-ever run
        effective_lookback_days = max((datetime.now() - since_dt).days, 1)

    print(f"Applications scan: scanning since {since_dt.strftime('%Y-%m-%d %H:%M')} "
          f"({'explicit' if lookback_days is not None else 'incremental — since last scan'})")

    # Scale the candidate cap with the lookback window — the original fixed
    # 80 was sized for the 14-day default; a 6-month ask against the same
    # cap would silently only see the most recent ~80 matches and miss the rest.
    max_candidates = min(80 + (effective_lookback_days // 14) * 40, 400)
    candidates = _fetch_candidate_outcome_emails(since_dt, max_candidates=max_candidates)
    print(f"Applications scan: {len(candidates)} candidate email(s) to classify")

    classified = _classify_outcome_emails(candidates, known_companies)
    print(f"Applications scan: {len(classified)} confidently identified as application-related")

    if APPLICATIONS_SHEET not in wb.sheetnames:
        print(f"[search_agent] {APPLICATIONS_SHEET} sheet not found — skipping")
        return []

    app_ws = wb[APPLICATIONS_SHEET]
    existing = set()
    for r in range(2, app_ws.max_row + 1):
        c, role = app_ws.cell(row=r, column=2).value, app_ws.cell(row=r, column=3).value
        if c:
            existing.add((c.lower(), (role or "").lower()))

    # Best-effort Serial # lookup by company name, for traceability back to
    # JobTracker — not required (many auto-detected rows won't match
    # confidently), just attached when it's unambiguous.
    company_to_serial = {}
    for r in range(2, ws.max_row + 1):
        c = ws.cell(row=r, column=3).value
        if c:
            company_to_serial.setdefault(c.lower(), ws.cell(row=r, column=1).value)

    added = []
    for item in classified:
        key = (item["company"].lower(), item["role"].lower())
        if key in existing:
            continue  # already logged, don't duplicate
        row_num = _applications_next_row(app_ws)
        serial = company_to_serial.get(item["company"].lower())
        if serial:
            app_ws.cell(row=row_num, column=1, value=serial)
        app_ws.cell(row=row_num, column=2, value=item["company"])
        app_ws.cell(row=row_num, column=3, value=item["role"])
        app_ws.cell(row=row_num, column=5, value="Email (auto-detected)")
        app_ws.cell(row=row_num, column=6, value=item["outcome"])
        app_ws.cell(row=row_num, column=8, value=f"Auto-detected from: {item['candidate']['subject']!r}")
        existing.add(key)
        added.append(item)

    if added and APPLICATIONS_TABLE in app_ws.tables:
        table = app_ws.tables[APPLICATIONS_TABLE]
        last_row = _applications_next_row(app_ws) - 1
        table.ref = f"A1:H{last_row}"
        table.autoFilter = AutoFilter(ref=table.ref)

    save_job_tracker(wb, ws)  # saves the whole workbook, both sheets

    # Record when this scan actually ran, so the NEXT incremental (default)
    # call knows to start from here rather than re-scanning the same window.
    _updated_cfg = dict(_app_cfg)
    _updated_cfg["last_applications_scan"] = datetime.now().isoformat()
    save_app_config(_updated_cfg)

    print(f"Applications scan: {len(added)} new row(s) added to {APPLICATIONS_SHEET}")
    return added


_STATUS_TO_OUTCOME_PATTERNS = [
    (r"reject|not selected|closed|cancelled", "Rejected"),
    (r"withdrawn", "No Response"),
    (r"interview", "Interview"),
    (r"offer", "Offer"),
]


def _normalize_historical_status(raw_status: str) -> str:
    """Maps free-text status values (e.g. 'Rejected (3/25)', 'Not selected
    (update 6/17)', 'In process (recruiter/HM engaged)') onto the tab's
    fixed Applied/Interview/Rejected/Offer/No Response dropdown. The
    original raw text is preserved in Notes regardless, so nothing is lost
    in the normalization — this is just what goes in the constrained cell."""
    lowered = (raw_status or "").lower()
    for pattern, outcome in _STATUS_TO_OUTCOME_PATTERNS:
        if re.search(pattern, lowered):
            return outcome
    return "Applied"  # covers "In progress", "Update received", "Applied - ...", etc.


def import_historical_applications(xlsx_path):
    """
    One-time bulk import from an already-compiled, already-verified
    application tracker (built via a careful multi-query Gmail search and
    thread-content review in a separate session — meaningfully more
    reliable than re-deriving the same data through this notebook's
    best-effort live classification pass). Dedupes against existing
    (company, role) pairs already in the tab, same as the live scan.
    """
    xlsx_path = Path(xlsx_path)
    source_wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    source_ws = source_wb[source_wb.sheetnames[0]]

    wb, ws = open_job_tracker()
    if APPLICATIONS_SHEET not in wb.sheetnames:
        print(f"[import] {APPLICATIONS_SHEET} sheet not found — aborting")
        return []

    app_ws = wb[APPLICATIONS_SHEET]
    existing = set()
    for r in range(2, app_ws.max_row + 1):
        c, role = app_ws.cell(row=r, column=2).value, app_ws.cell(row=r, column=3).value
        if c:
            existing.add((c.lower(), (role or "").lower()))

    company_to_serial = {}
    for r in range(2, ws.max_row + 1):
        c = ws.cell(row=r, column=3).value
        if c:
            company_to_serial.setdefault(c.lower(), ws.cell(row=r, column=1).value)

    added = []
    for r in range(2, source_ws.max_row + 1):
        company = source_ws.cell(row=r, column=1).value
        if not company:
            continue
        role = source_ws.cell(row=r, column=2).value or ""
        jd_brief = source_ws.cell(row=r, column=3).value or ""
        date_applied = source_ws.cell(row=r, column=4).value
        salary = source_ws.cell(row=r, column=5).value
        raw_status = source_ws.cell(row=r, column=6).value or ""

        key = (company.lower(), role.lower())
        if key in existing:
            continue

        row_num = _applications_next_row(app_ws)
        serial = company_to_serial.get(company.lower())
        if serial:
            app_ws.cell(row=row_num, column=1, value=serial)
        app_ws.cell(row=row_num, column=2, value=company)
        app_ws.cell(row=row_num, column=3, value=role)
        app_ws.cell(row=row_num, column=4, value=str(date_applied) if date_applied else None)
        app_ws.cell(row=row_num, column=5, value="Email (imported from historical tracker)")
        app_ws.cell(row=row_num, column=6, value=_normalize_historical_status(raw_status))
        notes_parts = [p for p in [jd_brief, f"Original status: {raw_status}" if raw_status else None,
                                    f"Salary: {salary}" if salary and salary != "Not mentioned" else None] if p]
        app_ws.cell(row=row_num, column=8, value=" | ".join(notes_parts))
        existing.add(key)
        added.append({"company": company, "role": role})

    if added and APPLICATIONS_TABLE in app_ws.tables:
        table = app_ws.tables[APPLICATIONS_TABLE]
        last_row = _applications_next_row(app_ws) - 1
        table.ref = f"A1:H{last_row}"
        table.autoFilter = AutoFilter(ref=table.ref)

    save_job_tracker(wb, ws)
    print(f"Import: {len(added)} historical application(s) added, {len(existing) - len(added)} already present / skipped")
    return added


def search_agent_run(applications_lookback_days: int = None):
    """
    Gmail -> JobTracker.xlsx, unfiltered. Reads only messages newer than the
    tracker's own last 'Timestamp Captured' (first run falls back to
    LOOKBACK_DAYS), appends new postings as new rows, then re-labels EVERY
    row's High Match column from search_config.json — no pre-filtering, no
    JD lookup here at all. That's Talent Agent's job, and only for rows
    already marked High Match.

    Also scans Gmail for application-status emails (confirmation, rejection,
    interview, offer) and appends confident matches to the Applications &
    Outcomes tab — this is Search Agent's job per spec ('data from Gmail'),
    not a manual JOB-e trigger. applications_lookback_days=None (default)
    scans incrementally since the last successful scan, not a fixed window
    re-covered every time — a real fix for token cost, since a fixed 14-day
    window re-scanned daily was re-fetching and re-classifying the same
    overlapping mail every single day. Pass an explicit number (e.g. 60,
    180) to force a deliberate deeper look-back.
    """
    wb, ws = open_job_tracker()

    last_ts = get_last_captured_timestamp(ws)
    since_dt = last_ts if last_ts else (datetime.now() - timedelta(days=LOOKBACK_DAYS))
    print(f"Searching Gmail since: {since_dt}")

    raw = fetch_alert_emails(since_dt)
    postings = parse_postings(raw)

    existing_urls = get_existing_urls(ws)
    new_postings = [p for p in postings if p.get("url") and p["url"] not in existing_urls]

    for p in new_postings:
        append_posting_row(ws, p)

    config = load_search_config()
    labeled = label_high_match(ws, config)

    save_job_tracker(wb, ws)
    print(f"Search Agent: {len(new_postings)} new posting(s) added, {labeled} row(s) labeled (of {len(postings)} parsed this run)")

    scan_gmail_for_application_outcomes(lookback_days=applications_lookback_days)

    return JOB_TRACKER_PATH

# =============================================================================
# TALENT AGENT
# =============================================================================

TALENT_AGENT_SKILL_PATH = SKILLS_DIR / "talent-agent-SKILL.md"

# LinkedIn/Indeed are never crawled directly — by design, not just preference.
_BLOCKED_CRAWL_DOMAINS = ("linkedin.com", "indeed.com")

_POSTING_UNAVAILABLE_PATTERNS = [
    r"this position has been filled", r"this job is no longer available",
    r"job posting has expired", r"this position is closed",
    r"no longer accepting applications", r"job not found",
    r"posting is not available", r"no longer active",
]

_BOILERPLATE_LINE_PATTERNS = [
    r"sign in( with email)?", r"join( now| or sign in)?", r"skip to main content",
    r"accept( close)?", r"close", r"language",
    r"(english|français|español|deutsch|italiano|português)( \([^)]+\))?",
    r"apply now!?", r"visit .* career page", r"search by \"?job title.*",
    r"\+ more options", r"loading\.\.\.", r"\*\s*\*\s*\*",
    r"(country/region|city|experience level|location)", r"all",
    r"x reset filters", r"select how often.*", r"create alert", r"×",
    r"sitemap", r"job applicant privacy notice", r"privacy",
    r"accessibility statement", r"cookie(s)? (policy|notice|settings)",
]


def _detect_posting_unavailable(text: str) -> bool:
    return any(re.search(p, text, re.IGNORECASE) for p in _POSTING_UNAVAILABLE_PATTERNS)


def _strip_boilerplate_lines(text: str) -> str:
    kept = []
    for line in text.split("\n"):
        visible = re.sub(r"^\s*[\*\-]\s*", "", line).strip()
        m = re.match(r'^\[([^\]]*)\]\([^)]*\)$', visible)
        if m:
            visible = m.group(1).strip()
        if not visible:
            continue
        if any(re.fullmatch(p, visible, re.IGNORECASE) for p in _BOILERPLATE_LINE_PATTERNS):
            continue
        kept.append(line)
    return "\n".join(kept)


def find_company_url(title: str, company: str, location: str, max_searches: int = 1):
    """Attempt 1 of 2: one web_search call to find the posting on the
    company's own careers page or a third-party ATS. Hard-blocks
    linkedin.com/indeed.com even if the model returns one anyway."""
    user_message = (
        f"Find the URL for this job posting on the COMPANY'S OWN careers "
        f"page or a third-party ATS listing (Lever, Greenhouse, Workday, "
        f"iCIMS, SmartRecruiters, etc.) — never LinkedIn or Indeed.\n\n"
        f"Title: {title}\nCompany: {company}\nLocation: {location or ''}\n\n"
        f"Do no more than {max_searches} search(es).\n\n"
        f"Respond in exactly this format, nothing else:\n"
        f"<url>the best URL found, or NOT_FOUND</url>"
    )
    try:
        response = client.messages.create(
            model=MODEL_ROUTING["jd_finder"],
            max_tokens=400,
            tools=[{"type": "web_search_20250305", "name": "web_search", "max_uses": max_searches}],
            messages=[{"role": "user", "content": user_message}],
        )
        full_text = "".join(b.text for b in response.content if b.type == "text")
        match = re.search(r"<url>\s*(.*?)\s*</url>", full_text, re.DOTALL)
        if not match:
            return None
        url = match.group(1).strip()
        if not url or url == "NOT_FOUND" or not url.startswith("http"):
            return None
        if any(domain in url for domain in _BLOCKED_CRAWL_DOMAINS):
            print(f"[talent_agent] search returned a blocked domain ({url}) — discarding")
            return None
        return url
    except Exception as e:
        print(f"[talent_agent] company-URL search failed: {e}")
        return None


def crawl_job_description(url: str, timeout_s: int = 30):
    """Attempt 2 of 2: fetches url with crawl4ai and returns cleaned text,
    or None. No further fallback beyond this single attempt."""
    try:
        from crawl4ai import AsyncWebCrawler
        from crawl4ai.async_configs import CrawlerRunConfig
        from crawl4ai.content_filter_strategy import PruningContentFilter
        from crawl4ai.markdown_generation_strategy import DefaultMarkdownGenerator
    except ImportError:
        print("[crawl4ai] not installed — pip install crawl4ai && crawl4ai-setup")
        return None

    async def _crawl():
        config = CrawlerRunConfig(
            page_timeout=timeout_s * 1000,
            markdown_generator=DefaultMarkdownGenerator(content_filter=PruningContentFilter(threshold=0.48)),
        )
        async with AsyncWebCrawler() as crawler:
            return await crawler.arun(url=url, config=config)

    try:
        result = asyncio.run(_crawl())
    except Exception as e:
        print(f"[crawl4ai] failed to fetch {url}: {e}")
        return None

    if not result.success or not result.markdown:
        reason = getattr(result, "error_message", None) or "no content returned"
        print(f"[crawl4ai] could not fetch {url}: {reason}")
        return None

    text = (result.markdown.fit_markdown or "").strip()
    if not text:
        print(f"[crawl4ai] {url} — pruning found no substantial content block, discarding")
        return None

    if _detect_posting_unavailable(text):
        print(f"[crawl4ai] {url} — posting appears to have been filled or closed, not a scraping failure")
        return "__CLOSED__"  # sentinel: crawl succeeded, posting confirmed dead — write Closed, not blank

    cleaned = _strip_boilerplate_lines(text)
    if len(cleaned) < 200:
        print(f"[crawl4ai] {url} — only {len(cleaned)} chars of real content after cleaning, discarding")
        return None

    return cleaned


def extract_structured_fields(raw_text: str, title: str, company: str):
    """One cheap, tightly-bounded model call: turns crawled page text into
    Date Posted / Open or Closed / Hiring Team / Pay / JD Summary, PLUS the
    JD's explicit stated requirements/qualifications as a separate list —
    used by scoring below for checklist-style matching (see
    score_against_context_hub), not stored as its own tracker column."""
    user_message = (
        f"From this job posting page content for '{title}' at '{company}', "
        f"extract these fields. Use exactly 'Unknown' for anything not "
        f"clearly present — never guess or fabricate.\n\n"
        f"<page_content>\n{raw_text[:8000]}\n</page_content>\n\n"
        f"Respond in exactly this format, nothing else:\n"
        f"<fields>\n"
        f"<date_posted>mm/dd/yyyy or Unknown</date_posted>\n"
        f"<status>Open or Closed or Unknown</status>\n"
        f"<hiring_team>name/team if mentioned, else Unknown</hiring_team>\n"
        f"<pay>salary/CTC range if stated, else Unknown</pay>\n"
        f"<summary>2-3 sentence summary of the role</summary>\n"
        f"<requirements>\n"
        f"- each explicitly stated required or preferred qualification, one per line\n"
        f"- verbatim or close to verbatim from the posting — do not paraphrase away specifics\n"
        f"- leave this section empty if the page doesn't state explicit qualifications\n"
        f"</requirements>\n"
        f"</fields>"
    )
    try:
        response = client.messages.create(
            model=MODEL_ROUTING["jd_finder"], max_tokens=800,
            messages=[{"role": "user", "content": user_message}],
        )
        text = "".join(b.text for b in response.content if b.type == "text")

        def _extract(tag):
            m = re.search(f"<{tag}>\\s*(.*?)\\s*</{tag}>", text, re.DOTALL)
            val = m.group(1).strip() if m else "Unknown"
            return None if val.lower() == "unknown" else val

        requirements_raw = _extract("requirements") or ""
        requirements = [l.strip("- ").strip() for l in requirements_raw.split("\n") if l.strip().startswith("-")]

        return {
            "date_posted": _extract("date_posted"), "status": _extract("status"),
            "hiring_team": _extract("hiring_team"), "pay": _extract("pay"),
            "summary": _extract("summary"), "requirements": requirements,
        }
    except Exception as e:
        print(f"[talent_agent] field extraction failed: {e}")
        return None


def load_context_hub():
    """Read every file in context_hub/ into cacheable text blocks — any filename, any format."""
    blocks = []
    files = sorted(
        p for p in CONTEXT_HUB_DIR.iterdir()
        if p.is_file() and p.name not in {".gitkeep", ".DS_Store"} and not p.name.startswith(".")
    )
    if not files:
        print(f"[talent_agent] warning: {CONTEXT_HUB_DIR} is empty — Context Hub has nothing to score against")
        return blocks
    for path in files:
        text = read_file_text(path)
        if text and text.strip():
            blocks.append(f"### {path.name}\n{text}")
    return blocks


def _get_match_keywords() -> list:
    """Live keyword list from the sidebar's Match Keywords panel — session
    state if the person edited it this session, else whatever's saved in
    config. Used as extra scoring signal, never a pre-filter/gate — that
    exact pattern (a narrow keyword list gating what gets scored at all)
    was already tried and found to lose real matches earlier in this
    project, which is why search_config.json's title_filters stayed empty."""
    return st.session_state.get(
        "match_keywords", _app_cfg.get("match_keywords", ["Strategy", "Business Development", "Operations", "ISVs", "Data Science"])
    )


def score_against_context_hub(row_data: dict, context_blocks: list, requirements: list = None):
    """
    Scores fitment with a short, decision-focused rationale — strictly the
    3 columns Talent Agent owns: Match Score, Strong Points, Weak Points.

    When requirements (explicit stated qualifications, from
    extract_structured_fields) are available, scores checklist-style —
    match each one individually (matched / partial / not matched) and
    derive the score from that ratio. This mirrors LinkedIn's own
    'Matches N of M required qualifications' feature, which Raj pointed to
    as the target pattern: it's a more grounded, more interpretable signal
    than an abstract number, and it's what actually produces a usable
    Strong/Weak Points breakdown instead of generic filler.

    Falls back to the general title/company/location assessment when no
    explicit requirements exist (no JD text found at all, or the page
    didn't state clear qualifications) — same conservative behavior as
    before in that case.
    """
    keywords = _get_match_keywords()
    threshold = st.session_state.get("fitment_threshold", FITMENT_THRESHOLD_PERCENT)

    # The threshold is now a real scoring lens, not just something JOB-e
    # references in conversation after the fact. A low threshold signals
    # someone open to a stretch or a deliberate pivot — score should credit
    # real transferable/adjacent experience more generously. A high
    # threshold signals "only count close, ready-now matches" — score
    # strictly. Either way this still has to be grounded in truth: the
    # instruction changes what counts as a positive signal, never invents one.
    if threshold <= 40:
        strictness_note = (
            f"Raj has set his match threshold to {threshold}% — low, which signals "
            f"he's open to a stretch role or is deliberately exploring a different "
            f"direction, not just close matches to his existing history. Score "
            f"generously for real transferable and adjacent experience: don't mark "
            f"something 'not matched' just because it isn't an exact prior title or "
            f"tool — if the underlying skill genuinely transfers, credit it. Still "
            f"never invent experience that isn't there; generosity applies to how "
            f"real experience gets interpreted, not to fabricating what's missing.\n\n"
        )
    elif threshold >= 70:
        strictness_note = (
            f"Raj has set his match threshold to {threshold}% — high, which signals "
            f"he wants only close, ready-now matches surfaced with confidence. Score "
            f"strictly: don't give credit for loosely adjacent experience, and be "
            f"willing to mark something a real gap even if a softer read might call "
            f"it a stretch-fit.\n\n"
        )
    else:
        strictness_note = ""  # middle range — no directional nudge, just the normal grounded assessment

    keyword_note = (
        f"Focus areas Raj has explicitly flagged as his current target direction: "
        f"{', '.join(keywords)}. Actively look for real transferable or adjacent "
        f"experience toward these — don't just check for an exact keyword match "
        f"and move on. If a posting is squarely in one of these areas and Raj has "
        f"genuinely relevant (even if not identical) experience, credit it clearly "
        f"in Strong Points. This is still grounded in truth, never a rubber stamp — "
        f"don't credit a connection that isn't real, but actively look for one that "
        f"is rather than defaulting to 'no direct match.'\n\n"
    ) if keywords else ""

    if requirements:
        req_list_text = "\n".join(f"{i+1}. {r}" for i, r in enumerate(requirements))
        user_message = (
            f"Check this posting's stated requirements against the Context Hub, "
            f"one at a time — same pattern as LinkedIn's own 'Matches N of M "
            f"required qualifications' feature.\n\n"
            f"{strictness_note}"
            f"{keyword_note}"
            f"Posting:\n{json.dumps(row_data, indent=2, default=str)}\n\n"
            f"Requirements stated in the posting:\n{req_list_text}\n\n"
            f"For each one, decide: matched (real, specific proof point exists), "
            f"partial (adjacent experience but not a full match), or not matched "
            f"(no real evidence). Compute the score as roughly "
            f"(matched + 0.5*partial) / total * 100, then round to a sensible number.\n\n"
            f"Respond in exactly this format, nothing else:\n"
            f"<match_score>0-100</match_score>\n"
            f"<strong_points>\n- one short bullet per matched requirement, citing the "
            f"specific real proof point, e.g. '15 years business development experience'\n"
            f"</strong_points>\n"
            f"<weak_points>\n- one short bullet per partial/not-matched requirement, plain "
            f"language, e.g. 'No AWS Bedrock experience'\n"
            f"- empty if every requirement matched\n</weak_points>"
        )
    else:
        user_message = (
            f"Score this posting's fitment against the Context Hub. No explicit "
            f"requirements list was found for this posting, so score "
            f"conservatively from title/company/location alone and say so in "
            f"your reasoning — don't pretend to JD-level confidence you don't have.\n\n"
            f"{strictness_note}"
            f"{keyword_note}"
            f"Posting:\n{json.dumps(row_data, indent=2, default=str)}\n\n"
            f"Respond in exactly this format, nothing else:\n"
            f"<match_score>0-100</match_score>\n"
            f"<strong_points>\n- short bullet, e.g. '15 years business development experience'\n"
            f"- (2-4 bullets total)\n</strong_points>\n"
            f"<weak_points>\n- short bullet, e.g. 'No AWS Bedrock experience'\n"
            f"- (0-4 bullets total, empty if no real gaps)\n</weak_points>"
        )

    try:
        output = call_agent("talent_agent", TALENT_AGENT_SKILL_PATH, user_message, context_blocks)
        score_raw = _parse_tagged(output, "match_score", "0")
        score = int(re.sub(r"\D", "", score_raw) or "0")
        strong = _parse_tagged(output, "strong_points", "").strip()
        weak = _parse_tagged(output, "weak_points", "").strip()
        return {"match_score": score, "strong_points": strong, "weak_points": weak, "error": None}
    except Exception as e:
        error_msg = f"{type(e).__name__}: {e}"
        print(f"[talent_agent] scoring failed: {error_msg}")
        return {"match_score": None, "strong_points": None, "weak_points": None, "error": error_msg}


def _parse_tagged(text: str, tag: str, default=None):
    m = re.search(f"<{tag}>\\s*(.*?)\\s*</{tag}>", text, re.DOTALL)
    return m.group(1).strip() if m else default


def talent_agent_process_row(row_num: int, context_blocks: list, ws=None, wb=None) -> dict:
    """
    Enriches and scores one row. Enrichment (find URL, crawl, extract
    structured fields) only runs if Job Description Summary is empty and
    High Match is set — 2 attempts total (1 search + 1 crawl), no fallback.
    Scoring ALWAYS runs regardless of enrichment outcome — a posting with
    no JD text is still scored conservatively from title/company/location,
    same as before.
    """
    close_after = False
    if wb is None:
        wb, ws = open_job_tracker()
        close_after = True

    cols = {name: 2 + JOB_TRACKER_COLUMNS.index(name) for name in JOB_TRACKER_COLUMNS}
    title = ws.cell(row=row_num, column=cols["Job Title"]).value
    company = ws.cell(row=row_num, column=cols["Company"]).value
    location = ws.cell(row=row_num, column=cols["Location"]).value
    high_match = ws.cell(row=row_num, column=cols["High Match"]).value
    jd_summary = ws.cell(row=row_num, column=cols["Job Description Summary"]).value

    enriched = False
    requirements = None
    if high_match == "High Match" and not jd_summary:
        url = find_company_url(title, company, location)
        if url:
            raw = crawl_job_description(url)
            if raw == "__CLOSED__":
                ws.cell(row=row_num, column=cols["Open or Closed"], value="Closed")
            elif raw:
                fields = extract_structured_fields(raw, title, company)
                if fields:
                    if fields["date_posted"]:
                        ws.cell(row=row_num, column=cols["Date Posted"], value=fields["date_posted"])
                    if fields["status"] in ("Open", "Closed"):
                        ws.cell(row=row_num, column=cols["Open or Closed"], value=fields["status"])
                    if fields["hiring_team"]:
                        ws.cell(row=row_num, column=cols["Hiring Team / Manager"], value=fields["hiring_team"])
                    if fields["pay"]:
                        ws.cell(row=row_num, column=cols["Pay / CTC / Salary"], value=fields["pay"])
                    if fields["summary"]:
                        ws.cell(row=row_num, column=cols["Job Description Summary"], value=fields["summary"])
                        jd_summary = fields["summary"]
                    requirements = fields.get("requirements") or None
                    enriched = True
        else:
            print(f"[talent_agent] row {row_num} ({title!r} @ {company!r}): no company URL found")

    row_data = {
        "title": title, "company": company, "location": location,
        "date_posted": ws.cell(row=row_num, column=cols["Date Posted"]).value,
        "status": ws.cell(row=row_num, column=cols["Open or Closed"]).value,
        "pay": ws.cell(row=row_num, column=cols["Pay / CTC / Salary"]).value,
        "jd_summary": jd_summary,
        "url": ws.cell(row=row_num, column=cols["Job URL"]).value,
    }
    score_result = score_against_context_hub(row_data, context_blocks, requirements=requirements)
    if score_result and score_result.get("error") is None:
        ws.cell(row=row_num, column=cols["Match Score"], value=score_result["match_score"])
        ws.cell(row=row_num, column=cols["Strong Points"], value=score_result["strong_points"])
        ws.cell(row=row_num, column=cols["Weak Points"], value=score_result["weak_points"])
    elif score_result and score_result.get("error"):
        print(f"[talent_agent] row {row_num}: scoring failed, leaving cells blank — {score_result['error']}")

    if close_after:
        save_job_tracker(wb, ws)

    return {"row": row_num, "title": title, "company": company, "enriched": enriched, "score": score_result}


def talent_agent_run(row_nums=None) -> list:
    """
    Processes specific rows if given, else every High Match row without a
    Match Score yet (bulk convenience for 'score all my high matches').
    Skips already-scored rows in bulk mode, saves after every row.
    """
    wb, ws = open_job_tracker()
    context_blocks = load_context_hub()

    if row_nums is None:
        match_col = 2 + JOB_TRACKER_COLUMNS.index("High Match")
        score_col = 2 + JOB_TRACKER_COLUMNS.index("Match Score")
        row_nums = [
            r for r in range(2, ws.max_row + 1)
            if ws.cell(row=r, column=match_col).value == "High Match"
            and ws.cell(row=r, column=score_col).value is None
        ]

    print(f"Talent Agent: {len(row_nums)} row(s) to process")
    results = []
    for i, row_num in enumerate(row_nums, 1):
        title = ws.cell(row=row_num, column=2).value
        print(f"  [{i}/{len(row_nums)}] row {row_num}: {title!r} ...")
        result = talent_agent_process_row(row_num, context_blocks, ws=ws, wb=wb)
        results.append(result)
        score_result = result["score"]
        if score_result and score_result.get("error") is None:
            print(f"      -> score {score_result['match_score']}%")
        elif score_result and score_result.get("error"):
            print(f"      -> SCORING FAILED: {score_result['error']}")
        save_job_tracker(wb, ws)  # save after every row

    print(f"\nTalent Agent: {len(results)} row(s) processed")
    return results


# =============================================================================
# COMPANY RESEARCH AGENT
# =============================================================================

COMPANY_RESEARCH_SKILL_PATH = SKILLS_DIR / "company-research-agent-SKILL.md"


def company_research_run(posting: dict, draft_cover_letter: str = "") -> str:
    jd_text = posting.get("jd_full_text") or posting.get("jd_snippet", "")
    user_message = (
        f"Research this company and critique the current draft cover letter "
        f"against what you find.\n\n"
        f"Company: {posting.get('company')}\n\n"
        f"Job description:\n{jd_text}\n\n"
        f"Current draft cover letter:\n{draft_cover_letter or '(not yet drafted)'}"
    )
    return call_agent("company_research_agent", COMPANY_RESEARCH_SKILL_PATH, user_message)


# =============================================================================
# RESUME AGENT
# =============================================================================

import tempfile

RESUME_AGENT_SKILL_PATH = SKILLS_DIR / "resume-agent-SKILL.md"


def _row_data_for(row_num: int, ws) -> dict:
    cols = {name: 2 + JOB_TRACKER_COLUMNS.index(name) for name in JOB_TRACKER_COLUMNS}
    return {
        "row": row_num,
        "title": ws.cell(row=row_num, column=cols["Job Title"]).value,
        "company": ws.cell(row=row_num, column=cols["Company"]).value,
        "location": ws.cell(row=row_num, column=cols["Location"]).value,
        "jd_summary": ws.cell(row=row_num, column=cols["Job Description Summary"]).value,
        "match_score": ws.cell(row=row_num, column=cols["Match Score"]).value,
        "strong_points": ws.cell(row=row_num, column=cols["Strong Points"]).value,
        "weak_points": ws.cell(row=row_num, column=cols["Weak Points"]).value,
        "url": ws.cell(row=row_num, column=cols["Job URL"]).value,
    }


def _safe_filename_part(s) -> str:
    return re.sub(r"[^\w\-]+", "_", str(s or "")).strip("_") or "unknown"


def draft_cover_letter(row_num: int) -> dict:
    """
    Drafts a cover letter for a specific row, on request only — never
    automatic. Spawns Company Research for a critique + one revision pass.
    Returns text directly; nothing is saved to disk here. Only writes
    Cover Created = Yes on the tracker; no draft persisted to state/.
    """
    wb, ws = open_job_tracker()
    row_data = _row_data_for(row_num, ws)
    context_blocks = load_context_hub()

    user_message = (
        f"Draft a cover letter for this posting, grounded strictly in the "
        f"Context Hub. Use the Talent Agent's scoring as a starting point "
        f"for what to emphasize (strong points) and what to address or "
        f"avoid overclaiming (weak points).\n\n"
        f"Posting + scoring:\n{json.dumps(row_data, indent=2, default=str)}\n\n"
        f"Respond in exactly this format, nothing else:\n"
        f"<cover_letter>\nfull cover letter text\n</cover_letter>"
    )
    draft = call_agent("resume_agent", RESUME_AGENT_SKILL_PATH, user_message, context_blocks)
    cover_text = _parse_tagged(draft, "cover_letter", "")

    posting_for_research = {"company": row_data["company"], "jd_full_text": row_data["jd_summary"]}
    critique = company_research_run(posting_for_research, draft_cover_letter=cover_text)

    revision_message = (
        f"Revise this cover letter based on the critique below. Only accept "
        f"points consistent with the Context Hub's positioning rules.\n\n"
        f"Current draft:\n{cover_text}\n\nCritique:\n{critique}\n\n"
        f"Respond in exactly the same <cover_letter> tag format."
    )
    revised = call_agent("resume_agent", RESUME_AGENT_SKILL_PATH, revision_message, context_blocks)
    final_text = _parse_tagged(revised, "cover_letter", cover_text)

    cols = {name: 2 + JOB_TRACKER_COLUMNS.index(name) for name in JOB_TRACKER_COLUMNS}
    ws.cell(row=row_num, column=cols["Cover Created"], value="Yes")
    save_job_tracker(wb, ws)

    return {"row": row_num, "company": row_data["company"], "title": row_data["title"], "text": final_text, "critique": critique}


def _parse_tagged(text: str, tag: str, default=None):
    m = re.search(f"<{tag}>\\s*(.*?)\\s*</{tag}>", text, re.DOTALL)
    return m.group(1).strip() if m else default


def _format_resume_changes_markdown(resume_changes_raw: str, row_data: dict) -> str:
    changes = re.findall(r"<change>(.*?)</change>", resume_changes_raw, re.DOTALL)
    lines = [f"# Proposed resume changes — {row_data.get('title')} @ {row_data.get('company')}\n"]
    if not changes:
        lines.append("_No specific changes proposed — existing resume already covers this posting well._")
    for i, block in enumerate(changes, 1):
        section = _parse_tagged(block, "section", "General")
        before = _parse_tagged(block, "before", "(new addition)")
        after = _parse_tagged(block, "after", "")
        reason = _parse_tagged(block, "reason", "")
        lines.append(f"## Change {i} — {section}")
        lines.append(f"**Before:** {before}")
        lines.append(f"**After:** {after}")
        if reason:
            lines.append(f"**Why:** {reason}")
        lines.append("")
    return "\n".join(lines)


def draft_resume_changes(row_num: int) -> dict:
    """
    Proposes specific before/after resume edits for a row, on request only.
    Individually-reviewable edits, not a full rewrite — matches the
    'options to edit and save' spec, not an auto-applied resume regeneration.
    """
    wb, ws = open_job_tracker()
    row_data = _row_data_for(row_num, ws)
    context_blocks = load_context_hub()

    user_message = (
        f"Propose specific before/after resume changes for this posting, "
        f"grounded strictly in the Context Hub — not a full rewrite.\n\n"
        f"Posting + scoring:\n{json.dumps(row_data, indent=2, default=str)}\n\n"
        f"Respond in exactly this format, nothing else:\n"
        f"<resume_changes>\n"
        f"<change><section>...</section><before>...</before><after>...</after><reason>...</reason></change>\n"
        f"(zero or more <change> blocks)\n"
        f"</resume_changes>"
    )
    output = call_agent("resume_agent", RESUME_AGENT_SKILL_PATH, user_message, context_blocks)
    changes_raw = _parse_tagged(output, "resume_changes", "")
    changes_markdown = _format_resume_changes_markdown(changes_raw, row_data)

    cols = {name: 2 + JOB_TRACKER_COLUMNS.index(name) for name in JOB_TRACKER_COLUMNS}
    ws.cell(row=row_num, column=cols["Resume Updated"], value="Yes")
    save_job_tracker(wb, ws)

    return {"row": row_num, "company": row_data["company"], "title": row_data["title"], "text": changes_markdown}


# ---------------------------------------------------------------------------
# Output compilation and delivery — text / docx / pdf, mail / save / both
# ---------------------------------------------------------------------------

def compile_docx(text: str, out_path: Path) -> Path:
    from docx import Document
    from docx.shared import Pt

    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    doc.add_paragraph("Rajarshi Majumder")
    doc.add_paragraph("469.236.3956 | rajarshi.majumder2@gmail.com")
    doc.add_paragraph(datetime.now().strftime("%B %d, %Y"))
    doc.add_paragraph("")
    for para in text.split("\n\n"):
        if para.strip():
            doc.add_paragraph(para.strip())

    doc.save(str(out_path))
    return out_path


def compile_pdf(text: str, out_path: Path) -> Path:
    from fpdf import FPDF

    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Helvetica", size=11)
    pdf.cell(0, 8, "Rajarshi Majumder", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 8, "469.236.3956 | rajarshi.majumder2@gmail.com", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 8, datetime.now().strftime("%B %d, %Y"), new_x="LMARGIN", new_y="NEXT")
    pdf.ln(4)
    pdf.multi_cell(0, 6, text)
    pdf.output(str(out_path))
    return out_path


def send_email(subject: str, body: str, attachments=None):
    msg = EmailMessage()
    msg["From"] = CEO_EMAIL
    msg["To"] = CEO_EMAIL
    msg["Subject"] = subject
    msg.set_content(body)

    for path in attachments or []:
        path = Path(path)
        if not path.exists():
            print(f"[resume_agent] warning: attachment not found, skipping: {path}")
            continue
        data = path.read_bytes()
        msg.add_attachment(data, maintype="application", subtype="octet-stream", filename=path.name)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(CEO_EMAIL, EMAIL_APP_PASSWORD)
        smtp.send_message(msg)


def deliver_resume_agent_output(draft: dict, kind: str, output_format: str = "text", delivery: str = "none") -> dict:
    """
    kind: "cover_letter" or "resume_changes"
    output_format: "text" (no file), "docx", or "pdf"
    delivery: "none" (text only), "save" (Agent output folder), "mail"
              (send only — uses a real OS temp file, deleted right after,
              never touches state/), "mail+save" (both)

    This is the only place a file gets written for Resume Agent output, and
    only when explicitly asked — matching 'nothing saved except in excel...
    no state' from the spec.
    """
    if output_format == "text" and delivery == "none":
        return {"text": draft["text"], "file": None}

    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    suffix = "CoverLetter" if kind == "cover_letter" else "ResumeChanges"
    base_name = f"{_safe_filename_part(draft['company'])}_{_safe_filename_part(draft['title'])}_{timestamp}_{suffix}"
    ext = "docx" if output_format == "docx" else "pdf"

    if delivery in ("save", "mail+save"):
        out_path = AGENT_OUTPUT_DIR / f"{base_name}.{ext}"
    else:
        out_path = Path(tempfile.gettempdir()) / f"{base_name}.{ext}"

    if output_format == "docx":
        compile_docx(draft["text"], out_path)
    elif output_format == "pdf":
        compile_pdf(draft["text"], out_path)

    result = {"text": draft["text"], "file": out_path if output_format != "text" else None}

    if delivery in ("mail", "mail+save"):
        subject = f"{'Cover Letter' if kind == 'cover_letter' else 'Resume Changes'} — {draft['title']} @ {draft['company']}"
        body_preview = draft["text"][:500] + ("..." if len(draft["text"]) > 500 else "")
        try:
            send_email(subject, body_preview, attachments=[out_path] if output_format != "text" else [])
            result["emailed"] = True
        except Exception as e:
            print(f"[resume_agent] email failed: {e}")
            result["emailed"] = False

    if delivery == "mail" and result["file"] and result["file"].exists():
        result["file"].unlink()  # ephemeral — only existed to attach to the email
        result["file"] = None

    return result

# =============================================================================
# JOB-E ORCHESTRATOR
# =============================================================================

JOB_E_SKILL_PATH = SKILLS_DIR / "job-e-SKILL.md"

JOB_E_TOOLS = [
    {
        "name": "query_tracker",
        "description": (
            "Reads rows from the JobTracker Excel, optionally filtered. Use this "
            "to answer questions about what's in the tracker, find rows to act on, "
            "or check completeness/quality before deciding what to do next. Returns "
            "row numbers you can pass to run_talent_agent / run_resume_agent / log_application. "
            "Every row includes date_posted and timestamp_captured (when the row was added to "
            "the tracker) — use captured_since for 'today' / 'this week' style questions rather "
            "than guessing from row numbers."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "company": {"type": "string", "description": "Filter to postings at this company (substring match, case-insensitive)"},
                "high_match_only": {"type": "boolean", "description": "Only rows marked High Match"},
                "min_score": {"type": "integer", "description": "Only rows with Match Score >= this value"},
                "missing_jd_summary": {"type": "boolean", "description": "Only rows with no Job Description Summary yet"},
                "missing_score": {"type": "boolean", "description": "Only rows with no Match Score yet"},
                "captured_since": {"type": "string", "description": "ISO date (YYYY-MM-DD) — only rows captured on or after this date. Use today's date for 'what's new today'."},
                "limit": {"type": "integer", "description": "Max rows to return, default 500 (comfortably above realistic tracker size — usually no need to raise this)"},
            },
        },
    },
    {
        "name": "run_talent_agent",
        "description": (
            "Enriches and scores specific rows (or, if row_nums is omitted, every "
            "High Match row without a score yet). Writes Date Posted, Open or "
            "Closed, Hiring Team, Pay, Job Description Summary, Match Score, "
            "Strong Points, Weak Points directly to the tracker."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "row_nums": {"type": "array", "items": {"type": "integer"}, "description": "Specific row numbers, or omit for bulk (all pending High Match rows)"},
            },
        },
    },
    {
        "name": "run_resume_agent",
        "description": (
            "Drafts a cover letter or resume changes for ONE specific row, only "
            "when explicitly requested. Optionally compiles as docx/pdf and "
            "delivers by email and/or saves to the Agent output folder."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "row_num": {"type": "integer", "description": "The JobTracker row to draft for"},
                "kind": {"type": "string", "enum": ["cover_letter", "resume_changes"]},
                "output_format": {"type": "string", "enum": ["text", "docx", "pdf"], "description": "Default text — ask the person if unspecified"},
                "delivery": {"type": "string", "enum": ["none", "save", "mail", "mail+save"], "description": "Default none — ask the person if they want to keep/send it"},
            },
            "required": ["row_num", "kind"],
        },
    },
    {
        "name": "log_application",
        "description": "Records that the person actually applied to a posting, in the Applications & Outcomes tab. Only call this when the person explicitly says they applied — never infer it from a cover letter being drafted.",
        "input_schema": {
            "type": "object",
            "properties": {
                "row_num": {"type": "integer", "description": "The JobTracker row this application is for"},
                "method": {"type": "string", "description": "e.g. 'Company site', 'Email', 'LinkedIn Easy Apply', 'Referral'"},
                "outcome": {"type": "string", "enum": ["Applied", "Interview", "Rejected", "Offer", "No Response"]},
                "notes": {"type": "string"},
            },
            "required": ["row_num"],
        },
    },
    {
        "name": "run_search_agent",
        "description": (
            "Runs a full Gmail sync: parses new LinkedIn/Indeed job alerts into the "
            "tracker, re-labels High Match across every row, and scans Gmail for "
            "application confirmation/rejection/interview/offer emails to update "
            "the Applications & Outcomes tab. This is the ONLY way new data enters "
            "the tracker from Gmail — use it when asked to check for new postings, "
            "sync email, or check for application updates. By default the "
            "application-scan is INCREMENTAL (since the last successful scan, not "
            "a fixed window) — cheap for daily use, no need to specify anything for "
            "a routine sync. Only pass applications_lookback_days when Raj "
            "explicitly asks to look back further, catch up after missing several "
            "days, or backfill older history — e.g. 'sync and look back 60 days'."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "applications_lookback_days": {"type": "integer", "description": "Only set this when Raj explicitly requests a specific lookback window (e.g. 30, 60, 180). Omit entirely for a normal day-to-day sync — the default incremental behavior (since last scan) is what keeps token cost low."},
            },
        },
    },
    {
        "name": "query_applications",
        "description": (
            "Reads rows from the Applications & Outcomes tab — actual applications "
            "logged (auto-detected from Gmail via run_search_agent, or manually via "
            "log_application), with their outcomes. Use this to answer 'what have I "
            "applied to,' 'what's my response rate,' 'is the application tracker "
            "populated,' or to check for an existing entry before logging a new one."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "company": {"type": "string", "description": "Filter by company (substring match)"},
                "outcome": {"type": "string", "enum": ["Applied", "Interview", "Rejected", "Offer", "No Response"]},
                "limit": {"type": "integer", "description": "Max rows to return, default 500 (comfortably above realistic tracker size — usually no need to raise this)"},
            },
        },
    },
    {
        "name": "read_context_hub",
        "description": (
            "Reads Raj's resume, proof points, positioning rules, and any other reference "
            "material in the Context Hub. Use this for general questions about his background "
            "or experience that aren't about a specific tracker posting — 'summarize my resume,' "
            "'what's my strongest proof point for X,' 'what does my positioning say about Y.' "
            "For questions about how he matches a SPECIFIC posting, prefer query_tracker's "
            "strong_points/weak_points instead — those are already scored against a real JD."
        ),
        "input_schema": {"type": "object", "properties": {}},
    },
    {"type": "web_search_20250305", "name": "web_search", "max_uses": 3},
]


def _tool_query_tracker(company=None, high_match_only=False, min_score=None,
                         missing_jd_summary=False, missing_score=False,
                         captured_since=None, limit=500):
    wb, ws = open_job_tracker()
    cols = {name: 2 + JOB_TRACKER_COLUMNS.index(name) for name in JOB_TRACKER_COLUMNS}

    since_dt = None
    if captured_since:
        try:
            since_dt = datetime.strptime(captured_since, "%Y-%m-%d")
        except ValueError:
            pass  # ignore a malformed date rather than erroring the whole query

    rows = []
    for r in range(2, ws.max_row + 1):
        title = ws.cell(row=r, column=cols["Job Title"]).value
        if not title:
            continue
        row_company = ws.cell(row=r, column=cols["Company"]).value
        if company and (not row_company or company.lower() not in row_company.lower()):
            continue
        if high_match_only and ws.cell(row=r, column=cols["High Match"]).value != "High Match":
            continue
        score = ws.cell(row=r, column=cols["Match Score"]).value
        if min_score is not None and (score is None or score < min_score):
            continue
        if missing_jd_summary and ws.cell(row=r, column=cols["Job Description Summary"]).value:
            continue
        if missing_score and score is not None:
            continue
        if since_dt is not None:
            captured = ws.cell(row=r, column=cols["Timestamp Captured"]).value
            if not captured or (hasattr(captured, "date") and captured < since_dt):
                continue

        rows.append({
            "row": r,
            "title": title, "company": row_company,
            "location": ws.cell(row=r, column=cols["Location"]).value,
            "high_match": ws.cell(row=r, column=cols["High Match"]).value,
            "match_score": score,
            "strong_points": ws.cell(row=r, column=cols["Strong Points"]).value,
            "weak_points": ws.cell(row=r, column=cols["Weak Points"]).value,
            "jd_summary": ws.cell(row=r, column=cols["Job Description Summary"]).value,
            "pay": ws.cell(row=r, column=cols["Pay / CTC / Salary"]).value,
            "status": ws.cell(row=r, column=cols["Open or Closed"]).value,
            "date_posted": ws.cell(row=r, column=cols["Date Posted"]).value,
            "timestamp_captured": str(ws.cell(row=r, column=cols["Timestamp Captured"]).value) if ws.cell(row=r, column=cols["Timestamp Captured"]).value else None,
            "cover_created": ws.cell(row=r, column=cols["Cover Created"]).value,
            "resume_updated": ws.cell(row=r, column=cols["Resume Updated"]).value,
            "url": ws.cell(row=r, column=cols["Job URL"]).value,
        })
        if len(rows) >= limit:
            break

    return rows


def _tool_run_talent_agent(row_nums=None):
    results = talent_agent_run(row_nums)
    output = []
    for r in results:
        score_result = r["score"]
        if score_result and score_result.get("error") is None:
            output.append({"row": r["row"], "title": r["title"], "company": r["company"],
                            "score": score_result["match_score"], "scoring_error": None})
        else:
            # Surface the real failure reason instead of just "no score" — this
            # is exactly what was invisible before: a row could fail scoring
            # entirely and the only trace was a print() in a terminal nobody
            # was watching. Now it's in the tool result JOB-e actually sees.
            output.append({"row": r["row"], "title": r["title"], "company": r["company"],
                            "score": None, "scoring_error": (score_result or {}).get("error", "Unknown failure")})
    return output


def _tool_run_resume_agent(row_num, kind, output_format="text", delivery="none"):
    """
    App-specific override of the notebook version: ALWAYS drafts only,
    never delivers — output_format/delivery arguments from JOB-e are
    ignored on purpose. Delivery is a UI action now (the Draft Review tab's
    Approve & Deliver button), not something decided through conversation,
    since a real review-and-edit step exists here that the notebook chat
    never had. The draft goes into st.session_state for that tab to pick
    up; JOB-e gets back a short status, not the full text, to keep the
    conversation from ballooning with content that's about to be shown
    properly anyway.
    """
    if kind == "cover_letter":
        draft = draft_cover_letter(row_num)
    else:
        draft = draft_resume_changes(row_num)

    st.session_state.pending_draft = {
        "row": row_num, "kind": kind, "company": draft["company"],
        "title": draft["title"], "text": draft["text"],
        "critique": draft.get("critique"),
    }

    return {
        "row": row_num, "company": draft["company"], "title": draft["title"],
        "status": "drafted — ready for review in the Draft Review tab",
    }


def _tool_log_application(row_num, method=None, outcome="Applied", notes=None):
    wb, ws = open_job_tracker()
    cols = {name: 2 + JOB_TRACKER_COLUMNS.index(name) for name in JOB_TRACKER_COLUMNS}
    company = ws.cell(row=row_num, column=cols["Company"]).value
    title = ws.cell(row=row_num, column=cols["Job Title"]).value
    serial = ws.cell(row=row_num, column=1).value
    row_added = add_application_outcome(serial, company, title, method=method, outcome=outcome, notes=notes)
    return {"row": row_added, "company": company, "title": title, "outcome": outcome}


def _tool_run_search_agent(applications_lookback_days: int = None):
    path = search_agent_run(applications_lookback_days=applications_lookback_days)
    return {"status": "synced", "tracker_path": str(path),
            "applications_lookback_days": applications_lookback_days or "incremental (since last scan)"}


def _tool_query_applications(company=None, outcome=None, limit=500):
    wb = openpyxl.load_workbook(JOB_TRACKER_PATH)
    if APPLICATIONS_SHEET not in wb.sheetnames:
        return []
    ws = wb[APPLICATIONS_SHEET]
    rows = []
    for r in range(2, ws.max_row + 1):
        c = ws.cell(row=r, column=2).value
        if not c:
            continue
        if company and company.lower() not in c.lower():
            continue
        row_outcome = ws.cell(row=r, column=6).value
        if outcome and row_outcome != outcome:
            continue
        rows.append({
            "row": r,
            "serial": ws.cell(row=r, column=1).value,
            "company": c,
            "role": ws.cell(row=r, column=3).value,
            "date_applied": str(ws.cell(row=r, column=4).value) if ws.cell(row=r, column=4).value else None,
            "method": ws.cell(row=r, column=5).value,
            "outcome": row_outcome,
            "outcome_date": str(ws.cell(row=r, column=7).value) if ws.cell(row=r, column=7).value else None,
            "notes": ws.cell(row=r, column=8).value,
        })
        if len(rows) >= limit:
            break
    return rows


def _tool_read_context_hub():
    blocks = load_context_hub()
    return {"context_hub_content": "\n\n".join(blocks)[:15000]}  # capped — this is a direct read, not scoring


_JOB_E_TOOL_DISPATCH = {
    "query_tracker": _tool_query_tracker,
    "run_search_agent": _tool_run_search_agent,
    "query_applications": _tool_query_applications,
    "run_talent_agent": _tool_run_talent_agent,
    "run_resume_agent": _tool_run_resume_agent,
    "log_application": _tool_log_application,
    "read_context_hub": _tool_read_context_hub,
}


def chat_with_jobe(user_message: str, history: list = None) -> tuple:
    """
    One turn of conversation with JOB-e. Runs the full tool-use loop
    (JOB-e can call multiple tools in sequence before answering) and
    returns (response_text, updated_history) — pass the history back in
    on the next call to keep the conversation going.
    """
    history = list(history or [])
    history.append({"role": "user", "content": user_message})

    system_prompt = load_skill(JOB_E_SKILL_PATH)
    current_threshold = st.session_state.get("fitment_threshold", FITMENT_THRESHOLD_PERCENT)
    system_prompt += (
        f"\n\nCurrent fitment threshold (set live via the sidebar slider): "
        f"{current_threshold}%. Use this actual number when discussing whether a "
        f"posting clears the bar, flagging low scores, or describing what counts "
        f"as a good match right now — it changes as Raj adjusts the slider."
    )

    for _ in range(8):  # hard cap on tool-call rounds per turn, avoid runaway loops
        response = client.messages.create(
            model=MODEL_ROUTING["job_e"],
            max_tokens=2000,
            system=system_prompt,
            tools=JOB_E_TOOLS,
            messages=history,
        )

        history.append({"role": "assistant", "content": response.content})

        # Track cumulative session token usage for the sidebar's cost estimate.
        if hasattr(response, "usage") and response.usage:
            used = (response.usage.input_tokens or 0) + (response.usage.output_tokens or 0)
            st.session_state["session_tokens"] = st.session_state.get("session_tokens", 0) + used

        if response.stop_reason != "tool_use":
            final_text = "".join(b.text for b in response.content if b.type == "text")
            return final_text, history

        tool_results = []
        for block in response.content:
            if block.type != "tool_use":
                continue
            fn = _JOB_E_TOOL_DISPATCH.get(block.name)
            if fn is None:
                # web_search is handled server-side by Anthropic — nothing to dispatch locally
                continue
            try:
                result = fn(**block.input)
                content = json.dumps(result, default=str)
            except Exception as e:
                content = json.dumps({"error": str(e)})
            tool_results.append({"type": "tool_result", "tool_use_id": block.id, "content": content})

        if tool_results:
            history.append({"role": "user", "content": tool_results})
        else:
            # Only web_search blocks were used — Anthropic already resolved those
            # server-side, so the next response.content will include the answer.
            continue

    return "(stopped after several tool-call rounds — try a narrower request)", history

# =============================================================================
# DASHBOARD DATA (rendering is Streamlit-native, in the UI section below)
# =============================================================================

def compute_dashboard_kpis():
    wb, ws = open_job_tracker()
    cols = {name: 2 + JOB_TRACKER_COLUMNS.index(name) for name in JOB_TRACKER_COLUMNS}

    total = high_match = scored = cover_created = resume_updated = 0
    scores = []
    for r in range(2, ws.max_row + 1):
        if not ws.cell(row=r, column=cols["Job Title"]).value:
            continue
        total += 1
        if ws.cell(row=r, column=cols["High Match"]).value == "High Match":
            high_match += 1
        score = ws.cell(row=r, column=cols["Match Score"]).value
        if score not in (None, ""):
            scored += 1
            try:
                scores.append(float(score))
            except (TypeError, ValueError):
                pass
        if ws.cell(row=r, column=cols["Cover Created"]).value == "Yes":
            cover_created += 1
        if ws.cell(row=r, column=cols["Resume Updated"]).value == "Yes":
            resume_updated += 1

    avg_score = sum(scores) / len(scores) if scores else 0

    applications, response_rate = 0, 0
    if APPLICATIONS_SHEET in wb.sheetnames:
        app_ws = wb[APPLICATIONS_SHEET]
        outcomes = []
        for r in range(2, app_ws.max_row + 1):
            if not app_ws.cell(row=r, column=2).value:
                continue
            applications += 1
            outcomes.append(app_ws.cell(row=r, column=6).value)
        positive = sum(1 for o in outcomes if o in ("Interview", "Offer"))
        response_rate = (positive / applications * 100) if applications else 0

    return {
        "total": total, "high_match": high_match, "scored": scored,
        "cover_created": cover_created, "resume_updated": resume_updated,
        "avg_score": avg_score, "applications": applications, "response_rate": response_rate,
    }


# =============================================================================
# STREAMLIT UI
# =============================================================================

def extract_display_history(history):
    """
    Turns the raw Anthropic message history (tool_use/tool_result blocks
    included) into (role, text) pairs for rendering. Tool results are
    skipped entirely — internal plumbing, not conversation. Tool calls
    render as a small muted status line rather than raw JSON.
    """
    display_turns = []
    for msg in history:
        role = msg["role"]
        content = msg["content"]

        if isinstance(content, str):
            display_turns.append((role, content))
            continue

        for block in content:
            if isinstance(block, dict):
                block_type = block.get("type")
                if block_type == "tool_result":
                    continue
                block_text = block.get("text", "")
                block_name = block.get("name", "")
            else:
                block_type = getattr(block, "type", None)
                block_text = getattr(block, "text", "")
                block_name = getattr(block, "name", "")

            if block_type == "text" and block_text.strip():
                display_turns.append((role, block_text))
            elif block_type == "tool_use":
                display_turns.append(("tool", f"using {block_name} ..."))

    return display_turns


# --- Session state -----------------------------------------------------------
if "history" not in st.session_state:
    st.session_state.history = []
if "pending_draft" not in st.session_state:
    st.session_state.pending_draft = None
if "_pending_prompt" not in st.session_state:
    st.session_state["_pending_prompt"] = None
if "session_tokens" not in st.session_state:
    st.session_state["session_tokens"] = 0

validate_secrets()


def _get_kpis_safe():
    """KPIs used by both the sidebar status grid and the Dashboard tab —
    computed once per rerun, never hardcoded."""
    try:
        return compute_dashboard_kpis()
    except Exception:
        return None


# =============================================================================
# SIDEBAR — status grid, sync, session info. Matches TimberLens's sidebar
# pattern, adapted: no per-agent multiselect, since JOB-e is the sole
# conversational interface by explicit design (no per-agent chat mode here).
# =============================================================================
with st.sidebar:
    # Sidebar-specific button override — Streamlit wraps button labels in a
    # nested <p> tag that doesn't inherit color from the parent <button> the
    # way the general .stButton>button rule above assumes. TimberLens's own
    # code flags this exact gotcha as a "critical fix" — without forcing the
    # inner <p> color too, sidebar button text renders low-contrast/invisible
    # even though the outer button correctly shows the blue background.
    st.markdown(
        """
        <style>
        [data-testid="stSidebar"] .stButton > button {
            background-color: #0b57d0 !important;
            color: #ffffff !important;
            border: 1px solid #0b57d0 !important;
            border-radius: 8px !important;
            font-weight: 500 !important;
            transition: all 0.25s ease !important;
        }
        [data-testid="stSidebar"] .stButton > button p {
            color: #ffffff !important;
        }
        [data-testid="stSidebar"] .stButton > button:hover {
            background-color: #0043b6 !important;
            border-color: #0043b6 !important;
            box-shadow: 0px 4px 8px rgba(11, 87, 208, 0.2) !important;
        }
        [data-testid="stSidebar"] .stButton > button:hover p {
            color: #ffffff !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("### Raj JSAW")
    st.markdown("<small style='color:#8a7560'>JOB-e Command Center · v1.0</small>", unsafe_allow_html=True)
    st.markdown("---")

    with st.expander("⚙️ Setup & Paths", expanded=not JOB_TRACKER_PATH.exists()):
        st.markdown("<small style='color:#8a7560'>For a new machine or a new user — fill in paths, or paste credentials directly below instead of pointing at existing files. Saves automatically.</small>", unsafe_allow_html=True)

        st.markdown("**Folders**")
        cfg_base_dir = st.text_input("Project folder", value=str(BASE_DIR))
        cfg_context_hub_dir = st.text_input("Resume / Context folder", value=str(CONTEXT_HUB_DIR),
                                             help="Where your resume, proof points, and reference docs live.")
        cfg_job_alerts_dir = st.text_input("Job Alerts folder", value=str(JOB_ALERTS_DIR),
                                            help="Where the tracker Excel file lives.")
        cfg_tracker_filename = st.text_input("Tracker filename (inside Job Alerts folder)",
                                              value=_app_cfg.get("tracker_filename", "Job Search-SOR.xlsx"))
        cfg_agent_output_dir = st.text_input("Delivered documents folder", value=str(AGENT_OUTPUT_DIR),
                                              help="Where approved cover letters/resume changes get saved.")

        st.markdown("---")
        st.markdown("**Anthropic API key**")
        key_method = st.radio("Provide via", ["File path", "Paste key directly"], horizontal=True, key="api_key_method")
        if key_method == "File path":
            cfg_api_key_path = st.text_input("API key file (.rtf)", value=str(ANTHROPIC_API_KEY_PATH))
            cfg_api_key_direct = ""
        else:
            cfg_api_key_direct = st.text_input("API key", value=_app_cfg.get("api_key_direct", ""), type="password",
                                                placeholder="sk-ant-...")
            cfg_api_key_path = _app_cfg.get("api_key_path", str(ANTHROPIC_API_KEY_PATH))

        st.markdown("---")
        st.markdown("**Gmail**")
        gmail_method = st.radio("Provide via", ["G-Mail.env file", "Enter directly"], horizontal=True, key="gmail_method")
        if gmail_method == "Enter directly":
            cfg_ceo_email = st.text_input("Gmail address", value=_app_cfg.get("ceo_email", ""))
            cfg_email_app_password = st.text_input("Gmail App Password", value=_app_cfg.get("email_app_password", ""),
                                                     type="password",
                                                     help="Not your regular Gmail password — generate one at myaccount.google.com/apppasswords")
        else:
            cfg_ceo_email, cfg_email_app_password = "", ""

        st.caption("⚠️ Credentials entered here are saved in plain text to a local config file on this machine — same sensitivity as any saved password, not encrypted.")

        if st.button("💾 Save & Reload", use_container_width=True):
            new_cfg = {
                "base_dir": cfg_base_dir,
                "context_hub_dir": cfg_context_hub_dir,
                "job_alerts_dir": cfg_job_alerts_dir,
                "tracker_filename": cfg_tracker_filename,
                "agent_output_dir": cfg_agent_output_dir,
                "api_key_path": cfg_api_key_path,
                "api_key_direct": cfg_api_key_direct,
                "ceo_email": cfg_ceo_email,
                "email_app_password": cfg_email_app_password,
            }
            save_app_config(new_cfg)
            st.success("Saved — reloading...")
            st.rerun()

    st.markdown("---")

    st.markdown("**System Status**")
    kpis_for_status = _get_kpis_safe()
    api_ok = bool(ANTHROPIC_API_KEY)
    excel_ok = JOB_TRACKER_PATH.exists()
    gmail_ok = bool(CEO_EMAIL and EMAIL_APP_PASSWORD)

    grid_html = '<div class="status-grid">'
    grid_html += f'<div class="status-cell {"status-pass" if api_ok else "status-fail"}">API</div>'
    grid_html += f'<div class="status-cell {"status-pass" if excel_ok else "status-fail"}">Excel</div>'
    grid_html += f'<div class="status-cell {"status-pass" if gmail_ok else "status-fail"}">Gmail</div>'
    grid_html += "</div>"
    st.markdown(grid_html, unsafe_allow_html=True)

    if not excel_ok:
        st.caption(f"⚠️ Tracker not found at:\n{JOB_TRACKER_PATH}")

    st.markdown("---")

    if st.button("🔄 Sync Gmail Now", use_container_width=True):
        with st.spinner("Syncing Gmail — new postings, High Match labeling, application outcomes..."):
            try:
                _tool_run_search_agent()
                st.success("Sync complete.")
            except Exception as e:
                st.error(f"Sync failed: {e}")
        st.rerun()

    st.markdown("---")
    if kpis_for_status:
        st.caption(f"📋 {kpis_for_status['total']} postings · {kpis_for_status['high_match']} high match")
        st.caption(f"✉️ {kpis_for_status['applications']} applications tracked")

    st.markdown("---")
    st.markdown("**Match Threshold**")
    _current_threshold = st.session_state.get("fitment_threshold", FITMENT_THRESHOLD_PERCENT)
    _new_threshold = st.slider(
        "Minimum match % JOB-e treats as a good fit", 0, 100, value=_current_threshold, step=5,
        label_visibility="collapsed",
    )
    st.caption(f"Currently {_new_threshold}% — JOB-e uses this live when discussing fit.")
    if _new_threshold != _current_threshold:
        st.session_state["fitment_threshold"] = _new_threshold
        _updated_cfg = dict(_app_cfg)
        _updated_cfg["fitment_threshold"] = _new_threshold
        save_app_config(_updated_cfg)

    if kpis_for_status:
        try:
            wb_thresh, ws_thresh = open_job_tracker()
            score_col_idx = 2 + JOB_TRACKER_COLUMNS.index("Match Score")
            above_count = sum(
                1 for r in range(2, ws_thresh.max_row + 1)
                if isinstance(ws_thresh.cell(row=r, column=score_col_idx).value, (int, float))
                and ws_thresh.cell(row=r, column=score_col_idx).value >= _new_threshold
            )
            st.caption(f"🎯 {above_count} scored posting(s) currently clear this bar")
        except Exception:
            pass

    st.markdown("---")
    st.markdown("**Match Keywords**")
    st.caption("Focus areas JOB-e and Talent Agent weigh when scoring — not a filter, just extra signal.")
    _match_keywords = list(_app_cfg.get("match_keywords", ["Strategy", "Business Development", "Operations", "ISVs", "Data Science"]))

    for _kw in list(_match_keywords):
        kcol1, kcol2 = st.columns([4, 1])
        kcol1.markdown(f"<span class='agent-chip'>{_kw}</span>", unsafe_allow_html=True)
        if kcol2.button("✕", key=f"remove_kw_{_kw}"):
            _match_keywords.remove(_kw)
            _updated_cfg = dict(_app_cfg)
            _updated_cfg["match_keywords"] = _match_keywords
            save_app_config(_updated_cfg)
            st.rerun()

    new_kw = st.text_input("Add a keyword", key="new_match_keyword", label_visibility="collapsed", placeholder="e.g. Cloud Infrastructure")
    if st.button("➕ Add keyword", use_container_width=True) and new_kw.strip():
        if new_kw.strip() not in _match_keywords:
            _match_keywords.append(new_kw.strip())
            _updated_cfg = dict(_app_cfg)
            _updated_cfg["match_keywords"] = _match_keywords
            save_app_config(_updated_cfg)
        st.rerun()

    st.markdown("---")
    total_tok = st.session_state.get("session_tokens", 0)
    # Rough blended Haiku/Sonnet estimate, same approximation TimberLens uses —
    # not exact, just a directional sense of session cost.
    est_cost = (total_tok / 1_000_000) * 2.0
    st.caption(f"📊 Session tokens: {total_tok:,}")
    st.caption(f"💰 Est. cost: ${est_cost:.4f}")

    st.markdown("---")
    if st.button("🗑️ Clear chat", use_container_width=True):
        st.session_state.history = []
        st.session_state.pending_draft = None
        st.rerun()

    st.markdown("---")
    with st.expander("📜 Session history", expanded=False):
        display_turns = extract_display_history(st.session_state.history)
        user_turns = [t for r, t in display_turns if r == "user"]
        if user_turns:
            for i, t in enumerate(user_turns, 1):
                snippet = t[:35] + "..." if len(t) > 35 else t
                st.markdown(f"*{i}. {snippet}*")
        else:
            st.caption("No conversation yet this session.")


# =============================================================================
# MAIN TABS
# =============================================================================
tab_chat, tab_drafts, tab_dashboard, tab_help = st.tabs(["ORCHESTRATOR", "DRAFT REVIEW", "DASHBOARD", "ABOUT & HELP"])


# --- Chat tab: text_area + explicit send button, matching TimberLens's ------
# actual input pattern — not a chat-bubble/chat_input widget.
with tab_chat:
    # Handle a Quick Command click from the previous rerun
    if st.session_state["_pending_prompt"]:
        prompt = st.session_state["_pending_prompt"]
        st.session_state["_pending_prompt"] = None
        with st.spinner(f"Working: {prompt}"):
            response_text, st.session_state.history = chat_with_jobe(prompt, history=st.session_state.history)
        if st.session_state.pending_draft:
            st.info("📝 A new draft is ready — see the Draft Review tab.")
        st.rerun()

    st.markdown("**Quick Commands**")
    st.caption("In workflow order. Sync / Score / Stats run directly — no chat round trip, no extra token cost.")
    qc1, qc2, qc3, qc4 = st.columns(4)

    with qc1:
        if st.button("🔄 Sync Mail", use_container_width=True,
                      help="Pull new job alerts and application-status updates from Gmail into the tracker. Runs directly, no LLM cost for the sync itself."):
            with st.spinner("Syncing Gmail..."):
                try:
                    result = _tool_run_search_agent()
                    st.success(f"Synced. Lookback: {result['applications_lookback_days']}.")
                except Exception as e:
                    st.error(f"Sync failed: {e}")

    with qc2:
        if st.button("🎯 Score High Matches", use_container_width=True,
                      help="Run Talent Agent on any High Match postings not yet scored. Uses the LLM per posting — real work, real cost, but no chat overhead on top."):
            with st.spinner("Scoring pending High Match postings..."):
                try:
                    results = talent_agent_run()
                    st.success(f"Scored {len(results)} posting(s).")
                except Exception as e:
                    st.error(f"Scoring failed: {e}")

    with qc3:
        if st.button("📊 Quick Stats", use_container_width=True,
                      help="Current tracker totals, computed directly from Excel. Zero token cost."):
            kpis_qc = _get_kpis_safe()
            if kpis_qc:
                st.info(
                    f"**{kpis_qc['total']}** postings · **{kpis_qc['high_match']}** high match · "
                    f"**{kpis_qc['scored']}** scored (avg {kpis_qc['avg_score']:.0f}%) · "
                    f"**{kpis_qc['applications']}** applications · **{kpis_qc['response_rate']:.0f}%** response rate"
                )
            else:
                st.warning("Could not load the tracker.")

    with qc4:
        if st.button("🧭 What Needs Attention", use_container_width=True,
                      help="Ask JOB-e to review the tracker for gaps or issues worth a look. This one does use the LLM — it requires actual judgment."):
            st.session_state["_pending_prompt"] = "Review the tracker for anything incomplete or low-quality and flag what needs attention."
            st.rerun()

    st.markdown("---")

    for role, text in extract_display_history(st.session_state.history):
        if role == "tool":
            st.markdown(f'<p class="status-line">🔧 {text}</p>', unsafe_allow_html=True)
        elif role == "user":
            st.markdown(f'<div class="user-box">🧑‍💼 <b>You</b><br>{text}</div>', unsafe_allow_html=True)
        else:
            st.markdown(f'<div class="report-box">💼 <b>JOB-e</b><br>{text}</div>', unsafe_allow_html=True)

    user_input = st.text_area(
        "Ask JOB-e...",
        placeholder='e.g. "Draft a cover letter for the Atos posting" or "What roles came in from Google this month?"',
        height=90,
        label_visibility="collapsed",
        key="main_chat_input",
    )

    def _handle_send_click():
        """
        Runs as a button on_click callback, which Streamlit executes in a
        separate phase BEFORE the next script rerun — this is the actual
        sanctioned way to clear a keyed widget's value. Setting
        st.session_state["main_chat_input"] anywhere in the normal
        top-to-bottom script body raises StreamlitAPIException once that
        widget has been instantiated in the current run, even if the write
        happens right before st.rerun() — confirmed by hitting this exact
        error live. A callback is a genuinely different execution phase,
        not just "later in the same run," so it's allowed here.
        """
        msg = st.session_state.get("main_chat_input", "").strip()
        if msg:
            st.session_state["_pending_prompt"] = msg
            st.session_state["main_chat_input"] = ""

    st.button("▶ Send to JOB-e", use_container_width=False, on_click=_handle_send_click)


# --- Draft Review tab ----------------------------------------------------------
with tab_drafts:
    draft = st.session_state.pending_draft

    if not draft:
        st.info("No draft pending. Ask JOB-e to draft a cover letter or resume changes for a specific posting.")
    else:
        st.markdown(f"### {draft['kind'].replace('_', ' ').title()} — {draft['title']} @ {draft['company']}")

        if draft.get("critique"):
            with st.expander("Company Research critique (already incorporated into this revision)"):
                st.markdown(f'<div class="report-box">{draft["critique"]}</div>', unsafe_allow_html=True)

        draft_id = f"{draft['row']}_{draft['kind']}_{hash(draft['text'])}"
        if st.session_state.get("_shown_draft_id") != draft_id:
            st.session_state["draft_editor"] = draft["text"]
            st.session_state["_shown_draft_id"] = draft_id

        edited_text = st.text_area("Draft (editable — your edits are what gets delivered)", height=380, key="draft_editor")

        col1, col2 = st.columns(2)
        with col1:
            output_format = st.selectbox("Format", ["text", "docx", "pdf"], key="draft_format")
        with col2:
            delivery = st.selectbox("Delivery", ["none", "save", "mail", "mail+save"], key="draft_delivery")

        b1, b2 = st.columns(2)
        with b1:
            if st.button("✅ Approve & Deliver", use_container_width=True):
                final_draft = dict(draft)
                final_draft["text"] = edited_text
                with st.spinner("Compiling and delivering..."):
                    result = deliver_resume_agent_output(
                        final_draft, kind=draft["kind"], output_format=output_format, delivery=delivery
                    )
                msg_parts = []
                if result.get("file"):
                    msg_parts.append(f"saved as `{Path(result['file']).name}`")
                if result.get("emailed"):
                    msg_parts.append("emailed")
                st.success("Delivered" + (f" — {', '.join(msg_parts)}." if msg_parts else " (text only, nothing saved)."))
                st.session_state.pending_draft = None
                st.session_state.pop("_shown_draft_id", None)
                st.rerun()
        with b2:
            if st.button("🗑️ Discard", use_container_width=True):
                st.session_state.pending_draft = None
                st.session_state.pop("_shown_draft_id", None)
                st.rerun()


# --- Dashboard tab --------------------------------------------------------------
with tab_dashboard:
    st.markdown("### 📊 Job Search Dashboard")

    kpis = _get_kpis_safe()
    if kpis is None:
        st.warning(f"Could not load the tracker at {JOB_TRACKER_PATH}. Check the path and try again.")
    else:
        row1 = [
            (kpis["total"], "Total Postings"), (kpis["high_match"], "High Match"),
            (kpis["scored"], "Scored"), (f"{kpis['avg_score']:.0f}%", "Avg Match Score"),
        ]
        cols = st.columns(4)
        for col, (val, lbl) in zip(cols, row1):
            with col:
                st.markdown(f'<div class="kpi-card"><div class="kpi-val">{val}</div><div class="kpi-lbl">{lbl}</div></div>', unsafe_allow_html=True)

        row2 = [
            (kpis["cover_created"], "Covers Created"), (kpis["resume_updated"], "Resumes Updated"),
            (kpis["applications"], "Applications Sent"), (f"{kpis['response_rate']:.0f}%", "Response Rate"),
        ]
        cols2 = st.columns(4)
        for col, (val, lbl) in zip(cols2, row2):
            with col:
                st.markdown(f'<div class="kpi-card"><div class="kpi-val">{val}</div><div class="kpi-lbl">{lbl}</div></div>', unsafe_allow_html=True)

        st.markdown("---")
        sheet_choice = st.selectbox("📋 View sheet", ["Job Tracker", "Applications & Outcomes"])
        try:
            df = pd.read_excel(JOB_TRACKER_PATH, sheet_name=sheet_choice)
            st.dataframe(df.dropna(how="all").fillna(""), use_container_width=True, height=420)
        except Exception as e:
            st.error(f"Could not load sheet: {e}")


# --- About & Help tab -----------------------------------------------------
with tab_help:
    st.markdown("### What this app does")
    st.markdown("""
JOB-e is a single conversational assistant sitting on top of your job search tracker.
Behind the scenes, three specialist agents do the actual work — **Search Agent** (pulls
new postings and application updates from Gmail), **Talent Agent** (scores postings
against your resume), and **Resume Agent** (drafts cover letters / resume changes on
request) — but you never talk to them directly. You talk to JOB-e, and it decides what
needs to happen.
""")

    st.markdown("### The normal workflow")
    st.markdown("""
1. **Sync Mail** — pulls new job alerts and application-status updates from Gmail into the Excel tracker.
2. **Run Talent Agent** — scores whatever's newly marked High Match against your resume and proof points.
3. **Check the Dashboard** — see the numbers: totals, scores, applications, response rate.
4. **What needs attention** — ask JOB-e to review the tracker for gaps or postings worth a closer look.
5. **Draft, review, and send** — only when you're ready for a specific posting, never automatic.
""")

    st.markdown("---")
    st.markdown("### First-time setup")

    with st.expander("🔑 Anthropic API key"):
        st.markdown("""
1. Go to [console.anthropic.com](https://console.anthropic.com) and sign in.
2. Navigate to **API Keys** and create a new key.
3. Either save it into a `.rtf` file and point **Setup & Paths** at that file, or paste the key directly into the **Setup & Paths** panel in the sidebar — either works, direct paste takes priority if both are set.
""")

    with st.expander("✉️ Gmail App Password (required — your regular Gmail password will NOT work)"):
        st.markdown("""
1. Your Google Account needs **2-Step Verification** turned on first (Google Account → Security).
2. Go to [myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords).
3. Create a new App Password (name it anything, e.g. "Raj JSAW").
4. Google shows you a 16-character password **once** — copy it immediately.
5. Enter your Gmail address and that App Password into **Setup & Paths** in the sidebar.
6. Also confirm **IMAP is enabled**: Gmail Settings → See all settings → Forwarding and POP/IMAP → Enable IMAP.
""")

    with st.expander("📊 The Excel tracker template"):
        st.markdown("""
The tracker is a single Excel file with two sheets:

- **Job Tracker** — one row per posting: title, company, location, dates, match score, strong/weak points, and whether a cover letter or resume changes have been drafted for it.
- **Applications & Outcomes** — one row per application you've actually submitted, with its status (Applied / Interview / Rejected / Offer / No Response), populated both automatically (Gmail scan) and manually.

Point **Setup & Paths** at your project folder and the exact tracker filename — if the app can't find it, the sidebar's System Status will show Excel as failing, and this is almost always a filename mismatch (check for underscores vs. spaces).
""")

    st.markdown("---")
    st.markdown("### A note on cost")
    st.markdown("""
Every chat message, every scoring pass, and every draft costs real API tokens — visible
in the sidebar's Session tokens / Est. cost. A few habits keep this reasonable:

- Let **Sync Mail** use its default incremental behavior (since your last sync) rather than requesting a long lookback every time — only ask for a specific longer window when you're actually catching up after time away.
- Run Talent Agent on specific rows you care about rather than a full sweep, unless you genuinely want to score everything at once.
- The Dashboard tab's numbers are computed directly from the Excel file — checking it costs nothing.
""")
