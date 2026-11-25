"""
RAN KPI Analytics Tool  ·  v13  (Nokia 5G planning data + sector azimuth map)
═══════════════════════════════════════════════
Author  : Built for RF/RAN engineers
Purpose : Universal vendor-agnostic RAN KPI analytics engine
Vendors : Nokia (WCDMA/LTE/5G), Ericsson, Huawei, ZTE, Samsung
Formats : XLS, XLSX, CSV, TSV (auto-detect encoding, separator, header rows)
"""

# ─── Standard library ────────────────────────────────────────────────────────
import re
import json
import hashlib
import warnings
import subprocess
import tempfile
import os
import shutil
import zipfile
import glob
import xml.etree.ElementTree as ET
from io import BytesIO
from datetime import datetime, timedelta
from collections import Counter
from typing import Optional, Tuple

warnings.filterwarnings("ignore")

# ─── Third-party ─────────────────────────────────────────────────────────────
import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
import plotly.io as pio
from plotly.subplots import make_subplots
from sklearn.ensemble import IsolationForest
from sklearn.preprocessing import StandardScaler
from sklearn.linear_model import LinearRegression
import requests

# Optional dependencies
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors as RC
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.platypus import (SimpleDocTemplate, Paragraph,
                                     Spacer, Table, TableStyle, PageBreak)
    HAS_REPORTLAB = True
except ImportError:
    HAS_REPORTLAB = False

try:
    import ruptures as rpt
    HAS_RUPTURES = True
except ImportError:
    HAS_RUPTURES = False

try:
    import xlrd
    HAS_XLRD = True
except ImportError:
    HAS_XLRD = False

try:
    import openpyxl       # noqa: F401
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG  (must be first Streamlit call)
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="RAN KPI Analytics",
    page_icon="📡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────────────
# THEME — clean light design (soft neutrals, indigo primary, amber accent)
# ─────────────────────────────────────────────────────────────────────────────
PRIMARY      = "#4f46e5"   # indigo
PRIMARY_DARK = "#312e81"
PRIMARY_MID  = "#4338ca"
PRIMARY_SOFT = "#eef2ff"   # tint for chips / hover / tags
AMBER        = "#f59e0b"
INK          = "#1f2937"
MUTED        = "#6b7280"
LINE         = "#e5e7ee"
CANVAS       = "#f6f7fb"
# Back-compat aliases (older code paths)
TEAL, TEAL_DARK, TEAL_MID = PRIMARY, PRIMARY_DARK, PRIMARY_MID

pio.templates["rf"] = go.layout.Template(layout=dict(
    font=dict(family="Inter, -apple-system, Segoe UI, sans-serif", color=INK),
    colorway=["#4f46e5", "#f59e0b", "#10b981", "#ef4444", "#0ea5e9", "#8b5cf6", "#f97316", "#64748b"],
    title=dict(font=dict(size=15, color=PRIMARY_DARK)),
    xaxis=dict(gridcolor="#eef0f6", linecolor="#d7dbe6", zeroline=False),
    yaxis=dict(gridcolor="#eef0f6", linecolor="#d7dbe6", zeroline=False),
    hoverlabel=dict(font_size=12),
    margin=dict(l=50, r=20, t=55, b=45),
))
RF_TPL = "plotly_white+rf"
pio.templates.default = RF_TPL

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
html, body, [class*="css"] {{ font-family:'Inter',-apple-system,BlinkMacSystemFont,sans-serif; color:{INK}; }}
.stApp {{ background:{CANVAS}; }}
#MainMenu, footer {{ visibility:hidden; }}
.block-container {{ padding:0 1.5rem 2rem 1.5rem !important; max-width:100% !important; background:transparent; }}

/* ── Header ──────────────────────────────────────────────────────────────── */
.rf-header {{
    background:#fff; padding:18px 28px 14px;
    margin:-0.5rem -1.5rem 1.5rem -1.5rem;
    border-bottom:1px solid {LINE}; border-top:4px solid {PRIMARY};
    box-shadow:0 1px 3px rgba(17,24,39,.05);
}}
.rf-header h1 {{ color:{INK} !important; font-size:1.7rem !important; font-weight:800 !important;
    margin:0 0 3px 0 !important; letter-spacing:-0.02em; }}
.rf-header p {{ color:{MUTED} !important; font-size:0.82rem !important; margin:0 !important; }}
.rf-badge {{ background:{PRIMARY_SOFT}; color:{PRIMARY_MID}; font-size:.62rem; font-weight:700;
    padding:3px 9px; border-radius:999px; vertical-align:middle; margin-left:10px;
    letter-spacing:.06em; border:1px solid #c7d2fe; }}
.rf-chips {{ display:flex; flex-wrap:wrap; gap:8px; margin:0 0 14px 0; }}
.rf-chip {{ background:#fff; border:1px solid {LINE}; color:{PRIMARY_MID}; font-size:.74rem;
    font-weight:600; padding:5px 12px; border-radius:999px; box-shadow:0 1px 2px rgba(17,24,39,.04); }}

/* ── Cards ───────────────────────────────────────────────────────────────── */
.rf-card {{ background:#fff; border:1px solid {LINE}; border-radius:12px; padding:18px 20px;
    margin-bottom:14px; box-shadow:0 1px 3px rgba(17,24,39,.05); }}
.rf-card-title {{ color:{PRIMARY_DARK}; font-size:0.95rem; font-weight:700;
    border-bottom:1px solid #eef0f6; padding-bottom:7px; margin-bottom:10px; }}
.rf-metric {{ background:#fff; border:1px solid {LINE}; border-top:3px solid {PRIMARY};
    border-radius:12px; padding:14px 16px; text-align:center;
    box-shadow:0 1px 3px rgba(17,24,39,.05); transition:box-shadow .15s, transform .15s; }}
.rf-metric:hover {{ box-shadow:0 8px 20px rgba(49,46,129,.12); transform:translateY(-2px); }}
.rf-metric .val {{ color:{PRIMARY}; font-size:1.55rem; font-weight:800; line-height:1.1; }}
.rf-metric .lbl {{ color:{MUTED}; font-size:0.72rem; margin-top:3px; font-weight:500; }}
.rf-metric .sub {{ color:#9ca3af; font-size:0.68rem; margin-top:1px; }}
.rf-sp-name {{ font-size:.74rem; font-weight:700; color:{PRIMARY_MID}; line-height:1.25; min-height:2.3em;
    text-transform:none; overflow:hidden; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; }}
.rf-sp-row {{ display:flex; align-items:baseline; justify-content:space-between; gap:8px; margin-top:2px; }}
.rf-sp-val {{ font-size:1.55rem; font-weight:800; color:{PRIMARY_DARK}; letter-spacing:-.02em; }}
.rf-sp-delta {{ font-size:.72rem; font-weight:700; padding:2px 8px; border-radius:999px; white-space:nowrap; }}
.rf-sp-meta {{ font-size:.66rem; color:{MUTED}; margin-top:1px; }}

/* ── Tabs (pill style) ───────────────────────────────────────────────────── */
.stTabs [data-baseweb="tab-list"] {{ gap:4px; background:transparent; padding:0 0 8px 0;
    border-bottom:1px solid {LINE}; overflow-x:auto; flex-wrap:nowrap; scrollbar-width:thin; }}
.stTabs [data-baseweb="tab"] {{ background:transparent; border-radius:8px; padding:7px 13px;
    font-size:0.76rem; font-weight:600; color:{MUTED} !important; border:none;
    white-space:nowrap; transition:all .15s; }}
.stTabs [data-baseweb="tab"]:hover {{ background:{PRIMARY_SOFT} !important; color:{PRIMARY_MID} !important; }}
.stTabs [aria-selected="true"] {{ background:{PRIMARY} !important; color:#fff !important; font-weight:700 !important; }}
.stTabs [data-baseweb="tab-highlight"], .stTabs [data-baseweb="tab-border"] {{ display:none; }}
.stTabs [data-baseweb="tab-panel"] {{ background:#fff; border:1px solid {LINE}; border-radius:12px;
    padding:20px 22px; margin-top:12px; box-shadow:0 1px 3px rgba(17,24,39,.05); }}
.stTabs [data-baseweb="tab-panel"] .stTabs [data-baseweb="tab-panel"] {{ box-shadow:none; background:#fff; }}

/* ── Sidebar (light) ─────────────────────────────────────────────────────── */
[data-testid="stSidebar"] {{ background:#fff; border-right:1px solid {LINE}; }}
[data-testid="stSidebar"] .block-container {{ padding-top:1rem !important; }}
[data-testid="stSidebar"] h2 {{ color:{PRIMARY_DARK} !important; font-size:0.95rem !important; font-weight:700 !important;
    border-bottom:2px solid {PRIMARY}; padding-bottom:6px; }}
[data-testid="stSidebar"] h3, [data-testid="stSidebar"] .stMarkdown h3 {{
    color:{MUTED} !important; font-size:0.68rem !important; font-weight:700 !important;
    text-transform:uppercase; letter-spacing:0.08em; border-bottom:1px solid #eef0f6;
    padding-bottom:3px; margin-top:14px !important; }}
[data-testid="stSidebar"] span[data-baseweb="tag"] {{ background:{PRIMARY_SOFT} !important; color:{PRIMARY_MID} !important; }}
[data-testid="stSidebar"] .stCaption {{ color:{MUTED} !important; font-size:0.68rem !important; }}

/* ── Buttons ─────────────────────────────────────────────────────────────── */
.stButton > button, .stDownloadButton > button {{
    background:{PRIMARY} !important; color:#fff !important; border:none !important;
    border-radius:8px !important; font-weight:600 !important; font-size:0.82rem !important;
    padding:7px 16px !important; box-shadow:0 1px 3px rgba(49,46,129,.25) !important;
    transition:transform .12s, box-shadow .12s, background .12s !important; }}
.stButton > button:hover, .stDownloadButton > button:hover {{
    background:{PRIMARY_MID} !important; transform:translateY(-1px);
    box-shadow:0 5px 14px rgba(49,46,129,.28) !important; }}
[data-testid="stSidebar"] .stButton > button {{
    background:{PRIMARY_SOFT} !important; color:{PRIMARY_MID} !important;
    border:1px solid #c7d2fe !important; box-shadow:none !important;
    font-size:.72rem !important; padding:5px 6px !important; }}
[data-testid="stSidebar"] .stButton > button:hover {{ background:{PRIMARY} !important; color:#fff !important; }}

/* ── Headings ────────────────────────────────────────────────────────────── */
h2 {{ color:{PRIMARY_DARK} !important; font-size:1.05rem !important; font-weight:700 !important;
    border-bottom:2px solid #eef0f6 !important; padding-bottom:5px !important; margin-top:4px !important; }}
h3 {{ color:#374151 !important; font-size:0.88rem !important; font-weight:600 !important; }}

/* ── Inputs ──────────────────────────────────────────────────────────────── */
.stSelectbox > div > div, .stMultiSelect > div > div, .stTextInput input, .stNumberInput input {{
    border:1px solid #d7dbe6 !important; border-radius:8px !important;
    background:#fff !important; color:{INK} !important; }}

/* ── Native metrics / expander / status ──────────────────────────────────── */
[data-testid="stMetricValue"] {{ color:{PRIMARY} !important; font-weight:700 !important; }}
[data-testid="stMetricLabel"] {{ color:{MUTED} !important; font-size:0.75rem !important; }}
.streamlit-expanderHeader {{ background:#fff !important; color:{PRIMARY_DARK} !important;
    font-weight:600 !important; border-radius:8px !important; border:1px solid {LINE} !important; }}
.stSuccess {{ background:#ecfdf5 !important; border-left:4px solid #10b981 !important; }}
.stInfo    {{ background:{PRIMARY_SOFT} !important; border-left:4px solid {PRIMARY} !important; }}
.stWarning {{ background:#fffbeb !important; border-left:4px solid {AMBER} !important; }}
.stError   {{ background:#fef2f2 !important; border-left:4px solid #ef4444 !important; }}

/* ── Tables ──────────────────────────────────────────────────────────────── */
[data-testid="stDataFrame"] {{ border:1px solid {LINE}; border-radius:10px; overflow:hidden; }}
.stDataFrame thead tr th {{ background:#f9fafb !important; color:#4b5563 !important; font-size:0.72rem !important;
    font-weight:700 !important; text-transform:uppercase !important; letter-spacing:0.04em !important;
    border-bottom:2px solid {PRIMARY} !important; }}

/* ── Chat ────────────────────────────────────────────────────────────────── */
[data-testid="stChatMessageContent"] {{ background:#f9fafb !important; border-radius:10px !important; border:1px solid {LINE} !important; }}
</style>
""", unsafe_allow_html=True)


# ═════════════════════════════════════════════════════════════════════════════
# CONSTANTS
# ═════════════════════════════════════════════════════════════════════════════

# ID columns only exist in first N columns of vendor files
ID_ZONE_LIMIT = 8

GRANULARITY = {
    "Raw (no aggregation)": None,
    "Hourly":  "h",
    "Daily":   "D",
    "Weekly":  "W",
    "Monthly": "ME",
}

DEFAULT_KPI_THRESHOLDS = {
    "avail":      (99.0,  95.0),   # % — higher is better
    "rrc sr":     (97.0,  95.0),
    "cssr":       (97.0,  95.0),
    "rab sr":     (97.0,  95.0),
    "hosr":       (97.0,  94.0),
    "drop":       (2.0,   5.0),    # % — lower is better
    "fail":       (2.0,   5.0),
    "thp":        (5.0,   1.0),    # Mbps
    "throughput": (5.0,   1.0),
    "rsrp":       (-100,  -110),   # dBm
    "rsrq":       (-12,   -15),
    "sinr":       (5,     0),
    "cqi":        (8,     5),
    "bler":       (10,    20),     # % — lower is better
}

DATE_PATTERNS = [
    r'\d{4}-\d{2}-\d{2}',           # 2015-05-15
    r'\d{2}/\d{2}/\d{4}',           # 05/15/2015
    r'\d{2}-\d{2}-\d{4}',           # 15-05-2015
    r'\d{4}/\d{2}/\d{2}',           # 2015/05/15
    r'\d{2}\.\d{2}\.\d{4}',       # 15.05.2015
    r'\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}',  # 2015-05-15T08:00
    r'^\d{8}$',                        # 20150515 — MUST be entire string (anchored)
]


# ═════════════════════════════════════════════════════════════════════════════
# FILE LOADING
# ═════════════════════════════════════════════════════════════════════════════

def _read_csv_bytes(data: bytes) -> pd.DataFrame:
    """Try every common encoding × separator combination."""
    # Sniff the delimiter from the first line so ';' exports (e.g. Nokia planning
    # dumps) are parsed on the first attempt instead of after three failed ones.
    head = data[:4096].decode('utf-8', errors='ignore').splitlines()[:1]
    seps = [',', ';', '\t', '|']
    if head:
        seps.sort(key=lambda c: -head[0].count(c))
    for enc in ('utf-8', 'latin-1', 'cp1252', 'utf-16'):
        for sep in seps:
            try:
                df = pd.read_csv(
                    BytesIO(data), encoding=enc, sep=sep,
                    header=None, low_memory=False, on_bad_lines='skip',
                )
                if df.shape[1] >= 2:
                    return df
            except Exception:
                continue
    raise ValueError("Cannot parse as CSV with any common encoding/separator.")


def _smart_header(df: pd.DataFrame) -> pd.DataFrame:
    """
    Detect the true header row in vendor RAN exports.

    Nokia 3-row pattern:
      Row 0 — sparse group labels ("Cell Availability", "RRC Connections" …)
      Row 1 — column names        ("Period start time", "WCEL name", "Cell avail" …)
      Row 2 — Nokia counter IDs   ("RNC_183c", "RNC_727a" …)   ← skipped
      Row 3+ — data

    Scoring: coverage × name_ratio  →  row with 100% coverage wins over
    the sparse label row (28% coverage).
    """
    if df is None or df.empty:
        return df

    # Already has string column names — pandas parsed the header already
    if not all(isinstance(c, (int, np.integer)) for c in df.columns):
        return df

    n = len(df.columns)

    def row_score(row) -> float:
        vals = [v for v in row if pd.notna(v) and str(v).strip()]
        if not vals:
            return 0.0
        coverage = len(vals) / max(n, 1)
        name_like = sum(
            1 for v in vals
            if isinstance(v, str)
            and not re.match(r'^-?[\d\.]+$', v.strip())
            and not re.match(r'^\d{2}[/\-\.]\d{2}[/\-\.]\d{4}', v.strip())
            and not re.match(r'^\d{4}[/\-\.]\d{2}[/\-\.]\d{2}', v.strip())
            and len(v.strip()) > 1
        )
        return coverage * (name_like / len(vals))

    best_row, best_score = 0, -1.0
    for i in range(min(8, len(df))):
        s = row_score(df.iloc[i])
        if s > best_score:
            best_score, best_row = s, i

    # Detect Nokia counter-ID row immediately after header
    skip = 0
    next_i = best_row + 1
    if next_i < len(df):
        nr = df.iloc[next_i]
        counter_like = sum(
            1 for v in nr
            if pd.notna(v) and isinstance(v, str)
            and re.search(r'[A-Za-z]{1,12}[_\-][0-9A-Fa-f]{3,}', str(v).strip())
        )
        nn = sum(1 for v in nr if pd.notna(v) and str(v).strip())
        if counter_like >= 3 or (0 < nn < n * 0.15):
            skip = 1

    # Build deduplicated column list
    raw_header = df.iloc[best_row].tolist()
    seen: dict = {}
    columns: list = []
    for val in raw_header:
        name = str(val).strip() if (pd.notna(val) and str(val).strip()) else ''
        if not name:
            name = f'Col_{len(columns)}'
        if name in seen:
            seen[name] += 1
            name = f'{name}_{seen[name]}'
        else:
            seen[name] = 0
        columns.append(name)

    # Slice data
    data_start = best_row + 1 + skip
    result = df.iloc[data_start:].copy()
    result.columns = columns
    result = result.reset_index(drop=True)

    # Drop junk rows (all-NaN or very sparse group-label rows)
    def is_junk(row) -> bool:
        nn = row.notna().sum()
        return nn == 0 or (nn <= 3 and nn < n * 0.12)

    result = result[~result.apply(is_junk, axis=1)].reset_index(drop=True)

    # Coerce clearly numeric columns
    for col in result.columns:
        try:
            converted = pd.to_numeric(result[col], errors='coerce')
            non_null  = result[col].notna().sum()
            if non_null > 0 and converted.notna().sum() / non_null >= 0.85:
                result[col] = converted
        except Exception:
            pass

    return result


def load_file(uploaded_file) -> Tuple[Optional[pd.DataFrame], str]:
    """Load any uploaded file and return (DataFrame, format_description)."""
    name = uploaded_file.name.lower()
    data = uploaded_file.read()   # read once into memory

    try:
        # ── CSV / TSV / TXT ──────────────────────────────────────────────────
        if name.endswith(('.csv', '.txt', '.tsv')):
            df = _read_csv_bytes(data)
            return _smart_header(df), f"CSV ({uploaded_file.name})"

        # ── Modern XLSX ───────────────────────────────────────────────────────
        elif name.endswith('.xlsx'):
            if not HAS_OPENPYXL:
                st.error(f"❌ {uploaded_file.name} needs openpyxl → `pip install openpyxl`")
                return None, ''
            df = pd.read_excel(BytesIO(data), engine='openpyxl', header=None)
            return _smart_header(df), f"XLSX ({uploaded_file.name})"

        # ── Legacy XLS (BIFF8) ────────────────────────────────────────────────
        elif name.endswith('.xls'):
            if HAS_XLRD:
                wb = xlrd.open_workbook(file_contents=data)
                ws = wb.sheet_by_index(0)
                rows = [ws.row_values(i) for i in range(ws.nrows)]
                df = pd.DataFrame(rows)
                return _smart_header(df), f"XLS/xlrd ({uploaded_file.name})"

            # LibreOffice headless fallback
            lo = shutil.which('libreoffice') or shutil.which('soffice')
            if lo:
                with tempfile.NamedTemporaryFile(suffix='.xls', delete=False) as tmp:
                    tmp.write(data)
                    tmp_path = tmp.name
                out_dir = tempfile.mkdtemp()
                try:
                    subprocess.run(
                        [lo, '--headless', '--convert-to', 'csv',
                         '--outdir', out_dir, tmp_path],
                        capture_output=True, timeout=90,
                    )
                    csvs = [f for f in os.listdir(out_dir) if f.endswith('.csv')]
                    if csvs:
                        csv_data = open(os.path.join(out_dir, csvs[0]), 'rb').read()
                        df = _read_csv_bytes(csv_data)
                        return _smart_header(df), f"XLS→CSV ({uploaded_file.name})"
                finally:
                    os.unlink(tmp_path)
                    shutil.rmtree(out_dir, ignore_errors=True)

            st.error(
                f"❌ **{uploaded_file.name}** is a legacy `.xls` file.  \n"
                "Install xlrd: `pip install xlrd`  \n"
                "Or convert to `.xlsx` / `.csv` first."
            )
            return None, ''

        # ── ODS ───────────────────────────────────────────────────────────────
        elif name.endswith('.ods'):
            df = pd.read_excel(BytesIO(data), engine='odf', header=None)
            return _smart_header(df), f"ODS ({uploaded_file.name})"

        # ── Unknown — sniff ───────────────────────────────────────────────────
        else:
            magic = data[:4]
            if magic == b'PK\x03\x04':      # ZIP → likely XLSX
                df = pd.read_excel(BytesIO(data), engine='openpyxl', header=None)
                return _smart_header(df), f"Auto-XLSX ({uploaded_file.name})"
            df = _read_csv_bytes(data)
            return _smart_header(df), f"Auto-CSV ({uploaded_file.name})"

    except Exception as exc:
        st.error(f"❌ Could not load **{uploaded_file.name}**: {exc}")
        return None, ''


# ═════════════════════════════════════════════════════════════════════════════
# KML / KMZ + GEO HELPERS
# ═════════════════════════════════════════════════════════════════════════════

def load_kml(uploaded_file) -> Tuple[Optional[pd.DataFrame], str]:
    """Read a KML or KMZ file into a DataFrame: KML_Name, Latitude, Longitude.
    Polygons / lines are reduced to their centroid."""
    data = uploaded_file.read()
    try:
        blobs = []
        if data[:2] == b'PK':                               # KMZ = zipped KML
            with zipfile.ZipFile(BytesIO(data)) as z:
                blobs = [z.read(n) for n in z.namelist() if n.lower().endswith('.kml')]
            if not blobs:
                st.error(f"❌ {uploaded_file.name}: no .kml found inside the KMZ"); return None, ''
        else:
            blobs = [data]

        rows = []
        for blob in blobs:
            root = ET.fromstring(blob.lstrip(b'\xef\xbb\xbf'))
            for pm in root.iter():
                if pm.tag.split('}')[-1] != 'Placemark':
                    continue
                name, pts, extra = None, [], {}
                for el in pm.iter():
                    tag = el.tag.split('}')[-1]
                    if tag == 'Data' and el.get('name'):            # ExtendedData
                        v = next((c.text for c in el if c.tag.split('}')[-1] == 'value'), None)
                        if v and v.strip():
                            extra[el.get('name')] = v.strip()
                    elif tag == 'SimpleData' and el.get('name') and el.text:
                        extra[el.get('name')] = el.text.strip()
                    if tag == 'name' and name is None:
                        name = (el.text or '').strip()
                    elif tag == 'coordinates' and el.text:
                        for tok in re.sub(r'\s*,\s*', ',', el.text.strip()).split():
                            parts = tok.split(',')
                            try:
                                pts.append((float(parts[0]), float(parts[1])))
                            except (ValueError, IndexError):
                                pass
                    elif tag == 'coord' and el.text:        # gx:coord = "lon lat alt"
                        parts = el.text.split()
                        try:
                            pts.append((float(parts[0]), float(parts[1])))
                        except (ValueError, IndexError):
                            pass
                pts = [(x, y) for x, y in pts if -180 <= x <= 180 and -90 <= y <= 90]
                if name and pts:
                    row = {'KML_Name': name,
                           'Latitude':  float(np.mean([y for _, y in pts])),
                           'Longitude': float(np.mean([x for x, _ in pts]))}
                    for k, v in extra.items():
                        row.setdefault(k, v)
                    rows.append(row)
        if not rows:
            st.error(f"❌ {uploaded_file.name}: no placemarks with coordinates found")
            return None, ''
        return pd.DataFrame(rows), f"KML ({uploaded_file.name}, {len(rows)} placemarks)"
    except Exception as e:
        st.error(f"❌ Could not read {uploaded_file.name}: {e}")
        return None, ''


def to_coord(series: pd.Series) -> pd.Series:
    """Coerce a lat/lon column to float. Handles text, decimal commas and
    N/S/E/W suffixes; anything unparseable becomes NaN."""
    if pd.api.types.is_numeric_dtype(series):
        return series.astype(float)
    t = series.astype(str).str.strip()
    sign = np.where(t.str.contains(r'[SsWw]\s*$'), -1.0, 1.0)
    t = (t.str.replace(r'\s*[NnSsEeWw]\s*$', '', regex=True)   # drop hemisphere letter
          .str.replace(',', '.', regex=False)
          .str.replace(r'[^0-9eE+\-.]', '', regex=True))
    return pd.to_numeric(t, errors='coerce') * sign


# ═════════════════════════════════════════════════════════════════════════════
# PLANNING-FILE HELPERS  (Nokia 5G cell parameters, sector azimuths, geo)
# ═════════════════════════════════════════════════════════════════════════════

# planning column -> sidebar label (shown top-down as a Region → … hierarchy)
PLAN_FILTER_LABELS = {
    'REGIONCODE':       '📍 Region',
    'CITY':             '📍 City',
    'DISTRICT':         '📍 District',
    'TEKNOLOJI':        '📶 Technology',
    'FREQUENCY_BAND':   '📶 Band',
    'CELL_ATLAS_STATE': '🏷️ Cell state',
}
# planning columns that may be used to join onto the KPI file
PLAN_JOIN_HINTS = {'CELL_NAME': 'cell_id', 'DISTNAME': 'cell_id',
                   'GNODEB_NAME': 'site_id', 'EMISYON_NAME': 'site_id'}
# wedge length multiplier per technology so co-sited layers stay distinguishable
TECH_RADIUS = {'FDD_NR': 1.0, 'DSS': 0.8, 'C_BAND': 0.6}
WEDGE_MAX_CELLS = 4000


class _MemFile:
    """Duck-types a Streamlit UploadedFile so load_file() can read raw bytes."""
    def __init__(self, name: str, data: bytes):
        self.name, self._d, self.size = name, data, len(data)
    def read(self): return self._d
    def getvalue(self): return self._d


class _LocalFile(_MemFile):
    def __init__(self, path: str):
        with open(path, 'rb') as fh:
            super().__init__(os.path.basename(path), fh.read())


def find_bundled_planning() -> Optional[str]:
    """Newest planning CSV sitting next to the script (or in ./planning)."""
    try:
        base = os.path.dirname(os.path.abspath(__file__))
    except NameError:
        base = os.getcwd()
    found = []
    for d in {base, os.path.join(base, 'planning'), os.getcwd(),
              os.path.join(os.getcwd(), 'planning')}:
        for pat in ('NOKIA_5G_Cell_Parameters*.csv', '*Cell_Parameters*.csv'):
            found += glob.glob(os.path.join(d, pat))
    return max(found, key=lambda x: os.path.basename(x)) if found else None


def rescale_degrees(s: pd.Series, limit: float) -> Tuple[pd.Series, int]:
    """Divide by 10**k (smallest k) so that 99 % of |values| <= limit.
    Handles coordinates stored as degrees x 10^5 (3988638 -> 39.88638)."""
    v = s.dropna().abs()
    if v.empty:
        return s, 0
    q = float(np.nanpercentile(v, 99))
    for k in range(0, 10):
        if q / (10 ** k) <= limit:
            return s / (10 ** k), k
    return s, 0


def _as_label(s: pd.Series) -> pd.Series:
    """Text labels for filter columns (78.0 -> '78', blanks stay missing)."""
    if pd.api.types.is_numeric_dtype(s):
        return s.map(lambda v: None if pd.isna(v)
                     else (str(int(v)) if float(v).is_integer() else str(v)))
    return s.map(lambda v: None if pd.isna(v) else str(v).strip())


def prepare_planning(df: pd.DataFrame) -> Tuple[pd.DataFrame, dict]:
    """Normalise a planning table.
    * scaled coordinates -> decimal degrees in Plan_Latitude / Plan_Longitude
    * Nokia 5G cell-parameter schema (CELL_NAME + AZIMUTH): every attribute
      column is pinned to type 'dimension' so it is never mistaken for a KPI or
      a time axis, join keys are hinted, filter columns become text labels.
    Returns (df, info)."""
    info = {'forced': {}, 'hints': {}, 'filters': [], 'notes': [], 'date': None}
    df = df.copy()
    up = {str(c).upper(): c for c in df.columns}
    lat_c = next((up[n] for n in ('LATITUDE_DEC', 'LATITUDE', 'LAT') if n in up), None)
    lon_c = next((up[n] for n in ('LONGITUDE_DEC', 'LONGITUDE', 'LON', 'LONG') if n in up), None)
    nokia = 'CELL_NAME' in up and 'AZIMUTH' in up

    if lat_c and lon_c:
        la, ka = rescale_degrees(to_coord(df[lat_c]), 90)
        lo, ko = rescale_degrees(to_coord(df[lon_c]), 180)
        if ka or ko or nokia:
            ok = la.between(-90, 90) & lo.between(-180, 180) & ~((la == 0) & (lo == 0))
            df = df.drop(columns=[lat_c, lon_c])
            df['Plan_Latitude'] = la.where(ok)
            df['Plan_Longitude'] = lo.where(ok)
            info['forced']['Plan_Latitude'] = 'latitude'
            info['forced']['Plan_Longitude'] = 'longitude'
            if ka or ko:
                info['notes'].append(
                    f"`{lat_c}` / `{lon_c}` were stored as scaled integers "
                    f"(÷10^{ka} / ÷10^{ko}) → converted to decimal degrees.")
            nbad = int((~ok).sum())
            if nbad:
                info['notes'].append(f"{nbad:,} planning rows have no usable coordinates.")

    if nokia:
        az = up['AZIMUTH']
        df[az] = pd.to_numeric(df[az], errors='coerce')
        df.loc[~df[az].between(0, 360), az] = np.nan
        df.loc[df[az] == 360, az] = 0
        for c in df.columns:
            info['forced'].setdefault(c, 'dimension')
        info['hints'] = {c: t for c, t in PLAN_JOIN_HINTS.items() if c in df.columns}
        for c in PLAN_FILTER_LABELS:
            if c in df.columns:
                df[c] = _as_label(df[c])
                info['filters'].append(c)
        if 'WDATE' in df.columns:
            info['date'] = str(df['WDATE'].dropna().astype(str).max())[:10]
    return df, info


@st.cache_data(show_spinner=False, max_entries=4)
def load_planning_cached(name: str, data: bytes):
    """Parse + prepare a planning file once per unique upload (not every rerun)."""
    df, fmt = load_file(_MemFile(name, data))
    if df is None or df.empty:
        return None, fmt, {}
    df, info = prepare_planning(df)
    return df, fmt, info


def sector_polygons(lat, lon, az, radius_m, beam_deg: float, n_arc: int = 9):
    """Closed wedge outlines (site → arc → site) for many sectors, joined with
    None separators so one Scattermap trace can draw them all."""
    out_lat, out_lon = [], []
    m_per_deg = 111_320.0
    for la, lo, a, r in zip(lat, lon, az, radius_m):
        th = np.radians(np.linspace(a - beam_deg / 2, a + beam_deg / 2, n_arc))
        d_lat = r * np.cos(th) / m_per_deg
        d_lon = r * np.sin(th) / (m_per_deg * max(np.cos(np.radians(la)), 0.01))
        out_lat += [la, *(la + d_lat), la, None]
        out_lon += [lo, *(lo + d_lon), lo, None]
    return out_lat, out_lon


def sector_tip(lat, lon, az, radius_m):
    """Point 70 % along the sector boresight (used as the hover/colour marker)."""
    th = np.radians(np.asarray(az, dtype=float))
    r = np.asarray(radius_m, dtype=float) * 0.7
    la = np.asarray(lat, dtype=float)
    return (la + r * np.cos(th) / 111_320.0,
            np.asarray(lon, dtype=float) + r * np.sin(th) / (111_320.0 * np.maximum(np.cos(np.radians(la)), 0.01)))


def _rgba(c: str, a: float) -> str:
    c = c.strip()
    if c.startswith('#') and len(c) == 7:
        return f"rgba({int(c[1:3],16)},{int(c[3:5],16)},{int(c[5:7],16)},{a})"
    m = re.findall(r'[\d.]+', c)
    return f"rgba({m[0]},{m[1]},{m[2]},{a})" if len(m) >= 3 else c



def _norm_key(values) -> pd.Series:
    """Upper-case, trimmed text; '1234.0' -> '1234'."""
    return (pd.Series(values).astype(str).str.strip().str.upper()
              .str.replace(r'\.0+$', '', regex=True))


def _alnum(x: str) -> str:
    return re.sub(r'[^A-Z0-9]', '', x)


def smart_map(kpi_vals, plan_vals, fuzzy: bool = True) -> Tuple[dict, str]:
    """Match KPI ids to planning names. Returns ({kpi_norm: plan_norm}, method).
    Tries: exact -> ignoring punctuation -> plan name inside KPI id
    -> KPI id inside plan name."""
    kn = list(dict.fromkeys(_norm_key(pd.Series(kpi_vals).dropna().unique())))
    pn = list(dict.fromkeys(_norm_key(pd.Series(plan_vals).dropna().unique())))
    pset = set(pn)
    m = {k: k for k in kn if k in pset}
    if m:
        return m, "exact"
    pal = {}
    for n in pn:
        pal.setdefault(_alnum(n), n)
    m = {k: pal[_alnum(k)] for k in kn if _alnum(k) and _alnum(k) in pal}
    if m:
        return m, "ignoring punctuation"
    if not fuzzy:          # contains-matching is O(n·m): only when nothing else matched
        return {}, "none"
    m = {}
    names = sorted((n for n in pn if len(n) >= 4), key=len, reverse=True)
    for k in kn:
        for n in names:
            if n in k:
                m[k] = n
                break
    if m:
        return m, "planning name contained in KPI id"
    m = {}
    ks = sorted((k for k in kn if len(k) >= 4), key=len, reverse=True)
    for n in pn:
        for k in ks:
            if k in n:
                m.setdefault(k, n)
                break
    if m:
        return m, "KPI id contained in planning name"
    return {}, "none"


def fuzzy_name_map(kpi_vals, plan_vals, min_len: int = 4) -> dict:
    """Map each KPI id to the longest plan name contained in it (e.g. site
    'LHR001' placemark matched to cell 'LHR001_A'). Keys/values upper-cased."""
    names = sorted({n for n in _norm_key(plan_vals) if len(n) >= min_len},
                   key=len, reverse=True)
    out = {}
    for v in _norm_key(kpi_vals).unique():
        for n in names:
            if n in v:
                out[v] = n
                break
    return out


# ═════════════════════════════════════════════════════════════════════════════
# COLUMN TYPE DETECTION
# ═════════════════════════════════════════════════════════════════════════════

# ── Strict ID keywords (only valid in first ID_ZONE_LIMIT columns) ────────────
_ID_EXACT: dict[str, str] = {}   # normalised_name → type
_ID_CONTAINS: list[tuple[str, str]] = []   # (substring, type)

_RAW_IDS = {
    'cell_id': [
        'wcel name','wcel id','wcelname','wcelid',
        'lncel name','lncel id','lncelname','lncelid',
        'eutrancell','eutrancellfdd','eutrancelltdd',
        'nrcell','nr cell','nrcelldu','nrcellcu',
        'nrcel name','nrcel id','nrcelname','nrcelid','nrcel',      # Nokia 5G NR cell
        'cell name','cellname','cell id','cellid','cell_id','cell_name',
        'localcellid','local cell id',
        'sector name','sector_id','sectorid',
    ],
    'site_id': [
        'wbts name','wbts id','wbtsname','wbtsid',
        'lnbts name','lnbts id','lnbtsname','lnbtsid',
        'enodeb name','enodeb id','enodebid',
        'nodeb name','nodeb id','nodebid',
        'node name','nodeid',
        'ne name','ne id','nename','neid',
        'enb name','enb id','enbname','enbid',
        'sitename','site_name','site_id','siteid',
        'bts name','bts_id','btsid',
        'gnb name','gnb_id','gnbid',
        'mrbts name','mrbts id','mrbtsname','mrbtsid','mrbts',      # Nokia managed radio BTS
        'nrbts name','nrbts id','nrbtsname','nrbtsid','nrbts',      # Nokia 5G gNB
    ],
    'rnc_id': ['rnc name','rnc id','rncname','rncid'],
    'plmn':   ['plmn name','plmnname','plmn id','plmn','operator'],
}

for _type, _kws in _RAW_IDS.items():
    for _kw in _kws:
        _norm = _kw.replace(' ', '_').replace('-', '_')
        _ID_EXACT[_norm] = _type
        _ID_CONTAINS.append((_norm, _type))

# ── Non-ID structural keywords (no position restriction) ──────────────────────
_STRUCTURAL_EXACT: dict[str, str] = {}
_STRUCTURAL_CONTAINS: list[tuple[str, str]] = []

_RAW_STRUCTURAL = {
    'latitude':  ['latitude','lat','lat_deg','y_coord'],
    'longitude': ['longitude','lon','lng','long','lon_deg','x_coord'],
    'date': [
        'period start time','period end time','period start','period end',
        'start time','end time','start_time','end_time',
        'report_time','report time','collection time',
        'measurement time','measurement date',
        'datetime','timestamp','date','period',
    ],
    'date_exact': ['time','hour','day','week','month','year'],  # only exact match
    'duration': [
        'duration','serv dur','cs serv dur','ps nrt serv dur',
        'nrt dur','min per drop','data per drop',
    ],
}

for _type, _kws in _RAW_STRUCTURAL.items():
    if _type == 'date_exact':
        for _kw in _kws:
            _STRUCTURAL_EXACT[_kw.replace(' ', '_')] = 'date'
    else:
        for _kw in _kws:
            _norm = _kw.replace(' ', '_').replace('-', '_')
            _STRUCTURAL_EXACT[_norm] = _type
            _STRUCTURAL_CONTAINS.append((_norm, _type))

# ── Plural network element names → count KPIs ─────────────────────────────────
_PLURAL_PATTERNS = [
    re.compile(p) for p in [
        r'^rncs$', r'^wbtss$', r'^wcels$', r'^cells$', r'^sites$',
        r'^enodebs$', r'^gnbs$', r'^sectors$',
        r'number[_\s]of[_\s]rnc', r'number[_\s]of[_\s]wbts',
        r'number[_\s]of[_\s]wcel', r'number[_\s]of[_\s]cell',
        r'nbr[_\s]rnc', r'nbr[_\s]wbts', r'nbr[_\s]wcel', r'nbr[_\s]cell',
        r'#[_\s]rnc', r'#[_\s]wbts', r'#[_\s]wcel',
    ]
]


def _col_norm(col: str) -> str:
    return col.lower().strip().replace(' ', '_').replace('-', '_')


def _looks_like_date_value(v: str) -> bool:
    return any(re.search(p, v) for p in DATE_PATTERNS)


def detect_column_type(col: str, series: pd.Series, position: int) -> Tuple[str, float]:
    """
    Return (type, confidence) for one column.

    Types: cell_id | site_id | rnc_id | plmn | latitude | longitude |
           date | duration | numeric_kpi | identifier | dimension | unknown
    """
    norm     = _col_norm(col)
    numeric  = pd.api.types.is_numeric_dtype(series)
    beyond   = position >= ID_ZONE_LIMIT
    non_null = series.notna().sum()

    # ── 0. Plural network-element count columns → always KPI ─────────────────
    for pat in _PLURAL_PATTERNS:
        if pat.search(norm):
            return 'numeric_kpi', 0.92

    # ── 1. Structural keywords (date, lat, lon, duration) ────────────────────
    if norm in _STRUCTURAL_EXACT:
        return _STRUCTURAL_EXACT[norm], 0.97

    for substr, stype in _STRUCTURAL_CONTAINS:
        if re.search(r'(?<![a-z])' + re.escape(substr) + r'(?![a-z])', norm):
            return stype, 0.90

    # ── 2. ID keywords — ONLY inside first ID_ZONE_LIMIT columns ─────────────
    if not beyond:
        if norm in _ID_EXACT:
            return _ID_EXACT[norm], 0.97
        for substr, itype in _ID_CONTAINS:
            if re.search(r'(?<![a-z])' + re.escape(substr) + r'(?![a-z])', norm):
                return itype, 0.90

    # ── 3. All numeric columns → KPI (the most common case for RAN files) ─────
    if numeric:
        vals = series.dropna()
        if len(vals) > 0 and not beyond:
            mn, mx = float(vals.min()), float(vals.max())
            # Lat/lon guard — only check inside ID zone for small ranges
            if -90 <= mn <= mx <= 90 and (mx - mn) < 5:
                return 'latitude', 0.62
            if -180 <= mn <= mx <= 180 and 5 <= (mx - mn) < 15:
                return 'longitude', 0.60
        return 'numeric_kpi', 0.92

    # ── 4. All-null column beyond ID zone → KPI placeholder ──────────────────
    if non_null == 0 and beyond:
        return 'numeric_kpi', 0.80

    # ── 5. Date heuristic on string values ───────────────────────────────────
    sample = series.dropna().astype(str).head(20)
    if len(sample) > 0:
        hits = sum(1 for v in sample if _looks_like_date_value(v))
        if hits / len(sample) >= 0.6:
            return 'date', round(hits / len(sample), 2)

    # ── 6. Sparse numeric strings (all-null or mixed) outside ID zone → KPI ──
    if beyond and non_null > 0:
        num_try = pd.to_numeric(series.dropna(), errors='coerce').notna().sum()
        if num_try / non_null >= 0.7:
            return 'numeric_kpi', 0.80

    # ── 7. String heuristics ─────────────────────────────────────────────────
    if not beyond and non_null > 0:
        uniqueness = series.nunique() / max(non_null, 1)
        if uniqueness > 0.5:
            return 'identifier', 0.55

    # Beyond ID zone with no other classification → treat as KPI placeholder
    if beyond:
        return 'numeric_kpi', 0.72
    return 'dimension', 0.50


def detect_all_columns(df: pd.DataFrame) -> Tuple[dict, dict]:
    """Returns (types dict, confidence dict) for all columns."""
    types, conf = {}, {}
    for i, col in enumerate(df.columns):
        types[col], conf[col] = detect_column_type(col, df[col], i)
    return types, conf


# ═════════════════════════════════════════════════════════════════════════════
# DATE PARSING
# ═════════════════════════════════════════════════════════════════════════════

def _excel_serial_to_date(series: pd.Series) -> pd.Series:
    """Convert Excel serial date numbers (e.g. 42139 = 2015-05-15)."""
    numeric = pd.to_numeric(series, errors='coerce')
    if numeric.notna().sum() > 0:
        mn = numeric.dropna().min()
        mx = numeric.dropna().max()
        if 25000 <= mn and mx <= 60000:
            return pd.to_datetime(numeric, unit='D', origin='1899-12-30', errors='coerce')
    return pd.to_datetime(series, errors='coerce')


def parse_date_column(series: pd.Series) -> pd.Series:
    """Parse a date column handling both string dates and Excel serial numbers."""
    if pd.api.types.is_numeric_dtype(series):
        return _excel_serial_to_date(series)

    sample = series.dropna().astype(str).head(10).tolist()
    fmt = None
    for sv in sample:
        if re.match(r'\d{2}/\d{2}/\d{4}', sv):   fmt = '%m/%d/%Y'; break
        if re.match(r'\d{4}-\d{2}-\d{2}', sv):   fmt = '%Y-%m-%d'; break
        if re.match(r'\d{2}-\d{2}-\d{4}', sv):   fmt = '%d-%m-%Y'; break
        if re.match(r'\d{2}\.\d{2}\.\d{4}', sv): fmt = '%d.%m.%Y'; break
        if re.match(r'\d{4}/\d{2}/\d{2}', sv):   fmt = '%Y/%m/%d'; break
        if re.match(r'\d{4}\d{2}\d{2}$', sv):    fmt = '%Y%m%d';   break

    try:
        if fmt:
            return pd.to_datetime(series, format=fmt, errors='coerce')
        else:
            return pd.to_datetime(series, errors='coerce')
    except Exception:
        return pd.to_datetime(series, errors='coerce')


def find_time_col(types: dict, df: pd.DataFrame) -> Optional[str]:
    """Return the best date column — skip numeric-only ones unless serial dates."""
    for col, t in types.items():
        if t != 'date' or col not in df.columns:
            continue
        parsed = parse_date_column(df[col])
        valid  = parsed.notna().sum() / max(df[col].notna().sum(), 1)
        if valid >= 0.5:
            return col
    return None


# ═════════════════════════════════════════════════════════════════════════════
# ANALYTICS FUNCTIONS
# ═════════════════════════════════════════════════════════════════════════════

def resample(df: pd.DataFrame, time_col: str,
             group_cols: list, kpi_cols: list, freq: str) -> pd.DataFrame:
    if not freq or time_col not in df.columns:
        return df
    try:
        grp = [pd.Grouper(key=time_col, freq=freq)] + [c for c in group_cols if c in df.columns]
        agg = {k: 'mean' for k in kpi_cols if k in df.columns}
        return df.groupby(grp, observed=True).agg(agg).reset_index() if agg else df
    except Exception:
        return df


def health_score(df: pd.DataFrame, kpi_cols: list,
                 cell_col: str, thresholds: dict) -> pd.DataFrame:
    if not cell_col or cell_col not in df.columns or not kpi_cols:
        return pd.DataFrame()
    avgs = df.groupby(cell_col, observed=True)[kpi_cols].mean()
    rows = []
    for cell, row in avgs.iterrows():
        crit = viol = 0
        for kpi in kpi_cols:
            kl  = kpi.lower()
            thr = next(((d, c) for pat, (d, c) in thresholds.items() if pat in kl), None)
            if thr is None:
                continue
            val = row.get(kpi, np.nan)
            if pd.isna(val):
                continue
            worse_high = any(x in kl for x in ['drop','fail','bler','util','error'])
            deg, crt = thr
            if worse_high:
                if crt is not None and val > abs(crt): crit += 1
                elif deg is not None and val > abs(deg): viol += 1
            else:
                if crt is not None and val < crt: crit += 1
                elif deg is not None and val < deg: viol += 1
        if crit:
            label, score = "🔴 Critical", max(0, 40 - crit * 10)
        elif viol:
            label, score = "🟡 Degraded", max(41, 75 - viol * 5)
        else:
            label, score = "🟢 Healthy", 100
        rows.append({'Cell': cell, 'Score': score, 'Status': label,
                     'Critical': crit, 'Degraded': viol})
    return pd.DataFrame(rows).set_index('Cell') if rows else pd.DataFrame()


ROOT_CAUSE_RULES = [
    ('drop',  'rsrp',  '📡 High drop rate + low RSRP → **Coverage issue** (weak signal)'),
    ('drop',  'sinr',  '📶 High drop rate + low SINR → **Interference issue**'),
    ('drop',  'hosr',  '🔄 High drop rate + low HO SR → **Handover failure**'),
    ('thp',   'prb',   '📊 Low throughput + high PRB → **Cell congestion**'),
    ('thp',   'rsrp',  '📡 Low throughput + low RSRP → **Coverage/path-loss**'),
    ('thp',   'cqi',   '📶 Low throughput + low CQI → **Poor radio conditions**'),
    ('rrc',   'sinr',  '📶 Low RRC SR + low SINR → **Uplink interference**'),
    ('erab',  'rsrp',  '📡 ERAB drop + low RSRP → **Cell edge/coverage hole**'),
    ('volte', 'sinr',  '🎙️ VoLTE drop + low SINR → **Interference on voice bearer**'),
    ('cssr',  'rsrp',  '📡 Low CSSR + low RSRP → **Coverage issue affecting setup**'),
]


def root_cause_hints(corr: pd.DataFrame, kpis: list, threshold: float = 0.5) -> list[str]:
    hints: set = set()
    for kpi in kpis:
        kl = kpi.lower()
        for rk, ck, hint in ROOT_CAUSE_RULES:
            if rk in kl and kpi in corr.columns:
                for other in corr.columns:
                    if ck in other.lower():
                        val = corr.loc[kpi, other] if kpi in corr.index else 0
                        if abs(val) >= threshold:
                            hints.add(hint)
    return sorted(hints)


def linear_forecast(df: pd.DataFrame, time_col: str,
                    kpi: str, horizon: int = 14) -> Tuple[pd.DataFrame, pd.DataFrame]:
    tmp = (df[[time_col, kpi]].dropna()
           .groupby(pd.Grouper(key=time_col, freq='D'))[kpi]
           .mean().reset_index().dropna())
    if len(tmp) < 5:
        return tmp, pd.DataFrame()
    tmp['_x'] = (tmp[time_col] - tmp[time_col].min()).dt.days.values
    model = LinearRegression().fit(tmp[['_x']], tmp[kpi])
    fitted = model.predict(tmp[['_x']])
    std = float(np.std(tmp[kpi].values - fitted))
    last_x = int(tmp['_x'].max())
    fx = np.arange(last_x + 1, last_x + 1 + horizon).reshape(-1, 1)
    fp = model.predict(fx)
    future_dates = [tmp[time_col].max() + timedelta(days=i) for i in range(1, horizon + 1)]
    fcast = pd.DataFrame({
        time_col: future_dates, kpi: fp,
        'lower': fp - 1.96 * std, 'upper': fp + 1.96 * std,
    })
    tmp = tmp.drop(columns=['_x'])
    tmp['lower'] = np.nan
    tmp['upper'] = np.nan
    return tmp, fcast


def change_points(series: pd.Series, n_max: int = 5) -> list[int]:
    arr = series.dropna().values
    if len(arr) < 10:
        return []
    if HAS_RUPTURES:
        try:
            bkps = rpt.Pelt(model='rbf').fit(arr).predict(pen=3)
            return [b for b in bkps if b < len(arr)][:n_max]
        except Exception:
            pass
    # Sliding-window fallback
    w = max(3, len(arr) // 20)
    diffs = [abs(arr[i:i+w].mean() - arr[i-w:i].mean())
             for i in range(w, len(arr) - w)]
    if not diffs:
        return []
    thr = np.mean(diffs) + 2 * np.std(diffs)
    result, last = [], -999
    for i, d in enumerate(diffs):
        if d > thr and (i + w) - last > w:
            result.append(i + w)
            last = i + w
    return result[:n_max]


# ═════════════════════════════════════════════════════════════════════════════
# AI QUERY
# ═════════════════════════════════════════════════════════════════════════════

GROQ_URL = "https://api.groq.com/openai/v1"
# Preferred models, best first. The list actually available to the key is
# fetched live, so retired model names no longer break the AI tab.
GROQ_PREFERRED = [
    "llama-3.3-70b-versatile",
    "openai/gpt-oss-120b",
    "openai/gpt-oss-20b",
    "meta-llama/llama-4-scout-17b-16e-instruct",
    "llama-3.1-8b-instant",
]


@st.cache_data(ttl=600, show_spinner=False)
def groq_available_models(api_key: str) -> list:
    """Chat-capable model ids this API key can use, preferred ones first."""
    try:
        r = requests.get(f"{GROQ_URL}/models",
                         headers={"Authorization": f"Bearer {api_key}"}, timeout=20)
        ids = [m["id"] for m in r.json().get("data", [])
               if m.get("active", True)
               and not any(x in m["id"].lower()
                           for x in ("whisper", "guard", "tts", "playai", "distil-whisper"))]
    except Exception:
        return list(GROQ_PREFERRED)
    first = [m for m in GROQ_PREFERRED if m in ids]
    return first + [m for m in ids if m not in first] or list(GROQ_PREFERRED)


def ai_query(question: str, df: pd.DataFrame,
             kpi_cols: list, cell_col: str, time_col: str,
             api_key: str) -> str:
    if not api_key:
        return (
            "⚠️ **API key not set.**\n\n"
            "Enter your **Groq API key** (free at groq.com) in the input field above, "
            "or add it to `.streamlit/secrets.toml`:\n```\n"
            "GROQ_API_KEY = \"gsk_...\"\n```"
        )

    # Build context with actual data
    ctx_parts = [
        f"Total rows: {len(df)}",
        f"KPI columns: {kpi_cols}",
        f"Cell column: {cell_col or 'not detected'}",
        f"Time column: {time_col or 'not detected'}",
    ]

    if kpi_cols:
        stats = df[kpi_cols].agg(['mean', 'min', 'max', 'std']).round(3).to_dict()
        ctx_parts.append(f"KPI statistics:\n{json.dumps(stats, indent=2)}")

    if cell_col and cell_col in df.columns and kpi_cols:
        avgs = (df.groupby(cell_col, observed=True)[kpi_cols]
                  .mean().round(3).reset_index())
        ctx_parts.append(
            f"Per-cell averages (up to 60 cells):\n"
            f"{avgs.head(60).to_string(index=False)}"
        )

    if time_col and time_col in df.columns:
        ctx_parts.append(
            f"Date range: {df[time_col].min().date()} to {df[time_col].max().date()}"
        )

    system = (
        "You are a senior RAN (Radio Access Network) performance analytics assistant.\n"
        "Answer questions using ONLY the data provided below.\n"
        "For ranking/comparison questions, produce a clean markdown table.\n"
        "Be concise and precise. If data is insufficient, say so.\n\n"
        "DATA CONTEXT:\n" + "\n".join(ctx_parts)
    )

    last_err = ""
    for model_id in groq_available_models(api_key)[:4]:     # try up to 4 models
        try:
            resp = requests.post(
                f"{GROQ_URL}/chat/completions",
                headers={"Content-Type": "application/json",
                         "Authorization": f"Bearer {api_key}"},
                json={"model": model_id, "max_tokens": 1024,
                      "messages": [{"role": "system", "content": system},
                                   {"role": "user",   "content": question}]},
                timeout=45,
            )
            data = resp.json()
            if "error" in data:
                last_err = data["error"].get("message", str(data["error"]))
                low = last_err.lower()
                if ("model" in low and ("not exist" in low or "decommission" in low
                                        or "no access" in low or "not found" in low)):
                    continue                                  # try next model
                return f"❌ API error: {last_err}"
            try:
                return data["choices"][0]["message"]["content"]
            except (KeyError, IndexError):
                return f"⚠️ Unexpected response: {str(data)[:400]}"
        except requests.exceptions.Timeout:
            return "❌ Request timed out. Try a simpler question."
        except Exception as exc:
            return f"❌ Error: {exc}"
    return (f"❌ None of the available Groq models worked. Last error: {last_err}\n\n"
            "Check your key at console.groq.com and that your account can use chat models.")


# ═════════════════════════════════════════════════════════════════════════════
# PDF REPORT
# ═════════════════════════════════════════════════════════════════════════════

def build_pdf(summary: pd.DataFrame, health: pd.DataFrame,
              worst: pd.DataFrame, anom_count: int,
              date_str: str, filter_str: str) -> bytes:
    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            leftMargin=2*cm, rightMargin=2*cm,
                            topMargin=2*cm, bottomMargin=2*cm)
    styles = getSampleStyleSheet()
    T  = ParagraphStyle('T',  parent=styles['Title'],   fontSize=18, spaceAfter=12)
    H2 = ParagraphStyle('H2', parent=styles['Heading2'], fontSize=13, spaceAfter=6,
                         textColor=RC.HexColor('#4338ca'))
    B  = styles['Normal']
    story = []
    story.append(Paragraph("📡 RAN KPI Analytics Report", T))
    story.append(Paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}", B))
    story.append(Paragraph(f"Period: {date_str}", B))
    story.append(Paragraph(f"Filters: {filter_str}", B))
    story.append(Spacer(1, 0.4*cm))
    story.append(Paragraph("KPI Summary Statistics", H2))
    if not summary.empty:
        td = [['KPI'] + list(summary.columns)]
        for idx, row in summary.iterrows():
            td.append([str(idx)] + [
                f"{v:.3f}" if isinstance(v, (float, np.floating)) else str(v)
                for v in row])
        t = Table(td, repeatRows=1)
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), RC.HexColor('#4338ca')),
            ('TEXTCOLOR',  (0,0), (-1,0), RC.white),
            ('FONTSIZE',   (0,0), (-1,-1), 8),
            ('ROWBACKGROUNDS', (0,1), (-1,-1),
             [RC.white, RC.HexColor('#f5f6fb')]),
            ('GRID', (0,0), (-1,-1), 0.25, RC.grey),
            ('ALIGN', (1,0), (-1,-1), 'RIGHT'),
        ]))
        story.append(t)
    story.append(Spacer(1, 0.3*cm))
    if not health.empty:
        story.append(Paragraph("Cell Health Summary", H2))
        for lbl, cnt in health['Status'].value_counts().items():
            story.append(Paragraph(f"  {lbl}: {cnt} cells", B))
    story.append(Spacer(1, 0.2*cm))
    story.append(Paragraph("Anomaly Detection", H2))
    story.append(Paragraph(f"Total anomalies detected: {anom_count}", B))
    if not worst.empty:
        story.append(PageBreak())
        story.append(Paragraph("Worst Performing Cells", H2))
        td = [list(worst.reset_index().columns)]
        for _, row in worst.reset_index().iterrows():
            td.append([
                f"{v:.3f}" if isinstance(v, (float, np.floating)) else str(v)
                for v in row])
        t = Table(td, repeatRows=1)
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), RC.HexColor('#c0392b')),
            ('TEXTCOLOR',  (0,0), (-1,0), RC.white),
            ('FONTSIZE',   (0,0), (-1,-1), 7),
            ('ROWBACKGROUNDS', (0,1), (-1,-1),
             [RC.white, RC.HexColor('#fde8e8')]),
            ('GRID', (0,0), (-1,-1), 0.25, RC.grey),
        ]))
        story.append(t)
    doc.build(story)
    return buf.getvalue()


# ═════════════════════════════════════════════════════════════════════════════
# HELPERS
# ═════════════════════════════════════════════════════════════════════════════

def to_excel(df: pd.DataFrame, sheet: str = 'Data') -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as w:
        df.to_excel(w, index=False, sheet_name=sheet)
    return out.getvalue()


def file_sig(files) -> str:
    return hashlib.md5(
        ''.join(f"{f.name}{f.size}" for f in files).encode()
    ).hexdigest()


def metric_card(label: str, value: str, sub: str = '') -> str:
    sub_html = f'<div class="sub">{sub}</div>' if sub else ''
    return (f'<div class="rf-metric">'
            f'<div class="val">{value}</div>'
            f'<div class="lbl">{label}</div>'
            f'{sub_html}</div>')


def vline_trace(x_val, y_min: float, y_max: float,
                color: str, label: str) -> go.Scatter:
    """Vertical line drawn as a Scatter trace — avoids add_vline Timestamp bug."""
    return go.Scatter(
        x=[x_val, x_val],
        y=[y_min, y_max * 1.03],
        mode='lines+text',
        line=dict(color=color, width=1.6, dash='dash'),
        text=['', label],
        textposition='top center',
        textfont=dict(color=color, size=10),
        showlegend=False,
        hoverinfo='skip',
    )


# ═════════════════════════════════════════════════════════════════════════════
# SESSION STATE
# ═════════════════════════════════════════════════════════════════════════════

_LOWER_SUBSTR = ('drop', 'fail', 'bler', 'error', 'util', 'loss', 'block', 'latency', 'delay',
                 'congest', 'retrans', 'abnormal', 'reject', 'timeout', 'interrupt',
                 'unavail', 'downtime', 'outage', 'discard', 'unsuccess', 'lost', 'violation')
# short abbreviations: matched as whole words only (so "fr" never hits "frequency")
_LOWER_TOKENS = {'fr', 'fails', 'failrate', 'dcr', 'cdr', 'ber', 'rlf', 'bler', 'retx', 'ploss'}


def lower_is_better(kpi: str) -> bool:
    """Direction guess from the KPI name (editable in the SLA Scorecard table)."""
    kl = kpi.lower()
    if any(x in kl for x in _LOWER_SUBSTR):
        return True
    return any(t in _LOWER_TOKENS for t in re.findall(r'[a-z]+', kl))


def _fmt_num(v: float) -> str:
    """Compact, readable number: 12,345 · 123.4 · 12.34 · 0.123"""
    if pd.isna(v):
        return "—"
    a = abs(v)
    if a >= 1000: return f"{v:,.0f}"
    if a >= 100:  return f"{v:,.1f}"
    if a >= 1:    return f"{v:,.2f}"
    return f"{v:.3f}"


def sparkline_tile(kpi: str, ts: pd.DataFrame, time_col: str, freq_label: str):
    """Return (header_html, plotly_figure) for one KPI tile.
    Colour is direction-aware: green = improving, red = degrading, grey = flat."""
    import html as _html
    y = ts[kpi].astype(float)
    lb = lower_is_better(kpi)
    last, prev = float(y.iloc[-1]), float(y.iloc[-2])
    pct = (last - prev) / abs(prev) * 100 if prev else (0.0 if last == prev else 100.0 * np.sign(last - prev))
    if abs(pct) < 0.5:
        col, bg, arrow = '#718096', 'rgba(113,128,150,.14)', '▬'
    else:
        good = (pct < 0) if lb else (pct > 0)
        col = '#2fa36b' if good else '#e05a4e'
        bg = 'rgba(47,163,107,.14)' if good else 'rgba(224,90,78,.14)'
        arrow = '▲' if pct > 0 else '▼'
    GOOD, BAD = '#2fa36b', '#e05a4e'
    i_min, i_max = int(y.values.argmin()), int(y.values.argmax())
    i_best, i_worst = (i_min, i_max) if lb else (i_max, i_min)
    lo, hi = float(y.min()), float(y.max())
    # Overall trend over the whole period (least-squares slope, direction-aware)
    n_pts = len(y)
    slope = float(np.polyfit(np.arange(n_pts), y.values, 1)[0]) if n_pts >= 3 else (last - prev)
    base = abs(float(y.mean())) or 1.0
    tot_pct = slope * max(n_pts - 1, 1) / base * 100
    if abs(tot_pct) < 1.0:
        tr_txt, tr_col = '▬ Stable', '#718096'
    else:
        t_good = (slope < 0) if lb else (slope > 0)
        tr_col = '#2fa36b' if t_good else '#e05a4e'
        tr_txt = ('↗ ' if slope > 0 else '↘ ') + ('Improving' if t_good else 'Degrading')
    # Y scale: ratio KPIs fixed 0-100; count/number KPIs 0 → max + headroom
    _ratio_hint = ('%', 'rate', 'ratio', 'sr', 'success', 'avail', 'accessib', 'retain',
                   'drop', 'fail', 'bler', 'ssr', 'util', 'percent', 'cssr', 'dcr')
    is_ratio = (any(h in kpi.lower() for h in _ratio_hint)
                and lo >= -0.001 and hi <= 100.001)
    if is_ratio:
        y0, y1 = 0.0, 100.0
    else:
        y0 = min(0.0, lo * 1.1)
        y1 = (hi * 1.2) if hi > 0 else (abs(hi) * 0.2 or 1.0)
    hover_fmt = '%{x|%d %b %H:%M}' if 'our' in freq_label else '%{x|%d %b}'

    fig = go.Figure()
    # invisible floor so the area is filled to the bottom of the plot, not to zero
    fig.add_trace(go.Scatter(x=ts[time_col], y=[y0] * len(ts), mode='lines',
                             line=dict(width=0), hoverinfo='skip', showlegend=False))
    fig.add_trace(go.Scatter(
        x=ts[time_col], y=y, mode='lines', fill='tonexty', showlegend=False,
        line=dict(color=col, width=2.2), fillcolor=_rgba(col, 0.13),
        hovertemplate=hover_fmt + '<br><b>%{y:,.3f}</b><extra></extra>'))
    fig.add_hline(y=float(y.mean()), line=dict(color='#a0aec0', width=1, dash='dot'))
    pts = [(i_best, GOOD, 6), (i_worst, BAD, 6)]
    if i_best != i_worst:
        fig.add_trace(go.Scatter(
            x=[ts[time_col].iloc[i] for i, _, _ in pts], y=[y.iloc[i] for i, _, _ in pts],
            mode='markers', showlegend=False, hoverinfo='skip',
            marker=dict(size=6, color=[c for _, c, _ in pts],
                        line=dict(width=1, color='#ffffff'))))
    fig.add_trace(go.Scatter(                               # latest point, emphasised
        x=[ts[time_col].iloc[-1]], y=[last], mode='markers', showlegend=False, hoverinfo='skip',
        marker=dict(size=9, color=col, line=dict(width=2, color='#ffffff'))))
    fig.update_layout(
        height=230, margin=dict(l=40, r=10, t=6, b=28), showlegend=False,
        xaxis=dict(visible=True, fixedrange=True, showgrid=False, tickfont=dict(size=10),
                   tickformat='%d %b %H:%M' if 'our' in freq_label else '%d %b', nticks=5),
        yaxis=dict(visible=True, range=[y0, y1], fixedrange=True, tickfont=dict(size=10),
                   gridcolor='#eef2f6', nticks=5),
        plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', hovermode='x')

    header = (f'<div class="rf-sp-name" title="{_html.escape(kpi)}">{_html.escape(kpi)}</div>'
              f'<div class="rf-sp-row"><span class="rf-sp-val">{_fmt_num(last)}</span>'
              f'<span class="rf-sp-delta" style="color:{col};background:{bg}">'
              f'{arrow} {pct:+.1f}%</span></div>'
              f'<div class="rf-sp-meta" style="font-weight:700;font-size:.78rem;color:{tr_col}">'
              f'{tr_txt} over period ({tot_pct:+.1f}%)</div>'
              f'<div class="rf-sp-meta">avg {_fmt_num(float(y.mean()))} · '
              f'min {_fmt_num(lo)} · max {_fmt_num(hi)} · {len(y)} pts</div>')
    return header, fig


def init_state():
    defaults = {
        'confirmed_hash':    '',
        'type_overrides':    {},   # col → type (from user clarification)
        'nl_history':        [],
        'anthropic_api_key': '',
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


# ═════════════════════════════════════════════════════════════════════════════
# MAIN APPLICATION
# ═════════════════════════════════════════════════════════════════════════════

def main():
    init_state()

    # ── Header ───────────────────────────────────────────────────────────────
    st.markdown("""
    <div class="rf-header">
        <h1>📡 RAN KPI Analytics Tool<span class="rf-badge">PRO · v13</span></h1>
        <p>Universal vendor-agnostic RAN performance analytics —
           Nokia · Ericsson · Huawei · ZTE · Samsung</p>
    </div>
    """, unsafe_allow_html=True)

    # ── File Upload ───────────────────────────────────────────────────────────
    with st.expander("📂 Upload Data Files", expanded=True):
        u1, u2 = st.columns(2)
        with u1:
            st.markdown("#### 📊 KPI Data Files *(required)*")
            kpi_files = st.file_uploader(
                "CSV / XLS / XLSX / TSV",
                type=['csv','xls','xlsx','txt','tsv'],
                accept_multiple_files=True, key='kpi_up')
        with u2:
            st.markdown("#### 🗺️ Planning / Site Files *(optional)*")
            st.caption("Lat/lon auto-extracted from KPI file when omitted. "
                       "KML / KMZ placemarks are supported (matched by name).")
            plan_files = st.file_uploader(
                "CSV / XLS / XLSX / TSV / KML / KMZ",
                type=['csv','xls','xlsx','txt','tsv','kml','kmz'],
                accept_multiple_files=True, key='plan_up')
            bundled = find_bundled_planning()
            if bundled:
                use_bundled = st.checkbox(
                    f"Use bundled planning file · {os.path.basename(bundled)}",
                    value=True, key='use_bundled',
                    help="Found next to the script. Ignored when you upload your own "
                         "planning CSV/XLSX.")
                has_upload = any(not f.name.lower().endswith(('.kml', '.kmz'))
                                 for f in (plan_files or []))
                plan_files = list(plan_files or [])
                if use_bundled and not has_upload:
                    plan_files.append(_LocalFile(bundled))

    if not kpi_files:
        st.info("👆 Upload at least one KPI file to begin.")
        return

    # ── Load KPI data ─────────────────────────────────────────────────────────
    with st.spinner("Loading KPI files…"):
        kpi_frames, kpi_fmts = [], []
        for f in kpi_files:
            df_tmp, fmt = load_file(f)
            if df_tmp is not None and not df_tmp.empty:
                kpi_frames.append(df_tmp)
                kpi_fmts.append(fmt)

    if not kpi_frames:
        st.error("❌ Could not load any KPI data."); return

    df_kpi = pd.concat(kpi_frames, ignore_index=True)
    st.success(f"✅ Loaded {len(kpi_files)} KPI file(s): {', '.join(kpi_fmts)}")

    # ── Load Planning data ────────────────────────────────────────────────────
    df_plan  = None
    plan_src = "none"
    kml_frames = []
    plan_info = {'forced': {}, 'hints': {}, 'filters': [], 'notes': [], 'date': None}

    if plan_files:
        with st.spinner("Loading planning files…"):
            plan_frames, plan_fmts = [], []
            for f in plan_files:
                if f.name.lower().endswith(('.kml', '.kmz')):
                    df_tmp, fmt = load_kml(f)
                    if df_tmp is not None and not df_tmp.empty:
                        kml_frames.append(df_tmp)
                else:
                    raw = f.getvalue() if hasattr(f, 'getvalue') else f.read()
                    df_tmp, fmt, inf_tmp = load_planning_cached(f.name, raw)
                    if inf_tmp:
                        plan_info['forced'].update(inf_tmp['forced'])
                        plan_info['hints'].update(inf_tmp['hints'])
                        plan_info['filters'] += [c for c in inf_tmp['filters']
                                                 if c not in plan_info['filters']]
                        plan_info['notes'] += inf_tmp['notes']
                        plan_info['date'] = inf_tmp['date'] or plan_info['date']
                if df_tmp is not None and not df_tmp.empty:
                    plan_frames.append(df_tmp)
                    plan_fmts.append(fmt)
        if plan_frames:
            df_plan  = pd.concat(plan_frames, ignore_index=True)
            plan_src = "file"
            st.success(f"✅ Loaded {len(plan_files)} planning file(s): {', '.join(plan_fmts)}"
                       + (f" · planning snapshot {plan_info['date']}" if plan_info['date'] else ""))
            for _n in plan_info['notes']:
                st.caption("ℹ️ " + _n)

    # ── Detect column types ───────────────────────────────────────────────────
    with st.spinner("Detecting column types…"):
        kpi_types,  kpi_conf  = detect_all_columns(df_kpi)
        plan_types, plan_conf = (detect_all_columns(df_plan)
                                 if df_plan is not None else ({}, {}))

    # Planning-file columns with a known meaning (Nokia 5G cell parameters)
    for _c, _t in plan_info['forced'].items():
        if _c in plan_types:
            plan_types[_c], plan_conf[_c] = _t, 1.0

    # KML-derived columns have fixed meaning
    for _c, _t in (('Latitude', 'latitude'), ('Longitude', 'longitude'),
                   ('KML_Name', 'identifier')):
        if kml_frames and _c in plan_types:
            plan_types[_c] = _t

    # Apply any saved user overrides
    for col, t in st.session_state.type_overrides.items():
        if col in kpi_types:  kpi_types[col]  = t
        if col in plan_types: plan_types[col] = t

    # ── Auto-extract geo from KPI if no planning file ─────────────────────────
    if plan_src == "none":
        lat_candidates  = [c for c, t in kpi_types.items() if t == 'latitude'  and c in df_kpi.columns]
        lon_candidates  = [c for c, t in kpi_types.items() if t == 'longitude' and c in df_kpi.columns]
        cell_candidates = [c for c, t in kpi_types.items() if t == 'cell_id'   and c in df_kpi.columns]
        site_candidates = [c for c, t in kpi_types.items() if t == 'site_id'   and c in df_kpi.columns]

        if lat_candidates and lon_candidates:
            keep = lat_candidates[:1] + lon_candidates[:1] + cell_candidates[:1] + site_candidates[:1]
            df_plan = (df_kpi[keep]
                       .drop_duplicates()
                       .dropna(subset=lat_candidates[:1] + lon_candidates[:1])
                       .reset_index(drop=True))
            plan_types = {
                lat_candidates[0]: 'latitude',
                lon_candidates[0]: 'longitude',
            }
            if cell_candidates: plan_types[cell_candidates[0]] = 'cell_id'
            if site_candidates: plan_types[site_candidates[0]] = 'site_id'
            plan_src = "auto"
            st.info(f"📍 Auto-detected geo data from KPI file — {len(df_plan):,} unique locations")

    # ── Column detection details ──────────────────────────────────────────────
    with st.expander("🔬 Column Detection Details", expanded=False):
        d1, d2 = st.columns(2)
        with d1:
            st.markdown("**KPI File Columns**")
            st.dataframe(pd.DataFrame({
                'Column': list(kpi_types),
                'Type':   list(kpi_types.values()),
                'Conf':   [f"{kpi_conf.get(c,0):.0%}" for c in kpi_types],
            }), use_container_width=True, height=260)
        with d2:
            if plan_types:
                lbl = 'Planning File' if plan_src == 'file' else 'Auto-extracted Geo'
                st.markdown(f"**{lbl} Columns**")
                st.dataframe(pd.DataFrame({
                    'Column': list(plan_types),
                    'Type':   list(plan_types.values()),
                    'Conf':   [f"{plan_conf.get(c,0):.0%}" for c in plan_types],
                }), use_container_width=True, height=260)

    # ── Clarification — only truly ambiguous non-numeric string columns ────────
    current_hash = file_sig(kpi_files)
    if current_hash != st.session_state.confirmed_hash:
        ambiguous = [
            col for col, conf in kpi_conf.items()
            if conf < 0.70
            and kpi_types.get(col) not in {'numeric_kpi','date','latitude','longitude','duration'}
            and col in df_kpi.columns
            and not pd.api.types.is_numeric_dtype(df_kpi[col])
            and df_kpi[col].notna().sum() > 0
        ]
        if ambiguous:
            st.warning(f"⚠️ {len(ambiguous)} column(s) need clarification.")
            with st.form("clarify"):
                answers = {}
                opts = ['Cell ID','Site ID','RNC','PLMN','Date/Time','Duration',
                        'Latitude','Longitude','KPI Metric','Identifier','Ignore']
                for col in ambiguous:
                    s = df_kpi[col].dropna().astype(str).head(3).tolist()
                    answers[col] = st.selectbox(
                        f"**`{col}`** ({kpi_types.get(col,'?')} · {kpi_conf.get(col,0):.0%}) "
                        f"— samples: `{', '.join(s)}`",
                        opts, key=f"cl_{col}")
                if st.form_submit_button("✅ Apply & Remember"):
                    mapping = {
                        'Cell ID':'cell_id','Site ID':'site_id','RNC':'rnc_id',
                        'PLMN':'plmn','Date/Time':'date','Duration':'duration',
                        'Latitude':'latitude','Longitude':'longitude',
                        'KPI Metric':'numeric_kpi','Identifier':'identifier',
                        'Ignore':'ignore',
                    }
                    overrides = {col: mapping[ans] for col, ans in answers.items()}
                    st.session_state.type_overrides.update(overrides)
                    st.session_state.confirmed_hash = current_hash
                    st.rerun()
        else:
            # No ambiguous columns — mark this file as confirmed
            st.session_state.confirmed_hash = current_hash

    # ── Merge KPI + Planning ──────────────────────────────────────────────────
    df = df_kpi.copy()

    if df_plan is not None and plan_src == "file":
        id_types = {'cell_id', 'site_id', 'identifier', 'rnc_id'}
        geo_cols = {c for c, t in plan_types.items() if t in ('latitude', 'longitude')}
        # KPI-side key candidates: id-like columns + any text / integer column
        kpi_keys = [c for c in df_kpi.columns
                    if kpi_types.get(c) in id_types
                    or (kpi_types.get(c) not in ('date', 'duration', 'latitude', 'longitude')
                        and (not pd.api.types.is_numeric_dtype(df_kpi[c])
                             or pd.api.types.is_integer_dtype(df_kpi[c])))][:25]
        plan_id_types = {**plan_types, **plan_info['hints']}
        plan_keys = sorted((c for c in df_plan.columns if c not in geo_cols),
                           key=lambda c: (c not in plan_info['hints'], 0))[:25]

        def _do_merge(kk, pk, mapping):
            left = df_kpi.copy()
            left['_jk'] = _norm_key(left[kk].values).map(mapping).values
            right = df_plan.copy()
            right['_jk'] = _norm_key(right[pk].values).values
            right = right.drop_duplicates('_jk')            # never multiply KPI rows
            return left.merge(right, on='_jk', how='left',
                              suffixes=('', '_plan')).drop(columns=['_jk'])

        # Preferred pairs first (id-typed on both sides), then everything else
        pref = [(k, q) for k in kpi_keys if kpi_types.get(k) in id_types
                for q in plan_keys if plan_id_types.get(q) in id_types]
        rest = [(k, q) for k in kpi_keys for q in plan_keys if (k, q) not in pref]
        def _search(pairs, fuzzy):
            out = []
            for k, q in pairs:
                mp, how = smart_map(df_kpi[k].dropna().unique(), df_plan[q].dropna().unique(), fuzzy)
                if mp:
                    out.append((k, q, mp, how))
            return out

        with st.spinner("Matching planning/KML names to KPI ids…"):
            # id-typed pairs first, cheap exact/punctuation matching before any fuzzy matching
            results = (_search(pref, False) or _search(rest, False)
                       or _search(pref, True) or _search(rest, True))

        # Prefer exact matches, then the most *unique* planning key: a cell-level key
        # (one row per cell) keeps per-sector azimuths; a site-level key would give
        # every cell of a site the same row.
        def _uniq(q):
            nn = df_plan[q].notna().sum()
            return df_plan[q].nunique() / nn if nn else 0.0
        results.sort(key=lambda r: ({"exact": 0, "ignoring punctuation": 1}.get(r[3], 2),
                                    -round(_uniq(r[1]), 2), -len(r[2])))

        if results:
            opts = [f"{k} ↔ {q} · {len(mp):,} ids · {how}" for k, q, mp, how in results[:8]]
            sel = st.radio("🔗 Join key:", opts, index=0)
            k, q, mp, how = results[opts.index(sel)]
            df = _do_merge(k, q, mp)
            st.success(f"✅ Joined **{k}** ↔ **{q}** ({how}) — {len(mp):,} KPI ids matched")
        else:
            st.warning("⚠️ Planning/KML file loaded but no names match any KPI column. "
                       "Compare the samples below and pick the columns manually.")
            c1, c2 = st.columns(2)
            with c1:
                mk = st.selectbox("KPI column", kpi_keys or list(df_kpi.columns), key='man_k')
                st.caption("KPI samples: " + ", ".join(
                    df_kpi[mk].dropna().astype(str).unique()[:6]))
            with c2:
                mq = st.selectbox("Planning / KML column", plan_keys, key='man_q')
                st.caption("Planning samples: " + ", ".join(
                    df_plan[mq].dropna().astype(str).unique()[:6]))
            mp, how = smart_map(df_kpi[mk].dropna().unique(), df_plan[mq].dropna().unique())
            if mp:
                df = _do_merge(mk, mq, mp)
                st.success(f"✅ Joined **{mk}** ↔ **{mq}** ({how}) — {len(mp):,} ids")
            else:
                st.error("Still no overlap for that pair. The KML names are probably "
                         "a different naming scheme (e.g. site names vs. cell IDs); "
                         "you need a mapping file (CSV with cell ID + Lat/Lon).")

    # ── Resolve all column roles ───────────────────────────────────────────────
    all_types = {**kpi_types, **plan_types}

    time_col     = find_time_col(all_types, df)
    cell_id_cols = [c for c, t in all_types.items() if t == 'cell_id'  and c in df.columns]
    site_id_cols = [c for c, t in all_types.items() if t == 'site_id'  and c in df.columns]
    rnc_cols     = [c for c, t in all_types.items() if t == 'rnc_id'   and c in df.columns]
    plmn_cols    = [c for c, t in all_types.items() if t == 'plmn'     and c in df.columns]
    dur_cols     = [c for c, t in all_types.items() if t == 'duration' and c in df.columns]
    lat_col      = next((c for c, t in all_types.items() if t == 'latitude'  and c in df.columns), None)
    lon_col      = next((c for c, t in all_types.items() if t == 'longitude' and c in df.columns), None)

    if plan_src == "file":      # planning coordinates take precedence over KPI-file ones
        _pla = next((c for c, t in plan_types.items() if t == 'latitude'  and c in df.columns), None)
        _plo = next((c for c, t in plan_types.items() if t == 'longitude' and c in df.columns), None)
        if _pla and _plo:
            lat_col, lon_col = _pla, _plo

    def _pcol(name):
        """Actual df column of a planning attribute (handles the _plan suffix)."""
        if plan_src != "file":
            return None
        real = name + '_plan' if name in df_kpi.columns else name
        return real if real in df.columns else None

    az_col = _pcol('AZIMUTH')

    # Parse & sort by time
    if time_col:
        df[time_col] = parse_date_column(df[time_col])
        df = df.dropna(subset=[time_col]).sort_values(time_col).reset_index(drop=True)

    # Build KPI column list — union of numeric dtypes + detector-tagged numeric_kpi
    _num_by_dtype = set(df.select_dtypes(include=np.number).columns)
    _num_by_tag   = {c for c, t in all_types.items() if t == 'numeric_kpi' and c in df.columns}
    _exclude      = (
        set(cell_id_cols) | set(site_id_cols) | set(rnc_cols) |
        set(plmn_cols)    | set(dur_cols)      |
        ({lat_col} if lat_col else set()) |
        ({lon_col} if lon_col else set()) |
        ({time_col} if time_col else set()) |
        {c for c in df.columns if 'ANOMALY' in c.upper()}
    )
    if plan_src == "file":      # PMAX, TAC, PHYSCELLID, … are attributes, not KPIs
        _exclude = _exclude | {c for c in df.columns if c not in df_kpi.columns}
    _all_kpi_candidates = list(_num_by_dtype | _num_by_tag)
    kpi_cols = [c for c in df.columns   # preserve original column order
                if c in _all_kpi_candidates and c not in _exclude]

    # ── Coerce all KPI cols to numeric — fixes 'float + str' crash ──────────
    # All-null Nokia KPI columns stay as object dtype after loading.
    # .agg(['mean','std',...]) crashes on object columns → force numeric.
    for _kc in list(kpi_cols):
        if not pd.api.types.is_numeric_dtype(df[_kc]):
            df[_kc] = pd.to_numeric(df[_kc], errors='coerce')
    # Also apply to df_f after filters (done lazily below via df reference)


    # ── SIDEBAR ───────────────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown("## 🎛️ Filters & Settings")

        def _qa(action):
            if action == 'reset':
                for _k in list(st.session_state.keys()):
                    if _k.startswith('f_'):
                        st.session_state[_k] = []
                st.session_state['ksel'] = kpi_cols[:min(5, len(kpi_cols))]
            elif action == 'all':
                st.session_state['ksel'] = list(kpi_cols)
            elif action == 'top5':
                st.session_state['ksel'] = kpi_cols[:min(5, len(kpi_cols))]
            elif action == 'none':
                st.session_state['ksel'] = []

        st.markdown("### ⚡ Quick Actions")
        qa1, qa2 = st.columns(2)
        qa1.button("↺ Reset filters", key='qa_reset', on_click=_qa, args=('reset',), use_container_width=True)
        qa2.button("✓ All KPIs", key='qa_all', on_click=_qa, args=('all',), use_container_width=True)
        qa3, qa4 = st.columns(2)
        qa3.button("★ First 5", key='qa_top5', on_click=_qa, args=('top5',), use_container_width=True)
        qa4.button("✕ Clear KPIs", key='qa_none', on_click=_qa, args=('none',), use_container_width=True)

        # Granularity
        st.markdown("### ⏱️ Granularity")
        gran_label = st.selectbox("gran", list(GRANULARITY), index=2,
                                   key='gran', label_visibility='collapsed')
        gran_freq  = GRANULARITY[gran_label]

        # Time period
        d_start = d_end = None
        if time_col:
            st.markdown("### 📅 Time Period")
            t_min = df[time_col].min().date()
            t_max = df[time_col].max().date()
            dr = st.date_input("dr", value=(t_min, t_max),
                               min_value=t_min, max_value=t_max,
                               key='dr', label_visibility='collapsed')
            if isinstance(dr, (list, tuple)) and len(dr) == 2:
                d_start, d_end = dr[0], dr[1]
            else:
                d_start, d_end = t_min, t_max

        # Network hierarchy filters
        st.markdown("### 📡 Network Filters")
        # Cascading hierarchy: every filter only offers values that exist under
        # the selections made in the filters above it (PLMN → RNC → site(s) → cell(s)).
        order = {c: i for i, c in enumerate(df.columns)}
        H = []
        for _pc, _pl in PLAN_FILTER_LABELS.items():      # Region → City → District → Tech → Band
            if _pc in plan_info['filters'] and _pcol(_pc):
                H.append((_pl, _pcol(_pc)))
        for lbl0, cols0 in (('PLMN', plmn_cols), ('RNC', rnc_cols)):
            c0 = next((c for c in cols0 if c in df.columns), None)
            if c0:
                H.append((lbl0, c0))
        for c in sorted(site_id_cols, key=lambda x: order.get(x, 0)):
            if c in df.columns:
                H.append((c, c))
        for c in sorted(cell_id_cols, key=lambda x: order.get(x, 0)):
            if c in df.columns:
                H.append((c, c))

        active_filters: dict = {}
        filter_col_map: dict = {}
        scope = pd.Series(True, index=df.index)
        for lbl, col in H:
            if col in filter_col_map.values():
                continue
            vals = sorted(df.loc[scope, col].dropna().astype(str).unique())
            if not vals:
                continue
            kkey = f'f_{lbl}'
            # drop stale selections that no longer exist under the parent choice
            if kkey in st.session_state:
                st.session_state[kkey] = [v for v in st.session_state[kkey] if v in vals]
            sel = st.multiselect(lbl, vals, default=[],
                                 placeholder=f"All ({len(vals):,})",
                                 key=kkey)
            filter_col_map[lbl] = col
            active_filters[lbl] = sel
            if sel:
                scope &= df[col].astype(str).isin(sel)

        # KPI selection
        st.markdown("### 📊 KPI Selection")
        selected_kpis = st.multiselect(
            "kpis", kpi_cols,
            default=list(kpi_cols),
            key='ksel', label_visibility='collapsed')

        # Health thresholds
        st.markdown("### 🏥 Health Thresholds")
        with st.expander("Configure", expanded=False):
            user_thresholds = {}
            for pat, (deg, crit) in DEFAULT_KPI_THRESHOLDS.items():
                matching = [k for k in kpi_cols if pat in k.lower()]
                if matching:
                    lbl2 = matching[0]
                    dv = st.number_input(f"{lbl2} — Degraded",
                                         value=float(deg), key=f'td_{pat}', step=0.5)
                    cv = st.number_input(f"{lbl2} — Critical",
                                         value=float(crit), key=f'tc_{pat}', step=0.5)
                    user_thresholds[pat] = (dv, cv)
            if not user_thresholds:
                user_thresholds = DEFAULT_KPI_THRESHOLDS

        # Anomaly
        st.markdown("### 🤖 Anomaly Sensitivity")
        contamination = st.slider("contam", 0.01, 0.20, 0.05, 0.01,
                                   key='contam', label_visibility='collapsed')

        # Forecast
        st.markdown("### 📈 Forecast Horizon")
        fcast_days = st.selectbox("fh", [7, 14, 30], index=1,
                                   key='fh', label_visibility='collapsed')

        st.markdown("---")
        st.caption("All filters apply to every tab.")

    # ── Apply filters ─────────────────────────────────────────────────────────
    df_f = df.copy()

    if time_col and d_start and d_end:
        df_f = df_f[
            (df_f[time_col].dt.date >= d_start) &
            (df_f[time_col].dt.date <= d_end)
        ]

    for lbl, sel in active_filters.items():
        if sel:
            col = filter_col_map[lbl]
            df_f = df_f[df_f[col].astype(str).isin(sel)]

    if df_f.empty:
        st.warning("⚠️ No data matches the current filters. Adjust the sidebar.")
        return

    # Resampled version for trend charts
    id_cols  = [c for c in (cell_id_cols[:1] + site_id_cols[:1]) if c in df_f.columns]
    df_plot  = (resample(df_f, time_col, id_cols, kpi_cols, gran_freq)
                if gran_freq and time_col else df_f.copy())

    if not selected_kpis:
        st.warning("Select at least one KPI from the sidebar.")
        return

    # Filter summary for exports / PDF
    active_parts = [f"{l}: {', '.join(v)}" for l, v in active_filters.items() if v]
    filter_str   = (', '.join(active_parts) if active_parts else "All network elements")
    filter_str  += f" | {gran_label}"
    date_str     = f"{d_start} → {d_end}" if d_start else "Full dataset"

    entity_col = next((c for c in (cell_id_cols + site_id_cols) if c in df_f.columns), None)
    chips = [f"📄 {len(df_f):,} rows"]
    if entity_col:
        chips.append(f"📡 {df_f[entity_col].nunique():,} {entity_col}")
    chips += [f"📊 {len(kpi_cols)} KPIs", f"📅 {date_str}", f"⏱️ {gran_label}"]
    if any(v for v in active_filters.values()):
        chips.append("🔎 filters active")
    st.markdown('<div class="rf-chips">' + ''.join(f'<span class="rf-chip">{c}</span>' for c in chips)
                + '</div>', unsafe_allow_html=True)

    # ── TABS ──────────────────────────────────────────────────────────────────
    TAB_NAMES = [
        "📊 Overview", "🏥 Health", "📈 Trends",
        "📉 Worst Cells", "🔗 Correlation", "🗺️ Map",
        "🤖 Anomaly", "↔️ Period Compare", "🔍 Change Points",
        "💬 AI Query",
        "🏆 SLA Scorecard", "🎯 Drill-down", "📊 Distribution",
        "⛔ Chronic & Movers", "🧬 Clustering", "🧪 Data Quality",
        "📥 Export", "📡 Planning",
    ]
    tabs = st.tabs(TAB_NAMES)

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 0 — OVERVIEW
    # ══════════════════════════════════════════════════════════════════════════
    with tabs[0]:
        # Metric cards
        c1, c2, c3, c4, c5, c6 = st.columns(6)
        with c1:
            st.markdown(metric_card("Filtered Records", f"{len(df_f):,}",
                                    f"of {len(df):,}"), unsafe_allow_html=True)
        with c2:
            n_cells = df_f[cell_id_cols[0]].nunique() if cell_id_cols else "—"
            st.markdown(metric_card("Unique Cells", str(n_cells)),
                        unsafe_allow_html=True)
        with c3:
            n_sites = df_f[site_id_cols[0]].nunique() if site_id_cols else "—"
            st.markdown(metric_card("Unique Sites", str(n_sites)),
                        unsafe_allow_html=True)
        with c4:
            if time_col:
                span = (df_f[time_col].max() - df_f[time_col].min()).days
                st.markdown(metric_card("Date Span", f"{span}d",
                                        f"{d_start} → {d_end}" if d_start else ""),
                            unsafe_allow_html=True)
            else:
                st.markdown(metric_card("Date Span", "—"), unsafe_allow_html=True)
        with c5:
            st.markdown(metric_card("KPIs Available", str(len(kpi_cols)),
                                    f"{len(selected_kpis)} selected"),
                        unsafe_allow_html=True)
        with c6:
            n_active = sum(1 for s in active_filters.values() if s)
            st.markdown(metric_card("Active Filters", str(n_active)),
                        unsafe_allow_html=True)

        st.markdown("---")

        # KPI summary table — one column per date (hour if span < 2 days)
        st.markdown("### 📋 KPI Summary by Date")
        if selected_kpis and time_col:
            _span = (df_f[time_col].max() - df_f[time_col].min()).total_seconds() / 86400
            _freq, _fmt = ('h', '%d-%b %H:%M') if _span < 2 else ('D', '%d-%b-%y')
            _dated = (df_f.set_index(time_col)[selected_kpis]
                      .apply(pd.to_numeric, errors='coerce')
                      .resample(_freq).mean().round(4))
            _dated = _dated.dropna(how='all')
            _dated.index = _dated.index.strftime(_fmt)
            summary = _dated.T                      # rows = KPIs, columns = dates
            summary.index.name = 'KPI'

            def _row_color(row):
                """Red→green per KPI row (reversed for lower-is-better KPIs)."""
                v = pd.to_numeric(row, errors='coerce')
                lo, hi = v.min(), v.max()
                if pd.isna(lo) or hi == lo:
                    return [''] * len(row)
                t = (v - lo) / (hi - lo)
                if lower_is_better(row.name):
                    t = 1 - t
                out = []
                for x in t:
                    if pd.isna(x):
                        out.append('')
                    else:
                        r = int(248 - 120 * x); g = int(150 + 80 * x); bl = int(150 - 20 * x)
                        out.append(f'background-color: rgb({r},{g},{bl})')
                return out

            st.dataframe(
                summary.style.apply(_row_color, axis=1)
                       .format("{:.2f}", na_rep="—"),
                use_container_width=True,
            )
        elif selected_kpis:
            st.info("No time column detected — cannot show KPIs by date.")
            summary = df_f[selected_kpis].agg(['mean']).T.round(4)

        # Sparklines — KPI tiles: latest value, change vs previous period, trend line
        if time_col and selected_kpis:
            st.markdown("### 📈 KPI Trends at a Glance")
            span_days = (df_f[time_col].max() - df_f[time_col].min()).total_seconds() / 86400
            sp_freq, sp_label = ('h', 'Hourly') if span_days < 2 else ('D', 'Daily')
            show = list(selected_kpis)
            st.caption(f"{sp_label} average · value = latest period, badge = change vs previous period "
                       "(green improving, red degrading, direction-aware) · "
                       "dotted line = period average · green/red dots = best/worst point"
                       + (f" · showing {len(show)} of {len(selected_kpis)} selected KPIs" if len(selected_kpis) > len(show) else ""))
            for row0 in range(0, len(show), 3):
                row_k = show[row0:row0 + 3]
                sp_cols = st.columns(3)
                for i, kpi in enumerate(row_k):
                    with sp_cols[i]:
                        spark = (df_f.groupby(pd.Grouper(key=time_col, freq=sp_freq))[kpi]
                                     .mean().dropna().reset_index())
                        try:
                            box = st.container(border=True)
                        except TypeError:                    # older Streamlit
                            box = st.container()
                        with box:
                            if len(spark) >= 2:
                                hdr, fig = sparkline_tile(kpi, spark, time_col, sp_label)
                                st.markdown(hdr, unsafe_allow_html=True)
                                st.plotly_chart(fig, use_container_width=True,
                                                config={'displayModeBar': False},
                                                key=f"spark_{row0 + i}_{kpi}")
                            else:
                                st.markdown(f'<div class="rf-sp-name">{kpi}</div>'
                                            '<div class="rf-sp-meta">Not enough data points '
                                            'for a trend (need ≥ 2 periods).</div>',
                                            unsafe_allow_html=True)

        # Active filter summary
        if any(v for v in active_filters.values()):
            st.markdown("### 🔍 Active Filters")
            for lbl, sel in active_filters.items():
                if sel:
                    st.markdown(f"**{lbl}:** {', '.join(str(s) for s in sel)}")

        with st.expander("👁️ Raw Data Preview (first 200 rows)", expanded=False):
            st.dataframe(df_f.head(200), use_container_width=True)

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 1 — HEALTH
    # ══════════════════════════════════════════════════════════════════════════
    with tabs[1]:
        st.markdown("## 🏥 RAN Cell Health Scoring")
        cell_col_h = cell_id_cols[0] if cell_id_cols else None
        hdf = health_score(df_f, selected_kpis, cell_col_h, user_thresholds)

        if hdf.empty:
            st.info("Health scoring requires a Cell ID column and KPI thresholds.")
        else:
            cts = hdf['Status'].value_counts()
            h1, h2, h3 = st.columns(3)
            with h1: st.metric("🟢 Healthy",  cts.get("🟢 Healthy",  0))
            with h2: st.metric("🟡 Degraded", cts.get("🟡 Degraded", 0))
            with h3: st.metric("🔴 Critical", cts.get("🔴 Critical", 0))

            fig_pie = px.pie(
                values=cts.values, names=cts.index, hole=0.55,
                color=cts.index,
                color_discrete_map={
                    "🟢 Healthy":  "#27a99a",
                    "🟡 Degraded": "#f59e0b",
                    "🔴 Critical": "#e05a4e",
                },
                template=RF_TPL,
                title="Cell Health Distribution",
            )
            fig_pie.update_traces(textinfo='label+percent')
            st.plotly_chart(fig_pie, use_container_width=True)

            st.markdown("### Cell Health Detail")
            st.dataframe(
                hdf.reset_index().sort_values('Score'),
                use_container_width=True,
            )

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 2 — TRENDS & FORECAST
    # ══════════════════════════════════════════════════════════════════════════
    with tabs[2]:
        st.markdown(f"## 📈 KPI Trends  ·  *{gran_label}*")
        t1, t2, t3 = st.tabs(["Line Charts", "Box Plots",
                               f"🔮 Forecast (+{fcast_days}d)"])

        with t1:
            if not time_col:
                st.info("No time column detected.")
            else:
                color_by = id_cols[0] if id_cols else None
                for kpi in selected_kpis:
                    fig = px.line(df_plot, x=time_col, y=kpi, color=color_by,
                                  title=kpi, template=RF_TPL)
                    fig.update_layout(hovermode='x unified',
                                      yaxis_title=kpi, xaxis_title=None)
                    st.plotly_chart(fig, use_container_width=True)

        with t2:
            for kpi in selected_kpis:
                fig = px.box(df_f,
                             x=id_cols[0] if id_cols else None,
                             y=kpi,
                             title=f"{kpi} — Distribution",
                             template=RF_TPL, points=False)
                st.plotly_chart(fig, use_container_width=True)

        with t3:
            if not time_col:
                st.info("Forecasting requires a time column.")
            else:
                fkpi = st.selectbox("KPI to forecast", selected_kpis, key='fkpi')
                hist, fcast = linear_forecast(df_f, time_col, fkpi, fcast_days)
                if fcast.empty:
                    st.warning("Need at least 5 daily data points for forecasting.")
                else:
                    fig_f = go.Figure()
                    fig_f.add_trace(go.Scatter(
                        x=hist[time_col], y=hist[fkpi], name='Historical',
                        line=dict(color=TEAL, width=2)))
                    fig_f.add_trace(go.Scatter(
                        x=fcast[time_col], y=fcast[fkpi], name='Forecast',
                        line=dict(color=AMBER, width=2, dash='dash')))
                    fig_f.add_trace(go.Scatter(
                        x=pd.concat([fcast[time_col], fcast[time_col][::-1]]),
                        y=pd.concat([fcast['upper'], fcast['lower'][::-1]]),
                        fill='toself',
                        fillcolor='rgba(245,158,11,0.15)',
                        line=dict(color='rgba(0,0,0,0)'),
                        name='95% CI',
                    ))
                    fig_f.update_layout(
                        title=f"{fkpi} — {fcast_days}-Day Linear Forecast",
                        template=RF_TPL, hovermode='x unified',
                    )
                    st.plotly_chart(fig_f, use_container_width=True)
                    st.caption("⚠️ Linear extrapolation — directional guidance only.")

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 3 — WORST CELLS
    # ══════════════════════════════════════════════════════════════════════════
    with tabs[3]:
        st.markdown("## 📉 Worst Cell Ranking")

        if not cell_id_cols:
            st.info("No Cell ID column detected.")
        else:
            cell_col_w = cell_id_cols[0]
            wc1, wc2, wc3 = st.columns([2, 1, 1])
            with wc1:
                rank_kpis = st.multiselect(
                    "KPIs to rank by", selected_kpis,
                    default=selected_kpis[:min(3, len(selected_kpis))],
                    key='wc_kpis')
            with wc2:
                rank_n = st.slider("Show worst N", 5, 100, 20, key='wc_n')
            with wc3:
                rank_mode = st.radio("Mode", ["Composite", "Single KPI"], key='wc_mode')

            if rank_kpis:
                cell_avg = (df_f.groupby(cell_col_w, observed=True)[rank_kpis]
                            .mean().round(4).reset_index())

                if rank_mode == "Single KPI":
                    single = st.selectbox("KPI", rank_kpis, key='wc_single')
                    worse_high = any(x in single.lower()
                                     for x in ['drop','fail','bler','error','util'])
                    worst = (cell_avg.nlargest(rank_n, single) if worse_high
                             else cell_avg.nsmallest(rank_n, single))
                    sort_col = single
                else:
                    norm = cell_avg[rank_kpis].copy()
                    for kpi in rank_kpis:
                        rng = norm[kpi].max() - norm[kpi].min()
                        if rng > 0:
                            norm[kpi] = (norm[kpi] - norm[kpi].min()) / rng
                        else:
                            norm[kpi] = 0.0
                        if not any(x in kpi.lower() for x in ['drop','fail','bler']):
                            norm[kpi] = 1.0 - norm[kpi]
                    cell_avg['Composite'] = norm[rank_kpis].mean(axis=1).round(4)
                    worst    = cell_avg.nlargest(rank_n, 'Composite')
                    sort_col = 'Composite'

                st.dataframe(
                    worst.style.background_gradient(
                        subset=[sort_col] if sort_col in worst.columns else [],
                        cmap='RdYlGn_r'),
                    use_container_width=True,
                )

                fig_bar = px.bar(
                    worst.head(20), x=cell_col_w, y=sort_col,
                    color=sort_col, color_continuous_scale='RdYlGn_r',
                    title=f"Worst {min(rank_n, 20)} Cells — {sort_col}",
                    template=RF_TPL,
                )
                fig_bar.update_layout(xaxis_tickangle=-45, showlegend=False,
                                      coloraxis_showscale=False)
                st.plotly_chart(fig_bar, use_container_width=True)

                # Trend comparison of worst cells
                if time_col and len(worst) > 0:
                    st.markdown("### 📊 Worst Cell Trends")
                    trend_kpi   = st.selectbox("Trend KPI", rank_kpis, key='wc_trend')
                    worst_names = worst[cell_col_w].astype(str).tolist()[:10]
                    tdf = df_plot[df_plot[cell_col_w].astype(str).isin(worst_names)]
                    if not tdf.empty:
                        fig_t = px.line(tdf, x=time_col, y=trend_kpi,
                                        color=cell_col_w,
                                        title=f"{trend_kpi} — Worst Cells",
                                        template=RF_TPL)
                        fig_t.update_layout(hovermode='x unified')
                        st.plotly_chart(fig_t, use_container_width=True)

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 4 — CORRELATION
    # ══════════════════════════════════════════════════════════════════════════
    with tabs[4]:
        st.markdown("## 🔗 KPI Correlation Matrix")
        if len(selected_kpis) < 2:
            st.info("Select at least 2 KPIs.")
        else:
            corr = df_f[selected_kpis].corr().round(3)
            fig_c = px.imshow(
                corr, color_continuous_scale='RdBu_r',
                zmin=-1, zmax=1, text_auto='.2f',
                title="KPI Correlation Matrix", template=RF_TPL,
            )
            fig_c.update_layout(coloraxis_colorbar_title="r")
            st.plotly_chart(fig_c, use_container_width=True)

            # Ranked pairs table
            pairs = []
            for i, k1 in enumerate(selected_kpis):
                for k2 in selected_kpis[i+1:]:
                    pairs.append({'KPI A': k1, 'KPI B': k2,
                                  'r': round(corr.loc[k1, k2], 3)})
            if pairs:
                pairs_df = pd.DataFrame(pairs).sort_values('r', key=abs,
                                                            ascending=False)
                st.dataframe(pairs_df.head(15), use_container_width=True)

            hints = root_cause_hints(corr, selected_kpis)
            if hints:
                st.markdown("### 💡 Root-Cause Hints")
                for h in hints:
                    st.warning(h)
            else:
                st.success("✅ No strong root-cause correlations at current threshold.")

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 5 — MAP
    # ══════════════════════════════════════════════════════════════════════════
    with tabs[5]:
        st.markdown("## 🗺️ Network Map")
        if not (lat_col and lon_col):
            st.info("No geo coordinates available. Upload a planning file (CSV/XLSX) "
                    "or a KML/KMZ with site placemarks, or make sure the KPI file "
                    "contains Latitude / Longitude columns.")
        else:
            m1, m2, m3 = st.columns([3, 1, 1])
            with m1:
                map_kpi = st.selectbox("Colour by KPI", selected_kpis, key='mkpi')
            with m2:
                map_dir = st.radio("Better when", ["Higher", "Lower"],
                                   horizontal=True, key='mdir')
            with m3:
                levels = [("Site", site_id_cols[0])] if site_id_cols else []
                levels = ([("Cell", cell_id_cols[0])] if cell_id_cols else []) + levels
                lvl = st.radio("Level", [n for n, _ in levels], horizontal=True,
                               key='mlvl') if len(levels) > 1 else (levels[0][0] if levels else None)
            key_col = dict(levels).get(lvl) if lvl else None

            # Sector wedges need azimuths from the planning file
            wedge_mode, beam, base_r = False, 65, 300
            if az_col:
                w1, w2, w3 = st.columns([2, 1, 1])
                with w1:
                    draw_mode = st.radio("Draw as", ["Sector wedges", "Markers"],
                                         horizontal=True, key='mdraw')
                wedge_mode = draw_mode == "Sector wedges"
                with w2:
                    beam = st.slider("Beamwidth (°)", 20, 120, 65, 5, key='mbeam',
                                     disabled=not wedge_mode)
                with w3:
                    base_r = st.slider("Wedge length (m)", 50, 2000, 300, 50, key='mrad',
                                       disabled=not wedge_mode)

            g = pd.DataFrame({
                'lat': to_coord(df_f[lat_col]),
                'lon': to_coord(df_f[lon_col]),
            }, index=df_f.index)
            # lat/lon accidentally swapped?
            if (g['lat'].abs() > 90).mean() > 0.5 and (g['lon'].abs() <= 90).mean() > 0.5:
                g = g.rename(columns={'lat': 'lon', 'lon': 'lat'})
                st.caption("ℹ️ Latitude/longitude looked swapped; corrected automatically.")
            g['kpi']  = pd.to_numeric(df_f[map_kpi], errors='coerce') if map_kpi in df_f.columns else np.nan
            g['name'] = df_f[key_col].astype(str) if key_col else ''
            g['az']   = pd.to_numeric(df_f[az_col], errors='coerce') if az_col else np.nan
            _tc, _bc  = _pcol('TEKNOLOJI'), _pcol('FREQUENCY_BAND')
            g['tech'] = df_f[_tc].astype(str) if _tc else ''
            g['band'] = df_f[_bc].astype(str) if _bc else ''
            n_raw = len(g)
            g = g[g['lat'].between(-90, 90) & g['lon'].between(-180, 180)
                  & ~((g['lat'] == 0) & (g['lon'] == 0))].dropna(subset=['lat', 'lon'])

            if g.empty:
                st.warning(f"No valid coordinates in the filtered data "
                           f"(checked {n_raw:,} rows; lat column `{lat_col}`, lon column `{lon_col}`). "
                           "If you loaded a KML/KMZ, check the join message above — the "
                           "placemark names must match cell or site IDs.")
            else:
                if not key_col:
                    g['name'] = g['lat'].round(5).astype(str) + ',' + g['lon'].round(5).astype(str)
                map_df = (g.groupby('name', observed=True)
                            .agg(lat=('lat', 'mean'), lon=('lon', 'mean'),
                                 kpi=('kpi', 'mean'), n=('kpi', 'size'),
                                 az=('az', 'first'), tech=('tech', 'first'),
                                 band=('band', 'first'))
                            .reset_index())

                has_kpi = map_df['kpi'].notna()
                scale = 'RdYlGn' if map_dir == "Higher" else 'RdYlGn_r'

                def _hover(r):
                    t = f"<b>{r['name']}</b><br>{map_kpi}: " + (
                        f"{r['kpi']:,.2f}" if pd.notna(r['kpi']) else "n/a")
                    if pd.notna(r['az']):
                        t += f"<br>Azimuth: {r['az']:.0f}°"
                    if r['tech'] and r['tech'] != 'nan':
                        t += f"<br>{r['tech']}" + (f" · band {r['band']}" if r['band'] not in ('', 'nan') else "")
                    return t
                map_df['text'] = map_df.apply(_hover, axis=1)

                # Grey markers: KML placemarks that did not match any KPI row
                grey = None
                if kml_frames:
                    k = pd.concat(kml_frames, ignore_index=True)
                    have = set(zip(map_df['lat'].round(5), map_df['lon'].round(5)))
                    grey = k[[(a_, b_) not in have for a_, b_ in
                              zip(k['Latitude'].round(5), k['Longitude'].round(5))]]

                use_new = hasattr(go, 'Scattermap')
                Sc = go.Scattermap if use_new else go.Scattermapbox
                fig_m = go.Figure()
                if grey is not None and not grey.empty:
                    fig_m.add_trace(Sc(lat=grey['Latitude'], lon=grey['Longitude'],
                                       mode='markers', name='No KPI match',
                                       marker=dict(size=7, color='#9aa5b1', opacity=0.6),
                                       text=grey['KML_Name'], hoverinfo='text'))

                # ── sector wedges ────────────────────────────────────────────
                draw_wedges = wedge_mode and lvl == "Cell" and map_df['az'].notna().any()
                if wedge_mode and lvl != "Cell":
                    st.caption("ℹ️ Sector wedges are drawn at **Cell** level — switch *Level* to Cell.")
                if draw_wedges and len(map_df) > WEDGE_MAX_CELLS:
                    st.warning(f"{len(map_df):,} cells is too many to draw as wedges "
                               f"(limit {WEDGE_MAX_CELLS:,}). Narrow the Region / City / District "
                               "filters in the sidebar — showing markers instead.")
                    draw_wedges = False
                colorbar_done = False
                if draw_wedges:
                    wd   = map_df[map_df['az'].notna()].sort_values('name').copy()
                    # co-sector layers (several carriers on one azimuth) get progressively
                    # shorter wedges so they stay visible instead of overlapping exactly
                    wd['_rk'] = wd.groupby([wd['lat'].round(5), wd['lon'].round(5), wd['az']]).cumcount()
                    rad  = (base_r * wd['tech'].map(TECH_RADIUS).fillna(1.0)
                            * (1 - 0.12 * wd['_rk']).clip(lower=0.4))
                    NB   = 10
                    try:
                        from plotly.colors import sample_colorscale
                        pal = list(sample_colorscale('RdYlGn', [i / (NB - 1) for i in range(NB)]))
                        if len(pal) != NB or not all(isinstance(c_, str) for c_ in pal):
                            raise ValueError("unexpected palette")
                    except Exception:
                        pal = ['#d73027', '#f46d43', '#fdae61', '#fee08b', '#ffffbf',
                               '#d9ef8b', '#a6d96a', '#66bd63', '#1a9850', '#006837']
                    if map_dir == "Lower":
                        pal = pal[::-1]
                    if wd['kpi'].notna().any():
                        lo_, hi_ = float(wd['kpi'].min()), float(wd['kpi'].max())
                        idx = ((wd['kpi'] - lo_) / max(hi_ - lo_, 1e-9) * (NB - 1)).round()
                    else:
                        idx = pd.Series(np.nan, index=wd.index)
                    groups = [(i, pal[i]) for i in range(NB)] + [(None, '#9aa5b1')]
                    for i, col_ in groups:
                        sub = wd[idx == i] if i is not None else wd[idx.isna()]
                        if sub.empty:
                            continue
                        pla, plo = sector_polygons(sub['lat'], sub['lon'], sub['az'],
                                                   rad[sub.index], beam)
                        fig_m.add_trace(Sc(lat=pla, lon=plo, mode='lines', fill='toself',
                                           fillcolor=_rgba(col_, 0.55),
                                           line=dict(width=0.7, color=_rgba(col_, 1.0)),
                                           hoverinfo='skip', showlegend=False))
                    # hover + colour-bar markers on the sector boresight
                    t_lat, t_lon = sector_tip(wd['lat'], wd['lon'], wd['az'], rad)
                    wd['t_lat'], wd['t_lon'] = t_lat, t_lon
                    wk = wd[wd['kpi'].notna()]
                    if not wk.empty:
                        fig_m.add_trace(Sc(
                            lat=wk['t_lat'], lon=wk['t_lon'], mode='markers', name=map_kpi,
                            marker=dict(size=6, color=wk['kpi'], colorscale=scale,
                                        cmin=float(wk['kpi'].min()), cmax=float(wk['kpi'].max()),
                                        showscale=True, colorbar=dict(title=map_kpi[:24])),
                            text=wk['text'], hoverinfo='text'))
                        colorbar_done = True
                    wn = wd[wd['kpi'].isna()]
                    if not wn.empty:
                        fig_m.add_trace(Sc(lat=wn['t_lat'], lon=wn['t_lon'], mode='markers',
                                           name='No KPI value',
                                           marker=dict(size=6, color='#9aa5b1'),
                                           text=wn['text'], hoverinfo='text'))
                    sites_xy = wd[['lat', 'lon']].round(5).drop_duplicates()
                    fig_m.add_trace(Sc(lat=sites_xy['lat'], lon=sites_xy['lon'], mode='markers',
                                       name='Site', marker=dict(size=4, color='#1f2d3a'),
                                       hoverinfo='skip'))
                    plain = map_df[map_df['az'].isna()]      # cells without azimuth → plain dots
                else:
                    plain = map_df

                pts = plain[plain['kpi'].notna()]
                if not pts.empty:
                    fig_m.add_trace(Sc(
                        lat=pts['lat'], lon=pts['lon'], mode='markers', name=map_kpi,
                        marker=dict(size=11, color=pts['kpi'], colorscale=scale,
                                    cmin=float(pts['kpi'].min()), cmax=float(pts['kpi'].max()),
                                    showscale=not colorbar_done,
                                    colorbar=dict(title=map_kpi[:24])),
                        text=pts['text'], hoverinfo='text', showlegend=not colorbar_done))
                nok = plain[plain['kpi'].isna()]
                if not nok.empty:
                    fig_m.add_trace(Sc(lat=nok['lat'], lon=nok['lon'], mode='markers',
                                       name='No KPI value' if not draw_wedges else 'No azimuth / KPI',
                                       marker=dict(size=8, color='#9aa5b1'),
                                       text=nok['text'], hoverinfo='text'))

                # Auto-fit view to the data
                lat_span = max(map_df['lat'].max() - map_df['lat'].min(), 1e-3)
                lon_span = max((map_df['lon'].max() - map_df['lon'].min())
                               * np.cos(np.radians(map_df['lat'].mean())), 1e-3)
                zoom = float(np.clip(np.log2(1300 / max(lat_span, lon_span)) - 0.5, 1, 16))
                view = dict(style='open-street-map', zoom=zoom,
                            center=dict(lat=float(map_df['lat'].mean()),
                                        lon=float(map_df['lon'].mean())))
                fig_m.update_layout(**{('map' if use_new else 'mapbox'): view},
                                    height=620, margin=dict(r=0, t=10, l=0, b=0),
                                    legend=dict(orientation='h', y=1.02, x=0))
                st.plotly_chart(fig_m, use_container_width=True)
                src = {'file': 'planning file / KML', 'auto': 'auto-extracted from KPI'}.get(plan_src, plan_src)
                extra = f" · {len(grey):,} placemarks without KPI match" if grey is not None and len(grey) else ""
                wnote = (f" · wedges: {beam}° beamwidth, {base_r} m; lower-frequency / "
                         f"co-sector layers drawn longer / shorter so they do not overlap") if draw_wedges else ""
                st.caption(f"{len(map_df):,} {lvl.lower() if lvl else 'location'}(s) shown "
                           f"(from {len(g):,} rows) · Geo source: {src}{extra}{wnote}")

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 6 — ANOMALY
    # ══════════════════════════════════════════════════════════════════════════
    with tabs[6]:
        st.markdown("## 🤖 AI Anomaly Detection")
        if len(selected_kpis) < 2:
            st.info("Select at least 2 KPIs for anomaly detection.")
        else:
            X       = df_f[selected_kpis].fillna(0)
            X_sc    = StandardScaler().fit_transform(X)
            model   = IsolationForest(contamination=contamination, random_state=42)
            labels  = model.fit_predict(X_sc)
            scores  = model.score_samples(X_sc)

            df_anom           = df_f.copy()
            df_anom['_label'] = labels
            df_anom['_score'] = scores
            anomalies         = df_anom[df_anom['_label'] == -1]

            a1, a2, a3 = st.columns(3)
            with a1: st.metric("Anomalies Found", f"{len(anomalies):,}")
            with a2: st.metric("Anomaly %",
                               f"{len(anomalies)/max(len(df_anom),1)*100:.1f}%")
            with a3:
                if cell_id_cols and cell_id_cols[0] in anomalies.columns:
                    st.metric("Affected Cells",
                              anomalies[cell_id_cols[0]].nunique())

            if not anomalies.empty and time_col:
                fig_a = go.Figure()
                normal = df_anom[df_anom['_label'] == 1]
                fig_a.add_trace(go.Scatter(
                    x=normal[time_col], y=normal['_score'],
                    mode='markers', name='Normal',
                    marker=dict(color='#27a99a', size=4, opacity=0.5)))
                fig_a.add_trace(go.Scatter(
                    x=anomalies[time_col], y=anomalies['_score'],
                    mode='markers', name='Anomaly',
                    marker=dict(color='#e05a4e', size=6, symbol='x')))
                fig_a.update_layout(title="Anomaly Scores Over Time",
                                    template=RF_TPL, hovermode='x unified')
                st.plotly_chart(fig_a, use_container_width=True)

            if not anomalies.empty:
                st.markdown("### 🔍 Top Anomalies")
                disp = []
                if time_col and time_col in anomalies.columns:
                    disp.append(time_col)
                disp.append('_score')
                disp += [c for c in cell_id_cols[:1] + site_id_cols[:1]
                          if c in anomalies.columns]
                disp += [c for c in selected_kpis if c in anomalies.columns]

                anom_show = anomalies[disp].sort_values('_score').head(30).copy()
                anom_show = anom_show.rename(columns={'_score': 'Anomaly Score'})
                if time_col and time_col in anom_show.columns:
                    anom_show[time_col] = anom_show[time_col].dt.strftime('%Y-%m-%d')
                st.dataframe(anom_show, use_container_width=True)

                # Root-cause hints for anomalous records
                a_means = anomalies[selected_kpis].mean()
                o_means = df_f[selected_kpis].mean()
                flagged = [k for k in selected_kpis
                           if o_means[k] != 0 and
                           abs(a_means[k] - o_means[k]) / abs(o_means[k]) > 0.10]
                hints = root_cause_hints(df_f[selected_kpis].corr(), flagged)
                if hints:
                    st.markdown("#### 💡 Possible Root Causes")
                    for h in hints:
                        st.warning(h)

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 7 — PERIOD COMPARE
    # ══════════════════════════════════════════════════════════════════════════
    with tabs[7]:
        st.markdown("## ↔️ Period-over-Period Comparison")
        if not time_col:
            st.info("Requires a time column.")
        else:
            g_min = df[time_col].min().date()
            g_max = df[time_col].max().date()
            mid   = g_min + (g_max - g_min) // 2

            pc1, pc2 = st.columns(2)
            with pc1:
                st.markdown("**Period A**")
                per_a = st.date_input("Period A", value=(g_min, mid),
                                      min_value=g_min, max_value=g_max, key='pera')
            with pc2:
                st.markdown("**Period B**")
                per_b = st.date_input("Period B", value=(mid, g_max),
                                      min_value=g_min, max_value=g_max, key='perb')

            if (isinstance(per_a, (list, tuple)) and len(per_a) == 2 and
                    isinstance(per_b, (list, tuple)) and len(per_b) == 2):
                dfa = df[(df[time_col].dt.date >= per_a[0]) &
                         (df[time_col].dt.date <= per_a[1])]
                dfb = df[(df[time_col].dt.date >= per_b[0]) &
                         (df[time_col].dt.date <= per_b[1])]

                if dfa.empty or dfb.empty:
                    st.warning("One or both periods contain no data.")
                else:
                    ma = dfa[selected_kpis].mean()
                    mb = dfb[selected_kpis].mean()
                    delta = ((mb - ma) / ma.replace(0, np.nan) * 100).round(2)

                    cmp = pd.DataFrame({
                        f"A  {per_a[0]}→{per_a[1]}": ma.round(4),
                        f"B  {per_b[0]}→{per_b[1]}": mb.round(4),
                        "Δ %": delta,
                    })

                    def colour_delta(v):
                        try:
                            f = float(v)
                            return 'color:#27a99a' if f > 0 else 'color:#e05a4e'
                        except Exception:
                            return ''

                    st.dataframe(
                        cmp.style.map(colour_delta, subset=["Δ %"]),
                        use_container_width=True,
                    )

                    fig_d = px.bar(
                        x=cmp.index, y=cmp["Δ %"],
                        color=cmp["Δ %"], color_continuous_scale='RdYlGn',
                        title="KPI Change  A → B (%)",
                        template=RF_TPL,
                        labels={"x": "KPI", "y": "Δ %"},
                    )
                    fig_d.update_coloraxes(showscale=False)
                    st.plotly_chart(fig_d, use_container_width=True)

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 8 — CHANGE POINTS
    # Uses go.Scatter vertical lines — avoids add_vline Timestamp bug entirely
    # ══════════════════════════════════════════════════════════════════════════
    with tabs[8]:
        st.markdown("## 🔍 Change-Point Detection")
        if not time_col:
            st.info("Requires a time column.")
        else:
            cp_kpi  = st.selectbox("KPI", selected_kpis, key='cpkpi')
            n_bkps  = st.slider("Max change points", 2, 10, 5, key='nbkps')

            ts = (df_f[[time_col, cp_kpi]]
                  .groupby(pd.Grouper(key=time_col, freq='D'))[cp_kpi]
                  .mean().dropna().reset_index())

            bkps = change_points(ts[cp_kpi], n_bkps)

            y_min = float(ts[cp_kpi].min()) if len(ts) else 0.0
            y_max = float(ts[cp_kpi].max()) if len(ts) else 1.0

            fig_cp = go.Figure()
            fig_cp.add_trace(go.Scatter(
                x=ts[time_col], y=ts[cp_kpi],
                name=cp_kpi, line=dict(color=TEAL, width=1.8)))

            CP_COLORS = ['#e74c3c','#f39c12','#9b59b6','#1abc9c','#e67e22',
                         '#e91e63','#00bcd4','#ff5722','#8bc34a','#ff9800']

            for i, bp in enumerate(bkps):
                if bp < len(ts):
                    fig_cp.add_trace(vline_trace(
                        x_val=ts[time_col].iloc[bp],
                        y_min=y_min, y_max=y_max,
                        color=CP_COLORS[i % len(CP_COLORS)],
                        label=f"CP{i+1}",
                    ))

            fig_cp.update_layout(
                title=f"{cp_kpi} — {len(bkps)} Change Point(s) Detected",
                template=RF_TPL, hovermode='x unified',
                yaxis=dict(range=[y_min * 0.97, y_max * 1.06]),
            )
            st.plotly_chart(fig_cp, use_container_width=True)

            if bkps:
                rows, prev = [], 0
                all_bkps   = bkps + [len(ts)]
                for i, bp in enumerate(all_bkps):
                    seg = ts[cp_kpi].iloc[prev:bp]
                    if len(seg) > 0:
                        seg_mean = round(seg.mean(), 4)
                        prev_mean = rows[-1]['Mean'] if rows else None
                        delta = (f"{((seg_mean-prev_mean)/abs(prev_mean)*100):+.1f}%"
                                 if prev_mean and prev_mean != 0 else "—")
                        rows.append({
                            'Seg':  f"S{i+1}",
                            'From': ts[time_col].iloc[prev].date() if prev < len(ts) else '—',
                            'To':   ts[time_col].iloc[min(bp, len(ts)-1)].date(),
                            'Mean': seg_mean,
                            'Δ vs prev': delta,
                            'Trend': '📈' if seg.iloc[-1] > seg.iloc[0] else '📉',
                        })
                    prev = bp
                st.dataframe(pd.DataFrame(rows), use_container_width=True)
                if not HAS_RUPTURES:
                    st.caption("💡 Install `ruptures` for PELT algorithm: "
                               "`pip install ruptures`")
            else:
                st.success("✅ No significant change points detected.")

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 9 — AI QUERY
    # ══════════════════════════════════════════════════════════════════════════
    with tabs[9]:
        st.markdown("## 💬 AI-Powered Natural Language Query")
        st.caption("Llama 3.3 (via Groq) answers using actual computed per-cell KPI averages from your data.")

        # API key management
        has_secret = False
        try:
            has_secret = bool(st.secrets.get("GROQ_API_KEY", ""))
        except Exception:
            pass

        if not has_secret:
            with st.expander("🔑 API Key Configuration",
                             expanded=not st.session_state.anthropic_api_key):
                k_in = st.text_input(
                    "Groq API Key (free at groq.com)",
                    value=st.session_state.anthropic_api_key,
                    type="password",
                    placeholder="gsk_...",
                    key='api_key_in',
                    help="Get free key from console.groq.com",
                )
                col_btn1, col_btn2 = st.columns([1, 4])
                with col_btn1:
                    if st.button("💾 Save", key='save_key'):
                        st.session_state.anthropic_api_key = k_in
                        st.success("✅ Key saved for this session.")
                        st.rerun()
                if not st.session_state.anthropic_api_key:
                    st.info(
                        "Alternatively, create `.streamlit/secrets.toml`:\n"
                        "```toml\nGROQ_API_KEY = \"gsk_...\"\n```"
                    )

        # Resolve API key
        api_key = ""
        try:
            api_key = st.secrets.get("GROQ_API_KEY", "")
        except Exception:
            pass
        if not api_key:
            api_key = st.session_state.anthropic_api_key

        # Example prompts
        with st.expander("💡 Example questions"):
            examples = [
                "Which 5 cells have the worst average drop rate?",
                "What is the average HSDPA throughput across all cells?",
                "Which cells have RRC SR below 95%?",
                "Summarise overall network health.",
                "Which KPI has the highest variance across cells?",
            ]
            ex_cols = st.columns(len(examples))
            for i, (ex, ec) in enumerate(zip(examples, ex_cols)):
                with ec:
                    if st.button(ex[:26] + "…" if len(ex) > 26 else ex,
                                 key=f'ex_{i}', use_container_width=True):
                        st.session_state.nl_history.append(
                            {"role": "user", "content": ex})
                        st.rerun()

        # Chat history
        for msg in st.session_state.nl_history:
            with st.chat_message(msg["role"]):
                st.markdown(msg["content"])

        # Chat input
        user_q = st.chat_input("Ask about your RAN data…")
        if user_q:
            st.session_state.nl_history.append({"role": "user", "content": user_q})
            with st.chat_message("user"):
                st.markdown(user_q)
            with st.chat_message("assistant"):
                with st.spinner("Analysing…"):
                    ccol = cell_id_cols[0] if cell_id_cols else ""
                    ans  = ai_query(user_q, df_f, selected_kpis,
                                    ccol, time_col or "", api_key)
                st.markdown(ans)
                st.session_state.nl_history.append(
                    {"role": "assistant", "content": ans})

        if st.session_state.nl_history:
            if st.button("🗑️ Clear chat", key='clr_chat'):
                st.session_state.nl_history = []
                st.rerun()

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 10 — EXPORT
    # ══════════════════════════════════════════════════════════════════════════
    with tabs[16]:
        st.markdown("## 📥 Export")

        # Recompute for export
        cell_col_e  = cell_id_cols[0] if cell_id_cols else None
        health_exp  = health_score(df_f, selected_kpis, cell_col_e, user_thresholds)
        worst_exp   = (health_exp.nsmallest(20, 'Score')
                       if not health_exp.empty else pd.DataFrame())

        summary_exp = df_f[selected_kpis].agg(
            ['mean','median','max','min','std']).T.round(4)
        summary_exp.columns = ['Mean','Median','Max','Min','Std Dev']

        # Count anomalies if they were computed in tab 6
        anom_count = 0  # tab 6 runs independently

        e1, e2, e3, e4 = st.columns(4)
        with e1:
            st.download_button(
                "📥 Filtered Data (Excel)",
                to_excel(df_f, "Filtered"),
                f"ran_filtered_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with e2:
            st.download_button(
                "📥 KPI Summary (Excel)",
                to_excel(summary_exp.reset_index(), "KPI Summary"),
                f"ran_summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with e3:
            if not health_exp.empty:
                st.download_button(
                    "📥 Health Scores (Excel)",
                    to_excel(health_exp.reset_index(), "Health"),
                    f"ran_health_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
        with e4:
            if HAS_REPORTLAB:
                try:
                    pdf = build_pdf(summary_exp, health_exp, worst_exp,
                                    anom_count, date_str, filter_str)
                    st.download_button(
                        "📄 Full PDF Report",
                        pdf,
                        f"ran_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                    )
                except Exception as pdf_err:
                    st.warning(f"PDF generation failed: {pdf_err}")
            else:
                st.info("PDF needs reportlab:\n`pip install reportlab`")

        st.markdown("---")
        st.markdown("### 📋 Export Preview")
        st.dataframe(summary_exp, use_container_width=True)

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 10 — SLA SCORECARD
    # ══════════════════════════════════════════════════════════════════════════
    with tabs[10]:
        st.markdown("## 🏆 SLA Scorecard")
        st.caption("Edit the target and direction per KPI. Compliance = share of samples "
                   "(and of cells) that meet the target.")
        cfg_rows = []
        for k in selected_kpis:
            kl  = k.lower()
            thr = next((d for pat, (d, _c) in user_thresholds.items() if pat in kl), None)
            med = df_f[k].median()
            tgt = float(thr) if thr is not None else (float(med) if pd.notna(med) else 0.0)
            cfg_rows.append({'KPI': k,
                             'Direction': 'Lower is better' if lower_is_better(k) else 'Higher is better',
                             'Target': round(tgt, 3)})
        cfg = st.data_editor(
            pd.DataFrame(cfg_rows), hide_index=True, use_container_width=True,
            disabled=['KPI'], key='sla_cfg',
            column_config={
                'Direction': st.column_config.SelectboxColumn(
                    options=['Higher is better', 'Lower is better']),
                'Target': st.column_config.NumberColumn(format="%.3f"),
            })
        res = []
        for _, r in cfg.iterrows():
            k, lb, t = r['KPI'], str(r['Direction']).startswith('Lower'), float(r['Target'])
            s = df_f[k].dropna()
            if s.empty:
                continue
            ok_s = (s <= t) if lb else (s >= t)
            row = {'KPI': k, 'Target': t, 'Network avg': round(float(s.mean()), 3),
                   'Samples OK %': round(float(ok_s.mean() * 100), 1)}
            if entity_col:
                ca  = df_f.groupby(entity_col, observed=True)[k].mean().dropna()
                okc = (ca <= t) if lb else (ca >= t)
                row['Cells OK %']    = round(float(okc.mean() * 100), 1)
                row['Cells failing'] = int((~okc).sum())
            sv = row['Samples OK %']
            row['Status'] = '🟢 Met' if sv >= 95 else ('🟡 At risk' if sv >= 85 else '🔴 Missed')
            res.append(row)
        if res:
            sla = pd.DataFrame(res)
            m1, m2, m3, m4 = st.columns(4)
            with m1:
                st.markdown(metric_card("Overall compliance",
                            f"{sla['Samples OK %'].mean():.1f}%"), unsafe_allow_html=True)
            with m2:
                st.markdown(metric_card("KPIs met", f"{(sla['Status'] == '🟢 Met').sum()}",
                            f"of {len(sla)}"), unsafe_allow_html=True)
            with m3:
                st.markdown(metric_card("At risk", f"{(sla['Status'] == '🟡 At risk').sum()}"),
                            unsafe_allow_html=True)
            with m4:
                st.markdown(metric_card("Missed", f"{(sla['Status'] == '🔴 Missed').sum()}"),
                            unsafe_allow_html=True)
            st.markdown("")
            colcfg = {'Samples OK %': st.column_config.ProgressColumn(
                          'Samples OK %', min_value=0, max_value=100, format="%.1f%%")}
            if 'Cells OK %' in sla:
                colcfg['Cells OK %'] = st.column_config.ProgressColumn(
                    'Cells OK %', min_value=0, max_value=100, format="%.1f%%")
            st.dataframe(sla, hide_index=True, use_container_width=True, column_config=colcfg)
            figc = px.bar(sla, x='KPI', y='Samples OK %', color='Samples OK %',
                          color_continuous_scale=['#e05a4e', '#f2b84b', '#2fa36b'],
                          range_color=[60, 100], template=RF_TPL,
                          title="SLA compliance by KPI")
            figc.add_hline(y=95, line_dash='dash', line_color='#718096',
                           annotation_text='95 % goal')
            figc.update_layout(coloraxis_showscale=False, yaxis_range=[0, 105])
            st.plotly_chart(figc, use_container_width=True)
            if time_col:
                tk = st.selectbox("Compliance trend for", sla['KPI'].tolist(), key='sla_tk')
                rr = cfg[cfg['KPI'] == tk].iloc[0]
                lb, t = str(rr['Direction']).startswith('Lower'), float(rr['Target'])
                tr = (df_f.assign(_d=df_f[time_col].dt.date).groupby('_d')[tk]
                          .apply(lambda s: ((s.dropna() <= t) if lb else (s.dropna() >= t)).mean() * 100)
                          .reset_index(name='Compliance %'))
                figt = px.area(tr, x='_d', y='Compliance %', template=RF_TPL,
                               title=f"{tk} — daily compliance")
                figt.update_layout(xaxis_title='', yaxis_range=[0, 100])
                st.plotly_chart(figt, use_container_width=True)
            st.download_button("⬇️ Download scorecard (CSV)", sla.to_csv(index=False).encode(),
                               "sla_scorecard.csv", "text/csv")

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 11 — CELL DRILL-DOWN
    # ══════════════════════════════════════════════════════════════════════════
    with tabs[11]:
        st.markdown("## 🎯 Cell / Site Drill-down")
        if not entity_col:
            st.info("No Cell or Site ID column detected.")
        else:
            ents = sorted(df_f[entity_col].dropna().astype(str).unique())
            d1, d2 = st.columns([1, 2])
            with d1:
                sel_e = st.selectbox(f"{entity_col}", ents, key='dd_e')
            with d2:
                dd_k = st.multiselect("KPIs", selected_kpis,
                                      default=selected_kpis[:min(4, len(selected_kpis))],
                                      key='dd_k')
            if dd_k:
                sub     = df_f[df_f[entity_col].astype(str) == sel_e]
                ent_avg = df_f.groupby(entity_col, observed=True)[dd_k].mean()
                ent_avg.index = ent_avg.index.astype(str)
                cols = st.columns(min(4, len(dd_k)))
                rows = []
                for i, k in enumerate(dd_k):
                    me  = sub[k].mean()
                    net = df_f[k].mean()
                    dlt = (me - net) / abs(net) * 100 if net else np.nan
                    lb  = lower_is_better(k)
                    rk  = ent_avg[k].rank(pct=True, ascending=not lb).get(sel_e, np.nan) * 100
                    rows.append({'KPI': k, 'This element': round(me, 3),
                                 'Network avg': round(net, 3), 'Δ vs network %': round(dlt, 1),
                                 'Better than % of cells': round(rk, 0)})
                    with cols[i % len(cols)]:
                        st.markdown(metric_card(k[:28], f"{me:,.2f}",
                                    f"Network {net:,.2f} · Δ {dlt:+.1f}%"), unsafe_allow_html=True)
                st.markdown("")
                st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)
                if time_col:
                    sp   = df_plot[df_plot[entity_col].astype(str) == sel_e].sort_values(time_col)
                    netp = df_plot.groupby(time_col)[dd_k].mean()
                    fig = make_subplots(rows=len(dd_k), cols=1, shared_xaxes=True,
                                        subplot_titles=dd_k, vertical_spacing=0.07)
                    for i, k in enumerate(dd_k, 1):
                        fig.add_trace(go.Scatter(x=sp[time_col], y=sp[k], name=sel_e,
                                      line=dict(color=TEAL, width=2), showlegend=(i == 1)), row=i, col=1)
                        fig.add_trace(go.Scatter(x=netp.index, y=netp[k], name='Network avg',
                                      line=dict(color=AMBER, width=1.5, dash='dot'),
                                      showlegend=(i == 1)), row=i, col=1)
                    fig.update_layout(height=max(320, 230 * len(dd_k)), template=RF_TPL,
                                      hovermode='x unified',
                                      legend=dict(orientation='h', y=1.04, x=0))
                    st.plotly_chart(fig, use_container_width=True)

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 12 — DISTRIBUTION & PATTERNS
    # ══════════════════════════════════════════════════════════════════════════
    with tabs[12]:
        st.markdown("## 📊 Distribution & Patterns")
        dk = st.selectbox("KPI", selected_kpis, key='ds_k')
        s  = df_f[dk].dropna()
        if s.empty:
            st.info("No data for this KPI.")
        else:
            pct = s.quantile([.01, .05, .25, .5, .75, .95, .99])
            pc  = st.columns(7)
            for c, (q, v) in zip(pc, pct.items()):
                with c:
                    st.markdown(metric_card(f"P{int(q * 100)}", f"{v:,.2f}"), unsafe_allow_html=True)
            st.markdown("")
            g1, g2 = st.columns(2)
            with g1:
                fh = px.histogram(s, nbins=50, marginal='box', template=RF_TPL,
                                  color_discrete_sequence=[TEAL], title=f"{dk} — distribution")
                fh.update_layout(showlegend=False, xaxis_title=dk)
                st.plotly_chart(fh, use_container_width=True)
            with g2:
                xs = np.sort(s.values)
                step = max(1, len(xs) // 2000)
                fc = go.Figure(go.Scatter(x=xs[::step], y=np.linspace(0, 100, len(xs))[::step],
                                          line=dict(color=TEAL, width=2)))
                fc.update_layout(template=RF_TPL, title=f"{dk} — cumulative distribution",
                                 xaxis_title=dk, yaxis_title='% of samples ≤ x')
                st.plotly_chart(fc, use_container_width=True)
            if time_col:
                t = df_f[[time_col, dk]].dropna()
                t['_dow'] = t[time_col].dt.day_name()
                order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
                if t[time_col].dt.hour.nunique() > 1:
                    t['_x'] = t[time_col].dt.hour
                    xlab = 'Hour of day'
                else:
                    t['_x'] = t[time_col].dt.isocalendar().week.astype(int)
                    xlab = 'Week of year'
                pv = t.pivot_table(index='_dow', columns='_x', values=dk, aggfunc='mean')
                pv = pv.reindex([d for d in order if d in pv.index])
                if pv.shape[0] > 0 and pv.shape[1] > 0:
                    cs = 'RdYlGn_r' if lower_is_better(dk) else 'RdYlGn'
                    fz = px.imshow(pv, aspect='auto', color_continuous_scale=cs, template=RF_TPL,
                                   labels=dict(x=xlab, y='', color=dk),
                                   title=f"{dk} — weekday × {xlab.lower()} pattern")
                    st.plotly_chart(fz, use_container_width=True)

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 13 — CHRONIC CELLS & MOVERS
    # ══════════════════════════════════════════════════════════════════════════
    with tabs[13]:
        st.markdown("## ⛔ Chronic Cells & Movers")
        if not entity_col or not time_col:
            st.info("This analysis needs a Cell/Site ID column and a date column.")
        else:
            sub1, sub2 = st.tabs(["⛔ Chronic offenders", "🔀 Movers (period vs period)"])
            dbase = df_f.assign(_d=df_f[time_col].dt.date)

            with sub1:
                c1, c2, c3 = st.columns(3)
                with c1:
                    ck = st.selectbox("KPI", selected_kpis, key='ch_k')
                with c2:
                    pp = st.slider("Worst % of cells each day", 5, 30, 10, key='ch_p')
                with c3:
                    minf = st.slider("Chronic if flagged on ≥ % of days", 20, 100, 50, key='ch_f')
                lb = lower_is_better(ck)
                dd = (dbase.groupby(['_d', entity_col], observed=True)[ck]
                           .mean().reset_index().dropna())
                if dd['_d'].nunique() < 3:
                    st.info("Need at least 3 days of data for a meaningful chronic analysis.")
                else:
                    q   = pp / 100
                    cut = dd.groupby('_d')[ck].transform(
                        lambda x: x.quantile(1 - q if lb else q))
                    dd['flag'] = (dd[ck] >= cut) if lb else (dd[ck] <= cut)
                    agg = (dd.groupby(entity_col, observed=True)
                             .agg(Days=('flag', 'size'), Flagged=('flag', 'sum'), Avg=(ck, 'mean'))
                             .reset_index())
                    agg['Flagged %'] = (agg['Flagged'] / agg['Days'] * 100).round(1)
                    agg['Avg'] = agg['Avg'].round(3)
                    chronic = (agg[(agg['Flagged %'] >= minf) & (agg['Days'] >= 3)]
                               .sort_values(['Flagged %', 'Flagged'], ascending=False))
                    k1, k2, k3 = st.columns(3)
                    with k1:
                        st.markdown(metric_card("Chronic cells", f"{len(chronic):,}"),
                                    unsafe_allow_html=True)
                    with k2:
                        st.markdown(metric_card("Share of network",
                                    f"{len(chronic) / max(len(agg), 1) * 100:.1f}%"),
                                    unsafe_allow_html=True)
                    with k3:
                        st.markdown(metric_card("Days analysed", f"{dd['_d'].nunique()}"),
                                    unsafe_allow_html=True)
                    st.markdown("")
                    st.dataframe(chronic.head(200), hide_index=True, use_container_width=True,
                                 column_config={'Flagged %': st.column_config.ProgressColumn(
                                     'Flagged %', min_value=0, max_value=100, format="%.1f%%")})
                    if not chronic.empty:
                        fb = px.bar(chronic.head(20), x=entity_col, y='Flagged %',
                                    color='Flagged %', color_continuous_scale='Reds',
                                    template=RF_TPL, title=f"Top chronic cells — {ck}")
                        fb.update_layout(xaxis_tickangle=-45, coloraxis_showscale=False)
                        st.plotly_chart(fb, use_container_width=True)
                    par = agg.sort_values('Flagged', ascending=False).reset_index(drop=True)
                    if par['Flagged'].sum() > 0:
                        par['Cum %'] = par['Flagged'].cumsum() / par['Flagged'].sum() * 100
                        par['Cells %'] = (par.index + 1) / len(par) * 100
                        fp = go.Figure(go.Scatter(x=par['Cells %'], y=par['Cum %'],
                                                  line=dict(color=TEAL, width=2.5), fill='tozeroy'))
                        fp.add_hline(y=80, line_dash='dash', line_color='#718096')
                        fp.update_layout(template=RF_TPL, title="Pareto — share of bad-cell events",
                                         xaxis_title='% of cells (worst first)',
                                         yaxis_title='% of all flagged events')
                        st.plotly_chart(fp, use_container_width=True)
                    st.download_button("⬇️ Download chronic list (CSV)",
                                       chronic.to_csv(index=False).encode(),
                                       "chronic_cells.csv", "text/csv")

            with sub2:
                days = sorted(dbase['_d'].unique())
                if len(days) < 2:
                    st.info("Need at least 2 days of data to compare periods.")
                else:
                    maxw = max(1, len(days) // 2)
                    m1, m2 = st.columns(2)
                    with m1:
                        mk = st.selectbox("KPI", selected_kpis, key='mv_k')
                    with m2:
                        w = st.slider("Window (days)", 1, min(14, maxw), min(7, maxw), key='mv_w')
                    recent, prev = days[-w:], days[-2 * w:-w]
                    a = dbase[dbase['_d'].isin(recent)].groupby(entity_col, observed=True)[mk].mean()
                    b = dbase[dbase['_d'].isin(prev)].groupby(entity_col, observed=True)[mk].mean()
                    mv = pd.concat([b.rename('Previous'), a.rename('Recent')], axis=1).dropna()
                    if mv.empty:
                        st.info("No cells have data in both periods.")
                    else:
                        mv['Δ'] = mv['Recent'] - mv['Previous']
                        mv['Δ %'] = (mv['Δ'] / mv['Previous'].abs().replace(0, np.nan) * 100).round(1)
                        mv['Impact'] = -mv['Δ'] if lower_is_better(mk) else mv['Δ']
                        mv = mv.round(3).reset_index()
                        st.caption(f"Recent: {recent[0]} → {recent[-1]}  ·  "
                                   f"Previous: {prev[0]} → {prev[-1]}")
                        worse = mv.nsmallest(15, 'Impact')
                        better = mv.nlargest(15, 'Impact')
                        b1, b2 = st.columns(2)
                        with b1:
                            fw = px.bar(worse, x='Impact', y=entity_col, orientation='h',
                                        template=RF_TPL, color_discrete_sequence=['#e05a4e'],
                                        title="Top decliners")
                            fw.update_layout(yaxis=dict(autorange='reversed', type='category'))
                            st.plotly_chart(fw, use_container_width=True)
                        with b2:
                            fi = px.bar(better, x='Impact', y=entity_col, orientation='h',
                                        template=RF_TPL, color_discrete_sequence=['#2fa36b'],
                                        title="Top improvers")
                            fi.update_layout(yaxis=dict(autorange='reversed', type='category'))
                            st.plotly_chart(fi, use_container_width=True)
                        st.dataframe(mv.sort_values('Impact'), hide_index=True,
                                     use_container_width=True)

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 14 — CLUSTERING
    # ══════════════════════════════════════════════════════════════════════════
    with tabs[14]:
        st.markdown("## 🧬 Cell Clustering")
        st.caption("Group cells with similar KPI behaviour (K-Means on standardised "
                   "cell averages, shown on a 2-D PCA map).")
        if not entity_col or len(selected_kpis) < 2:
            st.info("Needs a Cell/Site ID column and at least 2 selected KPIs.")
        else:
            cl_k = st.multiselect("KPIs used", selected_kpis, default=selected_kpis[:min(6, len(selected_kpis))],
                                  key='cl_kpis')
            kk_ = st.slider("Number of clusters", 2, 8, 4, key='cl_n')
            feat = df_f.groupby(entity_col, observed=True)[cl_k].mean() if len(cl_k) >= 2 else pd.DataFrame()
            if feat.empty or len(feat) < kk_ * 3:
                st.info("Not enough cells / KPIs for this many clusters.")
            else:
                from sklearn.cluster import KMeans
                from sklearn.decomposition import PCA
                X = feat.fillna(feat.median())
                X = X.loc[:, X.std() > 0]
                if X.shape[1] < 2:
                    st.info("Selected KPIs are (almost) constant — pick others.")
                else:
                    Z = StandardScaler().fit_transform(X)
                    lab = KMeans(n_clusters=kk_, n_init=10, random_state=42).fit_predict(Z)
                    pc2 = PCA(n_components=2, random_state=42).fit_transform(Z)
                    out = feat.copy()
                    out['Cluster'] = ['C' + str(i + 1) for i in lab]
                    pdf = pd.DataFrame({'PC1': pc2[:, 0], 'PC2': pc2[:, 1],
                                        'Cluster': out['Cluster'].values,
                                        'Cell': feat.index.astype(str)})
                    if len(pdf) > 5000:
                        pdf = pdf.sample(5000, random_state=1)
                    fs = px.scatter(pdf, x='PC1', y='PC2', color='Cluster', hover_name='Cell',
                                    template=RF_TPL, opacity=0.8, title="Cluster map (PCA)")
                    st.plotly_chart(fs, use_container_width=True)
                    prof = out.groupby('Cluster').mean().round(3)
                    prof.insert(0, 'Cells', out.groupby('Cluster').size())
                    st.markdown("### Cluster profiles")
                    st.dataframe(prof.style.background_gradient(cmap='RdYlGn', axis=0, subset=cl_k),
                                 use_container_width=True)
                    st.download_button("⬇️ Download assignments (CSV)",
                                       out.reset_index().to_csv(index=False).encode(),
                                       "cell_clusters.csv", "text/csv")

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 15 — DATA QUALITY
    # ══════════════════════════════════════════════════════════════════════════
    with tabs[15]:
        st.markdown("## 🧪 Data Quality")
        qrows = []
        for k in kpi_cols[:80]:
            s0 = df_f[k]
            sn = s0.dropna()
            miss = float(s0.isna().mean() * 100)
            if len(sn):
                q1, q3 = sn.quantile([.25, .75])
                iqr = q3 - q1
                out_p = float(((sn < q1 - 3 * iqr) | (sn > q3 + 3 * iqr)).mean() * 100) if iqr > 0 else 0.0
                zero_p = float((sn == 0).mean() * 100)
                const = sn.nunique() <= 1
            else:
                out_p = zero_p = 0.0
                const = True
            qrows.append({'KPI': k, 'Completeness %': round(100 - miss, 1),
                          'Zeros %': round(zero_p, 1), 'Outliers %': round(out_p, 2),
                          'Constant': '⚠️' if const else ''})
        dq = pd.DataFrame(qrows)
        dups = int(df_f.duplicated([time_col, entity_col]).sum()) if (time_col and entity_col) else 0
        q1c, q2c, q3c, q4c = st.columns(4)
        with q1c:
            st.markdown(metric_card("Rows", f"{len(df_f):,}"), unsafe_allow_html=True)
        with q2c:
            st.markdown(metric_card("Avg completeness",
                        f"{dq['Completeness %'].mean():.1f}%" if not dq.empty else "—"),
                        unsafe_allow_html=True)
        with q3c:
            st.markdown(metric_card("KPIs > 20% missing",
                        f"{int((dq['Completeness %'] < 80).sum()) if not dq.empty else 0}"),
                        unsafe_allow_html=True)
        with q4c:
            st.markdown(metric_card("Duplicate time+cell rows", f"{dups:,}"), unsafe_allow_html=True)
        st.markdown("")
        if not dq.empty:
            st.dataframe(dq, hide_index=True, use_container_width=True,
                         column_config={'Completeness %': st.column_config.ProgressColumn(
                             'Completeness %', min_value=0, max_value=100, format="%.1f%%")})
        if time_col:
            per_day = df_f.groupby(df_f[time_col].dt.date).size().reset_index(name='Rows')
            per_day.columns = ['Date', 'Rows']
            fd = px.bar(per_day, x='Date', y='Rows', template=RF_TPL,
                        color_discrete_sequence=[TEAL], title="Rows per day (completeness over time)")
            st.plotly_chart(fd, use_container_width=True)
            ts = pd.Series(sorted(df_f[time_col].dropna().unique()))
            if len(ts) > 2:
                diffs = ts.diff().dropna()
                exp = diffs.mode().iloc[0]
                gaps = ts[1:][diffs.values > exp * 1.5]
                if len(gaps):
                    st.warning(f"⏱️ {len(gaps)} time gap(s) detected (expected step {exp}). "
                               "First few after: " + ", ".join(str(g)[:16] for g in gaps[:5]))
                else:
                    st.success("⏱️ No gaps in the time series.")
        if entity_col and time_col:
            nd = df_f.groupby(entity_col, observed=True)[time_col].nunique()
            low = int((nd < 0.8 * nd.max()).sum())
            st.info(f"📡 {low:,} of {len(nd):,} elements have data for less than 80% of the days.")
        if not dq.empty:
            st.download_button("⬇️ Download quality report (CSV)", dq.to_csv(index=False).encode(),
                               "data_quality.csv", "text/csv")


    # ══════════════════════════════════════════════════════════════════════════
    # TAB 17 — PLANNING DATA
    # ══════════════════════════════════════════════════════════════════════════
    with tabs[17]:
        st.markdown("## 📡 Planning Data")
        if plan_src != "file" or df_plan is None or 'CELL_NAME' not in df_plan.columns:
            st.info("Load a Nokia 5G cell-parameter planning file (CSV with CELL_NAME, AZIMUTH, "
                    "LATITUDE_DEC, LONGITUDE_DEC …) to see the planning inventory and checks here.")
        else:
            P = df_plan.copy()
            site_c = 'GNODEB_NAME' if 'GNODEB_NAME' in P.columns else ('EMISYON_NAME' if 'EMISYON_NAME' in P.columns else None)
            n_cells = int(P['CELL_NAME'].notna().sum())
            has_xy  = P['Plan_Latitude'].notna() & P['Plan_Longitude'].notna() if 'Plan_Latitude' in P.columns else pd.Series(False, index=P.index)
            has_az  = P['AZIMUTH'].notna() if 'AZIMUTH' in P.columns else pd.Series(False, index=P.index)

            # KPI ↔ planning coverage
            _pc_cols = [c for c in df.columns if c not in df_kpi.columns]
            kcell = cell_id_cols[0] if cell_id_cols and cell_id_cols[0] in df.columns else None
            cov_txt, unmatched = "—", pd.DataFrame()
            if kcell and _pc_cols:
                m_ = df[_pc_cols].notna().any(axis=1)
                u_all = df[kcell].dropna().nunique()
                u_hit = df.loc[m_, kcell].dropna().nunique()
                cov_txt = f"{u_hit / max(u_all, 1) * 100:.1f}%"
                unmatched = (pd.DataFrame({'KPI cell': sorted(set(df[kcell].dropna().astype(str))
                                                              - set(df.loc[m_, kcell].dropna().astype(str)))}))

            c1, c2, c3, c4, c5, c6 = st.columns(6)
            with c1: st.markdown(metric_card("Cells", f"{n_cells:,}"), unsafe_allow_html=True)
            with c2: st.markdown(metric_card("Sites", f"{P[site_c].nunique():,}" if site_c else "—"), unsafe_allow_html=True)
            with c3: st.markdown(metric_card("Regions", f"{P['REGIONCODE'].nunique():,}" if 'REGIONCODE' in P.columns else "—"), unsafe_allow_html=True)
            with c4: st.markdown(metric_card("With coordinates", f"{has_xy.mean() * 100:.1f}%"), unsafe_allow_html=True)
            with c5: st.markdown(metric_card("With azimuth", f"{has_az.mean() * 100:.1f}%"), unsafe_allow_html=True)
            with c6: st.markdown(metric_card("KPI cells in planning", cov_txt), unsafe_allow_html=True)
            st.markdown("")

            # Inventory charts
            ch1, ch2 = st.columns(2)
            with ch1:
                if {'REGIONCODE', 'TEKNOLOJI'} <= set(P.columns):
                    inv = (P.dropna(subset=['REGIONCODE']).assign(TEKNOLOJI=lambda d: d['TEKNOLOJI'].fillna('n/a'))
                             .groupby(['REGIONCODE', 'TEKNOLOJI']).size().reset_index(name='Cells'))
                    st.plotly_chart(px.bar(inv, x='REGIONCODE', y='Cells', color='TEKNOLOJI',
                                           template=RF_TPL, title="Cells by region & technology"),
                                    use_container_width=True)
            with ch2:
                if 'AZIMUTH' in P.columns:
                    rose_tech = st.selectbox("Azimuth rose — technology",
                                             ['All'] + sorted(P['TEKNOLOJI'].dropna().unique()) if 'TEKNOLOJI' in P.columns else ['All'],
                                             key='rose_tech')
                    R = P if rose_tech == 'All' else P[P['TEKNOLOJI'] == rose_tech]
                    cnt, edges = np.histogram(R['AZIMUTH'].dropna(), bins=np.arange(0, 361, 10))
                    fr = go.Figure(go.Barpolar(r=cnt, theta=edges[:-1] + 5, width=10,
                                               marker_color=TEAL, marker_line_color='#fff'))
                    fr.update_layout(template=RF_TPL, title="Sector azimuth distribution",
                                     polar=dict(angularaxis=dict(direction='clockwise', rotation=90,
                                                                 tickvals=list(range(0, 360, 45))),
                                                radialaxis=dict(showticklabels=False)))
                    st.plotly_chart(fr, use_container_width=True)

            ch3, ch4 = st.columns(2)
            with ch3:
                if 'FREQUENCY_BAND' in P.columns:
                    bb = P['FREQUENCY_BAND'].fillna('n/a').value_counts().reset_index()
                    bb.columns = ['Band', 'Cells']
                    st.plotly_chart(px.bar(bb, x='Band', y='Cells', template=RF_TPL,
                                           color_discrete_sequence=[TEAL], title="Cells by band"),
                                    use_container_width=True)
            with ch4:
                if 'CELL_ATLAS_STATE' in P.columns:
                    sb = P['CELL_ATLAS_STATE'].fillna('n/a').value_counts().reset_index()
                    sb.columns = ['State', 'Cells']
                    st.plotly_chart(px.bar(sb, x='State', y='Cells', template=RF_TPL,
                                           color_discrete_sequence=[TEAL], title="Cells by Atlas state"),
                                    use_container_width=True)

            # Planning data-quality checks
            st.markdown("### 🧪 Planning checks")
            chk = [("Cells without coordinates", int((~has_xy).sum())),
                   ("Cells without azimuth", int((~has_az).sum())),
                   ("Rows with blank CELL_NAME", int(P['CELL_NAME'].isna().sum())),
                   ("Duplicate CELL_NAME", int(P['CELL_NAME'].dropna().duplicated().sum()))]
            if {'ADMINISTRATIVESTATE', 'OPERATIONALSTATE'} <= set(P.columns):
                chk.append(("Locked / disabled cells",
                            int(((P['ADMINISTRATIVESTATE'].astype(str).str.upper() == 'LOCKED')
                                 | (P['OPERATIONALSTATE'].astype(str).str.lower() == 'disabled')).sum())))
            if site_c and 'Plan_Latitude' in P.columns:
                xy = P.dropna(subset=['Plan_Latitude']).assign(
                    _k=lambda d: d['Plan_Latitude'].round(5).astype(str) + d['Plan_Longitude'].round(5).astype(str))
                chk.append(("Sites with cells at different coordinates",
                            int((xy.groupby(site_c)['_k'].nunique() > 1).sum())))
            _dk = [c for c in (site_c, 'FREQUENCY_BAND', 'NRARFCNDL', 'AZIMUTH') if c and c in P.columns]
            if site_c and 'AZIMUTH' in _dk:
                # same site + band + carrier + azimuth twice = likely duplicate record
                # (PCI is intentionally shared across layers of a sector, so it is not checked)
                dd = P.dropna(subset=_dk)
                chk.append(("Cells on the same site, carrier and azimuth (possible duplicates)",
                            int(dd.duplicated(_dk, keep=False).sum())))
            chk_df = pd.DataFrame(chk, columns=['Check', 'Count'])
            chk_df['Status'] = np.where(chk_df['Count'] == 0, '✅', '⚠️')
            st.dataframe(chk_df, hide_index=True, use_container_width=True)

            if not unmatched.empty:
                with st.expander(f"KPI cells with no planning record ({len(unmatched):,})"):
                    st.dataframe(unmatched, hide_index=True, use_container_width=True, height=240)
                    st.download_button("⬇️ Download list (CSV)", unmatched.to_csv(index=False).encode(),
                                       "kpi_cells_without_planning.csv", "text/csv")

            # Browser
            st.markdown("### 🗂️ Planning table")
            b1, b2, b3 = st.columns(3)
            def _opts(col): return sorted(P[col].dropna().unique()) if col in P.columns else []
            sel_r = b1.multiselect("Region", _opts('REGIONCODE'), key='pl_reg')
            sel_t = b2.multiselect("Technology", _opts('TEKNOLOJI'), key='pl_tech')
            q_txt = b3.text_input("Search cell / site name", key='pl_q')
            V = P
            if sel_r: V = V[V['REGIONCODE'].isin(sel_r)]
            if sel_t: V = V[V['TEKNOLOJI'].isin(sel_t)]
            if q_txt.strip():
                qq = q_txt.strip().upper()
                V = V[V['CELL_NAME'].astype(str).str.upper().str.contains(qq, regex=False)
                      | (V[site_c].astype(str).str.upper().str.contains(qq, regex=False) if site_c else False)]
            st.caption(f"{len(V):,} of {len(P):,} cells")
            st.dataframe(V.head(2000), use_container_width=True, height=360)
            st.download_button("⬇️ Download planning table (Excel)", to_excel(V, "Planning"),
                               "planning_cells.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    main()
