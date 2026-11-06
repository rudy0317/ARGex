import base64
import html
import logging
from io import BytesIO
import json
from pathlib import Path
from typing import Any
import zipfile
import streamlit as st
from src.exporter import _extract_logo_bytes, export_to_excel_bytes, read_existing_dp_codes
from src.scanner import DEFAULT_NUMBER_MAP, extract_dp_code_from_path, extract_number_slot

BASE_DIR = Path(__file__).parent
CONFIG_PATH = BASE_DIR / "rules.json"

EXCEL_COL_RANGES: dict[str, str] = {
    "1": "Kolom J–M",
    "2": "Kolom N–Q",
    "3": "Kolom R–U",
    "4": "Kolom V–Y",
    "5": "Kolom Z–AC",
    "6": "Kolom AD–AG",
    "7": "Kolom AH–AK",
}


def load_config() -> dict[str, Any]:
    if CONFIG_PATH.is_file():
        return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    return {
        "number_map": DEFAULT_NUMBER_MAP,
        "paths": {
            "input_dirs": ["./Folder Aset"],
            "template_excel": "./Template Reporting/SPN-CMP-NA-P02-F02 - Network Assurance Acceptence Test Form FTTH_ai.xlsx",
        },
        "location_regex": r"(?:^|[^a-zA-Z0-9])([a-zA-Z]{2,5})[\s\.\-~_|()\[\]]{0,6}([0-9]{4,8})(?=[^a-zA-Z0-9]|$)",
        "blacklist_prefixes": ["IMG", "IMAGE", "LINE", "ALBUM", "PHOTO", "DSC"],
    }


@st.cache_data(show_spinner=False)
def _decode_template_b64(b64: str) -> bytes:
    return base64.b64decode(b64)


def resolve_template_source(local_template_path: Path) -> Path | bytes | None:
    if local_template_path.is_file():
        return local_template_path
    try:
        if "TEMPLATE_EXCEL_B64" in st.secrets:
            return _decode_template_b64(st.secrets["TEMPLATE_EXCEL_B64"])
    except Exception:
        pass
    return st.session_state.get("custom_template_bytes")


@st.cache_data(show_spinner=False, max_entries=2)
def get_logo_b64(_template_src: Path | bytes | None, tpl_key: str) -> str | None:
    if _template_src is None:
        return None
    raw = _extract_logo_bytes(_template_src)
    if not raw:
        return None
    return base64.b64encode(raw).decode("ascii")


def build_renamed_photos_zip(dp_results: dict[str, dict[str, Any]]) -> bytes:
    buf = BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_STORED) as zout:
        for dp_code in sorted(dp_results.keys()):
            photos: dict[str, bytes] = dp_results[dp_code].get("photos", {})
            exts: dict[str, str] = dp_results[dp_code].get("exts", {})
            for category, raw_bytes in photos.items():
                ext = exts.get(category, ".jpg")
                zout.writestr(f"{dp_code}/{category}{ext}", raw_bytes)
    return buf.getvalue()


def build_complete_bundle_zip(dp_results: dict[str, dict[str, Any]], excel_bytes: bytes) -> bytes:
    buf = BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_STORED) as zout:
        zout.writestr("Report/Report_FTTH_Auto.xlsx", excel_bytes)
        for dp_code in sorted(dp_results.keys()):
            photos: dict[str, bytes] = dp_results[dp_code].get("photos", {})
            exts: dict[str, str] = dp_results[dp_code].get("exts", {})
            for category, raw_bytes in photos.items():
                ext = exts.get(category, ".jpg")
                zout.writestr(f"foto_renamed/{dp_code}/{category}{ext}", raw_bytes)
    return buf.getvalue()


MAX_PHOTO_PX = 1600
JPEG_QUALITY = 85


def compress_photo(raw: bytes, ext: str) -> tuple[bytes, str]:
    try:
        from PIL import Image, ImageOps

        with Image.open(BytesIO(raw)) as im:
            im = ImageOps.exif_transpose(im)
            im.thumbnail((MAX_PHOTO_PX, MAX_PHOTO_PX))
            out = BytesIO()
            im.convert("RGB").save(out, format="JPEG", quality=JPEG_QUALITY, optimize=True)
        data = out.getvalue()
        return (data, ".jpg") if len(data) < len(raw) or ext.lower() == ".png" else (raw, ext)
    except Exception:
        logging.exception("compress_photo gagal, pakai file asli")
        return raw, ext


def safe_build(label: str, fn: Any, *args: Any) -> Any:
    try:
        return fn(*args)
    except Exception as exc:
        logging.exception("Gagal membangun %s", label)
        st.error(
            f"Gagal membuat {label}: {type(exc).__name__}: {exc}. "
            "Coba kurangi jumlah DP / ukuran foto, lalu muat ulang halaman."
        )
        return None


def store_signature(dp_results: dict[str, dict[str, Any]]) -> tuple:
    return tuple(
        (dp, cat, len(raw))
        for dp in sorted(dp_results)
        for cat, raw in sorted(dp_results[dp].get("photos", {}).items())
    )


@st.cache_data(show_spinner="Menyusun Excel...", max_entries=2)
def cached_excel(
    _template: Path | bytes, tpl_key: str, sig: tuple, _dp_results: dict, _base: bytes | None, base_key: int
) -> bytes:
    return export_to_excel_bytes(_template, _dp_results, _base)


@st.cache_data(show_spinner=False, max_entries=4)
def cached_existing_codes(report_bytes: bytes) -> list[str]:
    return read_existing_dp_codes(report_bytes)


@st.cache_data(show_spinner=False, max_entries=2)
def cached_renamed_zip(sig: tuple, _dp_results: dict) -> bytes:
    return build_renamed_photos_zip(_dp_results)


@st.cache_data(show_spinner=False, max_entries=2)
def cached_bundle_zip(sig: tuple, tpl_key: str, _dp_results: dict, _excel: bytes, excel_key: int) -> bytes:
    return build_complete_bundle_zip(_dp_results, _excel)


@st.cache_data(show_spinner=False, max_entries=300)
def make_thumbnail(raw: bytes, max_px: int = 360) -> bytes:
    from PIL import Image, ImageOps

    with Image.open(BytesIO(raw)) as im:
        im = ImageOps.exif_transpose(im).convert("RGB")
        im.thumbnail((max_px, max_px))
        out = BytesIO()
        im.save(out, format="JPEG", quality=70)
        return out.getvalue()


@st.dialog("Hapus DP")
def confirm_delete_dp_dialog(dp_code: str) -> None:
    st.markdown(f"Hapus **DP {dp_code}** dari daftar report saat ini?")
    c_cancel, c_confirm = st.columns(2)
    with c_cancel:
        if st.button("Batal", use_container_width=True):
            st.rerun()
    with c_confirm:
        if st.button("Ya, Hapus", type="primary", use_container_width=True):
            st.session_state["dp_store_v2"].pop(dp_code, None)
            st.session_state["last_added_msg"] = f"DP {dp_code} dihapus."
            st.rerun()


@st.dialog("Reset Semua DP")
def confirm_reset_all_dialog(total_dp: int) -> None:
    st.markdown(f"Kosongkan seluruh **{total_dp} DP** dari daftar report saat ini?")
    c_cancel, c_confirm = st.columns(2)
    with c_cancel:
        if st.button("Batal", use_container_width=True):
            st.rerun()
    with c_confirm:
        if st.button("Ya, Kosongkan", type="primary", use_container_width=True):
            st.session_state["dp_store_v2"] = {}
            st.session_state["last_added_msg"] = ""
            st.session_state["last_skipped"] = []
            st.rerun()


st.set_page_config(
    page_title="Biznet — FTTH Acceptance Report Exporter",
    layout="wide",
)

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap');

    :root {
        --bg-page: #F4F7FB;
        --surface: #FFFFFF;
        --surface-soft: #F8FAFC;
        --biznet-navy: #002B5B;
        --biznet-blue: #0095DA;
        --biznet-blue-hover: #0077B6;
        --biznet-blue-soft: #E6F4FB;
        --ink-main: #0F172A;
        --ink-sub: #475569;
        --ink-muted: #64748B;
        --border: #E2E8F0;
        --border-strong: #CBD5E1;
        --ok-ink: #047857;
        --ok-bg: #ECFDF5;
        --ok-border: #A7F3D0;
        --warn-ink: #B45309;
        --warn-bg: #FFFBEB;
        --warn-border: #FDE68A;
        --ease: cubic-bezier(0.22, 1, 0.36, 1);
    }

    [data-testid="stToolbar"],
    [data-testid="stToolbarActions"],
    [data-testid="stDecoration"],
    #MainMenu,
    footer {
        visibility: hidden !important;
        height: 0 !important;
    }

    header[data-testid="stHeader"] {
        background: transparent !important;
        height: 0 !important;
        border: none !important;
    }

    .stApp, [data-testid="stAppViewContainer"] {
        background-color: var(--bg-page) !important;
        color: var(--ink-main) !important;
    }

    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif;
        color: var(--ink-main);
        font-variant-numeric: tabular-nums;
        -webkit-font-smoothing: antialiased;
    }

    .block-container {
        padding-top: 1.25rem !important;
        padding-bottom: 3.5rem !important;
        max-width: 1340px !important;
    }

    /* Biznet Corporate Top Navbar */
    .biznet-navbar {
        background: var(--surface);
        border: 1px solid var(--border);
        border-radius: 12px;
        padding: 0.9rem 1.4rem;
        margin-bottom: 1.5rem;
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 1rem;
        box-shadow: 0 2px 10px rgba(0, 43, 91, 0.04);
        flex-wrap: wrap;
    }
    .biznet-nav-left {
        display: flex;
        align-items: center;
        gap: 1.1rem;
    }
    .biznet-logo {
        height: 38px;
        width: auto;
        object-fit: contain;
        display: block;
    }
    .biznet-wordmark {
        font-size: 1.35rem;
        font-weight: 800;
        letter-spacing: -0.03em;
        color: var(--biznet-navy);
    }
    .biznet-wordmark span {
        color: var(--biznet-blue);
    }
    .biznet-divider {
        width: 1px;
        height: 28px;
        background: var(--border);
    }
    .biznet-nav-title {
        font-size: 1.02rem;
        font-weight: 700;
        color: var(--biznet-navy);
        margin: 0;
        line-height: 1.25;
    }
    .biznet-nav-sub {
        font-size: 0.78rem;
        color: var(--ink-muted);
        margin: 0.1rem 0 0 0;
    }
    .biznet-status-pill {
        font-size: 0.76rem;
        font-weight: 600;
        padding: 0.35rem 0.85rem;
        border-radius: 999px;
        background: var(--biznet-blue-soft);
        color: var(--biznet-navy);
        border: 1px solid rgba(0, 149, 218, 0.25);
    }

    /* Section Card Wrapper via Streamlit column styling */
    .section-heading {
        font-size: 0.95rem;
        font-weight: 700;
        color: var(--biznet-navy);
        margin: 0 0 0.75rem 0;
    }

    /* Tabs — Clean Biznet Pill/Segment Style */
    [data-testid="stTabs"] [data-baseweb="tab-list"] {
        background: var(--surface) !important;
        border: 1px solid var(--border) !important;
        padding: 4px !important;
        border-radius: 10px !important;
        gap: 4px !important;
        margin-bottom: 0.85rem !important;
    }
    [data-testid="stTabs"] [data-baseweb="tab"] {
        border-radius: 7px !important;
        padding: 0.45rem 0.95rem !important;
        font-size: 0.82rem !important;
        font-weight: 600 !important;
        color: var(--ink-sub) !important;
        border: none !important;
        background: transparent !important;
        transition: background-color 180ms var(--ease), color 180ms var(--ease) !important;
    }
    [data-testid="stTabs"] [data-baseweb="tab"]:hover {
        color: var(--biznet-navy) !important;
        background: var(--surface-soft) !important;
    }
    [data-testid="stTabs"] [aria-selected="true"] {
        background: var(--biznet-blue) !important;
        color: #FFFFFF !important;
    }
    [data-testid="stTabs"] [data-baseweb="tab-highlight"],
    [data-testid="stTabs"] [data-baseweb="tab-border"] {
        display: none !important;
    }

    /* File Uploader Dropzone — Clean White Card */
    [data-testid="stFileUploader"] {
        background: var(--surface) !important;
        border: 1px solid var(--border) !important;
        border-radius: 12px !important;
        padding: 0.85rem 1rem !important;
        box-shadow: 0 2px 8px rgba(0, 43, 91, 0.03) !important;
    }
    [data-testid="stFileUploader"] label {
        color: var(--biznet-navy) !important;
        font-weight: 600 !important;
    }
    [data-testid="stFileUploaderDropzone"] {
        background: var(--surface-soft) !important;
        border: 1.5px dashed var(--border-strong) !important;
        border-radius: 10px !important;
        padding: 1.25rem !important;
        transition: border-color 180ms var(--ease), background-color 180ms var(--ease) !important;
    }
    [data-testid="stFileUploaderDropzone"]:hover {
        border-color: var(--biznet-blue) !important;
        background: var(--biznet-blue-soft) !important;
    }
    [data-testid="stFileUploaderDropzone"] small,
    [data-testid="stFileUploaderDropzone"] span {
        color: var(--ink-sub) !important;
    }
    [data-testid="stFileUploaderDropzone"] button {
        border-radius: 8px !important;
        background: var(--surface) !important;
        border: 1px solid var(--biznet-blue) !important;
        color: var(--biznet-blue) !important;
        font-weight: 600 !important;
        font-size: 0.8rem !important;
    }
    [data-testid="stFileUploaderDropzone"] button:hover {
        background: var(--biznet-blue) !important;
        color: #FFFFFF !important;
    }

    /* Buttons — Biznet Cyan-Blue Primary & Clean White Secondary */
    .stButton > button,
    .stDownloadButton > button {
        border-radius: 8px !important;
        padding: 0.55rem 1rem !important;
        font-family: 'Plus Jakarta Sans', sans-serif !important;
        font-weight: 600 !important;
        font-size: 0.84rem !important;
        border: 1px solid var(--border-strong) !important;
        background: var(--surface) !important;
        color: var(--biznet-navy) !important;
        box-shadow: 0 1px 2px rgba(0, 43, 91, 0.04) !important;
        transition: all 160ms var(--ease) !important;
    }
    .stButton > button:hover,
    .stDownloadButton > button:hover {
        border-color: var(--biznet-blue) !important;
        color: var(--biznet-blue) !important;
        background: var(--surface-soft) !important;
        transform: translateY(-1px) !important;
    }
    .stButton > button:active,
    .stDownloadButton > button:active {
        transform: scale(0.99) !important;
    }
    .stButton > button[kind="primary"],
    .stDownloadButton > button[kind="primary"] {
        background: var(--biznet-blue) !important;
        color: #FFFFFF !important;
        border: 1px solid var(--biznet-blue) !important;
        box-shadow: 0 4px 12px rgba(0, 149, 218, 0.22) !important;
    }
    .stButton > button[kind="primary"]:hover,
    .stDownloadButton > button[kind="primary"]:hover {
        background: var(--biznet-blue-hover) !important;
        border-color: var(--biznet-blue-hover) !important;
        color: #FFFFFF !important;
    }

    /* Expanders, Inputs & Selectboxes — Force Clean Light Theme */
    [data-testid="stExpander"] {
        background: var(--surface) !important;
        border: 1px solid var(--border) !important;
        border-radius: 10px !important;
        color: var(--ink-main) !important;
        box-shadow: 0 1px 4px rgba(0, 43, 91, 0.03) !important;
    }
    [data-testid="stExpander"] summary,
    [data-testid="stExpander"] summary span,
    [data-testid="stExpander"] p,
    [data-testid="stCaptionContainer"],
    [data-testid="stMarkdownContainer"] p,
    label {
        color: var(--ink-main) !important;
    }
    [data-baseweb="select"] > div,
    [data-baseweb="input"] > div {
        background-color: var(--surface) !important;
        border-color: var(--border-strong) !important;
        color: var(--ink-main) !important;
        border-radius: 8px !important;
    }
    div[role="dialog"] {
        background: var(--surface) !important;
        border: 1px solid var(--border) !important;
        border-radius: 14px !important;
        color: var(--ink-main) !important;
        box-shadow: 0 20px 40px rgba(0, 43, 91, 0.12) !important;
    }

    /* Right Column Empty Card */
    .empty-card {
        background: var(--surface);
        border: 1px solid var(--border);
        border-radius: 12px;
        padding: 2.5rem 2rem;
        box-shadow: 0 2px 10px rgba(0, 43, 91, 0.03);
    }
    .empty-card h4 {
        color: var(--biznet-navy);
        font-size: 1rem;
        font-weight: 700;
        margin: 0 0 0.4rem 0;
    }
    .empty-card p {
        color: var(--ink-sub);
        font-size: 0.86rem;
        line-height: 1.6;
        margin: 0 0 1.1rem 0;
    }
    .slot-grid-ref {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(135px, 1fr));
        gap: 0.5rem;
    }
    .slot-ref-item {
        background: var(--surface-soft);
        border: 1px solid var(--border);
        border-radius: 8px;
        padding: 0.5rem 0.65rem;
        font-size: 0.78rem;
    }
    .slot-ref-item strong {
        color: var(--biznet-blue);
        display: block;
        font-size: 0.74rem;
    }
    .slot-ref-item span {
        color: var(--biznet-navy);
        font-weight: 600;
    }

    /* Table Header Bar */
    .table-bar {
        display: flex;
        align-items: center;
        gap: 0.55rem;
        flex-wrap: wrap;
        margin-bottom: 0.5rem;
    }
    .table-count {
        font-size: 1rem;
        font-weight: 700;
        color: var(--biznet-navy);
    }
    .stat-badge {
        font-size: 0.75rem;
        font-weight: 600;
        padding: 0.2rem 0.65rem;
        border-radius: 999px;
        border: 1px solid var(--border);
        background: var(--surface);
        color: var(--ink-sub);
    }
    .stat-badge.ok {
        color: var(--ok-ink);
        background: var(--ok-bg);
        border-color: var(--ok-border);
    }
    .stat-badge.warn {
        color: var(--warn-ink);
        background: var(--warn-bg);
        border-color: var(--warn-border);
    }

    /* Corporate Data Table */
    .dp-table-wrap {
        background: var(--surface);
        border: 1px solid var(--border);
        border-radius: 12px;
        overflow-x: auto;
        margin-bottom: 1rem;
        box-shadow: 0 2px 10px rgba(0, 43, 91, 0.03);
    }
    .dp-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 0.82rem;
    }
    .dp-table th {
        font-size: 0.73rem;
        font-weight: 700;
        color: var(--biznet-navy);
        text-align: left;
        padding: 0.7rem 0.75rem;
        border-bottom: 1px solid var(--border);
        background: var(--surface-soft);
        white-space: nowrap;
    }
    .dp-table td {
        padding: 0.65rem 0.75rem;
        border-bottom: 1px solid var(--border);
        font-size: 0.8rem;
        color: var(--ink-sub);
        white-space: nowrap;
    }
    .dp-table tbody tr:last-child td {
        border-bottom: none;
    }
    .dp-table tbody tr:hover {
        background-color: var(--surface-soft);
    }
    .cell-dp {
        font-weight: 700;
        color: var(--biznet-navy) !important;
    }
    .cell-ok {
        color: var(--ok-ink) !important;
        font-weight: 600;
    }
    .cell-miss {
        color: #94A3B8 !important;
    }

    [data-testid="stImage"] img {
        border-radius: 8px !important;
        border: 1px solid var(--border) !important;
    }

    /* Subtle Author Credit */
    .argex-wm {
        position: fixed;
        right: 16px;
        bottom: 12px;
        z-index: 45;
        font-size: 0.74rem;
        font-weight: 500;
        color: var(--ink-muted);
        background: var(--surface);
        border: 1px solid var(--border);
        border-radius: 999px;
        padding: 0.28rem 0.75rem;
        box-shadow: 0 2px 8px rgba(0, 43, 91, 0.06);
    }
    .argex-wm a {
        color: var(--biznet-blue);
        font-weight: 600;
        text-decoration: none;
    }
    .argex-wm a:hover {
        text-decoration: underline;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

cfg = load_config()
number_map: dict[str, str] = cfg.get("number_map", DEFAULT_NUMBER_MAP)
valid_keys = set(number_map.keys())
loc_regex = cfg.get("location_regex", r"([a-zA-Z]{2,5})([0-9]{4,8})")
blacklist = {w.upper() for w in cfg.get("blacklist_prefixes", [])}
local_template_path = BASE_DIR / cfg["paths"]["template_excel"]

if "dp_store_v2" not in st.session_state:
    st.session_state["dp_store_v2"] = {}
if "uploader_seq" not in st.session_state:
    st.session_state["uploader_seq"] = 0
if "last_added_msg" not in st.session_state:
    st.session_state["last_added_msg"] = ""
if "last_skipped" not in st.session_state:
    st.session_state["last_skipped"] = []

template_src = resolve_template_source(local_template_path)
tpl_key = str(template_src) if isinstance(template_src, Path) else f"bytes:{len(template_src or b'')}"
logo_b64 = get_logo_b64(template_src, tpl_key)

dp_store: dict[str, dict[str, Any]] = st.session_state["dp_store_v2"]
seq = st.session_state["uploader_seq"]

logo_html = (
    f'<img class="biznet-logo" src="data:image/png;base64,{logo_b64}" alt="Biznet">'
    if logo_b64
    else '<div class="biznet-wordmark">Biz<span>net</span></div>'
)

st.markdown(
    f"""
    <div class="biznet-navbar">
        <div class="biznet-nav-left">
            {logo_html}
            <div class="biznet-divider"></div>
            <div>
                <h1 class="biznet-nav-title">FTTH Acceptance Test Report Exporter</h1>
                <p class="biznet-nav-sub">Network Assurance · Form SPN/CMP/NA/P02/F02 (Captured Devices 01)</p>
            </div>
        </div>
        <div class="biznet-status-pill">{len(dp_store)} DP dalam Antrean</div>
    </div>
    <div class="argex-wm">coded by <a href="https://github.com/rudy0317" target="_blank" rel="noopener noreferrer">github.com/rudy0317</a></div>
    """,
    unsafe_allow_html=True,
)

col_ingest, col_inspect = st.columns([1, 1.55], gap="large")

with col_ingest:
    st.markdown('<div class="section-heading">Upload Folder DP</div>', unsafe_allow_html=True)

    if template_src is None:
        uploaded_tpl = st.file_uploader(
            "Master Template Excel (.xlsx) [1x per sesi]",
            type=["xlsx"],
            key="tpl_uploader",
        )
        if uploaded_tpl is not None:
            st.session_state["custom_template_bytes"] = uploaded_tpl.getvalue()
            st.rerun()

    tab_folder, tab_zip, tab_direct = st.tabs(
        [
            "Folder DP",
            "Archive .zip",
            "Satuan (1 DP)",
        ]
    )

    with tab_folder:
        folder_files = st.file_uploader(
            "Pilih Folder DP (1.jpg s/d 7.jpg)",
            type=["jpg", "jpeg", "png"],
            accept_multiple_files="directory",
            help="Pilih folder DP (contoh: LTS010504). File di luar nomor 1–7 diabaikan otomatis.",
            key=f"folder_uploader_{seq}",
        )
        if folder_files:
            added_dps: set[str] = set()
            skipped_now: list[str] = []
            for uf in sorted(folder_files, key=lambda x: x.name):
                p = Path(uf.name)
                if p.suffix.lower() not in (".jpg", ".jpeg", ".png"):
                    continue

                slot_num = extract_number_slot(p.stem, valid_keys)
                loc_code = extract_dp_code_from_path(p.parts, p.name, loc_regex, blacklist)

                if not slot_num:
                    skipped_now.append(f"{loc_code}/{p.name}")
                    continue

                category = number_map[slot_num]
                if loc_code not in dp_store:
                    dp_store[loc_code] = {
                        "photos": {},
                        "exts": {},
                        "filenames": {},
                    }

                photo_bytes, photo_ext = compress_photo(uf.getvalue(), p.suffix)
                dp_store[loc_code]["photos"][category] = photo_bytes
                dp_store[loc_code]["exts"][category] = photo_ext
                dp_store[loc_code]["filenames"][category] = p.name
                added_dps.add(loc_code)

            st.session_state["last_skipped"] = skipped_now
            if added_dps:
                st.session_state["last_added_msg"] = (
                    f"Berhasil menambahkan DP {', '.join(sorted(added_dps))}"
                )
            else:
                st.session_state["last_added_msg"] = (
                    "Tidak ditemukan file 1.jpg s/d 7.jpg di folder tersebut."
                )
            st.session_state["uploader_seq"] += 1
            st.rerun()

    with tab_zip:
        uploaded_zip = st.file_uploader(
            "Pilih Archive (.zip)",
            type=["zip"],
            help="Struktur di dalam ZIP: LTS010504/1.jpg s/d 7.jpg",
            key=f"zip_uploader_{seq}",
        )
        if uploaded_zip is not None:
            added_dps = set()
            skipped_now = []
            with zipfile.ZipFile(uploaded_zip, "r") as zf:
                for member in sorted(zf.namelist()):
                    if member.endswith("/") or "__MACOSX" in member:
                        continue
                    p = Path(member)
                    if p.suffix.lower() not in (".jpg", ".jpeg", ".png"):
                        continue

                    slot_num = extract_number_slot(p.stem, valid_keys)
                    loc_code = extract_dp_code_from_path(p.parts, p.name, loc_regex, blacklist)

                    if not slot_num:
                        skipped_now.append(f"{loc_code}/{p.name}")
                        continue

                    category = number_map[slot_num]
                    if loc_code not in dp_store:
                        dp_store[loc_code] = {
                            "photos": {},
                            "exts": {},
                            "filenames": {},
                        }

                    photo_bytes, photo_ext = compress_photo(zf.read(member), p.suffix)
                    dp_store[loc_code]["photos"][category] = photo_bytes
                    dp_store[loc_code]["exts"][category] = photo_ext
                    dp_store[loc_code]["filenames"][category] = p.name
                    added_dps.add(loc_code)

            st.session_state["last_skipped"] = skipped_now
            if added_dps:
                st.session_state["last_added_msg"] = (
                    f"Berhasil menambahkan DP {', '.join(sorted(added_dps))} dari ZIP"
                )
            st.session_state["uploader_seq"] += 1
            st.rerun()

    with tab_direct:
        manual_dp_code = (
            st.text_input("Kode DP", value="LTS010504", placeholder="LTS010504")
            .strip()
            .upper()
        )
        direct_files = st.file_uploader(
            "Pilih Foto 1 s/d 7 (.jpg / .png)",
            type=["jpg", "jpeg", "png"],
            accept_multiple_files=True,
            key=f"direct_uploader_{seq}",
        )
        if st.button("Tambahkan DP", type="primary", use_container_width=True) and direct_files and manual_dp_code:
            if manual_dp_code not in dp_store:
                dp_store[manual_dp_code] = {
                    "photos": {},
                    "exts": {},
                    "filenames": {},
                }
            skipped_now = []
            for uf in direct_files:
                p = Path(uf.name)
                slot_num = extract_number_slot(p.stem, valid_keys)
                if not slot_num:
                    skipped_now.append(f"{manual_dp_code}/{uf.name}")
                    continue
                category = number_map[slot_num]
                photo_bytes, photo_ext = compress_photo(uf.getvalue(), p.suffix)
                dp_store[manual_dp_code]["photos"][category] = photo_bytes
                dp_store[manual_dp_code]["exts"][category] = photo_ext
                dp_store[manual_dp_code]["filenames"][category] = p.name

            st.session_state["last_skipped"] = skipped_now
            st.session_state["last_added_msg"] = f"DP {manual_dp_code} berhasil ditambahkan."
            st.session_state["uploader_seq"] += 1
            st.rerun()

    with st.expander("Lanjutkan dari Report .xlsx Sebelumnya (opsional)"):
        base_up = st.file_uploader(
            "File .xlsx batch sebelumnya",
            type=["xlsx"],
            key="base_report_uploader",
        )
        if base_up is not None:
            st.session_state["base_report_bytes"] = base_up.getvalue()
        base_report_bytes: bytes | None = st.session_state.get("base_report_bytes")
        if base_report_bytes:
            try:
                base_codes = cached_existing_codes(base_report_bytes)
                st.info(f"Report lama berisi {len(base_codes)} DP. DP baru mulai dari No {len(base_codes) + 3}.")
            except Exception as exc:
                st.error(f"Gagal membaca report lama: {type(exc).__name__}")
                base_report_bytes = None
                base_codes = []
            if st.button("Lepas Report Lama", use_container_width=True):
                st.session_state["base_report_bytes"] = None
                st.rerun()
        else:
            base_codes = []

    if st.session_state["last_added_msg"]:
        st.success(st.session_state["last_added_msg"])

    if dp_store:
        st.markdown('<div class="section-heading" style="margin-top:1.1rem;">Unduh Laporan</div>', unsafe_allow_html=True)
        sig = store_signature(dp_store)
        renamed_zip_bytes = safe_build("ZIP foto", cached_renamed_zip, sig, dp_store) or b""
        dup_codes = sorted(set(dp_store) & set(base_codes))
        if dup_codes:
            st.warning(f"Dilewati (sudah ada di report lama): {', '.join(dup_codes)}")

        excel_bytes = (
            safe_build(
                "Excel",
                cached_excel,
                template_src,
                tpl_key,
                sig,
                dp_store,
                base_report_bytes,
                len(base_report_bytes or b""),
            )
            if (template_src or base_report_bytes)
            else None
        )

        if excel_bytes:
            bundle_zip_bytes = (
                safe_build("ZIP paket lengkap", cached_bundle_zip, sig, tpl_key, dp_store, excel_bytes, len(excel_bytes))
                or b""
            )
            st.download_button(
                label="Download Report Excel (.xlsx)",
                data=excel_bytes,
                file_name="Report_FTTH_Auto.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )
            dl_c1, dl_c2 = st.columns(2)
            with dl_c1:
                st.download_button(
                    label="Foto Renamed (.zip)",
                    data=renamed_zip_bytes,
                    file_name="Foto_Renamed_FTTH.zip",
                    mime="application/zip",
                    use_container_width=True,
                )
            with dl_c2:
                st.download_button(
                    label="Paket Lengkap (.zip)",
                    data=bundle_zip_bytes,
                    file_name="Bundle_Report_dan_Foto_FTTH.zip",
                    mime="application/zip",
                    use_container_width=True,
                )
        else:
            st.warning("Upload Master Template Excel (.xlsx) di atas untuk mengaktifkan export Excel.")
            st.download_button(
                label="Foto Renamed (.zip)",
                data=renamed_zip_bytes,
                file_name="Foto_Renamed_FTTH.zip",
                mime="application/zip",
                use_container_width=True,
            )

with col_inspect:
    st.markdown('<div class="section-heading">Daftar DP pada Report</div>', unsafe_allow_html=True)

    sorted_slots = sorted(number_map.items(), key=lambda x: int(x[0]))

    if not dp_store:
        slot_cards_html = "".join(
            f'<div class="slot-ref-item"><strong>{num_str}.jpg · {EXCEL_COL_RANGES.get(num_str, "")}</strong>'
            f"<span>{label}</span></div>"
            for num_str, label in sorted_slots
        )
        st.markdown(
            f"""
            <div class="empty-card">
                <h4>Belum ada folder DP yang dipilih</h4>
                <p>
                    Pilih folder DP di sebelah kiri untuk mulai menyusun laporan Excel.
                    Pastikan foto di dalam folder sudah diberi nama nomor 1 sampai 7 sesuai urutan kolom berikut:
                </p>
                <div class="slot-grid-ref">{slot_cards_html}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    else:
        ordered_categories = [
            number_map[k] for k in sorted(number_map.keys(), key=lambda x: int(x))
        ]
        complete_count = sum(
            1 for d in dp_store.values() if len(d.get("photos", {})) == len(ordered_categories)
        )
        incomplete_count = len(dp_store) - complete_count
        start_no = len(base_codes) + 3
        end_no = start_no + len(dp_store) - 1

        hdr_c1, hdr_c2, hdr_c3, hdr_c4 = st.columns([2.4, 1.2, 0.7, 0.7])
        with hdr_c1:
            badge_cls = "ok" if incomplete_count == 0 else "warn"
            badge_txt = (
                f"{complete_count}/{len(dp_store)} Lengkap"
                if incomplete_count == 0
                else f"{complete_count} Lengkap · {incomplete_count} Belum Lengkap"
            )
            st.markdown(
                f'<div class="table-bar">'
                f'<span class="table-count">{len(dp_store)} Folder DP</span>'
                f'<span class="stat-badge {badge_cls}">{badge_txt}</span>'
                f'<span class="stat-badge">Baris Excel No {start_no}–{end_no}</span>'
                f"</div>",
                unsafe_allow_html=True,
            )
        with hdr_c2:
            dp_to_remove = st.selectbox(
                "Pilih DP",
                options=sorted(dp_store.keys()),
                label_visibility="collapsed",
                key="select_dp_remove",
            )
        with hdr_c3:
            if st.button("Hapus", use_container_width=True) and dp_to_remove:
                confirm_delete_dp_dialog(dp_to_remove)
        with hdr_c4:
            if st.button("Reset", use_container_width=True):
                confirm_reset_all_dialog(len(dp_store))

        th_cells = "".join(f"<th>#{i} {cat}</th>" for i, cat in enumerate(ordered_categories, start=1))
        body_rows = []
        for idx_dp, dp_code in enumerate(sorted(dp_store.keys()), start=start_no):
            photos = dp_store[dp_code]["photos"]
            fnames = dp_store[dp_code].get("filenames", {})
            safe_code = html.escape(dp_code)
            slot_tds = []
            for idx, cat in enumerate(ordered_categories, start=1):
                if cat in photos:
                    fn = html.escape(fnames.get(cat, f"{idx}.jpg"))
                    slot_tds.append(f'<td class="cell-ok" title="{fn}">✓ {fn}</td>')
                else:
                    slot_tds.append('<td class="cell-miss">—</td>')
            body_rows.append(
                f"<tr>"
                f"<td>{idx_dp}</td>"
                f'<td class="cell-dp">{safe_code}</td>'
                f"<td>{len(photos)}/7</td>"
                f"{''.join(slot_tds)}"
                f"</tr>"
            )

        st.markdown(
            f'<div class="dp-table-wrap"><table class="dp-table">'
            f"<thead><tr><th>No</th><th>Kode DP</th><th>Slot</th>{th_cells}</tr></thead>"
            f"<tbody>{''.join(body_rows)}</tbody>"
            f"</table></div>",
            unsafe_allow_html=True,
        )

        with st.expander("Preview Foto per DP", expanded=True):
            for dp_code in sorted(dp_store.keys()):
                photos = dp_store[dp_code]["photos"]
                fnames = dp_store[dp_code].get("filenames", {})
                st.markdown(f"**DP {dp_code}** ({len(photos)}/7 foto)")
                p_cols = st.columns(7)
                for idx, cat in enumerate(ordered_categories):
                    with p_cols[idx]:
                        if cat in photos:
                            st.image(
                                make_thumbnail(photos[cat]),
                                caption=f"#{idx+1} {cat}",
                                use_container_width=True,
                            )
                        else:
                            st.caption(f"#{idx+1} Kosong")

        if st.session_state["last_skipped"]:
            with st.expander(f"File sisa yang diabaikan ({len(st.session_state['last_skipped'])})"):
                st.write(", ".join(st.session_state["last_skipped"]))
