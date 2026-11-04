import base64
from io import BytesIO
import json
from pathlib import Path
from typing import Any
import zipfile
import streamlit as st
from src.exporter import export_to_excel_bytes
from src.scanner import DEFAULT_NUMBER_MAP, extract_dp_code_from_path, extract_number_slot

BASE_DIR = Path(__file__).parent
CONFIG_PATH = BASE_DIR / "rules.json"

EXCEL_COL_RANGES: dict[str, str] = {
    "1": "Col J–M",
    "2": "Col N–Q",
    "3": "Col R–U",
    "4": "Col V–Y",
    "5": "Col Z–AC",
    "6": "Col AD–AG",
    "7": "Col AH–AK",
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


def resolve_template_source(local_template_path: Path) -> Path | bytes | None:
    """
    Resolusi 3 lapis untuk Master Template Excel (supaya file internal kantor aman tidak masuk repo publik):
    1. File lokal di ./Template Reporting/... (saat dijalankan di laptop lokal)
    2. Streamlit Cloud Secrets: st.secrets['TEMPLATE_EXCEL_B64']
    3. Upload manual di UI (tersimpan di st.session_state['custom_template_bytes'])
    """
    if local_template_path.is_file():
        return local_template_path
    try:
        if "TEMPLATE_EXCEL_B64" in st.secrets:
            return base64.b64decode(st.secrets["TEMPLATE_EXCEL_B64"])
    except Exception:
        pass
    return st.session_state.get("custom_template_bytes")


def build_renamed_photos_zip(dp_results: dict[str, dict[str, Any]]) -> bytes:
    """Bangun archive .zip berisi foto yang sudah di-rename per folder DP di dalam memori."""
    buf = BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for dp_code in sorted(dp_results.keys()):
            photos: dict[str, bytes] = dp_results[dp_code].get("photos", {})
            exts: dict[str, str] = dp_results[dp_code].get("exts", {})
            for category, raw_bytes in photos.items():
                ext = exts.get(category, ".jpg")
                zout.writestr(f"{dp_code}/{category}{ext}", raw_bytes)
    return buf.getvalue()


def build_complete_bundle_zip(dp_results: dict[str, dict[str, Any]], excel_bytes: bytes) -> bytes:
    """Bangun archive .zip lengkap berisi Report Excel + seluruh folder foto yang sudah di-rename."""
    buf = BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        zout.writestr("Report/Report_FTTH_Auto.xlsx", excel_bytes)
        for dp_code in sorted(dp_results.keys()):
            photos: dict[str, bytes] = dp_results[dp_code].get("photos", {})
            exts: dict[str, str] = dp_results[dp_code].get("exts", {})
            for category, raw_bytes in photos.items():
                ext = exts.get(category, ".jpg")
                zout.writestr(f"foto_renamed/{dp_code}/{category}{ext}", raw_bytes)
    return buf.getvalue()


@st.dialog("Konfirmasi Hapus DP")
def confirm_delete_dp_dialog(dp_code: str) -> None:
    """Modal pop-up konfirmasi sebelum menghapus 1 DP dari daftar report."""
    st.markdown(
        f"Apakah lu yakin ingin menghapus **`DP {dp_code}`** beserta seluruh fotonya dari daftar report saat ini?"
    )
    c_cancel, c_confirm = st.columns(2)
    with c_cancel:
        if st.button("Batal", use_container_width=True):
            st.rerun()
    with c_confirm:
        if st.button("Ya, Hapus DP", type="primary", use_container_width=True):
            st.session_state["dp_store_v2"].pop(dp_code, None)
            st.session_state["last_added_msg"] = f"DP {dp_code} telah dihapus dari daftar report."
            st.rerun()


@st.dialog("Konfirmasi Reset Semua DP")
def confirm_reset_all_dialog(total_dp: int) -> None:
    """Modal pop-up konfirmasi sebelum mengosongkan seluruh daftar DP."""
    st.markdown(
        f"Apakah lu yakin ingin mengosongkan seluruh **{total_dp} DP** dari daftar report saat ini? Tindakan ini tidak dapat dibatalkan."
    )
    c_cancel, c_confirm = st.columns(2)
    with c_cancel:
        if st.button("Batal", use_container_width=True):
            st.rerun()
    with c_confirm:
        if st.button("Ya, Kosongkan Semua", type="primary", use_container_width=True):
            st.session_state["dp_store_v2"] = {}
            st.session_state["last_added_msg"] = ""
            st.session_state["last_skipped"] = []
            st.rerun()


st.set_page_config(
    page_title="ARGex — FTTH Report Exporter",
    layout="wide",
)

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500;600&family=IBM+Plex+Sans:wght@400;500;600;700&display=swap');

    :root {
        --bg-canvas: #0D1117;
        --bg-surface: #161B22;
        --bg-elevated: #1C2128;
        --ink-primary: #E6EDF3;
        --ink-secondary: #9EA7B3;
        --ink-muted: #7D8590;
        --border-hairline: #30363D;
        --border-subtle: #21262D;
        --accent-primary: #38BDF8;
        --accent-dim: rgba(56, 189, 248, 0.12);
        --status-ok-bg: rgba(46, 160, 67, 0.15);
        --status-ok-ink: #56D364;
        --status-ok-border: rgba(46, 160, 67, 0.4);
        --status-warn-bg: rgba(210, 153, 34, 0.15);
        --status-warn-ink: #E3B341;
        --status-warn-border: rgba(210, 153, 34, 0.4);
    }

    ::selection {
        background-color: rgba(56, 189, 248, 0.28);
        color: #FFFFFF;
    }

    /* Force dark matte canvas across the entire Streamlit app & dialogs */
    .stApp, [data-testid="stAppViewContainer"], [data-testid="stHeader"] {
        background-color: var(--bg-canvas) !important;
        color: var(--ink-primary) !important;
    }
    div[role="dialog"] {
        background-color: var(--bg-surface) !important;
        border: 1px solid var(--border-hairline) !important;
        border-radius: 8px !important;
        color: var(--ink-primary) !important;
    }

    header[data-testid="stHeader"] {
        height: 2rem;
        border-bottom: none;
    }

    html, body, [class*="css"] {
        font-family: 'IBM Plex Sans', -apple-system, BlinkMacSystemFont, sans-serif;
        color: var(--ink-primary);
        font-variant-numeric: tabular-nums;
    }

    .block-container {
        padding-top: 1.25rem;
        padding-bottom: 2.5rem;
        max-width: 1180px;
    }

    /* Header bar */
    .argex-header {
        display: flex;
        justify-content: space-between;
        align-items: flex-end;
        padding-bottom: 0.9rem;
        margin-bottom: 1rem;
        border-bottom: 1px solid var(--border-hairline);
        gap: 1rem;
        flex-wrap: wrap;
    }
    .argex-title {
        font-size: 1.45rem;
        font-weight: 600;
        letter-spacing: -0.02em;
        color: var(--ink-primary);
        margin: 0;
        line-height: 1.25;
    }
    .argex-subtitle {
        font-size: 0.88rem;
        color: var(--ink-secondary);
        margin: 0.3rem 0 0 0;
        max-width: 72ch;
        line-height: 1.5;
    }
    .argex-subtitle code {
        font-family: 'IBM Plex Mono', monospace;
        font-size: 0.8rem;
        color: var(--accent-primary);
        background: var(--bg-surface);
        padding: 0.1rem 0.35rem;
        border: 1px solid var(--border-hairline);
        border-radius: 4px;
    }
    .argex-doc-badge {
        font-family: 'IBM Plex Mono', monospace;
        font-size: 0.75rem;
        font-weight: 500;
        color: var(--ink-secondary);
        background: var(--bg-surface);
        border: 1px solid var(--border-hairline);
        padding: 0.35rem 0.7rem;
        border-radius: 5px;
        white-space: nowrap;
    }

    /* Compact 7-slot specification bar */
    .slot-strip {
        display: grid;
        grid-template-columns: repeat(7, minmax(115px, 1fr));
        gap: 0;
        border: 1px solid var(--border-hairline);
        border-radius: 6px;
        background: var(--bg-surface);
        overflow-x: auto;
        margin-bottom: 1.2rem;
    }
    .slot-cell {
        padding: 0.55rem 0.7rem;
        border-right: 1px solid var(--border-subtle);
        display: flex;
        flex-direction: column;
        gap: 0.2rem;
    }
    .slot-cell:last-child {
        border-right: none;
    }
    .slot-top {
        display: flex;
        justify-content: space-between;
        align-items: center;
    }
    .slot-num {
        font-family: 'IBM Plex Mono', monospace;
        font-size: 0.74rem;
        font-weight: 600;
        color: var(--accent-primary);
        background: var(--accent-dim);
        padding: 0.08rem 0.42rem;
        border-radius: 4px;
    }
    .slot-col {
        font-family: 'IBM Plex Mono', monospace;
        font-size: 0.7rem;
        color: var(--ink-muted);
    }
    .slot-name {
        font-size: 0.84rem;
        font-weight: 500;
        color: var(--ink-primary);
    }

    /* Native Streamlit widget dark harmonization */
    [data-testid="stFileUploaderDropzone"] {
        background-color: var(--bg-surface) !important;
        border: 1px dashed var(--border-hairline) !important;
        border-radius: 6px !important;
        padding: 1.1rem 1.25rem !important;
    }
    [data-testid="stFileUploaderDropzone"]:hover {
        border-color: var(--accent-primary) !important;
        background-color: var(--bg-elevated) !important;
    }
    [data-testid="stFileUploaderDropzone"] small,
    [data-testid="stFileUploaderDropzone"] span {
        color: var(--ink-secondary) !important;
    }
    [data-testid="stExpander"] {
        background-color: var(--bg-surface) !important;
        border: 1px solid var(--border-hairline) !important;
        border-radius: 6px !important;
    }

    /* Empty state panel */
    .empty-panel {
        border: 1px solid var(--border-hairline);
        border-radius: 6px;
        background: var(--bg-surface);
        padding: 1.15rem 1.4rem;
        margin-top: 0.6rem;
        color: var(--ink-secondary);
        font-size: 0.87rem;
        line-height: 1.65;
    }
    .empty-panel strong {
        color: var(--ink-primary);
    }
    .empty-panel code {
        font-family: 'IBM Plex Mono', monospace;
        font-size: 0.8rem;
        background: var(--bg-canvas);
        padding: 0.1rem 0.35rem;
        border: 1px solid var(--border-hairline);
        border-radius: 4px;
        color: var(--accent-primary);
    }

    /* Summary status bar */
    .summary-strip {
        display: flex;
        align-items: center;
        gap: 0.65rem;
        flex-wrap: wrap;
        margin-bottom: 0.5rem;
    }
    .summary-title {
        font-size: 1.08rem;
        font-weight: 600;
        color: var(--ink-primary);
    }
    .summary-pill {
        font-family: 'IBM Plex Mono', monospace;
        font-size: 0.76rem;
        font-weight: 500;
        padding: 0.2rem 0.55rem;
        border-radius: 4px;
        border: 1px solid var(--border-hairline);
        background: var(--bg-surface);
        color: var(--ink-secondary);
    }
    .summary-pill.ok {
        background: var(--status-ok-bg);
        color: var(--status-ok-ink);
        border-color: var(--status-ok-border);
    }
    .summary-pill.warn {
        background: var(--status-warn-bg);
        color: var(--status-warn-ink);
        border-color: var(--status-warn-border);
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

st.markdown(
    """
    <div class="argex-header">
        <div>
            <h1 class="argex-title">ARGex — FTTH Acceptance Test Report Exporter</h1>
            <p class="argex-subtitle">
                Pilih folder DP berisi foto bernomor <code>1.jpg</code> s/d <code>7.jpg</code>. Setiap folder yang dipilih otomatis ditambahkan ke tabel report mulai dari baris <code>No 3</code>.
            </p>
        </div>
        <div class="argex-doc-badge">SPN/CMP/NA/P02/F02 · Captured Devices 01</div>
    </div>
    """,
    unsafe_allow_html=True,
)

sorted_slots = sorted(number_map.items(), key=lambda x: int(x[0]))
slot_cells_html = "".join(
    f'<div class="slot-cell">'
    f'<div class="slot-top"><span class="slot-num">{num_str}.jpg</span>'
    f'<span class="slot-col">{EXCEL_COL_RANGES.get(num_str, "")}</span></div>'
    f'<div class="slot-name">{label}</div>'
    f"</div>"
    for num_str, label in sorted_slots
)
st.markdown(f'<div class="slot-strip">{slot_cells_html}</div>', unsafe_allow_html=True)

template_src = resolve_template_source(local_template_path)
if template_src is None:
    uploaded_tpl = st.file_uploader(
        "Upload File Master Template Excel (.xlsx) [Cukup 1x per Sesi]",
        type=["xlsx"],
        help="Karena folder Template Reporting di-ignore demi keamanan repo publik, silakan pilih file template SPN-CMP-NA-P02-F02.xlsx atau simpan di Streamlit Secrets (TEMPLATE_EXCEL_B64).",
        key="tpl_uploader",
    )
    if uploaded_tpl is not None:
        st.session_state["custom_template_bytes"] = uploaded_tpl.getvalue()
        st.rerun()

dp_store: dict[str, dict[str, Any]] = st.session_state["dp_store_v2"]
seq = st.session_state["uploader_seq"]

if st.session_state["last_added_msg"]:
    st.success(st.session_state["last_added_msg"])

tab_folder, tab_zip, tab_direct = st.tabs(
    [
        "Tambah Folder DP (Tanpa ZIP)",
        "Upload Archive (.zip)",
        "Upload Foto Satuan (1 DP)",
    ]
)

with tab_folder:
    folder_files = st.file_uploader(
        "Pilih Folder DP untuk Ditambahkan ke Daftar Report",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files="directory",
        help="Klik 'Browse files' lalu pilih folder DP (misal LTS010504). Setelah masuk ke tabel, kotak ini otomatis di-reset agar siap menambah folder DP berikutnya.",
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

            dp_store[loc_code]["photos"][category] = uf.getvalue()
            dp_store[loc_code]["exts"][category] = p.suffix
            dp_store[loc_code]["filenames"][category] = p.name
            added_dps.add(loc_code)

        st.session_state["last_skipped"] = skipped_now
        if added_dps:
            st.session_state["last_added_msg"] = (
                f"Folder DP {', '.join(sorted(added_dps))} berhasil ditambahkan ke daftar report. "
                "Klik 'Browse files' lagi untuk menambah folder DP berikutnya."
            )
        else:
            st.session_state["last_added_msg"] = (
                "Tidak ditemukan file foto bernomor murni 1.jpg s/d 7.jpg di folder yang dipilih."
            )
        st.session_state["uploader_seq"] += 1
        st.rerun()

with tab_zip:
    uploaded_zip = st.file_uploader(
        "Upload Archive Folder DP (.zip)",
        type=["zip"],
        help="Contoh struktur di dalam ZIP: LTS010504/1.jpg s/d 7.jpg, LTS010603/1.jpg s/d 7.jpg",
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

                dp_store[loc_code]["photos"][category] = zf.read(member)
                dp_store[loc_code]["exts"][category] = p.suffix
                dp_store[loc_code]["filenames"][category] = p.name
                added_dps.add(loc_code)

        st.session_state["last_skipped"] = skipped_now
        if added_dps:
            st.session_state["last_added_msg"] = (
                f"Folder DP {', '.join(sorted(added_dps))} dari archive ZIP berhasil ditambahkan."
            )
        st.session_state["uploader_seq"] += 1
        st.rerun()

with tab_direct:
    manual_dp_code = (
        st.text_input("Kode DP", value="LTS010504", placeholder="Contoh: LTS010504")
        .strip()
        .upper()
    )

    direct_files = st.file_uploader(
        "Pilih Foto Nomor 1 s/d 7 (.jpg / .png)",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=True,
        key=f"direct_uploader_{seq}",
    )
    if st.button("Tambahkan DP ke Report", type="primary") and direct_files and manual_dp_code:
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
            dp_store[manual_dp_code]["photos"][category] = uf.getvalue()
            dp_store[manual_dp_code]["exts"][category] = p.suffix
            dp_store[manual_dp_code]["filenames"][category] = p.name

        st.session_state["last_skipped"] = skipped_now
        st.session_state["last_added_msg"] = f"DP {manual_dp_code} berhasil ditambahkan ke report."
        st.session_state["uploader_seq"] += 1
        st.rerun()

if not dp_store:
    st.markdown(
        """
        <div class="empty-panel">
            <strong>Belum ada folder DP di dalam daftar report.</strong><br>
            1. Klik tombol <code>Browse files</code> di atas lalu pilih folder DP (misal <code>LTS010504</code>).<br>
            2. Hanya foto bernomor murni <code>1.jpg</code> s/d <code>7.jpg</code> yang diambil; file sisa (<code>LINE_ALBUM_...</code>) diabaikan otomatis.<br>
            3. Di file Excel hasil export, baris <code>No 1</code> (OLT) dan <code>No 2</code> (RK) dikosongkan bersih, sedangkan daftar DP dimulai dari <code>No 3</code> (Baris 35).
        </div>
        """,
        unsafe_allow_html=True,
    )
else:
    st.divider()
    ordered_categories = [
        number_map[k] for k in sorted(number_map.keys(), key=lambda x: int(x))
    ]
    complete_count = sum(
        1 for d in dp_store.values() if len(d.get("photos", {})) == len(ordered_categories)
    )
    incomplete_count = len(dp_store) - complete_count

    hdr_col1, hdr_col2, hdr_col3, hdr_col4 = st.columns([2.8, 1.2, 0.8, 0.8])
    with hdr_col1:
        status_cls = "ok" if incomplete_count == 0 else "warn"
        status_txt = (
            f"{complete_count}/{len(dp_store)} DP Lengkap (7/7 Foto)"
            if incomplete_count == 0
            else f"{complete_count} Lengkap · {incomplete_count} Belum Lengkap"
        )
        st.markdown(
            f'<div class="summary-strip">'
            f'<span class="summary-title">Daftar DP di Report ({len(dp_store)} DP)</span>'
            f'<span class="summary-pill {status_cls}">{status_txt}</span>'
            f'<span class="summary-pill">Baris Excel: No 3 s/d No {len(dp_store) + 2}</span>'
            f"</div>",
            unsafe_allow_html=True,
        )
    with hdr_col2:
        dp_to_remove = st.selectbox(
            "Pilih DP",
            options=sorted(dp_store.keys()),
            label_visibility="collapsed",
            key="select_dp_remove",
        )
    with hdr_col3:
        if st.button("Hapus DP", type="secondary", use_container_width=True) and dp_to_remove:
            confirm_delete_dp_dialog(dp_to_remove)
    with hdr_col4:
        if st.button("Reset Semua", type="secondary", use_container_width=True):
            confirm_reset_all_dialog(len(dp_store))

    summary_rows = []
    for idx_dp, dp_code in enumerate(sorted(dp_store.keys()), start=3):
        photos = dp_store[dp_code]["photos"]
        fnames = dp_store[dp_code].get("filenames", {})
        row_data: dict[str, Any] = {
            "No": idx_dp,
            "Device Information": f"DP {dp_code}",
            "Kelengkapan": f"{len(photos)}/{len(ordered_categories)}",
        }
        for idx, cat in enumerate(ordered_categories, start=1):
            row_data[f"#{idx} {cat}"] = (
                f"OK ({fnames.get(cat, f'{idx}.jpg')})" if cat in photos else "Kosong"
            )
        summary_rows.append(row_data)

    st.dataframe(
        summary_rows,
        use_container_width=True,
        hide_index=True,
    )

    with st.expander("Preview Thumbnail Foto per DP", expanded=True):
        for dp_code in sorted(dp_store.keys()):
            photos = dp_store[dp_code]["photos"]
            fnames = dp_store[dp_code].get("filenames", {})
            st.markdown(f"**DP `{dp_code}`** ({len(photos)}/7 foto)")
            p_cols = st.columns(7)
            for idx, cat in enumerate(ordered_categories):
                with p_cols[idx]:
                    if cat in photos:
                        st.image(
                            photos[cat],
                            caption=f"#{idx+1} {cat} ({fnames.get(cat, '')})",
                            use_container_width=True,
                        )
                    else:
                        st.caption(f"#{idx+1} {cat}: Kosong")

    if st.session_state["last_skipped"]:
        with st.expander(
            f"File Sisa yang Diabaikan pada Upload Terakhir (Bukan 1.jpg s/d 7.jpg): {len(st.session_state['last_skipped'])} file"
        ):
            st.write(", ".join(st.session_state["last_skipped"]))

    renamed_zip_bytes = build_renamed_photos_zip(dp_store)
    excel_bytes = export_to_excel_bytes(template_src, dp_store) if template_src else None

    if excel_bytes:
        bundle_zip_bytes = build_complete_bundle_zip(dp_store, excel_bytes)
        btn_col1, btn_col2, btn_col3 = st.columns(3)
        with btn_col1:
            st.download_button(
                label="Download Report Excel (.xlsx)",
                data=excel_bytes,
                file_name="Report_FTTH_Auto.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )
        with btn_col2:
            st.download_button(
                label="Download Foto Renamed (.zip)",
                data=renamed_zip_bytes,
                file_name="Foto_Renamed_FTTH.zip",
                mime="application/zip",
                use_container_width=True,
            )
        with btn_col3:
            st.download_button(
                label="Download Paket Lengkap: Excel + Foto (.zip)",
                data=bundle_zip_bytes,
                file_name="Bundle_Report_dan_Foto_FTTH.zip",
                mime="application/zip",
                use_container_width=True,
            )
    else:
        st.warning(
            "Silakan upload file Master Template Excel (.xlsx) di bagian atas halaman terlebih dahulu untuk mengaktifkan tombol Download Report Excel."
        )
        st.download_button(
            label="Download Foto Renamed (.zip)",
            data=renamed_zip_bytes,
            file_name="Foto_Renamed_FTTH.zip",
            mime="application/zip",
        )
