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


def load_config() -> dict[str, Any]:
    if CONFIG_PATH.is_file():
        return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    return {
        "number_map": DEFAULT_NUMBER_MAP,
        "paths": {
            "template_excel": "./Template Reporting/SPN-CMP-NA-P02-F02 - Network Assurance Acceptence Test Form FTTH_ai.xlsx"
        },
        "location_regex": r"(?:^|[^a-zA-Z0-9])([a-zA-Z]{2,5})[\s\.\-~_|()\[\]]{0,6}([0-9]{4,8})(?=[^a-zA-Z0-9]|$)",
        "blacklist_prefixes": ["IMG", "IMAGE", "LINE", "ALBUM", "PHOTO", "DSC"],
    }


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


st.set_page_config(
    page_title="ARGex — FTTH Report Exporter (1-7)",
    layout="wide",
)

st.markdown(
    """
    <style>
    .block-container { padding-top: 1.5rem; padding-bottom: 2rem; max-width: 1200px; }
    table { font-size: 0.88rem; }
    </style>
    """,
    unsafe_allow_html=True,
)

cfg = load_config()
number_map: dict[str, str] = cfg.get("number_map", DEFAULT_NUMBER_MAP)
valid_keys = set(number_map.keys())
loc_regex = cfg.get("location_regex", r"([a-zA-Z]{2,5})([0-9]{4,8})")
blacklist = {w.upper() for w in cfg.get("blacklist_prefixes", [])}
template_path = BASE_DIR / cfg["paths"]["template_excel"]

st.title("ARGex — FTTH Excel Report Exporter (Mode 1–7)")
st.caption(
    "Upload file `.zip` berisi folder DP (masing-masing foto dinamai `1` s/d `7`), periksa tabel pemetaan, lalu unduh file Excel `.xlsx` maupun paket foto yang sudah di-rename."
)

with st.expander("Referensi Urutan Nomor Foto (1 s/d 7)", expanded=True):
    cols = st.columns(7)
    for idx, (num_str, label) in enumerate(sorted(number_map.items(), key=lambda x: int(x[0]))):
        cols[idx].metric(label=f"Nomor {num_str}", value=label)

uploaded_zip = st.file_uploader(
    "Upload Archive Folder DP (.zip)",
    type=["zip"],
    help="Contoh struktur di dalam ZIP: LTS010504/1.jpg s/d 7.jpg, LTS010603/1.jpg s/d 7.jpg",
)

if uploaded_zip is not None:
    dp_results: dict[str, dict[str, Any]] = {}
    skipped_files: list[str] = []

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
                skipped_files.append(f"{loc_code}/{p.name}")
                continue

            category = number_map[slot_num]
            if loc_code not in dp_results:
                dp_results[loc_code] = {
                    "photos": {},
                    "exts": {},
                    "input_dbm": "",
                    "output_dbm": "",
                }

            if category not in dp_results[loc_code]["photos"]:
                dp_results[loc_code]["photos"][category] = zf.read(member)
                dp_results[loc_code]["exts"][category] = p.suffix

    if not dp_results:
        st.error("Tidak ditemukan foto bernomor 1 s/d 7 (.jpg/.jpeg/.png) di dalam file ZIP.")
    else:
        st.subheader(f"Ringkasan Folder DP Terdeteksi ({len(dp_results)} DP)")

        ordered_categories = [
            number_map[k] for k in sorted(number_map.keys(), key=lambda x: int(x))
        ]
        summary_rows = []
        for dp_code in sorted(dp_results.keys()):
            photos = dp_results[dp_code]["photos"]
            row_data: dict[str, Any] = {"Kode DP": dp_code}
            for idx, cat in enumerate(ordered_categories, start=1):
                row_data[f"#{idx} {cat}"] = "OK" if cat in photos else "-"
            row_data["Input DP (dBm)"] = dp_results[dp_code]["input_dbm"]
            row_data["Output DP (dBm)"] = dp_results[dp_code]["output_dbm"]
            summary_rows.append(row_data)

        edited_rows = st.data_editor(
            summary_rows,
            use_container_width=True,
            hide_index=True,
            disabled=["Kode DP"] + [f"#{i} {c}" for i, c in enumerate(ordered_categories, start=1)],
        )

        for row in edited_rows:
            code = row["Kode DP"]
            if code in dp_results:
                dp_results[code]["input_dbm"] = str(row.get("Input DP (dBm)", "")).strip() or None
                dp_results[code]["output_dbm"] = str(row.get("Output DP (dBm)", "")).strip() or None

        if skipped_files:
            with st.expander(f"File Di-skip (Bukan Nomor 1-7): {len(skipped_files)} file"):
                st.write(", ".join(skipped_files))

        excel_bytes = export_to_excel_bytes(template_path, dp_results)
        renamed_zip_bytes = build_renamed_photos_zip(dp_results)

        if excel_bytes:
            bundle_zip_bytes = build_complete_bundle_zip(dp_results, excel_bytes)
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
