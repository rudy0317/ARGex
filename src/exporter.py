from copy import copy
from io import BytesIO
from pathlib import Path
from typing import Any
import zipfile
import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles.cell_style import StyleArray
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor

# Mapping kategori ke (col_start_0idx, col_end_0idx) pada sheet 'Captured Devices 01'
CATEGORY_COL_MAP: dict[str, tuple[int, int]] = {
    "Output RK": (9, 12),    # J .. M
    "Input DP": (13, 16),    # N .. Q
    "Output DP": (17, 20),   # R .. U
    "Loop DP": (21, 24),     # V .. Y
    "Basetray DP": (25, 28), # Z .. AC
    "DP": (29, 32),          # AD .. AG
    "Structure": (33, 36),   # AH .. AK
}

DP_COL_HEADERS: dict[int, str] = {
    10: "Output RK",
    14: "Input DP",
    18: "Output DP",
    22: "Loop DP",
    26: "Basetray DP",
    30: "DP",
    34: "Structure",
}


def _extract_logo_bytes(template_src: Path | bytes) -> bytes | None:
    """Ambil file logo Biznet dari dalam archive template xlsx (baik Path maupun bytes)."""
    try:
        z_target = BytesIO(template_src) if isinstance(template_src, (bytes, bytearray)) else template_src
        with zipfile.ZipFile(z_target, "r") as z:
            for candidate in ("xl/media/image1.png", "xl/media/image25.png"):
                if candidate in z.namelist():
                    return z.read(candidate)
    except Exception:
        return None
    return None


def _copy_dp_block_style(ws: Any, src_base_row: int, dst_base_row: int) -> None:
    """Salin struktur merge cell, border, font, dan label DP dari blok referensi (baris 35) ke baris baru."""
    for r_offset in range(10):
        src_r = src_base_row + r_offset
        dst_r = dst_base_row + r_offset
        ws.row_dimensions[dst_r].height = ws.row_dimensions[src_r].height
        for col_idx in range(2, 38):
            src_cell = ws.cell(src_r, col_idx)
            dst_cell = ws.cell(dst_r, col_idx)
            if src_cell.has_style:
                dst_cell.font = copy(src_cell.font)
                dst_cell.fill = copy(src_cell.fill)
                dst_cell.border = copy(src_cell.border)
                dst_cell.alignment = copy(src_cell.alignment)

    for col_idx, header_text in DP_COL_HEADERS.items():
        ws.cell(dst_base_row, col_idx).value = header_text

    end_r = dst_base_row + 8
    merge_ranges = [
        f"B{dst_base_row}:B{end_r}",
        f"C{dst_base_row}:I{end_r}",
        f"J{dst_base_row}:M{end_r}",
        f"N{dst_base_row}:Q{end_r}",
        f"R{dst_base_row}:U{end_r}",
        f"V{dst_base_row}:Y{end_r}",
        f"Z{dst_base_row}:AC{end_r}",
        f"AD{dst_base_row}:AG{end_r}",
        f"AH{dst_base_row}:AK{end_r}",
        f"B{dst_base_row + 9}:AK{dst_base_row + 9}",
    ]
    existing_merges = {str(rng) for rng in ws.merged_cells.ranges}
    for m_range in merge_ranges:
        if m_range not in existing_merges:
            ws.merge_cells(m_range)


def _shift_footer_down(ws: Any, first_row: int, shift: int) -> None:
    """
    Geser footer (mulai baris `first_row`) ke bawah sebanyak `shift` baris, termasuk merged cell,
    style, dan tinggi baris, supaya baris data baru tidak menabrak merged cell footer.
    """
    footer_merges = [
        (r.min_col, r.min_row, r.max_col, r.max_row)
        for r in list(ws.merged_cells.ranges)
        if r.min_row >= first_row
    ]
    for min_col, min_row, max_col, max_row in footer_merges:
        ws.unmerge_cells(start_row=min_row, start_column=min_col, end_row=max_row, end_column=max_col)

    last_row = ws.max_row
    max_col_idx = ws.max_column
    for r in range(last_row, first_row - 1, -1):
        ws.row_dimensions[r + shift].height = ws.row_dimensions[r].height
        for c in range(1, max_col_idx + 1):
            src = ws.cell(r, c)
            dst = ws.cell(r + shift, c)
            dst.value = src.value
            dst._style = copy(src._style)
            src.value = None
            src._style = StyleArray()

    for min_col, min_row, max_col, max_row in footer_merges:
        ws.merge_cells(
            start_row=min_row + shift,
            start_column=min_col,
            end_row=max_row + shift,
            end_column=max_col,
        )


def read_existing_dp_codes(report_bytes: bytes) -> list[str]:
    """Baca daftar kode DP yang sudah ada di sheet 'Captured Devices 01' dari report hasil export sebelumnya."""
    wb = openpyxl.load_workbook(BytesIO(report_bytes))
    if "Captured Devices 01" not in wb.sheetnames:
        return []
    return _existing_codes(wb["Captured Devices 01"])


def _existing_codes(ws_cap: Any) -> list[str]:
    codes: list[str] = []
    row = 35
    while True:
        no_val = ws_cap.cell(row, 2).value
        label = ws_cap.cell(row, 3).value
        if not isinstance(no_val, (int, float)) or not isinstance(label, str) or not label.strip():
            break
        codes.append(label.strip().removeprefix("DP ").strip())
        row += 10
    return codes


def build_workbook(
    template_src: Path | bytes,
    dp_results: dict[str, dict[str, Any]],
    base_report: bytes | None = None,
) -> openpyxl.Workbook | None:
    """
    Bangun objek Workbook openpyxl dari Path lokal atau raw bytes (st.secrets / upload UI).
    - Tanpa `base_report`: mulai dari template; baris 15 (No 1) & 25 (No 2) dikosongkan gambarnya,
      DP dimulai dari Baris 35 (No 3), 45 (No 4), 55 (No 5), dst.
    - Dengan `base_report` (xlsx hasil export sebelumnya): DP baru ditambahkan DI BAWAH DP yang sudah ada.
      DP yang kodenya sudah ada di report lama dilewati.
    Footer (Note/Authorization) di kedua sheet digeser otomatis sesuai jumlah DP.
    """
    if base_report is None and isinstance(template_src, Path) and not template_src.is_file():
        print(f"[WARN] File template Excel tidak ditemukan: {template_src}")
        return None

    is_base = base_report is not None
    if is_base:
        wb = openpyxl.load_workbook(BytesIO(base_report))
    else:
        logo_bytes = _extract_logo_bytes(template_src)
        wb_input = BytesIO(template_src) if isinstance(template_src, (bytes, bytearray)) else template_src
        wb = openpyxl.load_workbook(wb_input)

        for sname in wb.sheetnames:
            ws_any = wb[sname]
            b2_val = ws_any["B2"].value
            if isinstance(b2_val, str) and "DISPIMG" in b2_val:
                ws_any["B2"].value = None
                if logo_bytes:
                    logo_img = XLImage(BytesIO(logo_bytes))
                    logo_img.width = 140
                    logo_img.height = 65
                    ws_any.add_image(logo_img, "B2")

    if "Captured Devices 01" not in wb.sheetnames:
        print("[WARN] Sheet 'Captured Devices 01' tidak ditemukan.")
        return None

    ws_cap = wb["Captured Devices 01"]

    if is_base:
        existing_codes = _existing_codes(ws_cap)
    else:
        existing_codes = []
        kept_images = []
        for img in ws_cap._images:
            anc = img.anchor
            if hasattr(anc, "_from") and anc._from.row >= 12:
                continue
            kept_images.append(img)
        ws_cap._images = kept_images

        ws_cap.cell(15, 3).value = None
        ws_cap.cell(25, 3).value = None

        for old_base in (35, 45, 55):
            ws_cap.cell(old_base, 2).value = None
            ws_cap.cell(old_base, 3).value = None

    new_codes = [c for c in sorted(dp_results.keys()) if c not in set(existing_codes)]
    n_old = len(existing_codes)
    n_total = n_old + len(new_codes)

    extra_old = max(0, n_old - 3)
    extra_new = max(0, n_total - 3)
    if extra_new > extra_old:
        _shift_footer_down(ws_cap, 65 + extra_old * 10, (extra_new - extra_old) * 10)

    for offset, dp_code in enumerate(new_codes):
        idx = n_old + offset
        base_row = 35 + (idx * 10)
        if base_row > 55:
            _copy_dp_block_style(ws_cap, 35, base_row)

        ws_cap.cell(base_row, 2).value = idx + 3
        ws_cap.cell(base_row, 3).value = f"DP {dp_code}"

        photos_map: dict[str, Path | bytes] = dp_results[dp_code].get("photos", {})
        for category, (col_start, col_end) in CATEGORY_COL_MAP.items():
            img_src = photos_map.get(category)
            if not img_src:
                continue

            if isinstance(img_src, (bytes, bytearray)):
                xl_img = XLImage(BytesIO(img_src))
            elif isinstance(img_src, Path) and img_src.is_file():
                xl_img = XLImage(str(img_src))
            else:
                continue

            _from = AnchorMarker(
                col=col_start,
                colOff=30000,
                row=base_row - 1,
                rowOff=15000,
            )
            to = AnchorMarker(
                col=col_end,
                colOff=250000,
                row=base_row + 7,
                rowOff=150000,
            )
            xl_img.anchor = TwoCellAnchor(editAs="twoCell", _from=_from, to=to)
            ws_cap.add_image(xl_img)

    if "Form Measurement" in wb.sheetnames:
        ws_meas = wb["Form Measurement"]
        meas_first, meas_last_sample = 29, 34
        if not is_base:
            for clear_r in range(meas_first, meas_last_sample + 1):
                ws_meas.cell(clear_r, 11).value = None
                ws_meas.cell(clear_r, 14).value = None
                ws_meas.cell(clear_r, 15).value = None

        meas_extra_old = max(0, n_old - 6)
        meas_extra_new = max(0, n_total - 6)
        if meas_extra_new > meas_extra_old:
            _shift_footer_down(ws_meas, 35 + meas_extra_old, meas_extra_new - meas_extra_old)

        max_col_meas = ws_meas.max_column
        for offset, dp_code in enumerate(new_codes):
            idx = n_old + offset
            row_m = meas_first + idx
            if idx >= 6:
                for c in range(1, max_col_meas + 1):
                    ws_meas.cell(row_m, c)._style = copy(ws_meas.cell(meas_last_sample, c)._style)
                ws_meas.row_dimensions[row_m].height = ws_meas.row_dimensions[meas_last_sample].height
            ws_meas.cell(row_m, 2).value = idx + 1
            ws_meas.cell(row_m, 11).value = dp_code
            if in_dbm := dp_results[dp_code].get("input_dbm"):
                ws_meas.cell(row_m, 14).value = str(in_dbm).replace(".", ",")
            if out_dbm := dp_results[dp_code].get("output_dbm"):
                ws_meas.cell(row_m, 15).value = str(out_dbm).replace(".", ",")

    return wb


def export_to_excel(
    template_src: Path | bytes,
    output_excel_path: Path,
    dp_results: dict[str, dict[str, Any]],
) -> Path | None:
    """Simpan hasil workbook ke file .xlsx di disk."""
    wb = build_workbook(template_src, dp_results)
    if wb is None:
        return None
    output_excel_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_excel_path)
    return output_excel_path


def export_to_excel_bytes(
    template_src: Path | bytes,
    dp_results: dict[str, dict[str, Any]],
    base_report: bytes | None = None,
) -> bytes | None:
    """Kembalikan byte stream file .xlsx untuk di-download langsung lewat Web UI."""
    wb = build_workbook(template_src, dp_results, base_report)
    if wb is None:
        return None
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()
