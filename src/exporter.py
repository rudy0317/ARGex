from copy import copy
from io import BytesIO
from pathlib import Path
from typing import Any
import zipfile
import openpyxl
from openpyxl.drawing.image import Image as XLImage
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


def _extract_logo_bytes(template_path: Path) -> bytes | None:
    """Ambil file logo Biznet dari dalam archive template xlsx agar tidak error #NAME? akibat rumus WPS DISPIMG."""
    try:
        with zipfile.ZipFile(template_path, "r") as z:
            for candidate in ("xl/media/image1.png", "xl/media/image25.png"):
                if candidate in z.namelist():
                    return z.read(candidate)
    except Exception:
        return None
    return None


def _copy_dp_block_style(ws: Any, src_base_row: int, dst_base_row: int) -> None:
    """Salin struktur merge cell, border, font, dan label dari blok DP template ke baris baru."""
    for r_offset in range(10):
        src_r = src_base_row + r_offset
        dst_r = dst_base_row + r_offset
        ws.row_dimensions[dst_r].height = ws.row_dimensions[src_r].height
        for col_idx in range(2, 38):
            src_cell = ws.cell(src_r, col_idx)
            dst_cell = ws.cell(dst_r, col_idx)
            if r_offset == 0:
                dst_cell.value = src_cell.value
            if src_cell.has_style:
                dst_cell.font = copy(src_cell.font)
                dst_cell.fill = copy(src_cell.fill)
                dst_cell.border = copy(src_cell.border)
                dst_cell.alignment = copy(src_cell.alignment)

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
    ]
    existing_merges = {str(rng) for rng in ws.merged_cells.ranges}
    for m_range in merge_ranges:
        if m_range not in existing_merges:
            ws.merge_cells(m_range)


def export_to_excel(
    template_path: Path,
    output_excel_path: Path,
    dp_results: dict[str, dict[str, Any]],
) -> Path | None:
    """
    Masukkan hasil foto yang sudah diklasifikasi ke sheet 'Captured Devices 01'
    dan nilai redaman dBm ke sheet 'Form Measurement'.
    """
    if not template_path.is_file():
        print(f"[WARN] File template Excel tidak ditemukan: {template_path}")
        return None

    logo_bytes = _extract_logo_bytes(template_path)
    wb = openpyxl.load_workbook(template_path)

    # Ganti formula proprietary WPS =DISPIMG(...) di sel B2 dengan gambar logo standar Excel
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
        print("[WARN] Sheet 'Captured Devices 01' tidak ditemukan di template.")
        return None

    ws_cap = wb["Captured Devices 01"]

    # Hapus gambar sampel DP lama di baris >= 34 (0-indexed >= 33), pertahankan gambar OLT/RK di atasnya
    kept_images = []
    for img in ws_cap._images:
        anc = img.anchor
        if hasattr(anc, "_from") and anc._from.row >= 33:
            continue
        kept_images.append(img)
    ws_cap._images = kept_images

    # Bersihkan label DP sampel lama pada baris 35, 45, 55
    for old_base in (35, 45, 55):
        ws_cap.cell(old_base, 2).value = None
        ws_cap.cell(old_base, 3).value = None

    sorted_dp_codes = sorted(dp_results.keys())

    # Isi tiap blok DP mulai dari baris 35 (interval 10 baris: 35, 45, 55, ...)
    for idx, dp_code in enumerate(sorted_dp_codes):
        base_row = 35 + (idx * 10)
        if base_row > 55:
            _copy_dp_block_style(ws_cap, 35, base_row)

        ws_cap.cell(base_row, 2).value = idx + 3  # Nomor urut setelah OLT(1) dan RK(2)
        ws_cap.cell(base_row, 3).value = f"DP {dp_code}"

        photos_map: dict[str, Path] = dp_results[dp_code].get("photos", {})
        for category, (col_start, col_end) in CATEGORY_COL_MAP.items():
            img_path = photos_map.get(category)
            if not img_path or not img_path.is_file():
                continue

            xl_img = XLImage(str(img_path))
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

    # Sinkronkan juga kode DP & hasil bacaan OPM dBm ke sheet 'Form Measurement'
    if "Form Measurement" in wb.sheetnames:
        ws_meas = wb["Form Measurement"]
        # Bersihkan nilai sampel lama pada baris 29..34 untuk kolom DP Code (K), Input dBm (N), Output dBm (O)
        for clear_r in range(29, 35):
            ws_meas.cell(clear_r, 11).value = None
            ws_meas.cell(clear_r, 14).value = None
            ws_meas.cell(clear_r, 15).value = None

        for idx, dp_code in enumerate(sorted_dp_codes):
            row_m = 29 + idx
            ws_meas.cell(row_m, 2).value = idx + 1
            ws_meas.cell(row_m, 11).value = dp_code  # Col K: DP Code
            if in_dbm := dp_results[dp_code].get("input_dbm"):
                ws_meas.cell(row_m, 14).value = in_dbm.replace(".", ",")  # Col N: Input DP (dBm)
            if out_dbm := dp_results[dp_code].get("output_dbm"):
                ws_meas.cell(row_m, 15).value = out_dbm.replace(".", ",")  # Col O: Output DP (dBm)

    output_excel_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_excel_path)
    return output_excel_path
