from pathlib import Path
import re
import shutil
from typing import Any
from src.exporter import export_to_excel

DEFAULT_NUMBER_MAP: dict[str, str] = {
    "1": "Output RK",
    "2": "Input DP",
    "3": "Output DP",
    "4": "Loop DP",
    "5": "Basetray DP",
    "6": "DP",
    "7": "Structure",
}


def extract_dp_code_from_path(
    filepath_parts: tuple[str, ...],
    filename: str,
    location_regex: str,
    blacklist: set[str],
) -> str:
    """Ekstrak kode lokasi DP (contoh: LTS010504) dari nama folder induk atau nama file."""
    pattern = re.compile(location_regex, re.IGNORECASE)
    for part in reversed(filepath_parts[:-1]):
        for prefix, digits in pattern.findall(part):
            if prefix.upper() not in blacklist and digits.isdigit():
                return f"{prefix.upper()}{digits}"

    for prefix, digits in pattern.findall(filename):
        if prefix.upper() not in blacklist and digits.isdigit():
            return f"{prefix.upper()}{digits}"

    # Fallback: gunakan nama folder induk langsung jika tidak cocok regex
    if len(filepath_parts) >= 2:
        parent_name = filepath_parts[-2].strip()
        if parent_name and parent_name.upper() not in blacklist:
            return parent_name
    return "UNKNOWN_DP"


def extract_number_slot(stem: str, valid_keys: set[str]) -> str | None:
    """
    Ambil nomor urut foto ('1' s/d '7') HANYA dari file yang sudah di-rename murni angka
    (contoh: '1.jpg', '2.jpg', ..., '7.jpg' atau '01.jpg'..'07.jpg').
    File sisa yang belum di-rename (seperti 'LINE_ALBUM_..._2.jpg') otomatis diabaikan.
    """
    s = stem.strip()
    if s.isdigit():
        norm = str(int(s))
        return norm if norm in valid_keys else None

    return None


def run_batch_scan(
    config: dict[str, Any],
    input_override: list[str] | None = None,
    target_override: list[str] | None = None,
) -> None:
    """
    Mode Numbered (1-7): Petakan foto 1 s/d 7 per folder DP langsung ke kolom Excel
    tanpa membutuhkan OpenCV / Tesseract-OCR.
    """
    paths_cfg = config["paths"]
    input_dirs = [Path(p) for p in (input_override or paths_cfg.get("input_dirs", []))]
    output_dir = Path(paths_cfg.get("output_dir", "./foto_renamed"))
    template_excel = Path(paths_cfg.get("template_excel", ""))
    output_excel = Path(paths_cfg.get("output_excel", "./Report/Report_FTTH_Auto.xlsx"))

    number_map: dict[str, str] = config.get("number_map", DEFAULT_NUMBER_MAP)
    valid_keys = set(number_map.keys())
    loc_regex = config.get(
        "location_regex",
        r"(?:^|[^a-zA-Z0-9])([a-zA-Z]{2,5})[\s\.\-~_|()\[\]]{0,6}([0-9]{4,8})(?=[^a-zA-Z0-9]|$)",
    )
    blacklist = {w.upper() for w in config.get("blacklist_prefixes", [])}

    output_dir.mkdir(parents=True, exist_ok=True)

    all_files: list[Path] = []
    for in_dir in input_dirs:
        if in_dir.is_dir():
            found = [
                f
                for f in in_dir.rglob("*")
                if f.is_file() and f.suffix.lower() in (".jpg", ".jpeg", ".png")
            ]
            all_files.extend(found)
        else:
            print(f"[WARN] Folder input tidak ditemukan: {in_dir}")

    print("Mode Eksekusi         : Manual Numbered (1-7 -> Excel)")
    print(
        "Mapping Nomor         : "
        + ", ".join(f"{k}={v}" for k, v in sorted(number_map.items(), key=lambda x: int(x[0])))
    )
    print(f"Total Foto Ditemukan  : {len(all_files)} file.\n" + "=" * 70)

    def _natural_sort_key(p: Path) -> tuple[str, int, str]:
        slot = extract_number_slot(p.stem, valid_keys)
        return (str(p.parent), int(slot) if slot else 999, p.name)

    matched_count = 0
    skipped_count = 0
    dp_results: dict[str, dict[str, Any]] = {}

    for filepath in sorted(all_files, key=_natural_sort_key):
        loc_code = extract_dp_code_from_path(filepath.parts, filepath.name, loc_regex, blacklist)
        slot_num = extract_number_slot(filepath.stem, valid_keys)

        if not slot_num:
            print(f"[SKIP] {loc_code}/{filepath.name} | Nama file bukan nomor 1-7")
            skipped_count += 1
            continue

        category = number_map[slot_num]
        if target_override and category not in target_override:
            skipped_count += 1
            continue

        dest_folder = output_dir / loc_code
        dest_folder.mkdir(parents=True, exist_ok=True)
        dest_file = dest_folder / f"{category}{filepath.suffix}"

        if loc_code not in dp_results:
            dp_results[loc_code] = {"photos": {}, "input_dbm": None, "output_dbm": None}

        if category in dp_results[loc_code]["photos"]:
            print(f"[DUP-SKIP] {loc_code}/{filepath.name} | Nomor {slot_num} ({category}) sudah ada")
            skipped_count += 1
            continue

        shutil.copy2(filepath, dest_file)
        dp_results[loc_code]["photos"][category] = dest_file
        matched_count += 1
        print(f"[SLOT #{slot_num} -> {category}] {loc_code}/{filepath.name} --> {loc_code}/{dest_file.name}")

    print("=" * 70)
    print(f"Selesai! Berhasil Dipetakan: {matched_count} | Di-skip (Di luar 1-7): {skipped_count}")

    if dp_results and template_excel.is_file():
        saved_excel = export_to_excel(template_excel, output_excel, dp_results)
        if saved_excel:
            print(f"[EXCEL] Laporan otomatis berhasil dibuat: {saved_excel}")
            print(
                f"        Total blok DP tertulis di 'Captured Devices 01': {len(dp_results)} DP ({', '.join(sorted(dp_results.keys()))})"
            )
