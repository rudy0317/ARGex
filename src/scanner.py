from pathlib import Path
import re
import shutil
from typing import Any
import cv2
import pytesseract
from src.classifier import AuditClassifier
from src.exporter import export_to_excel


def setup_tesseract(cmd_path: str) -> None:
    """Set path executable Tesseract dari config atau fallback ke system PATH."""
    p = Path(cmd_path)
    if p.is_file():
        pytesseract.pytesseract.tesseract_cmd = str(p)
    elif fallback := shutil.which("tesseract"):
        pytesseract.pytesseract.tesseract_cmd = fallback


def run_batch_scan(
    config: dict[str, Any],
    input_override: list[str] | None = None,
    target_override: list[str] | None = None,
) -> None:
    """Scan semua folder input, klasifikasikan gambar, simpan per folder lokasi, dan export ke Excel."""
    paths_cfg = config["paths"]
    setup_tesseract(paths_cfg.get("tesseract_cmd", ""))

    input_dirs = [Path(p) for p in (input_override or paths_cfg.get("input_dirs", []))]
    output_dir = Path(paths_cfg.get("output_dir", "./foto_renamed"))
    template_excel = Path(paths_cfg.get("template_excel", ""))
    output_excel = Path(paths_cfg.get("output_excel", "./Report/Report_FTTH_Auto.xlsx"))
    active_targets = target_override or config.get("active_targets", ["DP"])

    classifier = AuditClassifier(config)
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

    print(f"Target Variabel Aktif : {', '.join(active_targets)}")
    print(f"Total Foto Ditemukan  : {len(all_files)} file.\n" + "=" * 70)

    matched_count = 0
    skipped_count = 0
    dp_results: dict[str, dict[str, Any]] = {}

    def _natural_sort_key(p: Path) -> tuple[str, int, str]:
        m = re.search(r"_(\d+)$", p.stem)
        return (str(p.parent), int(m.group(1)) if m else 0, p.name)

    for filepath in sorted(all_files, key=_natural_sort_key):
        img = cv2.imread(str(filepath))
        if img is None:
            print(f"[SKIP] Gagal membaca file: {filepath.name}")
            skipped_count += 1
            continue

        category, detail = classifier.classify(img, active_targets)
        loc_code = classifier.extract_folder_code(img, filepath.parts, filepath.name)

        if category:
            dest_folder = output_dir / loc_code
            dest_folder.mkdir(parents=True, exist_ok=True)

            ext = filepath.suffix
            dest_file = dest_folder / f"{category}{ext}"

            if loc_code not in dp_results:
                dp_results[loc_code] = {"photos": {}, "input_dbm": None, "output_dbm": None}

            if category in dp_results[loc_code]["photos"]:
                print(
                    f"[DUP-SKIP] {loc_code}/{filepath.name} | {dest_file.name} sudah terisi di sesi ini"
                )
                skipped_count += 1
                continue

            shutil.copy2(filepath, dest_file)
            print(
                f"[MATCH: {category}] {loc_code}/{filepath.name} --> {loc_code}/{dest_file.name} | {detail}"
            )
            matched_count += 1

            dp_results[loc_code]["photos"][category] = dest_file

            if dbm_match := re.search(r"-\d+\.\d+", detail):
                if category == "Input DP":
                    dp_results[loc_code]["input_dbm"] = dbm_match.group(0)
                elif category == "Output DP":
                    dp_results[loc_code]["output_dbm"] = dbm_match.group(0)
        else:
            print(f"[SKIP] {loc_code}/{filepath.name} | {detail}")
            skipped_count += 1

    print("=" * 70)
    print(f"Selesai! Berhasil Di-rename: {matched_count} | Di-skip (Bukan Target): {skipped_count}")

    if dp_results and template_excel.is_file():
        saved_excel = export_to_excel(template_excel, output_excel, dp_results)
        if saved_excel:
            print(f"[EXCEL] Laporan otomatis berhasil dibuat: {saved_excel}")
            print(f"        Total blok DP tertulis di 'Captured Devices 01': {len(dp_results)} DP ({', '.join(sorted(dp_results.keys()))})")
