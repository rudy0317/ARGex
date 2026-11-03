import argparse
import json
from pathlib import Path
from src.scanner import run_batch_scan

DEFAULT_CONFIG_PATH = Path(__file__).parent / "rules.json"


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Biznet Audit Photo Classifier & Auto-Renamer (Driven by rules.json)"
    )
    parser.add_argument(
        "--config",
        type=Path,
        default=DEFAULT_CONFIG_PATH,
        help="Path ke file konfigurasi JSON (default: rules.json)",
    )
    parser.add_argument(
        "--input",
        nargs="+",
        help="Override daftar folder input (contoh: --input './Folder Aset')",
    )
    parser.add_argument(
        "--target",
        nargs="+",
        help="Override target variabel yang dicari (contoh: --target DP)",
    )
    args = parser.parse_args()

    if not args.config.is_file():
        raise FileNotFoundError(f"File konfigurasi tidak ditemukan: {args.config}")

    with args.config.open("r", encoding="utf-8") as f:
        config = json.load(f)

    run_batch_scan(
        config=config,
        input_override=args.input,
        target_override=args.target,
    )


if __name__ == "__main__":
    main()