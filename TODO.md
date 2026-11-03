# ARGex — FTTH Network Assurance Audit Classifier & Auto-Exporter

Roadmap pengembangan dan daftar tugas terstruktur untuk **ARGex**.

---

## Phase 1: Core Vision & OCR Engine (Completed)
- [x] Modularisasi arsitektur (`main.py`, `rules.json`, `src/classifier.py`, `src/scanner.py`, `src/exporter.py`).
- [x] Ekstraksi kode lokasi DP (`LTSxxxxxx`) via regex & normalisasi karakter sablon (`stencil_char_map`).
- [x] Detektor **DP** (HSV black-box gate + white-stencil gate + multi-threshold Tesseract OCR).
- [x] Detektor **OPM (`Input DP` & `Output DP`)** dengan dekoder geometris 7-segment & aturan SOP (`< 10.0 dBm` Input, `>= 10.0 dBm` Output).
- [x] Detektor **Loop DP**, **Basetray DP**, dan **Structure** berbasis komposisi spasial HSV & Sobel vertical edge.
- [x] Preservasi metadata EXIF asli menggunakan `shutil.copy2` (tanpa re-kompresi JPEG).

## Phase 2: Auto-Exporter Excel Report (Completed)
- [x] Injeksi foto otomatis ke sheet `Captured Devices 01` menggunakan `TwoCellAnchor` sesuai kolom variabel.
- [x] Auto-clone blok 20 baris untuk DP ke-2 dan seterusnya pada `Captured Devices 01`.
- [x] Auto-fill nilai redaman `Input DP` dan `Output DP` (dBm) ke sheet `Form Measurement`.
- [x] Perbaikan formula gambar proprietary WPS Office (`=DISPIMG(...)`) di header `B2` dengan injeksi logo PNG asli.

## Phase 3: Multi-Dataset Calibration (Completed)
- [x] Kalibrasi dual-mode layar LCD OPM: **Backlit LCD** (`LTS010603`) dan **Unlit STN Olive-Green LCD** (`LTS010504`).
- [x] Pemangkasan otomatis ekor patchcord kuning tipis di atas bodi OPM agar bounding box LCD presisi.
- [x] Hard exclusion bodi OPM agar foto OPM tidak bocor (fall-through) ke `Loop DP` atau `Basetray DP`.
- [x] Disambiguasi `Basetray DP` vs `Loop DP` menggunakan rasio buffer tube oranye di sisi kiri (`max_left_orange_ratio`).
- [x] Natural sorting urutan nama file (`_1` s/d `_10`) dan deduplikasi per sesi.

---

## Phase 4: Local Desktop Utility UI (Next Up)
- [ ] Bangun UI desktop ringan berbasis `tkinter` / `ttk` (gaya *industrial/engineering utility*, zero AI-slop).
- [ ] Fitur pilih folder input batch (drag/browse multi-folder DP) & pilih path output Excel.
- [ ] Checklist toggle variabel target aktif (`DP`, `Input DP`, `Output DP`, `Loop DP`, `Basetray DP`, `Structure`, `Output RK`).
- [ ] Tabel preview hasil klasifikasi + nilai dBm + fitur **Manual Override** sebelum tombol *Export Excel* ditekan (jaga-jaga jika ada foto lapangan ekstrem).
- [ ] Background worker thread agar UI tidak freeze saat memproses puluhan foto.

## Phase 5: Portable Packaging untuk Laptop Kantor (Upcoming)
- [ ] Update `setup_tesseract()` di `src/scanner.py` agar memprioritaskan binary relatif `./tesseract/tesseract.exe` sebelum `C:\Program Files\Tesseract-OCR\tesseract.exe`.
- [ ] Buat script build `PyInstaller` mode `--onedir` (`build_portable.ps1`) yang membungkus Python + OpenCV + `openpyxl` + folder `tesseract` portable.
- [ ] Uji coba jalan tanpa instalasi / tanpa akses Admin di Windows.

## Phase 6: Ekspansi & CI/CD (Backlog)
- [ ] Kalibrasi detektor variabel ke-7: **`Output RK`** setelah sampel foto lapangan tersedia.
- [ ] Tambahkan unit test regresi untuk parser kode lokasi & dekoder 7-segment (`tests/`).
- [ ] Setup GitHub Actions workflow (`.github/workflows/ci.yml`) untuk linting (`ruff`) dan unit test otomatis setiap push.
