@echo off
setlocal
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Python belum terinstall. Install Python 3.12 dari https://www.python.org/downloads/
    echo         Centang "Add python.exe to PATH" saat install, lalu jalankan file ini lagi.
    pause
    exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
    echo [SETUP] Membuat virtual environment...
    python -m venv .venv
    if errorlevel 1 goto :fail
)

if not exist ".venv\.deps_ok" (
    echo [SETUP] Menginstall library, tunggu sebentar...
    ".venv\Scripts\python.exe" -m pip install --upgrade pip
    ".venv\Scripts\python.exe" -m pip install -r requirements.txt
    if errorlevel 1 goto :fail
    echo ok> ".venv\.deps_ok"
)

echo [RUN] Membuka ARGex di browser (http://localhost:8501). Tutup jendela ini untuk menghentikan.
".venv\Scripts\python.exe" -m streamlit run app.py --server.port 8501 --browser.gatherUsageStats false
goto :eof

:fail
echo [ERROR] Setup gagal. Cek koneksi internet lalu jalankan lagi.
pause
exit /b 1
