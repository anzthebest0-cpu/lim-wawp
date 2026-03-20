@echo off
setlocal
cd /d "%USERPROFILE%\Downloads"

if not exist "generate_laporan.py" (
    echo File generate_laporan.py tidak ada di Downloads.
    pause & exit /b
)

echo Mencari file data LIM terbaru...
for /f "delims=" %%f in ('dir /b /o-d LIM_data_*.json LIM_WAWP_*.json 2^>nul') do (
    set "LATEST=%%f"
    goto :found
)

:found
if not defined LATEST (
    echo Tidak ada file JSON ditemukan. Ekspor data dari aplikasi dulu.
    pause & exit /b
)

echo Membuat PDF dari: %LATEST%
python generate_laporan.py "%LATEST%"
if %errorlevel%==0 (
    echo PDF berhasil dibuat dan dibuka otomatis.
) else (
    echo Gagal. Pastikan: pip install reportlab Pillow
)
pause
