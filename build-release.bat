@echo off
setlocal

cd /d "%~dp0"

echo [1/4] Building application with PyInstaller...
python -m PyInstaller "SIMAK-TPP.spec" --noconfirm --clean
if errorlevel 1 (
    echo Build gagal.
    exit /b 1
)

echo [2/4] Menyiapkan folder files...
if exist "dist\SIMAK-TPP\files" rmdir /s /q "dist\SIMAK-TPP\files"
xcopy "files" "dist\SIMAK-TPP\files\" /e /i /y >nul
if errorlevel 1 (
    echo Gagal menyalin folder files.
    exit /b 1
)

echo [3/4] Menyiapkan browser Playwright Chromium...
set "PW_SRC=%LOCALAPPDATA%\ms-playwright\chromium-1243"
set "PW_DST=dist\SIMAK-TPP\_internal\playwright\driver\package\.local-browsers\chromium-1243"
if exist "%PW_SRC%" (
    if not exist "%PW_DST%" (
        echo Menyalin Chromium dari cache lokal ke build...
        xcopy "%PW_SRC%" "%PW_DST%\" /e /i /y >nul
        echo Browser Chromium berhasil disertakan ke paket build.
    ) else (
        echo Browser Chromium sudah ada di folder build.
    )
) else (
    echo Browser Chromium lokal tidak ditemukan, aplikasi akan menggunakan Chrome/Edge sistem.
)

echo [4/4] Selesai!
echo Hasil build: "%cd%\dist\SIMAK-TPP"
exit /b 0
