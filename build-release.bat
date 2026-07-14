@echo off
setlocal

cd /d "%~dp0"

echo [1/3] Building application...
python -m PyInstaller "SIMAK-TPP.spec" --noconfirm --clean
if errorlevel 1 (
    echo Build gagal.
    exit /b 1
)

echo [2/3] Menyiapkan folder files yang bisa diedit...
if exist "dist\SIMAK-TPP\files" rmdir /s /q "dist\SIMAK-TPP\files"
xcopy "files" "dist\SIMAK-TPP\files\" /e /i /y >nul
if errorlevel 1 (
    echo Gagal menyalin folder files.
    exit /b 1
)

echo [3/3] Selesai.
echo Hasil build: "%cd%\dist\SIMAK-TPP"
exit /b 0
