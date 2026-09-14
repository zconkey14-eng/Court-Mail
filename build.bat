@echo off
REM ============================================================
REM  Court Mail Sorter OCR - one-time build script
REM
REM  Run this ONCE, on any Windows computer that has Python
REM  installed. It produces a standalone folder containing
REM  CourtMailSorterOCR.exe that nobody else needs Python for -
REM  they just double-click the exe from the shared drive.
REM
REM  If this computer does not have Python, install it first
REM  from https://www.python.org/downloads/ (check "Add
REM  python.exe to PATH" during install), then run this again.
REM ============================================================

setlocal

cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo.
    echo Python was not found on this computer.
    echo Install it from https://www.python.org/downloads/
    echo During install, check the box "Add python.exe to PATH".
    echo Then run build.bat again.
    echo.
    pause
    exit /b 1
)

echo.
echo Creating a private build environment (build_venv)...
python -m venv build_venv
if errorlevel 1 goto :error

call build_venv\Scripts\activate.bat

echo.
echo Installing required packages...
python -m pip install --upgrade pip
pip install -r requirements.txt
if errorlevel 1 goto :error

echo.
echo Building CourtMailSorterOCR.exe ...
echo (This can take a few minutes the first time.)
pyinstaller --noconfirm --onedir --console ^
    --name CourtMailSorterOCR ^
    --collect-all pymupdf ^
    --collect-all pytesseract ^
    courtmail_ocr.py
if errorlevel 1 goto :error

echo.
echo ============================================================
echo  Build finished.
echo.
echo  Your program is here:
echo    dist\CourtMailSorterOCR\CourtMailSorterOCR.exe
echo.
echo  NEXT STEP - before anyone can run it, you must add a
echo  portable copy of Tesseract-OCR next to the exe. See
echo  README.md, section "Adding Tesseract-OCR", for how.
echo.
echo  Once that folder is added, copy the ENTIRE
echo  dist\CourtMailSorterOCR folder to the shared drive.
echo  Everyone runs CourtMailSorterOCR.exe from there directly -
echo  no install needed on their machines.
echo ============================================================
echo.
pause
exit /b 0

:error
echo.
echo Build failed - see the messages above.
echo.
pause
exit /b 1
