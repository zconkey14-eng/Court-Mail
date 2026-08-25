@echo off
setlocal enabledelayedexpansion

rem Run this ONCE, on any Windows computer that already has a working
rem Python install (this machine's Python is only used to build the .exe --
rem the .exe itself does NOT need Python on the machines that run it
rem afterward). Produces CourtMailSorter.exe, a single self-contained file.
rem
rem NOTE: Tesseract OCR itself is a separate program, not a Python library,
rem so PyInstaller cannot bundle it. Every machine that runs the .exe still
rem needs Tesseract-OCR installed once:
rem https://github.com/UB-Mannheim/tesseract/wiki

cd /d "%~dp0"

set "PYEXE="

where py >nul 2>nul
if not errorlevel 1 set "PYEXE=py -3"

if not defined PYEXE (
    for /f "delims=" %%X in ('where python 2^>nul') do (
        echo %%X | find /I "WindowsApps" >nul
        if errorlevel 1 if not defined PYEXE set "PYEXE=%%X"
    )
)

if not defined PYEXE (
    echo Could not find a working Python installation to build with.
    echo Open "Anaconda Prompt" and run this file from there instead.
    pause
    exit /b 1
)

echo Using Python: !PYEXE!
echo Installing build requirements...
!PYEXE! -m pip install --upgrade pip
!PYEXE! -m pip install -r requirements.txt pyinstaller
if errorlevel 1 (
    echo Failed to install required packages.
    pause
    exit /b 1
)

echo.
echo Building CourtMailSorter.exe -- this can take a minute or two...
!PYEXE! -m PyInstaller --onefile --name CourtMailSorter --distpath . --workpath build --specpath build courtmail.py
if errorlevel 1 (
    echo Build failed.
    pause
    exit /b 1
)

echo.
echo Done. CourtMailSorter.exe is now in this folder.
echo Anyone can double-click it directly -- no Python required on their machine.
echo (Tesseract-OCR must still be installed once on each machine -- see the
echo note at the top of this file.)
pause
