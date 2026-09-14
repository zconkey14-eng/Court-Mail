# Court Mail Sorter OCR - Packaging Guide

This turns `courtmail_ocr.py` into a standalone program that anyone on the
shared drive can run by double-clicking - no Python, no pip installs, no
Tesseract installer, nothing to set up on their own computer.

You do the build **once**, on any one Windows computer at work that has
Python. The output is a folder you copy to the shared drive; everyone else
just runs the `.exe` inside it.

## Why this is needed

- `pytesseract` (already in the script) is only a wrapper - it calls a real
  `tesseract.exe` program, which is a separate native install, not something
  pip can install.
- The script already knows how to find a **portable** copy of Tesseract: it
  looks for a `Tesseract-OCR` folder sitting right next to itself before
  checking anywhere else on the machine (see `BUNDLED_TESSERACT` in
  `courtmail_ocr.py`). So dropping a Tesseract-OCR folder next to the built
  exe is all that's needed - nobody has to install Tesseract themselves.
- Turning the script into a `.exe` (via PyInstaller) means coworkers don't
  need Python installed either.

## Step 1 - One-time build (do this on one Windows computer)

Requirements for this step only: Python 3.9+ installed on the build
computer, with internet access to download packages.

1. Copy this whole folder (`courtmail_ocr.py`, `requirements.txt`,
   `build.bat`) somewhere local (not the network drive - building on a
   local disk is much faster).
2. Double-click `build.bat`.
   - It creates a private `build_venv` folder so this doesn't touch any
     other Python setup on the computer.
   - It installs the required packages and runs PyInstaller.
   - When it finishes, your program is at
     `dist\CourtMailSorterOCR\CourtMailSorterOCR.exe`.

If `build.bat` says Python isn't found, install it from
[python.org/downloads](https://www.python.org/downloads/) - check
**"Add python.exe to PATH"** during install - then run `build.bat` again.

## Step 2 - Add a portable Tesseract-OCR folder

The exe alone isn't enough; it needs Tesseract's own program files sitting
right beside it in a folder named exactly `Tesseract-OCR`.

1. On the same build computer, install Tesseract normally, one time, from
   the Windows build here:
   https://github.com/UB-Mannheim/tesseract/wiki
   (Use the default install location, e.g.
   `C:\Program Files\Tesseract-OCR`.)
2. Copy that entire installed folder into your build output, so you have:

   ```
   dist\CourtMailSorterOCR\
       CourtMailSorterOCR.exe
       ...(other files PyInstaller put here)...
       Tesseract-OCR\
           tesseract.exe
           tessdata\
           ...
   ```

3. That's it - the `Tesseract-OCR` folder now travels with the program.
   You (or IT) never have to install Tesseract on any other computer again.

## Step 3 - Deploy to the shared drive

Copy the entire `dist\CourtMailSorterOCR` folder (exe + support files +
`Tesseract-OCR` folder, all together) to wherever on the shared drive people
should run it from. Keep everything in that one folder together - don't
separate the exe from `Tesseract-OCR` or the other files next to it.

## Step 4 - Using it (everyone else)

No install, no Python, nothing to set up. Just:

1. Open the shared folder.
2. Double-click `CourtMailSorterOCR.exe`.
3. It reads PDFs from `F:\Legal\MD\Court Mail\Input`, OCRs any pages that
   need it, sorts them, and writes results to
   `F:\Legal\MD\Court Mail\Output OCR Test` plus an Excel report.
4. Press Enter when it says "Press Enter to close" to close the window.

## Troubleshooting

- **"Tesseract is not installed on this machine"** - the `Tesseract-OCR`
  folder isn't sitting next to `CourtMailSorterOCR.exe`, or got renamed.
  Confirm the folder name is exactly `Tesseract-OCR` and it's a sibling of
  the exe, not nested inside another folder.
- **Antivirus flags or deletes the exe** - PyInstaller executables sometimes
  trigger false positives on locked-down work computers. If IT's antivirus
  quarantines it, you may need it whitelisted by path.
- **Different drive letters** - the script expects the shared drive mapped
  as `F:`. If any computer maps the same network share to a different
  letter, the Input/Output/xaa.csv paths at the top of `courtmail_ocr.py`
  will need updating (and a rebuild) to match, or IT should standardize the
  mapping to `F:` on every machine.
- **Rebuilding after a script change** - re-run `build.bat` (delete the old
  `build_venv`, `build`, and `dist` folders first if you want a totally
  clean build), then redo Step 2 and Step 3.
