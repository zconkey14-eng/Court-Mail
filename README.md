# Court Mail Sorter OCR - Packaging Guide

This turns `courtmail_ocr.py` into a standalone program that anyone on the
shared drive can run by double-clicking - no Python and no pip installs on
their own computer.

OCR is done by calling AWS Textract, so there's no local OCR engine to
install or bundle at all - every machine just needs the exe and network
access to AWS.

You do the build **once**, on any one Windows computer at work that has
Python. The output is a folder you copy to the shared drive; everyone else
just runs the `.exe` inside it.

**What lives where:** this repo holds the source script and the build
tooling only (`courtmail_ocr.py`, `build.bat`, `requirements.txt`). The
built `.exe` is build *output* - it belongs on the shared drive (Step 2),
never committed here.

## Step 1 - One-time build (do this on one Windows computer)

Requirements for this step only: Python 3.9+ installed on the build
computer, with internet access to download packages.

1. Copy this whole folder (`courtmail_ocr.py`, `requirements.txt`,
   `build.bat`) somewhere local (not the network drive - building on a
   local disk is much faster).
2. Double-click `build.bat`.
   - It creates a private `build_venv` folder so this doesn't touch any
     other Python setup on the computer.
   - It installs the required packages (including `boto3`, the AWS SDK)
     and runs PyInstaller.
   - When it finishes, your program is at
     `dist\CourtMailSorterOCR\CourtMailSorterOCR.exe`.

If `build.bat` says Python isn't found, install it from
[python.org/downloads](https://www.python.org/downloads/) - check
**"Add python.exe to PATH"** during install - then run `build.bat` again.

## Step 2 - Deploy to the shared drive

Copy the entire `dist\CourtMailSorterOCR` folder to wherever on the shared
drive people should run it from. Keep the exe together with the other files
PyInstaller put next to it in that folder.

## Step 3 - AWS credentials

Every machine needs to reach AWS Textract with a valid Access Key ID /
Secret Access Key that has `textract:DetectDocumentText` permission, and
needs the region set in `TEXTRACT_REGION` at the top of `courtmail_ocr.py`
to be one where Textract is available.

There are two ways to supply the keys - the script tries them in this
order:

1. **Already set on the machine** - if `AWS_ACCESS_KEY_ID` and
   `AWS_SECRET_ACCESS_KEY` are set as environment variables (or an AWS CLI
   profile is configured), the program uses them automatically and asks for
   nothing.
2. **Typed in at runtime** - if neither is found, the program prompts for
   the Access Key ID and Secret Access Key when it starts. They are held in
   memory for that run only and are never written to disk, so this is safe
   to do on a shared machine, but it does mean typing them in every time the
   program is run unless step 1 is set up instead.

## Step 4 - Using it (everyone else)

No install, no Python, nothing to set up beyond the AWS keys above. Just:

1. Open the shared folder.
2. Double-click `CourtMailSorterOCR.exe`.
3. If asked, enter the AWS Access Key ID and Secret Access Key.
4. It reads PDFs from `F:\Legal\MD\Court Mail\Input`, OCRs any pages that
   need it via Textract, sorts them, and writes results to
   `F:\Legal\MD\Court Mail\Output` plus an Excel report.
5. Press Enter when it says "Press Enter to close" to close the window.

## Cost and rate limits

Textract's `DetectDocumentText` call is billed per page - check current
pricing at https://aws.amazon.com/textract/pricing/ before running large
batches. Pages are sent `TEXTRACT_WORKERS` at a time (8 by default, set at
the top of `courtmail_ocr.py`); this is deliberately conservative against
AWS's default per-account request-rate quota for Textract, not the number
of cores on the machine. If large batches are hitting throttling errors,
request a Service Quota increase for Textract in the AWS account rather
than just raising this number.

**Spend cap:** `SPEND_LIMIT_USD` at the top of `courtmail_ocr.py` (default
`$30.00` per run) stops the program from sending any more pages to Textract
once estimated spend would cross it - the rest of that run's pages are
treated like any other unreadable page and land in Review instead. This is
a client-side estimate against the `TEXTRACT_PRICE_PER_PAGE_USD` constant
next to it, **not** a real AWS-enforced billing limit - it only stops this
program, and the estimate drifts if AWS pricing changes. For a true
account-wide hard limit, pair this with an AWS Budget alert (or a budget
action) on the account itself. The end-of-run summary in the console shows
how many pages were actually sent and the estimated cost.

## Troubleshooting

- **Every page fails with a credentials/signature error** - the Access Key
  ID or Secret Access Key entered was wrong, or the IAM user behind them
  doesn't have Textract permission. Run the program again to re-enter the
  keys, or fix the IAM permissions.
- **Antivirus flags or deletes the exe** - PyInstaller executables sometimes
  trigger false positives on locked-down work computers. If IT's antivirus
  quarantines it, you may need it whitelisted by path.
- **No internet / AWS unreachable** - Textract calls need outbound HTTPS
  access to AWS from the machine; a locked-down corporate firewall may need
  an allowance for the Textract endpoint in the configured region.
- **Different drive letters** - the script expects the shared drive mapped
  as `F:`. If any computer maps the same network share to a different
  letter, the Input/Output/xaa.csv paths at the top of `courtmail_ocr.py`
  will need updating (and a rebuild) to match, or IT should standardize the
  mapping to `F:` on every machine.
- **Rebuilding after a script change** - re-run `build.bat` (delete the old
  `build_venv`, `build`, and `dist` folders first if you want a totally
  clean build), then redo Step 2.
