"""Second look at the Trash folder: flag pages that may have been trashed by
mistake.

Re-reads every trashed page saved in the last N days and runs it back
through the CURRENT classification rules in courtmail_aws.py. A page is
flagged when:

  * the current rules would no longer trash it (a rule was changed since the
    run - e.g. a new Owen exception), or
  * it also contains a keep keyword. Trash is checked first, so a page that
    is a real Hearing Notice but also says "has been canceled" gets trashed;
    this is the main way a keeper ends up in Trash.

Pages with no text layer (scans that were OCR'd by Textract during the run
and saved as images) can't be re-read for free. They are counted, and only
re-OCR'd through Textract if OCR_IMAGE_ONLY_PAGES is True.

Needs to sit in the same folder as courtmail_aws.py.
"""

import time
from datetime import datetime
from pathlib import Path

import pymupdf
from openpyxl import Workbook

import courtmail_aws as sorter

# Trash sub-folders to skip. Duplicates are copies of pages that were kept.
SKIP_FOLDERS = {sorter.DUPLICATES_FOLDER}

# Re-OCR image-only trashed pages through Textract (costs money, capped by
# the same spend limit as the main program). Off by default.
OCR_IMAGE_ONLY_PAGES = False

DEFAULT_DAYS_BACK = 7


def ask_days_back():
    answer = input(
        f"How many days back to check? [{DEFAULT_DAYS_BACK}]: "
    ).strip()

    if not answer:
        return DEFAULT_DAYS_BACK

    try:
        return max(1, int(answer))
    except ValueError:
        print(f"Not a number - using {DEFAULT_DAYS_BACK}.")
        return DEFAULT_DAYS_BACK


def keep_keyword_hits(text):
    """Return [(doc_type, keyword)] for every keep keyword on the page."""

    text_lower = text.lower()
    text_glued = sorter.WHITESPACE.sub('', text_lower)
    hits = []

    for doc_type, keywords in sorter.DOCUMENT_TYPES.items():

        for keyword in keywords:

            if sorter.keyword_matches(keyword, text_lower, text_glued):
                hits.append((doc_type, keyword))

    return hits


def read_text_layer(pdf_path):
    """Return (cleaned text, error). Text is "" for an image-only page."""

    try:

        with pymupdf.open(pdf_path) as document:

            raw = "".join(page.get_text("text") for page in document)

    except Exception as error:  # pylint: disable=broad-exception-caught
        return "", str(error)

    return sorter.clean_text(raw), None


def ocr_pdf_text(pdf_path, client, spend_guard):
    """Textract a single-page saved PDF. Returns cleaned text."""

    with pymupdf.open(pdf_path) as document:
        image = sorter.render_page_image(document, 0)

    _, text = sorter.textract_page(client, 0, image, spend_guard)

    return text


def audit_page(pdf_path, trash_type, text):
    """Return a list of reasons this trashed page looks wrong (may be empty)."""

    reasons = []

    doc_type, _, is_trash = sorter.classify_page(text)

    if not is_trash:
        reasons.append(
            f"Current rules would now file it as: {doc_type}"
        )

    hits = keep_keyword_hits(text)

    if hits:
        found = "; ".join(f"{name} ({keyword})" for name, keyword in hits)
        reasons.append(f"Also contains keep keyword(s): {found}")

    return reasons


def snippet(text, width=160):
    return text[:width]


def write_report(flagged, path):
    workbook = Workbook()
    sheet = workbook.worksheets[0]
    sheet.title = "Possibly Wrongly Trashed"

    sheet.append([
        "Saved File", "Trash Folder", "Reason", "Saved Folder", "Text Start"
    ])

    for item in flagged:
        sheet.append([
            item["file"], item["trash_type"], item["reason"],
            item["folder"], item["text"]
        ])

    for column in sheet.columns:
        longest = max(len(str(c.value or "")) for c in column)
        sheet.column_dimensions[column[0].column_letter].width = min(
            longest + 2, 70
        )

    workbook.save(path)


def main():
    print()
    print("Court Mail - Trash Audit")
    print("=" * 60)

    trash_root = sorter.OUTPUT_FOLDER / "Trash"

    if not trash_root.exists():
        print(f"No Trash folder at {trash_root}")
        return

    days = ask_days_back()
    cutoff = time.time() - days * 86400

    pdfs = [
        path
        for path in trash_root.glob("*/*.pdf")
        if path.parent.name not in SKIP_FOLDERS
        and path.stat().st_mtime >= cutoff
    ]

    print(f"Checking {len(pdfs):,} trashed page(s) from the last {days} day(s).")
    print()

    client = None
    spend_guard = None

    if OCR_IMAGE_ONLY_PAGES:
        client = sorter.get_textract_client()
        spend_guard = sorter.SpendGuard(
            sorter.SPEND_LIMIT_USD, sorter.TEXTRACT_PRICE_PER_PAGE_USD
        )

    flagged = []
    image_only = 0
    unreadable = 0

    for index, pdf_path in enumerate(pdfs, start=1):

        if index % 200 == 0:
            print(f"  {index:,}/{len(pdfs):,}", end="\r", flush=True)

        trash_type = pdf_path.parent.name
        text, error = read_text_layer(pdf_path)

        if error:
            unreadable += 1
            continue

        if len(text) < sorter.MIN_TEXT_CHARS:

            image_only += 1

            if not OCR_IMAGE_ONLY_PAGES:
                continue

            text = ocr_pdf_text(pdf_path, client, spend_guard)

            if not text:
                continue

        reasons = audit_page(pdf_path, trash_type, text)

        if reasons:
            flagged.append({
                "file": pdf_path.name,
                "trash_type": trash_type,
                "reason": " | ".join(reasons),
                "folder": str(pdf_path.parent),
                "text": snippet(text),
            })

    print()
    print("=" * 60)
    print(f"Checked:            {len(pdfs):,}")
    print(f"Flagged:            {len(flagged):,}")
    print(f"Image-only pages:   {image_only:,}"
          + ("" if OCR_IMAGE_ONLY_PAGES else "  (not re-read - see OCR_IMAGE_ONLY_PAGES)"))
    print(f"Unreadable files:   {unreadable:,}")

    if spend_guard:
        print(f"Textract spend (est.): ~${spend_guard.spent_usd:.2f}")

    if not flagged:
        print()
        print("Nothing looks wrongly trashed.")
        return

    report_path = sorter.get_unique_filename(
        sorter.OUTPUT_FOLDER,
        f"trash_audit_{datetime.now():%Y-%m-%d}.xlsx"
    )

    write_report(flagged, report_path)

    print()
    print(f"Flagged pages saved to: {report_path}")


if __name__ == "__main__":

    try:
        main()
    except Exception:  # pylint: disable=broad-exception-caught
        import traceback
        traceback.print_exc()
    finally:
        input("\nPress Enter to close...")
