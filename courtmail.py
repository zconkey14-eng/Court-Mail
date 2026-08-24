import csv
import re
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from pypdf import PdfReader, PdfWriter

# Folder Settings

INPUT_FOLDER = Path(r"F:\Legal\MD\Court Mail\Input")
OUTPUT_FOLDER = Path(r"F:\Legal\MD\Court Mail\Output")

# Case Number -> File Number lookup source
# Columns used: "case" (court case number) and "file" (our file / account number)

XAA_CSV = Path(r"F:\Reports\X-Reports\xaa.csv")

# Used when a case number has no match in xaa.csv
MANUAL = "MANUAL"

# Used for trash pages - we never look up a file number for those
NOT_APPLICABLE = "N/A"

# Second and later copies of the same document go here instead of
# cluttering the keep folders
DUPLICATES_FOLDER = "Duplicates"


# Document Types

DOCUMENT_TYPES = {
    "Amended Judgment": [
        "DC-BNV-Q2",
        "DC-BNV-Q3P"
    ],

    "Renewal of Judgment": [
        "renewed on"
    ],

    "Vacated or Stricken": [
        "DC-BNV-I1",
        "vacated or stricken",
        "you will be notifid of any further proceedings."
    ],

    "Hearing Notice": [
        "Type of Proceeding: Hearing - Motion",
        "Type of Proceeding: hearing - Show Cause",
        # OCR turns the dash into anything: "Hearing .., Show Cause ."
        # Safe as a bare phrase because trash is checked first
        "Show Cause"
    ],
    "Trial Notice":[
        "Type of Proceeding: Trial"
    ],

    "Affidavit Judgment Denial": [
        "DC-BNV-R2",
        "the affidavit defective."
    ],

    "Denied Motion": [
        "DC-BNV-G1B",
        "Maryland denied the motion"
    ],

    "Notice of Intention to Defend": [
        "DC-BNV-B1",
        "Intention to Defend.",
        "Defendant's Response"
    ],

    "Lien of Judgment": [
        "DC-BNV-S3",
        "Lien of Judgment Recorded in"
    ]
}


# The DC-BNV-* form codes are the primary signal, but they sit in small print
# at the bottom of the page and OCR loses them often - 51 of the 52 pages in
# the 8-18-26 Review folder were notices we already have a category for, with
# a destroyed form code. The plain-English phrase from the body of the notice
# is the fallback, and it survives scanning far better.

TRASH_DOCUMENTS = {
    # Criminal mail for private individuals that gets mixed into the batch
    # (expungement notices, summonses, preliminary inquiries). Checked first so
    # it never gets filed as one of the firm's civil notices.
    "Criminal Court Notices": [
        "STATE OF MARYLAND VS"
    ],

    "Affidavit Date Notices": [
        "(Affidavit Judgment)"
    ],

    "Satisfaction updates": [
        "DC-BNV-SE",
        "was entered as satisfied"
    ],

    "Conference Resolution": [
        "Conference - Resolution",
        # On 8-13-26 Prince George's scans, OCR read "Conference" as
        # "Confere1ice". It is still the same non-actionable conference.
        "Confere1ice - Resolution"
    ],

    # Case moved to another county, which issues it a brand new case number
    "Case Transfer": [
        "DC-BNV-W1",
        "a new case number has been assigned"
    ],

    "Tentative hearings": [
        "DC-BNV-D1",
        "DC-BNV-P1",
        "Hearing - Tentative",
        "Pending Service on Defendant"
    ],

    "Judgment": [
        "DC-BNV-J1",
        "DC-BNV-R1",
        "Post-judgment interest will be assessed at the legal rate."
    ],

    "Granted Motion": [
        "DC-BNV-G1A",
        "DC-BNV-G1C",
        "granted the motion"
    ],

    # The court looked at the motion and did nothing yet
    "Motion Noted No Action": [
        "no action has been taken at this time"
    ],

    # The court signed the motion but it arrived too late to matter
    "Moot Motion": [
        "but took no action",
        "motion is moot"
    ],

    "Garnishment Documents": [
        "DC-BNV-U1",
        "DC-BNV-U4",
        "DC-BNV-E6",
        "answer to the writ of garnishment",
        # Post-judgment garnishment status notices are not work items.
        "garnishment has been terminated"
    ],

    "Served on Notices": [
        "DC-BNV-A-SERV",
        "named party was served on"
    ],

    "Unserved": [
        "DC-BNV-A-UNSERV",
        "was returned unserved"
    ],

    "Notice of Dismissals 3-506": [
        "DC-BNV-T7",
        "DC-BNV-T8",
        "DC-BNV-T2",
        "dismissed this case upon stipulated terms"
    ],

    # Dismissed at the motion, trial or hearing itself
    "Dismissed at Hearing": [
        "The case was dismissed"
    ],

    # Rule 3-507 dismissal for lack of prosecution
    "Dismissal for Inactivity": [
        "no activity for more than one year"
    ],

    # The court reinstating a case it had dismissed
    "Dismissal Vacated": [
        "vacated the Order of Dismissal"
    ],

    "Cancellation Notice": [
    "Reason for Cancellation",
    "has been canceled",
    "has been cancelled",
    "NOTICE OF CANCELLATION"

    ],
    "Satisfatcion of Lien":[
        "Satisfaction of Lien Entered was recorded"
    ],
    "Bankrupcy":[
        "Suggestion of Bankruptcy was filed",
        # Common OCR rendering of "Bankruptcy" in the court's scan layer.
        "Suggestion of Banlauptcy was filed"
    ],

}


# Create output folders for each document type

def create_folders():
    INPUT_FOLDER.mkdir(parents=True, exist_ok=True)
    OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)

    # Normal sorted-document folders
    for doc_type in DOCUMENT_TYPES:
        folder_path = OUTPUT_FOLDER / doc_type
        folder_path.mkdir(parents=True, exist_ok=True)

    # Review folder
    review_folder = OUTPUT_FOLDER / "Review"
    review_folder.mkdir(parents=True, exist_ok=True)

    # Trash folders
    trash_root = OUTPUT_FOLDER / "Trash"
    trash_root.mkdir(parents=True, exist_ok=True)

    for trash_type in TRASH_DOCUMENTS:
        trash_folder = trash_root / trash_type
        trash_folder.mkdir(parents=True, exist_ok=True)

    # Duplicates of documents we keep
    (trash_root / DUPLICATES_FOLDER).mkdir(parents=True, exist_ok=True)

  

# Clean Extracted Text

# OCR drops junk glyphs right into the middle of the phrases we classify on.
# Real examples from the scans: "Conference \u00b7- Resolution",
# "because\ufffdthere has been no activity", "stipulated terms\ufffd and
# dismissed". Anything outside plain ASCII is scanner noise on these notices,
# so smart quotes and dashes get folded back to their ASCII form and everything
# else becomes a space - which then collapses away with the whitespace around
# it, repairing the phrase.
OCR_NOISE = {
    ord('\u2018'): "'", ord('\u2019'): "'",
    ord('\u201c'): '"', ord('\u201d'): '"',
    ord('\u2010'): '-', ord('\u2011'): '-', ord('\u2012'): '-',
    ord('\u2013'): '-', ord('\u2014'): '-', ord('\u2015'): '-',
}

NON_ASCII = re.compile(r'[^\x20-\x7e]')


def clean_text(text):

    # Fold typographic characters back to ASCII, then wipe the rest of the noise
    text = text.translate(OCR_NOISE)
    text = NON_ASCII.sub(' ', text)

    # Remove extra whitespace and newlines
    text = re.sub(r'\s+', ' ', text)

    return text.strip()


# Read Page

def extract_page_text(page):

    try:

        text = page.extract_text()

        if text:
            return clean_text(text)

        else:
            return ""

    except Exception as error:

        print(f"Could not read page: {error}")

        return ""

# Page Classification

WHITESPACE = re.compile(r'\s+')


def keyword_matches(keyword, text_lower, text_glued):

    keyword = keyword.lower()

    if keyword in text_lower:
        return True

    # OCR drops stray spaces into the middle of words and form codes:
    # "District Court of Mary land denied the motion", "DC-BNV-G1 B".
    # Comparing both sides with every space removed catches those without
    # needing a hand-written variant for each way a scan can break.
    return WHITESPACE.sub('', keyword) in text_glued


def classify_page(text):

    text_lower = text.lower()
    text_glued = WHITESPACE.sub('', text_lower)

    # Check TRASH documents FIRST
    for doc_type, keywords in TRASH_DOCUMENTS.items():

        for keyword in keywords:

            if keyword_matches(keyword, text_lower, text_glued):
                return doc_type, keyword, True

    # Check normal documents SECOND
    for doc_type, keywords in DOCUMENT_TYPES.items():

        for keyword in keywords:

            if keyword_matches(keyword, text_lower, text_glued):
                return doc_type, keyword, False

    # Nothing matched
    return "Review", None, False


# Show the exact extracted PDF text around the keyword that caused a match
def show_match_context(text, keyword, context=80):
    if not keyword:
        return None

    text_lower = text.lower()
    keyword_lower = keyword.lower()

    position = text_lower.find(keyword_lower)

    if position == -1:
        return keyword

    start = max(0, position - context)
    end = min(len(text), position + len(keyword) + context)

    return text[start:end]


# Case Number Extraction

def sanitize_filename_part(text):

    text = re.sub(r'[\\/:*?"<>|]+', '_', text)

    text = re.sub(r'\s+', ' ', text)

    return text.strip().replace(' ', '_')


# OCR mangles the words "Case Number" badly on these scans. Real examples
# from one batch: "Case N um her:" (200 pages), "Case Num her:", "Case Norn
# her:", "Case N urn her:", "Case Nuniber:", "Case Nuinber:", "Case Nui.nber:",
# "Case N um bet:", "Case.Number:", "Case '.Number:".
# Rather than chase every spelling of "Number", match the word "Case" - which
# survives OCR reliably - then a short run of letters/spaces/punctuation, then
# the colon that always follows.
CASE_LABEL = re.compile(
    r'CASE[\sA-Z.,\'"]{0,12}?[:;]',
    re.IGNORECASE
)

# Maryland district / circuit format: D-05-CV-26-020690, C-03-JG-26-012302
# The trailing sequence is always exactly 6 digits (3,541 of 3,562 in the
# 8-18-26 batch - every shorter one was OCR truncating a real 6-digit number).
# Pinning it to 6 stops a match from absorbing digits out of the next word.
# The numeric slots accept letters so OCR damage can be repaired below.
MARYLAND_SHAPE = re.compile(
    r'([A-Z])-?([A-Z0-9]{1,3})-?([A-Z]{2})-?([A-Z0-9]{2})-?([A-Z0-9]{6})'
)

# Older format: 08-04-0003889-2017 - the middle group is always 7 characters
LEGACY_SHAPE = re.compile(
    r'([A-Z0-9]{2})-([A-Z0-9]{2})-([A-Z0-9]{7})-([A-Z0-9]{4})'
)

# Junk OCR drops between the parts of a case number, where a hyphen belongs:
#   "D-07-CV...:26-015829"  "D ... 08-CV-26-031959"  "D-05..;CV-26-038016"
SEPARATOR_JUNK = re.compile(r'[.,;:~_\'"\-]+')


def match_case_shape(candidate):

    # Returns a canonical case number, or None when the text does not look
    # like one at all. Rejecting non-matches is deliberate: the old code
    # took the first token after the label, which turned body text into
    # case numbers ("ABSOLUTE", "TRACKING", "LVNV", and bare "D").

    maryland = MARYLAND_SHAPE.match(candidate)

    if maryland:

        prefix, court, kind, year, sequence = maryland.groups()

        # Repair only the slots that must be numeric, so a case type
        # like SC is never turned into 5C
        court = repair_ocr_digits(court)
        year = repair_ocr_digits(year)
        sequence = repair_ocr_digits(sequence)

        if court.isdigit() and year.isdigit() and sequence.isdigit():
            return f"{prefix}-{court}-{kind}-{year}-{sequence}"

    legacy = LEGACY_SHAPE.match(candidate)

    if legacy:

        first, second, middle, year = legacy.groups()

        first = repair_ocr_digits(first)
        second = repair_ocr_digits(second)
        year = repair_ocr_digits(year)

        # The middle group is left exactly as scanned - 02-03-0CV1895-1990
        # is a real case number in xaa.csv, letters and all
        if first.isdigit() and second.isdigit() and year.isdigit():
            return f"{first}-{second}-{middle}-{year}"

    return None


def parse_case_number(candidate):

    # Drop junk the scanner leaves between the colon and the number,
    # including the U+FFFD replacement character
    candidate = re.sub(r'^[^A-Z0-9]+', '', candidate)

    # Spaces land anywhere, including inside a segment, so close them up
    candidate = re.sub(r'\s+', '', candidate)

    # Two ways OCR breaks the number itself, so try both spellings:
    #   1. junk sitting where a hyphen belongs
    #      "D.'..08-CV-25-037 402" -> "D-08-CV-25-037402"
    #   2. junk sitting inside a segment, so drop separators entirely
    #      "D-101-CV-24-01368.7"   -> "D101CV24013687"
    hyphenated = SEPARATOR_JUNK.sub('-', candidate).strip('-')
    glued = re.sub(r'[^A-Z0-9]+', '', candidate)

    for attempt in (hyphenated, glued):

        case_number = match_case_shape(attempt)

        if case_number:
            return case_number

    return None


def extract_case_number(text):

    # The header case number is the first one on the page. That matters on
    # Lien of Judgment notices, which carry two: the district court case in
    # the header and the circuit court JG number the lien was recorded
    # under. Of 203 such pages, the district number was in xaa.csv all 203
    # times and the JG number none, so the header is the one we want.

    match = CASE_LABEL.search(text)

    if not match:
        return None

    remainder = text[match.end():match.end() + 50].upper()

    case_number = parse_case_number(remainder)

    if not case_number:
        return None

    return sanitize_filename_part(case_number)


# Case Number -> File Number Lookup

def normalize_case_number(case_number):

    # Uppercase, drop stray characters the PDF text layer adds
    case_number = case_number.upper()

    case_number = case_number.replace("~", "")
    case_number = case_number.replace(".", "")

    # Scanned pages sometimes break "CV" apart: "D-01-C V-26-000123"
    case_number = re.sub(r'-C\s+V-', '-CV-', case_number)

    # Remove all remaining whitespace
    case_number = re.sub(r'\s+', '', case_number)

    return case_number.strip('-').strip()


def case_key_variants(case_number):

    # Every spelling of one case number that we are willing to treat as a
    # match. Used to build the lookup index AND to search it, so both sides
    # get collapsed the same way.

    case_number = normalize_case_number(case_number)

    if not case_number:
        return []

    variants = [case_number]

    # Same case number without any hyphens
    variants.append(case_number.replace("-", ""))

    # Maryland district court format: D-01-CV-26-020795 / D-041-CV-26-007632
    # Collapse leading zeros so D-041-... and D-41-... land on the same key
    maryland = re.match(
        r'^([A-Z]{1,4})-?(\d{1,3})-?(CV|SC|CC)-?(\d{2})-?(\d{1,7})$',
        case_number
    )

    if maryland:

        prefix, court, kind, year, sequence = maryland.groups()

        variants.append(
            f"{prefix}-{int(court)}-{kind}-{year}-{int(sequence)}"
        )

    # Older format seen in xaa.csv: 12-34-5678901-2026
    legacy = re.match(
        r'^(\d{2})-(\d{2})-(\d{7})-(\d{4})$',
        case_number
    )

    if legacy:

        sequence = legacy.group(3).lstrip('0')
        year = legacy.group(4)

        # 5678901-26  and  5678901-2026
        variants.append(f"{sequence}-{year[2:]}")
        variants.append(f"{sequence}-{year}")

    # Short format: 10768-2000 -> also try 10768-00
    short = re.match(r'^(\d{1,6})-(\d{4})$', case_number)

    if short:
        variants.append(f"{int(short.group(1))}-{short.group(2)[2:]}")

    # De-duplicate, keep order, drop empties
    seen = []

    for variant in variants:

        if variant and variant not in seen:
            seen.append(variant)

    return seen


def repair_ocr_digits(segment):

    # Letters the PDF text layer commonly produces instead of digits.
    # S->5 and T->1 were confirmed against xaa.csv: "D-0S~CV-26-026166"
    # and "D-08-CV-26.:.0t8879" both resolve to real file numbers once
    # repaired. Only ever applied to slots that must be numeric.
    return (
        segment
        .replace("L", "1")
        .replace("I", "1")
        .replace("O", "0")
        .replace("S", "5")
        .replace("T", "1")
    )


def ocr_repair_case_number(case_number):

    # Rebuild a case number where letters landed in numeric positions,
    # e.g. D-LLL-CV-25-010494 -> D-111-CV-25-010494
    #      D-0L-CV-25-015671  -> D-01-CV-25-015671
    #      0L-01-0022643-2023 -> 01-01-0022643-2023
    # Returns None when nothing needed repairing.

    case_number = normalize_case_number(case_number)

    if not case_number:
        return None

    # Maryland format - court, year and sequence must be numeric
    maryland = re.match(
        r'^([A-Z])-([A-Z0-9]{1,3})-([A-Z]{2})-([A-Z0-9]{2})-([A-Z0-9]{1,7})$',
        case_number
    )

    if maryland:

        prefix, court, kind, year, sequence = maryland.groups()

        repaired = (
            f"{prefix}-{repair_ocr_digits(court)}-{kind}-"
            f"{repair_ocr_digits(year)}-{repair_ocr_digits(sequence)}"
        )

    else:

        legacy = re.match(
            r'^([A-Z0-9]{2})-([A-Z0-9]{2})-([A-Z0-9]{7})-([A-Z0-9]{4})$',
            case_number
        )

        if legacy:

            repaired = "-".join(
                repair_ocr_digits(part)
                for part in legacy.groups()
            )

        else:

            # Unknown format - only repair segments that already mix
            # letters and digits, so alpha suffixes like -LSS stay intact
            parts = []

            for part in case_number.split("-"):

                if re.search(r'\d', part) and re.search(r'[A-Z]', part):
                    parts.append(repair_ocr_digits(part))
                else:
                    parts.append(part)

            repaired = "-".join(parts)

    if repaired == case_number:
        return None

    return repaired


def load_file_number_lookup():

    # Returns (exact_matches, variant_matches) dicts of case number -> file number

    exact_matches = {}
    variant_matches = {}

    if not XAA_CSV.exists():

        print(
            f"WARNING: Lookup file not found: {XAA_CSV}"
        )
        print(
            "File numbers will be reported as MANUAL."
        )

        return exact_matches, variant_matches

    print(f"Loading case number lookup: {XAA_CSV}")

    with open(
        XAA_CSV,
        "r",
        newline="",
        encoding="utf-8",
        errors="replace"
    ) as lookup_file:

        reader = csv.DictReader(lookup_file)

        for row in reader:

            case_number = (row.get("case") or "").strip()
            file_number = (row.get("file") or "").strip()

            if not case_number or not file_number:
                continue

            key = normalize_case_number(case_number)

            # First row wins so an exact hit is never overwritten
            if key not in exact_matches:
                exact_matches[key] = file_number

            for variant in case_key_variants(case_number):

                if variant not in variant_matches:
                    variant_matches[variant] = file_number

    print(
        f"Loaded {len(exact_matches):,} case numbers "
        f"({len(variant_matches):,} match keys)."
    )
    print()

    return exact_matches, variant_matches


def lookup_case(case_number, exact_matches, variant_matches):

    # Returns the file number for a case number, or MANUAL

    key = normalize_case_number(case_number)

    # 1. Exact case number as it appears in xaa.csv
    if key in exact_matches:
        return exact_matches[key]

    # 2. Fall back to the collapsed / reformatted spellings
    for variant in case_key_variants(case_number):

        if variant in variant_matches:

            print(
                f"Matched case {case_number} "
                f"using alternate format {variant}"
            )

            return variant_matches[variant]

    return MANUAL


def find_file_number(case_number, exact_matches, variant_matches):

    # Returns (file_number, case_number, was_corrected)

    if not case_number:
        return MANUAL, case_number, False

    file_number = lookup_case(
        case_number,
        exact_matches,
        variant_matches
    )

    if file_number != MANUAL:
        return file_number, case_number, False

    # 3. Last resort - letters that should have been digits
    repaired = ocr_repair_case_number(case_number)

    if repaired:

        file_number = lookup_case(
            repaired,
            exact_matches,
            variant_matches
        )

        if file_number != MANUAL:

            print(
                f"Corrected case number {case_number} -> {repaired}"
            )

            return file_number, repaired, True

    return MANUAL, case_number, False

# Preventing Files from being overwritten

def get_unique_filename(output_folder, filename):

    output_path = output_folder / filename

    if not output_path.exists():
        return output_path

    file_stem = output_path.stem
    file_extension = output_path.suffix

    number = 2

    while True:

        new_filename = f"{file_stem}_{number}{file_extension}"

        new_output_path = output_folder / new_filename

        if not new_output_path.exists():
            return new_output_path

        number += 1


# Save Page

def save_page(
    reader,
    page_number,
    doc_type,
    case_number=None,
    file_number=MANUAL,
    is_trash=False,
    is_duplicate=False
):

    # DUPLICATES - a copy of this document was already saved, so keep it
    # out of the keep folders. Checked first so it wins over doc_type.
    if is_duplicate:
        output_folder = OUTPUT_FOLDER / "Trash" / DUPLICATES_FOLDER

    # REVIEW documents
    elif doc_type == "Review":
        output_folder = OUTPUT_FOLDER / "Review"

    # TRASH documents
    elif is_trash:
        output_folder = OUTPUT_FOLDER / "Trash" / doc_type

    # Normal documents we want to keep
    else:
        output_folder = OUTPUT_FOLDER / doc_type

    # Make sure folder exists
    output_folder.mkdir(
        parents=True,
        exist_ok=True
    )

    suffix = sanitize_filename_part(doc_type)

    has_file_number = (
        bool(file_number)
        and file_number not in (MANUAL, NOT_APPLICABLE)
    )

    if has_file_number:
        # Our file number is what everything else is filed under,
        # so it wins whenever we have one
        filename = (
            f"{sanitize_filename_part(file_number)}_{suffix}.pdf"
        )

    elif case_number:
        # No file number - either it is a trash page we never looked
        # up, or the case number is not in xaa.csv. Either way the
        # court's case number is the only identifier we have.
        filename = f"{case_number}_{suffix}.pdf"

    else:
        filename = f"{suffix}_{page_number + 1}.pdf"

    output_path = get_unique_filename(
        output_folder,
        filename
    )

    writer = PdfWriter()

    writer.add_page(
        reader.pages[page_number]
    )

    with open(output_path, "wb") as output_file:
        writer.write(output_file)

    return output_path


# Process PDF

def document_identity(doc_type, case_number, file_number):

    # What makes two keep pages "the same document". The case number stays
    # in the key even though it is no longer in the filename - two different
    # cases can share one file number, and those are separate documents,
    # not duplicates.

    return (
        doc_type,
        case_number or "",
        file_number or ""
    )


def process_pdf(
    pdf_path,
    document_counts,
    report_data,
    exact_matches,
    variant_matches,
    seen_documents
):

    print()
    print("=" * 60)
    print(f"Processing: {pdf_path.name}")

    reader = PdfReader(pdf_path)

    total_pages = len(reader.pages)

    print(f"Total Pages: {total_pages}")
    print()

    # Process each page
    for page_number, page in enumerate(
        reader.pages,
        start=1
    ):

        print(
            f"Processing Page "
            f"{page_number}/{total_pages}"
        )

        text = extract_page_text(page)

        status = "Sorted"
        destination = "Sorted"
        is_trash = False
        is_duplicate = False
        case_number = None
        file_number = MANUAL
        case_was_corrected = False
        matched_text = None
        matched_context = None

        # ------------------------------------
        # 1. Make sure page has readable text
        # ------------------------------------

        if not text:

            print(
                f"Page {page_number} is blank or unreadable. "
                f"Moving to Review folder."
            )

            doc_type = "Review"
            destination = "Review"
            status = "REVIEW - No readable text"

        else:

            # ------------------------------------
            # 3. Classify document
            # ------------------------------------

            doc_type, matched_text, is_trash = classify_page(text)

            if is_trash:
                destination = "Trash"
            else:
                destination = "Sorted"
            if matched_text:
                matched_context = show_match_context(text, matched_text)
                print(F"Matched keyword: {matched_text}")
                print(F"Exact PDF text around match: {matched_context}")

            # ------------------------------------
            # 4. Extract case number
            # ------------------------------------

            case_number = extract_case_number(text)

            # ------------------------------------
            # 4b. Look up our file number
            #
            #     Trash pages get thrown away, so there is no
            #     reason to spend a lookup on them. The case
            #     number is still used to name the file.
            # ------------------------------------

            if is_trash:

                file_number = NOT_APPLICABLE

            else:

                (
                    file_number,
                    case_number,
                    case_was_corrected
                ) = find_file_number(
                    case_number,
                    exact_matches,
                    variant_matches
                )

            # ------------------------------------
            # 5. Unknown document
            # ------------------------------------

            if doc_type == "Review":

                destination = "Review"
                status = "REVIEW - Could not classify"

                print(
                    f"Page {page_number} could not be classified. "
                    f"Moving to Review folder."
                )

            # ------------------------------------
            # 6. Known document but no case number
            # ------------------------------------

            elif destination == "Sorted" and not case_number:

                print(
                    f"Page {page_number} does not contain "
                    f"a case number. Moving to Review folder."
                )

                doc_type = "Review"
                destination = "Review"
                status = "REVIEW - No Case Number"

            elif destination == "Trash":

                status = f"TRASH - {doc_type}"
                print(
                    f"Page {page_number} classified as a trash document."
                )

            # ------------------------------------
            # 7. Known document, case number found,
            #    but no file number in xaa.csv
            # ------------------------------------

            elif file_number == MANUAL:

                status = "MANUAL - No File Number"

                print(
                    f"Page {page_number} case {case_number} "
                    f"is not in {XAA_CSV.name}. Needs manual review."
                )

            else:

                print(
                    f"Page {page_number} successfully classified."
                )

            # ------------------------------------
            # 8. Another copy of something already saved
            # ------------------------------------

            if destination == "Sorted":

                identity = document_identity(
                    doc_type,
                    case_number,
                    file_number
                )

                if identity in seen_documents:

                    is_duplicate = True
                    destination = "Duplicate"
                    status = f"DUPLICATE - {doc_type}"

                    print(
                        f"Page {page_number} is another copy of "
                        f"{doc_type} for {file_number} / {case_number}. "
                        f"Moving to Trash/{DUPLICATES_FOLDER}."
                    )

                else:
                    seen_documents.add(identity)

            # Flag OCR repairs regardless of where the page ended up,
            # so a corrected case number can always be audited
            if case_was_corrected:
                status = f"{status} (Case Number Corrected)"

        # ------------------------------------
        # Add to document count
        # ------------------------------------

        if is_duplicate:
            document_counts[DUPLICATES_FOLDER] += 1
        else:
            document_counts[doc_type] += 1

        # ------------------------------------
        # Save page
        # ------------------------------------

        output_path = save_page(
            reader=reader,
            page_number=page_number - 1,
            doc_type=doc_type,
            case_number=case_number,
            file_number=file_number,
            is_trash=is_trash,
            is_duplicate=is_duplicate
        )

        # ------------------------------------
        # Add to Excel report
        # ------------------------------------

        report_data.append({
            "Original PDF": pdf_path.name,
            "Page Number": page_number,
            "Document Type": doc_type,
            "Case Number": case_number or "Not Found",
            "File Number": file_number,
            "Status": status,
            # Not a spreadsheet column - the Status text already shows it.
            # Used to keep duplicates off the keep tabs.
            "Is Duplicate": is_duplicate,
            "Matched Keyword": matched_text or "None",
            "Matched PDF Text": matched_context or "None",
            "Saved Filename": output_path.name,
            "Saved Folder": str(output_path.parent)
        })

        # ------------------------------------
        # Console output
        # ------------------------------------

        print(f"Category: {doc_type}")

        if case_number:
            print(f"Case Number: {case_number}")
        else:
            print("Case Number: Not Found")

        print(f"File Number: {file_number}")
        print(f"Matched Keyword: {matched_text or 'None'}")
        print(f"Matched PDF Text: {matched_context or 'None'}")
        print(f"Saved to: {output_path}")
        print()

# Create Excel Report

def auto_fit_columns(worksheet):
    for column_cells in worksheet.columns:

        max_length = 0
        column_letter = column_cells[0].column_letter

        for cell in column_cells:

            if cell.value is not None:
                cell_length = len(str(cell.value))

                if cell_length > max_length: 
                    max_length = cell_length
        worksheet.column_dimensions[column_letter].width = min(max_length + 2, 60)


def build_report_row(document, fieldnames):

    return [
        document.get(field, "")
        for field in fieldnames
    ]


def is_keeper(document):

    # A page we are actually filing: classified, not trash, not Review, and
    # not a second copy of something already filed.

    return (
        document["Document Type"] != "Review"
        and document["Document Type"] not in TRASH_DOCUMENTS
        and not document.get("Is Duplicate")
    )


def create_excel_report(report_data):

    report_path = get_unique_filename(
        OUTPUT_FOLDER,
        "report.xlsx"
    )

    # Create ONE workbook
    workbook = Workbook()

    # ==================================================
    # TAB 1 - COURT MAIL REPORT
    # ==================================================

    worksheet = workbook.active
    worksheet.title = "Court Mail Report"

    fieldnames = [
        "Original PDF",
        "Page Number",
        "Document Type",
        "Case Number",
        "File Number",
        "Status",
        "Matched Keyword",
        "Matched PDF Text",
        "Saved Filename",
        "Saved Folder"
    ]

    worksheet.append(fieldnames)

    # Red for Review documents
    red_fill = PatternFill(
        fill_type="solid",
        fgColor="FF6666"
    )

    # Green for normal documents we are keeping
    green_fill = PatternFill(
        fill_type="solid",
        fgColor="92D050"
    )

    # Yellow for documents we classified but could not match to a file number
    yellow_fill = PatternFill(
        fill_type="solid",
        fgColor="FFD966"
    )

    # ==================================================
    # ADD ALL DOCUMENTS TO MAIN TAB
    # ==================================================

    for document in report_data:

        worksheet.append(
            build_report_row(document, fieldnames)
        )

        # Get the row we JUST added
        current_row = worksheet.max_row

        # REVIEW = RED
        if document["Document Type"] == "Review":

            for cell in worksheet[current_row]:
                cell.fill = red_fill

        # DUPLICATE = no fill. It went to Trash, so it is not a keeper
        # even though it carries a real document type.
        elif document.get("Is Duplicate"):
            pass

        # KEEP DOCUMENT - yellow when we could not match a file number,
        # green when we could
        elif is_keeper(document):

            fill = (
                yellow_fill
                if document["File Number"] == MANUAL
                else green_fill
            )

            for cell in worksheet[current_row]:
                cell.fill = fill

    # ==================================================
    # TAB 2 - DOCUMENTS WE KEEP
    # ==================================================

    keep_sheet = workbook.create_sheet(
        title="Documents to Keep"
    )

    keep_sheet.append(fieldnames)

    # Everything we are keeping - classified, not trash, not Review, no
    # duplicates. Yellow rows are the ones we could not match to a file number.
    for document in report_data:

        if is_keeper(document):

            keep_sheet.append(
                build_report_row(document, fieldnames)
            )

            if document["File Number"] == MANUAL:

                keep_current_row = keep_sheet.max_row

                for cell in keep_sheet[keep_current_row]:
                    cell.fill = yellow_fill

    # TAB 3 - REVIEW ONLY

    review_sheet = workbook.create_sheet(
        title="Review Only"
    )

    review_sheet.append(fieldnames)

    # Only put Review documents on this tab
    for document in report_data:

        if document["Document Type"] == "Review":

            review_sheet.append(
                build_report_row(document, fieldnames)
            )

            # Get the row we JUST added
            review_current_row = review_sheet.max_row

            # Highlight that row red
            for cell in review_sheet[review_current_row]:
                cell.fill = red_fill

    # TAB 4 - NO FILE NUMBER

    manual_sheet = workbook.create_sheet(
        title="No File Number"
    )

    manual_sheet.append(fieldnames)

    # Classified documents whose case number is not in xaa.csv
    for document in report_data:

        if is_keeper(document) and document["File Number"] == MANUAL:

            manual_sheet.append(
                build_report_row(document, fieldnames)
            )

            manual_current_row = manual_sheet.max_row

            for cell in manual_sheet[manual_current_row]:
                cell.fill = yellow_fill

    # Autofit Cells

    auto_fit_columns(worksheet)
    auto_fit_columns(keep_sheet)
    auto_fit_columns(review_sheet)
    auto_fit_columns(manual_sheet)

    # SAVE WORKBOOK

    workbook.save(report_path)

    return report_path

# Main

def main():

    print()
    print("Court Mail Sorter")
    print("=" * 60)
    print()

    create_folders()


    # Case Number -> File Number Lookup

    exact_matches, variant_matches = load_file_number_lookup()


    # Doc Counter

    document_counts = {}

    for doc_type in DOCUMENT_TYPES.keys():

        document_counts[doc_type] = 0

    for trash_type in TRASH_DOCUMENTS.keys():

        document_counts[trash_type] = 0

    document_counts["Review"] = 0
    document_counts[DUPLICATES_FOLDER] = 0

    # CSV Report Data

    report_data = []

    # Identities of the keep documents saved so far, so a second copy can
    # be spotted. Shared across every PDF in the input folder.

    seen_documents = set()

    # Find PDF files

    pdf_files = list(
        INPUT_FOLDER.glob("*.pdf")
    )

    if not pdf_files:

        print(
            "No PDF files found in the input folder."
        )

        return

    print(
        f"Found {len(pdf_files)} PDF file(s) "
        f"in the input folder."
    )


    # Process PDF Files

    for pdf_path in pdf_files:

        try:

            process_pdf(
                pdf_path,
                document_counts,
                report_data,
                exact_matches,
                variant_matches,
                seen_documents
            )

        except Exception as error:

            print(
                f"Error processing "
                f"{pdf_path.name}: {error}"
            )


    # Create Excel Report

    report_path = create_excel_report(
        report_data
    )


    # Document Summary

    print()
    print("=" * 60)
    print("Document Summary:")

    for doc_type, count in document_counts.items():

        print(
            f"  {doc_type}: {count}"
        )

    total_documents = sum(
        document_counts.values()
    )

    print("-" * 60)

    print(
        f"  Total Documents: {total_documents}"
    )

    print("-" * 60)


    # File Number Summary

    keepers = [
        document
        for document in report_data
        if is_keeper(document)
    ]

    no_file_number = [
        document
        for document in keepers
        if document["File Number"] == MANUAL
    ]

    duplicates = [
        document
        for document in report_data
        if document.get("Is Duplicate")
    ]

    print("File Number Lookup:")

    print(
        f"  Matched: {len(keepers) - len(no_file_number)}"
    )

    print(
        f"  No file number (MANUAL): {len(no_file_number)}"
    )

    print(
        f"  Duplicates moved to Trash/{DUPLICATES_FOLDER}: "
        f"{len(duplicates)}"
    )

    print("-" * 60)


    # Finished

    print()
    print("Finished")
    print("=" * 60)
    print()

    print(
        f"Check the output folder: "
        f"{OUTPUT_FOLDER}"
    )

    print()

    print(
        f"Excel Report Saved at: "
        f"{report_path}"
    )

    print()

    input(
        "Press Enter to close..."
    )


# Start the program

if __name__ == "__main__":

    main()
