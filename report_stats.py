"""Lifetime stats across every court mail report ever generated.

Reads the "Court Mail Report" tab of every report*.xlsx in REPORT_FOLDER
(the tab that lists every page) and totals them up: how many of each kept
document type, and the keep vs trash split overall.
"""

from collections import Counter
from pathlib import Path

from openpyxl import load_workbook

# Where courtmail_aws.py writes report.xlsx, report_2.xlsx, ... (its
# OUTPUT_FOLDER).
REPORT_FOLDER = Path(r"F:\Legal\MD\Court Mail\Complete")

MAIN_SHEET = "Court Mail Report"


def bucket_for(status):
    """Sort a report row's Status text into keep / trash / duplicate / review."""

    status = (status or "").upper()

    if status.startswith("TRASH"):
        return "Trash"

    if status.startswith("DUPLICATE"):
        return "Duplicate"

    if status.startswith("REVIEW"):
        return "Review"

    # "Sorted", "MANUAL - No File Number", and any "(Case Number Corrected)"
    # variant of them are all pages that were filed.
    return "Keep"


def read_report(path, keep_types, buckets):
    """Add one report's rows to the running totals. Returns rows read."""

    workbook = load_workbook(path, read_only=True, data_only=True)

    if MAIN_SHEET not in workbook.sheetnames:
        workbook.close()
        return 0

    rows = workbook[MAIN_SHEET].iter_rows(values_only=True)
    header = next(rows, None)

    if not header or "Status" not in header or "Document Type" not in header:
        workbook.close()
        return 0

    status_col = header.index("Status")
    type_col = header.index("Document Type")
    count = 0

    for row in rows:

        if not row or row[status_col] is None:
            continue

        bucket = bucket_for(row[status_col])
        buckets[bucket] += 1
        count += 1

        if bucket == "Keep":
            keep_types[row[type_col]] += 1

    workbook.close()

    return count


def percent(part, whole):
    return f"{part / whole * 100:.1f}%" if whole else "n/a"


def main():
    print()
    print("Court Mail - Lifetime Stats")
    print("=" * 60)
    print(f"Reading reports from: {REPORT_FOLDER}")

    reports = sorted(
        path for path in REPORT_FOLDER.glob("report*.xlsx")
        if not path.name.startswith("~$")
    )

    if not reports:
        print("No report*.xlsx files found.")
        return

    keep_types = Counter()
    buckets = Counter()
    used = 0

    for path in reports:

        try:
            rows = read_report(path, keep_types, buckets)
        except Exception as error:  # pylint: disable=broad-exception-caught
            print(f"  Skipped {path.name}: {error}")
            continue

        if rows:
            used += 1
        else:
            print(f"  Skipped {path.name}: no Court Mail Report tab")

    total = sum(buckets.values())
    keep = buckets["Keep"]
    trash = buckets["Trash"]

    print(f"Reports counted: {used} of {len(reports)} ({total:,} pages)")
    print()

    print("Documents kept, by type:")
    width = max((len(str(name)) for name in keep_types), default=10)

    for name, count in sorted(keep_types.items(), key=lambda i: -i[1]):
        print(f"  {str(name):<{width}}  {count:>7,}  {percent(count, keep)}")

    print(f"  {'Total kept':<{width}}  {keep:>7,}")
    print()

    print("All pages by outcome:")

    for name in ("Keep", "Trash", "Duplicate", "Review"):
        print(
            f"  {name:<10} {buckets[name]:>7,}  "
            f"{percent(buckets[name], total)}"
        )

    print()
    print("Keep vs trash (excluding duplicates and review):")
    print(f"  Keep  {percent(keep, keep + trash)}  ({keep:,})")
    print(f"  Trash {percent(trash, keep + trash)}  ({trash:,})")
    print()
    print(
        "Note: a PDF processed in more than one run is counted once per "
        "run."
    )


if __name__ == "__main__":

    try:
        main()
    except Exception:  # pylint: disable=broad-exception-caught
        import traceback
        traceback.print_exc()
    finally:
        input("\nPress Enter to close...")
