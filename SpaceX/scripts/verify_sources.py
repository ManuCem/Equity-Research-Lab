"""
Is every number in the data file really printed on the page it cites?

data/spacex_data.csv is typed in by hand from the filings, so a typo is always possible.
This script opens each cited PDF page, reads its text with pdfplumber, and looks for the
number written the way filings write it (18674 -> "18,674", -4937 -> "4,937" inside brackets).

    python scripts/verify_sources.py

Rows that are skipped on purpose:
    - zeros (the filing prints a dash, not a number)
    - management commentary from the transcript (no page numbers, and the transcript is not published)

This is a check you run once after editing the data file; build_outputs.py does not need it.
Needs: pip install pdfplumber

Layout of this file (see CLAUDE.md):
    1. python / data code  -> read the data file, read the PDF pages, compare, and def main()
    2. html / chart code   -> none (this script only prints a report)
"""

import csv                                # reads the data file row by row
import re                                 # regular expressions: find "18,674" as a whole number, not inside "118,674"
from pathlib import Path                  # builds file paths that work on Windows, Mac and Linux

import pdfplumber                         # reads the text of a PDF page


# ===========================================================================
# python / data code
# ===========================================================================

HERE = Path(__file__).resolve().parent.parent           # the company folder (SpaceX/)
DATA_PATH = HERE / "data" / "spacex_data.csv"           # the hand-typed numbers
FILINGS_DIR = HERE / "filings"                          # where the PDFs live


def load_rows():
    """Read the data file and keep only the rows that can be checked against a PDF page."""
    with open(DATA_PATH, encoding="utf-8") as f:        # open the CSV as text
        rows = list(csv.DictReader(f))                  # each row becomes a dict keyed by the header names
    checkable = [r for r in rows                        # keep a row only if...
                 if r["source"].endswith(".pdf")        # ...it comes from a PDF (not the transcript)
                 and float(r["value"]) != 0]            # ...and is not a zero (zeros are printed as a dash)
    # returns a list of dicts, ~620 items, e.g.
    # [{'period': 'FY2025', 'metric': 'revenue', 'value': '18674', 'unit': 'USD m', 'statement': 'income_statement',
    #   'source': 's1a.pdf', 'page': '313', 'page_label': 'F-5', 'kind': 'filing', 'note': ''}, ...]
    return checkable


def read_pages(rows):
    """Extract the text of every page the data file cites, once per page."""
    wanted = {}                                         # file name -> set of page numbers we need
    for r in rows:                                      # collect the pages cited by each row
        wanted.setdefault(r["source"], set()).add(int(r["page"]))
    texts = {}                                          # (file, page) -> page text
    for name, pages in wanted.items():                  # open each PDF only once
        with pdfplumber.open(FILINGS_DIR / name) as pdf:
            for p in pages:                             # page numbers in the CSV start at 1, pdfplumber starts at 0
                texts[(name, p)] = pdf.pages[p - 1].extract_text(x_tolerance=1.5, y_tolerance=3) or ""
    # returns a dict with one entry per cited page (~30 pages), e.g.
    # {('s1a.pdf', 313): 'Space Exploration Technologies Corp.\nConsolidated Statements of Operations\n...',
    #  ('q3-2026-release.pdf', 9): 'Consolidated Financial Statements\n...'}
    return texts


def spellings(value):
    """The ways a filing may print a number: '18,674', and '85.7 billion' for round billions."""
    v = abs(float(value))                               # brackets carry the minus sign in filings, so search the size only
    forms = []                                          # list of strings to look for
    if v == int(v):                                     # whole numbers (all $ millions in this file)
        forms.append(f"{int(v):,}")                     # 18674 -> '18,674'
    else:                                               # decimals such as 2.5 (GW)
        forms.append(f"{v:g}")
    if v >= 1000 and v % 100 == 0:                      # round amounts that a release writes in words
        forms.append(f"{v / 1000:g} billion")           # 85700 -> '85.7 billion'
    # returns a list of 1-2 strings, e.g. spellings('-4937') -> ['4,937'];  spellings('85700') -> ['85,700', '85.7 billion']
    return forms


def check(rows, texts):
    """Look for each number on its cited page; return the rows that were NOT found."""
    missing = []                                        # rows we could not find
    for r in rows:                                      # one data row at a time
        text = texts[(r["source"], int(r["page"]))]     # the text of the cited page
        found = any(                                    # True if any spelling appears as a whole number
            re.search(r"(?<![\d,.])" + re.escape(s).replace("\\ ", r"\s+") + r"(?![,.]?\d)", text)  # not glued to
            for s in spellings(r["value"]))             # other digits; the space in '42.5 billion' may be a line break
        if not found:                                   # keep the failures for the report
            missing.append(r)
    # returns a list of dicts (the same shape as load_rows()), empty when every number was found, e.g.
    # []   or   [{'period': 'FY2025', 'metric': 'revenue', 'value': '18764', 'source': 's1a.pdf', 'page': '313', ...}]
    return missing


def main():
    rows = load_rows()                                  # step 1: the rows to check
    texts = read_pages(rows)                            # step 2: the text of every cited page
    missing = check(rows, texts)                        # step 3: compare
    print(f"Checked {len(rows)} numbers on {len(texts)} pages.")
    for r in missing:                                   # step 4: report each failure
        print(f"  NOT FOUND: {r['period']:<11} {r['metric']:<32} {r['value']:>8}  ({r['source']} p.{r['page']})")
    print("All numbers found on their cited pages." if not missing else f"{len(missing)} number(s) not found.")


# ===========================================================================
# html / chart code
# ===========================================================================
# (none: this script only prints a report)

if __name__ == "__main__":
    main()
