"""
SpaceX: a basic accounting analysis, built from ONE data file.

Every number comes from data/spacex_data.csv (typed in from the filings in filings/, with the
document and page of each number). This script reads that file and writes all the outputs, so
they always agree with each other:

    data/spacex_accounting.xlsx   statements, ratios (Excel formulas), checks, sources
    spacex.html                   the web page (charts drawn with Apache ECharts from lib/)
    README.md                     only the block between the "key-numbers" markers is rewritten

Run it from anywhere:   python scripts/build_outputs.py

Sign convention: every value keeps the sign printed in the filing. Brackets in a filing become a
minus sign here, so interest expense is negative (-1,945) and a tax provision is positive (718 =
a tax charge; -549 = a tax benefit). Amounts are in US$ millions unless the unit column says otherwise.

Layout of this file (see CLAUDE.md):
    1. python / data code  -> read the data, calculate margins, ratios and checks, write the Excel
                              workbook (its ratios and checks are live Excel formulas), and def main()
    2. html / chart code   -> builds the web page and its charts
"""

import csv                                # reads the data file
import json                               # turns Python lists/dicts into text the page's JavaScript can read
from html import escape                   # makes text safe inside HTML (turns "<" into "&lt;", etc.)
from pathlib import Path                  # builds file paths that work on Windows, Mac and Linux

from openpyxl import Workbook             # writes the .xlsx workbook
from openpyxl.comments import Comment     # the hover note on each typed-in number (its source page)
from openpyxl.formatting.rule import CellIsRule  # colours a cell when a condition is true (FAIL -> red)
from openpyxl.styles import Alignment, Font, PatternFill  # cell look: alignment, font, fill colour
from openpyxl.utils import get_column_letter      # column number -> letter (3 -> "C")


# ===========================================================================
# python / data code
# ===========================================================================

HERE = Path(__file__).resolve().parent.parent           # the company folder (SpaceX/)
DATA_PATH = HERE / "data" / "spacex_data.csv"           # the single source of every number
XLSX_PATH = HERE / "data" / "spacex_accounting.xlsx"    # the workbook we write
HTML_PATH = HERE / "spacex.html"                        # the web page we write
README_PATH = HERE / "README.md"                        # we rewrite only its key-numbers block

COMPANY = "SpaceX"                                      # shown in titles
TICKER = "SPCX"                                         # Nasdaq ticker
TOLERANCE = 2                                           # $m: filings round every line, so sums can be off by 1-2

# Every period in the data file: id -> (short label for charts, long label for tables)
PERIODS = {
    "FY2023": ("2023", "Year 2023"), "FY2024": ("2024", "Year 2024"), "FY2025": ("2025", "Year 2025"),
    "Q1-2025": ("Q1 25", "Q1 2025"), "Q2-2025": ("Q2 25", "Q2 2025"), "Q3-2025": ("Q3 25", "Q3 2025"),
    "Q4-2025": ("Q4 25*", "Q4 2025 (not reported)"),     # not in the files: kept only to show the gap
    "Q1-2026": ("Q1 26", "Q1 2026"), "Q2-2026": ("Q2 26", "Q2 2026"), "Q3-2026": ("Q3 26", "Q3 2026"),
    "6M-2025": ("6M 25", "Jan-Jun 2025"), "6M-2026": ("6M 26", "Jan-Jun 2026"),
    "9M-2025": ("9M 25", "Jan-Sep 2025"), "9M-2026": ("9M 26", "Jan-Sep 2026"),
    "2024-12-31": ("Dec 24", "31 Dec 2024"), "2025-12-31": ("Dec 25", "31 Dec 2025"),
    "2026-03-31": ("Mar 26", "31 Mar 2026"), "2026-06-30": ("Jun 26", "30 Jun 2026"),
    "2026-09-30": ("Sep 26", "30 Sep 2026"),
}
PERIOD_ORDER = list(PERIODS)                            # column order used everywhere
YEARS = ["FY2023", "FY2024", "FY2025"]                  # audited full years (S-1/A)
QUARTERS = ["Q1-2025", "Q2-2025", "Q3-2025", "Q4-2025", "Q1-2026", "Q2-2026", "Q3-2026"]  # Q4 25 = gap
BS_DATES = ["2024-12-31", "2025-12-31", "2026-03-31", "2026-06-30", "2026-09-30"]  # balance sheet dates
COST_COLS = ["FY2023", "FY2024", "FY2025", "9M-2025", "9M-2026", "Q3-2026"]  # "where does the money go"
CASH_COLS = ["FY2023", "FY2024", "FY2025", "9M-2025", "9M-2026"]  # periods with a full cash flow statement
RATIO_COLS = ["FY2023", "FY2024", "FY2025", "9M-2025", "9M-2026",   # columns of the ratio tables
              "Q1-2025", "Q2-2025", "Q3-2025", "Q1-2026", "Q2-2026", "Q3-2026"]
BS_FOR = {"FY2024": "2024-12-31", "FY2025": "2025-12-31", "9M-2026": "2026-09-30",  # balance sheet that ends
          "Q1-2026": "2026-03-31", "Q2-2026": "2026-06-30", "Q3-2026": "2026-09-30"}  # each flow period
PRIOR = {"FY2024": "FY2023", "FY2025": "FY2024", "9M-2026": "9M-2025",  # same period one year earlier
         "Q1-2026": "Q1-2025", "Q2-2026": "Q2-2025", "Q3-2026": "Q3-2025"}

SEGMENTS = {"space": "Space", "connectivity": "Connectivity", "ai": "AI / SI"}  # id -> name on the page

# Human labels for every metric id in the data file (used by the workbook and the page)
LABELS = {
    "revenue": "Revenue", "cost_of_revenue": "Cost of revenue", "rnd": "Research and development",
    "sga": "Selling, general and administrative", "restructuring": "Restructuring charges (credits)",
    "impairment": "Impairment", "total_costs": "Total costs and expenses",
    "operating_income": "Income (loss) from operations", "interest_expense": "Interest expense",
    "interest_income": "Interest income", "other_income": "Other income (expense), net",
    "pretax_income": "Income (loss) before income taxes",
    "income_tax": "Provision for (benefit from) income taxes", "net_income": "Net income (loss)",
    "depreciation_amortization": "Depreciation and amortization", "share_based_comp": "Share-based compensation",
    "share_based_comp_acquisitions": "Share-based compensation related to acquisitions",
    "acquisition_costs": "Acquisition-related costs", "adj_ebitda": "Adjusted EBITDA (non-GAAP)",
    "capex_total": "Total capex (segment table)",
    "cash": "Cash and cash equivalents", "marketable_securities": "Marketable securities",
    "accounts_receivable": "Accounts receivable, net", "inventory": "Inventory",
    "prepaid_other_current": "Prepaid expenses and other current assets",
    "total_current_assets": "Total current assets", "ppe_net": "Property, plant and equipment, net",
    "finance_lease_rou": "Finance lease right-of-use assets", "intangibles": "Intangible assets, net",
    "digital_assets": "Digital assets", "goodwill": "Goodwill", "deferred_tax_assets": "Deferred tax assets",
    "other_assets": "Other assets", "total_assets": "Total assets", "accounts_payable": "Accounts payable",
    "deferred_revenue_current": "Deferred revenue, current", "debt_current": "Debt and finance leases, current",
    "accrued_other_current": "Accrued expenses and other current liabilities",
    "total_current_liabilities": "Total current liabilities",
    "deferred_revenue_noncurrent": "Deferred revenue, net of current",
    "debt_noncurrent": "Debt and finance leases, net of current", "other_liabilities": "Other liabilities",
    "total_liabilities": "Total liabilities", "redeemable_preferred": "Redeemable convertible preferred stock",
    "common_stock_a": "Class A common stock (par value)", "common_stock_b": "Class B common stock (par value)",
    "common_stock_c": "Class C common stock (par value)", "apic": "Additional paid-in capital",
    "accumulated_deficit": "Accumulated deficit", "aoci": "Accumulated other comprehensive income",
    "total_equity": "Total shareholders' equity",
    "total_liabilities_and_equity": "Total liabilities, preferred stock and shareholders' equity",
    "cfo": "Net cash from operating activities", "change_deferred_revenue": "  of which: change in deferred revenue",
    "capex_cf": "Purchases of property, plant and equipment (capex)",
    "securities_purchases": "Purchases of marketable securities",
    "securities_maturities": "Maturities of marketable securities",
    "securities_sales": "Sales of marketable securities", "cfi": "Net cash from investing activities",
    "debt_proceeds": "Proceeds from debt and other financing",
    "debt_repayments": "Repayments of debt and other financing",
    "stock_proceeds": "Proceeds from issuing capital stock (before the IPO)",
    "ipo_proceeds": "Proceeds from IPO, net", "buybacks": "Repurchases of common and preferred stock",
    "cff": "Net cash from financing activities", "fx_effect": "Effect of exchange rates on cash",
    "net_change_cash": "Net change in cash (incl. restricted cash)",
    "cash_begin": "Cash at start of period (incl. restricted cash)",
    "cash_end": "Cash at end of period (incl. restricted cash)",
    "backlog": "Backlog", "ipo_net_proceeds": "IPO net proceeds (approx.)", "bond_issue": "Bond issue, June 2026",
    "cursor_purchase_consideration": "Cursor: GAAP purchase consideration",
    "cursor_post_acquisition_awards": "Cursor: equity awards for post-acquisition service",
    "cursor_revenue": "Cursor: revenue contributed 14 Aug - 30 Sep 2026",
    "guidance_capex_low": "Guidance: Q4 2026 capex, low end", "guidance_capex_high": "Guidance: Q4 2026 capex, high end",
    "guidance_revenue_low": "Guidance: Q1 2027 revenue, low end",
    "guidance_revenue_high": "Guidance: Q1 2027 revenue, high end",
    "backlog_next_12m": "Backlog expected within 12 months",
    "guidance_nameplate_compute_low": "Guidance: nameplate compute end-2026, low end",
    "guidance_nameplate_compute_high": "Guidance: nameplate compute end-2026, high end",
    "loss_debt_extinguishment": "  of which: loss on debt extinguishment (non-cash add-back)",
    "other_income_related_party": "Other income: related-party part", "nameplate_compute": "Nameplate compute (SI segment)",
    "connectivity_gm_one_off": "Q3 Connectivity gross margin: one-off benefit",
}
for _seg, _name in SEGMENTS.items():                   # add the 12 segment labels (4 measures x 3 segments)
    LABELS[f"seg_revenue_{_seg}"] = f"{_name} revenue"
    LABELS[f"seg_operating_income_{_seg}"] = f"{_name} income (loss) from operations"
    LABELS[f"seg_adj_ebitda_{_seg}"] = f"{_name} segment adjusted EBITDA"
    LABELS[f"seg_capex_{_seg}"] = f"{_name} capex"

IS_KEYS = ["revenue", "cost_of_revenue", "rnd", "sga", "restructuring", "impairment", "total_costs",
           "operating_income", "interest_expense", "interest_income", "other_income", "pretax_income",
           "income_tax", "net_income"]                  # income statement lines, top to bottom
COST_LINES = ["cost_of_revenue", "rnd", "sga", "restructuring", "impairment"]  # lines that make total_costs

# How each input sheet of the workbook is laid out: sheet -> [(section title, [metric ids])]
INPUT_LAYOUT = {
    "Income Statement": [
        ("Income statement ($ millions)", IS_KEYS),
        ("Adjusted EBITDA reconciliation, non-GAAP ($ millions)", ["depreciation_amortization", "share_based_comp",
         "share_based_comp_acquisitions", "acquisition_costs", "adj_ebitda"])],
    "Segments": [
        ("Revenue by segment ($ millions)", [f"seg_revenue_{s}" for s in SEGMENTS]),
        ("Income (loss) from operations by segment ($ millions)", [f"seg_operating_income_{s}" for s in SEGMENTS]),
        ("Segment adjusted EBITDA, non-GAAP ($ millions)", [f"seg_adj_ebitda_{s}" for s in SEGMENTS]),
        ("Capital expenditures by segment ($ millions)", [f"seg_capex_{s}" for s in SEGMENTS] + ["capex_total"])],
    "Balance Sheet": [
        ("Assets ($ millions)", ["cash", "marketable_securities", "accounts_receivable", "inventory",
         "prepaid_other_current", "total_current_assets", "ppe_net", "finance_lease_rou", "intangibles",
         "digital_assets", "goodwill", "deferred_tax_assets", "other_assets", "total_assets"]),
        ("Liabilities ($ millions)", ["accounts_payable", "deferred_revenue_current", "debt_current",
         "accrued_other_current", "total_current_liabilities", "deferred_revenue_noncurrent", "debt_noncurrent",
         "other_liabilities", "total_liabilities"]),
        ("Preferred stock and shareholders' equity ($ millions)", ["redeemable_preferred", "common_stock_a",
         "common_stock_b", "common_stock_c", "apic", "accumulated_deficit", "aoci", "total_equity",
         "total_liabilities_and_equity"])],
    "Cash Flow": [
        ("Operating activities ($ millions)", ["cfo", "change_deferred_revenue", "loss_debt_extinguishment"]),
        ("Investing activities ($ millions)", ["capex_cf", "securities_purchases", "securities_maturities",
         "securities_sales", "cfi"]),
        ("Financing activities ($ millions)", ["debt_proceeds", "debt_repayments", "stock_proceeds", "ipo_proceeds",
         "buybacks", "cff"]),
        ("Change in cash ($ millions)", ["fx_effect", "net_change_cash", "cash_begin", "cash_end"])],
}

# Figures that are NOT in the files (left empty on purpose) and other things worth knowing
DATA_GAPS = [
    "Q4 2025 on its own is not reported in these files. The full year 2025 and Jan-Sep 2025 are, so you can "
    "work it out yourself (full year minus nine months); the charts leave the gap visible instead.",
    "There is no balance sheet for 31 Dec 2023 or 30 Sep 2025 in these files, so balance-sheet ratios for 2023 and "
    "Jan-Sep 2025 are left empty.",
    "The Q2 and Q3 releases give cash flows only year-to-date (Jan-Jun, Jan-Sep), so operating cash flow and free "
    "cash flow for Q2 and Q3 on their own are left empty (Q1 is in the S-1/A).",
    "The Q2 2026 release gives only 'selected' cash flow lines: no capex line and no exchange-rate line. That is why "
    "one cash flow check fails for Jan-Jun (see the Checks sheet).",
    "The S-1/A PDF does not show its amendment number or the date it was filed with the SEC; its cover says only "
    "'Subject to completion, dated June 3, 2026'.",
    "What makes up Q3 2026 'other income' ($1,025m, of which $561m from a related party) and the $390m tax benefit "
    "is explained in the 10-Q, which is not among these files.",
    "The Q2 2026 release PDF says it 'corrects a typographical error that appeared in the version furnished to the "
    "SEC' (the Q2 2025 compute figure), so it is not byte-for-byte the SEC copy.",
]
DISCREPANCIES = [
    "Shares at 31 Dec 2025 (millions): the S-1/A annual balance sheet shows 1,954 Class A outstanding, 644 Class B, "
    "482 Class C; its Q1 comparative column shows 1,952 / 643 / 484; the Q2 and Q3 releases show 1,952 / 644 / 482. "
    "The dollar amounts are identical. Share counts are not used in this analysis.",
    "Transcript vs filings: every overlapping figure checked (revenue, net income, adjusted EBITDA, segment revenue, "
    "subscribers, ARPU, Cursor revenue, Q3 capex, deferred revenue of $17.5bn) matches the filings. The transcript "
    "disagrees with itself on GPU deployment pace (about 200,000 vs 20,000 GB300s a month); that is not a filing number.",
    "The adjusted EBITDA definition changed in Q3 2026: two new add-backs (share-based compensation related to "
    "acquisitions, and acquisition-related costs). Earlier periods had no such items, so the figures stay "
    "comparable, but the definition is management's choice.",
]


def load_data():
    """Read the data file into a lookup dict and keep the source of every number."""
    with open(DATA_PATH, encoding="utf-8") as f:        # open the CSV as text
        rows = list(csv.DictReader(f))                  # one dict per row, keyed by the header names
    data, source = {}, {}                               # (period, metric) -> value, and -> row (for its source)
    for r in rows:                                      # fill both dicts
        r["value"] = float(r["value"])                  # "18674" text -> 18674.0 number
        data[(r["period"], r["metric"])] = r["value"]   # the number itself
        source[(r["period"], r["metric"])] = r          # the whole row: file, page, note...
    # returns a tuple of (list of 658 row dicts, dict of 658 numbers, dict of 658 rows), e.g.
    # data[('FY2025', 'revenue')] -> 18674.0
    # source[('FY2025', 'revenue')] -> {'period': 'FY2025', 'metric': 'revenue', 'value': 18674.0, 'unit': 'USD m',
    #                                   'statement': 'income_statement', 'source': 's1a.pdf', 'page': '313', ...}
    return rows, data, source


def get(data, period, metric):
    """One number, or None when the files do not contain it."""
    # returns a float or None, e.g. get(data, 'FY2025', 'revenue') -> 18674.0;  get(data, 'Q4-2025', 'revenue') -> None
    return data.get((period, metric))


# --- ratio definitions -------------------------------------------------------
# Each ratio = (sum of numerator terms) / (sum of denominator terms). A term is (sign, metric, basis):
#   basis "flow"  -> the column's own period (e.g. FY2025)
#   basis "bs"    -> the balance sheet at the end of that period (FY2025 -> 2025-12-31)
#   basis "prior" -> the same period one year earlier (FY2025 -> FY2024)
# The same definition is evaluated in Python (for the page) and turned into an Excel formula (for the workbook).
RATIOS = [
    {"id": "gross_margin", "name": "Gross margin", "fmt": "pct",
     "plain": "Of every $1 of revenue, how much is left after the direct cost of delivering the product.",
     "how": "(Revenue - Cost of revenue) / Revenue",
     "num": [(1, "revenue", "flow"), (-1, "cost_of_revenue", "flow")], "den": [(1, "revenue", "flow")]},
    {"id": "operating_margin", "name": "Operating margin", "fmt": "pct",
     "plain": "Of every $1 of revenue, how much is left after ALL operating costs, including R&D and SG&A.",
     "how": "Income (loss) from operations / Revenue",
     "num": [(1, "operating_income", "flow")], "den": [(1, "revenue", "flow")]},
    {"id": "net_margin", "name": "Net margin", "fmt": "pct",
     "plain": "Of every $1 of revenue, how much ends up as the bottom line after interest, other items and tax.",
     "how": "Net income (loss) / Revenue",
     "num": [(1, "net_income", "flow")], "den": [(1, "revenue", "flow")]},
    {"id": "capex_to_cfo", "name": "Capex / operating cash flow", "fmt": "x",
     "plain": "Dollars invested in new long-lived assets for each $1 the operations bring in. Above 1x, the "
              "company needs cash from elsewhere (savings, shares or loans).",
     "how": "Purchases of property, plant and equipment / Net cash from operating activities",
     "num": [(-1, "capex_cf", "flow")], "den": [(1, "cfo", "flow")]},
    {"id": "current_ratio", "name": "Current ratio", "fmt": "x",
     "plain": "Short-term assets (cash, receivables, inventory) for each $1 of bills due within a year. "
              "Above 1x, the short-term bills are covered.",
     "how": "Total current assets / Total current liabilities (end of period)",
     "num": [(1, "total_current_assets", "bs")], "den": [(1, "total_current_liabilities", "bs")]},
    {"id": "debt_to_equity", "name": "Debt-to-equity", "fmt": "x",
     "plain": "Borrowed money for each $1 of shareholders' equity. Higher = more reliance on lenders.",
     "how": "(Debt and finance leases, current + non-current) / Total shareholders' equity (end of period)",
     "num": [(1, "debt_current", "bs"), (1, "debt_noncurrent", "bs")], "den": [(1, "total_equity", "bs")]},
]
COST_SHARES = [                                         # "where does the money go": each line as % of revenue
    {"id": "cost_of_revenue", "name": "Cost of revenue", "num": [(1, "cost_of_revenue", "flow")]},
    {"id": "rnd", "name": "Research & development", "num": [(1, "rnd", "flow")]},
    {"id": "sga", "name": "Selling, general & admin.", "num": [(1, "sga", "flow")]},
    {"id": "other_operating", "name": "Restructuring & impairment",
     "num": [(1, "restructuring", "flow"), (1, "impairment", "flow")]},
    {"id": "interest", "name": "Interest expense", "num": [(-1, "interest_expense", "flow")]},
]
for _c in COST_SHARES:                                  # every cost share shares the same denominator and format
    _c.update({"den": [(1, "revenue", "flow")], "fmt": "pct", "how": f"{_c['name']} / Revenue"})
GROWTH = {"id": "revenue_growth", "name": "Revenue growth vs a year earlier", "fmt": "pct",
          "how": "Revenue / Revenue in the same period a year earlier - 1",
          "num": [(1, "revenue", "flow"), (-1, "revenue", "prior")], "den": [(1, "revenue", "prior")]}


def term_period(column, basis):
    """Which period a ratio term reads for a given column (or None if that period is not in the files)."""
    if basis == "flow":                                 # the column's own period
        return column
    if basis == "bs":                                   # the balance sheet at the end of the column's period
        return BS_FOR.get(column)
    # returns a period id or None, e.g. term_period('FY2025', 'bs') -> '2025-12-31';  term_period('FY2023', 'bs') -> None
    return PRIOR.get(column)                            # basis "prior": the same period a year earlier


def evaluate(data, spec, column):
    """Value of a ratio definition for one column, or None when an input is missing or the denominator is zero."""
    def total(terms):                                   # sum of sign x value over the terms, None if any is missing
        values = [get(data, term_period(column, b), m) for _, m, b in terms]  # look each term up
        if any(v is None for v in values):              # a missing input -> no ratio (never estimate)
            return None
        return sum(sign * v for (sign, _, _), v in zip(terms, values))  # apply the signs and add up
    num, den = total(spec["num"]), total(spec["den"])   # numerator and denominator
    if num is None or not den:                          # missing input, or division by zero
        return None
    # returns a float (a fraction for %, a multiple for x) or None, e.g.
    # evaluate(data, RATIOS[0], 'FY2025') -> 0.4939  (gross margin 49.4%);  evaluate(data, RATIOS[4], 'FY2023') -> None
    return num / den


def ratio_tables(data):
    """All ratios, cost shares and growth rates for the RATIO_COLS columns."""
    tables = {}                                         # spec id -> {column: value}
    for spec in RATIOS + COST_SHARES + [GROWTH]:        # every definition
        tables[spec["id"]] = {c: evaluate(data, spec, c) for c in RATIO_COLS}  # one value per column
    # returns a dict of 12 dicts (6 ratios, 5 cost shares, 1 growth), each with 11 columns, e.g.
    # {'gross_margin': {'FY2023': 0.4118, 'FY2024': 0.4295, 'FY2025': 0.4939, '9M-2025': 0.4933, '9M-2026': 0.5805, ...},
    #  'current_ratio': {'FY2023': None, 'FY2024': 1.3661, 'FY2025': 1.4463, ..., '9M-2026': 3.0843, ...}, ...}
    return tables


# --- chart and table data ------------------------------------------------------

def bn(value):
    """$ millions -> $ billions, rounded for charts (None stays None)."""
    # returns a float or None, e.g. bn(18674) -> 18.67;  bn(None) -> None
    return None if value is None else round(value / 1000, 2)


def revenue_series(data, periods):
    """Revenue by segment and in total for a list of periods, in $ billions."""
    values = [[bn(get(data, p, f"seg_revenue_{s}")) for p in periods] for s in SEGMENTS]  # one list per segment
    total = [bn(get(data, p, "revenue")) for p in periods]  # consolidated revenue (None for the Q4 2025 gap)
    # returns a dict, e.g. for QUARTERS:
    # {'labels': ['Q1 25', 'Q2 25', 'Q3 25', 'Q4 25*', 'Q1 26', 'Q2 26', 'Q3 26'],
    #  'values': [[0.86, 0.75, 1.67, None, 0.62, 0.96, 1.77], [2.48, ...], [0.73, ...]],
    #  'total': [4.07, 4.07, 5.58, None, 4.69, 7.81, 14.58]}
    return {"labels": [PERIODS[p][0] for p in periods], "values": values, "total": total}


def margin_series(data, periods):
    """Gross, operating and net margin (in %) for a list of periods."""
    specs = RATIOS[:3]                                  # the three margin definitions
    values = [[None if (v := evaluate(data, s, p)) is None else round(v * 100, 1) for p in periods] for s in specs]
    # returns a dict, e.g. for YEARS:
    # {'labels': ['2023', '2024', '2025'], 'names': ['Gross margin', 'Operating margin', 'Net margin'],
    #  'values': [[41.2, 42.9, 49.4], [-33.7, 3.3, -13.9], [-44.6, 5.6, -26.4]]}
    return {"labels": [PERIODS[p][0] for p in periods], "names": [s["name"] for s in specs], "values": values}


def cost_series(data):
    """Each cost line as % of revenue, for the COST_COLS periods."""
    values = [[None if (v := evaluate(data, c, p)) is None else round(v * 100, 1) for p in COST_COLS]
              for c in COST_SHARES]                     # one list per cost line
    total = [round(sum(col), 1) for col in zip(*values)]  # height of each stacked bar
    # returns a dict, e.g.
    # {'labels': ['2023', '2024', '2025', '9M 25', '9M 26', 'Q3 26'], 'names': ['Cost of revenue', ...],
    #  'values': [[58.8, 57.1, 50.6, 50.7, 41.9, 37.6], [20.3, 24.7, 46.3, ...], ...], 'total': [150.0, 108.0, 124.3, ...]}
    return {"labels": [PERIODS[p][0] for p in COST_COLS], "names": [c["name"] for c in COST_SHARES],
            "values": values, "total": total}


def cash_series(data):
    """Operating cash flow, capex (as a negative outflow) and free cash flow, in $ billions."""
    cfo = [get(data, p, "cfo") for p in CASH_COLS]      # cash from operations
    capex = [get(data, p, "capex_cf") for p in CASH_COLS]  # purchases of PP&E (already negative)
    fcf = [a + b for a, b in zip(cfo, capex)]           # free cash flow = CFO + (negative) capex
    # returns a dict in $bn, e.g.
    # {'labels': ['2023', '2024', '2025', '9M 25', '9M 26'], 'cfo': [4.52, 5.78, 6.79, 4.12, 13.1],
    #  'capex': [-4.42, -11.16, -20.74, -13.85, -48.07], 'fcf': [0.1, -5.39, -13.95, -9.73, -34.97]}
    return {"labels": [PERIODS[p][0] for p in CASH_COLS], "cfo": [bn(v) for v in cfo],
            "capex": [bn(v) for v in capex], "fcf": [bn(v) for v in fcf]}


def debt_cash(data, date):
    """Cash + marketable securities, debt (incl. finance leases) and the net position at one date, in $ millions."""
    cash = get(data, date, "cash") + get(data, date, "marketable_securities")  # money held
    debt = get(data, date, "debt_current") + get(data, date, "debt_noncurrent")  # money owed to lenders
    # returns a tuple of 3 floats in $m, e.g. debt_cash(data, '2026-09-30') -> (89855.0, 38844.0, 51011.0)
    return cash, debt, cash - debt


def debt_series(data):
    """debt_cash() for every balance sheet date, in $ billions."""
    rows = [debt_cash(data, d) for d in BS_DATES]       # one (cash, debt, net) tuple per date
    # returns a dict in $bn, e.g.
    # {'labels': ['Dec 24', 'Dec 25', 'Mar 26', 'Jun 26', 'Sep 26'], 'cash': [12.19, 24.75, 23.68, 100.01, 89.86],
    #  'debt': [13.79, 22.9, 30.27, 39.36, 38.84], 'net': [-1.61, 1.85, -6.59, 60.65, 51.01]}
    return {"labels": [PERIODS[d][0] for d in BS_DATES], "cash": [bn(r[0]) for r in rows],
            "debt": [bn(r[1]) for r in rows], "net": [bn(r[2]) for r in rows]}


# Steps from adjusted EBITDA to net income: (label, metric, sign applied to the stored value)
BRIDGE_STEPS = [
    ("Depreciation & amortization", "depreciation_amortization", -1),
    ("Share-based compensation", "share_based_comp", -1),
    ("Share-based comp. from acquisitions", "share_based_comp_acquisitions", -1),
    ("Acquisition-related costs", "acquisition_costs", -1),
    ("Restructuring", "restructuring", -1),
    ("Impairment", "impairment", -1),
    ("= Operating income", None, "operating_income"),     # subtotal row, checked against the income statement
    ("Interest expense", "interest_expense", 1),          # stored negative already
    ("Interest income", "interest_income", 1),
    ("Other income (expense)", "other_income", 1),
    ("= Income before tax", None, "pretax_income"),
    ("Income tax (benefit = +)", "income_tax", -1),       # a provision is a charge, a benefit adds to profit
    ("= Net income", None, "net_income"),
]


def ebitda_bridge(data, period):
    """Walk from adjusted EBITDA to net income for one period, step by step."""
    running = get(data, period, "adj_ebitda")           # start from management's adjusted EBITDA
    steps = [{"label": "Adjusted EBITDA", "delta": running, "total": running, "kind": "total", "reported": running}]
    for label, metric, how in BRIDGE_STEPS:             # apply each reconciling item in turn
        if metric is None:                              # a subtotal: compare our running total with the filing
            steps.append({"label": label, "delta": running, "total": running, "kind": "total",
                          "reported": get(data, period, how)})
            continue
        value = get(data, period, metric)               # the item as stored (None if the line did not exist)
        delta = 0.0 if value is None else how * value   # missing line = the item did not exist that period
        running += delta                                # move the running total
        steps.append({"label": label, "delta": delta, "total": running, "kind": "step", "reported": None})
    # returns a list of 14 dicts (start, 10 items, 3 subtotals) in $m, e.g. for 'Q3-2026':
    # [{'label': 'Adjusted EBITDA', 'delta': 8809.0, 'total': 8809.0, 'kind': 'total', 'reported': 8809.0},
    #  {'label': 'Depreciation & amortization', 'delta': -3644.0, 'total': 5165.0, 'kind': 'step', 'reported': None},
    #  ... {'label': '= Net income', 'delta': 4404.0, 'total': 4404.0, 'kind': 'total', 'reported': 4404.0}]
    return steps


def sources_and_uses(data, period):
    """Where the cash came from and where it went, from the cash flow statement lines."""
    p = lambda m: get(data, period, m) or 0.0           # a line, or 0 when the filing shows a dash
    rows = [                                            # (label, amount in $m: + = cash in, - = cash out)
        ("Cash from operations", p("cfo")),
        ("IPO (net of fees)", p("ipo_proceeds")),
        ("New borrowing", p("debt_proceeds")),
        ("Shares sold before the IPO", p("stock_proceeds")),
        ("Capex (new assets)", p("capex_cf")),
        ("Debt repaid", p("debt_repayments")),
        ("Marketable securities bought, net", p("securities_purchases") + p("securities_maturities")
         + p("securities_sales")),
        ("Share buybacks", p("buybacks")),
    ]
    other = p("net_change_cash") - sum(v for _, v in rows)  # everything else, so the table adds up exactly
    rows.append(("All other lines (remainder)", other))
    rows.append(("= Change in cash", p("net_change_cash")))  # the bottom line of the cash flow statement
    # returns a list of 10 (label, $m) tuples, e.g. for '9M-2026':
    # [('Cash from operations', 13097.0), ('IPO (net of fees)', 85675.0), ('New borrowing', 51834.0), ...,
    #  ('All other lines (remainder)', -724.0), ('= Change in cash', 28262.0)]
    return rows


# --- checks ----------------------------------------------------------------------

def build_checks(data):
    """Every arithmetic check: does the sum of some lines equal the line the filing reports?"""
    checks = []                                         # list of check dicts

    def add(group, name, period, terms, target, optional=(), note=""):
        """Add one check if all its inputs exist. terms/optional: [(sign, period, metric)], target: (period, metric)."""
        terms = list(terms) + [t for t in optional if (t[1], t[2]) in data]  # optional lines only if present
        needed = [(p, m) for _, p, m in terms] + [target]  # every number the check reads
        if any(k not in data for k in needed):          # skip checks whose inputs are not in the files
            return
        calculated = sum(s * data[(p, m)] for s, p, m in terms)  # add up the lines with their signs
        diff = calculated - data[target]                # calculated minus reported
        status = "PASS" if abs(diff) < 0.5 else "PASS (rounding)" if abs(diff) <= TOLERANCE else "FAIL"
        checks.append({"group": group, "name": name, "period": period, "terms": terms, "target": target,
                       "calculated": calculated, "reported": data[target], "diff": diff, "status": status,
                       "note": note if status == "FAIL" else ""})

    flows = [p for p in PERIOD_ORDER if (p, "revenue") in data]  # every period with an income statement
    for p in flows:                                     # 1. segments add up to the consolidated total
        for measure, total in [("seg_revenue", "revenue"), ("seg_operating_income", "operating_income"),
                               ("seg_adj_ebitda", "adj_ebitda"), ("seg_capex", "capex_total")]:
            add("Segments add up to the total", f"{LABELS[total]} = sum of the 3 segments", p,
                [(1, p, f"{measure}_{s}") for s in SEGMENTS], (p, total))
        add("Segments add up to the total", "Segment capex total = capex in the cash flow statement", p,
            [(-1, p, "capex_cf")], (p, "capex_total"))
    for p in flows:                                     # 2. the income statement adds up
        g = "Income statement adds up"
        add(g, "Cost lines = total costs and expenses", p, [(1, p, m) for m in COST_LINES], (p, "total_costs"))
        add(g, "Revenue - total costs = operating income", p, [(1, p, "revenue"), (-1, p, "total_costs")],
            (p, "operating_income"))
        add(g, "Operating income + interest + other = income before tax", p,
            [(1, p, "operating_income"), (1, p, "interest_expense"), (1, p, "interest_income"), (1, p, "other_income")],
            (p, "pretax_income"))
        add(g, "Income before tax - tax = net income", p, [(1, p, "pretax_income"), (-1, p, "income_tax")],
            (p, "net_income"))
        add("Adjusted EBITDA reconciliation adds up", "Net income + add-backs = adjusted EBITDA", p,
            [(1, p, "net_income"), (1, p, "depreciation_amortization"), (1, p, "share_based_comp"),
             (1, p, "restructuring"), (1, p, "impairment"), (-1, p, "interest_expense"), (-1, p, "interest_income"),
             (-1, p, "other_income"), (1, p, "income_tax")], (p, "adj_ebitda"),
            optional=[(1, p, "share_based_comp_acquisitions"), (1, p, "acquisition_costs")])
    for d in BS_DATES:                                  # 3. the balance sheet balances
        g = "Balance sheet balances"
        add(g, "Current asset lines = total current assets", d, [(1, d, m) for m in
            ["cash", "marketable_securities", "accounts_receivable", "inventory", "prepaid_other_current"]],
            (d, "total_current_assets"))
        add(g, "Asset lines = total assets", d, [(1, d, m) for m in ["total_current_assets", "ppe_net",
            "finance_lease_rou", "intangibles", "digital_assets", "goodwill", "deferred_tax_assets", "other_assets"]],
            (d, "total_assets"))
        add(g, "Current liability lines = total current liabilities", d, [(1, d, m) for m in
            ["accounts_payable", "deferred_revenue_current", "debt_current", "accrued_other_current"]],
            (d, "total_current_liabilities"))
        add(g, "Liability lines = total liabilities", d, [(1, d, m) for m in ["total_current_liabilities",
            "deferred_revenue_noncurrent", "debt_noncurrent", "other_liabilities"]], (d, "total_liabilities"))
        add(g, "Equity lines = total shareholders' equity", d, [(1, d, m) for m in ["common_stock_a",
            "common_stock_b", "common_stock_c", "apic", "accumulated_deficit", "aoci"]], (d, "total_equity"))
        add(g, "Assets = liabilities + preferred stock + equity", d, [(1, d, "total_liabilities"),
            (1, d, "redeemable_preferred"), (1, d, "total_equity")], (d, "total_assets"))
        add(g, "Reported 'total liabilities and equity' = total assets", d, [(1, d, "total_liabilities_and_equity")],
            (d, "total_assets"))
    for p in [p for p in PERIOD_ORDER if (p, "cfo") in data]:  # 4. the cash flow statement reconciles
        g = "Cash flow reconciles"
        no_fx = (p, "fx_effect") not in data            # the Q2 release leaves the exchange-rate line out
        add(g, "Operating + investing + financing (+ FX) = change in cash", p,
            [(1, p, "cfo"), (1, p, "cfi"), (1, p, "cff")], (p, "net_change_cash"), optional=[(1, p, "fx_effect")],
            note="The Q2 release shows only 'selected' cash flow lines and leaves out the exchange-rate effect, "
                 "so the three sections cannot add up to the change in cash." if no_fx else "")
        add(g, "Cash at start + change = cash at end", p, [(1, p, "cash_begin"), (1, p, "net_change_cash")],
            (p, "cash_end"))
        if p.endswith("2026"):                          # every 2026 period starts where 2025 ended
            add(g, "2026 starts with the cash that 2025 ended with", p, [(1, "FY2025", "cash_end")], (p, "cash_begin"))
    for year in ["2025", "2026"]:                       # 5. quarters add up to year-to-date
        q1, q2, q3, h1, m9 = (f"Q1-{year}", f"Q2-{year}", f"Q3-{year}", f"6M-{year}", f"9M-{year}")
        for m in ["revenue", "operating_income", "net_income", "adj_ebitda", "capex_total"] + \
                 [f"seg_revenue_{s}" for s in SEGMENTS] + [f"seg_adj_ebitda_{s}" for s in SEGMENTS]:
            add("Quarters add up to year-to-date", f"Q1 + Q2 = Jan-Jun: {LABELS[m]}", h1, [(1, q1, m), (1, q2, m)], (h1, m))
            add("Quarters add up to year-to-date", f"Jan-Jun + Q3 = Jan-Sep: {LABELS[m]}", m9, [(1, h1, m), (1, q3, m)],
                (m9, m))
    for start, end, flow in [("2024-12-31", "2025-12-31", "FY2025"), ("2025-12-31", "2026-03-31", "Q1-2026"),
                             ("2026-03-31", "2026-06-30", "Q2-2026"), ("2026-06-30", "2026-09-30", "Q3-2026")]:
        add("Balance sheet links to the income statement",   # 6. no dividends: the deficit moves by net income
            f"Accumulated deficit at {PERIODS[start][1]} + net income = deficit at {PERIODS[end][1]}", flow,
            [(1, start, "accumulated_deficit"), (1, flow, "net_income")], (end, "accumulated_deficit"))
    # returns a list of ~230 check dicts, e.g.
    # [{'group': 'Segments add up to the total', 'name': 'Revenue = sum of the 3 segments', 'period': 'FY2023',
    #   'terms': [(1, 'FY2023', 'seg_revenue_space'), ...], 'target': ('FY2023', 'revenue'), 'calculated': 10387.0,
    #   'reported': 10387.0, 'diff': 0.0, 'status': 'PASS', 'note': ''}, ...]
    return checks


def headline(data, tables, checks):
    """The handful of numbers quoted in the text of the page and the README."""
    q_ni = [get(data, q, "net_income") for q in QUARTERS if get(data, q, "net_income") is not None]  # 6 quarters
    cash, debt, net = debt_cash(data, "2026-09-30")     # latest balance sheet
    h = {
        "rev_fy23": get(data, "FY2023", "revenue"), "rev_fy25": get(data, "FY2025", "revenue"),
        "rev_9m26": get(data, "9M-2026", "revenue"), "rev_q3_26": get(data, "Q3-2026", "revenue"),
        "rev_q3_25": get(data, "Q3-2025", "revenue"), "growth_q3": tables["revenue_growth"]["Q3-2026"],
        "ni_q3_26": get(data, "Q3-2026", "net_income"), "ni_9m26": get(data, "9M-2026", "net_income"),
        "profitable_quarters": sum(v > 0 for v in q_ni), "quarters": len(q_ni),
        "cfo_9m26": get(data, "9M-2026", "cfo"), "capex_9m26": -get(data, "9M-2026", "capex_cf"),
        "fcf_9m26": get(data, "9M-2026", "cfo") + get(data, "9M-2026", "capex_cf"),
        "cash_sep26": cash, "debt_sep26": debt, "net_sep26": net,
        "checks": len(checks), "checks_fail": sum(c["status"] == "FAIL" for c in checks),
        "checks_rounding": sum(c["status"] == "PASS (rounding)" for c in checks),
    }
    # returns a dict of ~20 numbers, e.g.
    # {'rev_fy23': 10387.0, 'rev_fy25': 18674.0, 'rev_9m26': 27083.0, 'rev_q3_26': 14575.0, 'growth_q3': 1.6139,
    #  'ni_q3_26': 4404.0, 'profitable_quarters': 1, 'quarters': 6, 'capex_9m26': 48069.0, 'fcf_9m26': -34972.0,
    #  'net_sep26': 51011.0, 'checks': 233, 'checks_fail': 2, ...}
    return h


def money(m, digits=1):
    """$ millions -> readable text: 18674 -> '$18.7bn', 608 -> '$608m', -4937 -> '-$4.9bn'."""
    sign = "-" if m < 0 else ""                         # put the minus sign before the dollar sign
    a = abs(m)                                          # size without the sign
    # returns a str, e.g. money(18674) -> '$18.7bn';  money(-413) -> '-$413m';  money(85675, 0) -> '$86bn'
    return f"{sign}${a / 1000:,.{digits}f}bn" if a >= 1000 else f"{sign}${a:,.0f}m"


def pct(x, digits=1):
    """Fraction -> percent text: 0.4939 -> '49.4%'."""
    # returns a str, e.g. pct(0.4939) -> '49.4%';  pct(None) -> 'n/a'
    return "n/a" if x is None else f"{x * 100:.{digits}f}%"


def update_readme(h, checks):
    """Rewrite only the key-numbers block of README.md, so the README always matches the data."""
    start, end = "<!-- key-numbers:start -->", "<!-- key-numbers:end -->"  # markers that frame the block
    if not README_PATH.exists():                        # no README yet: nothing to update
        return None
    text = README_PATH.read_text(encoding="utf-8")      # the current README
    if start not in text or end not in text:            # nothing to update if the markers were removed
        return None
    fails = [c for c in checks if c["status"] == "FAIL"]  # failed checks, listed by name
    lines = [
        f"- Revenue: {money(h['rev_fy23'])} (2023) -> {money(h['rev_fy25'])} (2025) -> "
        f"{money(h['rev_9m26'])} in Jan-Sep 2026 alone.",
        f"- Q3 2026: revenue {money(h['rev_q3_26'])} ({pct(h['growth_q3'], 0)} vs Q3 2025), "
        f"net income {money(h['ni_q3_26'])} (the only profitable quarter of the {h['quarters']} in these files).",
        f"- Jan-Sep 2026: operating cash flow {money(h['cfo_9m26'])}, capex {money(h['capex_9m26'])}, "
        f"free cash flow {money(h['fcf_9m26'])}.",
        f"- 30 Sep 2026: cash and marketable securities {money(h['cash_sep26'])}, debt {money(h['debt_sep26'])}, "
        f"net cash {money(h['net_sep26'])}.",
        f"- Checks: {h['checks']} run, {h['checks'] - h['checks_fail']} pass "
        f"({h['checks_rounding']} only within ${TOLERANCE}m rounding), {h['checks_fail']} fail"
        + (": " + "; ".join(f"{c['name']} ({PERIODS[c['period']][1]}, off by {c['diff']:+,.0f})" for c in fails)
           if fails else "") + ".",
    ]
    block = start + "\n_Generated by `scripts/build_outputs.py` from `data/spacex_data.csv`._\n\n" + \
        "\n".join(lines) + "\n" + end              # the new block
    before, rest = text.split(start, 1)                 # text before the block
    after = rest.split(end, 1)[1]                       # text after the block
    README_PATH.write_text(before + block + after, encoding="utf-8")  # write it back
    # returns nothing; rewrites the block in README.md, e.g.
    # - Revenue: $10.4bn (2023) -> $18.7bn (2025) -> $27.1bn in Jan-Sep 2026 alone.
    return None


# --- Excel workbook ------------------------------------------------------------

ARIAL = "Arial"                                          # one professional font throughout
BLUE = Font(name=ARIAL, color="0000FF")                 # typed-in inputs (financial-model convention)
BLACK = Font(name=ARIAL, color="000000")                # formulas
BOLD = Font(name=ARIAL, bold=True)                      # headers
TITLE = Font(name=ARIAL, bold=True, size=14)            # sheet titles
NOTE = Font(name=ARIAL, italic=True, color="595959", size=9)  # small explanations
HEAD_FILL = PatternFill("solid", fgColor="DDEBF7")      # light blue header band
SECTION_FILL = PatternFill("solid", fgColor="F2F2F2")   # light grey section rows
MONEY = '#,##0;(#,##0);"-"'                             # 18,674  (4,937)  -   (zeros show as a dash)
PCT = '0.0%;(0.0%);"-"'                                 # 49.4%  (13.9%)
MULT = '0.00"x";(0.00"x");"-"'                          # 1.45x


def sheet_title(ws, title, subtitle):
    """Write the title rows every sheet starts with."""
    ws["A1"] = title                                    # big title
    ws["A1"].font = TITLE
    ws["A2"] = subtitle                                 # one-line explanation
    ws["A2"].font = NOTE
    # returns nothing; writes A1 and A2 of the sheet
    return None


def write_input_sheet(wb, name, layout, data, source, cell):
    """One statement sheet: rows = line items, columns = periods; every number typed in blue with its source."""
    ws = wb.create_sheet(name)                          # new tab
    metrics = [m for _, ms in layout for m in ms]       # every line item on this sheet
    periods = [p for p in PERIOD_ORDER if any((p, m) in data for m in metrics)]  # columns that have data
    sheet_title(ws, f"{COMPANY} - {name}", "US$ millions. Blue = typed in from the filing (hover a number to see "
                "its document and page). Empty = not in these files. Brackets = negative.")
    ws.cell(4, 1, "Line item").font = BOLD              # header row
    ws.cell(4, 2, "Metric id").font = BOLD
    for j, p in enumerate(periods, start=3):            # one column per period, from column C
        c = ws.cell(4, j, PERIODS[p][1])                # long label, e.g. "Year 2025"
        c.font, c.alignment = BOLD, Alignment(horizontal="center", wrap_text=True)
    for j in range(1, len(periods) + 3):                # colour the header band
        ws.cell(4, j).fill = HEAD_FILL
    row = 5                                             # first data row
    for title, ms in layout:                            # one block per section
        ws.cell(row, 1, title).font = BOLD              # section title row
        for j in range(1, len(periods) + 3):
            ws.cell(row, j).fill = SECTION_FILL
        row += 1
        for m in ms:                                    # one row per line item
            ws.cell(row, 1, LABELS[m]).font = BLACK     # readable label
            ws.cell(row, 2, m).font = NOTE              # the id used in the CSV (to find it there)
            for j, p in enumerate(periods, start=3):    # fill the numbers that exist
                if (p, m) not in data:                  # not in the files -> leave the cell empty
                    continue
                c = ws.cell(row, j, data[(p, m)])       # the typed-in number
                c.font, c.number_format = BLUE, MONEY
                r = source[(p, m)]                      # where it came from
                c.comment = Comment(f"{r['source']}, PDF page {r['page']} (printed page {r['page_label']})"
                                    + (f". {r['note']}" if r["note"] else ""), "source")
                cell[(p, m)] = f"'{name}'!{get_column_letter(j)}{row}"  # remember the address for formulas
            row += 1
    ws.column_dimensions["A"].width = 52                # wide label column
    ws.column_dimensions["B"].width = 24
    for j in range(3, len(periods) + 3):                # equal number columns
        ws.column_dimensions[get_column_letter(j)].width = 13
    ws.freeze_panes = "C5"                              # keep labels and headers visible while scrolling
    # returns nothing; adds the sheet and fills `cell` with addresses, e.g.
    # cell[('FY2025', 'revenue')] -> "'Income Statement'!E6"
    return None


def excel_formula(spec, column, cell):
    """Turn a ratio definition into an Excel formula for one column, or None when an input is missing."""
    def total(terms):                                   # "+'Sheet'!C6-'Sheet'!C7" for the terms
        parts = []
        for sign, metric, basis in terms:               # one cell reference per term
            ref = cell.get((term_period(column, basis), metric))  # its address, if that number exists
            if ref is None:                             # missing input -> no formula
                return None
            parts.append(("+" if sign > 0 else "-") + ref)
        return "".join(parts)
    num, den = total(spec["num"]), total(spec["den"])   # numerator and denominator text
    if num is None or den is None:
        return None
    # returns a str or None, e.g. excel_formula(RATIOS[0], 'FY2025', cell)
    # -> "=IFERROR((+'Income Statement'!E6-'Income Statement'!E7)/(+'Income Statement'!E6),\"\")"
    return f'=IFERROR(({num})/({den}),"")'


def write_ratios_sheet(wb, data, cell):
    """Ratios as live Excel formulas that read the statement sheets."""
    ws = wb.create_sheet("Ratios")                      # new tab
    sheet_title(ws, f"{COMPANY} - Ratios", "Every value here is an Excel formula reading the statement sheets "
                "(black text). Empty = an input is not in these files.")
    row = 4                                             # first block starts on row 4

    def header(title, columns, extra):                  # a block header row
        nonlocal row
        ws.cell(row, 1, title).font = BOLD
        ws.cell(row, 2, "What it tells you").font = BOLD
        for j, p in enumerate(columns, start=3):
            ws.cell(row, j, PERIODS[p][1]).font = BOLD
        ws.cell(row, len(columns) + 3, extra).font = BOLD
        for j in range(1, len(columns) + 4):
            ws.cell(row, j).fill = HEAD_FILL
            ws.cell(row, j).alignment = Alignment(wrap_text=True, vertical="center")
        row += 1

    def spec_rows(specs, fmt_default=None):             # one row per definition
        nonlocal row
        for s in specs:
            ws.cell(row, 1, s["name"]).font = BLACK
            ws.cell(row, 2, s.get("plain", "")).font = NOTE
            ws.cell(row, 2).alignment = Alignment(wrap_text=True, vertical="top")
            for j, p in enumerate(RATIO_COLS, start=3):  # formula per column, when the inputs exist
                f = excel_formula(s, p, cell)
                if f:
                    c = ws.cell(row, j, f)
                    c.font = BLACK
                    c.number_format = PCT if s["fmt"] == "pct" else MULT
            ws.cell(row, len(RATIO_COLS) + 3, s["how"]).font = NOTE
            row += 1
        row += 1

    header("Six key ratios", RATIO_COLS, "How it is calculated")
    spec_rows(RATIOS)
    header("Where the money goes: each cost as % of revenue", RATIO_COLS, "How it is calculated")
    spec_rows(COST_SHARES)
    header("Growth", RATIO_COLS, "How it is calculated")
    spec_rows([GROWTH])
    header("Profit vs cash ($ millions)", RATIO_COLS, "How it is calculated")  # CFO, capex, free cash flow
    for label, metric, how in [("Net cash from operating activities", "cfo", "From the cash flow statement"),
                               ("Capex (purchases of PP&E)", "capex_cf", "From the cash flow statement (negative = spent)"),
                               ("Free cash flow", None, "Operating cash flow + capex (capex is negative)")]:
        ws.cell(row, 1, label).font = BLACK
        for j, p in enumerate(RATIO_COLS, start=3):
            refs = [cell.get((p, "cfo")), cell.get((p, "capex_cf"))] if metric is None else [cell.get((p, metric))]
            if all(refs):                               # write the formula only when every input exists
                c = ws.cell(row, j, "=" + "+".join(refs))
                c.font, c.number_format = BLACK, MONEY
        ws.cell(row, len(RATIO_COLS) + 3, how).font = NOTE
        row += 1
    row += 1
    header("Debt and cash ($ millions, end of period)", BS_DATES, "How it is calculated")
    for label, parts, how in [("Cash + marketable securities", ["+cash", "+marketable_securities"], "Balance sheet"),
                              ("Debt and finance leases", ["+debt_current", "+debt_noncurrent"], "Current + non-current"),
                              ("Net cash (negative = net debt)", ["+cash", "+marketable_securities", "-debt_current",
                                                                  "-debt_noncurrent"], "Cash + securities - debt")]:
        ws.cell(row, 1, label).font = BLACK
        for j, d in enumerate(BS_DATES, start=3):
            c = ws.cell(row, j, "=" + "".join(x[0] + cell[(d, x[1:])] for x in parts))
            c.font, c.number_format = BLACK, MONEY
        ws.cell(row, len(BS_DATES) + 3, how).font = NOTE
        row += 1
    ws.column_dimensions["A"].width = 40                # layout
    ws.column_dimensions["B"].width = 60
    for j in range(3, len(RATIO_COLS) + 3):
        ws.column_dimensions[get_column_letter(j)].width = 12
    ws.column_dimensions[get_column_letter(len(RATIO_COLS) + 3)].width = 60
    ws.freeze_panes = "C4"
    # returns nothing; adds the Ratios sheet (about 110 formulas)
    return None


def write_bridge_sheet(wb, data, cell):
    """The challenge: adjusted EBITDA -> net income, as formulas, for three periods."""
    ws = wb.create_sheet("EBITDA Bridge")               # new tab
    sheet_title(ws, f"{COMPANY} - From adjusted EBITDA to net income (challenge)",
                "Start from management's adjusted EBITDA and take away each excluded item. The subtotals must equal "
                "the reported income statement lines (Difference = 0).")
    cols = ["Q3-2026", "9M-2026", "FY2025"]             # the periods shown
    ws.cell(4, 1, "Step").font = BOLD
    for j, p in enumerate(cols, start=2):
        ws.cell(4, j, PERIODS[p][1]).font = BOLD
    for j in range(1, len(cols) + 2):
        ws.cell(4, j).fill = HEAD_FILL
    ws.cell(5, 1, "Adjusted EBITDA (non-GAAP)").font = BOLD
    for j, p in enumerate(cols, start=2):
        c = ws.cell(5, j, "=" + cell[(p, "adj_ebitda")])  # start value
        c.font, c.number_format = BLACK, MONEY
    row, total_row, items_start = 6, 5, 6               # current row, last subtotal row, first item after it
    for label, metric, how in BRIDGE_STEPS:
        if metric is None:                              # subtotal + comparison with the filing
            ws.cell(row, 1, f"Subtotal: {label.lstrip('= ')} (calculated)").font = BOLD  # no leading "=": Excel
                                                                                         # would read it as a formula
            for j, p in enumerate(cols, start=2):
                col = get_column_letter(j)
                c = ws.cell(row, j, f"={col}{total_row}+SUM({col}{items_start}:{col}{row - 1})")  # last subtotal + items
                c.font, c.number_format = BOLD, MONEY
            ws.cell(row + 1, 1, "   reported in the income statement").font = NOTE
            ws.cell(row + 2, 1, "   difference (should be 0)").font = NOTE
            for j, p in enumerate(cols, start=2):
                col = get_column_letter(j)
                c = ws.cell(row + 1, j, "=" + cell[(p, how)])
                c.font, c.number_format = BLACK, MONEY
                c = ws.cell(row + 2, j, f"={col}{row}-{col}{row + 1}")
                c.font, c.number_format = BLACK, MONEY
            total_row, items_start = row, row + 3       # the next block starts from this subtotal (skip 2 check rows)
            row += 3
            continue
        ws.cell(row, 1, ("less " if how < 0 else "plus ") + label).font = BLACK
        for j, p in enumerate(cols, start=2):
            ref = cell.get((p, metric))                 # None when the line did not exist that period
            c = ws.cell(row, j, ("=" + ("-" if how < 0 else "+") + ref) if ref else 0)
            c.font, c.number_format = BLACK if ref else BLUE, MONEY
            if not ref:
                c.comment = Comment("This line did not exist in this period's reconciliation, so it is 0.", "note")
        row += 1
    ws.column_dimensions["A"].width = 46
    for j in range(2, len(cols) + 2):
        ws.column_dimensions[get_column_letter(j)].width = 15
    # returns nothing; adds the EBITDA Bridge sheet
    return None


def write_checks_sheet(wb, checks, cell):
    """Every check as live formulas: calculated, reported, difference, PASS/FAIL."""
    ws = wb.create_sheet("Checks", 1)                   # second tab, right after About
    sheet_title(ws, f"{COMPANY} - Checks", "Each row adds up lines from the statement sheets (Calculated) and "
                "compares them with the line the filing reports. FAIL rows are red; the Note explains why.")
    ws["A3"] = "Tolerance ($m) for rounding:"
    ws["A3"].font = BOLD
    ws["C3"] = TOLERANCE                                # the one input of this sheet
    ws["C3"].font = BLUE
    ws["E3"] = "Checks"
    ws["F3"] = f'=COUNTA(H7:H{len(checks) + 6})'
    ws["E4"] = "FAIL"
    ws["F4"] = f'=COUNTIF(H7:H{len(checks) + 6},"FAIL")'
    ws["E5"] = "PASS (rounding)"
    ws["F5"] = f'=COUNTIF(H7:H{len(checks) + 6},"PASS (rounding)")'
    for a in ["E3", "E4", "E5"]:
        ws[a].font = BOLD
    headers = ["#", "Group", "Check", "Period", "Calculated", "Reported", "Difference", "Status", "Note"]
    for j, h in enumerate(headers, start=1):            # header row
        c = ws.cell(6, j, h)
        c.font, c.fill = BOLD, HEAD_FILL
    for i, ch in enumerate(checks, start=7):            # one row per check
        calc = "=" + "".join(("+" if s > 0 else "-") + cell[(p, m)] for s, p, m in ch["terms"])
        values = [i - 6, ch["group"], ch["name"], PERIODS[ch["period"]][1], calc, "=" + cell[ch["target"]],
                  f"=E{i}-F{i}", f'=IF(ABS(G{i})<0.5,"PASS",IF(ABS(G{i})<=$C$3,"PASS (rounding)","FAIL"))', ch["note"]]
        for j, v in enumerate(values, start=1):
            c = ws.cell(i, j, v)
            c.font = NOTE if j == 9 else BLACK
            if j in (5, 6, 7):
                c.number_format = MONEY
    last = len(checks) + 6                              # last check row
    red = PatternFill("solid", fgColor="FFC7CE")        # Excel's standard "bad" fill
    amber = PatternFill("solid", fgColor="FFEB9C")      # rounding
    ws.conditional_formatting.add(f"H7:H{last}", CellIsRule(operator="equal", formula=['"FAIL"'], fill=red,
                                                            font=Font(name=ARIAL, bold=True, color="9C0006")))
    ws.conditional_formatting.add(f"H7:H{last}", CellIsRule(operator="equal", formula=['"PASS (rounding)"'], fill=amber))
    widths = [5, 36, 62, 16, 13, 13, 12, 16, 70]
    for j, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = "A7"
    # returns nothing; adds the Checks sheet (4 formulas per check)
    return None


def write_sources_sheet(wb, rows):
    """The documents, then one row per number with its document and page."""
    ws = wb.create_sheet("Sources")
    sheet_title(ws, f"{COMPANY} - Sources", "Every number in this workbook, where it comes from. The full list is "
                "data/spacex_data.csv; filings/SOURCES.md describes each document.")
    docs = [("s1a.pdf", "Form S-1/A preliminary prospectus, 'subject to completion, dated June 3, 2026'",
             "Audited years 2023-2025; unaudited Q1 2026 and Q1 2025"),
            ("q2-2026-release.pdf", "Q2 2026 earnings press release (corrected version)", "Q2 2026, Q2 2025, Jan-Jun"),
            ("q3-2026-release.pdf", "Q3 2026 earnings press release", "Q3 2026, Q3 2025, Jan-Sep"),
            ("q3-2026-transcript.txt", "Q3 2026 earnings call transcript (third-party copy, not published)",
             "Management commentary and guidance only")]
    for j, h in enumerate(["File", "What it is", "Periods"], start=1):
        c = ws.cell(4, j, h)
        c.font, c.fill = BOLD, HEAD_FILL
    for i, d in enumerate(docs, start=5):
        for j, v in enumerate(d, start=1):
            ws.cell(i, j, v).font = BLACK
    r = 5 + len(docs) + 1                               # leave a blank row
    ws.cell(r, 1, "Figures not in these files, and other notes").font = BOLD
    for k, text in enumerate(DATA_GAPS + DISCREPANCIES, start=1):
        ws.cell(r + k, 1, f"{k}. {text}").font = NOTE
    r = r + len(DATA_GAPS) + len(DISCREPANCIES) + 2     # the per-number table starts here
    cols = ["period", "metric", "value", "unit", "statement", "source", "page", "page_label", "kind", "note"]
    for j, h in enumerate(cols, start=1):
        c = ws.cell(r, j, h)
        c.font, c.fill = BOLD, HEAD_FILL
    for i, row in enumerate(rows, start=r + 1):
        for j, h in enumerate(cols, start=1):
            v = row[h]
            if h == "page" and v:                       # page numbers as numbers, so they sort properly
                v = int(v)
            c = ws.cell(i, j, v)
            c.font = BLUE if h == "value" else BLACK
            if h == "value":
                c.number_format = '#,##0.##;(#,##0.##);"-"'
    for j, w in enumerate([22, 34, 12, 8, 26, 24, 7, 10, 22, 70], start=1):
        ws.column_dimensions[get_column_letter(j)].width = w
    # returns nothing; adds the Sources sheet (4 documents, notes, 658 number rows)
    return None


def write_other_sheet(wb, rows):
    """Facts that are not part of a statement, and management guidance (clearly labelled)."""
    ws = wb.create_sheet("Other Facts")
    sheet_title(ws, f"{COMPANY} - Other facts and management commentary",
                "Top: facts printed in the releases. Bottom: what management SAID on the Q3 call (transcript, not a "
                "filing, not audited).")
    heads = ["Item", "Value", "Unit", "Date / period", "Source", "Page", "Note"]
    r = 4
    for kind, title in [("filing", "From the filings"), ("management commentary", "Management commentary (Q3 2026 call)")]:
        ws.cell(r, 1, title).font = BOLD
        r += 1
        for j, h in enumerate(heads, start=1):
            c = ws.cell(r, j, h)
            c.font, c.fill = BOLD, HEAD_FILL
        r += 1
        for row in rows:
            if row["statement"] not in ("other", "guidance") or row["kind"] != kind:
                continue
            vals = [LABELS[row["metric"]], row["value"], row["unit"], row["period"], row["source"], row["page"], row["note"]]
            for j, v in enumerate(vals, start=1):
                c = ws.cell(r, j, v)
                c.font = BLUE if j == 2 else (NOTE if j == 7 else BLACK)
                if j == 2:
                    c.number_format = '#,##0.0#'
            r += 1
        r += 1
    for j, w in enumerate([48, 12, 8, 14, 24, 7, 80], start=1):
        ws.column_dimensions[get_column_letter(j)].width = w
    # returns nothing; adds the Other Facts sheet
    return None


def write_about_sheet(wb, h):
    """First tab: what the workbook is and how to read it."""
    ws = wb.active                                      # the default first sheet
    ws.title = "About"
    lines = [
        (f"{COMPANY} ({TICKER}) - basic accounting analysis", TITLE),
        ("Personal learning project, not investment advice. Facts, ratios and questions only: no buy/sell view, "
         "price target or valuation.", Font(name=ARIAL, bold=True, color="9C0006")),
        ("", BLACK),
        ("How this workbook is built", BOLD),
        ("Every number comes from data/spacex_data.csv, typed in from the filings with the document and page of each "
         "number. scripts/build_outputs.py writes this workbook, so do not edit it by hand: edit the CSV and re-run.", BLACK),
        ("Colours: blue = typed in from a filing (hover the cell for its source page); black = an Excel formula.", BLACK),
        ("Empty cell = the figure is not in these files (never estimated). Brackets = negative. Amounts in US$ millions.", BLACK),
        ("Signs follow the filings: interest expense is negative; a tax provision is positive (a charge), a tax "
         "benefit negative.", BLACK),
        ("", BLACK),
        ("Tabs", BOLD),
        ("Checks - does everything add up? Red rows failed; the Note column says why.", BLACK),
        ("Income Statement, Segments, Balance Sheet, Cash Flow - the numbers as reported, one column per period.", BLACK),
        ("Ratios - six key ratios, costs as % of revenue, growth, cash flow and debt, all as formulas.", BLACK),
        ("EBITDA Bridge - the challenge: from adjusted EBITDA to net income, step by step.", BLACK),
        ("Other Facts - backlog, IPO, bond, Cursor; plus management guidance from the call (labelled).", BLACK),
        ("Sources - the documents, the figures that are missing, and every number's page.", BLACK),
        ("", BLACK),
        (f"Checks: {h['checks']} run, {h['checks_fail']} fail, {h['checks_rounding']} pass only within "
         f"${TOLERANCE}m rounding (when this file was built).", BLACK),
        ("Period codes: FY = full year; Q1-Q3 = quarters; 6M = January-June; 9M = January-September.", BLACK),
    ]
    for i, (text, font) in enumerate(lines, start=1):
        c = ws.cell(i, 1, text)
        c.font = font
        c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.column_dimensions["A"].width = 130
    # returns nothing; fills the About sheet
    return None


def build_workbook(rows, data, source, checks, h):
    """Write data/spacex_accounting.xlsx: inputs, ratios, bridge, checks and sources."""
    wb = Workbook()                                     # empty workbook with one sheet
    write_about_sheet(wb, h)                            # tab 1
    cell = {}                                           # (period, metric) -> "'Sheet'!C6", filled by the input sheets
    for name, layout in INPUT_LAYOUT.items():           # the four statement sheets
        write_input_sheet(wb, name, layout, data, source, cell)
    write_ratios_sheet(wb, data, cell)                  # formulas that read the statements
    write_bridge_sheet(wb, data, cell)
    write_checks_sheet(wb, checks, cell)                # inserted as the second tab
    write_other_sheet(wb, rows)
    write_sources_sheet(wb, rows)
    wb.calculation.fullCalcOnLoad = True                # openpyxl stores no results: ask Excel to calculate on open
    wb.save(XLSX_PATH)                                  # write the file
    # returns nothing; writes data/spacex_accounting.xlsx (9 tabs)
    return None


def print_summary(h, checks):
    """Print the headline numbers and every failed check in the terminal."""
    print(f"Revenue {money(h['rev_fy23'])} (2023) -> {money(h['rev_fy25'])} (2025); Jan-Sep 2026 {money(h['rev_9m26'])}")
    print(f"Q3 2026 revenue {money(h['rev_q3_26'])} ({pct(h['growth_q3'], 0)} y/y), net income {money(h['ni_q3_26'])}")
    print(f"Jan-Sep 2026: CFO {money(h['cfo_9m26'])}, capex {money(h['capex_9m26'])}, FCF {money(h['fcf_9m26'])}")
    print(f"30 Sep 2026: net cash {money(h['net_sep26'])}")
    print(f"Checks: {h['checks']} run, {h['checks_fail']} FAIL, {h['checks_rounding']} PASS (rounding)")
    for c in checks:                                    # list every check that is not a clean pass
        if c["status"] != "PASS":
            print(f"  {c['status']:<16} {c['period']:<10} {c['name']}  (diff {c['diff']:+,.0f})")
    # returns nothing; prints ~8 lines, e.g.
    # Revenue $10.4bn (2023) -> $18.7bn (2025); Jan-Sep 2026 $27.1bn
    # Checks: 233 run, 2 FAIL, 9 PASS (rounding)
    return None


def main():
    rows, data, source = load_data()                    # step 1: read the one data file
    tables = ratio_tables(data)                         # step 2: ratios, cost shares, growth
    checks = build_checks(data)                         # step 3: does everything add up?
    h = headline(data, tables, checks)                  # step 4: the numbers quoted in the text
    print_summary(h, checks)                            # step 5: show them in the terminal
    build_workbook(rows, data, source, checks, h)       # step 6: the Excel workbook
    build_page(data, tables, checks, h)                 # step 7: the web page (code below)
    update_readme(h, checks)                            # step 8: the README key-numbers block
    print(f"Saved {XLSX_PATH.name}, {HTML_PATH.name} and the README block")


# ===========================================================================
# html / chart code
# ===========================================================================

CSS = """
:root{color-scheme:light dark;--page:#fbfaf7;--surface:#ffffff;--ink:#0b0b0b;--ink-2:#52514e;--muted:#6b6a65;--grid:#e1e0d9;
  --axis:#c3c2b7;--s1:#2a78d6;--s2:#eb6834;--s3:#1baf7a;--s4:#eda100;--s5:#e87ba4;--accent:#2a78d6;--warn:#9c4a00;--warn-bg:#fff4e5}
@media (prefers-color-scheme:dark){:root:not([data-theme=light]){--page:#0d0d0d;--surface:#1a1a19;--ink:#f0efec;--ink-2:#c3c2b7;
  --muted:#9a988f;--grid:#2c2c2a;--axis:#4a4a46;--s1:#3987e5;--s2:#d95926;--s3:#199e70;--s4:#c98500;--s5:#d55181;--accent:#3987e5;
  --warn:#ffb86b;--warn-bg:#2a1f12}}
:root[data-theme=dark]{--page:#0d0d0d;--surface:#1a1a19;--ink:#f0efec;--ink-2:#c3c2b7;--muted:#9a988f;--grid:#2c2c2a;--axis:#4a4a46;
  --s1:#3987e5;--s2:#d95926;--s3:#199e70;--s4:#c98500;--s5:#d55181;--accent:#3987e5;--warn:#ffb86b;--warn-bg:#2a1f12}
*{box-sizing:border-box}
body{margin:0;background:var(--page);color:var(--ink);font:15px/1.55 system-ui,-apple-system,Segoe UI,sans-serif}
main{max-width:1000px;margin:0 auto;padding:16px max(16px,env(safe-area-inset-right)) 40px max(16px,env(safe-area-inset-left))}
h1{font-size:1.5rem;line-height:1.25;margin:.4rem 0}
h2{font-size:1.15rem;line-height:1.3;margin:2.6rem 0 .5rem;padding-top:.8rem;border-top:1px solid var(--grid)}
h3{font-size:1rem;margin:1.2rem 0 .3rem}
p{margin:.4rem 0 .8rem}
.sub{color:var(--ink-2)}
a{color:var(--accent)}
a.back{display:inline-flex;align-items:center;min-height:44px;text-decoration:none}
.disclaimer{background:var(--warn-bg);color:var(--warn);border:1px solid currentColor;border-radius:8px;padding:10px 14px;margin:12px 0;font-weight:600}
.tiles{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:10px;margin:12px 0}
.tile{background:var(--surface);border:1px solid var(--grid);border-radius:10px;padding:10px 12px}
.tile b{display:block;font-size:1.35rem;font-variant-numeric:tabular-nums}
.tile span{color:var(--muted);font-size:.85rem}
.read{background:var(--surface);border-left:4px solid var(--accent);border-radius:6px;padding:10px 14px;margin:12px 0}
.q{font-weight:600;color:var(--ink-2);margin:.2rem 0 .4rem}
.mgmt{background:var(--surface);border:1px dashed var(--axis);border-radius:8px;padding:8px 12px;margin:12px 0;font-size:.93rem}
.mgmt b.tag{display:inline-block;font-size:.75rem;text-transform:uppercase;letter-spacing:.04em;color:var(--muted);margin-right:6px}
.chart{width:100%;height:400px;background:var(--surface);border:1px solid var(--grid);border-radius:10px}
.chart.tall{height:460px}
.note{color:var(--muted);font-size:.85rem}
.seg{display:inline-flex;border:1px solid var(--axis);border-radius:8px;overflow:hidden;margin:4px 0 8px}
.seg button{min-height:44px;min-width:88px;padding:0 14px;border:0;background:var(--surface);color:var(--ink-2);font:inherit;cursor:pointer}
.seg button[aria-pressed=true]{background:var(--accent);color:#fff;font-weight:600}
details{margin:.8rem 0}
summary{min-height:44px;display:flex;align-items:center;cursor:pointer;color:var(--ink-2);font-weight:600}
.scroll{overflow-x:auto;border:1px solid var(--grid);border-radius:10px;margin:.4rem 0}
table{border-collapse:collapse;width:100%;background:var(--surface);font-variant-numeric:tabular-nums;font-size:.9rem}
th,td{padding:7px 9px;border-bottom:1px solid var(--grid);text-align:right;white-space:nowrap}
th{color:var(--ink-2);font-weight:600}
th:first-child,td:first-child{text-align:left;position:sticky;left:0;background:var(--surface);white-space:normal;min-width:150px}
tr.total td{font-weight:700}
td.na{color:var(--muted)}
.ratio-plain{display:block;font-weight:400;color:var(--muted);font-size:.82rem;max-width:300px}
.challenge{border:2px solid var(--s2);border-radius:12px;padding:4px 16px 12px;margin-top:2.4rem}
.challenge h2{border-top:0;margin-top:1rem}
.badge{display:inline-block;background:var(--s2);color:#fff;border-radius:6px;padding:2px 8px;font-size:.8rem;font-weight:700;letter-spacing:.03em}
ol.steps li,ol.qs li,ul.plain li{margin:.45rem 0}
dl.gloss{display:grid;grid-template-columns:200px 1fr;gap:6px 16px;margin:.5rem 0}
dl.gloss dt{font-weight:600}
dl.gloss dd{margin:0;color:var(--ink-2)}
.hint-touch{display:none}
.swipe{display:none;margin:.2rem 0}
@media (pointer:coarse){.hint-mouse{display:none}.hint-touch{display:inline}}
footer{margin-top:2.5rem;padding-top:1rem;border-top:1px solid var(--grid);color:var(--muted);font-size:.85rem}
@media (max-width:640px){.swipe{display:block}h1{font-size:1.28rem}.chart{height:470px}.chart.tall{height:560px}th,td{padding:7px 6px}
  dl.gloss{grid-template-columns:1fr}dl.gloss dd{margin-bottom:.5rem}.seg{display:flex}.seg button{flex:1}}
"""

JS = """
const DATA = __DATA__;                                          // everything Python computed, as JSON
const narrowMQ = window.matchMedia("(max-width: 640px)");       // true on a phone-width screen
const isTouch = window.matchMedia("(pointer: coarse)").matches; // a finger instead of a mouse
const css = name => getComputedStyle(document.documentElement).getPropertyValue(name).trim();  // read a colour token
const INSTANCES = {};                                           // chart div id -> ECharts instance
const MODES = {revenue: "q", margins: "q"};                     // toggles: q = quarters, y = years
const B = v => v == null ? "n/a" : (v < 0 ? "-$" : "$") + Math.abs(v).toFixed(1) + "bn";  // 18.67 -> "$18.7bn"
const P = v => v == null ? "n/a" : v.toFixed(1) + "%";          // 49.4 -> "49.4%"
const M = v => (v < 0 ? "-$" : "$") + Math.abs(v).toLocaleString("en-US") + "m";  // 8809 -> "$8,809m"

function base(narrow, n) {                                      // options shared by every chart; n = legend entries
  return {
    animation: false, backgroundColor: "transparent", textStyle: {color: css("--ink-2")},
    legend: narrow ? {bottom: 6, left: 10, orient: "vertical", itemGap: 6, textStyle: {color: css("--ink-2")}}  // below the plot
                   : {top: 6, left: "center", textStyle: {color: css("--ink-2")}},   // above the plot, never on top of it
    tooltip: {trigger: "axis", triggerOn: isTouch ? "click" : "mousemove|click", confine: true,
              axisPointer: {type: "shadow"}, valueFormatter: v => v},
  };
}
const grid = (narrow, n) => narrow ? {left: 52, right: 14, top: 22, bottom: 34 + 22 * n} : {left: 64, right: 24, top: 48, bottom: 34};
const ax = extra => Object.assign({axisLine: {lineStyle: {color: css("--axis")}}, splitLine: {lineStyle: {color: css("--grid")}},
                                   axisLabel: {color: css("--muted"), hideOverlap: true}}, extra);
const color = i => css("--s" + (i + 1));                        // categorical slot i (0-based), fixed order
const gap = () => ({borderColor: css("--surface"), borderWidth: 1}); // 2px surface gap between stacked fills

function revenue(narrow) {                                      // revenue by segment, stacked bars
  const d = DATA.rev[MODES.revenue], n = DATA.segments.length, o = base(narrow, n);
  o.grid = grid(narrow, n);
  o.xAxis = ax({type: "category", data: d.labels, axisTick: {show: false}, axisLabel: {color: css("--muted"), interval: 0, fontSize: narrow ? 10 : 12}});
  o.yAxis = ax({type: "value", axisLabel: {color: css("--muted"), formatter: v => "$" + v + "bn"}});
  o.series = DATA.segments.map((s, i) => ({name: s, type: "bar", stack: "rev", barMaxWidth: 56,
    data: d.values[i], itemStyle: Object.assign({color: color(i)}, gap())}));
  o.series[n - 1].label = {show: true, position: "top", color: css("--ink-2"), fontSize: narrow ? 10 : 12,   // total above each bar
    formatter: p => d.total[p.dataIndex] == null ? "" : B(d.total[p.dataIndex])};
  o.tooltip.formatter = ps => { const i = ps[0].dataIndex;
    if (d.total[i] == null) return d.labels[i] + ": not reported on its own in these files";
    return "<b>" + d.labels[i] + "</b><br>" + ps.slice().reverse().map(p => p.marker + p.seriesName + ": " + B(p.value)).join("<br>") +
           "<br><b>Total: " + B(d.total[i]) + "</b>"; };
  return o;
}

function costs(narrow) {                                        // each cost as % of revenue, stacked
  const d = DATA.costs, n = d.names.length, o = base(narrow, n);
  o.grid = grid(narrow, n);
  o.xAxis = ax({type: "category", data: d.labels, axisTick: {show: false}, axisLabel: {color: css("--muted"), interval: 0}});
  o.yAxis = ax({type: "value", axisLabel: {color: css("--muted"), formatter: "{value}%"}});
  o.series = d.names.map((s, i) => ({name: s, type: "bar", stack: "c", barMaxWidth: 56, data: d.values[i],
    itemStyle: Object.assign({color: color(i)}, gap())}));
  o.series[0].markLine = {silent: true, symbol: "none", lineStyle: {color: css("--ink"), type: "dashed", width: 1.5},
    label: {formatter: "revenue = 100%", color: css("--ink"), position: "insideStartTop", backgroundColor: css("--surface"),
            padding: [1, 4], borderRadius: 3}, data: [{yAxis: 100}]};
  o.series[n - 1].label = {show: true, position: "top", color: css("--ink-2"), fontSize: narrow ? 10 : 12,
    formatter: p => Math.round(d.total[p.dataIndex]) + "%"};
  o.tooltip.formatter = ps => "<b>" + d.labels[ps[0].dataIndex] + "</b> (% of revenue)<br>" +
    ps.slice().reverse().map(p => p.marker + p.seriesName + ": " + P(p.value)).join("<br>") + "<br><b>All: " + P(d.total[ps[0].dataIndex]) + "</b>";
  return o;
}

function margins(narrow) {                                      // gross, operating and net margin, lines
  const d = DATA.margins[MODES.margins], n = d.names.length, o = base(narrow, n);
  o.grid = grid(narrow, n);
  o.tooltip.axisPointer = {type: "line"};
  o.xAxis = ax({type: "category", data: d.labels, boundaryGap: false, axisTick: {show: false},
               axisLabel: {color: css("--muted"), interval: 0, fontSize: narrow ? 10 : 12}});
  o.yAxis = ax({type: "value", axisLabel: {color: css("--muted"), formatter: "{value}%"}});
  o.series = d.names.map((s, i) => ({name: s, type: "line", data: d.values[i], symbol: "circle", symbolSize: 8,
    connectNulls: false, lineStyle: {width: 2, color: color(i)}, itemStyle: {color: color(i), borderColor: css("--surface"), borderWidth: 2}}));
  o.series[0].markLine = {silent: true, symbol: "none", lineStyle: {color: css("--axis"), width: 1}, label: {show: false}, data: [{yAxis: 0}]};
  o.tooltip.formatter = ps => "<b>" + d.labels[ps[0].dataIndex] + "</b><br>" +
    (ps[0].value == null ? "not reported on its own in these files" : ps.map(p => p.marker + p.seriesName + ": " + P(p.value)).join("<br>"));
  return o;
}

function cash(narrow) {                                         // operating cash in (up), capex out (down), free cash flow
  const d = DATA.cash, o = base(narrow, 3);
  o.grid = grid(narrow, 3);
  o.xAxis = ax({type: "category", data: d.labels, axisTick: {show: false}, axisLabel: {color: css("--muted"), interval: 0}});
  o.yAxis = ax({type: "value", axisLabel: {color: css("--muted"), formatter: v => (v < 0 ? "-$" : "$") + Math.abs(v) + "bn"}});
  o.series = [
    {name: "Operating cash flow (in)", type: "bar", stack: "f", barMaxWidth: 56, data: d.cfo, itemStyle: Object.assign({color: color(0), borderRadius: [4, 4, 0, 0]}, gap())},
    {name: "Capex (out)", type: "bar", stack: "f", barMaxWidth: 56, data: d.capex, itemStyle: Object.assign({color: color(1), borderRadius: [0, 0, 4, 4]}, gap())},
    {name: "Free cash flow (in - out)", type: "scatter", symbol: "diamond", symbolSize: 14,
     data: d.fcf.map((v, i) => ({value: v, label: {position: i === d.fcf.length - 1 ? "left" : "right"}})),  // last label inside the plot
     itemStyle: {color: css("--ink"), borderColor: css("--surface"), borderWidth: 2},
     label: {show: true, position: "right", color: css("--ink"), fontSize: narrow ? 10 : 12, formatter: p => B(p.value),
             backgroundColor: css("--surface"), padding: [1, 3], borderRadius: 3}},   // readable on top of a bar
  ];
  o.tooltip.formatter = ps => "<b>" + d.labels[ps[0].dataIndex] + "</b><br>" + ps.map(p => p.marker + p.seriesName + ": " + B(p.value)).join("<br>");
  return o;
}

function debt(narrow) {                                         // cash & securities vs debt at each balance sheet date
  const d = DATA.debt, o = base(narrow, 3);
  o.grid = grid(narrow, 3);
  o.xAxis = ax({type: "category", data: d.labels, axisTick: {show: false}, axisLabel: {color: css("--muted"), interval: 0}});
  o.yAxis = ax({type: "value", axisLabel: {color: css("--muted"), formatter: v => (v < 0 ? "-$" : "$") + Math.abs(v) + "bn"}});
  o.series = [
    {name: "Cash + marketable securities", type: "bar", barMaxWidth: 40, barGap: "10%", data: d.cash, itemStyle: {color: color(0), borderRadius: [4, 4, 0, 0]}},
    {name: "Debt (incl. finance leases)", type: "bar", barMaxWidth: 40, data: d.debt, itemStyle: {color: color(1), borderRadius: [4, 4, 0, 0]}},
    {name: "Net cash (cash - debt)", type: "scatter", symbol: "diamond", symbolSize: 14, data: d.net,
     itemStyle: {color: css("--ink"), borderColor: css("--surface"), borderWidth: 2},
     label: {show: !narrow, position: "right", color: css("--ink"), formatter: p => B(p.value),
             backgroundColor: css("--surface"), padding: [1, 3], borderRadius: 3}},   // readable on top of a bar
  ];
  o.tooltip.formatter = ps => "<b>" + d.labels[ps[0].dataIndex] + "</b><br>" + ps.map(p => p.marker + p.seriesName + ": " + B(p.value)).join("<br>");
  return o;
}

function bridge(narrow) {                                       // waterfall: adjusted EBITDA -> net income (horizontal)
  const d = DATA.bridge, o = base(narrow, 0);
  o.legend = {show: false};                                     // colours are explained in the text and the tooltip
  o.grid = {left: narrow ? 4 : 16, right: narrow ? 62 : 90, top: 10, bottom: narrow ? 10 : 30, containLabel: true};
  o.yAxis = ax({type: "category", data: d.labels, inverse: true, axisTick: {show: false},
               axisLabel: {color: css("--ink-2"), fontSize: narrow ? 10 : 12, width: narrow ? 100 : 230, overflow: "break"}});
  o.xAxis = ax({type: "value", axisLabel: {show: !narrow, color: css("--muted"), formatter: v => "$" + (v / 1000) + "bn"}});
  const kindColor = (k, v) => k === "total" ? css("--ink-2") : v < 0 ? color(1) : color(0);
  o.series = [
    {type: "bar", stack: "w", silent: true, data: d.base, itemStyle: {color: "transparent"}, tooltip: {show: false}},
    {type: "bar", stack: "w", barMaxWidth: 26, data: d.size.map((v, i) => ({value: v,
       itemStyle: {color: kindColor(d.kind[i], d.delta[i]), borderRadius: 3}})),
     label: {show: true, position: "right", color: css("--ink"), fontSize: narrow ? 10 : 12,
             formatter: p => (d.kind[p.dataIndex] === "total" ? "" : (d.delta[p.dataIndex] >= 0 ? "+" : "")) + M(d.delta[p.dataIndex])}},
  ];
  o.tooltip = {trigger: "item", triggerOn: isTouch ? "click" : "mousemove|click", confine: true,
    formatter: p => { const i = p.dataIndex; return "<b>" + d.labels[i] + "</b><br>" + (d.kind[i] === "total" ? "Subtotal: " + M(d.delta[i]) :
      (d.delta[i] >= 0 ? "Adds " : "Takes away ") + M(Math.abs(d.delta[i])) + "<br>Running total: " + M(d.total[i])); }};
  return o;
}

const BUILDERS = {revenue, costs, margins, cash, debt, bridge};  // chart div id -> function that builds its options
function draw(id) {                                             // (re)build one chart
  const div = document.getElementById(id);
  if (!INSTANCES[id]) INSTANCES[id] = echarts.init(div);
  INSTANCES[id].setOption(BUILDERS[id](narrowMQ.matches), true);
  INSTANCES[id].resize();
}
function drawAll() { Object.keys(BUILDERS).forEach(draw); }     // on load, rotation and theme change
document.querySelectorAll(".seg button").forEach(b => b.addEventListener("click", () => {   // Quarters / Years toggles
  MODES[b.dataset.chart] = b.dataset.mode;
  b.parentNode.querySelectorAll("button").forEach(x => x.setAttribute("aria-pressed", String(x === b)));
  draw(b.dataset.chart);
}));
narrowMQ.addEventListener("change", drawAll);                   // rebuild when a phone is rotated
window.matchMedia("(prefers-color-scheme: dark)").addEventListener("change", drawAll);  // rebuild in the new colours
window.addEventListener("resize", () => Object.values(INSTANCES).forEach(c => c.resize()));
drawAll();
"""

GLOSSARY = [                                             # (term, one plain sentence)
    ("Revenue", "Money earned from customers for what the company delivered in the period (not necessarily cash received yet)."),
    ("Segment", "A business unit the company reports on separately: here Space (rockets), Connectivity (Starlink) and AI / SI."),
    ("Cost of revenue", "The direct cost of delivering what was sold, such as hardware, launch costs and data-centre running costs."),
    ("Gross profit / margin", "Revenue minus cost of revenue; as a % of revenue it is the gross margin."),
    ("R&D", "Research and development: spending on new products such as Starship and new AI models."),
    ("SG&A", "Selling, general and administrative costs: sales, marketing, offices, management and support staff."),
    ("Operating income", "Revenue minus all operating costs: what the core business earned before interest and tax."),
    ("Net income", "The bottom line: profit (or loss) after every cost, interest, other gains and losses, and tax."),
    ("Depreciation & amortization", "Spreading the cost of long-lived assets (rockets, satellites, computers) over the years they are used; a cost that uses no cash that year."),
    ("Share-based compensation", "Paying employees in shares or options; it is a real cost to shareholders but not a cash payment."),
    ("Operating cash flow", "Cash actually collected from customers minus cash paid to run the business in the period."),
    ("Capex", "Capital expenditures: cash spent on long-lived assets such as data centres, satellites and launch pads."),
    ("Free cash flow", "Operating cash flow minus capex: the cash left over (or missing) after investing."),
    ("Deferred revenue", "Cash received from customers for work not done yet; it becomes revenue later."),
    ("Marketable securities", "Short-term investments (like Treasury bills) that can be sold quickly for cash."),
    ("Debt and finance leases", "Borrowed money plus long leases that work like loans (the company effectively owns the asset)."),
    ("Shareholders' equity", "What would be left for shareholders if all liabilities were paid: assets minus liabilities (and minus preferred stock shown outside equity)."),
    ("Redeemable convertible preferred stock", "Shares sold to early investors that could be bought back or converted; shown between liabilities and equity until the IPO converted them."),
    ("Goodwill", "The part of an acquisition price above the fair value of what was bought; it is written down (impaired) if the deal disappoints."),
    ("Impairment", "A write-down when an asset is worth less than its book value (for example the Twitter brand in 2023)."),
    ("Adjusted EBITDA (non-GAAP)", "Earnings before interest, tax, depreciation and amortization, also excluding items management picks; not defined by GAAP, the official US accounting rules."),
    ("Backlog", "Signed contracts not yet delivered and not yet counted as revenue."),
    ("YTD / 9M", "Year-to-date; 9M = the first nine months (January-September)."),
]


def table(headers, rows, total_rows=()):
    """A simple HTML table inside a sideways-scrolling box; the first column stays pinned."""
    head = "".join(f"<th>{escape(h)}</th>" for h in headers)
    body = "".join(
        f'<tr{" class=total" if i in total_rows else ""}>' + "".join(
            f'<td{" class=na" if c in ("n/a", "") else ""}>{c if str(c).startswith("<") else escape(str(c))}</td>' for c in r)
        + "</tr>" for i, r in enumerate(rows))
    # returns a str of HTML, e.g. '<div class="scroll"><table><tr><th>Segment</th>...</tr><tr><td>Space</td>...</tr></table></div>'
    hint = '<p class="note swipe">Swipe the table sideways to see every column.</p>' if len(headers) > 3 else ""
    return f'{hint}<div class="scroll"><table><tr>{head}</tr>{body}</table></div>'


def acct(v):
    """$ millions in accounting style for tables: 18674 -> '18,674', -4937 -> '(4,937)', None -> 'n/a'."""
    # returns a str, e.g. acct(-4937.0) -> '(4,937)'
    return "n/a" if v is None else f"({abs(v):,.0f})" if v < 0 else f"{v:,.0f}"


def fmt_ratio(v, kind):
    """A ratio for tables: fractions as %, multiples as 1.45x, None as an empty cell."""
    # returns a str, e.g. fmt_ratio(0.4939, 'pct') -> '49.4%';  fmt_ratio(1.446, 'x') -> '1.45x';  fmt_ratio(None, 'x') -> ''
    return "" if v is None else pct(v) if kind == "pct" else f"{v:.2f}x"


def tile(value, label):
    """One small stat box."""
    # returns a str of HTML, e.g. '<div class="tile"><b>$27.1bn</b><span>Revenue, Jan-Sep 2026</span></div>'
    return f'<div class="tile"><b>{escape(value)}</b><span>{escape(label)}</span></div>'


def mgmt(text):
    """A box that marks management commentary from the call transcript (not a filing)."""
    # returns a str of HTML with a visible label
    return f'<div class="mgmt"><b class="tag">Management says (Q3 call, not a filing)</b>{text}</div>'


def build_page(data, tables, checks, h):
    """Assemble the standalone HTML page and write it to disk."""
    g = lambda p, m: get(data, p, m)                    # short name for lookups
    t = tables                                          # short name
    seg_share = lambda p, s: g(p, f"seg_revenue_{s}") / g(p, "revenue")  # a segment's share of revenue
    bridge_steps = [s for s in ebitda_bridge(data, "Q3-2026")  # the challenge waterfall, without
                    if s["kind"] == "total" or s["delta"] != 0]  # zero items (no impairment in Q3)
    wf_base, wf_size = [], []                           # invisible base + visible bar for the waterfall
    for i, s in enumerate(bridge_steps):
        if s["kind"] == "total":                        # totals stand on zero
            wf_base.append(0); wf_size.append(s["total"])
        else:                                           # a step floats between the old and the new running total
            before = s["total"] - s["delta"]
            wf_base.append(min(before, s["total"])); wf_size.append(abs(s["delta"]))
    payload = {                                         # everything the page's JavaScript needs
        "segments": list(SEGMENTS.values()),
        "rev": {"q": revenue_series(data, QUARTERS), "y": revenue_series(data, YEARS)},
        "costs": cost_series(data),
        "margins": {"q": margin_series(data, QUARTERS), "y": margin_series(data, YEARS)},
        "cash": cash_series(data), "debt": debt_series(data),
        "bridge": {"labels": [s["label"] for s in bridge_steps], "kind": [s["kind"] for s in bridge_steps],
                   "delta": [s["delta"] for s in bridge_steps], "total": [s["total"] for s in bridge_steps],
                   "base": wf_base, "size": wf_size},
    }
    m9 = "9M-2026"                                      # the latest year-to-date period
    oi_q3, oth_q3, ii_q3, tax_q3 = g("Q3-2026", "operating_income"), g("Q3-2026", "other_income"), \
        g("Q3-2026", "interest_income"), g("Q3-2026", "income_tax")
    gm = t["gross_margin"]                              # gross margins by column
    cs = {c["id"]: t[c["id"]] for c in COST_SHARES}     # cost shares by column
    capex_mult = t["capex_to_cfo"][m9]                  # capex / CFO in Jan-Sep 2026
    _, _, net_dec25 = debt_cash(data, "2025-12-31")     # net cash before the IPO
    _, debt_mar, _ = debt_cash(data, "2026-03-31")
    _, debt_jun, _ = debt_cash(data, "2026-06-30")
    fails = [c for c in checks if c["status"] == "FAIL"]
    rounding = [c for c in checks if c["status"] == "PASS (rounding)"]

    # --- tables shown under "Details" ---
    rev_cols = YEARS + ["Q1-2025", "Q2-2025", "Q3-2025", "Q1-2026", "Q2-2026", "Q3-2026", "9M-2025", "9M-2026"]
    rev_rows = [[name] + [acct(g(p, f"seg_revenue_{s}")) for p in rev_cols] for s, name in SEGMENTS.items()]
    rev_rows.append(["Total revenue"] + [acct(g(p, "revenue")) for p in rev_cols])
    rev_rows.append(["Growth vs a year earlier"] + [pct(t["revenue_growth"].get(p)) if p in RATIO_COLS and
                    t["revenue_growth"].get(p) is not None else "" for p in rev_cols])
    rev_table = table(["$ millions"] + [PERIODS[p][0] for p in rev_cols], rev_rows, total_rows=(3,))
    cost_rows = [[c["name"]] + [pct(t[c["id"]][p]) for p in COST_COLS] for c in COST_SHARES]
    cost_rows.append(["Revenue ($ millions)"] + [acct(g(p, "revenue")) for p in COST_COLS])
    cost_table = table(["% of revenue"] + [PERIODS[p][0] for p in COST_COLS], cost_rows)
    prof_cols = YEARS + ["Q1-2025", "Q2-2025", "Q3-2025", "Q1-2026", "Q2-2026", "Q3-2026", "9M-2025", "9M-2026"]
    prof_rows = [["Gross profit"] + [acct(g(p, "revenue") - g(p, "cost_of_revenue")) for p in prof_cols],
                 ["Operating income"] + [acct(g(p, "operating_income")) for p in prof_cols],
                 ["Net income"] + [acct(g(p, "net_income")) for p in prof_cols]]
    prof_rows += [[r["name"]] + [pct(evaluate(data, r, p)) for p in prof_cols] for r in RATIOS[:3]]
    prof_table = table(["$ millions"] + [PERIODS[p][0] for p in prof_cols], prof_rows)
    cash_rows = [["Operating cash flow"] + [acct(g(p, "cfo")) for p in CASH_COLS],
                 ["Capex"] + [acct(g(p, "capex_cf")) for p in CASH_COLS],
                 ["Free cash flow"] + [acct(g(p, "cfo") + g(p, "capex_cf")) for p in CASH_COLS],
                 ["Net income (for comparison)"] + [acct(g(p, "net_income")) for p in CASH_COLS]]
    cash_table = table(["$ millions"] + [PERIODS[p][0] for p in CASH_COLS], cash_rows, total_rows=(2,))
    su = sources_and_uses(data, m9)
    su_table = table(["Jan-Sep 2026", "$ millions"], [[label, acct(v)] for label, v in su], total_rows=(len(su) - 1,))
    debt_rows = []
    for d in BS_DATES:
        c_, d_, n_ = debt_cash(data, d)
        debt_rows.append([PERIODS[d][1], acct(c_), acct(d_), acct(n_), acct(g(d, "total_equity"))])
    debt_table = table(["Date", "Cash + securities", "Debt", "Net cash", "Equity"], debt_rows)
    ratio_cols = ["FY2023", "FY2024", "FY2025", "9M-2025", "9M-2026", "Q3-2026"]
    ratio_rows = [[f'<b>{escape(r["name"])}</b><span class="ratio-plain">{escape(r["plain"])}</span>']
                  + [fmt_ratio(t[r["id"]][p], r["fmt"]) for p in ratio_cols] for r in RATIOS]
    ratio_table = table(["Ratio"] + [PERIODS[p][0] for p in ratio_cols], ratio_rows)
    bridge_cols = ["Q3-2026", "9M-2026", "FY2025"]
    bsteps = {p: ebitda_bridge(data, p) for p in bridge_cols}
    bridge_rows = [[s["label"]] + [acct(bsteps[p][i]["delta"]) for p in bridge_cols]
                   for i, s in enumerate(bsteps["Q3-2026"])]
    bridge_table = table(["$ millions"] + [PERIODS[p][1] for p in bridge_cols], bridge_rows,
                         total_rows=tuple(i for i, s in enumerate(bsteps["Q3-2026"]) if s["kind"] == "total"))
    check_groups = {}                                   # group -> [pass, rounding, fail]
    for c in checks:
        k = check_groups.setdefault(c["group"], [0, 0, 0])
        k[["PASS", "PASS (rounding)", "FAIL"].index(c["status"])] += 1
    check_table = table(["Check group", "Pass", "Pass (rounding)", "Fail"],
                        [[grp, v[0], v[1], v[2]] for grp, v in check_groups.items()])
    fail_items = "".join(f"<li><b>{escape(c['name'])}</b>, {escape(PERIODS[c['period']][1])}: off by "
                         f"{c['diff']:+,.0f} ($m). {escape(c['note'])}</li>" for c in fails)
    rounding_items = ", ".join(f"{escape(c['name'].split(': ')[-1])} ({escape(PERIODS[c['period']][1])}, {c['diff']:+,.0f})"
                               for c in rounding)
    gloss = "".join(f"<dt>{escape(a)}</dt><dd>{escape(b)}</dd>" for a, b in GLOSSARY)
    gaps = "".join(f"<li>{escape(x)}</li>" for x in DATA_GAPS)
    discs = "".join(f"<li>{escape(x)}</li>" for x in DISCREPANCIES)
    tiles = "".join([
        tile(money(h["rev_9m26"]), f"Revenue, Jan-Sep 2026 (all of 2025: {money(h['rev_fy25'])})"),
        tile(f"+{h['growth_q3'] * 100:.0f}%", "Revenue growth, Q3 2026 vs Q3 2025"),
        tile(money(h["ni_q3_26"]), f"Net income, Q3 2026 (the only profitable quarter of the {h['quarters']} in these files)"),
        tile(money(h["fcf_9m26"]), "Free cash flow, Jan-Sep 2026 (operating cash flow minus capex)"),
        tile(money(h["net_sep26"]), "Net cash, 30 Sep 2026 (cash and securities minus debt)"),
    ])
    sbc_acq_q3 = g("Q3-2026", "share_based_comp_acquisitions")
    q2gm = gm["Q2-2026"]
    q_op = [t["operating_margin"][q] for q in QUARTERS if t["operating_margin"].get(q) is not None]  # 6 quarters
    cost_total = {p: sum(cs[k][p] for k in cs) for p in COST_COLS}  # height of each cost bar (fraction of revenue)
    below_100 = [PERIODS[p][1] for p in COST_COLS if cost_total[p] <= 1]  # periods where costs < revenue
    rng = lambda lo, hi, unit="bn": (f"${g(*lo) / 1000:g}-{g(*hi) / 1000:g}bn" if unit == "bn"  # guidance range text
                                     else f"{g(*lo):g}-{g(*hi):g} {unit}")
    capex_guide = rng(("Q4-2026", "guidance_capex_low"), ("Q4-2026", "guidance_capex_high"))      # '$35-40bn'
    revenue_guide = rng(("Q1-2027", "guidance_revenue_low"), ("Q1-2027", "guidance_revenue_high"))  # '$33-37bn'
    compute_guide = rng(("2026-12-31", "guidance_nameplate_compute_low"),
                        ("2026-12-31", "guidance_nameplate_compute_high"), "GW")                     # '2.5-3 GW'

    html = f"""<!doctype html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>SpaceX Accounting Basics</title>
<meta name="description" content="A basic, beginner-friendly accounting analysis of SpaceX (SPCX) from its SEC filings: revenue, costs, profit, cash and debt.">
<style>{CSS}</style>
<!-- Apache ECharts, stored locally in lib/ so the page works offline and on GitHub Pages without a CDN -->
<script src="lib/echarts.js"></script>
</head><body><main>
<a class="back" href="../">&larr; Equity Research</a>
<h1>{COMPANY} ({TICKER}): a first look at the accounts</h1>
<p class="sub">A basic accounting analysis built only from SpaceX's SEC documents: the S-1/A prospectus (years 2023-2025
and Q1 2026) and the Q2 and Q3 2026 earnings releases. US dollars; data up to 30 September 2026.</p>
<div class="disclaimer">Personal learning project, not investment advice. This page shows facts, ratios and questions
only: no buy/sell view, no price target, no valuation.</div>

<div class="read"><b>The short version.</b> SpaceX runs three businesses: rockets (Space), Starlink internet
(Connectivity) and artificial intelligence (AI, renamed Super Intelligence or SI in Q3 2026). Revenue is growing very
fast: {money(h['rev_9m26'])} in the first nine months of 2026, already more than all of 2025 ({money(h['rev_fy25'])}).
Q3 2026 was the only quarter with a profit among the {h['quarters']} in these files. But the company spends far more on new equipment than its
operations bring in, and that gap was paid for by its June 2026 stock market listing (IPO) and by borrowing.</div>
<div class="tiles">{tiles}</div>
<p class="note">How to read this page: each section asks a question, answers it in a few sentences, then shows a chart.
Open "Details" for the numbers. Accounting terms are explained the first time they appear and in the
<a href="#glossary">glossary</a>. Negative numbers in tables are in brackets, as in the filings.
<span class="hint-mouse">Hover a chart to see values.</span><span class="hint-touch">Tap a chart to see values.</span></p>

<h2 id="size">1. How big is the business, and where does the revenue come from?</h2>
<div class="read"><b>Revenue</b> (money earned from customers for what was delivered in the period) grew from
{money(h['rev_fy23'])} in 2023 to {money(h['rev_fy25'])} in 2025. 2026 is a step change: {money(h['rev_q3_26'])} in Q3
alone, {pct(h['growth_q3'], 0)} more than in Q3 2025. SpaceX reports three <b>segments</b> (business units reported
separately). In 2025 Connectivity (Starlink) was the biggest, at {pct(seg_share('FY2025', 'connectivity'), 0)} of revenue.
By Q3 2026 AI / SI had become the largest, at {pct(seg_share('Q3-2026', 'ai'), 0)}, driven mainly by cloud services
agreements (renting computing power to other companies), Grok, and Cursor, an AI coding company bought in August 2026
({money(g('Q3-2026', 'cursor_revenue'))} of revenue since then).</div>
<div class="seg" role="group" aria-label="Revenue chart period"><button data-chart="revenue" data-mode="q" aria-pressed="true">Quarters</button><button data-chart="revenue" data-mode="y" aria-pressed="false">Years</button></div>
<div class="chart" id="revenue" role="img" aria-label="Revenue by segment, stacked bars, quarters or years"></div>
<p class="note">* Q4 2025 is not reported on its own in these files (see <a href="#questions">open questions</a>).
The label above each bar is total revenue.</p>
<details><summary>Details: revenue by segment ($ millions)</summary>{rev_table}
<p class="note">Sources: S-1/A segment notes (pages F-57, F-58, F-94, F-95); Q2 and Q3 releases, page 3.</p></details>

<h2 id="money-goes">2. Where does the money go?</h2>
<div class="read">Each bar shows the main costs for every $100 of revenue. <b>Cost of revenue</b> (the direct cost of
delivering what was sold: hardware, launches, data-centre running costs) fell from {pct(cs['cost_of_revenue']['FY2023'], 0)}
of revenue in 2023 to {pct(cs['cost_of_revenue']['Q3-2026'], 0)} in Q3 2026. <b>Research and development (R&amp;D)</b>
(spending on new products such as Starship and AI models) is the cost that grew most: from
{pct(cs['rnd']['FY2023'], 0)} of revenue in 2023 to {pct(cs['rnd']['FY2025'], 0)} in 2025. <b>SG&amp;A</b> (selling,
general and administrative costs: sales, marketing, offices, management) is smaller, about
{pct(cs['sga']['9M-2026'], 0)}. <b>Interest expense</b> (what the company pays its lenders) was
{pct(cs['interest']['FY2025'], 0)} of revenue in 2025. When a bar is taller than the 100% line, these costs were bigger
than revenue. That is true for every period shown except {', '.join(below_100)}, when they fell to
{pct(cost_total['Q3-2026'], 0)} because revenue grew much faster than costs.</div>
<div class="chart tall" id="costs" role="img" aria-label="Costs as a percentage of revenue, stacked bars"></div>
<p class="note">2023 includes a {money(g('FY2023', 'impairment'))} <b>impairment</b> (a write-down because an asset lost
value): the Twitter brand after its rename to X. "9M" = January-September.</p>
<details><summary>Details: costs as % of revenue</summary>{cost_table}
<p class="note">Sources: income statements (S-1/A page F-5; Q3 release page 9).</p></details>

<h2 id="profit">3. Is it profitable, and is that improving?</h2>
<div class="read">Three <b>margins</b> (profit as a share of revenue) tell the story. <b>Gross margin</b> (revenue minus
cost of revenue, as % of revenue) rose from {pct(gm['FY2023'], 0)} in 2023 to {pct(gm['Q3-2026'], 0)} in Q3 2026.
<b>Operating margin</b> (what is left after ALL operating costs, including R&amp;D and SG&amp;A) was negative in
{sum(v < 0 for v in q_op)} of the {len(q_op)} quarters shown and reached {pct(t['operating_margin']['Q3-2026'], 0)} in
Q3 2026, the highest in these files. <b>Net income</b>
(the bottom line, after interest, other gains and losses, and tax) was {money(h['ni_q3_26'])} in Q3 2026, larger than
operating income ({money(oi_q3)}) because of {money(oth_q3)} of other income, {money(ii_q3)} of interest income on the
cash pile, and a {money(-tax_q3)} tax benefit. January-September 2026 as a whole still shows a net loss of
{money(-h['ni_9m26'])}.</div>
<div class="seg" role="group" aria-label="Margin chart period"><button data-chart="margins" data-mode="q" aria-pressed="true">Quarters</button><button data-chart="margins" data-mode="y" aria-pressed="false">Years</button></div>
<div class="chart" id="margins" role="img" aria-label="Gross, operating and net margin, lines"></div>
<p class="note">Q1 2026's very low net margin includes a {money(-g('Q1-2026', 'other_income'))} "other expense", mostly
a {money(g('Q1-2026', 'loss_debt_extinguishment'))} loss from repaying debt early (S-1/A cash flow statement, page F-67).</p>
{mgmt(f"Q3's gross margin included one-off benefits (roughly {g('Q3-2026', 'connectivity_gm_one_off'):g} percentage "
      f"points in Connectivity, mainly tariff refunds); management expects Q4 2026 gross margin to look more like "
      f"Q2's ({pct(q2gm, 0)}).")}
<details><summary>Details: profit and margins ($ millions)</summary>{prof_table}</details>

<h2 id="profit-vs-cash">4. Profit vs cash: is the company generating or using cash?</h2>
<div class="read">Profit and cash are not the same thing. <b>Operating cash flow</b> (cash collected from customers
minus cash paid to run the business) was {money(h['cfo_9m26'])} in January-September 2026, although that period shows a
net loss of {money(-h['ni_9m26'])}. Why? Big costs such as <b>depreciation</b> (spreading the cost of rockets,
satellites and computers over their useful life) and <b>share-based compensation</b> (paying staff in shares) reduce
profit but use no cash; and customers paid {money(g(m9, 'change_deferred_revenue'))} in advance (<b>deferred
revenue</b>: cash received for work not done yet). Then the company spent {money(h['capex_9m26'])} on <b>capital
expenditures (capex)</b>: data centres, satellites, launch sites. That is {capex_mult:.1f} times its operating cash
flow. <b>Free cash flow</b> (operating cash flow minus capex) was {money(h['fcf_9m26'])}. So SpaceX is using cash, not
generating it, and has been since 2024.</div>
<div class="chart" id="cash" role="img" aria-label="Operating cash flow, capex and free cash flow"></div>
<p class="note">Up = cash coming in from operations; down = cash spent on capex; diamond = free cash flow. Only periods
with a full cash flow statement are shown (Q2 and Q3 on their own are not reported).</p>
<h3>Where did the cash come from in January-September 2026?</h3>
{su_table}
<p class="note">Source: Q3 release cash flow statement, page 12. Cash includes restricted cash.</p>
{mgmt(f"Q4 2026 capex is expected at {capex_guide}, on top of the {money(h['capex_9m26'])} already spent in "
      f"January-September.")}
<details><summary>Details: cash flow by period ($ millions)</summary>{cash_table}</details>

<h2 id="debt-and-cash">5. How much does it owe, and how much cash does it hold?</h2>
<div class="read">On 30 September 2026 SpaceX held {money(h['cash_sep26'])} of cash and <b>marketable securities</b>
(short-term investments that can be sold quickly) against {money(h['debt_sep26'])} of <b>debt and finance leases</b>
(borrowed money, plus long leases that work like loans). That leaves <b>net cash</b> of {money(h['net_sep26'])}. At the
end of 2025, before the IPO, the two were about even (net cash {money(net_dec25)}). In June 2026 the IPO raised about
{money(g('Q2-2026', 'ipo_net_proceeds'))} and the company also sold {money(g('Q2-2026', 'bond_issue'), 0)} of bonds.
Debt still rose by only {money(debt_jun - debt_mar)} between March and June, because older loans were repaid: debt
repayments were {money(-g(m9, 'debt_repayments'))} in January-September 2026.</div>
<div class="chart" id="debt" role="img" aria-label="Cash and securities versus debt at each balance sheet date"></div>
<p class="note">Interest expense was {money(-g(m9, 'interest_expense'))} in January-September 2026, partly offset by
{money(g(m9, 'interest_income'))} of interest income earned on the cash.</p>
<details><summary>Details: cash, debt and equity ($ millions)</summary>{debt_table}
<p class="note"><b>Shareholders' equity</b> (assets minus liabilities, what belongs to shareholders) jumps after 2025:
before the IPO, {money(g('2025-12-31', 'redeemable_preferred'))} of <b>redeemable convertible preferred stock</b>
(early investors' shares that could be bought back) sat outside equity; it converted into common shares by June
2026, and the IPO added more.</p></details>

<h2 id="ratios">6. Six ratios in plain English</h2>
<div class="read">Ratios turn the statements into a few comparable numbers. Margins (first three) show how much of each
dollar of revenue is kept. Capex / operating cash flow shows how much investment exceeds what the operations bring in.
The current ratio and debt-to-equity are balance-sheet ratios: they compare what the company owns and owes at a date.
Both improved sharply after the IPO.</div>
{ratio_table}
<p class="note">Balance-sheet ratios use the balance sheet at the end of each period; they are empty where that balance
sheet is not in these files (end-2023 and 30 September 2025). Debt-to-equity looks very high at the end of 2025
({fmt_ratio(t['debt_to_equity']['FY2025'], 'x')}) because the preferred stock sat outside equity then. The workbook
(Ratios tab) has the same ratios for every quarter, as Excel formulas.</p>

<h2 id="outlook">What management says comes next</h2>
<p class="note">From the Q3 2026 earnings call (transcript, paraphrased). These are management's statements, not
filed or audited figures. Where the call quotes past results, they match the filings.</p>
<ul class="plain">
<li><b>Q1 2027 revenue:</b> expected at {revenue_guide}, more than double Q3 2026's {money(h['rev_q3_26'])}.</li>
<li><b>Q4 2026 capex:</b> expected at {capex_guide}, driven by data centres in Memphis and Southaven.</li>
<li><b>Backlog</b> (signed contracts not yet delivered): {money(g('2026-09-30', 'backlog'), 0)} at 30 September 2026,
with about {money(g('2026-09-30', 'backlog_next_12m'), 0)} expected to become revenue within a year (the Q2 release
reported {money(g('2026-06-30', 'backlog'))} at 30 June).</li>
<li><b>Computing capacity:</b> {g('2026-09-30', 'nameplate_compute'):g} GW at 30 September 2026 (filing); management
targets {compute_guide} by the end of 2026.</li>
<li><b>Gross margin:</b> Q4 expected closer to Q2's level, because Q3 had one-off benefits.</li>
</ul>

<section class="challenge" id="challenge">
<span class="badge">CHALLENGE - optional, skip it if you like</span>
<h2>From adjusted EBITDA to net income, step by step (Q3 2026)</h2>
<div class="read"><b>EBITDA</b> means earnings before interest, taxes, depreciation and amortization. <b>Adjusted
EBITDA</b> also leaves out other items management chooses (here share-based compensation, acquisition costs and
restructuring). It is a <b>non-GAAP</b> measure: GAAP are the official US accounting rules, and this number is not
defined by them, so each company builds its own. Management likes it because it strips out non-cash and one-off items;
critics point out that those costs are real. The waterfall starts from Q3 2026 adjusted EBITDA of
{money(g('Q3-2026', 'adj_ebitda'))} and takes each excluded item back in, until it reaches the
{money(h['ni_q3_26'])} of net income in the income statement.</div>
<div class="chart tall" id="bridge" role="img" aria-label="Waterfall from adjusted EBITDA to net income, Q3 2026"></div>
<p class="note">Grey = subtotals; orange = takes money away; blue = adds money. Source: Q3 release, reconciliation
table (page 10) and income statement (page 9).</p>
<ol class="steps">
<li><b>Depreciation &amp; amortization (-{money(g('Q3-2026', 'depreciation_amortization'))}):</b> the yearly share of
the cost of GPUs, satellites and rockets. No cash leaves this quarter (the cash left when they were bought, as capex),
but the assets wear out, so it is a real cost.</li>
<li><b>Share-based compensation (-{money(g('Q3-2026', 'share_based_comp'))}):</b> staff paid in shares. No cash, but
existing shareholders own a smaller slice.</li>
<li><b>Share-based compensation from acquisitions (-{money(sbc_acq_q3)}):</b> shares granted to keep Cursor's staff.
The release says {money(g('Q3-2026', 'cursor_post_acquisition_awards'))} of such awards will be expensed over the next
few years, so this item will come back.</li>
<li><b>Acquisition-related costs (-{money(g('Q3-2026', 'acquisition_costs'))}) and restructuring
(-{money(g('Q3-2026', 'restructuring'))}):</b> lawyers' and advisers' fees for the Cursor deal, and costs of
reorganising the business.</li>
<li><b>= Operating income ({money(oi_q3)}):</b> matches the income statement exactly. This is where GAAP profit from the
core business stands.</li>
<li><b>Interest expense ({money(g('Q3-2026', 'interest_expense'))}) and interest income (+{money(ii_q3)}):</b> paid on
debt, earned on the cash and securities raised in the IPO.</li>
<li><b>Other income (+{money(oth_q3)}):</b> gains outside the core business ({money(g('Q3-2026', 'other_income_related_party'))} of it from a related party).
What it is, and whether it repeats, is explained in the 10-Q, not in these files.</li>
<li><b>Income tax (+{money(-tax_q3)}):</b> a tax <i>benefit</i> this quarter, so tax added to profit instead of
reducing it. = <b>Net income {money(h['ni_q3_26'])}</b>.</li>
</ol>
<details><summary>Details: the same bridge for three periods ($ millions)</summary>{bridge_table}
<p class="note">The definition gained two add-backs in Q3 2026 (acquisition share-based compensation and acquisition
costs); in 2025 those lines did not exist, so they are 0.</p></details>

<h2 id="questions">Open questions a beginner could investigate</h2>
<ol class="qs">
<li><b>What did Q4 2025 look like?</b> It is not reported on its own here. Work it out: full year 2025 minus
January-September 2025 (both are in the workbook). Was Q4 2025 profitable?</li>
<li><b>Why was net income bigger than operating income in Q3 2026?</b> Find what the {money(oth_q3)} of "other
income" and the {money(-tax_q3)} tax benefit were (the Q3 10-Q notes). Are they likely to repeat?</li>
<li><b>How much of the Q3 gross margin jump will last?</b> Gross margin went from {pct(q2gm, 1)} in Q2 to
{pct(gm['Q3-2026'], 1)} in Q3; management says part was one-off. Which segment drove it?</li>
<li><b>How will the cash gap be funded?</b> Capex was {money(h['capex_9m26'])} in January-September and management
expects {capex_guide} more in Q4 alone, against {money(h['cfo_9m26'])} of operating cash flow. How long does
{money(h['cash_sep26'])} of cash and securities last at that pace, and what new debt could be needed?</li>
<li><b>Where did the {money(g('Q2-2026', 'bond_issue'), 0)} bond money go?</b> Debt rose only
{money(debt_jun - debt_mar)} from March to June 2026. Which loans were repaid, and why was there a
{money(g('Q1-2026', 'loss_debt_extinguishment'))} loss on repaying debt early in Q1?</li>
<li><b>What is the {money(g('2026-09-30', 'goodwill'))} of goodwill?</b> <b>Goodwill</b> (the part of a purchase price
above the fair value of what was bought) rose from {money(g('2026-06-30', 'goodwill'))} in June, mostly from Cursor
({money(g('Q3-2026', 'cursor_purchase_consideration'))} purchase price). What would an impairment mean, as with the
Twitter brand in 2023?</li>
<li><b>How big is share-based compensation becoming?</b> {money(g(m9, 'share_based_comp') + g(m9, 'share_based_comp_acquisitions'))}
in January-September 2026. Compare it with operating income and with adjusted EBITDA.</li>
<li><b>Why is deferred revenue growing?</b> It reached
{money(g('2026-09-30', 'deferred_revenue_current') + g('2026-09-30', 'deferred_revenue_noncurrent'))} at 30 September
2026 (current + non-current). Who is paying in advance, and what does that mean for future revenue and cash?</li>
</ol>
</section>

<h2 id="glossary">Glossary</h2>
<dl class="gloss">{gloss}</dl>

<h2 id="notes">Data notes, checks and sources</h2>
<p>Every number on this page comes from one data file (<code>data/spacex_data.csv</code>) with its document and page.
The workbook <code>data/spacex_accounting.xlsx</code> recomputes the ratios as Excel formulas and runs
{h['checks']} arithmetic checks: segments add up to totals, the balance sheet balances, the cash flow reconciles,
quarters add up to year-to-date. Result: {h['checks'] - h['checks_fail']} pass ({h['checks_rounding']} only within
${TOLERANCE}m of rounding) and {h['checks_fail']} fail.</p>
<details><summary>Checks by group, and the failures</summary>{check_table}
<ul class="plain">{fail_items}</ul>
<p class="note">Within rounding (filings round each line to $1m): {rounding_items}.</p></details>
<details><summary>Figures not in these files</summary><ul class="plain">{gaps}</ul></details>
<details><summary>Differences between documents</summary><ul class="plain">{discs}</ul></details>
<details><summary>Sources</summary><ul class="plain">
<li><b>Form S-1/A</b> (preliminary prospectus, "subject to completion, dated June 3, 2026"): audited statements for
2023-2025 (pages F-4 to F-9), unaudited Q1 2026 and Q1 2025 (F-63 to F-68), segment notes (F-57, F-58, F-94, F-95),
adjusted EBITDA reconciliations (pages 119-121).</li>
<li><b>Q2 2026 earnings release</b> (quarter ended 30 June 2026): highlights (page 3), statements (pages 8-10).</li>
<li><b>Q3 2026 earnings release</b> (quarter ended 30 September 2026): highlights (page 3), statements (pages 9-12).</li>
<li><b>Q3 2026 earnings call transcript</b> (20 October 2026, third-party copy): management commentary only, paraphrased;
not republished.</li></ul></details>

<footer>Built by <code>SpaceX/scripts/build_outputs.py</code> from <code>data/spacex_data.csv</code>.
<b>Personal learning project, not investment advice.</b> No buy/sell view, price target or valuation.</footer>
<script>{JS.replace('__DATA__', json.dumps(payload))}</script>
</main></body></html>"""
    HTML_PATH.write_text(html, encoding="utf-8")        # write the page
    # returns nothing; writes spacex.html (~60 KB; ECharts is loaded from lib/echarts.js)
    return None


if __name__ == "__main__":
    main()
