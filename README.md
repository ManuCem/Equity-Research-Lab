# [Equity Research Lab](https://manucem.github.io/Equity-Research-Lab/)

Company financials rebuilt from the filings themselves — statements, ratios and operating metrics,
every figure traced back to the page it came from.

## Companies

### [Anthropic](https://manucem.github.io/Equity-Research-Lab/Anthropic/)

FY2023 to FY2025, the first halves of 2025 and 2026, and quarterly figures to Q2 2026, from the
Form S-1 filed 26 October 2026.

| | FY2025 | First half 2025 | First half 2026 |
|---|---|---|---|
| Revenue | $4,585m | $1,163m | $16,286m |
| Gross margin | 40% | 12% | **85%** |
| Operating margin | −176% | −184% | **−4%** |
| Adjusted operating margin (non-GAAP) | −106% | −159% | **10%** |
| Compute and infrastructure / revenue | 160% | 201% | **55%** |
| Free cash flow | −$2,768m | −$1,531m | **+$3,447m** |

- Quarterly revenue went from **$376m** in Q1 2025 to **$11,555m** in Q2 2026, and revenue for the
  last twelve months to June 2026 was **$19,708m**.
- Compute and infrastructure cost fell from **251%** of revenue in Q1 2025 to **48%** in Q2 2026,
  which is what turns the gross margin from 12% to 85%.
- Operating income turned positive in Q2 2026, at an 8% margin.
- The first-half 2026 net loss of **$80.7B** comes mostly from a **$72.4B** fair-value change on
  convertible notes, not from operations.
- Purchase commitments at June 2026 were **$265.6B**, and signed leases not yet started a further
  **$321.4B**, against cash and short-term investments of **$131.7B**.

**Conclusion:** the business reached operating breakeven in Q2 2026, but the commitments it has
signed are several times the cash it holds, and the S-1 gives no IPO price range.

### [Alphabet](https://manucem.github.io/Equity-Research-Lab/Alphabet/)

The September 2026 quarter, from the Q3 2026 earnings release, the Q3 2025 release and the Q3 2025
Form 10-Q.

| | Q3 2025 | Q3 2026 | Change |
|---|---|---|---|
| Revenues | $102.3B | **$127.6B** | +25% |
| Google Cloud revenue | $15.2B | **$30.7B** | +103% |
| Income from operations | $31.2B | **$40.6B** | +30% |
| Operating margin | 30.5% | **31.8%** | +1.3pp |
| Other income (expense), net | $12.8B | **−$8.2B** | — |
| Net income to common stockholders | $35.0B | **$26.2B** | −25% |
| Diluted EPS | $2.87 | **$2.12** | −26% |
| Capital expenditure | $24.0B | **$51.4B** | +114% |
| Free cash flow | $24.5B | **−$9.1B** | — |

- Revenue grew **25%** to **$127.6B**, with Google Cloud up **103%** to **$30.7B**; the release names
  TPU system sales as a meaningful contributor for the first time.
- Operating income rose **30%**, but net income to common stockholders fell **25%**, because other
  income swung from **+$12.8B** to **−$8.2B** on an **$8.7B** loss on equity securities.
- Capital expenditure reached **$51.4B** — **40.3%** of revenue and **121%** of operating cash flow —
  so free cash flow was **−$9.1B** for the quarter and trailing-twelve-month free cash flow fell from
  **$73.6B** to **$19.7B**.
- Share buybacks stopped: **$0** in the first nine months of 2026 against **$39.8B** a year earlier,
  while the company raised **$42.9B** of common stock, **$19.1B** of preferred and **$85.2B** of debt.
- Total assets passed **$1.00T** at 30 September 2026 and long-term debt went from **$46.5B** to
  **$125.1B**.

**Conclusion:** record revenue and a higher operating margin sat alongside a 25% fall in earnings and
a swing to negative free cash flow; GOOGL closed **12.6%** lower the next trading day.

### [SpaceX](https://manucem.github.io/Equity-Research-Lab/SpaceX/)

Q1 2025 to Q3 2026, from the S-1 (filed 20 May 2026) and the Q2 and Q3 2026 results releases.

| | Q1 2026 | Q2 2026 | Q3 2026 |
|---|---|---|---|
| Revenue | $4.7B | $7.8B | **$14.6B** |
| Gross margin | 49.1% | 55.3% | **62.4%** |
| Operating margin | −41.4% | −1.8% | **19.1%** |
| Free cash flow | −$9.1B | −$15.9B | **−$10.0B** |
| Capex / revenue | 2.2x | 2.4x | **1.3x** |
| Net debt / LTM Adj. EBITDA | 0.9x | −7.3x | **−3.4x** (net cash) |
| Current ratio | 1.2x | 5.1x | **3.1x** |

- Revenue reached **$14.6B** in Q3 2026, up **161%** year on year, with adjusted EBITDA of **$8.8B**
  (60.4% margin) and net income of **$4.4B**.
- SI became the largest of the three segments at **$7.6B** in Q3 2026, up from **$0.8B** in Q1, ahead
  of Connectivity at **$5.2B** and Space at **$1.8B**.
- Capex was **$48.1B** over the first nine months of 2026, **1.8x** revenue, so free cash flow stayed
  negative in every quarter even as the operating margin turned positive.
- Starlink consumer subscribers reached **13.6M** and nameplate compute **1.7 GW** at 30 September 2026.
- Total assets were **$254.0B** and shareholders' equity **$176.8B** at 30 September 2026.

**Conclusion:** the operating margin turned positive in Q3 2026, but capital spending still exceeds
revenue, so the open question is funding rather than profitability. The S-1 copy is truncated, so
Q1 2026 is derived as H1 minus Q2.

## How it is built

Each company is a folder with the same pipeline. The **filing** (a PDF) is read once and its tables
are typed into a build script, every number carrying the document and page it came from. That script
writes the **sourced data** as CSV and asserts the statements tie — segments back to the total,
cash flow back to free cash flow, derived quarters back to the reported nine months. A **model**
stage computes the ratios and writes the workbook (`<company>.xlsx`), which holds the figures, the
ratio formulas and a sheet of citations. A **report** stage writes `index.html` from that workbook,
tagging every figure with the cell it came from, and a **verify** stage reads the finished page back
and re-checks each tagged figure, each chart tooltip and any number that is not tied to a cell.

Only the page and the workbook are committed; the filings, extracted text and scripts stay local and
can be rebuilt with `python <Company>/scripts/run_all.py`.

## License

MIT — see [LICENSE](LICENSE).
