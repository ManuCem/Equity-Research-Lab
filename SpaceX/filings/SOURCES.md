# SpaceX: source documents

The four files used for this analysis. Every number in `../data/spacex_data.csv` names its file and page.
"PDF page" means the page number in a PDF viewer; the printed page number (for example F-5) is in the `page_label` column.

| File | Original name | What it is | Period covered | Pages used |
|---|---|---|---|---|
| `s1a.pdf` | `S-1A.pdf` | **Form S-1/A**, the preliminary IPO prospectus ("Subject to completion, dated June 3, 2026"; 555,555,555 Class A shares at an expected $135). 405 pages. | Audited years 2023, 2024, 2025; unaudited Q1 2026 and Q1 2025 | Statements F-4 to F-9 (PDF 312-317) and F-63 to F-68 (PDF 371-376); segment notes F-57, F-58, F-94, F-95 (PDF 365, 366, 402, 403); adjusted EBITDA reconciliations, printed pages 119-121 (PDF 145-147) |
| `q2-2026-release.pdf` | `Q2.pdf` | **Q2 2026 earnings press release** (not a 10-Q). 10 pages. | Quarter ended 30 Jun 2026 vs Q2 2025; Jan-Jun 2026 and 2025; balance sheet at 30 Jun 2026 | Highlights p.3; statements of operations and adjusted EBITDA p.8; balance sheet p.9; *selected* cash flow p.10; IPO and bond p.7; backlog p.1 |
| `q3-2026-release.pdf` | `Q3.pdf` | **Q3 2026 earnings press release** (not a 10-Q). 13 pages. | Quarter ended 30 Sep 2026 vs Q3 2025; Jan-Sep 2026 and 2025; balance sheet at 30 Sep 2026 | Highlights p.3; segment pages 4-6; Cursor p.7; statements of operations p.9; adjusted EBITDA p.10; balance sheet p.11; full cash flow statement p.12 |
| `q3-2026-transcript.txt` | `Transcript Q3.txt` | **Q3 2026 earnings call transcript**, 20 Oct 2026, a third-party copy (Motley Fool) | Management commentary and guidance | Used only for labelled management commentary. **Not published**: it is excluded by the root `.gitignore` |

## Notes

- The two releases are press releases, not 10-Qs. They contain the main statements but not the notes; anything that needs the notes (for example what "other income" was in Q3 2026) is listed as an open question.
- The S-1/A in this folder does not show its amendment number or SEC filing date.
- `q2-2026-release.pdf` says it "corrects a typographical error that appeared in the version furnished to the SEC" (a Q2 2025 compute figure), so it is not byte-for-byte the SEC copy. No number used here is affected.
- The AI segment is renamed "Super Intelligence (SI)" in the Q3 release. It is the same segment.
- Numbers come from the filings. The transcript is used only for commentary and guidance. Where the transcript quotes past results, they match the filings (see "Differences between documents" in `../spacex.html`).
