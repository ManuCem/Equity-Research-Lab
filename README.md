# Micro-Lab

## Companies

Basic accounting analyses of individual companies, built from their SEC filings. These are personal learning projects, not investment advice.

| Company | Ticker | Covers | Read it |
|---|---|---|---|
| [SpaceX](SpaceX/) | SPCX | Years 2023-2025, quarters to Q3 2026 | [spacex.html](SpaceX/spacex.html) |

### Folder convention (copy it for the next company)

Each company gets one self-contained folder:

```
<Company>/
├── README.md              what it is, main takeaways, how to reproduce, disclaimer
├── <company>.html         the web page (generated, charts from lib/)
├── filings/               source documents, renamed consistently, + SOURCES.md
│                          (call transcripts stay local: */filings/*transcript* is in .gitignore)
├── data/
│   ├── <company>_data.csv          ONE data file: every number with its document and page
│   └── <company>_accounting.xlsx   statements, ratios as formulas, checks, sources (generated)
├── scripts/
│   ├── build_outputs.py   CSV -> workbook + web page + README key-numbers block
│   └── verify_sources.py  is every CSV number printed on the page it cites?
└── lib/                   chart library saved locally (no CDN)
```

Rules: numbers come only from the files in `filings/`. A transcript is used only for clearly labelled management commentary. Missing figures are left empty and listed as open questions, never estimated. Generated files are rebuilt by the scripts and never edited by hand.
