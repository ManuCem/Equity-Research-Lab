# Equity-Research-Lab

## Primary Goal
Company financials rebuilt from primary filings (S-1s, quarterly results), so the numbers can be
traced back to the page they came from. This is a **learning project**: understand how a company
makes money, what it costs, and what the filing does and does not say.

Not a production system, and **not investment advice**.

## Working with Code

- **Break down code line by line** — explain what each line does, not just the overall function
- **Comment every line** of the data code, and the "why" wherever it isn't obvious
- **Every function ends with a documented return**: right at the `return`, a `#` comment showing the
  type, the shape, and a small example of the real values, laid out the way Python would print them.
  If a function returns nothing and only writes a file, say so and name the file.
- **`def main():` is the conductor** — it calls the other functions in order; the work lives in them.
- **Never edit a generated `index.html` by hand.** Change the script that writes it and re-run it.
- **What is committed is the deliverable, not the machinery.** Each company's `.gitignore` keeps the
  filings, raw text, scripts and most data local; only `index.html` and the workbook (`.xlsx`/`.csv`)
  are pushed, because GitHub Pages serves the page. So a page must be rebuilt **and** checked before
  committing: the committed HTML is the only copy anyone else sees.
- Figures in a page come from the model values (`data/model_values.json` and friends), never typed
  into the text by hand, so the page stays true after a rebuild.

## Phone-friendly HTML (always use this method)

Every page is published on GitHub Pages and read on a phone. Build it phone-ready and **prove it**.

- **Viewport:** `<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">`.
- **No sideways scrolling of the page.** Only a table may scroll sideways, inside its own
  `overflow-x: auto` box, with the first column pinned (`position: sticky`).
- **Finger-sized tap targets:** every link, button and legend entry about 44px tall (at least 40px).
- **Charts:** a separate phone layout (legend below the plot, never on top of it or to the right;
  first and last axis labels not cut off; taller box on a phone). Rebuild on rotation.
- **Never trap the finger:** charts keep `touch-action: pan-y` so a swipe still scrolls the page.
  Chart tooltips must open on **tap** as well as hover, and say "Tap" instead of "Hover" on touch.
- Do not use hover as the only way to reach information.

**How to prove it (run these every time, and fix what they show)**
1. Render the page in a **real 360px-wide frame** (`<iframe style="width:360px">` inside a test page,
   headless Chrome with `--allow-file-access-from-files`). A small headless window is silently cropped.
   On this machine headless Chrome only works with its own profile (`--user-data-dir=<scratch folder>`).
2. **Look at the screenshots** at 360px and at desktop width. The company pages are black in both
   colour schemes, so one shot each is enough; the homepage still needs dark and light
   (`--blink-settings=preferredColorScheme=0` / `=1`).
3. **Measure, don't guess:** `scrollWidth <= clientWidth`, every chart a sensible size, every link
   and button tall enough.
4. Tell me what to try on a real phone: touch behaviour can't be fully tested in a desktop browser.

## Company pages (always use this format)

**Every company page uses the same structure.** `Anthropic/index.html` and `SpaceX/index.html` are
the same page in different colours with different numbers: same shell, same panel pattern, same
order of ideas. If a change improves one, it is made to both.

The shared stylesheet is `assets/report.css` (inlined at build time) and the shared behaviour is
`assets/report.js`. **Page-specific CSS belongs in neither**: each page only overrides its four
colour tokens after the shared CSS.

**Look: the plain data report, on a black page.** Not the Macro Quant Lab project-page look — the
company pages are their own thing and must not be rebuilt to match it. No colour band, no key-figure
row, no rounded cards, no boxed callouts. The page is black (`--page: #0b0c0e`) in light mode and in
dark mode alike; the only thing borrowed from the Macro Quant Lab homepage is the type (DM Serif
Display titles, DM Sans text).

**The page carries the data, not a reading of it.** No short answer at the top, no "How to read it"
note, no "What it tells us" paragraph, no closing summary: the reader gets the numbers, the charts
and the filing they came from, and draws their own conclusion.

The company's colour appears in one place: a 4px rule across the top of the page. It is **one
token**, read on black, and the page overrides nothing else:

| Page | accent |
|---|---|
| Homepage (`index.html`) | slate band `#1d2733` / `#12181f`, steel cyan `#2f8fa8` / `#6cc2d8` (the homepage keeps its band and its light mode) |
| SpaceX | blue `#82b6f8` |
| Anthropic | clay `#e59873` |
| NVIDIA | green `#76b900` |

A new company gets a tone of its own, close to its brand colour, bright enough to read on black.

**Page structure, in this order:**
1. **Header:** the accent rule, back link "← All companies", then as the title **the logo of the company's
   main business followed by its ticker in plain text** (since 18 Nov 2026): Claude logo + "ANT", SpaceX +
   "SPCX", Google's four-colour G + "GOOGL", NVIDIA's eye + "NVDA" — the full company name goes in the h1's
   `aria-label` and the `<title>` — not a question and not a sentence — then one small line giving
   the filing, its date, the units and the workbook. Three lines in all.
   The logo is inline SVG inside the h1 (`<span class="logo">`, sized by `h1 .logo` in `report.css`): the
   Simple Icons mark (public domain) in the brand colour, or Google's four-colour G. **The Anthropic, SpaceX
   and Alphabet pages were patched by hand on 18 Nov 2026 (the h1 line and the `h1 .logo` CSS only): their
   build scripts on the other PC must write the same h1 before the next rebuild, or the rebuild undoes it.**
   **"Last update · 2 Nov 2026"** sits small (11px, muted) in the top right corner, on the back link's
   line (`.updated` in `report.css`). The date is fixed text in a `<span data-static="1">`, set by me
   for each company (the day its data was last updated), not the push date. Change it only when the
   company's data is updated, or when I give a new date.
   After the date, in the same corner, **two icons** (`.ic`): the **Excel logo** is a direct `download` link
   to the `.xlsx` (neither browsers nor GitHub preview Excel files), and the **GitHub logo** opens the repo.
   Each logo has its name written beside it ("Excel", "GitHub"), 44px tall. On a phone (≤640px) the whole
   corner (date, Excel, GitHub) moves to its own line under the back link, so nothing overlaps.
   On the homepage the featured card's image box is 4:3, the preview's own shape, so the chart is never cropped.
   **Tables read like Excel:** labels and formulas stay on one line (`white-space: nowrap`), so each column
   takes the width of its longest entry and every row is the same height; cells have room around them
   (`9px 16px`) and a faint line between columns. Only on a phone may the pinned label column wrap.
2. **Straight into the data.** No key figures, no introduction.
3. **One flat section per topic**, separated by a hairline, each with:
   1. **A plain section name as the heading** — "Revenue", "Balance sheet", "Operating metrics".
   2. **One factual headline** under it: a single bold line stating what the data shows, with the
      number in it. Nothing beyond what the figures say.
   3. The chart or table (plus the small "hover a bar / swipe the table" hint where it helps).
4. **References, not a section:** after the last section, a small quiet list (muted, no box, no
   heading): the filing with its date, the page-citation workbook, anything else. Each link at
   least 40px tall. **Where the data comes from must never take centre stage.**

`panel(num, question, dek, how, body, so)` in both build scripts still takes all six arguments so
the call sites read in order, but only `question` (the plain section name), `dek` (the factual
headline) and `body` are printed. `num`, `how` and `so` are accepted and dropped.

**Charts are drawn for a black page.** Colours in the generated SVG come from the CSS tokens
(`fill="var(--c1)"`, `stroke="var(--grid)"`), never hard-coded hex. The text halo is `var(--page)`,
not white. The tooltip is a dark panel with the page's own ink, never a white box.

**A chart is never stretched past about 520px.** Each chart is drawn once at a fixed viewBox around
360 units wide, so its labels grow with the box it is put in: at 460px (the width inside a `.two`
grid) the 11.5-unit text reads at about 16px, but stretched across a full 940px column the same text
comes out at 30px and the chart stops looking like part of the page. `.panel>svg[data-chart]` is
therefore capped at 520px, and both pages' charts land within a few pixels of each other. Check the
**first and last axis labels** fit inside the viewBox: the right gutter has to be wide enough for
half the last label.

**Keep what already works:** the data-cell tagging (`data-cell`, `data-static`) that lets
`compare_report.py` / `verify_report.py` check every figure against the workbook, the clickable
legends, the row-and-column table highlight, and the page citations.

**Tables select like Excel** (`report.js`, all three pages): a click on a cell outlines it in the page
accent and keeps its row and column highlighted; the arrow keys move the selection, and the table and
the page scroll themselves so the cell stays in view (no dragging the bottom scrollbar on a desktop);
Esc or a click outside a table clears it. While a cell is selected the hover
highlight is off (`body.has-sel`), so only one thing is ever lit. STEP rows are not selectable; a click on a label picks its
first figure.

## Workbooks: the finance-course layout (Anthropic since 2 Nov 2026)

The workbook is laid out like my finance course files (`finances-examples/`, local only, git-ignored):
**no Sources sheet, no Checks sheet, no page-citation column** — I do not want to see data sources or
checks, in the workbook or on the page.

- First sheet `Case`: what the company is, the periods, the colour code, and what each sheet holds.
- Then one sheet per part of the finances, numbered: `1 Income statement`, `2 Quarters`,
  `3 Balance sheet`, `4 Working capital`, `5 Cash flow`, `6 Solvency`, `7 Customers`,
  `8 Commitments`, `9 Share capital`.
- Each sheet: title in bold orange (`FF5A00`) in A1, a grey note in A2, an orange header row
  (white bold) "Concept | periods | (empty) | Formula" in row 4 (Anthropic: "… | (empty) | Explanation |
  Formula", the explanation first because it matters more), panes frozen at B5, Arial 10.
- Blocks open with a beige (`F3E6DA`) **"STEP n - …"** row. Green numbers = reported by the company,
  black = calculated with a live formula, **bold on yellow** = the key result of the step.
- Costs are negatives in brackets (`#,##0;\(#,##0\);\-`), so each result is the sum of the lines above.
- **Every line starts with its sign in brackets** (Anthropic since 14 Nov 2026): **(+)** added, **(−)** subtracted,
  **(±)** either way (a gain in some periods, a loss in others), **(=)** the result, **(R)** a ratio or a calculation,
  **(i)** information that is not added into anything. Each (=) is the (=) above it plus the lines in between.
- **Each STEP reads on its own:** a figure the step needs from elsewhere is brought in as its first lines
  ("(±) Operating income (EBIT) — From STEP 4"), so nobody has to scroll up. Example, adjusted operating income:
  (±) Operating income, (+) Stock-based compensation, (+) Donation matching, (+) Bartz settlement added back,
  (=) Adjusted operating income, (R) Adjusted operating margin. These repeated lines are marked `link` in
  `data/course_layout.json`: they stay in the workbook but not on the web page or in the CSV, which show each
  number once. The signs are printed in the workbook only; the page keeps the plain labels.
- Before any rebuild of the Anthropic workbook, run `Anthropic/scripts/capture_notes.py <first build>` to record
  my notes in `data/user_notes.json`; `annotate_workbook.py` puts them back on the same lines afterwards.
- The **Formula** column says in words how each calculated line is made ("EBIT / Revenue").
  Reported lines leave it empty. The Spanish term goes in brackets where the course uses one.
- `anthropic.csv` is the same content in one long table: sheet, step, concept, period, value,
  unit, formula.
- **Explanations, in the workbook only** (Anthropic since 14 Nov 2026): every line's name ends with a short
  plain-English note in brackets ("Compute and infrastructure (cloud servers, AI chips, leased data centres,
  power)"), an **Explanation** column right after the numbers (before Formula) holds the longer ones, each ending
  with an **"Example: …"** that says what is behind the line, taken from the S-1/A with its page (what a
  cost includes, the contract and the counterparty: "Five-year leases of Google's TPU AI chips, made by Broadcom…").
  **A number appears only when it is compared with another one** (9M2025 against 9M2026, a share of revenue
  before and after); never just the line's own figure, which the reader already sees. Web sources may add
  background, but the filing wins where they disagree, and each beige STEP row says what its block is about. They are added on top of the workbook, in place, by
  `Anthropic/scripts/annotate_workbook.py` (local): it backs up first, never writes over a filled cell, leaves
  any label I have edited as I wrote it, refreshes only the "Example: …" part it wrote (so examples follow the
  latest figures), and puts back the formulas' stored values. **The web page stays clean**:
  it prints the plain labels from `data/course_layout.json`, numbers only, no explanations.

The company page follows the workbook: one section per sheet in the same order, each sheet's
blocks as a table (STEP rows on `--step`, key rows on `--key`), with the charts at the top of their
section. **The page shows only the labels and the numbers**: the Formula column stays in the workbook and
is hidden on every page by the shared CSS (`th.fx,td.fx{display:none}`), and the arrow keys never step
into it. The table hint reads "Swipe the table sideways for every period." (the SpaceX and Alphabet
build scripts on the other PC still say "…and the formula.": change that line there before rebuilding). Under the last section there are only two quiet links,
the `.xlsx` and the `.csv`. All four companies use this layout (SpaceX and Alphabet since 3 Nov 2026, NVIDIA since 18 Nov 2026).

For SpaceX and Alphabet, `build_model.py` writes the model to `data/historicals.xlsx` and
`data/historicals_map.json`; `build_course.py` lays it out, unchanged, as `<company>.xlsx` and `.csv`, with
the shared local tool `tools/course_layout.py` (git-ignored), and stops if any figure differs from the model.
Their reported signs are kept as filed (costs positive), since the data must not change.

**I take my own notes inside the workbooks — never delete them.** The `.xlsx` files are not only build
output: I write notes in them (extra cells, comments, text beside the tables, extra sheets). When I ask
for a change to a workbook:
- **Before** rebuilding or overwriting it, open the current file and find everything I added that the
  build does not write (any cell outside the generated layout, cell comments, sheets the build does not make).
- Keep all of it: carry my notes over into the new file in the same place (same sheet, same cells), or, if
  the layout moved, next to the same line, and tell me exactly where each note ended up.
- If a note cannot be kept where it was, stop and ask me before going on. Never drop a note silently.
- Before any rebuild, keep a dated backup copy of the workbook in the company's `data/` folder
  (git-ignored), e.g. `data/backup/spacex_2026-11-03.xlsx`.

## Homepage (`index.html`, always keep this structure)

Written by hand (no script). It follows all the Phone-friendly rules above.

- **Header band** in slate with the steel-cyan accent: small label "Equity research · Primary
  filings", the title "Equity Research Lab", the description, then a "View the code on GitHub"
  button with the **GitHub logo** (inline SVG).
- **Featured company:** one large card overlapping the band with a "Featured" badge. Chosen by me,
  not automatically the newest. Only change it when I ask. Currently NVIDIA (since 18 Nov 2026).
- **"All companies" grid:** one card per other company — preview image, a ticker row (the logo of the
  company's main business drawn inline, the ticker in bold plain text, then the sector as the small accent
  label: AI, Big tech, Aerospace, Semiconductors — no exchange name), serif title, one or two sentences.
  The featured card has the same ticker row under its badge. **Two cards to a row**, each half the width, whatever the number of
  companies (`repeat(2, minmax(0, 1fr))`, one per row below 760px). Not `auto-fill`: with a single
  company that shrank the card to its minimum width and left it looking lost beside the featured one.
- **No "Open the page →" text.** The whole card is the link.
- **Hover animation (mouse only):** in the **company's own accent** (`--c`, from the card's
  `--c-l` / `--c-d`), the border lights up, a soft glow appears, the card lifts 4px, the image zooms
  slightly and the title turns that colour. Keep it inside `@media (hover: hover)` so a tap on a
  phone does not leave a card stuck highlighted.
- **Preview images** are made with `python tools/make_previews.py`, never by hand: a readable
  view of the top of the report that shows **how much the page holds**, so a non-technical reader sees a
  serious piece of work: the header with its corner, the first headline, the first chart with the start
  of its table beside it (label column capped, long labels end in "…"), and the page's next chart under
  the first. Only the hints are hidden; the page files are never changed. Frame 1200x900, saved
  4:3 (1120x840), shown in a 4:3
  box with `object-fit: cover; object-position: top center`. The box's background is the company
  pages' black (`#0b0c0e`), so no light edge shows around the image. Add a new company to `PAGES` in
  that script.

## README (`README.md`, always keep this structure)

Minimalist: **the website is the protagonist.**

1. **Title:** just `# [Equity Research Lab](https://manucem.github.io/Equity-Research-Lab/)` — the
   title itself is the link. No icon, no highlight box, no second copy of the address.
2. **Description:** the same sentence as the homepage.
3. **`## Companies`:** one `###` section per company, in the same order as the homepage. Each has a
   heading that is the company name linked to its live page, a **fairly wide table** of its
   figures, **4-5 bullet points** with the most important numbers, and a bold `**Conclusion:**` line.
4. **`## How it is built`:** the filing → text → model → page pipeline in a few lines.
5. **`## License`:** MIT, with the link.

**Never in the README:** screenshots, "Run it yourself" blocks, a "What's here" list, "Open the
page" links, or a yellow `<mark>` highlighter.

## Do Not Commit
- `.env` files or API keys
- Real account data, trade logs, or balances
- Any credentials or secrets
- Any code: build scripts, tools and helpers stay on this machine. Each company pushes only its
  `index.html`, workbook (`.xlsx`) and flat table (`.csv`), allow-listed in its own `.gitignore`.

## Git: never push or pull
**Do not run `git push` or `git pull`** (nor `git fetch`, `git merge` from the remote, or a forced
push). I sync with GitHub myself. Change files, check them, and at most commit locally when I ask;
then stop and tell me it is ready to push.

## Philosophy
Expect to try things that don't work, refactor, and document what you learn. Clarity over brevity.
