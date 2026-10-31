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

A new company gets a tone of its own, close to its brand colour, bright enough to read on black.

**Page structure, in this order:**
1. **Header:** the accent rule, back link "← All companies", **the company name on its own** as the
   title ("Anthropic", "SpaceX") — not a question and not a sentence — then one small line giving
   the filing, its date, the units and the workbook. Three lines in all.
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

## Homepage (`index.html`, always keep this structure)

Written by hand (no script). It follows all the Phone-friendly rules above.

- **Header band** in slate with the steel-cyan accent: small label "Equity research · Primary
  filings", the title "Equity Research Lab", the description, then a "View the code on GitHub"
  button with the **GitHub logo** (inline SVG).
- **Featured company:** one large card overlapping the band with a "Featured" badge. Chosen by me,
  not automatically the newest. Only change it when I ask.
- **"All companies" grid:** one card per other company — preview image, small accent label, serif
  title, one or two sentences. **Two cards to a row**, each half the width, whatever the number of
  companies (`repeat(2, minmax(0, 1fr))`, one per row below 760px). Not `auto-fill`: with a single
  company that shrank the card to its minimum width and left it looking lost beside the featured one.
- **No "Open the page →" text.** The whole card is the link.
- **Hover animation (mouse only):** in the **company's own accent** (`--c`, from the card's
  `--c-l` / `--c-d`), the border lights up, a soft glow appears, the card lifts 4px, the image zooms
  slightly and the title turns that colour. Keep it inside `@media (hover: hover)` so a tap on a
  phone does not leave a card stuck highlighted.
- **Preview images** are made with `python tools/make_previews.py`, never by hand: a readable
  close-up of the top of the page (the company name, the filing line and the start of the first
  section), 4:3 (1120x840), shown in a 4:3
  box with `object-fit: cover; object-position: top center`. The box's background is the company
  pages' black (`#0b0c0e`), so no light edge shows around the image. A shrunk
  screenshot of a whole page is unreadable in a card. Add a new company to `PAGES` in that script.

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

## Philosophy
Expect to try things that don't work, refactor, and document what you learn. Clarity over brevity.
