# Project Context for Claude Code

You are helping **Harshit Gandhi** maintain a personal equity research portfolio
site at https://harshitgandhi0212.github.io/equity-report/

Read this entire file before doing anything.

---

## CRITICAL RULES (do not violate)

1. **Never fabricate financial numbers.** If a value is not in the Screener
   Excel, the existing HTML JSON block, or explicitly provided by Harshit,
   leave it as `null` and add the field to the pending list. Wrong numbers
   in a financial report destroy credibility permanently. Better to ship a
   gap than a hallucination.

2. **Never auto-update judgment fields.** These belong to Harshit only:
   - `scores.*` (risk scores 0-100 per axis + composite)
   - `verdict.*` (rating, fair value, entry zone, stop loss)
   - All prose: catalysts vs risks, "Where I Could Be Wrong", street view,
     industry context, the bottom-line verdict paragraph

3. **Match the canonical template exactly.** `01_template_biocon.html` is
   the single source of truth for structure, CSS, and component patterns.
   Replicate. Don't invent. Don't "improve" the design.

4. **No yfinance. No Yahoo Finance. No CORS proxies. No external APIs.**
   Static data only. Source = Screener.in Excel exports + manual entries.
   This is a deliberate decision, not an oversight.

5. **One commit per logical change.** Clear commit messages. Show `git status`
   and `git diff` before committing. Ask before pushing.

6. **Ask before destructive actions.** Renaming files, deleting versions,
   force-pushing, deleting branches — confirm with Harshit first.

---

## PROJECT STRUCTURE

```
equity-report/
├── index.html                      Landing page (lists all reports)
├── README.md                       Repo readme
├── reports/
│   ├── biocon.html                 ← Canonical v3 (JSON-driven)
│   ├── hfcl.html
│   ├── jainrec.html
│   ├── tiindia.html
│   └── (more stocks as added)
├── data/                           Screener.in Excel exports
│   ├── Biocon.xlsx
│   ├── HFCL.xlsx
│   └── (one per stock per quarter)
├── tools/
│   └── update_report.py            Parser: Excel → JSON
└── _context/                       This folder (Claude Code reference)
    ├── 01_template_biocon.html
    ├── 02_update_workflow.md
    ├── 03_update_report.py
    └── 04_instructions.md
```

---

## REPORT ARCHITECTURE (every report follows this)

1. **`<head>`** — SEO + Open Graph + theme color + inline SVG favicon
2. **`<script id="reportData" type="application/json">`** — all numeric data
   lives here. Single source of truth. JS reads it and populates the data
   badge on page load.
3. **Data freshness badge** — steel-blue (NOT green/pulsing), explicitly
   says `STATIC DATA`, shows source + last update + next scheduled refresh
4. **Top nav** — back link to index + jump-to-report dropdown
5. **Page 1** —
   - Stock bar (ticker, price, change)
   - External links strip (Screener, NSE, BSE, IR, Trendlyne)
   - Analysis timestamp ribbon
   - Risk gauge SVG (220×140)
   - KPI strip (6 metric cards)
   - 12-month price chart (inline SVG)
   - 3 card grids: Valuation, Financial Health, Growth
   - Score breakdown bars (35% / 35% / 30% weighting)
6. **Page 2** —
   - Quarterly table (last 6 quarters)
   - Latest earnings card
   - Catalysts vs Risks (dual column)
   - Street View (analyst grid, with source attribution)
   - Industry Context
   - Verdict + rating bar
   - **Where I Could Be Wrong** (3 falsifiable conditions)
7. **Footer** — disclaimer (NOT SEBI-registered) + copy-summary button
8. **Scripts** — JSON-to-DOM population + clipboard copy

Reports are ~1500-1600 lines. Most is hand-written prose. The JSON controls
~25 numeric fields. Everything else is editorial.

---

## RISK GAUGE MATH (delicate, get this right)

Center `(100, 110)`, radius `80`. ViewBox `0 0 200 130`.

For a score S, needle endpoint:
```
angle_deg = (1 - S/100) * 180
x = 100 + 80 * cos(radians(angle_deg))
y = 110 - 80 * sin(radians(angle_deg))
```

Boundary checks:
- Score 0 → `(20, 110)` (left)
- Score 50 → `(100, 30)` (top)
- Score 100 → `(180, 110)` (right)

Known endpoints (use these exact values):
- BIOCON 54 → `(110.03, 30.63)`
- HFCL 58 → `(119.90, 32.51)`
- TIINDIA 61 → `(127.10, 34.73)`
- JAINREC 68 → `(142.87, 42.45)`

When updating gauge for a new score, change FOUR things:
1. Filled arc path's endpoint
2. Needle line's `x2`, `y2`
3. Needle-tip circle's `cx`, `cy`
4. Score number in `<div class="n">`

**Use `stroke-linecap="butt"` on filled arc** (not `round`). `round` adds
a visual cap that overshoots the score position.

---

## COLOR CONVENTIONS

```
--green: #3ddc97   BEAT, low risk, BUY
--amber: #ffb547   CAUTION, moderate risk, HOLD
--red:   #ff5470   MISS, high risk, SELL/AVOID
--gold:  #d4af37   premium accent, brand
```

Composite score bands:
- `< 40` = green (low risk)
- `40-60` = amber (moderate)
- `> 60` = red (high)

Fonts (no substitutes):
- Body: DM Sans
- Data/mono: JetBrains Mono
- Headers/serif: Fraunces

---

## JSON DATA BLOCK SCHEMA

```json
{
  "stock":         { ticker, name, exchange, code, bseCode, isin,
                     sector, subSector, index },
  "asOf":          { dataDate, dataDateDisplay, latestReportPeriod,
                     priorReportPeriod, analysisDate, analysisDateDisplay,
                     source, screenerExportFile, nextScheduledUpdate, version },
  "price":         { cmp, change, changePct, high52, low52,
                     yearReturn1y, beta },
  "valuation":     { marketCapCr, pe, pb, evEbitda, peg, dividendYield,
                     mcapSales, epsTtm, epsPriorYear, bookValue },
  "fundamentals":  { revenueTtmCr, revenueYoyPct, ebitdaTtmCr, ebitdaYoyPct,
                     ebitdaMarginPct, patTtmCr, patYoyPct, patPriorCr },
  "health":        { debtEquity, interestCoverage, cfoEbitdaPct,
                     promoterHolding, fiiHolding, diiHolding, publicHolding,
                     roe, roce, promoterPledgePct },
  "scores":        { valuation, valuationFlag, financialHealth,
                     financialHealthFlag, growth, growthFlag,
                     composite, compositeBand, compositeBandColor },
  "verdict":       { rating, ratingPosition, fairValue, entryZone,
                     stopLoss, positionCapPct, consensusTp,
                     consensusUpsidePct },
  "links":         { screener, nse, bse, investorRelations, trendlyne },
  "quarterly":     [ { period, revenueCr, operatingProfitCr,
                       operatingMarginPct, patCr }, ... ]
}
```

**Flag values** = `"green"` | `"amber"` | `"red"` (used for card border colors)
**compositeBandColor** = `"green"` (composite < 40) | `"amber"` (40-60) | `"red"` (> 60)
**ratingPosition** = 0-100 (position on the rating bar; 0 = strong sell, 100 = strong buy)

---

## WHAT THE PARSER SCRIPT DOES

`tools/update_report.py` reads Screener Excel → produces JSON.

**Auto-extracts** (~25 fields):
- CMP, market cap, P/E, P/B (bonus-adjusted via DERIVED row)
- Revenue + YoY, PAT + YoY, EBITDA + margin + YoY
- ROE, debt/equity, CFO/EBITDA, dividend yield, mcap/sales
- EPS TTM + prior, book value
- Last 6 quarters (revenue, operating profit, margin, PAT)

**Cannot extract** (must fill manually):
- 52W high/low (Excel only has annual closes)
- Beta, ROCE, interest coverage
- FII / DII / promoter holding (not in default Excel)
- EV/EBITDA, PEG (need extra inputs)
- Segment revenue breakdown
- Intraday change

**Never extracts (judgment)**:
- Risk scores, verdict, fair value, entry/stop, all prose

### Modes

```bash
# Preview only — no file writes
python tools/update_report.py biocon

# Patch HTML in place, MERGE mode (preserves judgment + manual fields)
python tools/update_report.py biocon --in-place
```

### Merge behavior (critical to understand)

When running with `--in-place` or `--merge`:
- Refreshes from Excel: all numeric data, valuation, fundamentals
- **Preserves from existing HTML**: scores, verdict, shareholding, 52W,
  beta, ROCE, EV/EBITDA, analysisDate

This is the killer feature. You can re-run every quarter without losing
your judgment work.

---

## WHAT YOU SHOULD NOT DO

- ❌ Add new design elements / "improve" the layout
- ❌ Use yfinance, Yahoo Finance, or any external API
- ❌ Add new charting libraries (we use inline SVG only)
- ❌ Add tracking pixels, analytics, third-party scripts
- ❌ Change font choices (DM Sans, JetBrains Mono, Fraunces)
- ❌ Shorten the disclaimer or remove "NOT SEBI-registered" language
- ❌ Rename `biocon.html` ↔ `biocon_risk_report.html` without asking
  (filenames are linked from index, breaking changes hurt live URLs)
- ❌ Auto-write the verdict, scores, or any prose
- ❌ Fabricate numbers when data is missing
- ❌ Commit + push without showing the diff first

## WHAT YOU SHOULD DO

- ✅ Match the canonical template's HTML structure tag-for-tag
- ✅ Preserve every comment in the source (they're documentation)
- ✅ Keep CSS inline in `<style>` blocks (no external stylesheet)
- ✅ When finishing changes, run `git status` and `git diff` first
- ✅ Surface pending fields clearly so Harshit knows what's missing
- ✅ If Screener Excel is missing for a stock, say so explicitly
- ✅ When patching the gauge, use the math formula above + verify endpoints
- ✅ Ask if uncertain — easier than reverting

---

## TYPICAL TASKS

### Quarterly data refresh (existing stock)

```bash
# User just dropped a new Excel into data/
python tools/update_report.py biocon --in-place

# Show what changed
git diff reports/biocon.html

# List pending manual fields (script prints these automatically)
```

After running, tell Harshit:
1. Which numeric fields changed materially (>10% delta)
2. Which fields still need manual entry (shareholding, 52W, etc.)
3. Whether material changes warrant prose revision
4. Don't commit until confirmed

### New stock (first report)

1. Add entry to `STOCK_CONFIG` in `tools/update_report.py`
2. Copy `reports/biocon.html` → `reports/<newstock>.html` (template)
3. Update gauge math for new score
4. Update SEO tags + ticker/code references throughout
5. Add to nav dropdown in all reports
6. Add card to `index.html`
7. Run parser to populate JSON: `python tools/update_report.py newstock --in-place`
8. Tell Harshit which fields are pending

### Fix the gauge

Use the formula. Use the known endpoints. Test by opening the file in browser.
If the needle doesn't visually align with the score, the math is wrong.

---

## COMMUNICATION STYLE

Harshit prefers:
- Direct, honest feedback over hedging
- Naming tradeoffs clearly
- Pushing back when something doesn't make sense
- No padding ("great idea", "absolutely", "you're right")
- Concrete reasons when you do agree

If you think Harshit is wrong, say so. If a task is impossible, say so.
If there's a smarter way, propose it.

He hates:
- Filler affirmations
- Vague reassurance
- Mid-task scope creep without flagging it

---

## FILES IN THIS FOLDER

| File | Purpose |
|------|---------|
| `01_template_biocon.html` | Canonical v3 report. Replicate this structure for new stocks. |
| `02_update_workflow.md`   | Quarterly refresh workflow + manual entry checklist |
| `03_update_report.py`     | Copy of the parser (production version lives in `tools/`) |
| `04_instructions.md`      | This file. Read first. |

---

## ONE FINAL THING

If you're about to do something and you're not sure if it matches Harshit's
existing approach — **stop and ask**. Time spent confirming costs nothing.
Time spent reverting a wrong choice costs the whole task.

Start every session by reading this file. Then read whatever else is relevant.
