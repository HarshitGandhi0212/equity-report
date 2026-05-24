# Quarterly Update Workflow

How to refresh report data after a Screener.in download.

---

## One-time setup

1. **Install Python 3.9+** (`python --version` to check)
2. **Install openpyxl:** `pip install openpyxl`
3. **Folder structure:**
   ```
   equity-report/
   ├── reports/biocon.html          ← report file
   ├── data/                         ← Screener Excel exports
   │   └── Biocon.xlsx
   └── tools/
       └── update_report.py          ← the parser
   ```

---

## Workflow (per stock, ~5 minutes total)

### 1 — Download from Screener

1. Visit `https://www.screener.in/company/<TICKER>/consolidated/`
2. Click **"Export to Excel"** (top right)
3. Save to `data/` — name doesn't matter much, the script picks the latest matching file. Default Screener filenames work (e.g., `Biocon.xlsx`).

### 2 — Run the parser

**Recommended: merge mode** (preserves your judgment fields):
```
python tools/update_report.py biocon --in-place
```

This:
- Reads `data/Biocon.xlsx` (latest match)
- Refreshes all 25+ numeric fields (CMP, P/E, EBITDA, ROE, etc.)
- **Preserves** your existing scores, verdict, fair value, shareholding, 52W range
- Patches `reports/biocon.html` directly

**Preview mode** (just see the JSON, no file writes):
```
python tools/update_report.py biocon
```

**Other options:**
```
# Specific Excel file
python tools/update_report.py biocon --excel data/old_export.xlsx

# Don't merge — overwrite everything (use when starting fresh)
python tools/update_report.py biocon --in-place  # without --merge: still merges by default
```

### 3 — Review the "fields pending" list

After running, the script prints fields that need manual fill:

```
⚠ Fields requiring manual fill:
   • price.change       (today's intraday move — not in Excel)
   • price.changePct
   • asOf.analysisDate  (set when you've reviewed and revised analysis)
   • verdict.rating     (only if you started fresh — merge preserves)
   • ...
```

Open `reports/biocon.html`, find the JSON block, fill in:

| Field | Where to get it |
|---|---|
| `price.change`, `price.changePct` | NSE quote page (today's move) |
| `price.high52`, `price.low52` | NSE quote page (52W intraday) |
| `price.beta` | NSE quote page or Trendlyne |
| `health.promoterHolding`, `fiiHolding`, `diiHolding` | Screener web page (shareholding tab) |
| `health.roce`, `health.interestCoverage` | Screener web ratios or compute from AR |
| `valuation.evEbitda`, `peg` | Compute manually or use Tijori/Trendlyne |
| `asOf.analysisDate`, `nextScheduledUpdate` | Today's date / next results date |

### 4 — Update the prose (the part the script can't do)

The script only updates the JSON block. The analysis itself — verdict text, catalysts vs risks, "Where I Could Be Wrong" — is hand-written and stays as is, unless you decide to revise it based on new data.

**Important:** if material numbers changed (e.g., ROE moved from 5% to 1%), you should revise the relevant prose. The script won't.

### 5 — Commit and push

```
git add data/Biocon.xlsx reports/biocon.html
git commit -m "Update Biocon: Q1FY27 data refresh"
git push
```

---

## What the script extracts (auto)

| Field | Source in Excel |
|---|---|
| CMP, Market Cap | META section |
| Revenue TTM, YoY % | Profit & Loss / Sales |
| PAT TTM, YoY % | Profit & Loss / Net profit |
| EBITDA, margin % | Computed: PBT + Interest + Depreciation − Other Income |
| EPS | Computed: Net profit / Adjusted Equity Shares |
| Book Value | Computed: (Capital + Reserves) / Adjusted Shares |
| P/E, P/B | Computed: CMP / EPS or BV |
| ROE | Computed: PAT / Net worth |
| Debt/Equity | Computed: Borrowings / Net worth |
| CFO/EBITDA | Computed: Cash from Ops / EBITDA |
| Dividend yield | Computed: (Dividend Amount / Shares) / CMP |
| Mcap/Sales | Computed: Mcap / Revenue |
| Quarterly (last 6 Q) | Quarters section |
| Latest report period | P&L Report Date |

## What's NOT in the Excel (must fill manually)

| Field | Where to get it |
|---|---|
| Intraday CMP change | NSE/BSE live quote |
| 52-week high/low | NSE/BSE quote page |
| Beta | NSE or Trendlyne |
| FII / DII / Promoter holding | Screener web page (shareholding tab) |
| ROCE | Screener web page (not in default Excel) |
| Interest Coverage | Compute from AR or Screener web |
| Promoter pledge | NSDL or Screener web |
| EV/EBITDA, PEG | Compute or use aggregators |
| Segment revenue breakdown | Annual Report or investor presentation |

## What requires judgment (never automated)

- Risk scores (0-100 per axis)
- Composite band classification
- Verdict (BUY/HOLD/AVOID)
- Fair value, entry zone, stop loss
- Position cap %
- Catalysts vs risks
- "Where I Could Be Wrong" conditions
- Sell-side commentary

---

## Honest caveats

### EBITDA definition
This script computes EBITDA as: `PBT + Interest + Depreciation − Other Income` (pure operating EBITDA). If you compare against Screener's web page or another source, numbers may differ by ±5% because of:
- Inclusion/exclusion of Other Income
- Treatment of exceptional/one-off items
- IFRS vs Ind AS lease accounting

**Pick one definition (we use this one), be consistent, document it.** That's more honest than chasing a "true" number that doesn't exist.

### Bonus-issue handling
The script uses Screener's `Adjusted Equity Shares in Cr` row from the DERIVED section. This is already bonus-adjusted, so per-share metrics (EPS, BV) automatically reflect splits and bonuses.

If you compared the Biocon report before/after the parser ran, you'd see P/E moved from 150.98× to 180.86×. That's because the old number used pre-bonus share count (120 Cr) and the new uses post-bonus (162 Cr). **The new number is correct.**

### What can break
1. **Screener changes their Excel template structure** → the section headers ("PROFIT & LOSS", "Quarters", etc.) might rename. Check `SECTION_HEADERS` in the script if extraction fails.
2. **Banks and financial-sector stocks** → use different metric labels (NIM, NPA, ROA). This script is optimised for industrials/pharma/tech. Banks need a different parser.
3. **Loss-making companies** → P/E will be negative or null. The script handles it gracefully (returns null) but you'll need a separate convention in the report.
4. **Newly-listed stocks with <2 years history** → many YoY computations return null because the prior period is empty. Expected behavior.

---

## Adding a new stock

Edit `tools/update_report.py`:

```python
STOCK_CONFIG = {
    # ... existing entries
    "premier": {
        "ticker": "PREMIERENE",
        "name": "Premier Energies Ltd",
        "exchange": "NSE",
        "code": "544238",
        "bseCode": "544238",
        "isin": "INE0LXG01016",
        "sector": "Renewable Energy",
        "subSector": "Solar PV Cells & Modules",
        "index": "Smallcap",
        "screenerSlug": "PREMIERENE",
        "bseSlug": "premier-energies-ltd/premierene",
        "investorRelations": "https://premierenergies.com/investor-relations/",
    },
}
```

Then:
1. Download the Screener Excel to `data/Premier.xlsx`
2. Make sure `reports/premier.html` exists with at least an empty `<script id="reportData">` block
3. Run: `python tools/update_report.py premier --in-place`

---

## Quick reference

| Command | What it does |
|---|---|
| `python tools/update_report.py biocon` | Preview JSON, no writes |
| `python tools/update_report.py biocon --in-place` | Patch HTML, preserve judgment fields |
| `python tools/update_report.py biocon --merge` | Print merged JSON to stdout |
| `python tools/update_report.py biocon --excel data/file.xlsx` | Use specific Excel |
| `python tools/update_report.py biocon --html custom.html` | Use specific HTML target |
