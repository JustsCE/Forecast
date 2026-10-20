# NORDE forecast

Monthly P&L forecast for the Horizon ledger, run entirely from a Claude chat through MCP: GoodHealth pulls the
actuals with one DAX, the chat pastes them into a Google Sheet, the sheet computes growth, baseline and the
P&L matrix with formulas, and submissions from chat override the baseline. No runtime, no files.

| File | Purpose |
|---|---|
| `taxonomy.py` | lines, entities (with division and active flag), divisions, the DAX |
| `sheet.py` | builder: emits the ordered Google Sheets/Drive MCP calls that create the workbook; a Claude Code session relays them |
| `TASK.md` | the chat procedure (refresh, submit, read, verify) — the GoodHealth task guide |

## Workbook `NORDE forecast`

| Tab | Holds |
|---|---|
| `Summary` | the Horizon P&L matrix for the scope in `B1` (TOTAL, a division, an entity): `line \| FY t−2 \| FY t−1 \| FY t−1 % \| YTD t \| YTD t−1 \| YTD % \| F-YoY % \| Jan..Dec t \| FY t \| FY t % \| Jan..Dec t+1 \| FY t+1 \| FY t+1 %`, totals, a Checks block |
| one tab per line | Calculations block (the line's metrics) and months down × blocks `TOTAL \| TOTAL <division> \| <entity>…`, each `Amount \| LY \| YoY % \| F-YoY %`; entity columns grouped under their division |
| `inputs` | append-only submissions; `H` = `target\|line\|month` key |
| `pull` / `actuals` | DAX rows: the latest pull / the history |
| `data` | pull months + actuals before the pull, the single source; keys `pair`, `pair\|month`, `line\|month`, `line\|year`, `line\|year\|ytd` |
| `grid` | one row per (line, entity), one column per month Jan t−2 .. Dec t+1: actuals, then forecasts; row sums and the keys `line\|division`, `scoped line` |
| `model` | current month, trend window, blend weights, per-line metrics, the scope list |
| `lines`, `entities` | taxonomy |
| `Data Sources` | one row per DAX pull |

Forecast months are blue (a conditional rule on the current month), % columns turn red below −10 % and green above +10 %,
zeros show blank, gridlines are hidden.

Every row-wise aggregation is `SUMIF` / `XLOOKUP` over a composite key: `SUMIFS` and `COUNTIFS` do not iterate inside
`ARRAYFORMULA`. Codes are text everywhere (`02.1`, `1.1`); code cells are text-formatted before anything is written into them.

## Rebuild

```
python norde/sheet.py <spreadsheet id>            # build.json: every call in order
python norde/sheet.py <spreadsheet id> --step N   # one call
```
Create an empty Google Sheet with the Drive connector, then relay the calls through the Google Sheets connector,
paste the DAX history into `actuals`, and read `Summary!A28:D38` until every check is TRUE.
