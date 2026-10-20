# norde-forecast — the chat procedure

Runs from a normal Claude chat with two connectors: GoodHealth (`powerbi_run_dax`) and Google Sheets
(`batch_clear_values`, `update_values`, `append_values`, `get_values`). The workbook computes everything;
the chat only moves rows in and reads the matrix out.

Workbook `NORDE forecast`: spreadsheet id `1mi9r8idTYrY3nHclNaNA-gryOgerhvTMb4ZiAoX34WI`
(Drive folder `NORDE forecast`, id `1Ahpg0qX9L2ODVu1q2l0DJfxg9ZpdF5xH`).
Dataset `NORDE`: workspace `971b901d-13b2-415f-93a6-7a39a4c03ea7`, dataset `ccda6673-253e-4525-9b3b-66ab678ba6d7`.

## 1. Refresh actuals

Run the DAX from `taxonomy.py` (`dax(since)`): columns `l` line code, `e` entity, `m` month `YYYY-MM`, `a` amount.

- **Mid-month refresh** (`since` = the first day of the previous month):
  1. `batch_clear_values` ranges `["pull!A2:D"]`
  2. `update_values` range `pull!A2` values `[[l, e, m, a], …]` straight from the result rows (`m` becomes a date on write).
- **Month close** (`since` = `YYYY-01` of t−2, full history):
  1. `batch_clear_values` ranges `["actuals!A2:D", "pull!A2:D"]`
  2. `update_values` range `actuals!A2` with all rows, in chunks of 500.
- Then `append_values` range `'Data Sources'!A1` with one row:
  `[written to, "Power BI NORDE (HOR Financials)", "ccda6673-253e-4525-9b3b-66ab678ba6d7", "powerbi_run_dax", <the DAX>, "since=YYYY-MM", <run at>, <who asked>, <rows>, "none", ""]`.

`data` = `pull` months plus `actuals` months before the pull; `model!B2` (current) = the latest month present.

## 2. Submit a forecast

`append_values` range `inputs!A1` values `[[target, line, month, amount, yoy, note, ts]]`:

| field | value |
|---|---|
| target | entity code (`02.1`), division name (`Serviss`) or `TOTAL` |
| line | line code (`1.6`) |
| month | `YYYY-MM`, within current .. Dec t+1 |
| amount | EUR, ledger sign; **or** leave empty and give |
| yoy | growth vs LY as a fraction (`0.05`) — applied to every entity in the target |
| note, ts | who said what, ISO timestamp |

The last row per (target, line, month) wins (`inputs!H` keys every row); a row with both fields uses the amount.
An entity row beats a division row beats a TOTAL row. A division or TOTAL amount is split over the entities
in proportion to their LY. Nothing is deleted; to withdraw an input, append a new row with the value you want.

## 3. Read the forecast

1. `update_values` range `Summary!B1` values `[["TOTAL"]]` (or a division name, or an entity code as text, e.g. `"02.1"`).
2. `get_values` range `Summary!A2:AJ25` → render as the Horizon P&L matrix, state the scope.
3. One line in depth: `get_values` range `'1.6 Ieņēmumi no servisa pakalpojuma'!A1:CQ41`.

## 4. Verify

`get_values` range `Summary!A28:D38` with the scope at `TOTAL`: every `ok` must be TRUE (FY t−2, FY t−1 and the closed months of t: data = Summary),
`current month` must be the month you expect, `pull months` must show the refreshed window.
No `#REF!`, `#N/A` or `#ERROR!` anywhere in the Summary.

## Taxonomy

| term | definition |
|---|---|
| line | P&L leaf line by code, statement order = code order; totals are formulas |
| entity | `ACCOUNT_NR`, in one division; active entities carry postings |
| current | the latest month in `data` (partial, MTD) — the first forecast month |
| trend | months up to current−2 (the last posted month may still be filling) |
| F-YoY % | per line, all entities: 0.4·YTD % + 0.3·M % + 0.2·M−1 % + 0.1·M−2 % (weights in `model!B6:B9`) |
| baseline | (1 + F-YoY %) × entity LY; LY of a t+1 month is the t value |
| forecast | the applicable input, else the baseline |
