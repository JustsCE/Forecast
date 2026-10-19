# NORDE forecast

Monthly P&L forecast for the Horizon ledger, driven from chat: the LLM (GoodHealth MCP) pulls actuals with one DAX,
submits forecast values, and the CLI writes `NORDE-forecast[-<scope>].xlsx` and publishes it into a Google Sheet, where
the formulas compute on any device. Plain Python + openpyxl over one folder — no cloud runtime; it runs wherever the
GoodHealth task runs Python.

## Taxonomy

| Term | Definition |
|---|---|
| `line` | P&L leaf line, `HOR Financials[PNL_LINE]` verbatim (`1.6 Ieņēmumi no servisa pakalpojuma`); statement order = code order |
| `entity` | `dim Strukturvienibas[ACCOUNT_NR]` (`02.1`); belongs to one `division` (`ACCOUNT_DIVISION`) — see `ENTITIES` in `actuals.py` |
| `month` | `YYYY-MM` |
| `amount` | EUR, ledger sign (revenue +, expense −), the `[amount]` measure with `Kontu reformācija` (year-end close) excluded |
| `current` | max month in actuals (partial) — the first forecast month |
| `horizon` | `current` .. Dec t+1, t = year of `current` |
| `trend` | months up to `current − 2`: the last posted month may still be filling (salaries, utilities post with a lag). YTD means Jan..current−2 everywhere and is labelled with its months (`YTD 2026 (01–08)`) |
| `metrics` | per line (entities summed): `FY t-2`, `FY t-1`, `FY t-1 %`, `YTD t`, `YTD t-1`, `YTD %`, the three trend months' `%`, and `F-YoY % = 0.4·YTD % + 0.3·M % + 0.2·M-1 % + 0.1·M-2 %`, each `% = cur/prev − 1` (0 when prev = 0); line level because entity splits get reclassed year to year |
| `baseline` | `(1 + F-YoY %) × entity LY(month)`; LY of a t+1 month is the t value (actual or forecast) |
| `scope` | the workbook's view: `TOTAL` (company: divisions with their entities) · a division (its entities) · an entity code |
| `forecast` | submitted amount if any, else baseline |

Totals `KOPĀ IEŅĒMUMI`, `KOPĀ IZDEVUMI`, `PEĻŅA` and division/company totals are derived (Excel formulas).

## Files

| File | Purpose |
|---|---|
| `actuals.py` | `ENTITIES`, `DIVISIONS`, the DAX, `parse` (DAX result → rows), `refresh` (replace months ≥ earliest incoming) |
| `forecast.py` | `metrics` (growth inputs + blended F-YoY %), `forecast` (baseline + submissions), `submit` (chat rows → entity grain; division/TOTAL split pro rata) |
| `workbook.py` | the xlsx per scope: `Summary` matrix + one sheet per line |
| `run.py` | CLI: JSON in (stdin), JSON out (stdout), files in `NORDE_DIR` (default `./forecast`); `publish` pushes an xlsx into a Google Sheet |
| `test_forecast.py` | `pytest norde/test_forecast.py` |

Folder: `actuals.csv` (line, entity, month, amount) · `forecast.csv` append-only submissions (entity, line, month, amount, ts, note) · `NORDE-forecast[-<scope>].xlsx`.

## Commands

```
python norde/run.py dax [since]              → {"dax": "EVALUATE SELECTCOLUMNS(...)"}     run it with powerbi_run_dax (default since = Jan t−2)
python norde/run.py actuals < result.json    → {"rows": 2979, "months": ["2024-01", "2026-10"]}
                                               stdin = the powerbi_run_dax result verbatim; months ≥ the earliest pulled are replaced
python norde/run.py submit  < rows.json      → {"submitted": [{entity, line, month, amount, ts, note}, …], "build": {…}}
python norde/run.py build [scope]            → {"file": "forecast/NORDE-forecast-Serviss.xlsx", "scope": "Serviss", "current": "2026-10", "horizon": [...], "submissions": 3}
python norde/run.py publish <scope> <id>     → Drive file metadata; the scope's xlsx replaces the Google Sheet <id> in place (same link)
```

`rows.json`:
```json
{"note": "CFO call 19.10", "rows": [
  {"entity": "Serviss", "line": "1.6", "month": "2026-11", "amount": 160000},
  {"entity": "02.1",    "line": "1.1", "month": "2026-12", "yoy": 0.05},
  {"entity": "TOTAL",   "line": "2.3", "month": "2027-01", "amount": -300000}]}
```
`entity` = entity code | division name | `TOTAL`; `line` = code or full text; `yoy` resolves against LY; a division/TOTAL
amount is split over its entities in proportion to their current forecast; any error → `{"error": …}`, nothing written.

Daily use: `dax` with `since` = current − 1 → `actuals` → `build`. Full reload: `dax` without `since`.

## Workbook

One workbook per scope: company `TOTAL`, a division, an entity.

- **Summary** — `P&L line | FY t−2 | FY t−1 | FY t−1 % | YTD t (mm–mm) | YTD t−1 | YTD % | F-YoY % | Jan..Dec t | FY t | FY t % | Jan..Dec t+1 | FY t+1 | FY t+1 %`,
  lines in statement order, totals and `%` columns as formulas, F-YoY % = the blended growth applied, forecast months blue,
  `%` columns red < −10 % / green > +10 %.
- **One sheet per line** — rows 1–12 the Calculations block (the line's `metrics`, all entities: what the baseline used),
  then rows Jan t .. Dec t+1 + `TOTAL t` / `TOTAL t+1` with column blocks `TOTAL | TOTAL <division> | <entity>… | …`
  (a division scope starts at its division, an entity scope is one block), each `Amount | LY | YoY % | F-YoY %`:
  YoY % on actual rows, F-YoY % on forecast rows. Entity amounts are values (forecast months blue); division and company
  totals are `=` sums; entity columns are grouped (collapsed) under their division.

## Run

Requires Python ≥ 3.8 and `openpyxl` (`google-auth` for `publish`). `NORDE_DIR` points the folder anywhere.
Inside a GoodHealth task: the LLM runs the DAX, saves the result to `result.json`, runs `actuals`, then `submit`/`build`/`publish`
as the conversation goes. The Google Sheets are the deliverable; the xlsx files are their source.

Google: create the three Sheets once (an xlsx dropped into Drive converts on open), share each with the service account's
e-mail, set `GOOGLE_SERVICE_ACCOUNT=/path/key.json`, then `publish <scope> <sheet id>` after every build. `publish` is
written against Drive `files.update` (media upload converts in place) and is not yet exercised against a live service account.

Not built (yet): actual-vs-forecast error sheet (`forecast.csv` keeps every submission for it), the GoodHealth task config, SharePoint upload.
