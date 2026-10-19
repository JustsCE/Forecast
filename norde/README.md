# NORDE forecast

Monthly P&L forecast for the Horizon ledger, driven from chat: the LLM (GoodHealth MCP) pulls actuals with one DAX,
submits forecast values, and the Lambda writes `NORDE-forecast.xlsx`. One Lambda, four actions, pure Python + openpyxl.

## Taxonomy

| Term | Definition |
|---|---|
| `line` | P&L leaf line, `HOR Financials[PNL_LINE]` verbatim (`1.6 Ieņēmumi no servisa pakalpojuma`); statement order = code order |
| `entity` | `dim Strukturvienibas[ACCOUNT_NR]` (`02.1`); belongs to one `division` (`ACCOUNT_DIVISION`) — see `ENTITIES` in `actuals.py` |
| `month` | `YYYY-MM` |
| `amount` | EUR, ledger sign (revenue +, expense −), the `[amount]` measure with `Kontu reformācija` (year-end close) excluded |
| `current` | max month in actuals (partial) — the first forecast month |
| `horizon` | `current` .. Dec t+1, t = year of `current` |
| `trend` | months up to `current − 2`: the last posted month may still be filling (salaries, utilities post with a lag) |
| `yoy` | per line (entities summed): `0.4·YTD + 0.3·M + 0.2·M-1 + 0.1·M-2` over the trend months, each term `cur/prev − 1` (0 when prev = 0); line level because entity splits get reclassed year to year |
| `baseline` | `(1 + yoy) × entity LY(month)`; LY of a t+1 month is the t value (actual or forecast) |
| `forecast` | submitted amount if any, else baseline |

Totals `KOPĀ IEŅĒMUMI`, `KOPĀ IZDEVUMI`, `PEĻŅA` and division/company totals are derived (Excel formulas).

## Files

| File | Purpose |
|---|---|
| `actuals.py` | `ENTITIES`, `DIVISIONS`, the DAX, `parse` (DAX result → rows), `refresh` (replace months ≥ earliest incoming) |
| `forecast.py` | `yoy`, `forecast` (baseline + submissions), `submit` (chat rows → entity grain; division/TOTAL split pro rata) |
| `workbook.py` | the xlsx: `Summary` matrix + one sheet per line |
| `lambda_function.py` | action dispatch + store (`S3_BUCKET`/`S3_PREFIX`, else local `forecast/`) |
| `test_forecast.py` | `pytest norde/test_forecast.py` |
| `deploy.sh` | CloudShell: vendor openpyxl, zip, create/update Lambda + role, smoke `dax` |

Store: `actuals.csv` (line, entity, month, amount) · `forecast.csv` append-only submissions (entity, line, month, amount, ts, note) · `NORDE-forecast.xlsx`.

## Actions

```json
{"action": "dax", "since": "2025-01"}
→ {"dax": "EVALUATE SELECTCOLUMNS(...)"}            run it with powerbi_run_dax (default since = Jan t−1)

{"action": "actuals", "dax": <powerbi_run_dax result verbatim>}
→ {"rows": 2840, "months": ["2025-01", "2026-10"]}   months ≥ the earliest pulled month are replaced

{"action": "submit", "note": "CFO call 19.10", "rows": [
  {"entity": "Serviss",  "line": "1.6", "month": "2026-11", "amount": 160000},
  {"entity": "02.1",     "line": "1.1", "month": "2026-12", "yoy": 0.05},
  {"entity": "TOTAL",    "line": "2.3", "month": "2027-01", "amount": -300000}]}
→ {"submitted": [{entity, line, month, amount, ts, note}, …], "build": {…}}
   entity = entity code | division name | TOTAL; line = code or full text; yoy resolves against LY;
   a division/TOTAL amount is split over its entities in proportion to their current forecast; any error → {"error": …}, nothing written

{"action": "build"}
→ {"key": ".../NORDE-forecast.xlsx", "current": "2026-10", "horizon": ["2026-10", "2027-12"], "submissions": 3}
```

Daily use: `dax` with `since` = current − 1 → `actuals` → `build`. Full reload: `dax` without `since`.

## Workbook

- **Summary** — `P&L line | FY t−1 | Jan..Dec t | FY t | FY % | Jan..Dec t+1 | FY t+1 | FY %`, company level, lines in
  statement order, totals as formulas, forecast months blue, FY % red < −10 % / green > +10 %.
- **One sheet per line** — rows Jan t .. Dec t+1 + `TOTAL t` / `TOTAL t+1`; column blocks
  `TOTAL | TOTAL <division> | <entity>… | TOTAL <division> | …`, each `Amount | LY | YoY %`.
  Entity amounts are values (forecast months blue); division and company totals are `=` sums; YoY is a formula with
  red/green rules; entity columns are grouped (collapsed) under their division total.

## Deploy

```
git clone <repo> && bash Forecast/norde/deploy.sh          # CloudShell; S3_BUCKET / S3_PREFIX env override the defaults
```
Defaults: bucket `goodhealth.lv`, prefix `agents/norde/justs.ceplis/working/forecast/` (Norde's working lane).

Local: `cd norde && python3 lambda_function.py '{"action":"build"}'` reads/writes `./forecast/`.

Not built (yet): actual-vs-forecast error sheet (`forecast.csv` keeps every submission for it), GoodHealth task wrapper, SharePoint upload.
