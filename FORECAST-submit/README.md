# FORECAST-submit

Core submission processor. Reads all submitted forecast Excel files from S3, resolves multi-source priority (region > shoptype > shop), distributes formats using historical shares, extrapolates missing months, and writes the combined forecast CSV.

## Runtime

| Property | Value |
|----------|-------|
| Runtime | Python 3.12 |
| Handler | `lambda_function.lambda_handler` |
| Timeout | 303 s |
| Memory | 5000 MB |

## How It Works

### 1. File Reading
- **Shop files** (`submit_shop_*.xlsx`): Reads product sheets, extracts blue-cell forecast + format breakdowns.
- **Shoptype files** (`submit_shoptype_*.xlsx`): Reads per-destination blocks with shop-level blue cells.
- **Region files** (`submit_region_*.xlsx`): Reads transposed layout with shoptype/shop blocks per product.

### 2. Multi-Source Priority (`build_combined`)
When the same (shop, product, destination) appears in multiple file types:
- Newest file timestamp wins at the file level.
- Period-level fallback: if the winning source lacks a specific period, data from other sources is retained.
- Shoptype totals are suppressed when all member shops have individual shop files.

### 3. Format Distribution
For rows with `format=Unknown`:
- Computes historical format shares from actuals (last N months + last Q4).
- Falls back from shop-level to shoptype-level shares.
- Injects expected shares from `Forecast_Input.xlsx` for new formats.

### 4. Monthly Extrapolation
For months without explicit forecast data:
- Computes YoY growth from forecasted months.
- Applies growth to last-year actuals (shop-level or regional fallback).
- Distributes extrapolated totals across format/frame combinations.

## S3 Data

- **Reads:** `Forecast/Seperate Forecasts/submit_*.xlsx`, `Forecast/actuals.csv`, `Forecast/Forecast_Input.xlsx`
- **Writes:** `Forecast/Seperate Forecasts/seperate_forecasts_combined.csv`

## Output CSV Fields

`forecasted_shop`, `shoptype`, `destination`, `period`, `product`, `format`, `frame`, `forecast`, `forecast_type`, `source_file`, `source_timestamp`

## Dependencies

- `boto3`, `pandas`, `numpy`, `openpyxl`, `dateutil`
- `urllib.request` (Slack reply)
- SSM Parameter: `/forecast/slack-bot-token`
