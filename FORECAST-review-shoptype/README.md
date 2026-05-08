# FORECAST-review-shoptype

Generates a shoptype-level review Excel workbook with dual-region summaries and per-product detail sheets including top-5 shop breakdowns.

## Runtime

| Property | Value |
|----------|-------|
| Runtime | Python 3.12 |
| Handler | `lambda_function.lambda_handler` |
| Timeout | 303 s |
| Memory | 5000 MB |

## How It Works

1. Parses the shoptype name from event (`text` or `shoptype` field).
2. Reads `actuals.csv` and `seperate_forecasts_combined.csv` from S3.
3. **Summary sheet** (two tables, one per region):
   - **EU+RoW table**: Top 15 products with YTD, FY, Last 3W metrics, MoM growth.
   - **US+CA table**: Same layout below.
   - Destination labels under each table.
4. **Product sheets** (one per top product, both regions):
   - **Calculations block**: metrics from the regional summary.
   - **Last 6w block**: weekly breakdown with YoY.
   - **Monthly block** (M-2 to M+11): Actuals, Actuals LY, YoY %, F-YoY %, Forecast.
   - **Top-5 shops** (from actuals) + additional shops from forecast CSV:
     - Each shop: Forecast formula `=IFERROR((1+F-YoY%)*Actuals LY, 0)`, Actuals LY, F-YoY %.
     - "Other [shoptype]" residual row (no blue fill, no F-YoY).
   - Product-level Forecast row = SUM of individual shop forecast rows.

## Blue Cell Convention

- Forecast cells for the current month + next 5 months get blue fill (`#DCE6F1`) with thin borders.
- Shop-level rows that are direct user input get blue styling.
- "Other [shoptype]" and product-level totals do NOT get blue styling (they are derived).

## S3 Data

- **Reads:** `Forecast/actuals.csv`, `Forecast/Seperate Forecasts/seperate_forecasts_combined.csv`

## Environment Variables

| Variable | Description |
|----------|-------------|
| `SLACK_BOT_TOKEN` | Slack bot OAuth token |
| `SLACK_CHANNEL_ID` | Target Slack channel(s), comma-separated |

## Dependencies

- `boto3`, `pandas`, `numpy`, `openpyxl`
- `urllib.request` (Slack file upload)
