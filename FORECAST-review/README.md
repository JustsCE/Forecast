# FORECAST-review

Generates a shop-level review Excel workbook and uploads it to Slack. Contains a summary sheet and per-product detail sheets with actuals, forecasts, format breakdowns, and weekly trends.

## Runtime

| Property | Value |
|----------|-------|
| Runtime | Python 3.12 |
| Handler | `lambda_function.lambda_handler` |
| Timeout | 303 s |
| Memory | 5000 MB |

## How It Works

1. Parses the shop name from event (`text` or `shop` field).
2. Reads `actuals.csv` and `seperate_forecasts_combined.csv` from S3.
3. **Summary sheet:**
   - Top 15 products ranked by YTD volume (FY t-1 >= 1000).
   - Columns: `% of Total YTD`, `YTD t`, `YTD t-1`, `YTD % growth`, `FY t-1`, `FY t-2`, `FY % growth`, `Last 3w % growth`, last 3 month MoM growth.
   - "Other products" aggregation row + TOTAL row.
   - Revenue-weighted `% of Total YTD` when revenue data is available.
   - Destination row below the summary.
4. **Product sheets** (one per top product):
   - **Calculations block**: all summary metrics transposed.
   - **Last 6w block**: weekly actuals, LY actuals, YoY %.
   - **Monthly block** (M-2 to M+11): Actuals, Actuals LY, YoY %, F-YoY %, Forecast (blue cells for forecast months).
   - **Available formats**: format/frame breakdown with blue-cell forecast values from the combined CSV.
5. Uploads the workbook to Slack via `files.getUploadURLExternal` + `files.completeUploadExternal`.

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
