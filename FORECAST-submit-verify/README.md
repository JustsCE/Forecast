# FORECAST-submit-verify

Validates uploaded forecast Excel files before they are processed. Supports three file types: shop-level (`submit_shop_*.xlsx`), shoptype-level (`submit_shoptype_*.xlsx`), and region-level (`submit_region_*.xlsx`).

## Runtime

| Property | Value |
|----------|-------|
| Runtime | Python 3.12 |
| Handler | `lambda_function.lambda_handler` |
| Timeout | 120 s |
| Memory | 3000 MB |

## How It Works

1. **File acquisition:** If called from Slack (`user_id` present), downloads the most recent matching file from the Slack channel and uploads it to S3 with a timestamp. Otherwise accepts an `s3_key` directly.
2. **Configuration:** Reads `current_month` from `Forecast_Input.xlsx` info sheet.
3. **Validation (shop files):**
   - Forecast for current month must not be lower than actuals.
   - Format row totals must match the Forecast row when both exist.
   - No gaps within the forecasted month range.
4. **Validation (shoptype files):**
   - Partial coverage check: shops with some months filled but others marked `na`.
   - Forecast >= actuals for current month.
5. **Region files:** Pass through without validation.
6. **On failure:** Posts grouped error report to Slack, deletes the failed file from S3.
7. **On success:** Posts a summary (per-product forecast totals, YoY %, delta vs previous version), then invokes `FORECAST-submit` to process the file.

## S3 Data

- **Reads:** `Forecast/Forecast_Input.xlsx`, `Forecast/Seperate Forecasts/seperate_forecasts_combined.csv`
- **Writes:** `Forecast/Seperate Forecasts/submit_*_{timestamp}.xlsx`

## Blue Cell Convention

Forecast values are identified by cells with blue fill (`#DCE6F1`). Only blue cells are treated as user-submitted forecast data.

## Dependencies

- `boto3`, `openpyxl`, `pandas`
- `urllib.request` (Slack API)
- SSM Parameter: `/forecast/slack-bot-token`
