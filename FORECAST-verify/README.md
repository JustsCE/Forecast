# FORECAST-verify

Returns the current forecast status to Slack: latest actuals date, anchor cutoff, and a shop-level coverage summary table.

## Runtime

| Property | Value |
|----------|-------|
| Runtime | Python 3.12 |
| Handler | `lambda_function.lambda_handler` |
| Timeout | 183 s |
| Memory | 3000 MB |

## How It Works

1. Reads `actuals.csv` to determine the latest actuals date.
2. Reads `Forecast_Input.xlsx` info sheet to get the anchor cutoff (cell C2).
3. Reads `seperate_forecasts_combined.csv` and builds a summary:
   - Filters to rows where `shoptype != forecasted_shop` (separately forecasted shops).
   - Pivots by `(product, forecasted_shop)` with `US+CA` and `EU+RoW` columns.
4. Formats a Slack message with a monospaced table showing forecast totals per shop/product/region.

## Example Output

```
📊 Forecast Status
• Actuals latest date: 2026-05-07
• Anchor cutoff (Forecast Input): 2026-19

Separate forecasts (shop ≠ shoptype):
Product    Shop           US+CA   EU+RoW
-------------------------------------
Canvas     sendmoments      120    5,400
Postcard   ORWO               0   12,000
```

## S3 Data

- **Reads:** `Forecast/actuals.csv`, `Forecast/Forecast_Input.xlsx`, `Forecast/Seperate Forecasts/seperate_forecasts_combined.csv`

## Dependencies

- `boto3`, `pandas`
- `urllib.request` (Slack posting)
- SSM Parameter: `/forecast/slack-bot-token`
