# FORECAST-verify

Generates an Actuals vs Forecast comparison Excel workbook at the weekly level, broken down by destination, shop, and product. Writes the result to S3.

## Runtime

| Property | Value |
|----------|-------|
| Runtime | Python 3.12 |
| Handler | `lambda_function.lambda_handler` |
| Timeout | 183 s |
| Memory | 3000 MB |

## How It Works

1. Reads `forecast.csv` from S3, filters to `source == "separate"` rows (shop-level manual forecasts).
2. Reads `actuals.csv` from S3, determines the current year (`t`).
3. Computes three aggregations grouped by `(destination_region, forecasted_shop, forecast_product, week)`:
   - **Actuals** (year t): sum of `actuals` column
   - **Actuals LY** (year t-1, week-shifted to align with t): for YoY calculations
   - **Forecast** (year t): sum of `FQTY` column
4. Produces an Excel workbook with **one sheet per product**.
5. Within each sheet, blocks for each `(destination, shop)` combo:

```
Destination: EU+RoW | Shop: sendmoments
             2026-01   2026-02   2026-03   ...
Actuals        1,200     1,400     1,100   ...
Forecast       1,300     1,350     1,150   ...
F-YoY %        8.3%     -3.6%      4.5%   ...    = Forecast / Actuals_LY - 1
YoY %         10.1%      5.2%      3.1%   ...    = Actuals / Actuals_LY - 1
Error           -100        50       -50   ...    = Actuals - Forecast
Error %        -8.3%      3.6%     -4.5%   ...    = Error / Forecast
```

6. Writes the workbook to S3.
7. Posts a text confirmation to Slack via `response_url`.

## S3 Data

- **Reads:** `Forecast/forecast.csv`, `Forecast/actuals.csv`
- **Writes:** `Forecast/verify_actuals_vs_forecast.xlsx`

## Dimensions

Comparison is grouped by: `destination_region` x `forecasted_shop` x `forecast_product` x `week`

Week columns show all ISO weeks from year t present in either dataset.

## Dependencies

- `boto3`, `pandas`, `openpyxl`
- `urllib.request` (Slack text notification)
