# Forecast Pipeline

AWS Lambda functions powering the Picanova demand forecast system. Triggered via Slack slash commands, these functions form a pipeline that ingests actuals from Redshift, accepts manual forecasts from Excel files, validates submissions, computes format/shoptype distributions, and produces review workbooks.

## Architecture

```
Slack slash commands
        |
  FORECAST-review-trigger  (API Gateway → Lambda dispatcher)
        |
        ├── /submit_forecast  → FORECAST-submit-verify → FORECAST-submit
        ├── /review_shop      → FORECAST-review-trigger-verify → FORECAST-review
        ├── /review_shoptype  → FORECAST-review-shoptype
        ├── /review_region    → FORECAST-review-region
        ├── /calculate_forecast → FORECAST-calculate
        └── /verify_forecast  → FORECAST-verify
```

**Data flow:**
1. **FORECAST-inputs** — Syncs `Forecast_Input.xlsx` from SharePoint, downloads actuals from Redshift via `FORECAST-actuals`, applies product/shop/format mappings.
2. **FORECAST-actuals** — Streams raw actuals from Redshift into `s3://bi-automations/Forecast/actuals.csv` using multipart upload.
3. **FORECAST-submit-verify** — Validates uploaded Excel forecast files (shop/shoptype/region), checks forecast vs actuals, format mismatches, coverage gaps.
4. **FORECAST-submit** — Processes validated files: reads all submit files from S3, resolves shop/shoptype/region priority, distributes formats, extrapolates missing months, writes combined CSV.
5. **FORECAST-review** — Generates shop-level review Excel with summary + per-product sheets (actuals, forecast, YoY, last 6 weeks).
6. **FORECAST-review-shoptype** — Generates shoptype-level review Excel with dual-region summary (EU+RoW / US+CA) and per-product sheets with top-5 shop breakdowns.
7. **FORECAST-review-region** — Generates region-level review Excel with weekly + monthly tables broken down by shoptype and shop.
8. **FORECAST-calculate** — Full pipeline: computes remaining (unforecasted) products, distributes by format/PCS/DOW/shoptype shares, produces daily `forecast.csv`, uploads to S3 + SharePoint.
9. **FORECAST-verify** — Generates an Actuals vs Forecast comparison Excel workbook at the weekly level, broken down by destination, shop, and product. Writes to S3.
10. **FORECAST-review-trigger** — API Gateway entry point that parses Slack slash commands and dispatches to worker Lambdas.
11. **FORECAST-review-trigger-verify** — Validates shop/shoptype names against actuals before dispatching review workers.

## Scheduled Trigger

A Step Functions state machine **FORECAST-steps** orchestrates the daily data refresh:

```
Start → FORECAST-actuals → FORECAST-inputs → End
```

- **State machine ARN:** `arn:aws:states:eu-central-1:497892281264:execution:FORECAST-steps`
- **IAM role:** `arn:aws:iam::497892281264:role/service-role/StepFunctions-FORECAST-steps-role-n6ik54ptu`
- **Schedule:** EventBridge Scheduler `FORECAST-actuals` — cron `0 8 * * ? *` (daily at 08:00 UTC+3 / 05:00 UTC)
- **State transitions:** 4 (Start → Actuals → Inputs → End)

This ensures `actuals.csv` is refreshed from Redshift and mappings are applied from SharePoint every morning before users interact with the forecast.

## S3 Bucket

All data lives in `s3://bi-automations/Forecast/`:
- `actuals.csv` — mapped actuals from Redshift
- `Forecast_Input.xlsx` — configuration (products, formats, shops, anchor weeks)
- `Seperate Forecasts/` — submitted Excel files + `seperate_forecasts_combined.csv`
- `forecast.csv` — final daily forecast output

## Runtime

- **Python 3.12** on all functions
- Common dependencies: `boto3`, `pandas`, `numpy`, `openpyxl`
- `FORECAST-actuals` additionally uses `psycopg2` (Redshift)
- `FORECAST-inputs` and `FORECAST-calculate` use `urllib3` (Microsoft Graph API for SharePoint)
