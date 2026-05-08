# FORECAST-review-trigger-verify

Validates shop/shoptype names before dispatching review workers. Uses S3 Select to query `actuals.csv` for known values and provides fuzzy suggestions on mismatch.

## Runtime

| Property | Value |
|----------|-------|
| Runtime | Python 3.12 |
| Handler | `lambda_function.lambda_handler` |
| Timeout | 30 s |
| Memory | 256 MB |

## How It Works

1. Receives `{command, text, user_id, response_url, target_lambda}` from `FORECAST-review-trigger`.
2. Determines the validation column:
   - `/review_shoptype` → `shoptype` column
   - `/review_shop` → `forecasted_shop` column
3. Runs S3 Select against `actuals.csv` to get distinct values for the last 2 years.
4. **Exact match** → posts acknowledgement to Slack, invokes the target worker Lambda.
5. **No match** → runs fuzzy matching (prefix, substring, midpoint) and posts suggestions to Slack.
6. **S3 Select failure** → skips validation and invokes the worker anyway.

## S3 Data

- **Reads:** `s3://bi-automations/Forecast/actuals.csv` (via S3 Select)

## Dependencies

- `boto3` (S3 Select, Lambda invocation)
- `urllib.request` (Slack posting)
