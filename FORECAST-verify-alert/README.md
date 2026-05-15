# FORECAST-verify-alert

Weekly alert Lambda — runs every Tuesday morning (EventBridge).

Reads `forecast.csv` and `actuals.csv` from S3. For each product/region, checks
the last 4 completed weeks (at or before the cutoff week in `Forecast_Input.xlsx`).
Computes Error % = (Actuals − Forecast) / Forecast. Posts a Slack alert for every
product that exceeds its volume-based threshold.

## Volume tiers

| Avg weekly actuals (YTD) | Alert threshold |
|--------------------------|----------------|
| ≥ 10 000                 | 10%            |
| ≥ 1 000                  | 15%            |
| < 1 000                  | 20%            |

Override via SSM `/forecast/verify-alert/thresholds` (JSON array):
```json
[
  {"min_vol": 10000, "pct": 0.10, "label": "high"},
  {"min_vol": 1000,  "pct": 0.15, "label": "medium"},
  {"min_vol": 0,     "pct": 0.20, "label": "low"}
]
```

## Runtime

| Property | Value |
|----------|-------|
| Runtime  | Python 3.12 |
| Handler  | `lambda_function.lambda_handler` |
| Timeout  | 60 s |
| Memory   | 512 MB |

## SSM parameters

| Name | Type | Description |
|------|------|-------------|
| `/forecast/slack-bot-token` | SecureString | Slack bot OAuth token |
| `/forecast/slack-channel-id` | String | Channel ID for alerts (e.g. `C0XXXXXX`) |
| `/forecast/verify-alert/thresholds` | String | Optional JSON tier override |

## S3

- **Reads:** `Forecast/forecast.csv`, `Forecast/actuals.csv`, `Forecast/Forecast_Input.xlsx`
- **Writes:** nothing

## EventBridge schedule

```
cron(0 3 ? * TUE *)
```
Fires 03:00 UTC = 06:00 Riga (EET+3) every Tuesday.

## IAM

Needs `s3:GetObject` on `bi-automations/Forecast/*` and `ssm:GetParameter` on
`/forecast/slack-bot-token`, `/forecast/slack-channel-id`, `/forecast/verify-alert/thresholds`.
