# FORECAST-review-trigger

Slack slash command dispatcher. Sits behind API Gateway and routes incoming commands to the appropriate worker Lambda.

## Runtime

| Property | Value |
|----------|-------|
| Runtime | Python 3.12 |
| Handler | `lambda_function.lambda_handler` |
| Timeout | 30 s |
| Memory | 128 MB |

## Supported Commands

| Slash Command | Worker Lambda | Payload |
|---------------|---------------|---------|
| `/review_shop <name>` | FORECAST-review-trigger-verify → FORECAST-review | `{shop: name}` |
| `/review_shoptype <name>` | FORECAST-review-shoptype | `{text: name}` |
| `/review_region <name>` | FORECAST-review-region | `{text: name}` |
| `/submit_forecast` | FORECAST-submit-verify | `{user_id, channel_id, response_url}` |
| `/calculate_forecast` | FORECAST-calculate | `{response_url}` |
| `/verify_forecast` | FORECAST-verify | `{response_url}` |

## How It Works

1. Receives `POST` from API Gateway with base64-encoded form body.
2. Parses Slack fields: `command`, `text`, `user_id`, `channel_id`, `response_url`.
3. Looks up the target worker Lambda ARN from `WORKER_LAMBDAS` map.
4. Posts an acknowledgement to Slack via `response_url`.
5. Invokes the worker Lambda asynchronously (`InvocationType=Event`).
6. Returns `200` with `"Working on it..."` to satisfy Slack's 3-second response requirement.

**Special routing for `/review_shop`:** Instead of invoking the worker directly, it first calls `FORECAST-review-trigger-verify` to validate the shop name against `actuals.csv`.

## Environment Variables

None required — worker ARNs are hardcoded.

## Dependencies

- `boto3` (Lambda invocation)
- `urllib.request` (Slack response_url posting)
