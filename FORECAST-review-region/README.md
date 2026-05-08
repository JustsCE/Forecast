# FORECAST-review-region

Generates a region-level review Excel workbook with weekly and monthly tables broken down by shoptype and individual shops. Supports `EU+RoW` and `US+CA`.

## Runtime

| Property | Value |
|----------|-------|
| Runtime | Python 3.12 |
| Handler | `lambda_function.lambda_handler` |
| Timeout | 303 s |
| Memory | 3000 MB |

## How It Works

1. Parses the region from event (normalizes `US+CA`, `EU+RoW` variants).
2. Reads `actuals.csv` and `seperate_forecasts_combined.csv` from S3.
3. **Summary sheet:**
   - Top products with YTD, FY, Last 3W metrics, MoM growth.
   - "Other products" row + TOTAL row.
   - Below: full list of non-top products with their own TOTAL.
   - Conditional formatting: red fill for YTD growth < -10%, green for > +10%.
4. **Product sheets** (one per top product):
   - **Column layout**: `TOTAL | [gap] | TOTAL shoptype1 | shop1 | shop2 | Other shoptype1 | [gap] | TOTAL shoptype2 | ...`
   - **Week table** (last 15 weeks): current year, prior year, YoY % per block.
   - **Month table** (M-2 to M+5): same structure, with forecast values for future months from the combined CSV.
   - Shoptype TOTAL columns use SUM formulas referencing shop columns.
   - Grand TOTAL column uses SUM of shoptype TOTALs.
   - TOTAL rows at bottom of each table.
   - Column grouping: shop-level columns are grouped (collapsible, hidden by default).
   - Future month actuals cells get blue fill (`#DCE6F1`).
   - Conditional YoY formatting: red < -10%, green > +10%.

## Forecast Integration

For months >= current month, values come from `seperate_forecasts_combined.csv` aggregated at three levels:
- Product total (`fc_m_total`)
- Shoptype within product (`fc_m_st`)
- Individual shop within product (`fc_m_shop`)

"Other [shoptype]" columns are computed as shoptype total minus sum of named shops.

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
