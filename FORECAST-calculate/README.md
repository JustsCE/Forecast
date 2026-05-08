# FORECAST-calculate

The full forecast calculation pipeline. Computes forecasts for all products that are NOT covered by separate (manual) submissions, distributes them by format, PCS region, day-of-week, and shoptype, then combines with separate forecasts to produce a daily `forecast.csv`.

## Runtime

| Property | Value |
|----------|-------|
| Runtime | Python 3.12 |
| Handler | `lambda_function.lambda_handler` |
| Timeout | 303 s |
| Memory | 5000 MB |

## How It Works

### 1. Coverage Analysis
- Identifies which `(shoptype, shop, product, region)` combos are covered by separate forecasts.
- Remaining (unforecasted) combos are the target of this Lambda.

### 2. Set Growth Calculation
For remaining products, computes blended YoY growth:
- 30% YTD growth + 20% Last 3W growth + 25/15/10% last 3 month MoM growth.
- Forecast = (1 + YoY) * LY actuals.

### 3. Weekly Distribution
- Uses anchor weeks from `Forecast_Input.xlsx` to map forecast-year weeks to prior-year actuals.
- Splits yearly forecast into "fc weeks" (anchor-based) and "dist. weeks" (residual distribution).
- Special products use substitute distribution from another product's weekly pattern.

### 4. Daily Distribution
- Applies day-of-week shares from anchor week actuals.
- Distributes across PCS regions using monthly shares from `EU pcs` / `US pcs` sheets.

### 5. Format Distribution
- Computes format shares from historical actuals (last N months for non-Q4, last Q4 for Q4 periods).
- Injects expected shares for new formats from `Forecast_Input.xlsx`.

### 6. Shoptype Distribution
- Distributes remaining forecast across shoptypes proportional to historical actuals.

### 7. Separate Forecast Daily Distribution
- Takes monthly totals from `seperate_forecasts_combined.csv`.
- Distributes to weekly using anchor-week LY actuals.
- Distributes to daily using DOW shares.
- Applies PCS shares.

### 8. Output
- Combines remaining + separate forecasts.
- Prepends YTD actuals (from Jan 1 of prior year).
- Writes `forecast.csv` to S3 and SharePoint.
- Writes `remaining_forecasts.xlsx` diagnostic workbook to S3.

## S3 Data

- **Reads:** `Forecast/actuals.csv`, `Forecast/Forecast_Input.xlsx`, `Forecast/Seperate Forecasts/seperate_forecasts_combined.csv`
- **Writes:** `Forecast/forecast.csv`, `Forecast/Seperate Forecasts/remaining_forecasts.xlsx`

## SharePoint Uploads

- `forecast_{timestamp}.csv` — daily forecast
- `actuals.csv` — current year actuals with `separate_shops` column

## Environment Variables

| Variable | Description |
|----------|-------------|
| `CLIENT_ID` | Azure AD app client ID |
| `CLIENT_SECRET` | Azure AD app client secret |
| `TENANT_ID` | Azure AD tenant ID |
| `SITE_HOSTNAME` | SharePoint site hostname |
| `SITE_PATH` | SharePoint site path |
| `FOLDER_PATH` | SharePoint upload folder |

## Dependencies

- `boto3`, `pandas`, `numpy`, `openpyxl`, `dateutil`, `urllib3`
