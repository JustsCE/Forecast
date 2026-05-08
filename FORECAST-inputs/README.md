# FORECAST-inputs

Syncs `Forecast_Input.xlsx` from SharePoint and applies product, shop, and format mappings to the raw actuals CSV.

## Runtime

| Property | Value |
|----------|-------|
| Runtime | Python 3.12 |
| Handler | `lambda_function.lambda_handler` |
| Timeout | 363 s |
| Memory | 5000 MB |

## How It Works

1. Authenticates with Microsoft Graph API using OAuth2 client credentials.
2. Downloads `Forecast_Input.xlsx` from SharePoint (`picanovagmbh.sharepoint.com/sites/BI`).
3. Uploads the Excel file to S3.
4. Loads `actuals.csv` from S3 (raw output from `FORECAST-actuals`).
5. Applies mappings from the Excel workbook:
   - **Product mapping** (`products` sheet): `msf_product` → `forecast_product`
   - **Shop grouping** (`adhoc` sheet): `shop` → `forecasted_shop`
   - **Format cleaning** (`EU formats` / `US formats` sheets): `format_db` + `frame_color` → `format_clean`
6. Overwrites `actuals.csv` with the mapped version.
7. Writes diagnostic files: `new_products.csv` (unmapped products) and `new_formats.csv` (unmapped formats).

## Environment Variables

| Variable | Description |
|----------|-------------|
| `CLIENT_ID` | Azure AD app client ID |
| `CLIENT_SECRET` | Azure AD app client secret |
| `TENANT_ID` | Azure AD tenant ID |
| `SITE_HOSTNAME` | SharePoint site hostname |
| `SITE_PATH` | SharePoint site path |
| `FOLDER_PATH` | Path to file within SharePoint |
| `S3_BUCKET` | Target S3 bucket (`bi-automations`) |

## S3 Data

- **Reads:** `Forecast/actuals.csv`
- **Writes:** `Forecast/Forecast_Input.xlsx`, `Forecast/actuals.csv` (overwritten with mappings), `Forecast/new_products.csv`, `Forecast/new_formats.csv`

## Dependencies

- `boto3`, `pandas`, `numpy`, `urllib3`
