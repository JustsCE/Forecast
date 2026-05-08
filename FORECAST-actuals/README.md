# FORECAST-actuals

Extracts raw actuals from Redshift and streams them to S3 as `actuals.csv` using multipart upload.

## Runtime

| Property | Value |
|----------|-------|
| Runtime | Python 3.12 |
| Handler | `lambda_function.lambda_handler` |
| Timeout | 363 s |
| Memory | 2000 MB |

## How It Works

1. Opens a multipart upload to `s3://bi-automations/Forecast/actuals.csv`.
2. Connects to Redshift using `psycopg2` with SSL.
3. Executes a SQL query against `dw.v_sf_facts_groupon` joined with `dw.dim_date`:
   - Aggregates `quantity` as actuals
   - Derives `week` (ISO week with year-end corrections for 2024-2026)
   - Derives `shoptype` from business rules (ORWO, sendmoments, D2C, Reseller, Enterprise)
   - Maps `pcs_region` from production site
   - Date range: `2024-01-01` to current date
4. Streams results in batches of 50,000 rows → CSV → multipart parts (min 6 MB each).
5. Completes the multipart upload. On failure, aborts to prevent orphaned S3 state.
6. Includes a safety timeout check to stop gracefully if Lambda is about to expire.

## Environment Variables

| Variable | Description |
|----------|-------------|
| `DB_HOST` | Redshift cluster endpoint |
| `DB_PORT` | Redshift port (default 5439) |
| `DB_DATABASE` | Database name |
| `DB_USER` | Database user |
| `DB_PASSWORD` | Database password |

## S3 Data

- **Writes:** `s3://bi-automations/Forecast/actuals.csv`

## Output Columns

`actuals`, `fulldate`, `dayname`, `week`, `yyyy-mm`, `year`, `destination_region`, `pcs_region`, `format_db`, `frame_color`, `msf_product`, `shop`, `shoptype`, `revenuenet`, `revenuenet_ship`

## Dependencies

- `boto3`, `psycopg2`
