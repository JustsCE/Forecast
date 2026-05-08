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

1. Reads `Forecast_Input.xlsx` info sheet C2 for the cutoff week.
2. Reads `forecast.csv` from S3 (all rows — both "remaining" and "separate" sources).
3. Reads `actuals.csv` from S3, determines the current year (`t`).
4. **Actuals shop remapping (residual approach):**
   - Named shops: actuals rows where `forecasted_shop` matches a forecast shop AND `forecasted_shop != shoptype`.
   - "Other [shoptype]": computed as shoptype total actuals minus sum of named shops. Matches `FORECAST-review-region` logic.
   - This handles cases like ORWO where FORECAST-inputs groups all sub-shops into `forecasted_shop=ORWO`, but the forecast has separate `ORWO` + `Other ORWO` line items.
5. Builds aggregations grouped by `(destination_region, forecasted_shop, forecast_product, week)`:
   - **Actuals** (year t)
   - **Actuals LY** (year t-1, week-shifted to align with t)
   - **Forecast** (year t): sum of `FQTY`
6. Products sorted by YTD actuals total (descending). Only shops present in `forecast.csv` are shown.
7. Produces an Excel workbook:
   - **Summary sheet**: product × destination (no shop), split by region (EU+RoW first, 10 blank rows, then US+CA).
   - **Per-product sheets**: shop breakdown, split by region with 10-row gap.

## Excel Layout

Week headers in row 1, frozen at B2. Each block:

```
EU+RoW                                          (section header, bold 13pt)
sendmoments                                     (shop label, bold, bottom border)
  Actuals      1,200    1,400    1,100   ...     INTEGER_FMT
  Forecast     1,300    1,350    1,150   ...     INTEGER_FMT
  Actuals LY   1,100    1,330    1,070   ...     INTEGER_FMT
  F-YoY %      18.2%     1.5%     7.5%  ...     = Forecast / Actuals_LY - 1
  YoY %         9.1%     5.3%     2.8%  ...     = Actuals / Actuals_LY - 1
  Error         -100       50      -50  ...     = Actuals - Forecast
  Error %       -7.7%     3.6%    -4.3% ...     = Error / Forecast
  (blank row)
```

### Formatting

- **Pre-cutoff columns** (up to and including cutoff week): `#F2F2F2` fill, `#595959` font color.
- **Cutoff week column**: thick right border.
- **Week grouping**: all weeks before cutoff - 6 are collapsed (outline level 1).
- **% zero guards**: each metric zeroes only on its own denominator — F-YoY% is 0 only when `f == 0` or `a_ly == 0`, not when actuals are 0.

## S3 Data

- **Reads:** `Forecast/forecast.csv`, `Forecast/actuals.csv`, `Forecast/Forecast_Input.xlsx`
- **Writes:** `Forecast/verify_actuals_vs_forecast.xlsx`

## Dimensions

Comparison grouped by: `destination_region` × `forecasted_shop` × `forecast_product` × `week`

Week columns show all ISO weeks from year t present in either dataset.

## Dependencies

- `boto3`, `pandas`, `openpyxl`
