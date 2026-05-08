import re
import json
import urllib.request
from io import BytesIO
from dataclasses import dataclass
from typing import List, Tuple

import boto3
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment, Border, Side, Font
from openpyxl.utils import get_column_letter


def lambda_handler(event, context):
    @dataclass(frozen=True)
    class Config:
        bucket: str = "bi-automations"
        forecast_key: str = "Forecast/forecast.csv"
        actuals_key: str = "Forecast/actuals.csv"
        out_key: str = "Forecast/verify_actuals_vs_forecast.xlsx"

    CFG = Config()
    S3 = boto3.client("s3")

    # ── Style constants (matching review-shoptype) ────────────────────
    WHITE_FILL = PatternFill(fill_type="solid", fgColor="FFFFFF")
    THIN = Side(style="thin", color="000000")
    HEADER_BOTTOM_BORDER = Border(bottom=THIN)
    BOLD_FONT = Font(bold=True)
    LEFT_ALIGN = Alignment(horizontal="left")
    RIGHT_ALIGN = Alignment(horizontal="right")
    PERCENT_FMT = "0.0%"
    INTEGER_FMT = "#,##0"

    _WEEK_RE = re.compile(r"^(?P<y>\d{4})-(?P<w>\d{2})$")

    # ── Helpers ───────────────────────────────────────────────────────

    def parse_week_key(s: str) -> Tuple[int, int]:
        if not isinstance(s, str):
            return (-1, -1)
        m = _WEEK_RE.match(s.strip())
        return (int(m.group("y")), int(m.group("w"))) if m else (-1, -1)

    def shift_week_year(week_key: str, new_year: int) -> str:
        if not isinstance(week_key, str) or "-" not in week_key:
            return f"{new_year}-00"
        _, ww = week_key.split("-", 1)
        return f"{new_year}-{ww.zfill(2)}"

    def safe_sheet_base(name: str) -> str:
        bad = r'[:\\/?*\[\]]'
        s = re.sub(bad, " ", str(name)).strip()
        s = re.sub(r"\s+", " ", s)
        return s[:31] if s else "Sheet"

    def dedupe_sheet_name(base: str, used: set) -> str:
        name = base
        i = 2
        while name in used:
            suffix = f" {i}"
            name = (base[:31 - len(suffix)] + suffix) if len(base) + len(suffix) > 31 else base + suffix
            i += 1
        used.add(name)
        return name

    def to_num(s: pd.Series) -> pd.Series:
        return pd.to_numeric(s, errors="coerce").fillna(0).astype(float)

    def slack_reply(response_url: str, text: str):
        if not response_url:
            return
        req = urllib.request.Request(
            response_url,
            data=json.dumps({"response_type": "in_channel", "text": text}).encode(),
            headers={"Content-Type": "application/json"},
            method="POST",
        )
        urllib.request.urlopen(req, timeout=10).read()

    def get_val(agg, key):
        try:
            return float(agg.loc[key])
        except (KeyError, TypeError):
            return 0.0

    # ── Data loading ──────────────────────────────────────────────────

    response_url = event.get("response_url", "")

    try:
        # Read forecast.csv — all rows (remaining + separate)
        fc = pd.read_csv(S3.get_object(Bucket=CFG.bucket, Key=CFG.forecast_key)["Body"])
        fc["forecasted_shop"] = fc["forecasted_shop"].astype(str).str.strip()
        fc.loc[fc["forecasted_shop"].isin(["", "nan", "None"]), "forecasted_shop"] = "Unassigned"
        fc["forecasted_shop"] = fc["forecasted_shop"].fillna("Unassigned")
        fc["FQTY"] = to_num(fc["FQTY"])
        fc["iso_week"] = fc["iso_week"].astype(str).str.strip()
        fc["forecast_product"] = fc["forecast_product"].astype(str).str.strip()
        fc["destination_region"] = fc["destination_region"].astype(str).str.strip()
        print(f"[verify] forecast.csv: {len(fc)} rows")

        # Read actuals.csv
        act_raw = pd.read_csv(S3.get_object(Bucket=CFG.bucket, Key=CFG.actuals_key)["Body"])
        act_raw["fulldate"] = pd.to_datetime(act_raw["fulldate"], errors="coerce")
        act_raw = act_raw[act_raw["fulldate"].notna()].copy()
        act_raw["actuals"] = to_num(act_raw["actuals"])
        act_raw["week"] = act_raw["week"].astype(str).str.strip()
        act_raw["forecast_product"] = act_raw["forecast_product"].astype(str).str.strip()
        act_raw["destination_region"] = act_raw["destination_region"].astype(str).str.strip()
        act_raw["forecasted_shop"] = act_raw["forecasted_shop"].astype(str).str.strip()
        print(f"[verify] actuals.csv: {len(act_raw)} rows")

        # Determine year t
        t = int(act_raw["fulldate"].dt.year.max())
        print(f"[verify] year t = {t}")

        # Display weeks: union of both datasets, year t only, sorted
        fc_weeks = set(fc.loc[fc["iso_week"].str.startswith(f"{t}-"), "iso_week"].unique())
        act_weeks = set(act_raw.loc[act_raw["week"].str.startswith(f"{t}-"), "week"].unique())
        all_weeks: List[str] = sorted(fc_weeks | act_weeks, key=parse_week_key)
        print(f"[verify] display weeks: {len(all_weeks)} (fc={len(fc_weeks)}, act={len(act_weeks)})")

        if not all_weeks:
            slack_reply(response_url, "No week data found for the current year.")
            return {"ok": False, "error": "no weeks"}

        # ── Build aggregations ────────────────────────────────────────

        grp_cols = ["destination_region", "forecasted_shop", "forecast_product"]

        # Actuals year t
        act_t = act_raw[act_raw["week"].str.startswith(f"{t}-")]
        agg_actuals = (
            act_t.groupby(grp_cols + ["week"], dropna=False)["actuals"]
            .sum().astype(float)
        )

        # Actuals LY (year t-1), shift week index to year t for alignment
        act_ly = act_raw[act_raw["week"].str.startswith(f"{t - 1}-")]
        agg_actuals_ly_raw = (
            act_ly.groupby(grp_cols + ["week"], dropna=False)["actuals"]
            .sum().astype(float)
        )
        # Rebuild index with shifted weeks
        if not agg_actuals_ly_raw.empty:
            idx = agg_actuals_ly_raw.index
            new_weeks = [shift_week_year(w, t) for w in idx.get_level_values("week")]
            agg_actuals_ly = agg_actuals_ly_raw.copy()
            agg_actuals_ly.index = pd.MultiIndex.from_arrays(
                [idx.get_level_values(c) for c in grp_cols] + [new_weeks],
                names=grp_cols + ["week"],
            )
        else:
            agg_actuals_ly = agg_actuals_ly_raw

        # Forecast year t
        fc_t = fc[fc["iso_week"].str.startswith(f"{t}-")]
        agg_forecast = (
            fc_t.groupby(grp_cols + ["iso_week"], dropna=False)["FQTY"]
            .sum().astype(float)
        )
        agg_forecast.index = agg_forecast.index.rename({"iso_week": "week"})

        print(f"[verify] agg sizes: actuals={len(agg_actuals)}, actuals_ly={len(agg_actuals_ly)}, forecast={len(agg_forecast)}")

        # ── Discover products and combos ──────────────────────────────

        products = sorted(fc["forecast_product"].dropna().unique().tolist())
        print(f"[verify] products: {len(products)}")

        # Build (dest, shop) combos per product from both datasets
        fc_combos = set(
            fc_t[grp_cols].drop_duplicates()
            .apply(lambda r: (r["destination_region"], r["forecasted_shop"], r["forecast_product"]), axis=1)
        )
        act_combos = set(
            act_t[act_t["forecasted_shop"].isin(fc["forecasted_shop"].unique())]
            [grp_cols].drop_duplicates()
            .apply(lambda r: (r["destination_region"], r["forecasted_shop"], r["forecast_product"]), axis=1)
        )
        all_combos = fc_combos | act_combos

        product_combos = {}
        for dest, shop, prod in all_combos:
            product_combos.setdefault(prod, set()).add((dest, shop))

        # ── Build Excel workbook ──────────────────────────────────────

        wb = Workbook()
        # Remove default sheet
        if "Sheet" in wb.sheetnames:
            wb.remove(wb["Sheet"])

        used_names = set()

        for product in products:
            combos = product_combos.get(product, set())
            if not combos:
                continue

            sname = dedupe_sheet_name(safe_sheet_base(product), used_names)
            ws = wb.create_sheet(sname)

            current_row = 1

            for dest, shop in sorted(combos):
                # ── Header row: Destination + Shop ────────────────
                header_text = f"Destination: {dest} | Shop: {shop}"
                ws.cell(row=current_row, column=1, value=header_text).font = BOLD_FONT
                for c in range(1, len(all_weeks) + 3):
                    cell = ws.cell(row=current_row, column=c)
                    cell.fill = WHITE_FILL
                    cell.border = HEADER_BOTTOM_BORDER
                    cell.font = BOLD_FONT
                current_row += 1

                # ── Week header row ───────────────────────────────
                ws.cell(row=current_row, column=1, value="").fill = WHITE_FILL
                for wi, wk in enumerate(all_weeks):
                    c = wi + 2
                    cell = ws.cell(row=current_row, column=c, value=wk)
                    cell.font = BOLD_FONT
                    cell.border = HEADER_BOTTOM_BORDER
                    cell.alignment = RIGHT_ALIGN
                    cell.fill = WHITE_FILL
                current_row += 1

                # ── Compute values for each week ──────────────────
                actuals_vals = []
                forecast_vals = []
                fyoy_vals = []
                yoy_vals = []
                error_vals = []
                error_pct_vals = []

                for wk in all_weeks:
                    key = (dest, shop, product, wk)
                    a = get_val(agg_actuals, key)
                    f = get_val(agg_forecast, key)
                    a_ly = get_val(agg_actuals_ly, key)

                    fyoy = (f / a_ly - 1.0) if a_ly != 0 else 0.0
                    yoy = (a / a_ly - 1.0) if a_ly != 0 else 0.0
                    err = a - f
                    err_pct = (err / f) if f != 0 else 0.0

                    actuals_vals.append(a)
                    forecast_vals.append(f)
                    fyoy_vals.append(fyoy)
                    yoy_vals.append(yoy)
                    error_vals.append(err)
                    error_pct_vals.append(err_pct)

                # ── Write metric rows ─────────────────────────────
                metrics = [
                    ("Actuals",  actuals_vals,  INTEGER_FMT),
                    ("Forecast", forecast_vals, INTEGER_FMT),
                    ("F-YoY %",  fyoy_vals,     PERCENT_FMT),
                    ("YoY %",    yoy_vals,      PERCENT_FMT),
                    ("Error",    error_vals,    INTEGER_FMT),
                    ("Error %",  error_pct_vals, PERCENT_FMT),
                ]

                for label, vals, fmt in metrics:
                    cell_label = ws.cell(row=current_row, column=1, value=label)
                    cell_label.font = BOLD_FONT
                    cell_label.alignment = LEFT_ALIGN
                    cell_label.fill = WHITE_FILL

                    for wi, v in enumerate(vals):
                        c = wi + 2
                        cell = ws.cell(row=current_row, column=c, value=v)
                        cell.number_format = fmt
                        cell.alignment = RIGHT_ALIGN
                        cell.fill = WHITE_FILL

                    current_row += 1

                # ── Blank separator row ───────────────────────────
                for c in range(1, len(all_weeks) + 2):
                    ws.cell(row=current_row, column=c).fill = WHITE_FILL
                current_row += 1

            # ── Auto-width columns ────────────────────────────────────
            max_col = len(all_weeks) + 1
            for c in range(1, max_col + 1):
                max_len = 0
                for r in range(1, ws.max_row + 1):
                    v = ws.cell(row=r, column=c).value
                    if v is None:
                        continue
                    if isinstance(v, str):
                        max_len = max(max_len, len(v))
                    elif isinstance(v, (int, float)):
                        nf = ws.cell(row=r, column=c).number_format
                        if nf == PERCENT_FMT:
                            max_len = max(max_len, 7)
                        else:
                            try:
                                max_len = max(max_len, len(f"{v:,.0f}"))
                            except (ValueError, TypeError):
                                max_len = max(max_len, len(str(v)))
                ws.column_dimensions[get_column_letter(c)].width = max(10, min(18, max_len + 2))

            # Column A (labels) should be wider
            ws.column_dimensions["A"].width = max(18, ws.column_dimensions["A"].width)

        # ── Save to S3 ────────────────────────────────────────────────

        buf = BytesIO()
        wb.save(buf)
        S3.put_object(
            Bucket=CFG.bucket,
            Key=CFG.out_key,
            Body=buf.getvalue(),
            ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        print(f"[verify] Wrote {len(products)} product sheets to s3://{CFG.bucket}/{CFG.out_key}")

        return {
            "ok": True,
            "s3_key": CFG.out_key,
            "products": len(products),
            "weeks": len(all_weeks),
        }

    except Exception as e:
        print(f"[verify] ERROR: {e}")
        import traceback; traceback.print_exc()
        return {"ok": False, "error": str(e)}
