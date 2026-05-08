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
        input_key: str = "Forecast/Forecast_Input.xlsx"
        out_key: str = "Forecast/verify_actuals_vs_forecast.xlsx"

    CFG = Config()
    S3 = boto3.client("s3")

    # ── Style constants ───────────────────────────────────────────────
    WHITE_FILL = PatternFill(fill_type="solid", fgColor="FFFFFF")
    THIN = Side(style="thin", color="000000")
    THICK = Side(style="thick", color="000000")
    HEADER_BOTTOM_BORDER = Border(bottom=THIN)
    BOLD_FONT = Font(bold=True)
    BOLD_FONT_LARGE = Font(bold=True, size=13)
    LEFT_ALIGN = Alignment(horizontal="left")
    RIGHT_ALIGN = Alignment(horizontal="right")
    INDENT_ALIGN = Alignment(horizontal="left", indent=1)
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

    def get_val(agg, key):
        try:
            return float(agg.loc[key])
        except (KeyError, TypeError):
            return 0.0

    def write_week_header_row(ws, all_weeks, row=1):
        """Write yyyy-ww headers in row 1, starting at column B."""
        ws.cell(row=row, column=1, value="").fill = WHITE_FILL
        for wi, wk in enumerate(all_weeks):
            c = wi + 2
            cell = ws.cell(row=row, column=c, value=wk)
            cell.font = BOLD_FONT
            cell.border = HEADER_BOTTOM_BORDER
            cell.alignment = RIGHT_ALIGN
            cell.fill = WHITE_FILL

    def write_section_header(ws, current_row, label, max_c):
        """Write a bold region section header spanning the row."""
        cell = ws.cell(row=current_row, column=1, value=label)
        cell.font = BOLD_FONT_LARGE
        cell.fill = WHITE_FILL
        cell.alignment = LEFT_ALIGN
        for c in range(2, max_c + 1):
            ws.cell(row=current_row, column=c).fill = WHITE_FILL
        return current_row + 1

    def write_block(ws, current_row, label, all_weeks, actuals_vals,
                    forecast_vals, fyoy_vals, yoy_vals, error_vals, error_pct_vals):
        """Write a shop or product block (no destination prefix). Returns next row."""
        max_c = len(all_weeks) + 1

        # ── Label row (shop name or product name) ────────────────
        cell = ws.cell(row=current_row, column=1, value=label)
        cell.font = BOLD_FONT
        cell.fill = WHITE_FILL
        cell.border = HEADER_BOTTOM_BORDER
        for c in range(2, max_c + 1):
            ws.cell(row=current_row, column=c).fill = WHITE_FILL
            ws.cell(row=current_row, column=c).border = HEADER_BOTTOM_BORDER
        current_row += 1

        # ── Metric rows ──────────────────────────────────────────
        metrics = [
            ("Actuals",  actuals_vals,  INTEGER_FMT),
            ("Forecast", forecast_vals, INTEGER_FMT),
            ("F-YoY %",  fyoy_vals,     PERCENT_FMT),
            ("YoY %",    yoy_vals,      PERCENT_FMT),
            ("Error",    error_vals,    INTEGER_FMT),
            ("Error %",  error_pct_vals, PERCENT_FMT),
        ]

        for metric_label, vals, fmt in metrics:
            cell_label = ws.cell(row=current_row, column=1, value=metric_label)
            cell_label.font = BOLD_FONT
            cell_label.alignment = INDENT_ALIGN
            cell_label.fill = WHITE_FILL

            for wi, v in enumerate(vals):
                c = wi + 2
                cell = ws.cell(row=current_row, column=c, value=v)
                cell.number_format = fmt
                cell.alignment = RIGHT_ALIGN
                cell.fill = WHITE_FILL

            current_row += 1

        # ── Blank separator row ──────────────────────────────────
        for c in range(1, max_c + 1):
            ws.cell(row=current_row, column=c).fill = WHITE_FILL
        current_row += 1

        return current_row

    def apply_cutoff_border(ws, all_weeks, cutoff_week):
        """Add thick right border on the cutoff week column for all data rows."""
        if cutoff_week not in all_weeks:
            return
        col_idx = all_weeks.index(cutoff_week) + 2  # +2 because col A is labels, B is first week
        for r in range(1, ws.max_row + 1):
            cell = ws.cell(row=r, column=col_idx)
            b = cell.border or Border()
            cell.border = Border(left=b.left, right=THICK, top=b.top, bottom=b.bottom)

    def apply_week_grouping(ws, all_weeks, cutoff_week):
        """Group (collapse) all week columns before cutoff_week - 6."""
        if cutoff_week not in all_weeks:
            return
        cutoff_idx = all_weeks.index(cutoff_week)
        group_end_idx = cutoff_idx - 6  # show 6 weeks before cutoff
        if group_end_idx <= 0:
            return
        # Columns to group: from col 2 (first week) to col (group_end_idx + 1)
        for i in range(0, group_end_idx):
            col = i + 2  # week columns start at col 2
            col_dim = ws.column_dimensions[get_column_letter(col)]
            col_dim.outlineLevel = 1
            col_dim.hidden = True

    def auto_width(ws, n_weeks):
        """Set column widths based on content."""
        max_col = n_weeks + 1
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
        ws.column_dimensions["A"].width = max(18, ws.column_dimensions["A"].width)

    def compute_block_values(all_weeks, agg_act, agg_fc, agg_ly, key_prefix):
        """Compute 6 metric lists for a given key prefix."""
        actuals_vals, forecast_vals = [], []
        fyoy_vals, yoy_vals, error_vals, error_pct_vals = [], [], [], []

        for wk in all_weeks:
            key = (*key_prefix, wk)
            a = get_val(agg_act, key)
            f = get_val(agg_fc, key)
            a_ly = get_val(agg_ly, key)

            err = a - f

            if a == 0 or f == 0:
                fyoy = 0.0
                yoy = 0.0
                err_pct = 0.0
            else:
                fyoy = (f / a_ly - 1.0) if a_ly != 0 else 0.0
                yoy = (a / a_ly - 1.0) if a_ly != 0 else 0.0
                err_pct = err / f

            actuals_vals.append(a)
            forecast_vals.append(f)
            fyoy_vals.append(fyoy)
            yoy_vals.append(yoy)
            error_vals.append(err)
            error_pct_vals.append(err_pct)

        return actuals_vals, forecast_vals, fyoy_vals, yoy_vals, error_vals, error_pct_vals

    def write_blank_rows(ws, current_row, n, max_c):
        for _ in range(n):
            for c in range(1, max_c + 1):
                ws.cell(row=current_row, column=c).fill = WHITE_FILL
            current_row += 1
        return current_row

    # ── Data loading ──────────────────────────────────────────────────

    response_url = event.get("response_url", "")

    try:
        # Read Forecast_Input.xlsx info sheet → cutoff week (C2)
        input_bytes = S3.get_object(Bucket=CFG.bucket, Key=CFG.input_key)["Body"].read()
        info = pd.read_excel(BytesIO(input_bytes), sheet_name="info", header=None)
        cutoff_week = str(info.iat[1, 2]).strip()  # e.g. "2026-19"
        print(f"[verify] cutoff_week from Forecast_Input: {cutoff_week}")

        # Read forecast.csv
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

        t = int(act_raw["fulldate"].dt.year.max())
        print(f"[verify] year t = {t}")

        fc_weeks = set(fc.loc[fc["iso_week"].str.startswith(f"{t}-"), "iso_week"].unique())
        act_weeks = set(act_raw.loc[act_raw["week"].str.startswith(f"{t}-"), "week"].unique())
        all_weeks: List[str] = sorted(fc_weeks | act_weeks, key=parse_week_key)
        print(f"[verify] display weeks: {len(all_weeks)}")

        if not all_weeks:
            return {"ok": False, "error": "no weeks"}

        # ── Build shop-level aggregations ─────────────────────────────

        shop_grp = ["destination_region", "forecasted_shop", "forecast_product"]

        act_t = act_raw[act_raw["week"].str.startswith(f"{t}-")]
        agg_actuals = (
            act_t.groupby(shop_grp + ["week"], dropna=False)["actuals"]
            .sum().astype(float)
        )

        act_ly = act_raw[act_raw["week"].str.startswith(f"{t - 1}-")]
        agg_ly_raw = (
            act_ly.groupby(shop_grp + ["week"], dropna=False)["actuals"]
            .sum().astype(float)
        )
        if not agg_ly_raw.empty:
            idx = agg_ly_raw.index
            new_weeks = [shift_week_year(w, t) for w in idx.get_level_values("week")]
            agg_actuals_ly = agg_ly_raw.copy()
            agg_actuals_ly.index = pd.MultiIndex.from_arrays(
                [idx.get_level_values(c) for c in shop_grp] + [new_weeks],
                names=shop_grp + ["week"],
            )
        else:
            agg_actuals_ly = agg_ly_raw

        fc_t = fc[fc["iso_week"].str.startswith(f"{t}-")]
        agg_forecast = (
            fc_t.groupby(shop_grp + ["iso_week"], dropna=False)["FQTY"]
            .sum().astype(float)
        )
        agg_forecast.index = agg_forecast.index.rename({"iso_week": "week"})

        # ── Build product-level (summary) aggregations ────────────────

        prod_grp = ["destination_region", "forecast_product"]

        agg_act_prod = (
            act_t.groupby(prod_grp + ["week"], dropna=False)["actuals"]
            .sum().astype(float)
        )

        agg_ly_prod_raw = (
            act_ly.groupby(prod_grp + ["week"], dropna=False)["actuals"]
            .sum().astype(float)
        )
        if not agg_ly_prod_raw.empty:
            idx = agg_ly_prod_raw.index
            new_weeks = [shift_week_year(w, t) for w in idx.get_level_values("week")]
            agg_ly_prod = agg_ly_prod_raw.copy()
            agg_ly_prod.index = pd.MultiIndex.from_arrays(
                [idx.get_level_values(c) for c in prod_grp] + [new_weeks],
                names=prod_grp + ["week"],
            )
        else:
            agg_ly_prod = agg_ly_prod_raw

        agg_fc_prod = (
            fc_t.groupby(prod_grp + ["iso_week"], dropna=False)["FQTY"]
            .sum().astype(float)
        )
        agg_fc_prod.index = agg_fc_prod.index.rename({"iso_week": "week"})

        # ── Discover products and combos (forecast-only) ──────────────

        fc_combos = set(
            fc_t[shop_grp].drop_duplicates()
            .apply(lambda r: (r["destination_region"], r["forecasted_shop"], r["forecast_product"]), axis=1)
        )

        product_combos = {}
        for dest, shop, prod in fc_combos:
            product_combos.setdefault(prod, set()).add((dest, shop))

        product_dests = {}
        for prod, combos in product_combos.items():
            product_dests[prod] = sorted(set(d for d, s in combos))

        ytd_by_product = act_t.groupby("forecast_product")["actuals"].sum()
        products = sorted(
            product_combos.keys(),
            key=lambda p: ytd_by_product.get(p, 0),
            reverse=True,
        )
        print(f"[verify] products: {len(products)} (sorted by YTD)")

        # ── Build Excel workbook ──────────────────────────────────────

        wb = Workbook()
        if "Sheet" in wb.sheetnames:
            wb.remove(wb["Sheet"])

        used_names = set()
        max_c = len(all_weeks) + 1

        # ══════════════════════════════════════════════════════════════
        #  SUMMARY SHEET — product x destination, split by region
        # ══════════════════════════════════════════════════════════════

        ws_sum = wb.create_sheet(dedupe_sheet_name("Summary", used_names))
        write_week_header_row(ws_sum, all_weeks, row=1)

        current_row = 2

        # EU+RoW section
        eu_prods = [(p, "EU+RoW") for p in products if "EU+RoW" in product_dests.get(p, [])]
        if eu_prods:
            current_row = write_section_header(ws_sum, current_row, "EU+RoW", max_c)
            for product, dest in eu_prods:
                vals = compute_block_values(
                    all_weeks, agg_act_prod, agg_fc_prod, agg_ly_prod,
                    key_prefix=(dest, product),
                )
                current_row = write_block(ws_sum, current_row, product, all_weeks, *vals)

        # 10 blank rows
        if eu_prods:
            current_row = write_blank_rows(ws_sum, current_row, 10, max_c)

        # US+CA section
        us_prods = [(p, "US+CA") for p in products if "US+CA" in product_dests.get(p, [])]
        if us_prods:
            current_row = write_section_header(ws_sum, current_row, "US+CA", max_c)
            for product, dest in us_prods:
                vals = compute_block_values(
                    all_weeks, agg_act_prod, agg_fc_prod, agg_ly_prod,
                    key_prefix=(dest, product),
                )
                current_row = write_block(ws_sum, current_row, product, all_weeks, *vals)

        auto_width(ws_sum, len(all_weeks))
        apply_cutoff_border(ws_sum, all_weeks, cutoff_week)
        apply_week_grouping(ws_sum, all_weeks, cutoff_week)
        ws_sum.freeze_panes = "B2"

        # ══════════════════════════════════════════════════════════════
        #  PER-PRODUCT SHEETS — shop breakdown, split by region
        # ══════════════════════════════════════════════════════════════

        for product in products:
            combos = product_combos.get(product, set())
            if not combos:
                continue

            sname = dedupe_sheet_name(safe_sheet_base(product), used_names)
            ws = wb.create_sheet(sname)
            write_week_header_row(ws, all_weeks, row=1)

            eu_combos = sorted((d, s) for d, s in combos if d == "EU+RoW")
            us_combos = sorted((d, s) for d, s in combos if d == "US+CA")

            current_row = 2

            if eu_combos:
                current_row = write_section_header(ws, current_row, "EU+RoW", max_c)
                for dest, shop in eu_combos:
                    vals = compute_block_values(
                        all_weeks, agg_actuals, agg_forecast, agg_actuals_ly,
                        key_prefix=(dest, shop, product),
                    )
                    current_row = write_block(ws, current_row, shop, all_weeks, *vals)

            if eu_combos and us_combos:
                current_row = write_blank_rows(ws, current_row, 10, max_c)

            if us_combos:
                current_row = write_section_header(ws, current_row, "US+CA", max_c)
                for dest, shop in us_combos:
                    vals = compute_block_values(
                        all_weeks, agg_actuals, agg_forecast, agg_actuals_ly,
                        key_prefix=(dest, shop, product),
                    )
                    current_row = write_block(ws, current_row, shop, all_weeks, *vals)

            auto_width(ws, len(all_weeks))
            apply_cutoff_border(ws, all_weeks, cutoff_week)
            apply_week_grouping(ws, all_weeks, cutoff_week)
            ws.freeze_panes = "B2"

        # ── Save to S3 ────────────────────────────────────────────────

        buf = BytesIO()
        wb.save(buf)
        S3.put_object(
            Bucket=CFG.bucket,
            Key=CFG.out_key,
            Body=buf.getvalue(),
            ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        print(f"[verify] Wrote {len(products)} product sheets + Summary to s3://{CFG.bucket}/{CFG.out_key}")

        return {
            "ok": True,
            "s3_key": CFG.out_key,
            "products": len(products),
            "weeks": len(all_weeks),
            "cutoff_week": cutoff_week,
        }

    except Exception as e:
        print(f"[verify] ERROR: {e}")
        import traceback; traceback.print_exc()
        return {"ok": False, "error": str(e)}
