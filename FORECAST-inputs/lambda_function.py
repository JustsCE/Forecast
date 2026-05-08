import os
import json
import logging
import time
import urllib3
import boto3
from io import BytesIO, StringIO
import pandas as pd
import numpy as np

logger = logging.getLogger()
logger.setLevel(logging.INFO)

http = urllib3.PoolManager()

CLIENT_ID = os.environ["CLIENT_ID"]
CLIENT_SECRET = os.environ["CLIENT_SECRET"]
TENANT_ID = os.environ["TENANT_ID"]
SITE_HOSTNAME = os.environ["SITE_HOSTNAME"]
SITE_PATH = os.environ["SITE_PATH"]
FOLDER_PATH = os.environ["FOLDER_PATH"]
S3_BUCKET = os.environ["S3_BUCKET"]
S3_PREFIX = "Forecast/"


def get_access_token() -> str:
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = (
        f"client_id={CLIENT_ID}"
        f"&client_secret={CLIENT_SECRET}"
        f"&scope=https://graph.microsoft.com/.default"
        f"&grant_type=client_credentials"
    )
    resp = http.request(
        "POST", url, body=data,
        headers={"Content-Type": "application/x-www-form-urlencoded"},
    )
    return json.loads(resp.data.decode("utf-8"))["access_token"]


def load_forecast_input_from_s3(bucket: str, key: str):
    s3 = boto3.client("s3")
    obj = s3.get_object(Bucket=bucket, Key=key)
    excel_input_path = BytesIO(obj["Body"].read())

    forecast_products = pd.read_excel(excel_input_path, sheet_name="products")
    forecast_products.columns = forecast_products.columns.str.strip()
    excel_input_path.seek(0)

    grouped_shop = pd.read_excel(
        excel_input_path, sheet_name="adhoc", header=9, usecols="G:H"
    ).dropna()
    grouped_shop.columns = ["shop", "forecasted_shop"]
    excel_input_path.seek(0)

    clean_formats_eu = pd.read_excel(excel_input_path, sheet_name="EU formats")
    excel_input_path.seek(0)

    clean_formats_us = pd.read_excel(excel_input_path, sheet_name="US formats")
    excel_input_path.seek(0)

    clean_formats = pd.concat([clean_formats_eu, clean_formats_us], ignore_index=True)
    clean_formats.columns = clean_formats.columns.str.strip()

    ignore_formats = pd.read_excel(
        excel_input_path, sheet_name="adhoc", header=9, usecols="E"
    ).dropna()
    ignore_formats.columns = ["ignore_format"]

    return forecast_products, grouped_shop, clean_formats, ignore_formats


def apply_mappings(actuals_raw, forecast_products, grouped_shop, clean_formats, ignore_formats):
    t0 = time.time()
    df = actuals_raw.merge(
        forecast_products[["forecast_product", "msf_product"]], how="left", on="msf_product"
    )
    df["forecast_product"] = df["forecast_product"].fillna("NEW PRODUCT")
    logger.info(f"Product map: {time.time() - t0:.2f}s")

    t0 = time.time()
    df = df.merge(grouped_shop[["shop", "forecasted_shop"]], how="left", on="shop")
    df["forecasted_shop"] = df["forecasted_shop"].fillna(df["shop"])
    logger.info(f"Shop map: {time.time() - t0:.2f}s")

    t0 = time.time()
    df = df.merge(
        clean_formats[["destination", "product", "format_db", "frame", "format_clean"]],
        left_on=["destination_region", "forecast_product", "format_db", "frame_color"],
        right_on=["destination", "product", "format_db", "frame"],
        how="left",
    )
    logger.info(f"Format map: {time.time() - t0:.2f}s")

    t0 = time.time()
    ignore_mask = df["forecast_product"].isin(ignore_formats["ignore_format"])
    excluded_mask = df["forecast_product"] == "EXCLUDED PRODUCT"
    new_mask = df["forecast_product"] == "NEW PRODUCT"

    not_needed = ignore_mask | excluded_mask | new_mask
    df.loc[not_needed, "format_clean"] = "na"
    df.loc[not_needed, "frame"] = "na"

    df["format_clean"].replace("", np.nan, inplace=True)
    df["frame"].replace("", np.nan, inplace=True)
    df.loc[
        df["format_clean"].isna() | df["frame"].isna(),
        ["format_clean", "frame"],
    ] = "NEW FORMAT"
    df = df.drop(columns=["destination", "product"])
    logger.info(f"Format cleanup: {time.time() - t0:.2f}s")

    return df


def lambda_handler(event, context):
    token = get_access_token()
    headers = {"Authorization": f"Bearer {token}"}

    site_url = f"https://graph.microsoft.com/v1.0/sites/{SITE_HOSTNAME}:{SITE_PATH}"
    site_data = json.loads(http.request("GET", site_url, headers=headers).data.decode("utf-8"))
    site_id = site_data["id"]

    drives_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
    drives_data = json.loads(http.request("GET", drives_url, headers=headers).data.decode("utf-8"))
    drive_id = next((d["id"] for d in drives_data.get("value", []) if d["name"] == "Documents"), None)
    if not drive_id:
        return {"statusCode": 404, "body": "Documents drive not found"}

    file_meta_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{FOLDER_PATH}"
    file_meta = json.loads(http.request("GET", file_meta_url, headers=headers).data.decode("utf-8"))
    if "error" in file_meta:
        return {"statusCode": 404, "body": f"File not found: {file_meta['error']}"}

    file_name = file_meta.get("name", "sharepoint_file.xlsx")
    file_bytes = http.request("GET", file_meta["@microsoft.graph.downloadUrl"]).data

    s3 = boto3.client("s3")
    s3_excel_key = f"{S3_PREFIX}{file_name}"
    s3.put_object(Bucket=S3_BUCKET, Key=s3_excel_key, Body=file_bytes)
    logger.info(f"Excel uploaded to s3://{S3_BUCKET}/{s3_excel_key}")

    # Load raw actuals
    t0 = time.time()
    obj = s3.get_object(Bucket=S3_BUCKET, Key=f"{S3_PREFIX}actuals.csv")
    actuals_raw = pd.read_csv(obj["Body"])
    logger.info(f"Actuals loaded: {len(actuals_raw)} rows in {time.time() - t0:.2f}s")

    # Load mappings
    t0 = time.time()
    forecast_products, grouped_shop, clean_formats, ignore_formats = \
        load_forecast_input_from_s3(S3_BUCKET, s3_excel_key)
    logger.info(f"Excel parse: {time.time() - t0:.2f}s")

    # Apply mappings
    actuals = apply_mappings(actuals_raw, forecast_products, grouped_shop, clean_formats, ignore_formats)

    # Overwrite actuals.csv with mapped version
    t0 = time.time()
    csv_buffer = StringIO()
    actuals.to_csv(csv_buffer, index=False)
    s3.put_object(
        Bucket=S3_BUCKET,
        Key=f"{S3_PREFIX}actuals.csv",
        Body=csv_buffer.getvalue(),
    )
    logger.info(f"Mapped actuals uploaded: {len(actuals)} rows in {time.time() - t0:.2f}s")

    # Unmapped product/format checks — surface items that fell through Forecast_Input.xlsx
    t0 = time.time()
    qty_col = "actuals" if "actuals" in actuals.columns else None

    new_products_mask = actuals["forecast_product"] == "NEW PRODUCT"
    np_group_cols = ["msf_product", "shop", "destination_region"]
    np_df = actuals.loc[new_products_mask, np_group_cols + ([qty_col] if qty_col else [])]
    if qty_col:
        new_products = (
            np_df.groupby(np_group_cols, dropna=False)
            .agg(row_count=(qty_col, "size"), total_quantity=(qty_col, "sum"))
            .reset_index()
            .sort_values("total_quantity", ascending=False)
        )
    else:
        new_products = (
            np_df.groupby(np_group_cols, dropna=False)
            .size()
            .reset_index(name="row_count")
            .sort_values("row_count", ascending=False)
        )

    new_formats_mask = actuals["format_clean"] == "NEW FORMAT"
    nf_group_cols = ["destination_region", "forecast_product", "format_db", "frame_color"]
    nf_df = actuals.loc[new_formats_mask, nf_group_cols + ([qty_col] if qty_col else [])]
    if qty_col:
        new_formats = (
            nf_df.groupby(nf_group_cols, dropna=False)
            .agg(row_count=(qty_col, "size"), total_quantity=(qty_col, "sum"))
            .reset_index()
            .sort_values("total_quantity", ascending=False)
        )
    else:
        new_formats = (
            nf_df.groupby(nf_group_cols, dropna=False)
            .size()
            .reset_index(name="row_count")
            .sort_values("row_count", ascending=False)
        )

    for name, df_out in [("new_products.csv", new_products), ("new_formats.csv", new_formats)]:
        buf = StringIO()
        df_out.to_csv(buf, index=False)
        s3.put_object(
            Bucket=S3_BUCKET,
            Key=f"{S3_PREFIX}{name}",
            Body=buf.getvalue(),
        )
        logger.info(f"Wrote {name}: {len(df_out)} distinct rows")
    logger.info(f"Checks written in {time.time() - t0:.2f}s")

    return {
        "statusCode": 200,
        "body": f"Synced Excel + mapped {len(actuals)} actuals rows",
        "new_products": int(len(new_products)),
        "new_formats": int(len(new_formats)),
    }