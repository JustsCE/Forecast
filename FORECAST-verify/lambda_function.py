import json
import urllib.request
import boto3
from io import BytesIO, StringIO

import pandas as pd

S3 = boto3.client("s3")
BUCKET = "bi-automations"
ACTUALS_KEY = "Forecast/actuals.csv"
INPUT_XLSX_KEY = "Forecast/Forecast_Input.xlsx"
COMBINED_KEY = "Forecast/Seperate Forecasts/seperate_forecasts_combined.csv"


def get_slack_token() -> str:
    ssm = boto3.client("ssm")
    return ssm.get_parameter(
        Name="/forecast/slack-bot-token", WithDecryption=True
    )["Parameter"]["Value"]


def post_to_channel(token: str, channel_id: str, text: str):
    req = urllib.request.Request(
        "https://slack.com/api/chat.postMessage",
        data=json.dumps({"channel": channel_id, "text": text}).encode(),
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {token}",
        },
        method="POST",
    )
    urllib.request.urlopen(req, timeout=10).read()


def slack_reply(response_url: str, text: str):
    req = urllib.request.Request(
        response_url,
        data=json.dumps({"response_type": "in_channel", "text": text}).encode(),
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    urllib.request.urlopen(req, timeout=10).read()


def send_to_slack(token: str, channel_id: str, response_url: str, text: str):
    """Send message, chunking at 3800 chars. If a chunk contains an unclosed
    code block (odd number of ``` markers), close it and reopen in next chunk."""
    def dispatch(msg):
        if channel_id:
            post_to_channel(token, channel_id, msg)
        elif response_url:
            slack_reply(response_url, msg)

    # Split into lines so we never break mid-line
    lines = text.split("\n")
    chunk_lines = []
    chunk_len   = 0
    in_code     = False

    for line in lines:
        line_len = len(line) + 1  # +1 for the \n
        # Flush if adding this line would exceed limit
        if chunk_len + line_len > 3800 and chunk_lines:
            out = "\n".join(chunk_lines)
            if in_code:
                out += "\n```"   # close before sending
            dispatch(out)
            chunk_lines = []
            chunk_len   = 0
            if in_code:
                chunk_lines.append("```")  # reopen in next chunk
                chunk_len = 4

        chunk_lines.append(line)
        chunk_len += line_len
        if line.strip() == "```":
            in_code = not in_code

    if chunk_lines:
        dispatch("\n".join(chunk_lines))


def get_max_actuals_date() -> str:
    raw = S3.get_object(Bucket=BUCKET, Key=ACTUALS_KEY)["Body"].read().decode("utf-8", "replace")
    df = pd.read_csv(StringIO(raw), usecols=["fulldate"])
    df["fulldate"] = pd.to_datetime(df["fulldate"], errors="coerce")
    max_date = df["fulldate"].max()
    return max_date.strftime("%Y-%m-%d") if pd.notna(max_date) else "unknown"


def get_anchor_cutoff() -> str:
    raw = S3.get_object(Bucket=BUCKET, Key=INPUT_XLSX_KEY)["Body"].read()
    info = pd.read_excel(BytesIO(raw), sheet_name="info", header=None)
    return str(info.iat[1, 2]).strip()  # C2 = row index 1, col index 2


def get_combined_summary() -> pd.DataFrame:
    raw = S3.get_object(Bucket=BUCKET, Key=COMBINED_KEY)["Body"].read().decode("utf-8", "replace")
    df = pd.read_csv(StringIO(raw))

    # Only include rows where the shop was separately forecasted (shoptype != forecasted_shop)
    df = df[df["shoptype"].str.strip() != df["forecasted_shop"].str.strip()]

    df["forecast"] = pd.to_numeric(df["forecast"], errors="coerce").fillna(0)

    summary = (
        df.groupby(["product", "forecasted_shop", "destination"], as_index=False)["forecast"]
        .sum()
        .pivot_table(index=["product", "forecasted_shop"], columns="destination", values="forecast", aggfunc="sum")
        .fillna(0)
        .reset_index()
    )

    for col in ["US+CA", "EU+RoW"]:
        if col not in summary.columns:
            summary[col] = 0

    summary["US+CA"]  = summary["US+CA"].round(0).astype(int)
    summary["EU+RoW"] = summary["EU+RoW"].round(0).astype(int)
    summary = summary.sort_values(["product", "forecasted_shop"])[["product", "forecasted_shop", "US+CA", "EU+RoW"]]
    return summary


def format_slack_message(max_date: str, anchor: str, summary: pd.DataFrame) -> str:
    # Format numbers first so we can measure their width
    us_vals  = [f"{v:,}" for v in summary["US+CA"]]
    eu_vals  = [f"{v:,}" for v in summary["EU+RoW"]]

    # Dynamic column widths based on actual content
    w_product = max(len("Product"), summary["product"].astype(str).str.len().max())
    w_shop    = max(len("Shop"),    summary["forecasted_shop"].astype(str).str.len().max())
    w_us      = max(len("US+CA"),   max(len(v) for v in us_vals))
    w_eu      = max(len("EU+RoW"),  max(len(v) for v in eu_vals))

    def row_str(product, shop, us, eu):
        return f"{product:<{w_product}}  {shop:<{w_shop}}  {us:>{w_us}}  {eu:>{w_eu}}"

    header  = row_str("Product", "Shop", "US+CA", "EU+RoW")
    divider = "-" * len(header)

    rows = [header, divider]
    for (_, r), us, eu in zip(summary.iterrows(), us_vals, eu_vals):
        rows.append(row_str(str(r["product"]), str(r["forecasted_shop"]), us, eu))

    table = "\n".join(rows)

    lines = [
        "*📊 Forecast Status*",
        f"• *Actuals latest date:* `{max_date}`",
        f"• *Anchor cutoff (Forecast Input):* `{anchor}`",
        "",
        "*Separate forecasts (shop \u2260 shoptype):*",
        "```",
        table,
        "```",
    ]
    return "\n".join(lines)


def lambda_handler(event, context):
    channel_id   = event.get("channel_id", "")
    response_url = event.get("response_url", "")

    token = get_slack_token()

    try:
        max_date = get_max_actuals_date()
        anchor   = get_anchor_cutoff()
        summary  = get_combined_summary()
        message  = format_slack_message(max_date, anchor, summary)
    except Exception as e:
        message = f"❌ Error building forecast status: {e}"

    send_to_slack(token, channel_id, response_url, message)

    return {"statusCode": 200, "body": message}