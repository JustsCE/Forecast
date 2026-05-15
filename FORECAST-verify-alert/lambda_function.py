"""FORECAST-verify-alert

Weekly alert Lambda (Tuesday morning, EventBridge).
Checks Error % per product/region for completed past weeks.
Applies volume-based thresholds. Posts to Slack if any breach.

Volume tiers (avg weekly actuals for current year):
  high:   >= 10 000  →  alert if |Error %| > 10%
  medium: >= 1 000   →  alert if |Error %| > 15%
  low:    < 1 000    →  alert if |Error %| > 20%

SSM params:
  /forecast/slack-bot-token    (SecureString) Slack bot token
  /forecast/slack-channel-id               Channel to post alerts
  /forecast/verify-alert/thresholds  (optional) JSON override of tiers
"""
import json
import datetime
import urllib.request
from io import BytesIO

import boto3
import pandas as pd

BUCKET = "bi-automations"
FORECAST_KEY = "Forecast/forecast.csv"
ACTUALS_KEY = "Forecast/actuals.csv"
INPUT_KEY = "Forecast/Forecast_Input.xlsx"

DEFAULT_TIERS = [
    {"min_vol": 10_000, "pct": 0.10, "label": "high"},
    {"min_vol": 1_000,  "pct": 0.15, "label": "medium"},
    {"min_vol": 0,      "pct": 0.20, "label": "low"},
]

WEEKS_TO_CHECK = 4  # look back up to this many completed weeks


def get_tier(avg_weekly_vol: float, tiers: list) -> dict:
    for t in tiers:
        if avg_weekly_vol >= t["min_vol"]:
            return t
    return tiers[-1]


def post_slack(token: str, channel: str, text: str):
    req = urllib.request.Request(
        "https://slack.com/api/chat.postMessage",
        data=json.dumps({"channel": channel, "text": text}).encode(),
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {token}",
        },
        method="POST",
    )
    urllib.request.urlopen(req, timeout=10).read()


def lambda_handler(event, context):
    S3 = boto3.client("s3")
    ssm = boto3.client("ssm")

    slack_token = ssm.get_parameter(
        Name="/forecast/slack-bot-token", WithDecryption=True
    )["Parameter"]["Value"]
    channel = ssm.get_parameter(
        Name="/forecast/slack-channel-id"
    )["Parameter"]["Value"]

    tiers = DEFAULT_TIERS
    try:
        raw = ssm.get_parameter(
            Name="/forecast/verify-alert/thresholds"
        )["Parameter"]["Value"]
        tiers = json.loads(raw)
    except ssm.exceptions.ParameterNotFound:
        pass

    # ── Cutoff week from Forecast_Input.xlsx ──────────────────────────
    try:
        input_bytes = S3.get_object(Bucket=BUCKET, Key=INPUT_KEY)["Body"].read()
        info = pd.read_excel(BytesIO(input_bytes), sheet_name="info", header=None)
        cutoff_week = str(info.iat[1, 2]).strip()
    except Exception as e:
        print(f"[verify-alert] could not read cutoff_week: {e}, using current week")
        iso = datetime.date.today().isocalendar()
        cutoff_week = f"{iso[0]}-{iso[1]:02d}"

    print(f"[verify-alert] cutoff_week={cutoff_week}")

    today = datetime.date.today()
    year = today.year

    # ── Load raw data ─────────────────────────────────────────────────
    fc = pd.read_csv(BytesIO(S3.get_object(Bucket=BUCKET, Key=FORECAST_KEY)["Body"].read()))
    act = pd.read_csv(BytesIO(S3.get_object(Bucket=BUCKET, Key=ACTUALS_KEY)["Body"].read()))

    fc["iso_week"] = fc["iso_week"].astype(str).str.strip()
    fc["forecast_product"] = fc["forecast_product"].astype(str).str.strip()
    fc["destination_region"] = fc["destination_region"].astype(str).str.strip()
    fc["FQTY"] = pd.to_numeric(fc["FQTY"], errors="coerce").fillna(0)

    act["week"] = act["week"].astype(str).str.strip()
    act["forecast_product"] = act["forecast_product"].astype(str).str.strip()
    act["destination_region"] = act["destination_region"].astype(str).str.strip()
    act["actuals"] = pd.to_numeric(act["actuals"], errors="coerce").fillna(0)

    fc_yr = fc[fc["iso_week"].str.startswith(f"{year}-")]
    act_yr = act[act["week"].str.startswith(f"{year}-")]

    # ── Determine weeks to check ──────────────────────────────────────
    completed = sorted(
        w for w in act_yr["week"].unique()
        if w <= cutoff_week
    )
    check_weeks = completed[-WEEKS_TO_CHECK:] if completed else []
    print(f"[verify-alert] checking weeks: {check_weeks}")

    if not check_weeks:
        print("[verify-alert] no completed weeks — nothing to check")
        return {"ok": True, "alerts": 0}

    grp = ["forecast_product", "destination_region"]

    agg_fc = (
        fc_yr.groupby(grp + ["iso_week"])["FQTY"]
        .sum()
        .rename_axis(grp + ["week"])
    )
    agg_act = (
        act_yr.groupby(grp + ["week"])["actuals"]
        .sum()
    )

    # Avg weekly actuals YTD per product/dest (for volume tier)
    avg_vol = agg_act.groupby(level=grp).mean()

    # ── Find breaches ─────────────────────────────────────────────────
    alerts = []
    checked = set()

    for prod in fc_yr["forecast_product"].unique():
        for dest in fc_yr.loc[fc_yr["forecast_product"] == prod, "destination_region"].unique():
            key = (prod, dest)
            if key in checked:
                continue
            checked.add(key)

            vol = float(avg_vol.get(key, 0))
            tier = get_tier(vol, tiers)

            for week in check_weeks:
                f = float(agg_fc.get((*key, week), 0))
                a = float(agg_act.get((*key, week), 0))
                if f == 0:
                    continue
                err_pct = (a - f) / f
                if abs(err_pct) > tier["pct"]:
                    alerts.append({
                        "product": prod,
                        "dest": dest,
                        "week": week,
                        "err_pct": err_pct,
                        "actuals": a,
                        "forecast": f,
                        "tier": tier["label"],
                        "threshold": tier["pct"],
                    })

    print(f"[verify-alert] {len(alerts)} breaches found")

    if not alerts:
        return {"ok": True, "alerts": 0}

    alerts.sort(key=lambda x: abs(x["err_pct"]), reverse=True)

    # ── Build Slack message ───────────────────────────────────────────
    lines = [f":warning: *Forecast Error Alert — {today.strftime('%b %d, %Y')}*", ""]
    for a in alerts:
        sign = "+" if a["err_pct"] > 0 else ""
        lines.append(
            f"• *{a['product']}* ({a['dest']}) · W{a['week']}: "
            f"{sign}{a['err_pct']:.1%} error "
            f"[limit: {a['threshold']:.0%}, vol: {a['tier']}]"
        )

    post_slack(slack_token, channel, "\n".join(lines))
    print(f"[verify-alert] posted to Slack")
    return {"ok": True, "alerts": len(alerts)}
