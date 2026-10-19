"""NORDE-forecast Lambda — four actions over one store (S3 when S3_BUCKET is set, else a local folder).

{"action": "dax", "since"?: "YYYY-MM"}   -> {"dax": query}                  run it with powerbi_run_dax
{"action": "actuals", "dax": <result>}    -> {"rows", "months"}              refresh actuals.csv from the result
{"action": "submit", "rows": [...]}       -> {"submitted": [...], "build"}   rows: {entity|division|TOTAL, line, month, amount|yoy, note?}
{"action": "build"}                        -> {"key", "current", "horizon"}   write NORDE-forecast.xlsx
"""
import csv
import io
import json
import os
import sys
from datetime import datetime, timezone
from pathlib import Path

import actuals as A
import forecast as F
import workbook as W

BUCKET = os.environ.get("S3_BUCKET")
PREFIX = os.environ.get("S3_PREFIX", "forecast/")
ACTUALS, FORECAST, XLSX = "actuals.csv", "forecast.csv", "NORDE-forecast.xlsx"


def read(name):
    if BUCKET:
        import boto3
        s3 = boto3.client("s3")
        try:
            return s3.get_object(Bucket=BUCKET, Key=PREFIX + name)["Body"].read()
        except s3.exceptions.NoSuchKey:
            return b""
    p = Path(PREFIX, name)
    return p.read_bytes() if p.exists() else b""


def write(name, data):
    if BUCKET:
        import boto3
        boto3.client("s3").put_object(Bucket=BUCKET, Key=PREFIX + name, Body=data)
    else:
        Path(PREFIX).mkdir(parents=True, exist_ok=True)
        Path(PREFIX, name).write_bytes(data)


def rows(name):
    return list(csv.DictReader(io.StringIO(read(name).decode())))


def table(name, data, fields):
    buf = io.StringIO()
    w = csv.DictWriter(buf, fields, lineterminator="\n")
    w.writeheader()
    w.writerows(data)
    write(name, buf.getvalue().encode())


def load():
    """actuals {(entity, line, month): amount}, subs {(entity, line, month): amount} (latest submission wins)."""
    actuals = {(r["entity"], r["line"], r["month"]): float(r["amount"]) for r in rows(ACTUALS)}
    subs = {(r["entity"], r["line"], r["month"]): float(r["amount"]) for r in rows(FORECAST)}
    return actuals, subs


def build():
    actuals, subs = load()
    write(XLSX, W.build(actuals, subs))
    current = F.current_month(actuals)
    return {"key": PREFIX + XLSX, "current": current, "horizon": [current, F.horizon(current)[-1]], "submissions": len(subs)}


def lambda_handler(event, context=None):
    action = event["action"]
    if action == "dax":
        return {"dax": A.dax(event.get("since") or f"{datetime.now(timezone.utc).year - 1}-01")}
    if action == "actuals":
        data = A.refresh(rows(ACTUALS), A.parse(event["dax"]))
        table(ACTUALS, data, ["line", "entity", "month", "amount"])
        return {"rows": len(data), "months": [min(r["month"] for r in data), max(r["month"] for r in data)]}
    if action == "submit":
        actuals, subs = load()
        try:
            out = F.submit(event["rows"], actuals, subs)
        except ValueError as e:
            return {"error": str(e)}
        ts, note = datetime.now(timezone.utc).isoformat(timespec="seconds"), event.get("note", "")
        log = rows(FORECAST) + [dict(entity=e, line=l, month=m, amount=round(a, 2), ts=ts, note=note) for e, l, m, a in out]
        table(FORECAST, log, ["entity", "line", "month", "amount", "ts", "note"])
        return {"submitted": log[-len(out):], "build": build()}
    if action == "build":
        return build()
    return {"error": f"unknown action {action!r}"}


if __name__ == "__main__":
    print(json.dumps(lambda_handler(json.loads(sys.argv[1])), ensure_ascii=False, indent=1))
