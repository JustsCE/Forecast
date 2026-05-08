"""
FORECAST-review-trigger-verify

Async validation Lambda invoked by FORECAST-review-trigger.
Validates shop/shoptype names against actuals.csv via S3 Select.
If valid: invokes the worker Lambda.
If invalid: posts error with suggestions to Slack via response_url.
"""
import json
import urllib.request
from datetime import datetime
import boto3

s3_client = boto3.client("s3")
lambda_client = boto3.client("lambda")

BUCKET = "bi-automations"
ACTUALS_KEY = "Forecast/actuals.csv"


def _s3_select_values(column):
    """Use S3 Select to get distinct values from actuals.csv for recent data."""
    cutoff_year = datetime.now().year - 1
    query = f'SELECT s."{column}" FROM s3object s WHERE CAST(s."year" AS INT) >= {cutoff_year}'
    try:
        resp = s3_client.select_object_content(
            Bucket=BUCKET, Key=ACTUALS_KEY,
            Expression=query,
            ExpressionType="SQL",
            InputSerialization={"CSV": {"FileHeaderInfo": "Use"}, "CompressionType": "NONE"},
            OutputSerialization={"CSV": {}},
        )
        raw = b""
        for event in resp["Payload"]:
            if "Records" in event:
                raw += event["Records"]["Payload"]
        values = set()
        for line in raw.decode("utf-8").strip().split("\n"):
            v = line.strip().strip('"')
            if v:
                values.add(v)
        return values
    except Exception:
        return set()


def _normalize(name):
    return name.lower().replace(" ", "")


def _fuzzy_suggestions(query, known_values, max_results=5):
    q = _normalize(query)
    if not q:
        return []

    scored = []
    q_prefix = q[:5]
    q_mid_start = max(0, len(q) // 2 - 2)
    q_mid = q[q_mid_start:q_mid_start + 5]

    for val in known_values:
        v = _normalize(val)
        score = 0
        if v == q:
            score = 100
        elif v.startswith(q_prefix):
            score = 60
        elif q in v:
            score = 50
        elif q_prefix in v:
            score = 40
        elif q_mid and q_mid in v:
            score = 30
        elif v[:5] and q.startswith(v[:5]):
            score = 20
        if score > 0:
            scored.append((score, val))

    scored.sort(key=lambda x: (-x[0], x[1]))
    return [val for _, val in scored[:max_results]]


def post_to_slack(response_url, text):
    payload = {"response_type": "in_channel", "text": text}
    req = urllib.request.Request(
        response_url,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    urllib.request.urlopen(req, timeout=10).read()


def lambda_handler(event, context):
    command       = event.get("command", "")
    text          = event.get("text", "")
    user_id       = event.get("user_id", "")
    response_url  = event.get("response_url", "")
    target_lambda = event.get("target_lambda", "")

    # Determine which column to validate against
    if command == "/review_shoptype":
        column = "shoptype"
        label = "shop type"
    else:
        column = "forecasted_shop"
        label = "shop"

    # Build worker payload
    if command == "/review_shoptype":
        worker_payload = {"shoptype": text}
    elif command == "/review_region":
        worker_payload = {"region": text}
    else:
        worker_payload = {"shop": text}

    # S3 Select to get known values
    known = _s3_select_values(column)

    if not known:
        # S3 Select failed — don't block, proceed to worker
        post_to_slack(response_url, f"<@{user_id}> requested {label} review of {text}")
        lambda_client.invoke(
            FunctionName=target_lambda,
            InvocationType="Event",
            Payload=json.dumps(worker_payload).encode("utf-8"),
        )
        return {"ok": True, "validation": "skipped"}

    # Exact match — proceed
    if text in known:
        post_to_slack(response_url, f"<@{user_id}> requested {label} review of {text}")
        lambda_client.invoke(
            FunctionName=target_lambda,
            InvocationType="Event",
            Payload=json.dumps(worker_payload).encode("utf-8"),
        )
        return {"ok": True, "validation": "exact_match"}

    # Fuzzy search
    suggestions = _fuzzy_suggestions(text, known)
    if suggestions:
        bullet_list = "\n".join(f"  • `{s}`" for s in suggestions)
        post_to_slack(response_url,
                      f"The {label} `{text}` wasn't found. Did you mean:\n{bullet_list}")
    else:
        post_to_slack(response_url,
                      f"The {label} `{text}` wasn't found and no similar names could be suggested.")

    return {"ok": False, "validation": "not_found", "suggestions": suggestions}
