import json
import base64
import urllib.parse
import urllib.request
import boto3

lambda_client = boto3.client("lambda")

WORKER_LAMBDAS = {
    "/review_shop":        "arn:aws:lambda:eu-central-1:497892281264:function:FORECAST-review",
    "/review_shoptype":    "arn:aws:lambda:eu-central-1:497892281264:function:FORECAST-review-shoptype",
    "/review_region":      "arn:aws:lambda:eu-central-1:497892281264:function:FORECAST-review-region",
    "/submit_forecast":    "arn:aws:lambda:eu-central-1:497892281264:function:FORECAST-submit-verify",
    "/calculate_forecast": "arn:aws:lambda:eu-central-1:497892281264:function:FORECAST-calculate",
    "/verify_forecast":    "arn:aws:lambda:eu-central-1:497892281264:function:FORECAST-verify",
}

VERIFY_LAMBDA = "arn:aws:lambda:eu-central-1:497892281264:function:FORECAST-review-trigger-verify"


def parse_slash(event):
    body = event.get("body") or ""
    if event.get("isBase64Encoded"):
        body = base64.b64decode(body).decode("utf-8")
    p = urllib.parse.parse_qs(body)
    return {
        "command":      p.get("command",      [""])[0].strip(),
        "text":         p.get("text",         [""])[0].strip(),
        "user_name":    p.get("user_name",    [""])[0].strip(),
        "user_id":      p.get("user_id",      [""])[0].strip(),
        "channel_id":   p.get("channel_id",   [""])[0].strip(),
        "response_url": p.get("response_url", [""])[0].strip(),
    }


def post_to_slack_response_url(response_url, text):
    payload = {"response_type": "in_channel", "text": text}
    req = urllib.request.Request(
        response_url,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    urllib.request.urlopen(req, timeout=5).read()


def lambda_handler(event, context):
    d = parse_slash(event)
    command      = d["command"]
    text         = d["text"]
    user_id      = d["user_id"]
    channel_id   = d["channel_id"]
    response_url = d["response_url"]

    target_lambda = WORKER_LAMBDAS.get(command)
    if not target_lambda:
        return {
            "statusCode": 200,
            "headers": {"Content-Type": "application/json"},
            "body": json.dumps({"response_type": "ephemeral", "text": f"Unknown command: {command}"}),
        }

    if command == "/submit_forecast":
        message_text = f"<@{user_id}> submitted a forecast file"
        post_to_slack_response_url(response_url, message_text)
        lambda_client.invoke(
            FunctionName=target_lambda,
            InvocationType="Event",
            Payload=json.dumps({
                "user_id": user_id, "channel_id": channel_id,
                "response_url": response_url,
            }).encode("utf-8"),
        )
    elif command == "/calculate_forecast":
        message_text = f"<@{user_id}> started Calculating the full forecast..."
        post_to_slack_response_url(response_url, message_text)
        lambda_client.invoke(
            FunctionName=target_lambda,
            InvocationType="Event",
            Payload=json.dumps({"response_url": response_url}).encode("utf-8"),
        )
    elif command == "/verify_forecast":
        message_text = f"<@{user_id}> requested forecast status..."
        post_to_slack_response_url(response_url, message_text)
        lambda_client.invoke(
            FunctionName=target_lambda,
            InvocationType="Event",
            Payload=json.dumps({"response_url": response_url}).encode("utf-8"),
        )
    elif command in ("/review_shop", "/review_shoptype", "/review_region"):
        if not text:
            return {
                "statusCode": 200,
                "headers": {"Content-Type": "application/json"},
                "body": json.dumps({"response_type": "ephemeral",
                                     "text": f"Please provide a name. Usage: `{command} <name>`"}),
            }

        if command == "/review_shop":
            # Shop-level: validate name via verify Lambda
            lambda_client.invoke(
                FunctionName=VERIFY_LAMBDA,
                InvocationType="Event",
                Payload=json.dumps({
                    "command":       command,
                    "text":          text,
                    "user_id":       user_id,
                    "response_url":  response_url,
                    "target_lambda": target_lambda,
                }).encode("utf-8"),
            )
        else:
            # Region and shoptype: worker lambdas handle their own validation/normalization
            post_to_slack_response_url(response_url, f"<@{user_id}> requested {command.lstrip('/')} for `{text}`...")
            lambda_client.invoke(
                FunctionName=target_lambda,
                InvocationType="Event",
                Payload=json.dumps({"text": text}).encode("utf-8"),
            )
    else:
        post_to_slack_response_url(response_url, f"<@{user_id}> requested {text}")
        lambda_client.invoke(
            FunctionName=target_lambda,
            InvocationType="Event",
            Payload=json.dumps({"text": text}).encode("utf-8"),
        )

    return {
        "statusCode": 200,
        "headers": {"Content-Type": "application/json"},
        "body": json.dumps({"response_type": "ephemeral", "text": "Working on it..."}),
    }
