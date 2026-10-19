#!/usr/bin/env bash
# CloudShell:  git clone <repo> && bash Forecast/norde/deploy.sh
# Creates or updates Lambda NORDE-forecast (python3.12, openpyxl vendored, S3 store), then smoke-tests it.
set -euo pipefail
cd "$(dirname "$0")"
FN=NORDE-forecast
BUCKET=${S3_BUCKET:-goodhealth.lv}
PREFIX=${S3_PREFIX:-agents/norde/justs.ceplis/working/forecast/}
ROLE=arn:aws:iam::$(aws sts get-caller-identity --query Account --output text):role/$FN

rm -rf build && mkdir build
pip install -q openpyxl -t build
cp actuals.py forecast.py workbook.py lambda_function.py build/
(cd build && zip -qr ../function.zip .)

if ! aws iam get-role --role-name $FN >/dev/null 2>&1; then
  aws iam create-role --role-name $FN --assume-role-policy-document \
    '{"Version":"2012-10-17","Statement":[{"Effect":"Allow","Principal":{"Service":"lambda.amazonaws.com"},"Action":"sts:AssumeRole"}]}' >/dev/null
  aws iam attach-role-policy --role-name $FN --policy-arn arn:aws:iam::aws:policy/service-role/AWSLambdaBasicExecutionRole
  aws iam put-role-policy --role-name $FN --policy-name s3 --policy-document \
    "{\"Version\":\"2012-10-17\",\"Statement\":[{\"Effect\":\"Allow\",\"Action\":[\"s3:GetObject\",\"s3:PutObject\"],\"Resource\":\"arn:aws:s3:::$BUCKET/$PREFIX*\"}]}"
  sleep 10
fi

if aws lambda get-function --function-name $FN >/dev/null 2>&1; then
  aws lambda update-function-code --function-name $FN --zip-file fileb://function.zip >/dev/null
  aws lambda wait function-updated --function-name $FN
  aws lambda update-function-configuration --function-name $FN --environment "Variables={S3_BUCKET=$BUCKET,S3_PREFIX=$PREFIX}" >/dev/null
else
  aws lambda create-function --function-name $FN --runtime python3.12 --handler lambda_function.lambda_handler \
    --role $ROLE --zip-file fileb://function.zip --timeout 120 --memory-size 1024 \
    --environment "Variables={S3_BUCKET=$BUCKET,S3_PREFIX=$PREFIX}" >/dev/null
fi
aws lambda wait function-updated --function-name $FN
rm -rf build function.zip

aws lambda invoke --function-name $FN --cli-binary-format raw-in-base64-out --payload '{"action":"dax"}' /dev/stdout
