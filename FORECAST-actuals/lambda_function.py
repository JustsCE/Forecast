import os
import csv
import io
import time
import logging
import psycopg2
import boto3

logger = logging.getLogger()
logger.setLevel(logging.INFO)

BUCKET = "bi-automations"
KEY = "Forecast/actuals.csv"

FETCH_SIZE = 50000
# S3 multipart requires each part >= 5MB (except the last)
MIN_PART_BYTES = 6 * 1024 * 1024  # 6MB safety

QUERY = """
select 
    sum(a.quantity)::double precision as actuals,
    a.fulldate,
    dd.dayname,
    case 
        when dd.datekey in ('2026-12-28','2026-12-29','2026-12-30','2026-12-31') then '2026-53'
        when dd.datekey in ('2025-12-29','2025-12-30','2025-12-31') then '2025-53'
        when dd.datekey in ('2024-12-30','2024-12-31') then '2024-53'
        else dd."yyyy-ww" 
    end as "week",
    left(a.fulldate, 7) as "yyyy-mm",
    left(a.fulldate,4) as "year",
    a.destination_region,
    case
        when a.data_source in ('ORWO','sendmoments') then a.production_site
        when a.production_site in ('MerchRocket','Allcop','Elanders','LaserTryk','Print Logistic','VR Print','EXTERNAL EU') then 'EXTERNAL EU'
        when a.production_site = 'Wolfen' then 'ORWO'
        when a.production_site = 'PCS PL' then 'SZZ'
        when a.production_site = 'PCS CGN' then 'SZZ'
        when a.production_site = 'PCS PX' then 'USA'
        when a.production_site = 'PCS CMH' then 'USA'
        when a.production_site = 'PCS MI' then 'USA'
        else 'Other' 
    end as pcs_region,
    a.format_db,
    case when a.frame_color = 'none' then 'na' else a.frame_color end as frame_color,
    a.msf_product,
    a.shop,
    case
        WHEN lower(a.shop) = 'rossmann' then 'ORWO'
        WHEN lower(a.shop) like '%orwo%' then 'ORWO'
        when lower(a.shop) like '%sendmoments%' then 'sendmoments'
        WHEN a.shoptype IN ('D2C (New Business)','D2C (Marketplace)','D2C (Core Business)') THEN 'D2C'
        WHEN a.shoptype IN ('Reseller (API)','Reseller (B2B shop)') THEN 'Reseller'
        WHEN a.shoptype = 'Enterprise' THEN 'Enterprise'
        ELSE 'Other' 
    end as shoptype,
    sum(a.revenuenet),
    sum(a.revenuenet_ship)
from dw.v_sf_facts_groupon a
left join dw.dim_date dd on a.fulldate = dd.datekey
where a.fulldate >= '2024-01-01' and a.fulldate < current_date
group by 2,3,4,5,6,7,8,9,10,11,12,13
"""

def lambda_handler(event, context):
    s3 = boto3.client("s3")

    conn = None
    cur = None
    upload_id = None
    parts = []
    part_number = 0

    rows_written = 0
    header_written = False

    # We accumulate bytes until >= MIN_PART_BYTES, then upload a multipart part
    buffer = io.BytesIO()

    def upload_part_if_needed(force: bool = False):
        nonlocal part_number, buffer, parts
        size = buffer.tell()
        if size == 0:
            return
        if not force and size < MIN_PART_BYTES:
            return

        part_number += 1
        buffer.seek(0)
        body = buffer.read()

        t0 = time.time()
        resp = s3.upload_part(
            Bucket=BUCKET,
            Key=KEY,
            UploadId=upload_id,
            PartNumber=part_number,
            Body=body,
        )
        parts.append({"ETag": resp["ETag"], "PartNumber": part_number})
        logger.info(f"Uploaded multipart part {part_number} ({len(body)} bytes) in {time.time()-t0:.2f}s")

        buffer = io.BytesIO()

    try:
        # Start multipart upload (single final object, no leftover part files)
        try:
            resp = s3.create_multipart_upload(
                Bucket=BUCKET,
                Key=KEY,
                ContentType="text/csv",
            )
            upload_id = resp["UploadId"]
            logger.info(f"Started multipart upload: UploadId={upload_id}")
        except Exception:
            logger.exception("Failed to start S3 multipart upload")
            raise

        # Connect
        try:
            conn = psycopg2.connect(
                host=os.environ["DB_HOST"],
                port=int(os.environ.get("DB_PORT", "5439")),
                dbname=os.environ["DB_DATABASE"],
                user=os.environ["DB_USER"],
                password=os.environ["DB_PASSWORD"],
                connect_timeout=10,
                sslmode="require",
                keepalives=1,
                keepalives_idle=30,
                keepalives_interval=10,
                keepalives_count=5,
            )
        except Exception:
            logger.exception("DB connect failed")
            raise

        # Use normal cursor (more robust in Lambda)
        cur = conn.cursor()

        # Execute query (Fix 3)
        try:
            t0 = time.time()
            cur.execute(QUERY)
            logger.info(f"Query executed in {time.time() - t0:.2f}s")
        except Exception:
            logger.exception("Query failed")
            raise

        # Stream results -> CSV -> multipart
        batch = 0
        while True:
            try:
                rows = cur.fetchmany(FETCH_SIZE)
            except Exception:
                logger.exception("fetchmany failed")
                raise

            if not rows:
                break

            batch += 1
            rows_written += len(rows)

            # Build CSV for this batch (text), then write bytes into our multipart buffer
            sio = io.StringIO()
            w = csv.writer(sio)

            if not header_written:
                header = [d[0] for d in cur.description]
                w.writerow(header)
                header_written = True

            w.writerows(rows)
            data = sio.getvalue().encode("utf-8")
            buffer.write(data)

            # Upload parts as buffer grows
            upload_part_if_needed(force=False)

            if batch % 10 == 0:
                logger.info(f"Batch {batch}: total_rows={rows_written}, buffer_bytes={buffer.tell()}")

            # Optional: safety stop if Lambda is about to time out
            try:
                if context.get_remaining_time_in_millis() < 15000:
                    logger.warning("Stopping early to avoid timeout; completing upload with what we have.")
                    break
            except Exception:
                pass

        # Flush last part (can be < 5MB)
        upload_part_if_needed(force=True)

        # Complete multipart upload
        try:
            t0 = time.time()
            s3.complete_multipart_upload(
                Bucket=BUCKET,
                Key=KEY,
                UploadId=upload_id,
                MultipartUpload={"Parts": parts},
            )
            logger.info(f"Completed multipart upload in {time.time()-t0:.2f}s | parts={len(parts)}")
        except Exception:
            logger.exception("Failed to complete multipart upload")
            raise

        return {
            "rows_written": rows_written,
            "s3_path": f"s3://{BUCKET}/{KEY}",
            "parts_used_in_multipart": len(parts),
        }

    except Exception:
        # If anything fails, abort the multipart upload so S3 doesn't keep orphaned MPU state
        if upload_id:
            try:
                s3.abort_multipart_upload(Bucket=BUCKET, Key=KEY, UploadId=upload_id)
                logger.warning("Aborted multipart upload due to error")
            except Exception:
                logger.exception("Failed to abort multipart upload")
        raise

    finally:
        if cur is not None:
            try:
                cur.close()
            except Exception as e:
                logger.warning(f"Cursor close failed (ignored): {e}")
        if conn is not None:
            try:
                conn.close()
            except Exception as e:
                logger.warning(f"Connection close failed (ignored): {e}")