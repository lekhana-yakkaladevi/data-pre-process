"""
kb_upload.py
------------
Upload clean Markdown files (produced by doc2kb.py) to S3
and optionally trigger an Amazon Bedrock KB sync.

Run doc2kb.py first to generate the _clean.md files,
then run this to push them to AWS.

INSTALL:
    pip install boto3

USAGE:
    # Upload one file
    python kb_upload.py report_clean.md --s3 s3://my-bucket/kb-docs/

    # Upload all clean files in a folder
    python kb_upload.py ./docs/*_clean.md --s3 s3://my-bucket/kb-docs/

    # Upload + trigger Bedrock KB sync
    python kb_upload.py ./docs/*_clean.md \
        --s3 s3://my-bucket/kb-docs/ \
        --kb_id abc123def456 \
        --datasource_id xyz789

    # With a specific AWS profile
    python kb_upload.py ./docs/*_clean.md \
        --s3 s3://my-bucket/kb-docs/ \
        --profile infoservices-dev

HOW TO FIND kb_id AND datasource_id:
    AWS Console → Bedrock → Knowledge Bases → your KB
    → Overview tab    : Knowledge Base ID
    → Data Sources tab: Data Source ID
"""

import argparse
import logging
import os
import sys
from pathlib import Path

logging.basicConfig(level=logging.INFO, format="%(levelname)s | %(message)s")
log = logging.getLogger(__name__)


def _get_boto_session(profile: str = None):
    try:
        import boto3
        return boto3.Session(profile_name=profile) if profile else boto3.Session()
    except ImportError:
        log.error("boto3 not installed. Run:  pip install boto3")
        sys.exit(1)


# ─────────────────────────────────────────────────────────────────────────────
# S3 UPLOAD
# ─────────────────────────────────────────────────────────────────────────────

def upload_to_s3(
    local_path: str,
    s3_uri: str,
    aws_profile: str = None,
) -> bool:
    """
    Upload a single file to S3.
    s3_uri format:  s3://bucket-name/optional/prefix/
    """
    from botocore.exceptions import BotoCoreError, ClientError

    s3_uri = s3_uri.rstrip('/')
    without_scheme = s3_uri.replace("s3://", "", 1)
    parts = without_scheme.split("/", 1)
    bucket = parts[0]
    prefix = parts[1] if len(parts) > 1 else ""
    key = f"{prefix}/{Path(local_path).name}".lstrip("/")

    try:
        session = _get_boto_session(aws_profile)
        s3 = session.client("s3")
        s3.upload_file(local_path, bucket, key)
        log.info(f"  Uploaded → s3://{bucket}/{key}")
        return True
    except (BotoCoreError, ClientError) as e:
        log.error(f"  Upload failed for {Path(local_path).name}: {e}")
        return False


# ─────────────────────────────────────────────────────────────────────────────
# BEDROCK KB SYNC
# ─────────────────────────────────────────────────────────────────────────────

def trigger_kb_sync(
    kb_id: str,
    datasource_id: str,
    aws_profile: str = None,
) -> bool:
    """
    Start a Bedrock Knowledge Base ingestion job.
    Bedrock will re-index the S3 bucket and pick up new/updated .md files.
    """
    from botocore.exceptions import BotoCoreError, ClientError

    try:
        session = _get_boto_session(aws_profile)
        client = session.client("bedrock-agent")
        response = client.start_ingestion_job(
            knowledgeBaseId=kb_id,
            dataSourceId=datasource_id,
        )
        job = response["ingestionJob"]
        log.info(f"  KB sync started | job: {job['ingestionJobId']} | status: {job['status']}")
        log.info(f"  Monitor: AWS Console → Bedrock → Knowledge Bases → {kb_id}")
        return True
    except (BotoCoreError, ClientError) as e:
        log.error(f"  KB sync failed: {e}")
        return False


# ─────────────────────────────────────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Upload _clean.md files to S3 and optionally sync Bedrock KB.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python kb_upload.py report_clean.md --s3 s3://my-bucket/kb-docs/
  python kb_upload.py ./docs/*_clean.md --s3 s3://my-bucket/kb-docs/
  python kb_upload.py ./docs/*_clean.md \\
      --s3 s3://my-bucket/kb-docs/ \\
      --kb_id abc123def456 \\
      --datasource_id xyz789 \\
      --profile infoservices-dev
        """
    )
    parser.add_argument("files", nargs="+", help="_clean.md file(s) to upload")
    parser.add_argument(
        "--s3", required=True, metavar="S3_URI",
        help="S3 URI, e.g. s3://my-bucket/kb-docs/"
    )
    parser.add_argument(
        "--kb_id", metavar="KNOWLEDGE_BASE_ID",
        help="Bedrock Knowledge Base ID (triggers sync after upload)"
    )
    parser.add_argument(
        "--datasource_id", metavar="DATA_SOURCE_ID",
        help="Bedrock Data Source ID tied to your S3 bucket"
    )
    parser.add_argument(
        "--profile", metavar="AWS_PROFILE",
        help="AWS CLI profile name (uses default if omitted)"
    )
    args = parser.parse_args()

    # Validate KB sync args
    if args.kb_id and not args.datasource_id:
        log.error("--kb_id requires --datasource_id. Both must be provided to trigger sync.")
        sys.exit(1)

    # Upload files
    uploaded = []
    for f in args.files:
        if not os.path.isfile(f):
            log.warning(f"Skipping (not found): {f}")
            continue
        if not f.endswith("_clean.md"):
            log.warning(f"Skipping (not a _clean.md file): {f}")
            continue
        success = upload_to_s3(f, s3_uri=args.s3, aws_profile=args.profile)
        if success:
            uploaded.append(f)

    print(f"\n✓ {len(uploaded)}/{len(args.files)} file(s) uploaded → {args.s3}")

    # Trigger KB sync once after all uploads (not per file)
    if uploaded and args.kb_id:
        print("Triggering Bedrock KB sync...")
        trigger_kb_sync(args.kb_id, args.datasource_id, aws_profile=args.profile)


if __name__ == "__main__":
    main()
