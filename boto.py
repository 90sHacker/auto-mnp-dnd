import boto3  # pip install boto3
from pathlib import Path

# Let's use Amazon S3

fp = Path("attachments")

def upload_files(file_path):
    s3 = boto3.client("s3")

    s3.upload_file(
    Filename=file_path,
    Bucket="sample-bucket-90210",
    Key="20220602.pdf",
    )

upload_files('/Users/bolatito.shobanke/Documents/Terragon/MNP_DND/attachments/Blache_Fullstack_Software_Engineer.pdf')
