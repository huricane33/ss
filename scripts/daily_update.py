import os
import base64
import json
import pandas as pd
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from sqlalchemy import create_engine, text
import io

def get_google_creds():
    """
    Decodes the base64-encoded service account JSON from SERVICE_ACCOUNT_BASE64
    and returns a Credentials object for Google APIs.
    """
    encoded = os.getenv("SERVICE_ACCOUNT_BASE64")
    if not encoded:
        raise ValueError("Missing SERVICE_ACCOUNT_BASE64 environment variable!")
    decoded_bytes = base64.b64decode(encoded)
    service_info = json.loads(decoded_bytes.decode("utf-8"))
    creds = Credentials.from_service_account_info(service_info)
    return creds

def get_drive_service():
    """
    Builds and returns a Google Drive service client using the credentials.
    """
    creds = get_google_creds()
    return build('drive', 'v3', credentials=creds)

def download_latest_file(folder_id, destination):
    """
    Fetches the newest file in the specified Google Drive folder,
    downloads it, and saves it to 'destination' (e.g., "daily_products.xlsx").
    Returns the local file path if successful, or None if the folder is empty.
    """
    service = get_drive_service()

    # List the newest file in the folder
    results = service.files().list(
        q=f"'{folder_id}' in parents",
        orderBy="createdTime desc",
        pageSize=1,
        fields="files(id, name, createdTime)"
    ).execute()

    files = results.get('files', [])
    if not files:
        print(f"No files found in folder {folder_id}.")
        return None

    latest = files[0]  # The newest file
    file_id = latest["id"]
    file_name = latest["name"]

    print(f"Found latest file: {file_name} (ID: {file_id}). Downloading...")

    request = service.files().get_media(fileId=file_id)
    with open(destination, 'wb') as fh:
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            status, done = downloader.next_chunk()

    print(f"Downloaded '{file_name}' to '{destination}'.")
    return destination

def upsert_products(df):
    """
    Takes a DataFrame with columns:
      - product_id
      - product_name
      - vendor_name
      - category
      - barcode
    and upserts each row into the 'products' table in the PostgreSQL database
    specified by DATABASE_URL.
    """
    # Get the database URL from environment variables
    db_url = os.getenv("DATABASE_URL")
    if not db_url:
        raise ValueError("DATABASE_URL environment variable is not set!")

    # Ensure the URL uses the correct dialect prefix
    if db_url.startswith("postgres://"):
        db_url = db_url.replace("postgres://", "postgresql://", 1)

    # Create the SQLAlchemy engine
    engine = create_engine(db_url)

    # Upsert SQL with ON CONFLICT
    sql = text("""
        INSERT INTO products (product_id, product_name, vendor_name, category, barcode)
        VALUES (:pid, :pname, :vname, :cat, :bc)
        ON CONFLICT (product_id)
        DO UPDATE SET
            product_name = EXCLUDED.product_name,
            vendor_name  = EXCLUDED.vendor_name,
            category     = EXCLUDED.category,
            barcode      = EXCLUDED.barcode;
    """)

    with engine.begin() as conn:
        for _, row in df.iterrows():
            conn.execute(sql, {
                "pid": row["product_id"],
                "pname": row["product_name"],
                "vname": row["vendor_name"],
                "cat": row["category"],
                "bc":  row.get("barcode", "")
            })

def main():
    """
    Main script logic:
    1. Read FOLDER_ID from environment.
    2. Download the newest file from Google Drive, saving as 'daily_products.xlsx'.
    3. Read the .xlsx file.
    4. Rename columns to match the screenshot structure.
    5. Upsert into the DB.
    """
    folder_id = os.getenv("FOLDER_ID")
    if not folder_id:
        raise ValueError("Missing FOLDER_ID environment variable!")

    # Adjust the destination file to an XLSX file
    local_file = download_latest_file(folder_id, "daily_products.xlsx")
    if not local_file:
        print("No file downloaded. Exiting.")
        return

    # Read the XLSX file using the openpyxl engine
    df = pd.read_excel(local_file, engine="openpyxl")

    # Rename columns to match the screenshot
    column_map = {
        "BARCODE": "barcode",
        "ITEM ID": "product_id",
        "nama item": "product_name",
        "category": "category",
        "vendor_name": "vendor_name"
    }
    df.rename(columns=column_map, inplace=True)

    # Upsert to DB
    upsert_products(df)
    print("Daily product update complete!")

if __name__ == "__main__":
    main()