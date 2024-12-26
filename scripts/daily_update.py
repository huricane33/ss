import os
import pandas as pd
from sqlalchemy import create_engine, text
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google.oauth2.service_account import Credentials

FOLDER_ID = "YOUR_DRIVE_FOLDER_ID"
DATABASE_URL = os.getenv("DATABASE_URL", "postgresql://user:pass@host:port/dbname")

def get_drive_service():
    SCOPES = ['https://www.googleapis.com/auth/drive']
    creds = Credentials.from_service_account_file('service_account.json', scopes=SCOPES)
    return build('drive', 'v3', credentials=creds)

def download_latest_file(folder_id, destination):
    service = get_drive_service()
    # List newest file
    results = service.files().list(
        q=f"'{folder_id}' in parents",
        orderBy="createdTime desc",
        pageSize=1,
        fields="files(id, name, createdTime)"
    ).execute()
    files = results.get('files', [])
    if not files:
        print("No files in folder.")
        return None
    latest = files[0]
    file_id = latest["id"]
    request = service.files().get_media(fileId=file_id)
    fh = open(destination, 'wb')
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()
    print(f"Downloaded {latest['name']} to {destination}")
    return destination

def upsert_products(df):
    engine = create_engine(DATABASE_URL)

    # Simple row-by-row approach
    sql = text("""
        INSERT INTO products (product_id, product_name, vendor_name, category, barcode)
        VALUES (:pid, :pname, :vname, :cat, :bc)
        ON CONFLICT (product_id)
        DO UPDATE SET
            product_name = EXCLUDED.product_name,
            vendor_name = EXCLUDED.vendor_name,
            category = EXCLUDED.category,
            barcode = EXCLUDED.barcode;
    """)

    with engine.begin() as conn:
        for _, row in df.iterrows():
            conn.execute(sql, {
                "pid": row["product_id"],
                "pname": row["product_name"],
                "vname": row["vendor_name"],
                "cat": row["category"],
                "bc": row.get("barcode", "")
            })

def main():
    # 1. Download the latest product file
    local_file = download_latest_file(FOLDER_ID, "daily_products.xlsx")
    if not local_file:
        return

    # 2. Read the file from row 9 onward.
    #    The first 8 rows are skipped, so row 9 becomes the header row for the DataFrame.
    df = pd.read_excel(local_file, skiprows=8)

    # 3. Rename the columns:
    #    C -> barcode
    #    D -> product_id
    #    E -> product_name
    #    I -> category
    #    L -> vendor_name
    df.rename(columns={
        "C": "barcode",
        "D": "product_id",
        "E": "product_name",
        "I": "category",
        "L": "vendor_name"
    }, inplace=True)

    # 4. Upsert into the DB
    upsert_products(df)
    print("Daily product update complete!")

if __name__ == "__main__":
    main()