"""
download_from_drive.py
──────────────────────
Downloads existing Master_File.xlsx and Raw_Data.xlsx from Google Drive
BEFORE the screener runs — so the script can append to them correctly.
"""

import os
import json

IS_GITHUB_ACTIONS = os.environ.get("GITHUB_ACTIONS", "false").lower() == "true"
if IS_GITHUB_ACTIONS:
    FOLDER_PATH = "/tmp/screener_data"
else:
    FOLDER_PATH = os.path.join(os.path.dirname(__file__), "data")
os.makedirs(FOLDER_PATH, exist_ok=True)

creds_json_str = os.environ.get("GOOGLE_CREDENTIALS", "")
gdrive_folder  = os.environ.get("GDRIVE_FOLDER_ID", "")

if not creds_json_str:
    print("⚠️ GOOGLE_CREDENTIALS not set — skipping download.")
    exit(0)

try:
    from google.oauth2.service_account import Credentials
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseDownload
    import io

    creds = Credentials.from_service_account_info(
        json.loads(creds_json_str),
        scopes=["https://www.googleapis.com/auth/drive"]
    )
    service = build("drive", "v3", credentials=creds)

    files_to_download = ["Master_File.xlsx", "Raw_Data.xlsx"]

    for filename in files_to_download:
        query = f"name='{filename}' and trashed=false"
        if gdrive_folder:
            query += f" and '{gdrive_folder}' in parents"

        results  = service.files().list(q=query, fields="files(id, name)").execute()
        existing = results.get("files", [])

        if existing:
            file_id   = existing[0]["id"]
            request   = service.files().get_media(fileId=file_id)
            fh        = io.FileIO(os.path.join(FOLDER_PATH, filename), "wb")
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done:
                _, done = downloader.next_chunk()
            print(f"✅ Downloaded '{filename}' from Google Drive.")
        else:
            print(f"ℹ️  '{filename}' not found in Drive — will be created fresh.")

except Exception as e:
    print(f"⚠️ Download failed: {e}")
