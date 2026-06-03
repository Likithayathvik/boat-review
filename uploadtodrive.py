# upload_to_drive.py — Saves a new dated file every week to Google Drive folder
import os
import json
from datetime import datetime
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

XLSX_FILE     = 'Android_HearablesApp_Review.xlsx'
DRIVE_FOLDER_ID = os.environ.get('DRIVE_FOLDER_ID', '')  # ← folder ID from secret

def upload():
    creds_json = json.loads(os.environ['GOOGLE_CREDENTIALS_JSON'])
    creds = Credentials.from_service_account_info(
        creds_json,
        scopes=['https://www.googleapis.com/auth/drive']
    )
    service = build('drive', 'v3', credentials=creds)

    # Name file with current date — e.g. Android_HearablesApp_Review_2026-06-02.xlsx
    dated_name = f"Android_HearablesApp_Review_{datetime.now().strftime('%Y-%m-%d')}.xlsx"

    media = MediaFileUpload(
        XLSX_FILE,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

    meta = {
        'name':    dated_name,
        'parents': [DRIVE_FOLDER_ID]  # ← saves inside your folder
    }

    f = service.files().create(body=meta, media_body=media).execute()
    print(f"Saved: {dated_name} → Drive folder")
    print(f"File ID: {f['id']}")

if __name__ == '__main__':
    upload()