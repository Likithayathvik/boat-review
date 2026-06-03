import os
import json
from datetime import datetime
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

XLSX_FILE       = 'Android_HearablesApp_Review.xlsx'
DRIVE_FOLDER_ID = os.environ.get('DRIVE_FOLDER_ID', '')

def upload():
    creds_json = json.loads(os.environ['GOOGLE_CREDENTIALS_JSON'])
    creds = Credentials.from_service_account_info(
        creds_json,
        scopes=['https://www.googleapis.com/auth/drive']
    )
    service = build('drive', 'v3', credentials=creds)

    dated_name = f"Android_HearablesApp_Review_{datetime.now().strftime('%Y-%m-%d')}.xlsx"

    media = MediaFileUpload(
        XLSX_FILE,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

    meta = {
        'name':    dated_name,
        'parents': [DRIVE_FOLDER_ID],
        'driveId': DRIVE_FOLDER_ID,
    }

    # Try shared drive first, fallback to regular
    try:
        f = service.files().create(
            body=meta,
            media_body=media,
            supportsAllDrives=True,
            fields='id'
        ).execute()
        print(f"Saved: {dated_name}")
        print(f"File ID: {f['id']}")
    except Exception as e:
        print(f"Upload error: {e}")
        raise

if __name__ == '__main__':
    upload()
