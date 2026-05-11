"""
apps_script_uploader.py
Uploads files to Google Drive through a deployed Google Apps Script web app.
"""

import base64
import requests


def upload_file(*, web_app_url: str, token: str, folder_id: str, file_path: str, file_name: str, mimetype: str):
    with open(file_path, "rb") as f:
        content_b64 = base64.b64encode(f.read()).decode("ascii")

    response = requests.post(
        web_app_url,
        json={
            "token": token,
            "folderId": folder_id,
            "fileName": file_name,
            "mimeType": mimetype,
            "content": content_b64,
        },
        timeout=120,
    )
    response.raise_for_status()
    result = response.json()
    if not result.get("ok"):
        raise ValueError(result.get("error", "Apps Script upload failed"))
    return result
