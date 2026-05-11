"""
drive_uploader.py
Uploads the generated PPTX back to the same Google Drive folder using OAuth.
"""

SCOPES = ["https://www.googleapis.com/auth/drive.file"]
AUTH_URI = "https://accounts.google.com/o/oauth2/auth"
TOKEN_URI = "https://oauth2.googleapis.com/token"


def build_client_config(client_id: str, client_secret: str) -> dict:
    return {
        "web": {
            "client_id": client_id,
            "client_secret": client_secret,
            "auth_uri": AUTH_URI,
            "token_uri": TOKEN_URI,
        }
    }


def make_auth_url(client_id: str, client_secret: str, redirect_uri: str) -> str:
    from google_auth_oauthlib.flow import Flow

    flow = Flow.from_client_config(
        build_client_config(client_id, client_secret),
        scopes=SCOPES,
        redirect_uri=redirect_uri,
    )
    auth_url, _ = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent",
    )
    return auth_url


def exchange_code_for_credentials(
    *,
    client_id: str,
    client_secret: str,
    redirect_uri: str,
    code: str,
) -> dict:
    from google_auth_oauthlib.flow import Flow

    flow = Flow.from_client_config(
        build_client_config(client_id, client_secret),
        scopes=SCOPES,
        redirect_uri=redirect_uri,
    )
    flow.fetch_token(code=code)
    creds = flow.credentials
    return {
        "token": creds.token,
        "refresh_token": creds.refresh_token,
        "token_uri": creds.token_uri,
        "client_id": creds.client_id,
        "client_secret": creds.client_secret,
        "scopes": creds.scopes,
    }


def credentials_from_info(info: dict):
    from google.oauth2.credentials import Credentials

    return Credentials(
        token=info.get("token"),
        refresh_token=info.get("refresh_token"),
        token_uri=info.get("token_uri", TOKEN_URI),
        client_id=info.get("client_id"),
        client_secret=info.get("client_secret"),
        scopes=info.get("scopes", SCOPES),
    )


def upload_file_to_folder(
    *,
    credentials_info: dict,
    folder_id: str,
    file_path: str,
    file_name: str,
    mimetype: str,
) -> dict:
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaFileUpload

    creds = credentials_from_info(credentials_info)
    service = build("drive", "v3", credentials=creds)

    metadata = {
        "name": file_name,
        "parents": [folder_id],
    }
    media = MediaFileUpload(
        file_path,
        mimetype=mimetype,
        resumable=True,
    )

    uploaded = (
        service.files()
        .create(
            body=metadata,
            media_body=media,
            fields="id,name,webViewLink",
            supportsAllDrives=True,
        )
        .execute()
    )
    return uploaded


def upload_pptx_to_folder(
    *,
    credentials_info: dict,
    folder_id: str,
    file_path: str,
    file_name: str,
) -> dict:
    return upload_file_to_folder(
        credentials_info=credentials_info,
        folder_id=folder_id,
        file_path=file_path,
        file_name=file_name,
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )
