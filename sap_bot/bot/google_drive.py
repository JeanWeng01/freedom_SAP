"""Google Drive integration — download PDFs from a shared folder."""

import os
import io
import json
import logging
import tempfile

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

log = logging.getLogger(__name__)

DRIVE_FOLDER_ID = os.environ.get("GOOGLE_DRIVE_FOLDER_ID", "1nkRDml96iRsgFC7K0N3RcC-VmHoDHyk7")
SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]


def _get_credentials():
    """Load Google service account credentials."""
    creds_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if creds_json:
        info = json.loads(creds_json)
        return Credentials.from_service_account_info(info, scopes=SCOPES)

    creds_file = os.environ.get("GOOGLE_SERVICE_ACCOUNT_FILE", "service_account.json")
    if os.path.isfile(creds_file):
        return Credentials.from_service_account_file(creds_file, scopes=SCOPES)

    raise RuntimeError("No Google credentials found")


def _get_drive_service():
    creds = _get_credentials()
    return build("drive", "v3", credentials=creds)


def find_file(filename: str) -> dict | None:
    """Search for a file by name in the configured Drive folder.

    Args:
        filename: e.g. "VIM64D.pdf"

    Returns file metadata dict with 'id' and 'name', or None if not found.
    """
    service = _get_drive_service()
    query = f"name = '{filename}' and '{DRIVE_FOLDER_ID}' in parents and trashed = false"

    result = service.files().list(
        q=query,
        fields="files(id, name, mimeType)",
        pageSize=5,
    ).execute()

    files = result.get("files", [])
    if not files:
        log.warning("File '%s' not found in Drive folder %s", filename, DRIVE_FOLDER_ID)
        return None

    if len(files) > 1:
        log.warning("Multiple files named '%s' found — using first", filename)

    log.info("Found file: %s (id: %s)", files[0]["name"], files[0]["id"])
    return files[0]


def download_file(filename: str) -> str | None:
    """Download a file from Google Drive to a temp directory.

    Args:
        filename: e.g. "VIM64D.pdf"

    Returns the local file path, or None if not found.
    The caller is responsible for deleting the temp file after use.
    """
    file_meta = find_file(filename)
    if not file_meta:
        return None

    service = _get_drive_service()
    request = service.files().get_media(fileId=file_meta["id"])

    # Download to a temp file
    temp_dir = tempfile.mkdtemp(prefix="sap_pod_")
    local_path = os.path.join(temp_dir, filename)

    with open(local_path, "wb") as f:
        downloader = MediaIoBaseDownload(f, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()

    log.info("Downloaded %s to %s", filename, local_path)
    return local_path


def cleanup_temp_file(path: str):
    """Delete a temp file and its parent directory."""
    if not path:
        return
    try:
        os.remove(path)
        parent = os.path.dirname(path)
        if parent and os.path.isdir(parent) and parent.startswith(tempfile.gettempdir()):
            os.rmdir(parent)
        log.debug("Cleaned up temp file: %s", path)
    except Exception as e:
        log.warning("Could not clean up %s: %s", path, e)
