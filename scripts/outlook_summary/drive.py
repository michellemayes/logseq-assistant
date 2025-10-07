from __future__ import annotations

import io
import logging
from typing import Dict, Optional

import os

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaInMemoryUpload, MediaIoBaseDownload

from .config import get_required_env


def build_drive_service():
    credentials_path = get_required_env("GOOGLE_SERVICE_ACCOUNT_FILE")
    delegate_user = os.getenv("GOOGLE_DELEGATED_USER")
    scopes = [
        "https://www.googleapis.com/auth/drive.file",
        "https://www.googleapis.com/auth/drive.metadata",
    ]

    credentials = service_account.Credentials.from_service_account_file(
        credentials_path, scopes=scopes
    )
    if delegate_user:
        credentials = credentials.with_subject(delegate_user)

    return build("drive", "v3", credentials=credentials)


def ensure_drive_folder(service, folder_name: str, folder_id_override: Optional[str]) -> str:
    if folder_id_override:
        return folder_id_override

    logging.debug("Looking up Drive folder '%s'", folder_name)
    query = (
        "name = '{}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    ).format(folder_name.replace("'", "\\'"))
    response = (
        service.files()
        .list(
            q=query,
            spaces="drive",
            fields="files(id, name)",
            includeItemsFromAllDrives=True,
            supportsAllDrives=True,
        )
        .execute()
    )
    files = response.get("files", [])
    if files:
        return files[0]["id"]

    logging.info("Creating Drive folder '%s'", folder_name)
    file_metadata = {
        "name": folder_name,
        "mimeType": "application/vnd.google-apps.folder",
    }
    folder = (
        service.files()
        .create(body=file_metadata, fields="id", supportsAllDrives=True)
        .execute()
    )
    return folder["id"]


def find_drive_file(service, folder_id: str, filename: str) -> Optional[Dict]:
    escaped_name = filename.replace("'", "\\'")
    query = (
        "name = '{}' and '{}' in parents and trashed = false"
    ).format(escaped_name, folder_id)
    response = (
        service.files()
        .list(
            q=query,
            spaces="drive",
            fields="files(id, name, webViewLink)",
            includeItemsFromAllDrives=True,
            supportsAllDrives=True,
        )
        .execute()
    )
    files = response.get("files", [])
    return files[0] if files else None


def download_drive_file_text(service, file_id: str) -> str:
    request = service.files().get_media(fileId=file_id, supportsAllDrives=True)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return fh.getvalue().decode("utf-8")


def create_drive_markdown(
    service,
    folder_id: str,
    filename: str,
    content: str,
) -> Dict:
    file_metadata = {
        "name": filename,
        "mimeType": "text/markdown",
        "parents": [folder_id],
    }
    media = MediaInMemoryUpload(content.encode("utf-8"), mimetype="text/markdown")

    logging.debug("Creating %s in Drive folder %s", filename, folder_id)
    return (
        service.files()
        .create(
            body=file_metadata,
            media_body=media,
            fields="id, webViewLink",
            supportsAllDrives=True,
        )
        .execute()
    )


def update_drive_markdown(service, file_id: str, content: str) -> Dict:
    media = MediaInMemoryUpload(content.encode("utf-8"), mimetype="text/markdown")
    logging.debug("Updating Drive file %s", file_id)
    return (
        service.files()
        .update(
            fileId=file_id,
            media_body=media,
            fields="id, webViewLink",
            supportsAllDrives=True,
        )
        .execute()
    )


__all__ = [
    "build_drive_service",
    "ensure_drive_folder",
    "find_drive_file",
    "download_drive_file_text",
    "create_drive_markdown",
    "update_drive_markdown",
]
