import datetime
import io
import logging
import os
import re
from typing import Dict, List, Optional

import msal
import requests
from bs4 import BeautifulSoup
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaInMemoryUpload, MediaIoBaseDownload
from openai import OpenAI

GRAPH_APP_SCOPE = ["https://graph.microsoft.com/.default"]
DEFAULT_DELEGATED_SCOPES = ["Mail.ReadWrite", "offline_access"]
DEFAULT_TRIGGER_CATEGORY = "AI Summarize"
DEFAULT_PROCESSED_CATEGORY = "AI Summarized"
DEFAULT_MARKDOWN_FOLDER = "AI Email Summaries"
DEFAULT_SECRETS_FILE = "secrets.env"
DEFAULT_TOKEN_CACHE_FILE = ".ms_token_cache.json"
OPENAI_DEFAULT_MODEL = "gpt-4o-mini"


def load_env_file(file_path: str) -> None:
    path = (file_path or "").strip()
    if not path:
        return

    if not os.path.exists(path):
        logging.debug("Secrets file %s not found; skipping load", path)
        return

    logging.info("Loading secrets from %s", path)
    with open(path, "r", encoding="utf-8") as handle:
        for raw_line in handle:
            line = raw_line.strip()
            if not line or line.startswith("#"):
                continue
            if "=" not in line:
                logging.debug("Skipping malformed secrets line: %s", line)
                continue

            key, value = line.split("=", 1)
            key = key.strip()
            if key.startswith("export "):
                key = key[len("export ") :].strip()

            value = value.strip().strip("'\"")
            os.environ.setdefault(key, value)


def get_required_env(var_name: str) -> str:
    value = os.getenv(var_name)
    if not value:
        raise EnvironmentError(f"Missing required environment variable: {var_name}")
    return value


def get_auth_mode() -> str:
    requested = (os.getenv("MS_AUTH_MODE") or "").strip().lower()
    if requested:
        return requested
    return "client_credentials" if os.getenv("MS_CLIENT_SECRET") else "device_code"


def build_token_cache(cache_path: str) -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    if cache_path and os.path.exists(cache_path):
        try:
            with open(cache_path, "r", encoding="utf-8") as handle:
                cache.deserialize(handle.read())
        except Exception as err:  # noqa: BLE001
            logging.warning("Failed to load token cache %s: %s", cache_path, err)
    return cache


def persist_token_cache(cache: msal.SerializableTokenCache, cache_path: str) -> None:
    if not cache_path or not cache.has_state_changed:
        return
    directory = os.path.dirname(cache_path)
    if directory:
        os.makedirs(directory, exist_ok=True)
    with open(cache_path, "w", encoding="utf-8") as handle:
        handle.write(cache.serialize())


def delegated_scopes() -> List[str]:
    custom = os.getenv("MS_DELEGATED_SCOPES")
    if custom:
        scopes = [
            part.strip()
            for part in re.split(r"[\s,]+", custom)
            if part.strip()
        ]
    else:
        scopes = list(DEFAULT_DELEGATED_SCOPES)
    if "offline_access" not in scopes:
        scopes.append("offline_access")
    return scopes


def acquire_graph_token() -> str:
    client_id = get_required_env("MS_CLIENT_ID")
    tenant_id = get_required_env("MS_TENANT_ID")
    authority = f"https://login.microsoftonline.com/{tenant_id}"

    auth_mode = get_auth_mode()
    logging.debug("Using Microsoft Graph auth mode: %s", auth_mode)

    if auth_mode == "client_credentials":
        client_secret = get_required_env("MS_CLIENT_SECRET")
        app = msal.ConfidentialClientApplication(
            client_id=client_id,
            client_credential=client_secret,
            authority=authority,
        )
        token_response = app.acquire_token_for_client(scopes=GRAPH_APP_SCOPE)
        if "access_token" not in token_response:
            raise RuntimeError(f"Failed to acquire Graph token: {token_response}")
        return token_response["access_token"]

    if auth_mode == "device_code":
        cache_path = os.getenv("MS_TOKEN_CACHE_FILE", DEFAULT_TOKEN_CACHE_FILE)
        cache = build_token_cache(cache_path)
        app = msal.PublicClientApplication(
            client_id=client_id,
            authority=authority,
            token_cache=cache,
        )
        scopes = delegated_scopes()
        accounts = app.get_accounts()
        result = None
        if accounts:
            logging.debug("Attempting silent token acquisition for account %s", accounts[0].get("username"))
            result = app.acquire_token_silent(scopes=scopes, account=accounts[0])

        if not result:
            logging.info("Initiating device code flow for delegated Graph access")
            flow = app.initiate_device_flow(scopes=scopes)
            if "user_code" not in flow:
                raise RuntimeError(f"Failed to initiate device flow: {flow}")
            logging.warning(flow["message"])
            result = app.acquire_token_by_device_flow(flow)

        if "access_token" not in result:
            raise RuntimeError(f"Failed to acquire delegated token: {result}")

        persist_token_cache(cache, cache_path)
        return result["access_token"]

    raise ValueError(
        "Unsupported MS_AUTH_MODE. Use 'client_credentials', 'device_code', or leave unset."
    )


def graph_request(token: str, method: str, url: str, **kwargs) -> requests.Response:
    headers = kwargs.pop("headers", {})
    headers["Authorization"] = f"Bearer {token}"
    headers.setdefault("Accept", "application/json")
    headers.setdefault("Content-Type", "application/json")

    logging.debug("Graph %s %s", method.upper(), url)
    response = requests.request(method=method, url=url, headers=headers, **kwargs)
    if not response.ok:
        raise RuntimeError(
            f"Graph API call failed ({response.status_code}): {response.text}"
        )
    return response


def fetch_categorized_messages(token: str) -> List[Dict]:
    user_id = get_required_env("MS_GRAPH_USER_ID")
    trigger_category = os.getenv("OUTLOOK_TRIGGER_CATEGORY", DEFAULT_TRIGGER_CATEGORY)
    select_fields = (
        "id,subject,from,receivedDateTime,categories,body,replyTo"
    )
    params = {
        "$top": os.getenv("OUTLOOK_FETCH_LIMIT", "10"),
        "$select": select_fields,
        "$filter": f"categories/any(c:c eq '{trigger_category}')",
        "$orderby": "receivedDateTime desc",
    }

    url = f"https://graph.microsoft.com/v1.0/users/{user_id}/mailFolders/Inbox/messages"
    response = graph_request(token, "get", url, params=params)
    payload = response.json()
    messages = payload.get("value", [])
    logging.info("Found %s messages with category '%s'", len(messages), trigger_category)
    return messages


def html_to_text(html: str) -> str:
    soup = BeautifulSoup(html or "", "html.parser")
    for tag in soup(["script", "style"]):
        tag.decompose()
    text = soup.get_text(separator=" ")
    return re.sub(r"\s+", " ", text).strip()


def get_openai_client() -> OpenAI:
    api_key = get_required_env("OPENAI_API_KEY")
    base_url = os.getenv("OPENAI_BASE_URL")
    if base_url:
        return OpenAI(api_key=api_key, base_url=base_url)
    return OpenAI(api_key=api_key)


def summarize_email(client: OpenAI, subject: str, body_text: str) -> str:
    model = os.getenv("OPENAI_MODEL", OPENAI_DEFAULT_MODEL)
    system_prompt = (
        "You are an assistant that summarizes Outlook emails for busy knowledge workers. "
        "Provide a concise markdown-formatted summary with bullet key points and explicit next steps."
    )
    user_prompt = (
        "Summarize the following email. Highlight the sender's intent, key facts, explicit or implied "
        "requests, and recommended follow-ups. Mark follow ups in a new section with TODO in front. If next steps are unclear, mention assumptions.\n\n"
        f"Subject: {subject or 'No subject'}\n\n"
        f"Body: {body_text}"
    )

    logging.debug("Requesting summary from OpenAI model %s", model)
    response = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        temperature=0.2,
        max_tokens=400,
    )
    return response.choices[0].message.content.strip()


def ordinal(day: int) -> str:
    if 10 <= day % 100 <= 20:
        suffix = "th"
    else:
        suffix = {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")
    return f"{day}{suffix}"


def wikilink_today() -> str:
    now = datetime.datetime.now()
    month = now.strftime("%b")
    return f"[[{month} {ordinal(now.day)}, {now.year}]]"


def sanitize_filename(name: str) -> str:
    cleaned = re.sub(r"[\\/:*?\"<>|]", "-", name)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned[:180] if len(cleaned) > 180 else cleaned


def render_initial_markdown(message: Dict, summary: str, date_link: str) -> str:
    sender = message.get("from", {}).get("emailAddress", {})
    sender_name = sender.get("name", "Unknown sender")
    sender_address = sender.get("address", "")
    received = message.get("receivedDateTime")
    subject = message.get("subject", "No subject")

    lines = [
        date_link,
        "",
        f"# {subject}",
        "",
        f"- **From:** {sender_name} ({sender_address})".strip(),
    ]
    if received:
        lines.append(f"- **Received:** {received}")

    lines.extend([
        "",
        "## AI Summary",
        summary,
    ])

    return "\n".join(lines).strip() + "\n"


def render_update_section(message: Dict, summary: str, date_link: str) -> str:
    sender = message.get("from", {}).get("emailAddress", {})
    sender_name = sender.get("name", "Unknown sender")
    sender_address = sender.get("address", "")
    received = message.get("receivedDateTime")

    lines = [
        date_link,
        "",
        "## AI Summary Update",
        "",
        f"- **From:** {sender_name} ({sender_address})".strip(),
    ]
    if received:
        lines.append(f"- **Received:** {received}")

    lines.extend([
        "",
        summary,
    ])

    return "\n".join(lines).strip() + "\n"


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


def ensure_drive_folder(service, folder_name: str) -> str:
    folder_id = os.getenv("GOOGLE_DRIVE_FOLDER_ID")
    if folder_id:
        return folder_id

    logging.debug("Looking up Drive folder '%s'", folder_name)
    query = (
        "name = '{}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    ).format(folder_name.replace("'", "\\'"))
    response = (
        service.files()
        .list(q=query, spaces="drive", fields="files(id, name)")
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
    folder = service.files().create(body=file_metadata, fields="id").execute()
    return folder["id"]


def find_drive_file(service, folder_id: str, filename: str) -> Optional[Dict]:
    escaped_name = filename.replace("'", "\\'")
    query = (
        "name = '{}' and '{}' in parents and trashed = false"
    ).format(escaped_name, folder_id)
    response = (
        service.files()
        .list(q=query, spaces="drive", fields="files(id, name, webViewLink)")
        .execute()
    )
    files = response.get("files", [])
    return files[0] if files else None


def download_drive_file_text(service, file_id: str) -> str:
    request = service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return fh.getvalue().decode("utf-8")


def create_drive_markdown(service, folder_id: str, filename: str, content: str) -> Dict:
    file_metadata = {
        "name": filename,
        "mimeType": "text/markdown",
        "parents": [folder_id],
    }
    media = MediaInMemoryUpload(content.encode("utf-8"), mimetype="text/markdown")

    logging.debug("Creating %s in Drive folder %s", filename, folder_id)
    return (
        service.files()
        .create(body=file_metadata, media_body=media, fields="id, webViewLink")
        .execute()
    )


def update_drive_markdown(service, file_id: str, content: str) -> Dict:
    media = MediaInMemoryUpload(content.encode("utf-8"), mimetype="text/markdown")
    logging.debug("Updating Drive file %s", file_id)
    return (
        service.files()
        .update(fileId=file_id, media_body=media, fields="id, webViewLink")
        .execute()
    )


def append_section(existing_content: str, new_section: str) -> str:
    existing = existing_content.rstrip()
    section = new_section.strip()
    if not existing:
        return section + "\n"
    return f"{existing}\n\n---\n\n{section}\n"


def mark_message_processed(
    token: str,
    message_id: str,
    categories: List[str],
) -> None:
    trigger_category = os.getenv("OUTLOOK_TRIGGER_CATEGORY", DEFAULT_TRIGGER_CATEGORY)
    processed_category = os.getenv(
        "OUTLOOK_PROCESSED_CATEGORY", DEFAULT_PROCESSED_CATEGORY
    )

    updated_categories = [c for c in categories if c != trigger_category]
    if processed_category not in updated_categories:
        updated_categories.append(processed_category)

    user_id = get_required_env("MS_GRAPH_USER_ID")
    url = f"https://graph.microsoft.com/v1.0/users/{user_id}/messages/{message_id}"
    payload = {"categories": updated_categories, "isRead": True}
    graph_request(token, "patch", url, json=payload)


def process_messages():
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    secrets_file = os.getenv("OUTLOOK_SECRETS_FILE", DEFAULT_SECRETS_FILE)
    load_env_file(secrets_file)

    token = acquire_graph_token()
    client = get_openai_client()
    drive_service = build_drive_service()

    target_folder_name = os.getenv("GOOGLE_DRIVE_FOLDER_NAME", DEFAULT_MARKDOWN_FOLDER)
    folder_id = ensure_drive_folder(drive_service, target_folder_name)

    for message in fetch_categorized_messages(token):
        message_id = message.get("id")
        subject = message.get("subject") or "No subject"
        logging.info("Processing message %s", message_id)
        try:
            html_body = message.get("body", {}).get("content", "")
            plain_text = html_to_text(html_body)
            summary = summarize_email(client, subject, plain_text)
            date_link = wikilink_today()

            safe_subject = sanitize_filename(subject) or "untitled"
            filename = f"{safe_subject}.md"

            existing_file = find_drive_file(drive_service, folder_id, filename)
            if existing_file:
                logging.info("Updating existing summary %s", existing_file.get("webViewLink"))
                existing_content = download_drive_file_text(
                    drive_service, existing_file["id"]
                )
                new_section = render_update_section(message, summary, date_link)
                combined = append_section(existing_content, new_section)
                upload_info = update_drive_markdown(
                    drive_service, existing_file["id"], combined
                )
            else:
                logging.info("Creating new summary for subject '%s'", subject)
                markdown = render_initial_markdown(message, summary, date_link)
                upload_info = create_drive_markdown(
                    drive_service, folder_id, filename, markdown
                )

            logging.info(
                "Stored Drive file %s (%s)",
                filename,
                upload_info.get("webViewLink"),
            )

            mark_message_processed(
                token,
                message_id,
                message.get("categories", []),
            )
        except HttpError as err:
            logging.error("Google Drive error for message %s: %s", message_id, err)
        except Exception as err:  # noqa: BLE001
            logging.error("Failed to process message %s: %s", message_id, err)


if __name__ == "__main__":
    process_messages()
