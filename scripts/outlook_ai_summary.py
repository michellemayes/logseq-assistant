import datetime
import io
import json
import logging
import os
import re
from typing import Any, Dict, List, Optional

import msal
import requests
from bs4 import BeautifulSoup
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaInMemoryUpload, MediaIoBaseDownload
from openai import OpenAI

GRAPH_APP_SCOPE = ["https://graph.microsoft.com/.default"]
DEFAULT_DELEGATED_SCOPES = ["Mail.ReadWrite"]
DEFAULT_TRIGGER_CATEGORY = "AI Summarize"
DEFAULT_PROCESSED_CATEGORY = "AI Summarized"
DEFAULT_MARKDOWN_FOLDER = "AI Email Summaries"
DEFAULT_SECRETS_FILE = "secrets.env"
DEFAULT_TOKEN_CACHE_FILE = ".ms_token_cache.json"
OPENAI_DEFAULT_MODEL = "gpt-4o-mini"
SUBJECT_PREFIX_PATTERN = re.compile(r"^\s*(re|fw|fwd|aw|wg):\s*", re.IGNORECASE)
SUBJECT_BRACKET_PREFIX = re.compile(r"^\s*\[[^\]]*\]\s*")
INTERNAL_EMAIL_DOMAINS_ENV = "INTERNAL_EMAIL_DOMAINS"


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
    # if "offline_access" not in scopes:
    #     scopes.append("offline_access")
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
            error = result.get("error")
            description = result.get("error_description", "")
            if error == "invalid_client" and "AADSTS7000218" in description:
                raise RuntimeError(
                    "Failed to acquire delegated token: AADSTS7000218. Enable "
                    "'Allow public client flows' on the Azure AD app registration "
                    "or set MS_CLIENT_SECRET and MS_AUTH_MODE=client_credentials to use "
                    "application permissions."
                )
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
        "id,subject,from,receivedDateTime,sentDateTime,categories,body,replyTo,toRecipients"
    )
    params = {
        "$top": os.getenv("OUTLOOK_FETCH_LIMIT", "10"),
        "$select": select_fields,
        "$filter": f"categories/any(c:c eq '{trigger_category}')",
        "$orderby": "receivedDateTime desc",
    }

    url = f"https://graph.microsoft.com/v1.0/users/{user_id}/messages"
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


def summarize_email(client: OpenAI, subject: str, body_text: str) -> Dict[str, Any]:
    model = os.getenv("OPENAI_MODEL", OPENAI_DEFAULT_MODEL)
    system_prompt = (
        "You are an assistant that summarizes Outlook emails for busy knowledge workers. "
        "Respond with a compact JSON object containing five fields: "
        "'summary' (2-3 sentence overview), 'key_points' (concise bullet strings), 'todos' "
        "(actionable follow-ups without any TODO prefix), 'context_notes' (assumptions or "
        "background, may be empty), and 'topics' (major themes or entities useful for linking). "
        "Do not include markdown or text outside the JSON."
    )
    user_prompt = (
        "Summarize the following email for Logseq. Highlight the sender's intent, critical facts, explicit or implied "
        "requests, and recommended follow-ups. Return JSON only.\n\n"
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
        response_format={"type": "json_object"},
    )
    content = response.choices[0].message.content.strip()

    try:
        payload = json.loads(content)
    except json.JSONDecodeError:
        logging.warning("OpenAI summary was not valid JSON; returning fallback text")
        payload = {
            "summary": content,
            "key_points": [],
            "todos": [],
            "context_notes": [],
            "topics": [],
        }

    return normalize_summary_payload(payload)


def normalize_summary_payload(payload: Dict[str, Any]) -> Dict[str, Any]:
    def normalize_list(key: str) -> List[str]:
        value = payload.get(key, [])
        if isinstance(value, str):
            value = [value]
        if not isinstance(value, list):
            return []
        cleaned = []
        for item in value:
            if isinstance(item, str):
                text = item.strip()
                if text:
                    cleaned.append(text)
        return cleaned

    summary_text = payload.get("summary")
    if isinstance(summary_text, list):
        summary_text = " ".join(str(item).strip() for item in summary_text if item)
    if not isinstance(summary_text, str):
        summary_text = ""
    summary_text = summary_text.strip()

    normalized = {
        "summary": summary_text,
        "key_points": normalize_list("key_points"),
        "todos": normalize_list("todos") or normalize_list("follow_ups"),
        "context_notes": normalize_list("context_notes"),
        "topics": normalize_list("topics"),
    }

    if (
        not normalized["summary"]
        and normalized["key_points"]
    ):
        normalized["summary"] = "; ".join(normalized["key_points"][:2])

    if not normalized["summary"]:
        normalized["summary"] = "(No summary returned)"

    # Deduplicate topics while preserving order
    seen = set()
    unique_topics = []
    for topic in normalized["topics"]:
        key = topic.lower()
        if key in seen or not topic.strip():
            continue
        seen.add(key)
        unique_topics.append(topic.strip())
    normalized["topics"] = unique_topics

    return normalized


def strip_subject_prefixes(subject: str) -> str:
    if not subject:
        return "No subject"
    cleaned = subject
    while True:
        new_value = SUBJECT_PREFIX_PATTERN.sub("", cleaned)
        if new_value == cleaned:
            break
        cleaned = new_value
    cleaned = SUBJECT_BRACKET_PREFIX.sub("", cleaned).strip()
    return cleaned or "No subject"


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


def current_run_timestamp() -> str:
    return datetime.datetime.now().isoformat(timespec="seconds")


def sanitize_filename(name: str) -> str:
    cleaned = re.sub(r"[\\/:*?\"<>|]", "-", name)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned[:180] if len(cleaned) > 180 else cleaned


def internal_domains() -> List[str]:
    raw = os.getenv(INTERNAL_EMAIL_DOMAINS_ENV, "")
    domains: List[str] = []
    for part in raw.split(","):
        value = part.strip().lower()
        if value:
            domains.append(value)
    return domains


def is_internal_email(address: Optional[str]) -> bool:
    if not address or "@" not in address:
        return False
    domain = address.split("@", 1)[1].lower()
    for allowed in internal_domains():
        if domain == allowed or domain.endswith(f".{allowed}"):
            return True
    return False


def format_person_link(name: Optional[str], email: Optional[str]) -> str:
    name = (name or "").strip()
    email = (email or "").strip()

    if is_internal_email(email):
        display = name or email.split("@", 1)[0]
        parts = [part for part in re.split(r"[\s,]+", display) if part]
        if len(parts) >= 2:
            first = parts[0].capitalize()
            last_initial = parts[-1][0].upper()
            label = f"{first} {last_initial}"
        else:
            local_parts = email.split("@", 1)[0].split(".")
            first = local_parts[0].capitalize() if local_parts else display
            last_initial = local_parts[-1][0].upper() if len(local_parts) > 1 else ""
            label = f"{first} {last_initial}".strip()
        label = label or email.split("@", 1)[0]
        return f"[[{label}]]"

    if name and email and name.lower() != email.lower():
        return f"{name} ({email})"
    if name:
        return name
    if email:
        return email
    return "Unknown"


def render_initial_markdown(
    message: Dict, summary: Dict[str, Any], date_link: str, subject: str
) -> str:
    sender = message.get("from", {}).get("emailAddress", {})
    sender_name = sender.get("name")
    sender_address = sender.get("address")
    received = message.get("receivedDateTime") or message.get("sentDateTime")
    recipients = format_recipients(message.get("toRecipients", []))

    sender_display = format_person_link(sender_name, sender_address)

    lines = [
        "tags:: email",
        "",
        f"- {date_link}",
        f"\t- Subject: {subject}",
        f"\t- From: {sender_display}",
    ]
    if recipients:
        lines.append(f"\t- To: {recipients}")
    if received:
        lines.append(f"\t- Received: {received}")

    lines.extend(format_summary_sections(summary))

    return "\n".join(lines).strip() + "\n"


def render_update_section(
    message: Dict,
    summary: Dict[str, Any],
    date_link: str,
    subject: str,
    updated_at: str,
) -> str:
    sender = message.get("from", {}).get("emailAddress", {})
    sender_name = sender.get("name")
    sender_address = sender.get("address")
    received = message.get("receivedDateTime") or message.get("sentDateTime")
    recipients = format_recipients(message.get("toRecipients", []))

    sender_display = format_person_link(sender_name, sender_address)

    lines = [
        f"- {date_link}",
        f"\t- Update for: {subject}",
        f"\t- From: {sender_display}",
    ]
    if recipients:
        lines.append(f"\t- To: {recipients}")
    if received:
        lines.append(f"\t- Received: {received}")
    if updated_at:
        lines.append(f"\t- Updated: {updated_at}")

    lines.extend(format_summary_sections(summary))

    return "\n".join(lines).strip() + "\n"


def format_summary_sections(summary: Dict[str, Any]) -> List[str]:
    lines: List[str] = []

    topics = summary.get("topics", [])

    def link_text(value: str) -> str:
        return link_topics(value, topics)

    summary_text = summary.get("summary", "").strip()
    if summary_text:
        lines.append(f"\t- **Summary:** {link_text(summary_text)}")

    key_points = summary.get("key_points", [])
    if key_points:
        lines.append("\t- **Key Points:**")
        for point in key_points:
            lines.append(f"\t\t- {link_text(point)}")

    context_notes = summary.get("context_notes", [])
    if context_notes:
        lines.append("\t- **Context:**")
        for note in context_notes:
            lines.append(f"\t\t- {link_text(note)}")

    todos = summary.get("todos", [])
    if todos:
        lines.append("\t- **Tasks:**")
        for todo in todos:
            todo_text = todo.strip()
            if todo_text.lower().startswith("todo "):
                todo_text = todo_text[5:].strip()
            lines.append(f"\t\t- TODO {link_text(todo_text)}")

    return lines


def link_topics(text: str, topics: List[str]) -> str:
    if not text or not topics:
        return text

    linked_text = text
    for raw_topic in sorted(topics, key=len, reverse=True):
        topic = raw_topic.strip()
        if not topic:
            continue
        if re.search(r"\w", topic):
            pattern = re.compile(
                rf"(?<!\[\[)\b({re.escape(topic)})\b(?!\]\])",
                re.IGNORECASE,
            )
        else:
            pattern = re.compile(
                rf"(?<!\[\[)({re.escape(topic)})(?!\]\])",
                re.IGNORECASE,
            )

        def replacer(match: re.Match) -> str:
            matched = match.group(0)
            return f"[[{matched}]]"

        linked_text = pattern.sub(replacer, linked_text)

    return linked_text


def format_recipients(recipients: List[Dict]) -> str:
    display_parts: List[str] = []
    for recipient in recipients or []:
        email_address = recipient.get("emailAddress", {})
        name = (email_address.get("name") or "").strip()
        address = (email_address.get("address") or "").strip()
        display_parts.append(format_person_link(name, address))
    return ", ".join(display_parts)


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
        raw_subject = message.get("subject") or "No subject"
        subject = strip_subject_prefixes(raw_subject)
        logging.info("Processing message %s", message_id)
        try:
            html_body = message.get("body", {}).get("content", "")
            plain_text = html_to_text(html_body)
            summary_payload = summarize_email(client, subject, plain_text)
            date_link = wikilink_today()
            updated_at = current_run_timestamp()

            safe_subject = sanitize_filename(subject) or "untitled"
            filename = f"{safe_subject}.md"

            existing_file = find_drive_file(drive_service, folder_id, filename)
            if existing_file:
                logging.info("Updating existing summary %s", existing_file.get("webViewLink"))
                existing_content = download_drive_file_text(
                    drive_service, existing_file["id"]
                )
                new_section = render_update_section(
                    message, summary_payload, date_link, subject, updated_at
                )
                combined = append_section(existing_content, new_section)
                upload_info = update_drive_markdown(
                    drive_service, existing_file["id"], combined
                )
            else:
                logging.info("Creating new summary for subject '%s'", subject)
                markdown = render_initial_markdown(
                    message, summary_payload, date_link, subject
                )
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
            content = getattr(err, "content", b"")
            message = "Google Drive error for message %s: %s"
            if (
                err.resp.status == 403
                and isinstance(content, (bytes, bytearray))
                and b"storageQuotaExceeded" in content
            ):
                logging.error(
                    "Google Drive storage quota exceeded while processing message %s. "
                    "Service accounts do not include personal Drive storage. Upload into "
                    "a shared drive, share a user-owned folder and set GOOGLE_DRIVE_FOLDER_ID, "
                    "or supply GOOGLE_DELEGATED_USER with domain-wide delegation so uploads "
                    "count against a user quota.",
                    message_id,
                )
            else:
                logging.error(message, message_id, err)
        except Exception as err:  # noqa: BLE001
            logging.error("Failed to process message %s: %s", message_id, err)


if __name__ == "__main__":
    process_messages()
