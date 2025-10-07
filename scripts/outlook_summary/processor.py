from __future__ import annotations

import logging
import os

from googleapiclient.errors import HttpError

from .auth import acquire_graph_token
from .config import (
    load_env_file,
    project_names,
)
from .constants import DEFAULT_MARKDOWN_FOLDER, DEFAULT_TRIGGER_CATEGORY
from .drive import (
    build_drive_service,
    create_drive_markdown,
    download_drive_file_text,
    ensure_drive_folder,
    find_drive_file,
    update_drive_markdown,
)
from .graph import (
    debug_log_recent_categories,
    fetch_categorized_messages,
    mark_message_processed,
)
from .renderer import (
    append_section,
    current_run_timestamp,
    html_to_text,
    render_initial_markdown,
    render_update_section,
    sanitize_filename,
    strip_subject_prefixes,
    wikilink_today,
)
from .summary import get_openai_client, summarize_email


def process_messages() -> None:
    log_level_name = os.getenv("LOG_LEVEL", "INFO").upper()
    log_level = getattr(logging, log_level_name, logging.INFO)
    logging.basicConfig(level=log_level, format="%(levelname)s: %(message)s")

    load_env_file(os.getenv("OUTLOOK_SECRETS_FILE"))

    token = acquire_graph_token()
    client = get_openai_client()
    drive_service = build_drive_service()

    target_folder_name = os.getenv("GOOGLE_DRIVE_FOLDER_NAME", DEFAULT_MARKDOWN_FOLDER)
    folder_id = ensure_drive_folder(
        drive_service,
        target_folder_name,
        os.getenv("GOOGLE_DRIVE_FOLDER_ID"),
    )
    project_terms = project_names()

    trigger_category = os.getenv("OUTLOOK_TRIGGER_CATEGORY", DEFAULT_TRIGGER_CATEGORY)
    fetch_limit = int(os.getenv("OUTLOOK_FETCH_LIMIT", "10"))

    messages = fetch_categorized_messages(token, fetch_limit, trigger_category)
    if not messages:
        logging.debug(
            "No messages matched category '%s'. Checking recent messages for diagnostics...",
            trigger_category,
        )
        debug_log_recent_categories(token, fetch_limit)
        return

    for message in messages:
        message_id = message.get("id")
        subject = strip_subject_prefixes(message.get("subject") or "No subject")
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
                    message,
                    summary_payload,
                    date_link,
                    subject,
                    updated_at,
                    project_terms,
                )
                combined = append_section(existing_content, new_section)
                upload_info = update_drive_markdown(
                    drive_service, existing_file["id"], combined
                )
            else:
                logging.info("Creating new summary for subject '%s'", subject)
                markdown = render_initial_markdown(
                    message,
                    summary_payload,
                    date_link,
                    subject,
                    project_terms,
                )
                upload_info = create_drive_markdown(
                    drive_service, folder_id, filename, markdown
                )

            logging.info(
                "Stored Drive file %s (%s)", filename, upload_info.get("webViewLink")
            )

            mark_message_processed(
                token,
                message_id,
                message.get("categories", []),
                trigger_category=trigger_category,
            )
        except HttpError as err:
            logging.error("Google Drive error for message %s: %s", message_id, err)
        except Exception as err:  # noqa: BLE001
            logging.exception("Failed to process message %s: %s", message_id, err)


__all__ = ["process_messages"]
