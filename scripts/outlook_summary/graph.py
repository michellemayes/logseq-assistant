from __future__ import annotations

import logging
from typing import Dict, List

import requests

from .auth import acquire_graph_token
from .config import get_required_env
from .constants import DEFAULT_TRIGGER_CATEGORY, DEFAULT_PROCESSED_CATEGORY


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


def fetch_categorized_messages(
    token: str, fetch_limit: int, trigger_category: str
) -> List[Dict]:
    user_id = get_required_env("MS_GRAPH_USER_ID")
    select_fields = (
        "id,subject,from,receivedDateTime,sentDateTime,categories,body,replyTo,toRecipients"
    )
    params = {
        "$top": str(fetch_limit),
        "$select": select_fields,
        "$filter": f"categories/any(c:c eq '{trigger_category}')",
        "$orderby": "receivedDateTime desc",
    }

    url = f"https://graph.microsoft.com/v1.0/users/{user_id}/messages"
    response = graph_request(token, "get", url, params=params)
    payload = response.json()
    messages = payload.get("value", [])
    if messages:
        logging.info(
            "Found %s messages with category '%s' via Graph filter",
            len(messages),
            trigger_category,
        )
        return messages

    logging.debug(
        "Server-side filter returned no matches for '%s'; scanning recent messages locally",
        trigger_category,
    )

    params = {
        "$top": str(fetch_limit),
        "$select": select_fields,
        "$orderby": "receivedDateTime desc",
    }
    response = graph_request(token, "get", url, params=params)
    payload = response.json()
    candidates = payload.get("value", [])

    matched: List[Dict] = []
    for message in candidates:
        categories = message.get("categories") or []
        if any(cat.strip() == trigger_category for cat in categories):
            matched.append(message)

    logging.info(
        "Found %s messages with category '%s' after scanning %s recent messages",
        len(matched),
        trigger_category,
        len(candidates),
    )

    if logging.getLogger().isEnabledFor(logging.DEBUG):
        for message in candidates:
            logging.debug(
                "Candidate: subject='%s', categories=%s",
                message.get("subject"),
                message.get("categories"),
            )

    return matched


def debug_log_recent_categories(token: str, fetch_limit: int = 10) -> None:
    if not logging.getLogger().isEnabledFor(logging.DEBUG):
        return

    user_id = get_required_env("MS_GRAPH_USER_ID")
    select_fields = "id,subject,categories,receivedDateTime,parentFolderId"
    params = {
        "$top": str(fetch_limit),
        "$select": select_fields,
        "$orderby": "receivedDateTime desc",
    }
    url = f"https://graph.microsoft.com/v1.0/users/{user_id}/messages"
    try:
        response = graph_request(token, "get", url, params=params)
    except Exception as exc:  # noqa: BLE001
        logging.debug("Failed to fetch recent messages for debugging: %s", exc)
        return

    payload = response.json()
    for entry in payload.get("value", []):
        logging.debug(
            "Recent message candidate: subject='%s', categories=%s, received=%s",
            entry.get("subject"),
            entry.get("categories"),
            entry.get("receivedDateTime"),
        )


def mark_message_processed(
    token: str,
    message_id: str,
    categories: List[str],
    trigger_category: str = DEFAULT_TRIGGER_CATEGORY,
    processed_category: str = DEFAULT_PROCESSED_CATEGORY,
) -> None:
    updated_categories = [c for c in categories if c != trigger_category]
    if processed_category not in updated_categories:
        updated_categories.append(processed_category)

    user_id = get_required_env("MS_GRAPH_USER_ID")
    url = f"https://graph.microsoft.com/v1.0/users/{user_id}/messages/{message_id}"
    payload = {"categories": updated_categories, "isRead": True}
    graph_request(token, "patch", url, json=payload)


__all__ = [
    "acquire_graph_token",
    "fetch_categorized_messages",
    "graph_request",
    "mark_message_processed",
]
