from __future__ import annotations

import json
import logging
import os
from typing import Any, Dict

from openai import OpenAI

from .constants import OPENAI_DEFAULT_MODEL
from .config import get_required_env


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
        "Respond with a compact JSON object containing four fields: "
        "'summary' (2-3 sentence overview), 'key_points' (concise bullet strings), 'todos' "
        "(actionable follow-ups without any TODO prefix), and 'context_notes' (assumptions or "
        "background, may be empty). "
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
        }

    return normalize_summary_payload(payload)


def normalize_summary_payload(payload: Dict[str, Any]) -> Dict[str, Any]:
    def normalize_list(key: str) -> list[str]:
        value = payload.get(key, [])
        if isinstance(value, str):
            value = [value]
        if not isinstance(value, list):
            return []
        cleaned: list[str] = []
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
    }

    if not normalized["summary"] and normalized["key_points"]:
        normalized["summary"] = "; ".join(normalized["key_points"][:2])

    if not normalized["summary"]:
        normalized["summary"] = "(No summary returned)"

    return normalized


__all__ = ["get_openai_client", "summarize_email"]
