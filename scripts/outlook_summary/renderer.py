from __future__ import annotations

import datetime
import re
from typing import Any, Dict, List, Optional

from bs4 import BeautifulSoup

from .config import internal_domains
from .constants import SUBJECT_BRACKET_PREFIX, SUBJECT_PREFIX_PATTERN


def html_to_text(html: str) -> str:
    soup = BeautifulSoup(html or "", "html.parser")
    for tag in soup(["script", "style"]):
        tag.decompose()
    text = soup.get_text(separator=" ")
    return re.sub(r"\s+", " ", text).strip()


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


def sanitize_filename(name: str) -> str:
    cleaned = re.sub(r"[\\/:*?\"<>|]", "-", name)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned[:180] if len(cleaned) > 180 else cleaned


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


def format_recipients(recipients: List[Dict]) -> str:
    display_parts: List[str] = []
    for recipient in recipients or []:
        email_address = recipient.get("emailAddress", {})
        name = (email_address.get("name") or "").strip()
        address = (email_address.get("address") or "").strip()
        display_parts.append(format_person_link(name, address))
    return ", ".join(display_parts)


def link_projects(text: str, project_terms: List[str]) -> str:
    if not text or not project_terms:
        return text

    linked = text
    for term in sorted(project_terms, key=len, reverse=True):
        clean_term = term.strip()
        if not clean_term:
            continue
        escaped = re.escape(clean_term)
        if re.search(r"\w", clean_term):
            pattern = re.compile(
                rf"(?<!\[\[)\b({escaped})\b(?!\]\])",
                re.IGNORECASE,
            )
        else:
            pattern = re.compile(
                rf"(?<!\[\[)({escaped})(?!\]\])",
                re.IGNORECASE,
            )

        def repl(match: re.Match) -> str:
            return f"[[{match.group(0)}]]"

        linked = pattern.sub(repl, linked)

    return linked


def format_summary_sections(summary: Dict[str, Any], project_terms: List[str]) -> List[str]:
    lines: List[str] = []

    summary_text = summary.get("summary", "").strip()
    if summary_text:
        lines.append(f"\t- **Summary:** {link_projects(summary_text, project_terms)}")

    key_points = summary.get("key_points", [])
    if key_points:
        lines.append("\t- **Key Points:**")
        for point in key_points:
            lines.append(f"\t\t- {link_projects(point, project_terms)}")

    context_notes = summary.get("context_notes", [])
    if context_notes:
        lines.append("\t- **Context:**")
        for note in context_notes:
            lines.append(f"\t\t- {link_projects(note, project_terms)}")

    todos = summary.get("todos", [])
    if todos:
        lines.append("\t- **Tasks:**")
        for todo in todos:
            todo_text = todo.strip()
            if todo_text.lower().startswith("todo "):
                todo_text = todo_text[5:].strip()
            lines.append(f"\t\t- TODO {link_projects(todo_text, project_terms)}")

    return lines


def render_initial_markdown(
    message: Dict,
    summary: Dict[str, Any],
    date_link: str,
    subject: str,
    project_terms: List[str],
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

    lines.extend(format_summary_sections(summary, project_terms))

    return "\n".join(lines).strip() + "\n"


def render_update_section(
    message: Dict,
    summary: Dict[str, Any],
    date_link: str,
    subject: str,
    updated_at: str,
    project_terms: List[str],
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

    lines.extend(format_summary_sections(summary, project_terms))

    return "\n".join(lines).strip() + "\n"


def append_section(existing_content: str, new_section: str) -> str:
    existing = existing_content.rstrip()
    section = new_section.strip()
    if not existing:
        return section + "\n"
    return f"{existing}\n\n{section}\n"


__all__ = [
    "append_section",
    "current_run_timestamp",
    "format_person_link",
    "html_to_text",
    "render_initial_markdown",
    "render_update_section",
    "sanitize_filename",
    "strip_subject_prefixes",
    "wikilink_today",
]
