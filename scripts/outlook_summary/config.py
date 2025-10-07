from __future__ import annotations

import logging
import os
import re
from pathlib import Path
from typing import List

import msal

from .constants import (
    DEFAULT_DELEGATED_SCOPES,
    DEFAULT_SECRETS_FILE,
    DEFAULT_TOKEN_CACHE_FILE,
    INTERNAL_EMAIL_DOMAINS_ENV,
    MS_AUTH_MODE_ENV,
    MS_CLIENT_SECRET_ENV,
    MS_DELEGATED_SCOPES_ENV,
    MS_TOKEN_CACHE_FILE_ENV,
    PROJECT_NAMES_ENV,
)


def load_env_file(file_path: str | None = None) -> None:
    path = Path((file_path or DEFAULT_SECRETS_FILE).strip())
    if not path:
        return
    if not path.exists():
        logging.debug("Secrets file %s not found; skipping load", path)
        return

    logging.info("Loading secrets from %s", path)
    with path.open("r", encoding="utf-8") as handle:
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
    requested = (os.getenv(MS_AUTH_MODE_ENV) or "").strip().lower()
    if requested:
        return requested
    return "client_credentials" if os.getenv(MS_CLIENT_SECRET_ENV) else "device_code"


def build_token_cache(cache_path: str | None = None) -> msal.SerializableTokenCache:
    path = cache_path or os.getenv(MS_TOKEN_CACHE_FILE_ENV, DEFAULT_TOKEN_CACHE_FILE)
    cache = msal.SerializableTokenCache()
    if path and os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as handle:
                cache.deserialize(handle.read())
        except Exception as err:  # noqa: BLE001
            logging.warning("Failed to load token cache %s: %s", path, err)
    return cache


def persist_token_cache(cache: msal.SerializableTokenCache, cache_path: str | None = None) -> None:
    if not cache_path:
        cache_path = os.getenv(MS_TOKEN_CACHE_FILE_ENV, DEFAULT_TOKEN_CACHE_FILE)
    if not cache_path or not cache.has_state_changed:
        return
    directory = os.path.dirname(cache_path)
    if directory:
        os.makedirs(directory, exist_ok=True)
    with open(cache_path, "w", encoding="utf-8") as handle:
        handle.write(cache.serialize())


def delegated_scopes() -> List[str]:
    custom = os.getenv(MS_DELEGATED_SCOPES_ENV)
    if custom:
        scopes = [
            part.strip()
            for part in re.split(r"[\s,]+", custom)
            if part.strip()
        ]
    else:
        scopes = list(DEFAULT_DELEGATED_SCOPES)
    return scopes


def internal_domains() -> List[str]:
    raw = os.getenv(INTERNAL_EMAIL_DOMAINS_ENV, "")
    domains: List[str] = []
    for part in raw.split(","):
        value = part.strip().lower()
        if value:
            domains.append(value)
    return domains


def project_names() -> List[str]:
    raw = os.getenv(PROJECT_NAMES_ENV, "")
    names: List[str] = []
    for part in re.split(r"[\n,]", raw):
        value = part.strip()
        if value:
            names.append(value)
    return names
