# LogSeq Outlook Assistant

This utility polls an Outlook mailbox for messages assigned to a specific category, generates an AI summary, and uploads the result as a Markdown file to Google Drive. Every note is formatted for Logseq knowledge graphs: wiki-style date links, `tags:: email`, nested bullets for key points/context/tasks, configurable project wikilinks, and internal teammates rendered as `[[First L]]`. Subsequent emails with the same subject reuse the original Markdown file so the thread stays consolidated. The workflow keeps inbox triage, follow-ups, and knowledge capture in sync without manual copy/paste, and it supports both application-level (client credentials) and user-scoped (delegated/device-code) Microsoft Graph access.

## Prerequisites

1. **Python environment** – Python 3.9+ is recommended.
2. **Microsoft Graph application** – [Register an Azure AD app](docs/azure_ad_app_registration.md), choose the appropriate permission model (application or delegated), and collect:
   - `MS_CLIENT_ID`
   - `MS_TENANT_ID`
   - `MS_CLIENT_SECRET` (app permissions only)
   - `MS_GRAPH_USER_ID` (object ID or user principal name of the mailbox owner)
3. **OpenAI (or Azure OpenAI) access** – [Configure the API key](docs/openai_access.md) and optionally `OPENAI_BASE_URL`/`OPENAI_MODEL` for Azure OpenAI deployments.
4. **Google Drive service account** – [Create a service account and key](docs/google_drive_service_account.md), grant it access to the upload folder, and obtain the JSON key path.
5. **Email category** – [Set up the Outlook trigger and processed categories](docs/outlook_category_setup.md) so emails you want processed receive the appropriate labels.

## Installation

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### Configure secrets

The script loads environment variables from `secrets.env` (ignored by git) before execution. Copy the template file and populate it with your values:

```bash
cp secrets.example.env secrets.env
# edit secrets.env with your credentials
```

- For application permissions, include `MS_CLIENT_SECRET` (the script defaults to `MS_AUTH_MODE=client_credentials`).
- For delegated permissions, omit the client secret and set `MS_AUTH_MODE=device_code`. The first run will prompt you to complete the device sign-in; subsequent runs reuse the token cache.
- In Azure AD, enable **Allow public client flows** on the app registration when using delegated/device-code authentication.

You can change the secrets file path by setting `OUTLOOK_SECRETS_FILE` before running the script.

## Required environment variables

Provide the following either in `secrets.env` or through your shell environment:

| Variable | Description |
| --- | --- |
| `MS_CLIENT_ID` | Azure AD application (client) ID |
| `MS_TENANT_ID` | Azure AD tenant ID |
| `MS_GRAPH_USER_ID` | User ID or UPN of the mailbox to monitor |
| `OPENAI_API_KEY` | API key for OpenAI or Azure OpenAI |
| `GOOGLE_SERVICE_ACCOUNT_FILE` | Path to the service-account JSON key |

## Optional configuration

| Variable | Description | Default |
| --- | --- | --- |
| `MS_CLIENT_SECRET` | Required for client credentials (application permissions) | *none* |
| `MS_AUTH_MODE` | `client_credentials` or `device_code` | auto-detected |
| `MS_TOKEN_CACHE_FILE` | Persistent token cache path for delegated auth | `.ms_token_cache.json` |
| `MS_DELEGATED_SCOPES` | Custom Graph scopes for delegated auth | `Mail.ReadWrite offline_access` |
| `OUTLOOK_SECRETS_FILE` | Path to the secrets file to load | `secrets.env` |
| `OUTLOOK_TRIGGER_CATEGORY` | Category that triggers processing | `AI Summarize` |
| `OUTLOOK_PROCESSED_CATEGORY` | Category appended after upload | `AI Summarized` |
| `OUTLOOK_FETCH_LIMIT` | Maximum emails to fetch per run | `10` |
| `GOOGLE_DRIVE_FOLDER_ID` | Explicit Drive folder ID to upload into | auto-created |
| `GOOGLE_DRIVE_FOLDER_NAME` | Folder name when auto-creating | `AI Email Summaries` |
| `GOOGLE_DELEGATED_USER` | Email to impersonate when using domain-wide delegation | *none* |
| `PROJECT_NAMES` | Comma-separated project names to auto-link in summaries | *none* |
| `INTERNAL_EMAIL_DOMAINS` | Comma-separated domains treated as internal for Logseq links | *none* |
| `OPENAI_MODEL` | Model name for summarization | `gpt-4o-mini` |
| `OPENAI_BASE_URL` | Override endpoint (Azure OpenAI, proxies, etc.) | *none* |

> ⚠️ Google service accounts do not include personal Drive storage. Either upload into a shared drive, share a user-owned folder and set `GOOGLE_DRIVE_FOLDER_ID`, or enable domain-wide delegation and set `GOOGLE_DELEGATED_USER` so uploads consume a user quota.

## Running the collector

```bash
python scripts/outlook_ai_summary.py
```

The script will:

- Load secrets from `secrets.env` (or the file referenced by `OUTLOOK_SECRETS_FILE`).
- Acquire a Microsoft Graph token using either client credentials or device code based on your configuration.
- Fetch messages from the mailbox inbox that have the trigger category.
- Create or update a Drive Markdown note named after the email subject, prepending a current-date wiki link (`[[Oct 6th, 2025]]`) to each summary section.
- Convert the email body to plain text, request an AI-generated Logseq-ready outline (tab-indented bullets with `TODO` tasks), and append new context to existing notes instead of creating duplicates.
- Mark the email as read, remove the trigger category, and add the processed category.

> 📝 Device-code mode prints a one-time sign-in URL and code if silent token acquisition fails (e.g., on first run or after cache reset).

## Scheduling

Run the script on a schedule using your preferred scheduler (cron, Windows Task Scheduler, GitHub Actions with secrets, etc.). Ensure the environment variables and credentials are available in that runtime. For delegated auth, schedule runs on the same machine or environment that has access to the token cache.

See [docs/scheduling.md](docs/scheduling.md) for a macOS cron setup that activates the virtualenv and logs output.

For a cloud-hosted alternative triggered by Power Automate, follow [docs/cloud_deployment.md](docs/cloud_deployment.md).

## Testing notes

The script assumes network access to Microsoft Graph, OpenAI, and Google Drive APIs. Run it in an environment with the appropriate firewall permissions and secrets configured.
