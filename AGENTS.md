# Repository Guidelines

## Project Structure & Module Organization
- `scripts/outlook_ai_summary.py` – thin entrypoint; invokes the orchestration package.
- `scripts/outlook_summary/` – core modules (auth, config, Graph access, Drive upload, OpenAI summarisation, Logseq renderer, pipeline).
- `docs/` – operational playbooks (Azure AD setup, cron scheduling, cloud deployment, etc.).
- `secrets.example.env` – template for local/environment variables; copy to `secrets.env` when running locally.

## Build, Test, and Development Commands
- `python -m venv .venv && source .venv/bin/activate` – create and activate the project virtualenv.
- `pip install -r requirements.txt` – install runtime dependencies (Google API clients, MSAL, OpenAI SDK, BeautifulSoup).
- `python scripts/outlook_ai_summary.py` – run the summariser once using the categories and credentials in `secrets.env`.
- `LOG_LEVEL=DEBUG python scripts/outlook_ai_summary.py` – verbose run that prints Graph/OpenAI diagnostics when triaging issues.

## Coding Style & Naming Conventions
- Follow PEP 8: 4-space indentation, snake_case for functions/variables, PascalCase for classes.
- Keep modules focused; new behaviour should live in a dedicated file under `scripts/outlook_summary/` rather than the entrypoint.
- Prefer descriptive filenames (e.g., `drive.py`, `renderer.py`); avoid abbreviations unless industry-standard.

## Testing Guidelines
- No automated tests exist yet; when adding them, place unit tests under `tests/` mirroring the package structure.
- Use `pytest` for new suites; aim to mock network calls to Microsoft Graph, Google Drive, and OpenAI.
- Name tests after the behaviour under scrutiny (`test_renderer_formats_tasks`), and ensure they run with `pytest -q`.

## Commit & Pull Request Guidelines
- Write imperative, concise commit messages (e.g., “Add configurable project names for summary linking”).
- Each PR should describe scope, list notable configuration changes, and attach screenshots/log excerpts when altering output or behaviour.
- Reference related issues or tickets in the PR body; keep history linear (rebase or squash before merge).

## Security & Configuration Tips
- Never commit real secrets; use `secrets.env` locally and Key Vault/secret managers in cloud deployments.
- Rotate service-account keys and Azure client secrets regularly; document changes in the deployment runbooks.
- When debugging, avoid logging raw email content or tokens; scrub sensitive data before sharing traces.
