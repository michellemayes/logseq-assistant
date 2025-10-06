# Scheduling the Outlook Summarizer on macOS

Automate the script with a cron job so your Outlook summaries stay up-to-date without manual runs.

## 1. Create a wrapper script

Cron runs with a minimal environment, so use a short shell script to activate your virtualenv, set the working directory, and execute the Python script.

Create `~/bin/run_outlook_summary.sh` with the following contents (adjust paths if your repo or virtualenv live elsewhere):

```bash
#!/bin/zsh
source /Users/michelle/Repos/logseq-assistant/.venv/bin/activate
cd /Users/michelle/Repos/logseq-assistant
python scripts/outlook_ai_summary.py >> ~/Library/Logs/outlook_ai_summary.log 2>&1
```

Make it executable:

```bash
chmod +x ~/bin/run_outlook_summary.sh
```

> 📝 Ensure the script references the correct virtualenv, repository path, and desired log location. The log file lets you review output or errors after cron runs.

## 2. Install the cron job

Edit your cron table:

```bash
crontab -e
```

Add an entry to run the script on your desired cadence. The example below executes every 15 minutes:

```cron
*/15 * * * * /Users/michelle/bin/run_outlook_summary.sh
```

Always use absolute paths inside cron entries, and keep the environment setup within the wrapper script.

Save and exit (`:wq` in vim). Confirm the job is registered:

```bash
crontab -l
```

## 3. Monitor and adjust

- Tail the log to confirm execution:

  ```bash
  tail -f ~/Library/Logs/outlook_ai_summary.log
  ```

- Adjust the schedule by editing the cron entry (`crontab -e`).
- If you need environment variables beyond what `secrets.env` loads (e.g., proxy settings), export them inside `run_outlook_summary.sh` before the `python` command.

## Alternatives

Prefer launchd on macOS? Create a LaunchAgent plist instead of using cron, pointing it to the same wrapper script. The wrapper approach still applies: ensure the environment is ready before invoking Python.
