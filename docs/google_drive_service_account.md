# Create a Google Drive Service Account for Markdown Uploads

Follow these steps to provision a Google Cloud service account that can upload Markdown files to the desired Google Drive folder. The resulting JSON key file path is used as `GOOGLE_SERVICE_ACCOUNT_FILE`.

## Prerequisites

- Access to a Google Cloud project (owner or editor role).
- Permission to enable the Google Drive API.
- Access to the target Google Drive folder (or authority to share it with the service account).

## 1. Enable the Google Drive API

1. Go to the [Google Cloud Console](https://console.cloud.google.com/).
2. Select the project you will use for this integration.
3. Navigate to **APIs & Services** → **Library**.
4. Search for **Google Drive API** and click **Enable**.

## 2. Create the service account

1. In the Google Cloud Console, open **IAM & Admin** → **Service Accounts**.
2. Click **+ CREATE SERVICE ACCOUNT**.
3. Provide a descriptive **Service account name** (e.g., `outlook-ai-summarizer`).
4. Click **Create and continue**.
5. (Optional) Assign roles if needed for other workloads—no additional roles are required for Drive access when using service account keys.
6. Click **Done**.

## 3. Generate a JSON key

1. On the service account list, click the account you just created.
2. Go to the **Keys** tab.
3. Click **Add key** → **Create new key**.
4. Choose **JSON** and click **Create**.
5. Save the downloaded JSON file securely; this file path becomes `GOOGLE_SERVICE_ACCOUNT_FILE`.

## 4. Grant Drive folder access

Service accounts are not members of your Google Workspace by default. Grant access in one of the following ways:

> ⚠️ Service accounts do not receive personal Drive storage. Uploads must target a shared drive or a folder owned by a user with available quota. You can also enable domain-wide delegation and set `GOOGLE_DELEGATED_USER` to impersonate that user.

### Option A: Shared drive (recommended)

1. Create or open a [shared drive](https://support.google.com/a/users/answer/9310249) (team drive).
2. Add the service account email (`<service-account-name>@<project-id>.iam.gserviceaccount.com`) as a **Content manager** (or higher) member of the shared drive.
3. Place or create the destination folder within the shared drive and capture its ID for `GOOGLE_DRIVE_FOLDER_ID`.

### Option B: Share a user-owned folder and impersonate the owner

1. Share the folder with the service account and keep the human user as the owner.
2. Enable domain-wide delegation (see Option B, steps 1–5 above).
3. Set `GOOGLE_DELEGATED_USER` to the owner’s email so uploads count against that user’s quota.

### Domain-wide delegation setup (if needed)

If you want the service account to impersonate users across the domain:

1. In the service account details, enable **Domain-wide delegation** and note the **Client ID**.
2. In the [Google Admin Console](https://admin.google.com/), go to **Security** → **API controls** → **Domain-wide delegation**.
3. Click **Add new** and enter the client ID.
4. Set the OAuth scopes (space-separated):

   ```text
   https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/drive.metadata
   ```

5. Save the configuration.
6. Set `GOOGLE_DELEGATED_USER` to the email address you want the script to impersonate.

## 5. Update `secrets.env`

Add the key path and optional delegation user to your secrets file:

```ini
GOOGLE_SERVICE_ACCOUNT_FILE=/absolute/path/to/service-account-key.json
# Optional: only if using domain-wide delegation
GOOGLE_DELEGATED_USER=user@yourdomain.com
```

## 6. Test Drive access

Run the script after setting the secrets:

```bash
python scripts/outlook_ai_summary.py
```

If the upload fails, double-check that the service account has edit access to the folder and that the Drive API is enabled.

## Security reminders

- Treat the JSON key like a password—store it securely, and do not commit it to source control.
- Rotate service account keys periodically. Delete old keys in the Console and update `GOOGLE_SERVICE_ACCOUNT_FILE` accordingly.
- Limit the service account’s access to only the folders it needs for this automation.
