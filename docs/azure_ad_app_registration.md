# Register an Azure AD Application for Outlook AI Summaries

Follow these steps to register an Azure Active Directory (Azure AD) application that can read and update Outlook mail via Microsoft Graph. The same registration supports either application permissions (daemon-style) or delegated permissions (user-scoped) depending on how you configure the script.

## Prerequisites

- Azure AD tenant admin (or permissions to register apps and grant consent).
- Access to the [Azure portal](https://portal.azure.com/).
- Exchange Online mailbox that the automation will summarize.

## 1. Create the application registration

1. Sign in to the Azure portal and open **Azure Active Directory**.
2. Select **App registrations** → **New registration**.
3. Enter a friendly **Name** (e.g., `Outlook AI Summaries`).
4. Choose **Accounts in this organizational directory only** (single-tenant) unless you need multi-tenant access.
5. Leave **Redirect URI** empty (not required for device code or client credential flows).
6. Click **Register**.
7. On the overview page, record:
   - **Application (client) ID** → `MS_CLIENT_ID`
   - **Directory (tenant) ID** → `MS_TENANT_ID`

## 2. Decide on the permission model

| Scenario | Add this Graph permission | Script auth mode |
| --- | --- | --- |
| Background service or automation covering multiple mailboxes | **Application** → `Mail.ReadWrite` | `MS_AUTH_MODE=client_credentials` (default when `MS_CLIENT_SECRET` is set) |
| User-scoped access (your case) tied to a single signed-in mailbox | **Delegated** → `Mail.ReadWrite` | `MS_AUTH_MODE=device_code` (or another delegated flow) |

> ✅ You can add both permission types, but only the one matching `MS_AUTH_MODE` is used.

### Add the Microsoft Graph permission

1. In the app registration, open **API permissions** → **Add a permission** → **Microsoft Graph**.
2. Choose **Delegated permissions** or **Application permissions** based on the table above.
3. Search for `Mail.ReadWrite`, check the box, and click **Add permissions**.

## 3. Create credentials (only for application permissions)

If you plan to use `MS_AUTH_MODE=client_credentials`:

1. Go to **Certificates & secrets**.
2. Under **Client secrets**, click **New client secret**.
3. Provide a description, choose an expiry, and click **Add**.
4. Copy the **Value** immediately—store it as `MS_CLIENT_SECRET`.

> 🔐 Client secrets are not required for delegated/device-code flows.

## 4. Grant admin consent

1. Still on **API permissions**, click **Grant admin consent for &lt;Tenant&gt;**.
2. Confirm the prompt. For delegated permissions, this pre-approves the scope so users are not prompted individually.

## 5. (Optional) Restrict mailbox access (application permissions only)

Application permissions allow tenant-wide mailbox access unless restricted. To limit scope:

1. Use the Exchange Online PowerShell module.
2. Create a mail-enabled security group containing the allowed mailbox owners.
3. Run:

   ```powershell
   Connect-ExchangeOnline
   New-ApplicationAccessPolicy -AppId <MS_CLIENT_ID> -PolicyScopeGroupId ai-summarizer-mailboxes@yourdomain.com -AccessRight RestrictAccess -Description "Allow Outlook summarizer app"
   Test-ApplicationAccessPolicy -AppId <MS_CLIENT_ID> -Identity user@yourdomain.com
   ```

## 6. Configure the summarization script

### Environment variables (common)

Add to `secrets.env` or your secret store:

```ini
MS_CLIENT_ID=<Application client ID>
MS_TENANT_ID=<Directory tenant ID>
MS_GRAPH_USER_ID=<Mailbox UPN or ID>
```

### Option A: Application permissions (client credentials)

```ini
MS_CLIENT_SECRET=<Client secret value>
# MS_AUTH_MODE can be omitted or set explicitly
MS_AUTH_MODE=client_credentials
```

Run the script—token acquisition happens silently using the client secret.

### Option B: Delegated permissions (device code flow)

```ini
MS_AUTH_MODE=device_code
# Optional overrides
# MS_TOKEN_CACHE_FILE=.ms_token_cache.json
# MS_DELEGATED_SCOPES=Mail.ReadWrite offline_access
```

1. Execute the script interactively once:

   ```bash
   python scripts/outlook_ai_summary.py
   ```

2. When prompted, follow the device code instructions (visit the URL, enter the code, and sign in with the mailbox account).
3. The token cache (defaults to `.ms_token_cache.json`) stores refresh tokens so subsequent runs—scheduled or automated—can renew access silently.

> ℹ️ Delegated permissions allow the app to operate only on the signed-in user’s mailbox. No Exchange application access policy is required.

## Maintenance tips

- Rotate client secrets before they expire (application permissions).
- For delegated flows, rerun the device code flow if you clear the token cache or revoke refresh tokens.
- Revisit **API permissions** and click **Grant admin consent** after adding new scopes.
- Remove the app registration and any Exchange policies if you retire the automation.
