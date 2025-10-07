# Deploying the Outlook Summarizer as an Azure Function (Power Automate Trigger)

This guide explains how to run the existing `scripts/outlook_ai_summary.py` inside Azure Functions and invoke it from Power Automate. The Python code remains unchanged; we wrap it in a lightweight HTTP-triggered function that Power Automate calls when Outlook categories change.

## Architecture overview

1. **Power Automate flow** detects an Outlook message tagged with the trigger category and calls an Azure Function endpoint.
2. **Azure Function** (Python) loads the same environment variables used locally, runs `scripts/outlook_ai_summary.py`, and returns status output.
3. **Azure Key Vault** stores secrets (OpenAI key, Google service-account JSON, Graph credentials if using client credentials). The Function app uses managed identity to retrieve them.
4. **Google Drive** and **Microsoft Graph** behave exactly as in the local solution.

## Prerequisites

- Azure subscription with permission to create Function Apps, Storage Accounts, and Key Vault access policies.
- Python 3.11 or 3.10 locally (matching the Azure Functions runtime you choose).
- Azure CLI (`az`) and Azure Functions Core Tools installed.
- The repository cloned locally, including `scripts/outlook_ai_summary.py` and `requirements.txt`.
- Service-account JSON file for Google Drive and Azure AD app registration as documented earlier.

## 1. Provision Azure resources

1. Set context variables (replace names/locations as needed):

   ```bash
   RESOURCE_GROUP=rg-logseq-summarizer
   LOCATION=eastus
   STORAGE_ACCOUNT=logsummarizer$RANDOM
   FUNCTION_APP=logseq-summarizer-func
   KEY_VAULT=kv-logseq-summarizer
   az group create --name "$RESOURCE_GROUP" --location "$LOCATION"
   az storage account create --name "$STORAGE_ACCOUNT" --location "$LOCATION" --resource-group "$RESOURCE_GROUP" --sku Standard_LRS
   az keyvault create --name "$KEY_VAULT" --resource-group "$RESOURCE_GROUP" --location "$LOCATION"
   az functionapp create --name "$FUNCTION_APP" --resource-group "$RESOURCE_GROUP" --storage-account "$STORAGE_ACCOUNT" --consumption-plan-location "$LOCATION" --runtime python --functions-version 4
   ```

2. Enable managed identity on the Function app and allow it to read secrets:

   ```bash
   az functionapp identity assign --name "$FUNCTION_APP" --resource-group "$RESOURCE_GROUP"
   PRINCIPAL_ID=$(az functionapp identity show --name "$FUNCTION_APP" --resource-group "$RESOURCE_GROUP" --query principalId -o tsv)
   az keyvault set-policy --name "$KEY_VAULT" --object-id "$PRINCIPAL_ID" --secret-permissions get list
   ```

## 2. Store secrets in Key Vault

Upload each secret you previously used in `secrets.env`. Examples:

```bash
az keyvault secret set --vault-name "$KEY_VAULT" --name MS-CLIENT-ID --value <GUID>
az keyvault secret set --vault-name "$KEY_VAULT" --name MS-TENANT-ID --value <GUID>
 az keyvault secret set --vault-name "$KEY_VAULT" --name OPENAI-API-KEY --value <KEY>
 az keyvault secret set --vault-name "$KEY_VAULT" --name GOOGLE-SERVICE-ACCOUNT --file /path/to/service-account.json
 az keyvault secret set --vault-name "$KEY_VAULT" --name INTERNAL-EMAIL-DOMAINS --value "pushnami.com"
```

Repeat for any optional variables (trigger category names, OpenAI model overrides, etc.). If you are using the device-code flow, note that the token cache file is not viable in a stateless function; prefer the client-credentials setup in the cloud.

## 3. Prepare the Function project locally

1. From the repo root, create a Functions project with an HTTP trigger (this scaffolding lives outside the existing Git history if you prefer to keep the repo clean):

   ```bash
   func init azure-function --python
   cd azure-function
   func new --name runSummary --template "HTTP trigger" --authlevel function
   ```

2. Replace `runSummary/__init__.py` with a wrapper that invokes the existing script:

   ```python
   import json
   import logging
   import os
   import subprocess
   from pathlib import Path

   import azure.functions as func

   REPO_ROOT = Path(__file__).resolve().parents[2] / "logseq-assistant"
   SCRIPT_PATH = REPO_ROOT / "scripts" / "outlook_ai_summary.py"

   def main(req: func.HttpRequest) -> func.HttpResponse:
       logging.info("Triggering Outlook AI summarizer")
       env = os.environ.copy()
       result = subprocess.run(
           ["python", str(SCRIPT_PATH)],
           cwd=str(REPO_ROOT),
           env=env,
           capture_output=True,
           text=True,
       )
       payload = {
           "returncode": result.returncode,
           "stdout": result.stdout,
           "stderr": result.stderr,
       }
       status = 200 if result.returncode == 0 else 500
       return func.HttpResponse(json.dumps(payload), status_code=status, mimetype="application/json")
   ```

   > This wrapper leaves `scripts/outlook_ai_summary.py` unchanged. It simply shells out to the existing script with the current environment.

3. Update `requirements.txt` in the function project to include dependencies from the repo:

   ```bash
   # inside azure-function/requirements.txt
   azure-functions
   -r ../requirements.txt
   ```

4. Copy the repository into the function folder for deployment (or reference a git submodule). The HTTP function needs access to `scripts/outlook_ai_summary.py` and its packages.

## 4. Configure Function App settings

Map application settings to Key Vault secrets so the script sees the same environment variables it expects locally:

```bash
az functionapp config appsettings set --name "$FUNCTION_APP" --resource-group "$RESOURCE_GROUP" --settings \
    MS_CLIENT_ID="@Microsoft.KeyVault(SecretUri=https://$KEY_VAULT.vault.azure.net/secrets/MS-CLIENT-ID/)" \
    MS_TENANT_ID="@Microsoft.KeyVault(SecretUri=https://$KEY_VAULT.vault.azure.net/secrets/MS-TENANT-ID/)" \
    OPENAI_API_KEY="@Microsoft.KeyVault(SecretUri=https://$KEY_VAULT.vault.azure.net/secrets/OPENAI-API-KEY/)" \
    GOOGLE_SERVICE_ACCOUNT_FILE="/tmp/service-account.json" \
    INTERNAL_EMAIL_DOMAINS="@Microsoft.KeyVault(SecretUri=https://$KEY_VAULT.vault.azure.net/secrets/INTERNAL-EMAIL-DOMAINS/)" \
    MS_AUTH_MODE=client_credentials
```

For the Google JSON key, download it at startup with an init script or store the JSON string in Key Vault and write it to disk at runtime. A simple approach is to add this near the top of `runSummary/__init__.py`:

```python
from azure.identity import ManagedIdentityCredential
from azure.keyvault.secrets import SecretClient

credential = ManagedIdentityCredential()
client = SecretClient(vault_url=f"https://{os.environ['KEY_VAULT_NAME']}.vault.azure.net", credential=credential)
secret = client.get_secret("GOOGLE-SERVICE-ACCOUNT")
with open("/tmp/service-account.json", "w") as handle:
    handle.write(secret.value)

os.environ.setdefault("GOOGLE_SERVICE_ACCOUNT_FILE", "/tmp/service-account.json")
```

(Adjust the code if you prefer to set `KEY_VAULT_NAME` as an app setting.)

## 5. Deploy the Function

From the `azure-function` directory:

```bash
func azure functionapp publish $FUNCTION_APP
```

After deployment, test the endpoint (URL shown in the publish output) with a tool like curl or Postman. Ensure it returns `returncode: 0` when at least one categorized email exists.

## 6. Build the Power Automate flow

1. In Power Automate, create a new automated cloud flow.
2. Trigger: **When a new email arrives (V3)** or **When a new email arrives in a shared mailbox**, scoped to your mailbox and filtered for the trigger category (`Has Attachments` → `No`, `Importance` → `Any`, etc.); use advanced options to filter by category or use a condition later.
3. Add a condition to check `Contains(Category, 'AI Summarize')` (or your custom trigger name).
4. In the `Yes` branch, add an **HTTP** action:
   - Method: `POST`
   - URI: `https://<function-app>.azurewebsites.net/api/runSummary?code=<function-key>`
   - Headers: `Content-Type: application/json`
   - Body: optionally pass metadata (`{ 
       "messageId": "@{triggerOutputs()?['body/messageid']}"
     }`) if you want traceability.
5. Optionally, log the HTTP response to track success/failure. Use the response body to send notifications or retry if the function returns a non-200 status.

## 7. Testing and monitoring

- Use the Function App monitoring blade in Azure portal to check recent invocations and logs.
- Inspect the `~/Library/Logs/outlook_ai_summary.log`-style output now returned in the HTTP response. Consider forwarding logs to Application Insights for centralized monitoring.
- Simulate an Outlook category change to confirm the flow triggers as expected.

## 8. Operational tips

- Keep the Function runtime in sync with the Python version you test locally.
- Run `pip freeze > requirements.txt` inside the function project to capture exact dependency versions when you upgrade libraries.
- For deterministic behavior, pin the OpenAI model (e.g., `OPENAI_MODEL=gpt-4o-mini`) in app settings.
- If you later need to process multiple messages per invocation, consider extending the HTTP payload so Power Automate can pass the specific message ID to process.
- Enable retries or create a second Power Automate flow that reacts to non-200 responses to avoid missing summaries.

## 9. Cleanup

To delete resources when finished:

```bash
az group delete --name "$RESOURCE_GROUP" --yes --no-wait
```

This removes the Function App, Storage Account, and Key Vault in one operation.

---

With this deployment in place, Outlook category changes fire from Power Automate, the Azure Function reuses your existing summarizer, and Google Drive/Logseq stay up to date without running local cron jobs.
