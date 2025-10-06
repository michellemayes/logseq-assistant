# Configure OpenAI or Azure OpenAI Access

The summarization script uses the OpenAI Chat Completions API to produce Markdown summaries. Follow these steps to obtain the `OPENAI_API_KEY` and optional settings (`OPENAI_MODEL`, `OPENAI_BASE_URL`).

## Option A: OpenAI API (api.openai.com)

1. Sign in to [platform.openai.com](https://platform.openai.com/).
2. Navigate to **View API keys** from the profile menu.
3. Click **Create new secret key**.
4. Copy the generated key and store it securely—this becomes `OPENAI_API_KEY`.
5. (Optional) Set `OPENAI_MODEL` if you prefer a specific model variant. The script defaults to `gpt-4o-mini`.

### Usage limits and billing

- Ensure your OpenAI account has sufficient quota and billing enabled.
- Monitor usage at **Usage** → **View usage** within the OpenAI dashboard.

## Option B: Azure OpenAI Service

1. Ensure Azure OpenAI is enabled in your Azure subscription.
2. Deploy a Chat Completions-capable model (e.g., GPT-4o or GPT-3.5 Turbo) in the Azure portal.
3. From the Azure OpenAI resource:
   - Copy the **Endpoint** (e.g., `https://my-openai-resource.openai.azure.com/`).
   - Create an **API key** under **Keys and Endpoint**.
4. Set the following values in `secrets.env` or your environment:

   ```ini
   OPENAI_API_KEY=<Azure OpenAI key>
   OPENAI_BASE_URL=https://my-openai-resource.openai.azure.com/
   OPENAI_MODEL=<deployment name>
   ```

   In Azure OpenAI, `OPENAI_MODEL` must match the **deployment name** you assigned.

5. (Optional) Configure network rules or private endpoints so the runtime can reach the Azure OpenAI endpoint.

## Security tips

- Treat the API key like a password; never commit it to source control.
- Rotate keys periodically and update `secrets.env` accordingly.
- Limit access by running the script in an environment where the key is stored securely (e.g., encrypted secrets manager, environment variables injected at runtime).

## Verify connectivity

After setting the variables, run the script with a sample email or a smoke test to ensure OpenAI responses are successful:

```bash
python scripts/outlook_ai_summary.py
```

If you encounter issues, enable debug logging (`export LOG_LEVEL=DEBUG`) to capture the API response details.
