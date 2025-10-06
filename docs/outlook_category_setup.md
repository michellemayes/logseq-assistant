# Configure Outlook Categories and Automation

Categorizing emails is the trigger that instructs the summarization script to process a message. Follow these steps to create (or confirm) the trigger category and optionally automate its assignment.

## 1. Create the trigger category

1. Open Outlook on the web ([https://outlook.office.com](https://outlook.office.com)) or the desktop client.
2. Locate **Categories**:
   - **Web**: Select a message → click the **Categories** icon in the command bar.
   - **Desktop**: Right-click a message → choose **Categorize** → **All Categories**.
3. Click **Manage categories** / **New**.
4. Enter the category name specified in your configuration (default is `AI Summarize`).
5. Choose a color (optional) and save.

## 2. (Optional) Create an Outlook rule to auto-tag emails

Automate category assignment so the script runs without manual intervention.

### Outlook on the web

1. Go to **Settings** (gear icon) → **View all Outlook settings** → **Mail** → **Rules**.
2. Click **Add new rule**.
3. Name the rule (e.g., `Route to AI Summarizer`).
4. Add conditions (sender, subject, keywords, etc.).
5. Add an action → **Categorize** → select the trigger category (`AI Summarize`).
6. Add additional actions as needed (e.g., stop processing more rules).
7. Save the rule.

### Outlook desktop (Windows/Mac)

1. In the ribbon, select **Rules** → **Manage Rules & Alerts**.
2. Click **New Rule** and choose the appropriate template (e.g., “Apply rule on messages I receive”).
3. Configure the conditions that should trigger summarization.
4. Under actions, select **assign it to the category** and pick `AI Summarize`.
5. Finish and enable the rule.

## 3. Verify the processed category (optional)

The script removes the trigger category after processing and adds a completion category (default `AI Summarized`). You can create this category in Outlook to make the status visible.

1. Repeat the steps in section 1, naming the category `AI Summarized`.
2. Monitor your inbox to confirm processed emails are re-labeled.

## 4. Confirm mailbox ID for the script

The script requires the user principal name (UPN) or object ID of the mailbox as `MS_GRAPH_USER_ID`. To confirm:

- Use the mailbox email address (e.g., `user@yourdomain.com`), or
- Retrieve the object ID from the **Azure AD** user profile if you prefer a GUID.

## 5. Test the end-to-end flow

1. Ensure you have finished the Azure AD app registration and permissions setup.
2. Assign the trigger category (`AI Summarize`) to a new inbox message.
3. Run the script:

   ```bash
   python scripts/outlook_ai_summary.py
   ```

4. Verify the email is re-categorized to `AI Summarized` (or your configured name) and that a Markdown file is created/updated in Google Drive.

## Tips

- Refine your Outlook rule conditions over time to capture the right messages.
- Consider combining the trigger category with mailbox rules that flag or move messages for follow-up.
- If multiple people share the workflow, document the category names so everyone uses the same triggers.
