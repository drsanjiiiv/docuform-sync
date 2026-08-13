# DocuForm Sync Support 🛠️

Thank you for using DocuForm Sync! We want to ensure your form synchronization workflows run seamlessly. Because our application runs entirely inside your Google Workspace environment, we track all feedback and system bugs publicly on GitHub.

---

## 🔐 Managing & Deleting Your Data

DocuForm Sync does not store your data on external servers. Your source data lives in your own Google Sheets, and your form options live in your own Google Form.

**To remove DocuForm Sync data:**

1. **Saved mapping configurations:** Use the **↺ Start Over** or the reset option in the DocuForm Sync sidebar, or remove the add-on entirely to clear the per-user configuration.
2. **Scheduled runs / triggers:** Disable Auto-Sync in the **⚙️ Auto-Sync** tab, or review and remove triggers at any time under **Extensions > Apps Script > Triggers** in your form.
3. **Remove the add-on entirely:** Right-click the add-on icon in Google Forms and choose "Remove", then revoke access under your Google Account permissions (myaccount.google.com > Security > Third-party access).

Once the add-on is removed, no trace of DocuForm Sync processing remains in your files.

---

## 🐛 How to Report an Issue or Bug

If a sync isn't running correctly, or if you encounter a system error, please follow these steps to log an issue:

1. Click on the **Issues** tab at the top of this GitHub repository page.
2. Click the green **New Issue** button.
3. Provide a clear title (e.g., *"Dropdown options not updating after sync"*).
4. Describe the problem, including any error messages you see in your sidebar.
5. Click **Submit new issue**.

---

## 🔒 Data & Privacy Note
When reporting bugs, please **never** paste screenshots or logs that contain sensitive or personal business data (such as email addresses or private spreadsheet data).
