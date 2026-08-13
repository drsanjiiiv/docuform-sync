# DocuForm Sync ⚡

A modern, highly intuitive Google Forms add-on that keeps your **Dropdown**, **Multiple Choice**, **Checkbox**, and **Grid** options in perfect sync with any Google Sheet column — no more manual copy-paste whenever your data changes.

The add-on runs as a clean, card-based sidebar inside your Google Form editor, letting you map sheet columns to questions in seconds.

---

## ✨ Key Differentiators & Features

Unlike traditional sync tools, DocuForm Sync gives you complete control straight from the sidebar:

* **4-Step Sync Wizard:** Connect your Form & Sheet → Map columns to questions → Preview the options → Sync instantly or save the configuration.
* **Supported Question Types:**
  * 🎯 **Dropdown** (`ListItem`)
  * 🔘 **Multiple Choice**
  * ☑️ **Checkboxes**
  * 🧩 **Grid** (columns)
* **Smart Column Mapping:** Optional **Capacity column** (exclude rows where a quantity is ≤ 0), **Sorting** (A→Z / Z→A, with an optional separate sort column), **+ blank option**, and **Start/End row ranges**.
* **Auto-Sync Engine:** Keep options refreshed on a schedule (every 5 minutes up to daily) or **on every form submit** — fully multi-form safe, with a single shared time-driven trigger.
* **Sync History:** Every run (manual, scheduled, or submit-triggered) is logged with processed/error counts right in the sidebar.

---

## 🛠️ Project Structure
* `Code.gs` - Menu handlers, sidebar launcher, and shared config helpers.
* `SyncEngine.js` - Core engine that reads sheet columns and writes form choices.
* `Router.js` - RPC layer exposed to the sidebar via `google.script.run`.
* `BackgroundService.js` - Auto-sync execution and on-form-submit trigger handling.
* `TriggerService.js` - Time-driven trigger creation and management.
* `UI_Main.html` - The sidebar dashboard (Wizard / Auto-Sync / History tabs).
* `appsscript.json` - Apps Script manifest (runtime + OAuth scopes).
* `Privacy-Policy.md` - Our data protection compliance standards.
* `Terms-of-Service.md` - Terms and usage guidelines.

---

## 📂 Spreadsheet Picker Setup

The **📂 Browse** button uses the [Google Picker API](https://developers.google.com/drive/picker) with the `drive.file` scope, so users grant access to exactly the sheet they select. Configure your Google Cloud project once:

1. Create/link a **standard GCP project** (script.google.com → Project Settings → Change project) — note the **project number**.
2. In the Cloud Console for that project:
   - Configure the **OAuth consent screen** (External, app name + support email).
   - Enable the **Google Picker API** and the **Google Drive API**.
   - Create a **Browser API key** (no referrer restriction).
3. Set `PICKER_APP_ID` in `Code.gs` to your Cloud project number (a public value, safe to commit).
4. Run the **🔑 Set Picker Key** item from the add-on menu (or `SET_PICKER_KEY()` in the editor) and paste the API key.

> 🔐 **Security:** the Browser API key is stored in **Script Properties** — it is never committed to this repository. If a key is ever exposed, delete it and create a new one in the Cloud Console, then update it via the menu item.

---

## 🛡️ Privacy & Security
DocuForm Sync is designed for privacy-first operation. All reading, filtering, and writing happens directly inside your native Google Workspace environment — data never leaves your Google account.

> 🔒 Data stays in your Google Workspace, nothing stored externally.

---

## 📬 Support & Issue Tracking
Encountered a bug or want to suggest a new feature? Please use the **Issues** tab at the top of this repository to log and track your feedback.
