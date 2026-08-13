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

The **📂 Browse** button uses the [Google Picker API](https://developers.google.com/drive/picker) with the `drive.file` scope, so users grant access to exactly the sheet they select. To enable it, configure your Google Cloud project and fill in two constants at the top of `Code.gs`:

1. Open the Apps Script project → **Project Settings** → note the **Google Cloud Platform (GCP) Project number**.
2. In the Cloud Console for that project:
   - Enable the **Picker API**.
   - Create a **Browser API key** (no referrer restriction).
3. In `Code.gs`, set:
   - `PICKER_APP_ID = "<your Cloud project number>"` (e.g. `123456789012`)
   - `PICKER_DEVELOPER_KEY = "<your Browser API key>"`

Until these are set, the Browse button shows a hint and the manual "paste a Sheet ID" / "⚡ Use linked sheet" paths still work.

---

## 🛡️ Privacy & Security
DocuForm Sync is designed for privacy-first operation. All reading, filtering, and writing happens directly inside your native Google Workspace environment — data never leaves your Google account.

> 🔒 Data stays in your Google Workspace, nothing stored externally.

---

## 📬 Support & Issue Tracking
Encountered a bug or want to suggest a new feature? Please use the **Issues** tab at the top of this repository to log and track your feedback.
