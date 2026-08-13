# DocuForm Sync ⚡

A modern, highly intuitive Google Forms add-on that keeps your **Dropdown**, **Multiple Choice**, **Checkbox**, and **Grid** options in perfect sync with any Google Sheet column — no more manual copy-paste whenever your data changes.

The add-on runs as a clean, card-based sidebar inside your Google Form editor, letting you map sheet columns to questions in seconds.

---

## ✨ Key Differentiators & Features

Unlike traditional sync tools, DocuForm Sync gives you complete control straight from the sidebar:

* **Simple Question-by-Question Flow:** Each supported question gets a **Populate from range** toggle — a `+` opens a 3-step wizard (Select Sheet → Select Range with live preview → Name range) and **Save & Sync** updates it instantly.
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
* `UI_Main.html` - The sidebar dashboard (Questions / Auto-repopulate / Sync).
* `appsscript.json` - Apps Script manifest (runtime + OAuth scopes).
* `Privacy-Policy.md` - Our data protection compliance standards.
* `Terms-of-Service.md` - Terms and usage guidelines.

---

## 📂 Choosing a Spreadsheet

The **Select Sheet** step works **out of the box** — no API keys or Cloud setup needed:

1. **📂 Browse** opens a searchable spreadsheet picker (with Load more for large Drives), or
2. Paste a Sheet **link / ID**, or
3. Use **⚡ Linked sheet** (the form's response destination).

The full picker is handled server-side inside the add-on, so there is nothing for end users (e.g. MSME staff) to configure.

---

## 🛡️ Privacy & Security
DocuForm Sync is designed for privacy-first operation. All reading, filtering, and writing happens directly inside your native Google Workspace environment — data never leaves your Google account.

> 🔒 Data stays in your Google Workspace, nothing stored externally.

---

## 📬 Support & Issue Tracking
Encountered a bug or want to suggest a new feature? Please use the **Issues** tab at the top of this repository to log and track your feedback.
