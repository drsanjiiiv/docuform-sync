/**
 * DocuForm Sync - Google Forms add-on
 * Keep form dropdown / multiple-choice / checkbox options in sync
 * with any Google Sheet column. Deploy as a Google Forms add-on.
 */

var APP_NAME = "DocuForm Sync";
var PRIVACY_URL = "https://apps.pwmai.com/privacy-policy/";

var CFG_PREFIX = "DocuFormSync_Config_";
var OLD_CFG_PREFIX = "DocuFormSync_Mappings_";

// ============================ MENU (container-bound / testing) ============================

function onOpen(e) {
  try {
    var ui = FormApp.getUi();
    try {
      ui.createAddonMenu()
        .addItem("⚡ Open DocuForm Sync", "INITIALIZE_ADDON_SIDEBAR")
        .addSeparator()
        .addItem("🔑 Set Picker Key", "SET_PICKER_KEY")
        .addSeparator()
        .addItem("❓ Help", "showHelp")
        .addToUi();
    } catch (err) {
      // Not yet installed as an add-on (bound-script testing).
      ui.createMenu("⚡ DocuForm Sync")
        .addItem("Open DocuForm Sync", "INITIALIZE_ADDON_SIDEBAR")
        .addSeparator()
        .addItem("Set Picker Key", "SET_PICKER_KEY")
        .addSeparator()
        .addItem("Help", "showHelp")
        .addToUi();
    }
  } catch (err) {
    Logger.log("[DocuForm Sync] onOpen error: " + err.toString());
  }
}

function onInstall(e) {
  onOpen(e);
}

// ============================ SIDEBAR LAUNCHER ============================

/**
 * Opens the DocuForm Sync HTML sidebar in the Form editor.
 * Called from the onOpen menu, the homepage card button, or the script editor.
 */
function INITIALIZE_ADDON_SIDEBAR() {
  try {
    var form = FormApp.getActiveForm();
    if (form) {
      PropertiesService.getUserProperties().setProperty("lastUsedFormId", form.getId());
    }
    var html = HtmlService.createHtmlOutputFromFile("UI_Main")
      .setTitle("⚡ " + APP_NAME)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    FormApp.getUi().showSidebar(html);
  } catch (err) {
    Logger.log("[DocuForm Sync] INITIALIZE_ADDON_SIDEBAR error: " + err.toString());
    FormApp.getUi().alert("Could not open the sidebar.\n" + err.message);
  }
}

// ============================ HELP ============================

function showHelp() {
  FormApp.getUi().alert(
    "📚 " + APP_NAME + " - Help",
    "⚡ How it works:\n\n" +
      "1. Open the sidebar (⚡ Open DocuForm Sync).\n" +
      "2. Pick the Google Sheet whose column(s) should feed your form.\n" +
      "3. Map each question to a column.\n" +
      "4. Sync now, or enable Auto-Sync to keep options refreshed.\n\n" +
      "🔒 Data stays in your Google Workspace, nothing stored externally.",
    FormApp.getUi().ButtonSet.OK
  );
}

// ============================ CONFIG HELPERS (shared) ============================

function getFormConfig(formId) {
  if (!formId) return null;
  var props = PropertiesService.getUserProperties();
  var raw = props.getProperty(CFG_PREFIX + formId);
  if (raw) {
    try { return JSON.parse(raw); } catch (e) {}
  }
  var oldRaw = props.getProperty(OLD_CFG_PREFIX + formId);
  if (oldRaw) {
    try { return JSON.parse(oldRaw); } catch (e) {}
  }
  return null;
}

function saveFormConfig(formId, config) {
  PropertiesService.getUserProperties().setProperty(CFG_PREFIX + formId, JSON.stringify(config));
}

function deleteFormConfig(formId) {
  var props = PropertiesService.getUserProperties();
  props.deleteProperty(CFG_PREFIX + formId);
  props.deleteProperty(OLD_CFG_PREFIX + formId);
}

function getAllFormConfigs() {
  var props = PropertiesService.getUserProperties();
  var keys = props.getKeys();
  var out = [];
  for (var i = 0; i < keys.length; i++) {
    var k = keys[i];
    var id = null;
    if (k.indexOf(CFG_PREFIX) === 0) id = k.substring(CFG_PREFIX.length);
    else if (k.indexOf(OLD_CFG_PREFIX) === 0) id = k.substring(OLD_CFG_PREFIX.length);
    if (!id) continue;
    var raw = props.getProperty(k);
    try {
      var c = JSON.parse(raw);
      c.formId = c.formId || id;
      out.push(c);
    } catch (e) {}
  }
  return out;
}

// ============================ FILE PICKER ============================

/**
 * Picker config for the Google Picker API in the sidebar.
 * PICKER_APP_ID is the Cloud project number (public — safe to commit).
 * The Browser API key is stored in Script Properties, NEVER in the repo.
 * Set it once via the "🔑 Set Picker Key" menu item (or SET_PICKER_KEY()).
 */
var PICKER_APP_ID = "939872818516";
var PICKER_KEY_PROP = "DocuFormSync_Picker_DevKey";

function getPickerConfig() {
  return {
    appId: PICKER_APP_ID,
    developerKey: PropertiesService.getScriptProperties().getProperty(PICKER_KEY_PROP) || "",
    token: ScriptApp.getOAuthToken()
  };
}

function SET_PICKER_KEY() {
  try {
    var ui = FormApp.getUi();
    var res = ui.prompt(
      "🔑 Set Picker API Key",
      "Paste your Browser API key (starts with AIza...):",
      ui.ButtonSet.OK_CANCEL
    );
    if (res.getSelectedButton() === ui.Button.OK) {
      var key = (res.getResponseText() || "").trim();
      if (key) {
        PropertiesService.getScriptProperties().setProperty(PICKER_KEY_PROP, key);
        ui.alert("Saved. The 📂 Browse picker will use this key.");
      } else {
        ui.alert("No key entered. Skipped.");
      }
    }
  } catch (e) {
    Logger.log("[DocuForm Sync] SET_PICKER_KEY error: " + e.toString());
  }
}