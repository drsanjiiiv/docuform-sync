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
        .addItem("❓ Help", "showHelp")
        .addToUi();
    } catch (err) {
      // Not yet installed as an add-on (bound-script testing).
      ui.createMenu("⚡ DocuForm Sync")
        .addItem("Open DocuForm Sync", "INITIALIZE_ADDON_SIDEBAR")
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

/**
 * Opens the center-screen picker wizard dialog for a specific question.
 * Called from the sidebar "+" button / "Populate from range" checkbox.
 */
function OPEN_PICKER_DIALOG(questionId) {
  try {
    var form = FormApp.getActiveForm();
    var tpl = HtmlService.createTemplateFromFile("UI_Dialog");
    tpl.formId = form ? form.getId() : "";
    tpl.questionId = questionId || "";
    var html = tpl
      .evaluate()
      .setWidth(760)
      .setHeight(660)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    FormApp.getUi().showModalDialog(html, "⚡ " + APP_NAME);
  } catch (e) {
    Logger.log("[DocuForm Sync] OPEN_PICKER_DIALOG error: " + e.toString());
    try {
      FormApp.getUi().alert("Could not open the picker.\n" + e.message);
    } catch (e2) {}
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
  bumpConfigVersion(formId);
}

function deleteFormConfig(formId) {
  var props = PropertiesService.getUserProperties();
  props.deleteProperty(CFG_PREFIX + formId);
  props.deleteProperty(OLD_CFG_PREFIX + formId);
  bumpConfigVersion(formId);
}

// ============================ CONFIG VERSION (sidebar refresh) ============================

var VER_PROP = "DocuFormSync_Version";

function bumpConfigVersion(formId) {
  if (!formId) return;
  var p = PropertiesService.getUserProperties();
  var v = parseInt(p.getProperty(VER_PROP + "_" + formId), 10) || 0;
  p.setProperty(VER_PROP + "_" + formId, String(v + 1));
}

function getConfigVersion(formId) {
  if (!formId) return 0;
  return parseInt(PropertiesService.getUserProperties().getProperty(VER_PROP + "_" + formId), 10) || 0;
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

// ============================ PER-QUESTION MAPPINGS ============================
// Config shape: { formId, sheetId, sheetName,
//                 questions: { "<questionId>": {enabled, sheetId, sheetName, column,
//                   capacityColumn, sortColumn, sortOrder, startRow, endRow, includeBlank, rangeName} },
//                 namedRanges: { "<rangeName>": {...range} },
//                 autoSyncEnabled, refreshInterval, refreshOnSubmit }

function getQuestionsConfig(formId) {
  var c = getFormConfig(formId) || {};
  return c.questions || {};
}

function saveQuestionConfig(formId, questionId, qcfg) {
  var c = getFormConfig(formId) || { formId: formId };
  c.questions = c.questions || {};
  c.questions[String(questionId)] = qcfg;
  saveFormConfig(formId, c);
}

function deleteQuestionConfig(formId, questionId) {
  var c = getFormConfig(formId);
  if (c && c.questions) {
    delete c.questions[String(questionId)];
    saveFormConfig(formId, c);
  }
}

function getNamedRanges(formId) {
  var c = getFormConfig(formId) || {};
  return c.namedRanges || {};
}

function saveNamedRange(formId, rangeName, range) {
  var c = getFormConfig(formId) || { formId: formId };
  c.namedRanges = c.namedRanges || {};
  c.namedRanges[String(rangeName)] = range;
  saveFormConfig(formId, c);
}