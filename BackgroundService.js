/**
 * BackgroundService - Auto-sync execution + trigger management for
 * time-driven and on-form-submit runs. Never depends on an active
 * spreadsheet, so it works reliably outside the sidebar.
 */

var PROP_NAMESPACE = "DocuFormSync_Config";

var KEYS = {
  SPREADSHEET_RANGE: "lastUsedRange",
  FORM_ID: "lastUsedFormId",
  QUESTION_ID: "lastUsedQuestionId",
  LAST_SYNC_TIMESTAMP: "lastSyncTimestamp"
};

/**
 * Entry point for both time-driven and on-form-submit triggers.
 * @param {Object} e - Trigger event. e.source is the Form on submit runs.
 */
function performAutoSync(e) {
  const props = PropertiesService.getUserProperties();
  const engine = new SyncEngine();

  let eventFormId = null;
  if (e && e.source && typeof e.source.getId === "function") {
    try { eventFormId = e.source.getId(); } catch (err) {}
  }
  const savedFormId = props.getProperty(KEYS.FORM_ID);

  let configs = [];
  if (eventFormId) {
    const cfg = getFormConfig(eventFormId);
    if (cfg && Array.isArray(cfg.mappings) && cfg.mappings.length > 0) {
      configs.push(cfg);
    }
  } else if (savedFormId) {
    const cfg = getFormConfig(savedFormId);
    if (cfg && Array.isArray(cfg.mappings)) configs.push(cfg);
  }
  if (configs.length === 0) {
    configs = getAllFormConfigs().filter(c => Array.isArray(c.mappings) && c.mappings.length > 0);
  }

  if (configs.length === 0) {
    Logger.log("[DocuForm Sync] Auto-Sync skipped: no saved configurations found.");
    return;
  }

  const mode = eventFormId ? "submit" : "auto";
  let anySuccess = false;

  for (const cfg of configs) {
    const formId = cfg.formId || savedFormId;
    const sheetId = cfg.sheetId || props.getProperty(KEYS.SPREADSHEET_RANGE);
    if (!formId || !sheetId) {
      Logger.log("[DocuForm Sync] Auto-Sync skipped: missing form/sheet ID for a config.");
      continue;
    }
    try {
      engine.resetTimer();
      const result = engine.applyMappings(formId, cfg.mappings || [], sheetId);
      engine.recordSyncEntry({
        formId: formId,
        sheetsId: sheetId,
        mappingsCount: (cfg.mappings || []).length,
        processed: result.processed,
        errors: result.errors,
        success: result.success,
        mode: mode
      });
      if (result.processed > 0 || result.success) {
        anySuccess = true;
        props.setProperty(KEYS.LAST_SYNC_TIMESTAMP, new Date().toISOString());
      }
      Logger.log(`[DocuForm Sync] ${mode} sync for ${formId}: ${result.message}`);
    } catch (err) {
      Logger.log(`[DocuForm Sync] ERROR syncing ${formId}: ${err.toString()}`);
    }
  }
}

function persistSyncConfiguration(range, formId, questionId) {
  try {
    const userProps = PropertiesService.getUserProperties();
    if (range) userProps.setProperty(KEYS.SPREADSHEET_RANGE, range);
    if (formId) userProps.setProperty(KEYS.FORM_ID, formId);
    if (questionId) userProps.setProperty(KEYS.QUESTION_ID, questionId);
    Logger.log("[DocuForm Sync] Configuration persisted.");
    return true;
  } catch (e) {
    Logger.log("[DocuForm Sync] Failed to persist config: " + e.toString());
    return false;
  }
}

function getSavedConfiguration() {
  const userProps = PropertiesService.getUserProperties();
  return {
    range: userProps.getProperty(KEYS.SPREADSHEET_RANGE) || "",
    formId: userProps.getProperty(KEYS.FORM_ID) || "",
    questionId: userProps.getProperty(KEYS.QUESTION_ID) || "",
    lastSync: userProps.getProperty(KEYS.LAST_SYNC_TIMESTAMP) || "Never"
  };
}

// ============================ ON-SUBMIT TRIGGER ============================

function setupOnSubmitTrigger(formId) {
  try {
    const form = FormApp.openById(formId);
    const existingTriggers = ScriptApp.getProjectTriggers();
    for (const t of existingTriggers) {
      if (t.getHandlerFunction() === "performAutoSync" &&
          t.getTriggerSourceId() === formId &&
          t.getEventType() === ScriptApp.EventType.ON_FORM_SUBMIT) {
        Logger.log("[DocuForm Sync] On-submit trigger already exists.");
        return { success: true, message: "Trigger already exists." };
      }
    }
    ScriptApp.newTrigger("performAutoSync")
      .forForm(form)
      .onFormSubmit()
      .create();
    Logger.log("[DocuForm Sync] On-submit trigger created.");
    return { success: true, message: "On-submit trigger created." };
  } catch (e) {
    Logger.log("[DocuForm Sync] Failed to create on-submit trigger: " + e.toString());
    return { success: false, message: e.message };
  }
}

function removeOnSubmitTrigger(formId) {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    let count = 0;
    for (const t of triggers) {
      if (t.getHandlerFunction() === "performAutoSync" &&
          t.getEventType() === ScriptApp.EventType.ON_FORM_SUBMIT &&
          (!formId || t.getTriggerSourceId() === formId)) {
        ScriptApp.deleteTrigger(t);
        count++;
      }
    }
    return { success: true, message: `Removed ${count} trigger(s).` };
  } catch (e) {
    return { success: false, message: e.message };
  }
}