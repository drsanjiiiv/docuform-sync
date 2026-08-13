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
    if (cfg && cfg.questions && Object.keys(cfg.questions).length > 0) {
      configs.push(cfg);
    }
  } else if (savedFormId) {
    const cfg = getFormConfig(savedFormId);
    if (cfg && cfg.questions) configs.push(cfg);
  }
  if (configs.length === 0) {
    configs = getAllFormConfigs().filter(c => c.questions && Object.keys(c.questions).length > 0);
  }

  if (configs.length === 0) {
    Logger.log("[DocuForm Sync] Auto-Sync skipped: no saved configurations found.");
    return;
  }

  const mode = eventFormId ? "submit" : "auto";
  let anySuccess = false;

  for (const cfg of configs) {
    const formId = cfg.formId || savedFormId;
    if (!formId) continue;
    const qs = cfg.questions || {};
    const mappings = [];
    Object.keys(qs).forEach(qid => {
      const q = qs[qid];
      if (q && q.enabled && q.sheetId) mappings.push(Object.assign({ questionId: qid }, q));
    });
    if (mappings.length === 0) {
      Logger.log("[DocuForm Sync] Auto-Sync skipped: no enabled mappings for a config.");
      continue;
    }
    try {
      engine.resetTimer();
      let processed = 0;
      let errors = 0;
      const bySheet = {};
      mappings.forEach(m => { (bySheet[m.sheetId] = bySheet[m.sheetId] || []).push(m); });
      Object.keys(bySheet).forEach(sid => {
        const result = engine.applyMappings(formId, bySheet[sid], sid);
        processed += result.processed;
        errors += result.errors;
      });
      engine.recordSyncEntry({
        formId: formId,
        sheetsId: "",
        mappingsCount: mappings.length,
        processed: processed,
        errors: errors,
        success: errors === 0 && processed > 0,
        mode: mode
      });
      if (processed > 0) {
        anySuccess = true;
        props.setProperty(KEYS.LAST_SYNC_TIMESTAMP, new Date().toISOString());
      }
      Logger.log(`[DocuForm Sync] ${mode} sync for ${formId}: processed ${processed}/${mappings.length}.`);
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
    let form = null;
    try { form = FormApp.getActiveForm(); } catch (e) {}
    if (!form && formId) form = FormApp.openById(formId);
    if (!form) return { success: false, message: "Could not resolve the form for the on-submit trigger." };
    const formSourceId = form.getId();
    const existingTriggers = ScriptApp.getProjectTriggers();
    for (const t of existingTriggers) {
      if (t.getHandlerFunction() === "performAutoSync" &&
          t.getTriggerSourceId() === formSourceId &&
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