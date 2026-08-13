/**
 * Router - Public RPC layer exposed to the sidebar via google.script.run.
 * Every function returns a serializable { success, ... } object.
 */

function _wrap(result) {
  if (result instanceof Error) {
    return { success: false, error: result.message };
  }
  if (result && typeof result === "object" && "success" in result) {
    return result;
  }
  return { success: true, data: result };
}

// ============================ IDENTIFICATION ============================

function getFormTitle(formId) {
  try {
    const id = (formId || "").trim();
    const form = id ? (FormApp.getActiveForm() || FormApp.openById(id)) : FormApp.getActiveForm();
    if (!form) return { success: false, error: "No form found." };
    return { success: true, data: form.getTitle() };
  } catch (e) {
    return { success: false, error: "Could not open Form. Check the ID or permissions." };
  }
}

function getActiveFormId() {
  try {
    const form = FormApp.getActiveForm();
    return { success: true, data: form ? form.getId() : null };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function getLinkedSpreadsheet(formId) {
  try {
    const id = (formId || "").trim();
    const form = id ? (FormApp.getActiveForm() || FormApp.openById(id)) : FormApp.getActiveForm();
    if (!form) return { success: false, error: "No form found." };
    const dest = form.getDestination(FormApp.DestinationType.SPREADSHEET);
    if (!dest) {
      return { success: false, error: "No linked spreadsheet found. Add a response destination in the form, or paste a Sheet ID." };
    }
    return { success: true, data: dest.getId() };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function getActiveSheetId() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) return { success: true, data: ss.getId() };
    return { success: false, error: "No active spreadsheet. This add-on runs in Google Forms — use 'Use linked sheet' or paste a Sheet ID." };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function getSheetNames(sheetId) {
  try {
    const ss = SpreadsheetApp.openById(sheetId);
    const sheets = ss.getSheets().map(s => s.getName());
    return { success: true, data: sheets };
  } catch (e) {
    return { success: false, error: "Could not open spreadsheet. Check the ID or permissions. " + e.message };
  }
}

// ============================ QUESTIONS & CHOICES ============================

function getFormQuestions(formId) {
  try {
    const engine = new SyncEngine();
    return engine.getFormQuestions(formId);
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function fetchData(range, sheetId) {
  try {
    if (!range || typeof range !== "string" || range.trim() === "") {
      return { success: false, error: "Invalid or empty range provided." };
    }
    const engine = new SyncEngine();
    return _wrap(engine.fetchData(range.trim(), sheetId));
  } catch (e) {
    return { success: false, error: "Failed to read spreadsheet range: " + e.message };
  }
}

function updateForm(formId, questionId, data) {
  try {
    const engine = new SyncEngine();
    return _wrap(engine.updateForm(formId, questionId, data));
  } catch (e) {
    return { success: false, error: "Failed to update Form item: " + e.message };
  }
}

function populateChoices(formId, questionId, choices) {
  try {
    const engine = new SyncEngine();
    return engine.populateChoices(formId, questionId, choices);
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function previewMapping(formId, mappingJson, sheetId) {
  try {
    const mapping = JSON.parse(mappingJson);
    if (!formId || !mapping) {
      return { success: false, error: "Invalid parameters." };
    }
    const engine = new SyncEngine();
    return engine.previewMapping(formId, mapping, sheetId);
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function applyMappings(formId, mappingsJson, sheetId) {
  try {
    const mappings = JSON.parse(mappingsJson);
    if (!formId || !mappings || !Array.isArray(mappings)) {
      return { success: false, processed: 0, errors: 0, message: "Invalid parameters." };
    }
    const engine = new SyncEngine();
    const result = engine.applyMappings(formId.trim(), mappings, sheetId);
    engine.recordSyncEntry({
      formId: formId.trim(),
      sheetsId: sheetId,
      mappingsCount: mappings.length,
      processed: result.processed,
      errors: result.errors,
      success: result.success,
      mode: "manual"
    });
    return result;
  } catch (e) {
    return { success: false, processed: 0, errors: 0, message: e.message };
  }
}

// ============================ CONFIG ============================

function saveMappingsConfig(formId, configJson, sheetId) {
  try {
    const config = JSON.parse(configJson);
    if (formId) {
      const existing = getFormConfig(formId) || {};
      const merged = Object.assign({}, existing, config);
      merged.formId = formId;
      if (sheetId) merged.sheetId = sheetId;
      saveFormConfig(formId, merged);
    }
    const userProps = PropertiesService.getUserProperties();
    if (formId) userProps.setProperty("lastUsedFormId", formId);
    if (sheetId) userProps.setProperty("lastUsedRange", sheetId);
    return { success: true, message: "Configuration saved." };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function loadMappingsConfig(formId) {
  try {
    const config = getFormConfig(formId);
    return { success: true, data: config };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function resetAllConfig(formId) {
  try {
    if (formId) deleteFormConfig(formId);
    return { success: true, message: "Configuration cleared." };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ============================ AUTO-SYNC ============================

function toggleAutoSync(enabled, intervalMinutes) {
  try {
    if (typeof enabled !== "boolean") {
      return { success: false, message: "Invalid parameter: expected a boolean value." };
    }
    if (enabled) {
      return handleTrigger(intervalMinutes || 60);
    }
    return removeTrigger();
  } catch (e) {
    return { success: false, message: "Failed to toggle Auto Sync: " + e.message };
  }
}

function setAutoSyncConfig(formId, settingsJson) {
  try {
    const settings = JSON.parse(settingsJson);
    const messages = [];

    if (formId) {
      const cfg = getFormConfig(formId) || { formId: formId };
      cfg.autoSyncEnabled = !!settings.enabled;
      cfg.refreshInterval = settings.refreshInterval || 60;
      cfg.refreshOnSubmit = !!settings.refreshOnSubmit;
      saveFormConfig(formId, cfg);
    }

    if (settings.refreshOnSubmit && formId) {
      messages.push(setupOnSubmitTrigger(formId).message);
    } else if (formId) {
      messages.push(removeOnSubmitTrigger(formId).message);
    }

    _syncGlobalAutoTrigger();
    messages.push("Auto-Sync " + (settings.enabled ? "enabled" : "disabled") + ".");

    return { success: true, messages: messages };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Keeps a single global time-driven trigger for the project, active only while
 * at least one saved configuration has Auto-Sync enabled. Multi-form safe.
 */
function _syncGlobalAutoTrigger() {
  removeTrigger();
  const configs = getAllFormConfigs();
  for (const cfg of configs) {
    if (cfg.autoSyncEnabled) {
      handleTrigger(cfg.refreshInterval || 60);
      break;
    }
  }
}

function getAutoSyncConfig(formId) {
  try {
    const cfg = formId ? getFormConfig(formId) : null;
    const status = getTriggerStatus();
    const saved = getSavedConfiguration();
    return {
      success: true,
      data: {
        enabled: !!(cfg && cfg.autoSyncEnabled),
        refreshInterval: (cfg && cfg.refreshInterval) || 60,
        refreshOnSubmit: !!(cfg && cfg.refreshOnSubmit),
        triggerExists: status.exists,
        frequency: status.frequency || null,
        lastSync: saved.lastSync || "Never"
      }
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ============================ HISTORY ============================

function getSyncHistory(formId) {
  try {
    const engine = new SyncEngine();
    return engine.getSyncHistory(formId);
  } catch (e) {
    return { success: false, error: e.message };
  }
}