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

/**
 * Fetches the active user's OAuth token to initialize the Google Picker API client-side.
 * @return {Object} { success: true, token: string } or { success: false, error: string }
 */
function getOAuthToken() {
  try {
    return {
      success: true,
      token: ScriptApp.getOAuthToken()
    };
  } catch (err) {
    return {
      success: false,
      error: err.toString()
    };
  }
}

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
    let destId = null;
    try {
      destId = form.getDestinationId();
    } catch (e) {}
    if (!destId) {
      return { success: false, error: "No linked spreadsheet found. Add a response destination in the form, or paste a Sheet ID / Browse." };
    }
    return { success: true, data: destId };
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

function clearQuestionChoices(formId, questionId) {
  try {
    const engine = new SyncEngine();
    return engine.clearQuestionChoices(formId, questionId);
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

// ============================ SIMPLE FLOW (Form Ranger style) ============================

function getSheetHeaders(sheetId, sheetName) {
  try {
    const ss = SpreadsheetApp.openById(sheetId);
    const sheet = sheetName ? ss.getSheetByName(sheetName) : ss.getSheets()[0];
    if (!sheet) return { success: false, error: 'Sheet not found.' };
    const lastCol = sheet.getLastColumn();
    const headers = lastCol > 0
      ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h == null ? '' : h).trim())
      : [];
    return { success: true, data: headers };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function previewColumn(sheetId, sheetName, columnLetter) {
  try {
    const engine = new SyncEngine();
    const colIndex = engine._columnLetterToIndex(columnLetter);
    const ss = SpreadsheetApp.openById(sheetId);
    const sheet = sheetName ? ss.getSheetByName(sheetName) : ss.getSheets()[0];
    if (!sheet) return { success: false, error: 'Sheet not found.' };
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { success: true, data: [], stats: { blanksRemoved: 0, duplicatesRemoved: 0 }, problemRows: { blanks: [], duplicates: [] } };
    }
    const raw = sheet.getRange(2, colIndex, lastRow - 1, 1).getValues().flat();
    let lastDataIdx = -1;
    for (let i = 0; i < raw.length; i++) {
      if (String(raw[i] == null ? '' : raw[i]).trim() !== '') lastDataIdx = i;
    }
    const seen = {};
    const firstRowOf = {};
    const values = [];
    const blankRows = [];
    const duplicateRows = [];
    let blanks = 0;
    let dups = 0;
    for (let i = 0; i < raw.length; i++) {
      const row = 2 + i;
      const v = String(raw[i] == null ? '' : raw[i]).trim();
      if (v === '') {
        blanks++;
        if (i < lastDataIdx) blankRows.push({ row: row, value: '' });
        continue;
      }
      const key = v.toLowerCase();
      if (seen[key]) {
        dups++;
        duplicateRows.push({ row: row, value: v, duplicateOfRow: firstRowOf[key] });
        continue;
      }
      seen[key] = true;
      firstRowOf[key] = row;
      values.push(v);
    }
    return {
      success: true,
      data: values,
      stats: { blanksRemoved: blanks, duplicatesRemoved: dups },
      problemRows: { blanks: blankRows, duplicates: duplicateRows }
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function saveQuestionMapping(formId, mappingJson) {
  try {
    const m = JSON.parse(mappingJson);
    if (!formId || !m || !m.questionId || !m.sheetId || !m.column) {
      return { success: false, error: 'Invalid mapping. Sheet and column are required.' };
    }
    const qcfg = {
      enabled: true,
      sheetId: m.sheetId,
      sheetName: m.sheetName || '',
      column: m.column,
      capacityColumn: m.capacityColumn || '',
      sortColumn: m.sortColumn || '',
      sortOrder: m.sortOrder || 'none',
      startRow: m.startRow || 1,
      endRow: m.endRow || 0,
      includeBlank: !!m.includeBlank,
      rangeName: m.rangeName || ''
    };
    saveQuestionConfig(formId, String(m.questionId), qcfg);
    if (m.rangeName && m.sheetId && m.column) {
      saveNamedRange(formId, m.rangeName, {
        sheetId: m.sheetId,
        sheetName: m.sheetName || '',
        column: m.column,
        capacityColumn: qcfg.capacityColumn,
        sortColumn: qcfg.sortColumn,
        sortOrder: qcfg.sortOrder,
        startRow: qcfg.startRow,
        endRow: qcfg.endRow,
        includeBlank: qcfg.includeBlank
      });
    }
    return { success: true, message: 'Question mapping saved.' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function clearQuestionMapping(formId, questionId) {
  try {
    if (formId && questionId) deleteQuestionConfig(formId, questionId);
    return { success: true, message: 'Mapping cleared.' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function reuseSavedRange(formId, questionId, rangeName) {
  try {
    if (!formId || !questionId || !rangeName) {
      return { success: false, error: 'Missing range or question.' };
    }
    const c = getFormConfig(formId) || {};
    const r = (c.namedRanges || {})[String(rangeName)];
    if (!r) return { success: false, error: 'Saved range "' + rangeName + '" not found.' };
    const qcfg = {
      enabled: true,
      sheetId: r.sheetId,
      sheetName: r.sheetName || '',
      column: r.column,
      capacityColumn: r.capacityColumn || '',
      sortColumn: r.sortColumn || '',
      sortOrder: r.sortOrder || 'none',
      startRow: r.startRow || 1,
      endRow: r.endRow || 0,
      includeBlank: !!r.includeBlank,
      rangeName: String(rangeName)
    };
    saveQuestionConfig(formId, String(questionId), qcfg);
    return syncQuestion(formId, String(questionId));
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function loadQuestionMappings(formId) {
  try {
    const c = getFormConfig(formId) || {};
    return {
      success: true,
      data: {
        questions: c.questions || {},
        namedRanges: c.namedRanges || {},
        sheetId: c.sheetId || '',
        autoSync: {
          autoSyncEnabled: !!c.autoSyncEnabled,
          refreshInterval: c.refreshInterval || 60,
          refreshOnSubmit: !!c.refreshOnSubmit
        }
      }
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function syncQuestion(formId, questionId) {
  try {
    const qcfg = getQuestionsConfig(formId)[String(questionId)];
    if (!qcfg || !qcfg.enabled || !qcfg.sheetId) {
      return { success: false, processed: 0, errors: 0, message: 'No mapping configured for this question.' };
    }
    const engine = new SyncEngine();
    const result = engine.applyMappings(formId, [Object.assign({ questionId: String(questionId) }, qcfg)], qcfg.sheetId);
    engine.recordSyncEntry({
      formId: formId,
      sheetsId: qcfg.sheetId,
      mappingsCount: 1,
      processed: result.processed,
      errors: result.errors,
      success: result.success,
      mode: 'manual'
    });
    return result;
  } catch (e) {
    return { success: false, processed: 0, errors: 0, message: e.message };
  }
}

function syncAll(formId) {
  try {
    const qs = getQuestionsConfig(formId);
    const engine = new SyncEngine();
    const mappings = [];
    Object.keys(qs).forEach(qid => {
      const q = qs[qid];
      if (q && q.enabled && q.sheetId) mappings.push(Object.assign({ questionId: qid }, q));
    });
    if (mappings.length === 0) {
      return { success: false, processed: 0, errors: 0, message: 'No enabled question mappings to sync.' };
    }
    let processed = 0;
    let errors = 0;
    let cleanedBlanks = 0;
    let cleanedDuplicates = 0;
    const results = [];
    const bySheet = {};
    mappings.forEach(m => { (bySheet[m.sheetId] = bySheet[m.sheetId] || []).push(m); });
    Object.keys(bySheet).forEach(sid => {
      const res = engine.applyMappings(formId, bySheet[sid], sid);
      processed += res.processed;
      errors += res.errors;
      cleanedBlanks += (res.cleaned && res.cleaned.blanks) || 0;
      cleanedDuplicates += (res.cleaned && res.cleaned.duplicates) || 0;
      results.push.apply(results, res.results);
    });
    const success = errors === 0 && processed > 0;
    engine.recordSyncEntry({
      formId: formId,
      sheetsId: '',
      mappingsCount: mappings.length,
      processed: processed,
      errors: errors,
      success: success,
      mode: 'manual'
    });
    return {
      success: success,
      processed: processed,
      errors: errors,
      results: results,
      cleaned: { blanks: cleanedBlanks, duplicates: cleanedDuplicates },
      message: `Processed ${processed}/${mappings.length} mappings.`
    };
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
    if (settings.enabled) {
      const trig = _getTriggerResult();
      if (trig && trig.success === false) {
        messages.push("Warning: schedule trigger could not be created (" + (trig.message || "unknown") + ").");
      }
    }
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
var _lastTriggerResult = null;

function _getTriggerResult() {
  return _lastTriggerResult;
}

function _syncGlobalAutoTrigger() {
  _lastTriggerResult = null;
  removeTrigger();
  const configs = getAllFormConfigs();
  for (const cfg of configs) {
    if (cfg.autoSyncEnabled) {
      _lastTriggerResult = handleTrigger(cfg.refreshInterval || 60);
      break;
    }
  }
}

function getAutoSyncConfig(formId) {
  try {
    let cfg = null;
    try { cfg = formId ? getFormConfig(formId) : null; } catch (e) {}
    let status = { exists: false, handler: null, frequency: null };
    let triggerWarning = "";
    try {
      status = getTriggerStatus();
    } catch (e) {
      triggerWarning = e.message;
      Logger.log("[DocuForm Sync] getTriggerStatus failed: " + e.toString());
    }
    let saved = { lastSync: "Never" };
    try { saved = getSavedConfiguration(); } catch (e) {}
    return {
      success: true,
      data: {
        enabled: !!(cfg && cfg.autoSyncEnabled),
        refreshInterval: (cfg && cfg.refreshInterval) || 60,
        refreshOnSubmit: !!(cfg && cfg.refreshOnSubmit),
        triggerExists: status.exists,
        frequency: status.frequency || null,
        lastSync: saved.lastSync || "Never",
        triggerWarning: triggerWarning
      }
    };
  } catch (e) {
    Logger.log("[DocuForm Sync] getAutoSyncConfig failed: " + e.toString());
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