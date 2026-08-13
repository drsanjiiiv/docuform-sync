/**
 * SyncEngine - Core logic for reading Sheet columns and writing form choices.
 * No dependency on an "active" spreadsheet, so it also works from
 * time-driven and on-form-submit triggers.
 */
class SyncEngine {

  constructor() {
    this._maxMs = 340000;
    this._startAt = Date.now();
  }

  _remainingMs() {
    return Math.max(0, this._maxMs - (Date.now() - this._startAt));
  }

  resetTimer() {
    this._startAt = Date.now();
    return this;
  }

  // ============================ SHEET READING ============================

  fetchData(rangeNotation, sheetId) {
    this.resetTimer();
    if (!rangeNotation) {
      return { success: false, error: "No Range Provided." };
    }
    let ss = null;
    try {
      ss = sheetId ? SpreadsheetApp.openById(sheetId) : SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) return { success: false, error: "No active spreadsheet. Paste a Sheet ID instead." };
    } catch (e) {
      return { success: false, error: "Could not open spreadsheet: " + e.message };
    }
    let range;
    try {
      range = ss.getRange(rangeNotation);
    } catch (e) {
      return { success: false, error: "Invalid Range: Could not find the specified area." };
    }
    const values = range.getValues();
    if (values.length === 0 || (values.length === 1 && values[0].length === 0)) {
      return { success: false, error: "Empty Range: No data found in selected area." };
    }
    const cleanData = values.filter(row => row.some(cell => cell !== "" && cell !== null));
    return { success: true, data: cleanData };
  }

  fetchColumnData(sheetId, sheetName, columnLetter, startRow, endRow) {
    try {
      const sheet = this._openSheet(sheetId, sheetName);
      if (!sheet) return { success: false, error: `Sheet "${sheetName}" not found.` };
      const raw = this._readColumn(sheet, columnLetter, startRow || 1, endRow || 0);
      const values = raw.filter(v => v !== "" && v !== null);
      return { success: true, data: values };
    } catch (e) {
      return { success: false, error: e.message };
    }
  }

  _openSheet(sheetId, sheetName) {
    const ss = SpreadsheetApp.openById(sheetId);
    let sheet = null;
    if (sheetName) {
      sheet = ss.getSheetByName(sheetName);
    } else {
      sheet = ss.getSheets()[0];
    }
    return sheet;
  }

  _readColumn(sheet, columnLetter, startRow, endRow) {
    const colIndex = this._columnLetterToIndex(columnLetter);
    const lastRow = sheet.getLastRow();
    const start = Math.max(startRow || 1, 1);
    let end = (endRow && endRow > 0) ? endRow : lastRow;
    end = Math.min(end, lastRow);
    if (start > end) return [];
    return sheet.getRange(start, colIndex, end - start + 1, 1).getValues().flat();
  }

  /**
   * Reads the mapped column(s) over the same row range so choice / capacity /
   * sort columns stay aligned, then applies capacity filtering, sorting and
   * the optional blank option.
   */
  _prepareChoices(sheetId, mapping) {
    const sheet = this._openSheet(sheetId, mapping.sheetName);
    if (!sheet) {
      return { error: `Sheet "${mapping.sheetName || '(first tab)'}" not found in the target spreadsheet.` };
    }
    const lastRow = sheet.getLastRow();
    const start = Math.max((mapping.startRow || 1), 1);
    let end = (mapping.endRow && mapping.endRow > 0) ? mapping.endRow : lastRow;
    end = Math.min(end, lastRow);
    if (start > end) {
      return { choices: [], removedChoices: 0, totalRows: 0, rawCount: 0 };
    }
    const totalRows = end - start + 1;
    const rawChoices = this._readColumn(sheet, mapping.column, start, end);
    let rawCaps = null;
    let rawSort = null;
    if (mapping.capacityColumn) {
      rawCaps = this._readColumn(sheet, mapping.capacityColumn, start, end);
    }
    if (mapping.sortColumn) {
      rawSort = this._readColumn(sheet, mapping.sortColumn, start, end);
    }

    const pairs = [];
    let removedChoices = 0;
    let blankRemoved = 0;
    let duplicateRemoved = 0;
    const seen = {};
    for (let i = 0; i < rawChoices.length; i++) {
      const v = String(rawChoices[i] == null ? "" : rawChoices[i]).trim();
      if (v === "") {
        blankRemoved++;
        continue;
      }
      if (rawCaps) {
        const maxVal = Number(rawCaps[i]);
        if (!isNaN(maxVal) && maxVal <= 0) {
          removedChoices++;
          continue;
        }
      }
      const key = v.toLowerCase();
      if (seen[key]) {
        duplicateRemoved++;
        continue;
      }
      seen[key] = true;
      let sortKey = v;
      if (rawSort) {
        sortKey = String(rawSort[i] == null ? "" : rawSort[i]).trim();
      }
      pairs.push({ choice: v, sortKey: sortKey });
    }

    const order = mapping.sortOrder === "desc" ? -1 : 1;
    if (mapping.sortOrder && mapping.sortOrder !== "none") {
      pairs.sort((a, b) => order * String(a.sortKey).localeCompare(String(b.sortKey), undefined, { sensitivity: "base", numeric: true }));
    }

    let choices = pairs.map(p => p.choice);
    if (mapping.includeBlank) {
      choices.unshift("");
    }
    return {
      choices,
      removedChoices,
      blankRemoved,
      duplicateRemoved,
      totalRows,
      rawCount: pairs.length
    };
  }

  // ============================ FORM READING ============================

  getFormQuestions(formId) {
    try {
      const form = this._openForm(formId);
      if (!form) return { success: false, error: "No active Form found." };
      const items = form.getItems();
      const questions = [];
      for (const item of items) {
        const type = item.getType();
        const supported = [
          FormApp.ItemType.MULTIPLE_CHOICE,
          FormApp.ItemType.LIST,
          FormApp.ItemType.CHECKBOX,
          FormApp.ItemType.GRID
        ];
        questions.push({
          id: item.getId().toString(),
          title: item.getTitle(),
          type: type,
          typeName: this._getTypeName(type),
          supported: supported.includes(type)
        });
      }
      return { success: true, data: questions };
    } catch (e) {
      return { success: false, error: e.message };
    }
  }

  _openForm(formId) {
    const id = (formId || "").trim();
    try {
      const active = FormApp.getActiveForm();
      if (active) return active;
    } catch (e) {}
    if (id) {
      return FormApp.openById(id);
    }
    return null;
  }

  // ============================ WRITE CHOICES ============================

  populateChoices(formId, questionId, choices) {
    try {
      const form = this._openForm(formId);
      if (!form) return { success: false, message: "No active Form found." };
      const targetItem = this._findItem(form, questionId);
      if (!targetItem) {
        return { success: false, message: "Question ID not found." };
      }
      const type = targetItem.getType();
      const seen = {};
      let hasBlank = false;
      const cleanChoices = [];
      (choices || []).forEach(c => {
        const v = c == null ? "" : String(c).trim();
        if (v === "") {
          if (!hasBlank) { hasBlank = true; cleanChoices.push(v); }
          return;
        }
        const key = v.toLowerCase();
        if (seen[key]) return;
        seen[key] = true;
        cleanChoices.push(v);
      });

      if (cleanChoices.length === 0) {
        return { success: false, message: "No valid choices to populate." };
      }

      switch (type) {
        case FormApp.ItemType.MULTIPLE_CHOICE:
          targetItem.asMultipleChoiceItem().setChoiceValues(this._toFormValues(cleanChoices));
          break;
        case FormApp.ItemType.LIST:
          targetItem.asListItem().setChoiceValues(this._toFormValues(cleanChoices));
          break;
        case FormApp.ItemType.CHECKBOX:
          targetItem.asCheckboxItem().setChoiceValues(this._toFormValues(cleanChoices));
          break;
        case FormApp.ItemType.GRID:
          targetItem.asGridItem().setColumns(cleanChoices.filter(c => c !== ""));
          break;
        default:
          return { success: false, message: "Unsupported question type for choice population." };
      }
      return { success: true, message: `Populated ${cleanChoices.length} choices.` };
    } catch (e) {
      return { success: false, message: e.message };
    }
  }

  _findItem(form, questionId) {
    const items = form.getItems();
    const target = String(questionId);
    for (const item of items) {
      if (item.getId().toString() === target) return item;
    }
    return null;
  }

  _toFormValues(choices) {
    return choices.map(c => c === "" ? "\u200B" : c);
  }

  // ============================ SYNC ============================

  applyMappings(formId, mappings, sheetId) {
    this.resetTimer();
    if (!sheetId) return { success: false, processed: 0, errors: 0, results: [], message: "No Spreadsheet ID configured." };
    let processed = 0;
    let errors = 0;
    let cleanedBlanks = 0;
    let cleanedDuplicates = 0;
    const results = [];

    for (let i = 0; i < mappings.length; i++) {
      if (this._remainingMs() <= 3000) {
        results.push({ index: i, status: "timeout", message: "Execution time limit reached." });
        break;
      }
      const mapping = mappings[i];
      try {
        if (!mapping.questionId || !mapping.column) {
          results.push({ index: i, status: "skipped", message: "Missing questionId or column." });
          errors++;
          continue;
        }
        if (mapping.enabled === false) {
          results.push({ index: i, status: "skipped", message: "Mapping disabled." });
          continue;
        }
        const prep = this._prepareChoices(sheetId, mapping);
        if (prep.error) {
          results.push({ index: i, status: "error", message: prep.error });
          errors++;
          continue;
        }
        if (prep.choices.length === 0) {
          results.push({ index: i, status: "skipped", message: "No choices found in the selected column." });
          continue;
        }
        cleanedBlanks += prep.blankRemoved || 0;
        cleanedDuplicates += prep.duplicateRemoved || 0;
        const popResult = this.populateChoices(formId, mapping.questionId, prep.choices);
        if (popResult.success) {
          processed++;
          let note = popResult.message;
          if (prep.blankRemoved || prep.duplicateRemoved) {
            note += " (removed " + (prep.blankRemoved || 0) + " blank, " + (prep.duplicateRemoved || 0) + " duplicate)";
          }
          results.push({ index: i, status: "ok", message: note });
        } else {
          errors++;
          results.push({ index: i, status: "failed", message: popResult.message });
        }
      } catch (e) {
        errors++;
        results.push({ index: i, status: "error", message: e.message });
      }
    }
    return {
      success: errors === 0 && processed > 0,
      processed,
      errors,
      results,
      cleaned: { blanks: cleanedBlanks, duplicates: cleanedDuplicates },
      message: `Processed ${processed}/${mappings.length} mappings.`
    };
  }

  previewMapping(formId, mapping, sheetId) {
    try {
      if (!sheetId) return { success: false, error: "No Spreadsheet ID provided." };
      const prep = this._prepareChoices(sheetId, mapping);
      if (prep.error) return { success: false, error: prep.error };
      return {
        success: true,
        data: {
          choices: prep.choices,
          totalRows: prep.totalRows,
          validChoices: prep.choices.filter(c => c !== "").length,
          blankIncluded: prep.choices.length > 0 && prep.choices[0] === "",
          removedChoices: prep.removedChoices,
          blankRemoved: prep.blankRemoved,
          duplicateRemoved: prep.duplicateRemoved
        }
      };
    } catch (e) {
      return { success: false, error: e.message };
    }
  }

  // ============================ HISTORY ============================

  getSyncHistory(formId) {
    try {
      const userProps = PropertiesService.getUserProperties();
      const historyRaw = userProps.getProperty("DocuFormSync_SyncHistory");
      const history = historyRaw ? JSON.parse(historyRaw) : [];
      const filtered = formId ? history.filter(h => h.formId === formId) : history;
      return { success: true, data: filtered.slice(-50) };
    } catch (e) {
      return { success: false, error: e.message };
    }
  }

  recordSyncEntry(entry) {
    try {
      const userProps = PropertiesService.getUserProperties();
      const historyRaw = userProps.getProperty("DocuFormSync_SyncHistory");
      const history = historyRaw ? JSON.parse(historyRaw) : [];
      history.push({
        timestamp: new Date().toISOString(),
        formId: entry.formId,
        sheetsId: entry.sheetsId,
        mappingsCount: entry.mappingsCount || 0,
        processed: entry.processed || 0,
        errors: entry.errors || 0,
        success: entry.success || false,
        mode: entry.mode || "manual"
      });
      if (history.length > 200) history.splice(0, history.length - 200);
      userProps.setProperty("DocuFormSync_SyncHistory", JSON.stringify(history));
      return true;
    } catch (e) {
      return false;
    }
  }

  // ============================ SINGLE QUESTION EDIT ============================

  updateForm(formId, questionId, data) {
    let form;
    try {
      form = this._openForm(formId);
      if (!form) return { success: false, message: "No active Form found." };
    } catch (e) {
      return { success: false, message: "Access Denied: Cannot open Form." };
    }
    let title = null;
    let description = null;
    if (typeof data === "object") {
      title = data.title ?? data.name;
      description = data.description ?? data.text;
    } else if (typeof data === "string") {
      description = data;
    }
    const targetItem = this._findItem(form, questionId);
    if (!targetItem) {
      return { success: false, message: "Question ID not found." };
    }
    try {
      const itemType = targetItem.getType();
      let editableItem;
      switch (itemType) {
        case FormApp.ItemType.TEXT_ITEM:
        case FormApp.ItemType.PARAGRAPH_TEXT_ITEM:
          editableItem = targetItem.asTextItem();
          break;
        case FormApp.ItemType.MULTIPLE_CHOICE:
          editableItem = targetItem.asMultipleChoiceItem();
          break;
        case FormApp.ItemType.CHECKBOX:
          editableItem = targetItem.asCheckboxItem();
          break;
        case FormApp.ItemType.LIST:
          editableItem = targetItem.asListItem();
          break;
        default:
          return { success: false, message: "Unsupported question type." };
      }
      let updated = false;
      if (title !== null && editableItem.setTitle) {
        editableItem.setTitle(String(title));
        updated = true;
      }
      if (description !== null && editableItem.setDescription) {
        editableItem.setDescription(String(description));
        updated = true;
      }
      if (!updated) {
        return { success: true, message: "No changes made (data matches or empty)." };
      }
      return { success: true, message: "Question updated." };
    } catch (e) {
      return { success: false, message: "Internal update error." };
    }
  }

  // ============================ UTILITIES ============================

  _columnLetterToIndex(letter) {
    let index = 0;
    const upper = String(letter || "A").toUpperCase().replace(/[^A-Z]/g, "");
    if (!upper) return 1;
    for (let i = 0; i < upper.length; i++) {
      index = index * 26 + (upper.charCodeAt(i) - 64);
    }
    return index;
  }

  _getTypeName(type) {
    const names = {};
    names[FormApp.ItemType.MULTIPLE_CHOICE] = "Multiple Choice";
    names[FormApp.ItemType.LIST] = "Dropdown";
    names[FormApp.ItemType.CHECKBOX] = "Checkboxes";
    names[FormApp.ItemType.GRID] = "Grid";
    names[FormApp.ItemType.TEXT_ITEM] = "Text";
    names[FormApp.ItemType.PARAGRAPH_TEXT_ITEM] = "Paragraph";
    names[FormApp.ItemType.SCALE_ITEM] = "Scale";
    names[FormApp.ItemType.DATE_ITEM] = "Date";
    names[FormApp.ItemType.TIME_ITEM] = "Time";
    return names[type] || "Unknown";
  }
}