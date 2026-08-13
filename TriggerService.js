/**
 * @fileoverview TriggerService - Manages Google Apps Script time-driven triggers.
 *              Responsible for creating, deleting, and checking the status of
 *              the "performAutoSync" trigger.
 */

// ============================ CONSTANTS ============================

/** The name of the handler function to trigger */
const TRIGGER_HANDLER = 'performAutoSync';

/** Namespace/Tag identifier for our specific triggers */
const TRIGGER_TAG = 'DocuFormSync_AutoSync';

// ============================ PUBLIC API ============================

/**
 * Main entry point called by UI or other services.
 * Attempts to create the trigger if it doesn't exist.
 * 
 * @param {number} intervalMinutes - The desired interval in minutes (e.g., 15, 60).
 * @returns {{success: boolean, message: string}}
 */
function handleTrigger(intervalMinutes = 60) {
  try {
    Logger.log(`[TriggerService] Requested: Create/Update trigger for every ${intervalMinutes} minutes.`);

    // Remove any existing time-driven trigger so interval changes take effect.
    const existingTrigger = findActiveTrigger();
    if (existingTrigger) {
      ScriptApp.deleteTrigger(existingTrigger);
      Logger.log(`[TriggerService] Removed previous trigger to apply new interval.`);
    }

    // Create new trigger
    const result = _createTrigger(intervalMinutes);
    
    if (result.success) {
      Logger.log(`[TriggerService] New trigger created successfully.`);
    } else {
      Logger.log(`[TriggerService] Failed to create trigger: ${result.message}`);
    }

    return result;

  } catch (error) {
    Logger.log(`[TriggerService] Critical Error in handleTrigger: ${error.toString()}`);
    return {
      success: false,
      message: `Error: ${error.message}`
    };
  }
}

/**
 * Removes all time-driven triggers associated with DocuForm Sync.
 * @returns {{success: boolean, message: string}}
 */
function removeTrigger() {
  try {
    Logger.log(`[TriggerService] Requested: Delete all DocuForm Sync triggers.`);
    
    const deletedCount = _deleteTriggers();

    if (deletedCount > 0) {
      return {
        success: true,
        message: `Removed ${deletedCount} trigger(s).`
      };
    } else {
      return {
        success: true,
        message: "No DocuForm Sync triggers found to remove."
      };
    }

  } catch (error) {
    Logger.log(`[TriggerService] Critical Error in removeTrigger: ${error.toString()}`);
    return {
      success: false,
      message: `Deletion Failed: ${error.message}`
    };
  }
}

/**
 * Checks the current status of the trigger system.
 * @returns {{exists: boolean, handler: string|null, frequency: string|null}}
 */
function getTriggerStatus() {
  const trigger = findActiveTrigger();
  
  if (trigger) {
    return {
      exists: true,
      handler: trigger.getHandlerFunction(),
      frequency: getFrequencyDescription(trigger)
    };
  }

  return {
    exists: false,
    handler: null,
    frequency: null
  };
}

// ============================ INTERNAL HELPERS ============================

/**
 * Finds an active time-driven trigger that matches our handler function.
 * @returns {ScriptApp.Trigger|null}
 */
function findActiveTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === TRIGGER_HANDLER && 
        triggers[i].getEventType() === ScriptApp.EventType.TIME_DRIVEN) {
      return triggers[i];
    }
  }
  
  return null;
}

/**
 * Creates a new time-driven trigger.
 * 
 * @param {number} intervalMinutes - Minutes between executions.
 * @returns {{success: boolean, message: string}}
 */
function _createTrigger(intervalMinutes) {
  try {
    if (typeof intervalMinutes !== 'number' || intervalMinutes < 1) {
      return { success: false, message: "Invalid interval provided." };
    }

    // Add-ons can only schedule a time-driven trigger once per hour at most.
    // Clamp sub-hour values so legacy configs don't fail to create a trigger.
    if (intervalMinutes < 60) intervalMinutes = 60;

    const triggerBuilder = ScriptApp.newTrigger(TRIGGER_HANDLER).timeBased();

    // Add-ons can only schedule a time-driven trigger once per hour at most, so
    // valid intervals are 1h / 6h / 12h / 24h.
    if (intervalMinutes >= 1440) {
      triggerBuilder.everyDays(1);
    } else if (intervalMinutes >= 720) {
      triggerBuilder.everyHours(12);
    } else if (intervalMinutes >= 360) {
      triggerBuilder.everyHours(6);
    } else {
      triggerBuilder.everyHours(1);
    }

    triggerBuilder.create();

    return { success: true, message: "Trigger Created" };

  } catch (e) {
    Logger.log(`[TriggerService] Failed to create trigger: ${e.toString()}`);
    return { success: false, message: `Creation Error: ${e.message}` };
  }
}

/**
 * Deletes all triggers associated with our handler function.
 * @returns {number} Number of deleted triggers.
 */
function _deleteTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  let count = 0;

  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === TRIGGER_HANDLER) {
      ScriptApp.deleteTrigger(triggers[i]);
      count++;
    }
  }
  
  return count;
}

/**
 * Helper to convert a trigger to a readable frequency string.
 * @param {ScriptApp.Trigger} trigger
 * @returns {string}
 */
function getFrequencyDescription(trigger) {
  const source = trigger.getTriggerSource();
  const eventType = trigger.getEventType();
  
  if (source === ScriptApp.TriggerSource.CLOCK && eventType === ScriptApp.EventType.TIME_DRIVEN) {
    return "Time-driven";
  }
  
  return "Unknown";
}
