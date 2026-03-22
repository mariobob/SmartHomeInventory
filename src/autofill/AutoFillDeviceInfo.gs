/**
 * Smart Home Inventory - Auto-Fill Device Information
 * 
 * This script automatically populates device fields when a Model is entered.
 * It uses a predefined map of models to their associated device information.
 * 
 * @author mariobob
 * @version 1.0
 */

// =============================================================================
// MAIN TRIGGER FUNCTION
// =============================================================================

/**
 * Triggered when a cell is edited in the spreadsheet.
 * Automatically fills device information when Model is updated.
 * 
 * @param {Object} e - The edit event object
 */
function onEdit(e) {
  try {
    // Get event details
    const range = e.range;
    const sheet = range.getSheet();
    const row = range.getRow();
    const col = range.getColumn();
    
    // Only process edits in the Items sheet
    if (sheet.getName() !== CONFIG.SHEET_NAME) {
      return;
    }
    
    // Only process edits in data rows (not header)
    if (row < CONFIG.DATA_START_ROW) {
      return;
    }
    
    // Only process edits to the Model column
    if (col !== COLUMN_MAP["Model"]) {
      return;
    }
    
    // Get the entered model value
    const modelValue = range.getValue();
    
    // If model is empty, do nothing
    if (!modelValue || modelValue === "" || modelValue === "-") {
      return;
    }
    
    // Look up the model in the device map (including custom devices)
    const deviceInfo = getDeviceInfo(modelValue);
    
    // If model not found, highlight the cell and leave a note with instructions
    if (!deviceInfo) {
      range.setBackground('#FFF3CD');
      range.setNote(
        '⚠️ Unknown model: "' + modelValue + '"\n' +
        'Use: 🏠 Smart Home Tools → ➕ Add Model to Database'
      );
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Model "' + modelValue + '" not found — cell highlighted yellow.\n' +
        'Open Smart Home Tools → ➕ Add Model to Database to add it.',
        'Unknown Model',
        7
      );
      return;
    }

    // Clear any previous unknown-model highlight when model IS found
    range.setBackground(null);
    range.clearNote();

    // Auto-fill the device information
    fillDeviceInformation(sheet, row, deviceInfo);

    // Show success notification
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Device information auto-filled for model "${modelValue}"`,
      "Success",
      3
    );
    
  } catch (error) {
    // Log error for debugging
    Logger.log(`Error in onEdit: ${error.message}`);
    
    // Show user-friendly error message
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `An error occurred: ${error.message}`,
      "Error",
      5
    );
  }
}

// =============================================================================
// HELPER FUNCTIONS
// =============================================================================

/**
 * Fills device information in the specified row.
 * 
 * @param {Sheet} sheet - The sheet object
 * @param {number} row - The row number to fill
 * @param {Object} deviceInfo - The device information object
 */
function fillDeviceInformation(sheet, row, deviceInfo) {
  // Use shared FILLABLE_FIELDS from Config.gs
  const fieldsToFill = FILLABLE_FIELDS;
  
  // Prepare batch update
  const updates = [];
  
  fieldsToFill.forEach(field => {
    const col = COLUMN_MAP[field];
    const currentValue = sheet.getRange(row, col).getValue();
    const newValue = deviceInfo[field] || "";
    
    // Only update if:
    // 1. Current value is empty, OR
    // 2. Current value is "-" (placeholder)
    if (currentValue === "" || currentValue === "-") {
      updates.push({
        row: row,
        col: col,
        value: newValue
      });
    }
  });
  
  // Apply all updates at once (more efficient)
  updates.forEach(update => {
    sheet.getRange(update.row, update.col).setValue(update.value);
  });
}
