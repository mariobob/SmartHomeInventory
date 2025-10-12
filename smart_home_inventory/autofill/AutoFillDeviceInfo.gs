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
// COLUMN MAPPING
// =============================================================================
// Maps field names to their column indices (1-based)

const COLUMN_MAP = {
  "Room": 1,
  "Device type": 2,
  "Manufacturer": 3,
  "Nickname": 4,
  "Actual device name": 5,
  "Mac address": 6,
  "Zigbee ID": 7,
  "Model": 8,
  "Serial number": 9,
  "Luminous flux": 10,
  "Power": 11,
  "Plug type": 12,
  "Row done?": 13,
  "Smart": 14,
  "Installed?": 15,
  "Old name": 16,
  "Notes": 17
};

// =============================================================================
// CONFIGURATION
// =============================================================================

const CONFIG = {
  SHEET_NAME: "Items",
  HEADER_ROW: 1,
  DATA_START_ROW: 2
};

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
    
    // Look up the model in the device map
    const deviceInfo = getDeviceInfo(modelValue);
    
    // If model not found, show a gentle notification
    if (!deviceInfo) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `Model "${modelValue}" not found in database. Please check the model number.`,
        "Model Not Found",
        5
      );
      return;
    }
    
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
  const fieldsToFill = [
    "Device type",
    "Manufacturer",
    "Actual device name",
    "Zigbee ID",
    "Luminous flux",
    "Power",
    "Plug type",
    "Smart"
  ];
  
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
