/**
 * Smart Home Inventory - Shared Configuration
 * 
 * Contains configuration constants and column mappings shared across all scripts.
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
  "Notes": 17,
  "Status": 18
};

// =============================================================================
// GENERAL CONFIGURATION
// =============================================================================

const CONFIG = {
  SHEET_NAME: "Items",
  HEADER_ROW: 1,
  DATA_START_ROW: 2
};

// =============================================================================
// FIELDS TO AUTO-FILL
// =============================================================================
// Fields that should be auto-filled from the database when a model is entered

const FILLABLE_FIELDS = [
  "Device type",
  "Manufacturer",
  "Actual device name",
  "Zigbee ID",
  "Luminous flux",
  "Power",
  "Plug type",
  "Smart"
];

// =============================================================================
// UTILITY FUNCTIONS
// =============================================================================

/**
 * Gets the column letter for a given field name.
 * 
 * @param {string} fieldName - The name of the field
 * @return {string} The column letter (e.g., "A", "B", "AA")
 */
function getColumnLetter(fieldName) {
  const colNum = COLUMN_MAP[fieldName];
  if (!colNum) {
    throw new Error(`Field "${fieldName}" not found in COLUMN_MAP`);
  }
  
  let letter = '';
  let num = colNum;
  
  while (num > 0) {
    const remainder = (num - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    num = Math.floor((num - 1) / 26);
  }
  
  return letter;
}

/**
 * Gets cell reference (e.g., "A5") for a given field and row.
 * 
 * @param {string} fieldName - The name of the field
 * @param {number} row - The row number
 * @return {string} The cell reference (e.g., "B10")
 */
function getCellReference(fieldName, row) {
  return getColumnLetter(fieldName) + row;
}
