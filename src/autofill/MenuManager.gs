/**
 * Smart Home Inventory - Menu Manager
 * 
 * Manages custom menu items for the spreadsheet.
 * 
 * @author mariobob
 * @version 1.0
 */

/**
 * Creates custom menu when the spreadsheet is opened.
 * This function is automatically triggered by Apps Script.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('🏠 Smart Home Tools')
    .addSubMenu(ui.createMenu('✅ Validation')
      .addItem('Validate All Items', 'validateAllItems')
      .addItem('Fix All Discrepancies', 'fixAllDiscrepancies')
      .addSeparator()
      .addItem('Show Database Stats', 'showDatabaseStats'))
    .addSeparator()
    .addItem('🔀 Sort Items by Room', 'sortItemsWithinRooms')
    .addItem('➕ Add Model to Database', 'showAddModelDialog')
    .addSeparator()
    .addSubMenu(ui.createMenu('⚙️ Setup')
      .addItem('Enable dropdown auto-add', 'installDropdownAutoAddTrigger')
      .addItem('Disable dropdown auto-add', 'disableDropdownAutoAddTrigger'))
    .addSubMenu(ui.createMenu('📷 Photo Intake')
      .addItem('Set up photo intake', 'setupPhotoIntake')
      .addItem('Process inbox now', 'processPhotoInbox')
      .addItem('Turn off auto-intake', 'disablePhotoIntake'))
    .addSubMenu(ui.createMenu('ℹ️ Help')
      .addItem('About Auto-Fill', 'showAutoFillHelp')
      .addItem('About Validation', 'showValidationHelp'))
    .addToUi();
}

/**
 * Shows information about the auto-fill feature.
 */
function showAutoFillHelp() {
  const ui = SpreadsheetApp.getUi();
  
  const helpText = `
AUTO-FILL FEATURE
================

When you enter a Model number in the "Model" column, the following fields 
will be automatically filled from the device database:

• Device type
• Manufacturer
• Actual device name
• Zigbee ID
• Luminous flux
• Power
• Plug type
• Smart

HOW TO USE:
1. Enter or paste a Model number in the Model column (column H)
2. Press Enter or click outside the cell
3. Watch as the fields auto-populate!

NOTES:
• Only empty fields or fields with "-" will be filled
• If the model is not found, you'll see a notification
• The auto-fill happens instantly (no need to wait)

DATABASE SIZE: ${Object.keys(DEVICE_MAP).length} models
  `;
  
  ui.alert('Auto-Fill Help', helpText, ui.ButtonSet.OK);
}

/**
 * Shows information about the validation feature.
 */
function showValidationHelp() {
  const ui = SpreadsheetApp.getUi();
  
  const helpText = `
VALIDATION FEATURE
==================

The validation tool checks all items in your inventory against the 
device database and identifies discrepancies.

WHAT IT CHECKS:
• Compares actual values with expected database values
• Identifies models not in the database
• Finds mismatched device information

HOW TO USE:
1. Go to: 🏠 Smart Home Tools → ✅ Validation → Validate All Items
2. Wait for the validation to complete
3. Review the detailed HTML report that appears
4. (Optional) Use "Fix All Discrepancies" to auto-correct issues

VALIDATION REPORT INCLUDES:
• Summary statistics
• List of all discrepancies by row
• Cell references for easy navigation
• Expected vs actual values comparison

TIP: The full report is also logged to the script logs 
(View → Logs in the script editor).
  `;
  
  ui.alert('Validation Help', helpText, ui.ButtonSet.OK);
}

/**
 * Shows statistics about the device database.
 */
function showDatabaseStats() {
  const stats = getDeviceDatabaseStats();
  
  let message = `DEVICE DATABASE STATISTICS\n`;
  message += `==========================\n\n`;
  message += `Total Models: ${stats.totalModels}\n\n`;
  
  message += `DEVICE TYPES:\n`;
  Object.keys(stats.deviceTypes).sort().forEach(type => {
    message += `  • ${type}: ${stats.deviceTypes[type]}\n`;
  });
  
  message += `\nMANUFACTURERS:\n`;
  Object.keys(stats.manufacturers).sort().forEach(mfr => {
    message += `  • ${mfr}: ${stats.manufacturers[mfr]}\n`;
  });
  
  const ui = SpreadsheetApp.getUi();
  ui.alert('Database Statistics', message, ui.ButtonSet.OK);
}
