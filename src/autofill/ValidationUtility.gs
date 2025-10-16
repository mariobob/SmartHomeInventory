/**
 * Smart Home Inventory - Validation Utility
 * 
 * Validates device information in the inventory against the device database.
 * Identifies discrepancies and provides detailed reports.
 * 
 * @author mariobob
 * @version 1.0
 */

// =============================================================================
// VALIDATION FUNCTIONS
// =============================================================================

/**
 * Validates all items in the inventory against the device database.
 * Can be triggered from the custom menu.
 */
function validateAllItems() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  
  if (!sheet) {
    showAlert("Error", `Sheet "${CONFIG.SHEET_NAME}" not found.`);
    return;
  }
  
  // Show progress notification
  ss.toast("Starting validation...", "Validating Inventory", -1);
  
  // Get all data from the sheet
  const lastRow = sheet.getLastRow();
  
  if (lastRow < CONFIG.DATA_START_ROW) {
    showAlert("No Data", "No items found to validate.");
    return;
  }
  
  // Collect discrepancies
  const discrepancies = [];
  let totalItems = 0;
  let itemsWithModel = 0;
  let itemsValidated = 0;
  let itemsWithErrors = 0;
  
  // Iterate through all data rows
  for (let row = CONFIG.DATA_START_ROW; row <= lastRow; row++) {
    totalItems++;
    
    // Get the model value
    const modelCell = sheet.getRange(row, COLUMN_MAP["Model"]);
    const modelValue = modelCell.getValue();
    
    // Skip if no model or placeholder
    if (!modelValue || modelValue === "" || modelValue === "-") {
      continue;
    }
    
    itemsWithModel++;
    
    // Look up the model in the database
    const expectedInfo = getDeviceInfo(modelValue);
    
    // If model not in database, skip it (not an error, just not in database yet)
    if (!expectedInfo) {
      continue;
    }
    
    itemsValidated++;
    
    // Check each fillable field
    const rowDiscrepancies = validateRow(sheet, row, modelValue, expectedInfo);
    
    if (rowDiscrepancies.length > 0) {
      discrepancies.push(...rowDiscrepancies);
      itemsWithErrors++;
    }
  }
  
  // Generate and display report
  displayValidationReport(discrepancies, {
    totalItems: totalItems,
    itemsWithModel: itemsWithModel,
    itemsValidated: itemsValidated,
    itemsWithErrors: itemsWithErrors
  });
}

/**
 * Validates a single row against expected device information.
 * 
 * @param {Sheet} sheet - The sheet object
 * @param {number} row - The row number
 * @param {string} modelValue - The model number
 * @param {Object} expectedInfo - Expected device information from database
 * @return {Array} Array of discrepancies found
 */
function validateRow(sheet, row, modelValue, expectedInfo) {
  const discrepancies = [];
  
  // Check each fillable field
  FILLABLE_FIELDS.forEach(field => {
    const col = COLUMN_MAP[field];
    const actualValue = sheet.getRange(row, col).getValue();
    const expectedValue = expectedInfo[field] || "";
    
    // Skip validation if:
    // 1. Both are empty
    // 2. Actual is empty or "-" (not yet filled)
    if (actualValue === "" || actualValue === "-") {
      // This is okay - field not filled yet
      return;
    }
    
    // Convert to strings for comparison
    const actualStr = String(actualValue).trim();
    const expectedStr = String(expectedValue).trim();
    
    // Check if values match
    if (actualStr !== expectedStr && expectedStr !== "") {
      discrepancies.push({
        row: row,
        model: modelValue,
        field: field,
        issue: "Value mismatch",
        expected: expectedStr,
        actual: actualStr,
        cellRef: getCellReference(field, row)
      });
    }
  });
  
  return discrepancies;
}

/**
 * Displays the validation report to the user.
 * 
 * @param {Array} discrepancies - Array of discrepancy objects
 * @param {Object} stats - Validation statistics
 */
function displayValidationReport(discrepancies, stats) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Hide progress toast
  ss.toast("Validation complete!", "Status", 1);
  
  if (discrepancies.length === 0) {
    showAlert(
      "Validation Complete ✓",
      `All items validated successfully!\n\n` +
      `Total items: ${stats.totalItems}\n` +
      `Items with model: ${stats.itemsWithModel}\n` +
      `Items validated: ${stats.itemsValidated}\n\n` +
      `No discrepancies found.`
    );
    return;
  }
  
  // Generate detailed report
  const reportLines = [];
  reportLines.push(`VALIDATION REPORT`);
  reportLines.push(`================\n`);
  reportLines.push(`Total items: ${stats.totalItems}`);
  reportLines.push(`Items with model: ${stats.itemsWithModel}`);
  reportLines.push(`Items validated: ${stats.itemsValidated}`);
  reportLines.push(`Items with errors: ${stats.itemsWithErrors}`);
  reportLines.push(`Total discrepancies: ${discrepancies.length}\n`);
  reportLines.push(`DISCREPANCIES FOUND:`);
  reportLines.push(`-------------------\n`);
  
  // Group by row for better readability
  const discrepanciesByRow = {};
  discrepancies.forEach(d => {
    if (!discrepanciesByRow[d.row]) {
      discrepanciesByRow[d.row] = [];
    }
    discrepanciesByRow[d.row].push(d);
  });
  
  // Format each row's discrepancies
  Object.keys(discrepanciesByRow).sort((a, b) => parseInt(a) - parseInt(b)).forEach(row => {
    const items = discrepanciesByRow[row];
    const model = items[0].model;
    
    reportLines.push(`Row ${row} (Model: ${model}):`);
    
    items.forEach(item => {
      reportLines.push(`  • ${item.field} [${item.cellRef}]`);
      reportLines.push(`    Issue: ${item.issue}`);
      reportLines.push(`    Expected: "${item.expected}"`);
      reportLines.push(`    Actual: "${item.actual}"`);
    });
    
    reportLines.push('');
  });
  
  // Log to console for detailed review
  const fullReport = reportLines.join('\n');
  Logger.log(fullReport);
  
  // Show summary in dialog
  const htmlOutput = HtmlService.createHtmlOutput(generateHtmlReport(discrepanciesByRow, stats))
    .setWidth(800)
    .setHeight(600);
  
  ss.show(htmlOutput);
}

/**
 * Generates an HTML report for display in a dialog.
 * 
 * @param {Object} discrepanciesByRow - Discrepancies grouped by row
 * @param {Object} stats - Validation statistics
 * @return {string} HTML content
 */
function generateHtmlReport(discrepanciesByRow, stats) {
  const totalDiscrepancies = Object.values(discrepanciesByRow).reduce((sum, arr) => sum + arr.length, 0);
  
  let html = `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body {
      font-family: 'Google Sans', Roboto, Arial, sans-serif;
      padding: 20px;
      background-color: #f8f9fa;
      margin: 0;
    }
    .header {
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      color: white;
      padding: 20px;
      border-radius: 8px;
      margin-bottom: 20px;
    }
    .header h1 {
      margin: 0 0 10px 0;
      font-size: 24px;
    }
    .stats {
      display: grid;
      grid-template-columns: repeat(2, 1fr);
      gap: 10px;
      margin-bottom: 20px;
    }
    .stat-card {
      background: white;
      padding: 15px;
      border-radius: 8px;
      box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }
    .stat-label {
      font-size: 12px;
      color: #666;
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }
    .stat-value {
      font-size: 28px;
      font-weight: bold;
      color: #333;
      margin-top: 5px;
    }
    .stat-value.error {
      color: #dc3545;
    }
    .stat-value.success {
      color: #28a745;
    }
    .discrepancies {
      background: white;
      border-radius: 8px;
      padding: 20px;
      box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }
    .row-group {
      margin-bottom: 20px;
      border-left: 4px solid #dc3545;
      padding-left: 15px;
    }
    .row-header {
      font-weight: bold;
      color: #333;
      margin-bottom: 10px;
      font-size: 16px;
    }
    .discrepancy-item {
      margin-bottom: 15px;
      padding: 12px;
      background: #f8f9fa;
      border-radius: 6px;
    }
    .field-name {
      font-weight: bold;
      color: #667eea;
      display: inline-block;
      min-width: 150px;
    }
    .cell-ref {
      color: #666;
      font-size: 12px;
      margin-left: 10px;
      font-family: 'Courier New', monospace;
      background: #e9ecef;
      padding: 2px 6px;
      border-radius: 3px;
    }
    .value-comparison {
      margin-top: 8px;
      display: grid;
      grid-template-columns: 80px 1fr;
      gap: 5px;
      font-size: 14px;
    }
    .label {
      color: #666;
      font-weight: 500;
    }
    .expected {
      color: #28a745;
      font-family: 'Courier New', monospace;
    }
    .actual {
      color: #dc3545;
      font-family: 'Courier New', monospace;
    }
    .no-issues {
      text-align: center;
      padding: 40px;
      color: #28a745;
      font-size: 18px;
    }
    .footer {
      margin-top: 20px;
      padding: 15px;
      background: #e9ecef;
      border-radius: 8px;
      font-size: 12px;
      color: #666;
    }
  </style>
</head>
<body>
  <div class="header">
    <h1>📋 Inventory Validation Report</h1>
    <div>Complete analysis of device information against database</div>
  </div>
  
  <div class="stats">
    <div class="stat-card">
      <div class="stat-label">Total Items</div>
      <div class="stat-value">${stats.totalItems}</div>
    </div>
    <div class="stat-card">
      <div class="stat-label">Items with Model</div>
      <div class="stat-value">${stats.itemsWithModel}</div>
    </div>
    <div class="stat-card">
      <div class="stat-label">Items Validated</div>
      <div class="stat-value success">${stats.itemsValidated}</div>
    </div>
    <div class="stat-card">
      <div class="stat-label">Discrepancies Found</div>
      <div class="stat-value error">${totalDiscrepancies}</div>
    </div>
  </div>
  
  <div class="discrepancies">
    <h2 style="margin-top: 0;">Discrepancies Detail</h2>
`;

  const rows = Object.keys(discrepanciesByRow).sort((a, b) => parseInt(a) - parseInt(b));
  
  if (rows.length === 0) {
    html += `<div class="no-issues">✓ No issues found! All validated items match the database.</div>`;
  } else {
    rows.forEach(row => {
      const items = discrepanciesByRow[row];
      const model = items[0].model;
      
      html += `
      <div class="row-group">
        <div class="row-header">Row ${row} • Model: ${model}</div>
      `;
      
      items.forEach(item => {
        html += `
        <div class="discrepancy-item">
          <div>
            <span class="field-name">${item.field}</span>
            <span class="cell-ref">${item.cellRef}</span>
          </div>
          <div class="value-comparison">
            <span class="label">Expected:</span>
            <span class="expected">${escapeHtml(item.expected)}</span>
            <span class="label">Actual:</span>
            <span class="actual">${escapeHtml(item.actual)}</span>
          </div>
        </div>
        `;
      });
      
      html += `</div>`;
    });
  }
  
  html += `
  </div>
  
  <div class="footer">
    💡 <strong>Tip:</strong> You can fix these issues by re-entering the Model number in each row, 
    or manually updating the incorrect values. Check the Apps Script logs (View → Logs) for a text version of this report.
  </div>
</body>
</html>
  `;
  
  return html;
}

/**
 * Escapes HTML special characters.
 * 
 * @param {string} text - Text to escape
 * @return {string} Escaped text
 */
function escapeHtml(text) {
  const map = {
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#039;'
  };
  return String(text).replace(/[&<>"']/g, m => map[m]);
}

/**
 * Shows a simple alert dialog.
 * 
 * @param {string} title - Dialog title
 * @param {string} message - Dialog message
 */
function showAlert(title, message) {
  const ui = SpreadsheetApp.getUi();
  ui.alert(title, message, ui.ButtonSet.OK);
}

// =============================================================================
// FIX UTILITIES
// =============================================================================

/**
 * Fixes all discrepancies by updating cells with database values.
 * WARNING: This will overwrite existing data!
 */
function fixAllDiscrepancies() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Fix All Discrepancies',
    'This will overwrite all incorrect values with database values.\n\n' +
    'Are you sure you want to continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  
  if (!sheet) {
    showAlert("Error", `Sheet "${CONFIG.SHEET_NAME}" not found.`);
    return;
  }
  
  ss.toast("Fixing discrepancies...", "Working", -1);
  
  const lastRow = sheet.getLastRow();
  let fixedCount = 0;
  
  for (let row = CONFIG.DATA_START_ROW; row <= lastRow; row++) {
    const modelValue = sheet.getRange(row, COLUMN_MAP["Model"]).getValue();
    
    if (!modelValue || modelValue === "" || modelValue === "-") {
      continue;
    }
    
    const deviceInfo = getDeviceInfo(modelValue);
    
    if (deviceInfo) {
      // Use the existing fillDeviceInformation function but force update
      FILLABLE_FIELDS.forEach(field => {
        const col = COLUMN_MAP[field];
        const expectedValue = deviceInfo[field] || "";
        const actualValue = sheet.getRange(row, col).getValue();
        
        if (String(actualValue).trim() !== String(expectedValue).trim() && expectedValue !== "") {
          sheet.getRange(row, col).setValue(expectedValue);
          fixedCount++;
        }
      });
    }
  }
  
  ss.toast(`Fixed ${fixedCount} field(s)`, "Complete", 3);
  showAlert("Fix Complete", `Successfully fixed ${fixedCount} discrepancies.`);
}
