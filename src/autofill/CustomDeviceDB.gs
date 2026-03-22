/**
 * Smart Home Inventory - Custom Device Database
 *
 * Allows users to add device models that are not in DeviceDatabase.gs.
 * Entries are stored in Script Properties and persist across sessions.
 *
 * getDeviceInfo() in DeviceDatabase.gs automatically checks custom devices
 * as a fallback after the main DEVICE_MAP, so auto-fill works for all entries.
 *
 * @author mariobob
 * @version 1.0
 */

const CUSTOM_DEVICES_KEY = 'CUSTOM_DEVICES';

// =============================================================================
// MENU-CALLABLE FUNCTIONS
// =============================================================================

/**
 * Opens the "Add Model to Database" sidebar.
 * Called from the Smart Home Tools menu.
 */
function showAddModelDialog() {
  const html = HtmlService.createHtmlOutputFromFile('AddModelDialog')
    .setTitle('Add Model to Database')
    .setWidth(340);
  SpreadsheetApp.getUi().showSidebar(html);
}

// =============================================================================
// FUNCTIONS CALLED BY THE DIALOG VIA google.script.run
// =============================================================================

/**
 * Returns the model from the active cell (if it is in the Model column)
 * plus whether it was flagged as unknown (yellow background).
 * Called by the dialog on load to pre-fill the Model field.
 *
 * @return {{ model: string, isUnknown: boolean }}
 */
function getCurrentModelFromSelection() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) return { model: '', isUnknown: false };
  const cell = sheet.getActiveCell();
  if (cell.getColumn() !== COLUMN_MAP['Model']) return { model: '', isUnknown: false };
  const model = String(cell.getValue() || '').trim();
  // '#fff3cd' is the yellow highlight set by onEdit for unknown models
  const isUnknown = cell.getBackground().toLowerCase() === '#fff3cd';
  return { model: model, isUnknown: isUnknown };
}

/**
 * Saves a new model entry to Script Properties.
 * Called by google.script.run from the dialog's Save button.
 *
 * @param {Object} formData
 *   { model, deviceType, manufacturer, actualName, zigbeeId, lumens, power, plugType, smart }
 * @return {string} A human-readable success message shown in the dialog.
 */
function saveModelToDatabase(formData) {
  if (!formData || !formData.model) throw new Error('Model is required');
  if (!formData.actualName) throw new Error('Actual device name is required');

  const entry = {
    'Device type':        formData.deviceType   || '',
    'Manufacturer':       formData.manufacturer  || '',
    'Actual device name': formData.actualName    || '',
    'Zigbee ID':          formData.zigbeeId      || '-',
    'Luminous flux':      formData.lumens        || 'N/A',
    'Power':              formData.power         || '',
    'Plug type':          formData.plugType      || '',
    'Smart':              formData.smart         || 'Yes'
  };

  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(CUSTOM_DEVICES_KEY);
  const custom = raw ? JSON.parse(raw) : {};
  custom[formData.model] = entry;
  props.setProperty(CUSTOM_DEVICES_KEY, JSON.stringify(custom));

  // If the active cell still contains this model, apply auto-fill immediately
  _applyFillAfterSave_(formData.model, entry);

  return '✅ "' + formData.model + '" saved. Auto-fill is now active for this model.';
}

/**
 * Returns all custom entries (used by the dialog to show a count / list).
 *
 * @return {Object} Map of model → device info
 */
function getCustomDeviceList() {
  const raw = PropertiesService.getScriptProperties().getProperty(CUSTOM_DEVICES_KEY);
  return raw ? JSON.parse(raw) : {};
}

// =============================================================================
// INTERNAL HELPERS
// =============================================================================

/**
 * After saving a model, immediately apply auto-fill to the active row
 * if it still holds that model value.
 */
function _applyFillAfterSave_(model, deviceInfo) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) return;
    const cell = sheet.getActiveCell();
    if (cell.getColumn() !== COLUMN_MAP['Model']) return;
    if (String(cell.getValue() || '').trim() !== model) return;
    const row = cell.getRow();
    // Clear the unknown-model highlight and note
    cell.setBackground(null).clearNote();
    // Fill the device fields
    fillDeviceInformation(sheet, row, deviceInfo);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Auto-fill applied for "' + model + '"',
      'Auto-Fill',
      3
    );
  } catch (e) {
    Logger.log('_applyFillAfterSave_ error: ' + e.message);
  }
}
