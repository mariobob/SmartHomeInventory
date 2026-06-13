/**
 * Smart Home Inventory - Sort Utility
 *
 * Sorts items within each room by Device type → Manufacturer → Nickname,
 * using the ordering defined in each column's data-validation dropdown.
 * Preserves the physical room order (first-occurrence order).
 * Empty-room rows (status devices after migration) are always sorted last.
 */

// =============================================================================
// SORT
// =============================================================================

function sortItemsWithinRooms() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  const firstDataRow = CONFIG.DATA_START_ROW;
  const lastRow      = sheet.getLastRow();
  const lastCol      = sheet.getLastColumn();

  if (lastRow < firstDataRow) {
    ss.toast('No data rows found.', 'Sort');
    return;
  }

  // Read ordering from each column's validation dropdown
  const deviceTypeOrder   = getValidationListOrder(sheet, COLUMN_MAP['Device type']);
  const manufacturerOrder = getValidationListOrder(sheet, COLUMN_MAP['Manufacturer']);

  const range       = sheet.getRange(firstDataRow, 1, lastRow - firstDataRow + 1, lastCol);
  const values      = range.getValues();
  const backgrounds = range.getBackgrounds();
  const fontColors  = range.getFontColors();
  const fontWeights = range.getFontWeights();
  const notes       = range.getNotes();

  // 0-based column indices
  const ROOM_IDX  = COLUMN_MAP['Room']         - 1;
  const TYPE_IDX  = COLUMN_MAP['Device type']  - 1;
  const MFR_IDX   = COLUMN_MAP['Manufacturer'] - 1;
  const NICK_IDX  = COLUMN_MAP['Nickname']     - 1;

  const rows = values.map((v, i) => ({
    values:       v,
    backgrounds:  backgrounds[i],
    fontColors:   fontColors[i],
    fontWeights:  fontWeights[i],
    notes:        notes[i],
    room:         String(v[ROOM_IDX]  || ''),
    deviceType:   String(v[TYPE_IDX]  || ''),
    manufacturer: String(v[MFR_IDX]  || ''),
    nickname:     String(v[NICK_IDX]  || '')
  }));

  // Preserve room order by first occurrence; empty Room → always last
  const roomOrder = [];
  const roomSeen  = {};
  rows.forEach(r => {
    if (r.room && !roomSeen[r.room]) {
      roomSeen[r.room] = true;
      roomOrder.push(r.room);
    }
  });
  const roomRank = {};
  roomOrder.forEach((r, i) => { roomRank[r] = i; });

  rows.sort((a, b) => {
    // 1. Room (empty → last)
    const rA = roomRank[a.room] ?? 99999;
    const rB = roomRank[b.room] ?? 99999;
    if (rA !== rB) return rA - rB;

    // 2. Device type by validation-rule order; unknowns sort after known, then alpha
    const tA = deviceTypeOrder[a.deviceType]  ?? 99999;
    const tB = deviceTypeOrder[b.deviceType]  ?? 99999;
    if (tA !== tB) return tA - tB;
    if (tA === 99999) {
      const fb = a.deviceType.localeCompare(b.deviceType);
      if (fb !== 0) return fb;
    }

    // 3. Manufacturer by validation-rule order; unknowns sort after known, then alpha
    const mA = manufacturerOrder[a.manufacturer] ?? 99999;
    const mB = manufacturerOrder[b.manufacturer] ?? 99999;
    if (mA !== mB) return mA - mB;
    if (mA === 99999) {
      const fb = a.manufacturer.localeCompare(b.manufacturer);
      if (fb !== 0) return fb;
    }

    // 4. Nickname alphabetically
    return a.nickname.localeCompare(b.nickname);
  });

  // Write back values + formatting
  range.setValues(rows.map(r => r.values));
  range.setBackgrounds(rows.map(r => r.backgrounds));
  range.setFontColors(rows.map(r => r.fontColors));
  range.setFontWeights(rows.map(r => r.fontWeights));
  range.setNotes(rows.map(r => r.notes));

  ss.toast(
    'Sorted ' + rows.length + ' rows by room → device type → manufacturer → nickname',
    'Sort Complete'
  );
}

/**
 * Returns a {value: index} map built from a column's VALUE_IN_LIST validation.
 * Returns {} if the column has no dropdown validation.
 *
 * @param {Sheet}  sheet  The sheet to inspect.
 * @param {number} col    1-based column index.
 * @return {Object}       Map of string value → sort rank (0 = first in dropdown).
 */
function getValidationListOrder(sheet, col) {
  try {
    const validation = sheet.getRange(2, col).getDataValidation();
    if (!validation) return {};
    const type   = validation.getCriteriaType();
    const values = validation.getCriteriaValues();

    if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
      const list = values[0];
      if (!Array.isArray(list)) return {};
      const order = {};
      list.forEach((v, i) => { order[String(v)] = i; });
      return order;
    }

    if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
      const range      = values[0];           // Range object
      const rangeVals  = range.getValues();   // 2-D array
      const order = {};
      let i = 0;
      rangeVals.forEach(row => row.forEach(cell => {
        if (cell !== '') { order[String(cell)] = i++; }
      }));
      return order;
    }

    return {};
  } catch (e) {
    Logger.log('getValidationListOrder error (col ' + col + '): ' + e.message);
    return {};
  }
}

// =============================================================================
// ONE-TIME RESTORATION
// =============================================================================

/**
 * Restores the Items sheet from the 2026-06-13 backup spreadsheet.
 * Run ONCE to undo the sort-induced data corruption.
 * Automatically rebuilds the Status column after restoring.
 *
 * Safe to run even if Status column already exists — clears everything first.
 */
function restoreFromBackup() {
  const BACKUP_ID   = '1_hMZyAp9UARxEdG7qO2r1kuDYx53OiQ29c4ZdVKBTgQ';
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  // Open backup
  ss.toast('Opening backup spreadsheet…', 'Restore', 10);
  let backupSS, backupSheet;
  try {
    backupSS    = SpreadsheetApp.openById(BACKUP_ID);
    backupSheet = backupSS.getSheetByName(CONFIG.SHEET_NAME);
  } catch (e) {
    ss.toast('ERROR: Cannot open backup — ' + e.message, 'Restore Failed', 15);
    return;
  }
  if (!backupSheet) {
    ss.toast('ERROR: Items sheet not found in backup.', 'Restore Failed', 10);
    return;
  }

  const srcLastRow = backupSheet.getLastRow();
  const srcLastCol = backupSheet.getLastColumn();   // 17 (no Status col in backup)

  // Read values + cell notes from backup
  ss.toast('Reading ' + (srcLastRow - 1) + ' rows from backup…', 'Restore', 30);
  const srcRange  = backupSheet.getRange(1, 1, srcLastRow, srcLastCol);
  const data      = srcRange.getValues();
  const cellNotes = srcRange.getNotes();

  // Clear all content and cell notes (cols 1–18) in the target sheet
  const dstLastRow = Math.max(targetSheet.getLastRow(), srcLastRow);
  const clearRange = targetSheet.getRange(1, 1, dstLastRow, 18);
  clearRange.clearContent();
  clearRange.setNotes(
    Array.from({ length: dstLastRow }, () => Array(18).fill(''))
  );

  // Write backup data (17 cols) and notes
  ss.toast('Writing backup data…', 'Restore', 20);
  targetSheet.getRange(1, 1, srcLastRow, srcLastCol).setValues(data);
  targetSheet.getRange(1, 1, srcLastRow, srcLastCol).setNotes(cellNotes);

  // Rebuild Status column (col 18) from the clean backup data
  ss.toast('Rebuilding Status column…', 'Restore', 10);
  setupStatusColumn();

  ss.toast(
    '✅ ' + (srcLastRow - 1) + ' rows restored from backup. Status column rebuilt.',
    'Restore Complete', 12
  );
}

// =============================================================================
// ONE-TIME MIGRATION: setupStatusColumn
// =============================================================================

/**
 * Adds / rebuilds column 18 "Status".
 * Rows whose Room contains a status-like value (Extra, Unassigned, For sale,
 * Gone, Malfunctioned, Old home) get that value as Status and Room is cleared.
 * All other rows get Status = "Active".
 *
 * Run after restoreFromBackup() — the backup has real room names,
 * so the migration always works correctly from a clean state.
 */
function setupStatusColumn() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  const STATUS_COL = COLUMN_MAP['Status'];  // 18
  const ROOM_COL   = COLUMN_MAP['Room'];    // 1

  const STATUS_ROOMS = new Set([
    'Extra', 'Unassigned', 'For sale', 'Gone', 'Malfunctioned', 'Old home'
  ]);

  const lastRow = sheet.getLastRow();

  // 1. Header
  sheet.getRange(1, STATUS_COL)
    .setValue('Status')
    .setFontWeight('bold')
    .setBackground('#d9ead3');

  // 2. Build and write status values
  const roomValues   = sheet.getRange(2, ROOM_COL, lastRow - 1, 1).getValues();
  const statusValues = roomValues.map(([room]) =>
    STATUS_ROOMS.has(room) ? [room] : ['Active']
  );
  sheet.getRange(2, STATUS_COL, lastRow - 1, 1).setValues(statusValues);

  // 3. Clear Room for status rows
  const updatedRooms = roomValues.map(([room]) =>
    STATUS_ROOMS.has(room) ? [''] : [room]
  );
  sheet.getRange(2, ROOM_COL, lastRow - 1, 1).setValues(updatedRooms);

  // 4. Summary
  const moved = statusValues.filter(([s]) => s !== 'Active').length;
  ss.toast(
    `Status column set.\n${moved} rows moved Room → Status.\n${lastRow - 1 - moved} rows set Active.`,
    'Setup Complete', 8
  );
}
// END MIGRATIONS
