/**
 * Smart Home Inventory - Dropdown Auto-Add
 *
 * When you type a value into a cell whose column has a list dropdown (data
 * validation) and the value is not yet in that list, this offers to add it to
 * the dropdown automatically (Yes/No), so you never have to edit the validation
 * rule by hand.
 *
 * It runs from an INSTALLABLE on-edit trigger, because simple triggers cannot
 * show dialogs. Enable it once via:
 *   Smart Home Tools -> Setup -> Enable dropdown auto-add
 *
 * For the prompt to be able to fire, a dropdown column must ACCEPT (and flag)
 * values that are not in the list, instead of hard-rejecting them. Enabling the
 * feature switches any "reject input" list columns on the Items sheet to
 * "show warning" mode automatically.
 *
 * Works for both "List of items" and "List from a range" validations, on any
 * column of the Items sheet.
 */

/**
 * Installs the on-edit trigger that powers dropdown auto-add and relaxes any
 * reject-input dropdowns to show-warning. Run once from the Setup menu. The
 * first run asks you to authorize the script.
 */
function installDropdownAutoAddTrigger() {
  const ss = SpreadsheetApp.getActive();
  // Remove any existing copy so we never stack duplicate triggers.
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'onEditDropdownAutoAdd') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('onEditDropdownAutoAdd')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  const relaxedCols = relaxListColumnsToWarning_();

  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Dropdown auto-add enabled',
    'Done. When you type a value into a dropdown column that is not in the list ' +
    'yet, you will be asked whether to add it.\n\n' +
    (relaxedCols > 0
      ? (relaxedCols + ' dropdown column(s) were switched from "reject input" to ' +
         '"show warning" so new values can be typed in. An unlisted value now ' +
         'shows a small corner flag until you add it.')
      : 'Your dropdown columns already allowed typing new values.'),
    ui.ButtonSet.OK
  );
}

/**
 * Turns the feature off by deleting its trigger. (Validation mode is left as
 * "show warning"; set it back to "reject input" by hand if you prefer.)
 */
function disableDropdownAutoAddTrigger() {
  let removed = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'onEditDropdownAutoAdd') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Dropdown auto-add disabled',
    removed ? 'The auto-add prompt has been turned off.' : 'It was not enabled.',
    ui.ButtonSet.OK
  );
}

/**
 * Installable on-edit handler. Offers to add an unknown value to its column's
 * dropdown list.
 *
 * @param {Object} e - the edit event object
 */
function onEditDropdownAutoAdd(e) {
  try {
    if (!e || !e.range) return;
    const range = e.range;
    const sheet = range.getSheet();

    // Items sheet, a single data cell only.
    if (sheet.getName() !== CONFIG.SHEET_NAME) return;
    if (range.getNumRows() !== 1 || range.getNumColumns() !== 1) return;
    if (range.getRow() < CONFIG.DATA_START_ROW) return;

    // Must be a list-style dropdown.
    const dv = range.getDataValidation();
    if (!dv) return;
    const type = dv.getCriteriaType();
    const VIL = SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST;
    const VIR = SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE;
    if (type !== VIL && type !== VIR) return;

    // The value the user typed (works even if "reject input" cleared the cell).
    let val = (e.value !== undefined && e.value !== null) ? e.value : range.getValue();
    val = String(val == null ? '' : val).trim();
    if (val === '' || val === '-') return;

    // Current allowed values; if it is already valid, do nothing.
    const allowed = getAllowedValues_(dv, type, VIR);
    if (allowed.indexOf(val) !== -1) return;

    // Ask the user.
    const ui = SpreadsheetApp.getUi();
    const header = String(sheet.getRange(1, range.getColumn()).getValue() || 'this').trim();
    const resp = ui.alert(
      'Add to dropdown?',
      'The value "' + val + '" is not in the "' + header + '" dropdown.\n\nAdd it to the list?',
      ui.ButtonSet.YES_NO
    );
    if (resp !== ui.Button.YES) return;

    // Add it.
    if (type === VIL) {
      addToListOfItems_(sheet, range.getColumn(), dv, allowed, val);
    } else {
      addToSourceRange_(dv, val);
    }

    // Re-affirm the typed value (in case reject-input wiped it) and confirm.
    range.setValue(val);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Added "' + val + '" to the "' + header + '" dropdown.', 'Dropdown updated', 5
    );
  } catch (err) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Could not update dropdown: ' + err.message, 'Error', 6
    );
  }
}

/** Returns the array of currently-allowed values for a list/range rule. */
function getAllowedValues_(dv, type, VIR) {
  const args = dv.getCriteriaValues();
  if (type === VIR) {
    return args[0].getValues()
      .map(function (r) { return String(r[0] == null ? '' : r[0]).trim(); })
      .filter(function (s) { return s !== ''; });
  }
  return (args[0] || []).map(function (s) { return String(s == null ? '' : s).trim(); });
}

/** Appends a value to a "List of items" rule across the whole data column. */
function addToListOfItems_(sheet, col, dv, allowed, val) {
  const args = dv.getCriteriaValues();
  const showDropdown = (args[1] !== false);
  const list = allowed.slice();
  list.push(val);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(list, showDropdown)
    .setAllowInvalid(true)
    .build();
  const n = sheet.getLastRow() - CONFIG.DATA_START_ROW + 1;
  const rows = n > 0 ? n : 1;
  sheet.getRange(CONFIG.DATA_START_ROW, col, rows, 1).setDataValidation(rule);
}

/** Appends a value to the source range backing a "List from a range" rule. */
function addToSourceRange_(dv, val) {
  const rng = dv.getCriteriaValues()[0];
  const srcSheet = rng.getSheet();
  const startRow = rng.getRow();
  const startCol = rng.getColumn();
  const numRows = rng.getNumRows();
  const colVals = srcSheet.getRange(startRow, startCol, numRows, 1).getValues();
  let emptyIdx = -1;
  for (let i = 0; i < colVals.length; i++) {
    if (String(colVals[i][0] == null ? '' : colVals[i][0]).trim() === '') { emptyIdx = i; break; }
  }
  if (emptyIdx >= 0) {
    srcSheet.getRange(startRow + emptyIdx, startCol).setValue(val);
  } else {
    srcSheet.getRange(startRow + numRows, startCol).setValue(val);  // extend below the range
  }
}

/**
 * Switches any "reject input" list/range dropdowns on the Items data range to
 * "show warning" (allow invalid), preserving each rule's list and dropdown
 * setting. Returns the number of columns that were changed.
 */
function relaxListColumnsToWarning_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEET_NAME);
  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  const nRows = lastRow - CONFIG.DATA_START_ROW + 1;
  if (nRows < 1 || lastCol < 1) return 0;
  const VIL = SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST;
  const VIR = SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE;
  const rng = sheet.getRange(CONFIG.DATA_START_ROW, 1, nRows, lastCol);
  const grid = rng.getDataValidations();
  let cellsChanged = 0;
  const colsTouched = {};
  for (let r = 0; r < grid.length; r++) {
    for (let c = 0; c < grid[r].length; c++) {
      const d = grid[r][c];
      if (!d) continue;
      const t = d.getCriteriaType();
      if ((t === VIL || t === VIR) && !d.getAllowInvalid()) {
        grid[r][c] = d.copy().setAllowInvalid(true).build();
        cellsChanged++;
        colsTouched[c] = true;
      }
    }
  }
  if (cellsChanged > 0) rng.setDataValidations(grid);
  return Object.keys(colsTouched).length;
}
