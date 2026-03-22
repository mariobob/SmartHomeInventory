/**
 * Smart Home Inventory - Sort Utility
 *
 * Sorts items within each room by Device type → Nickname (A-Z),
 * preserving the physical room order (first-occurrence order).
 */

function sortItemsWithinRooms() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Items');
  const firstDataRow = 2;
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < firstDataRow) {
    SpreadsheetApp.getActiveSpreadsheet().toast('No data rows found.', 'Sort');
    return;
  }

  const range = sheet.getRange(firstDataRow, 1, lastRow - firstDataRow + 1, lastCol);
  const values = range.getValues();
  const backgrounds = range.getBackgrounds();
  const fontColors = range.getFontColors();
  const fontWeights = range.getFontWeights();
  const notes = range.getNotes();

  // Build row objects
  const rows = values.map((v, i) => ({
    values: v,
    backgrounds: backgrounds[i],
    fontColors: fontColors[i],
    fontWeights: fontWeights[i],
    notes: notes[i],
    room: v[0] || '',
    deviceType: v[1] || '',
    nickname: v[3] || ''
  }));

  // Preserve room order by first occurrence
  const roomOrder = [];
  const roomSeen = {};
  rows.forEach(r => {
    if (r.room && !roomSeen[r.room]) {
      roomSeen[r.room] = true;
      roomOrder.push(r.room);
    }
  });
  const roomRank = {};
  roomOrder.forEach((r, i) => { roomRank[r] = i; });

  rows.sort((a, b) => {
    const rDiff = (roomRank[a.room] ?? 999) - (roomRank[b.room] ?? 999);
    if (rDiff !== 0) return rDiff;
    const tDiff = a.deviceType.localeCompare(b.deviceType);
    if (tDiff !== 0) return tDiff;
    return a.nickname.localeCompare(b.nickname);
  });

  // Write back
  range.setValues(rows.map(r => r.values));
  range.setBackgrounds(rows.map(r => r.backgrounds));
  range.setFontColors(rows.map(r => r.fontColors));
  range.setFontWeights(rows.map(r => r.fontWeights));
  range.setNotes(rows.map(r => r.notes));

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Sorted ' + rows.length + ' rows by room → device type → nickname',
    'Sort Complete'
  );
}

// ONE-TIME MIGRATION: setupStatusColumn
// Run once to add the Status column (col 18) and migrate existing room values.
// After running, delete this function (or leave it — it's safe to re-run).
function setupStatusColumn() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Items');
  const STATUS_COL = 18;
  const ROOM_COL   = 1;

  // Status-like values that live in the Room column but are really statuses
  const STATUS_ROOMS = new Set([
    'Extra', 'Unassigned', 'For sale', 'Gone', 'Malfunctioned', 'Old home'
  ]);

  const lastRow = sheet.getLastRow();

  // 1. Add / overwrite column 18 header
  sheet.getRange(1, STATUS_COL).setValue('Status')
    .setFontWeight('bold')
    .setBackground('#d9ead3'); // light green, matching other header colours

  // 2. Build status values for all data rows
  const roomValues   = sheet.getRange(2, ROOM_COL, lastRow - 1, 1).getValues();
  const statusValues = roomValues.map(([room]) => {
    if (STATUS_ROOMS.has(room)) return [room];   // move room → status
    return ['Active'];                            // real location → Active
  });

  // 3. Write Status column
  sheet.getRange(2, STATUS_COL, lastRow - 1, 1).setValues(statusValues);

  // 4. Clear the Room cell for rows where Room was a status value
  const updatedRooms = roomValues.map(([room]) => {
    return STATUS_ROOMS.has(room) ? [''] : [room];
  });
  sheet.getRange(2, ROOM_COL, lastRow - 1, 1).setValues(updatedRooms);

  // 5. Summary
  const moved = statusValues.filter(([s]) => s !== 'Active').length;
  ss.toast(
    `Status column created.\n${moved} rows moved from Room → Status.\n${lastRow - 1 - moved} rows set to Status = Active.`,
    'Setup Complete',
    8
  );
}
// END ONE-TIME MIGRATION
