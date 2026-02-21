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
