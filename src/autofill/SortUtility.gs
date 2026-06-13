/**
 * Smart Home Inventory - Sort Utility
 *
 * Sorts all data rows by Room → Device type → Manufacturer → Nickname.
 *
 * The order for each of the first three keys is defined EXPLICITLY in the
 * config block below (it is NOT alphabetical and is NOT read from the
 * dropdown). Any value that is not listed sorts AFTER all listed values,
 * alphabetically. Rows with an empty Room sort last. Nickname is a natural,
 * numeric-aware sort ("lamp 2" before "lamp 10").
 *
 * To change the order, just edit the arrays in the SORT ORDER CONFIG block.
 */

// =============================================================================
// SORT ORDER CONFIG  — edit these arrays to change the sort order
// =============================================================================

// Physical room order (most important). Rooms not listed here sort after these
// alphabetically; rows with an empty Room sort last.
const ROOM_ORDER = [
  'Balcony', 'Bedroom', 'Wardrobe', 'Bathroom', 'Staircase', 'Hallway',
  'Storage', 'Lower bathroom', 'Office', 'Studio', 'Kitchen', 'Living room',
  'Terrace', 'Backyard', 'Garage door', 'Garage', 'Front door', 'Front yard',
  'Porch', 'Apartment', 'On the go', 'Unassigned', 'Extra', 'Old home',
  'For sale', 'Gone', 'Malfunctioned'
];

// Device-type order. Types not listed (e.g. Hub, Appliance, Server, Printer,
// Console, Computer, Phone, Watch, Tag, Cord) sort after these alphabetically.
const DEVICE_TYPE_ORDER = [
  'Light', 'Plug', 'Sensor', 'Switch', 'Speaker', 'Wifi', 'Climate', 'Camera', 'TV'
];

// Manufacturer order. Manufacturers not listed sort after these alphabetically.
const MANUFACTURER_ORDER = [
  'Philips Hue', 'Google', 'Xiaomi', 'Amazon', 'Samsung'
];

// =============================================================================
// SORT
// =============================================================================

/**
 * Builds a comparable rank for a value against an explicit order array.
 * Returns [tier, key]:
 *   tier 0 = listed   → key is the index in the array (numeric)
 *   tier 1 = unlisted → key is the lowercased value (alphabetical)
 *   tier 2 = empty    → sorts last
 */
function rankByOrder_(value, orderArr) {
  const v = String(value || '').trim();
  if (!v) return [2, ''];
  const i = orderArr.indexOf(v);
  if (i >= 0) return [0, i];
  return [1, v.toLowerCase()];
}

function compareRank_(a, b) {
  if (a[0] !== b[0]) return a[0] - b[0];          // tier first
  if (a[0] === 0)    return a[1] - b[1];          // both listed → numeric index
  return String(a[1]).localeCompare(String(b[1])); // both unlisted → alphabetical
}

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

  const range       = sheet.getRange(firstDataRow, 1, lastRow - firstDataRow + 1, lastCol);
  const values      = range.getValues();
  const backgrounds = range.getBackgrounds();
  const fontColors  = range.getFontColors();
  const fontWeights = range.getFontWeights();
  const notes       = range.getNotes();

  // 0-based column indices
  const ROOM_IDX = COLUMN_MAP['Room']         - 1;
  const TYPE_IDX = COLUMN_MAP['Device type']  - 1;
  const MFR_IDX  = COLUMN_MAP['Manufacturer'] - 1;
  const NICK_IDX = COLUMN_MAP['Nickname']     - 1;

  const rows = values.map((v, i) => ({
    values:       v,
    backgrounds:  backgrounds[i],
    fontColors:   fontColors[i],
    fontWeights:  fontWeights[i],
    notes:        notes[i],
    room:         v[ROOM_IDX],
    deviceType:   v[TYPE_IDX],
    manufacturer: v[MFR_IDX],
    nickname:     String(v[NICK_IDX] || '')
  }));

  rows.sort((a, b) => {
    let c = compareRank_(rankByOrder_(a.room, ROOM_ORDER),
                         rankByOrder_(b.room, ROOM_ORDER));
    if (c) return c;

    c = compareRank_(rankByOrder_(a.deviceType, DEVICE_TYPE_ORDER),
                     rankByOrder_(b.deviceType, DEVICE_TYPE_ORDER));
    if (c) return c;

    c = compareRank_(rankByOrder_(a.manufacturer, MANUFACTURER_ORDER),
                     rankByOrder_(b.manufacturer, MANUFACTURER_ORDER));
    if (c) return c;

    return a.nickname.localeCompare(b.nickname, undefined,
                                    { numeric: true, sensitivity: 'base' });
  });

  // Write back values + formatting (preserves row colors, font colors/weights, notes)
  range.setValues(rows.map(r => r.values));
  range.setBackgrounds(rows.map(r => r.backgrounds));
  range.setFontColors(rows.map(r => r.fontColors));
  range.setFontWeights(rows.map(r => r.fontWeights));
  range.setNotes(rows.map(r => r.notes));

  ss.toast(
    'Sorted ' + rows.length + ' rows by Room → Device type → Manufacturer → Nickname.',
    'Sort Complete'
  );
}
