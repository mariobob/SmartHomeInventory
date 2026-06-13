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
 * HOW IT APPLIES THE SORT
 * The sort is applied as the minimum set of whole-row moves needed to reach
 * the target order (Sheet.moveRows). This keeps Google Sheets version history
 * clean — e.g. adding three rows and sorting shows up as three row MOVES, not a
 * full-sheet value rewrite — and it preserves each row's values, formatting,
 * notes and data validation automatically. If a very large reorder is required
 * (more than MOVE_LIMIT moves) it falls back to a single bulk rewrite for speed.
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

// If reaching the target order needs more than this many row moves, fall back to
// a single bulk rewrite instead (faster for big reshuffles).
const MOVE_LIMIT = 60;

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

/** Full row comparator: Room → Device type → Manufacturer → Nickname. */
function itemCompare_(a, b) {
  let c = compareRank_(rankByOrder_(a.room, ROOM_ORDER),
                       rankByOrder_(b.room, ROOM_ORDER));
  if (c) return c;
  c = compareRank_(rankByOrder_(a.dt, DEVICE_TYPE_ORDER),
                   rankByOrder_(b.dt, DEVICE_TYPE_ORDER));
  if (c) return c;
  c = compareRank_(rankByOrder_(a.mf, MANUFACTURER_ORDER),
                   rankByOrder_(b.mf, MANUFACTURER_ORDER));
  if (c) return c;
  return a.nick.localeCompare(b.nick, undefined, { numeric: true, sensitivity: 'base' });
}

function sortItemsWithinRooms() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  const firstRow = CONFIG.DATA_START_ROW;
  const lastRow  = sheet.getLastRow();
  const lastCol  = sheet.getLastColumn();
  const n = lastRow - firstRow + 1;

  if (n < 2) { ss.toast('Nothing to sort.', 'Sort'); return; }

  // Read only the key columns to decide the order.
  const RM = COLUMN_MAP['Room']         - 1;
  const DT = COLUMN_MAP['Device type']  - 1;
  const MF = COLUMN_MAP['Manufacturer'] - 1;
  const NK = COLUMN_MAP['Nickname']     - 1;

  const values = sheet.getRange(firstRow, 1, n, lastCol).getValues();
  const items = values.map((v, i) => ({
    idx: i, room: v[RM], dt: v[DT], mf: v[MF], nick: String(v[NK] || '')
  }));

  // Target order (stable sort keeps equal rows in their current relative order).
  const desired = items.slice().sort(itemCompare_);
  const want  = desired.map(it => it.idx);   // target: physical position -> original index
  const order = items.map(it => it.idx);     // current physical order

  let alreadySorted = true;
  for (let i = 0; i < n; i++) { if (order[i] !== want[i]) { alreadySorted = false; break; } }
  if (alreadySorted) { ss.toast('Already in order.', 'Sort'); return; }

  // Plan the minimum moves (selection from the top) on a simulated array.
  const sim = order.slice();
  const moves = [];
  for (let k = 0; k < n; k++) {
    if (sim[k] === want[k]) continue;
    let p = -1;
    for (let j = k + 1; j < n; j++) { if (sim[j] === want[k]) { p = j; break; } }
    moves.push([p, k]);                       // move row at sim-pos p up to pos k
    sim.splice(k, 0, sim.splice(p, 1)[0]);
  }

  // Apply as whole-row moves when reasonable — clean version history, and
  // formatting / notes / validation travel with each row automatically.
  if (moves.length <= MOVE_LIMIT) {
    for (let m = 0; m < moves.length; m++) {
      const p = moves[m][0], k = moves[m][1];
      const srcRow = firstRow + p;
      sheet.moveRows(sheet.getRange(srcRow + ':' + srcRow), firstRow + k);
    }
    SpreadsheetApp.flush();
    ss.toast('Sorted with ' + moves.length + ' row move' + (moves.length === 1 ? '' : 's') + '.',
             'Sort Complete');
    return;
  }

  // Fallback: single bulk rewrite for large reorders (preserves formatting/notes).
  const range       = sheet.getRange(firstRow, 1, n, lastCol);
  const backgrounds = range.getBackgrounds();
  const fontColors  = range.getFontColors();
  const fontWeights = range.getFontWeights();
  const notes       = range.getNotes();

  const rows = items.map((it, i) => ({
    it: it, v: values[i], bg: backgrounds[i], fc: fontColors[i], fw: fontWeights[i], nt: notes[i]
  }));
  rows.sort((a, b) => itemCompare_(a.it, b.it));

  range.setValues(rows.map(r => r.v));
  range.setBackgrounds(rows.map(r => r.bg));
  range.setFontColors(rows.map(r => r.fc));
  range.setFontWeights(rows.map(r => r.fw));
  range.setNotes(rows.map(r => r.nt));
  ss.toast('Sorted ' + n + ' rows (bulk rewrite).', 'Sort Complete');
}
