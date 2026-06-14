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
 * It applies the MINIMUM number of whole-row moves needed to reach the target
 * order: it keeps the longest subsequence of rows that are already in the right
 * relative order (the LIS) fixed, and moves only the rows that are genuinely out
 * of place (Sheet.moveRows). So adding a few rows / nudging a few values and
 * sorting shows up in version history as just those few row MOVES — not a
 * full-sheet rewrite — and formatting / notes / data-validation travel with each
 * row automatically. Only if a genuinely large reorder is needed (more than
 * MOVE_LIMIT moves) does it fall back to a single bulk rewrite, for speed.
 *
 * To change the order, just edit the arrays in the SORT ORDER CONFIG block.
 */

// =============================================================================
// SORT ORDER CONFIG  — edit these arrays to change the sort order
// =============================================================================

// Physical room order (most important). Rooms not listed here sort after these
// alphabetically; rows with an empty Room sort last.
const ROOM_ORDER = [
  'Balcony',
  'Bedroom',
  'Wardrobe',
  'Bathroom',
  'Staircase',
  'Hallway',
  'Storage',
  'Lower bathroom',
  'Office',
  'Studio',
  'Kitchen',
  'Living room',
  'Terrace',
  'Backyard',
  'Garage door',
  'Garage',
  'Front door',
  'Front yard',
  'Porch',
  'Apartment',
  'On the go',
  'Unassigned',
  'Extra',
  'Old home',
  'For sale',
  'Gone',
  'Malfunctioned'
];

// Device-type order. Types not listed (e.g. Hub, Appliance, Server, Printer,
// Console, Computer, Phone, Watch, Tag, Cord) sort after these alphabetically.
const DEVICE_TYPE_ORDER = [
  'Light',
  'Plug',
  'Sensor',
  'Switch',
  'Speaker',
  'Wifi',
  'Climate',
  'Camera',
  'TV'
];

// Manufacturer order. Manufacturers not listed sort after these alphabetically.
const MANUFACTURER_ORDER = [
  'Philips Hue',
  'Google',
  'Xiaomi',
  'Amazon',
  'Samsung'
];

// If reaching the target order needs more than this many row moves, fall back to
// a single bulk rewrite instead (faster for big one-time reshuffles).
const MOVE_LIMIT = 60;

// =============================================================================
// COMPARATOR
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

/**
 * Indices of a longest strictly-increasing subsequence of seq.
 * These positions are already in the right relative order, so they stay put.
 */
function lisIndices_(seq) {
  const n = seq.length;
  const tails = [];
  const prev = new Array(n).fill(-1);
  for (let i = 0; i < n; i++) {
    const x = seq[i];
    let lo = 0, hi = tails.length;
    while (lo < hi) { const m = (lo + hi) >> 1; if (seq[tails[m]] < x) lo = m + 1; else hi = m; }
    if (lo > 0) prev[i] = tails[lo - 1];
    tails[lo] = i;
  }
  let k = tails.length ? tails[tails.length - 1] : -1;
  const res = [];
  while (k >= 0) { res.push(k); k = prev[k]; }
  return res.reverse();
}

// =============================================================================
// SORT
// =============================================================================

function sortItemsWithinRooms() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  const firstRow = CONFIG.DATA_START_ROW;
  const lastRow  = sheet.getLastRow();
  const lastCol  = sheet.getLastColumn();
  const n = lastRow - firstRow + 1;

  if (n < 2) { ss.toast('Nothing to sort.', 'Sort'); return; }

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
  const want = desired.map(it => it.idx);   // target: physical position -> original index

  let alreadySorted = true;
  for (let i = 0; i < n; i++) { if (want[i] !== i) { alreadySorted = false; break; } }
  if (alreadySorted) { ss.toast('Already in order.', 'Sort'); return; }

  // Keep the longest run of rows already in the right relative order; move the rest.
  const targetPos = new Array(n);
  want.forEach((idx, pos) => { targetPos[idx] = pos; });
  const seq = [];
  for (let i = 0; i < n; i++) seq.push(targetPos[i]);   // order is the identity 0..n-1
  const keep = new Set(lisIndices_(seq));               // physical indices (= original idx) to keep

  // Plan the moves on a simulated array that mirrors Sheet.moveRows exactly.
  const cur = [];
  for (let i = 0; i < n; i++) cur.push(i);
  const moves = [];
  for (let k = 0; k < n; k++) {
    const R = want[k];
    if (keep.has(R)) continue;                  // anchor row — leave it
    const p = cur.indexOf(R);
    let destRel;
    if (k === 0) { if (p === 0) continue; destRel = 0; }
    else {
      const aPos = cur.indexOf(want[k - 1]);    // the row that should precede R
      if (p === aPos + 1) continue;             // already right after it
      destRel = aPos + 1;
    }
    moves.push([p, destRel]);
    // simulate moveRows(p -> insert before destRel)
    const el = cur.splice(p, 1)[0];
    const ins = (p < destRel) ? destRel - 1 : destRel;
    cur.splice(ins, 0, el);
  }

  if (moves.length === 0) { ss.toast('Already in order.', 'Sort'); return; }

  // Apply as whole-row moves when reasonable — clean version history; formatting,
  // notes and validation travel with each row automatically.
  if (moves.length <= MOVE_LIMIT) {
    for (let m = 0; m < moves.length; m++) {
      const p = moves[m][0], destRel = moves[m][1];
      sheet.moveRows(sheet.getRange((firstRow + p) + ':' + (firstRow + p)), firstRow + destRel);
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
  ss.toast('Sorted ' + n + ' rows (bulk rewrite — ' + moves.length + ' moves needed).', 'Sort Complete');
}
