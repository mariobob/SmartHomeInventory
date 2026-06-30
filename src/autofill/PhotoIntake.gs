/**
 * Smart Home Inventory - Photo Intake
 *
 * Drop a photo of a device label into the Drive "Inbox" folder and this
 * pipeline reads the label, matches the model against the Device Database,
 * appends a flagged row to the Items sheet (auto-filling known models), then
 * files the photo into a "Processed" folder and links it in the row's Notes.
 *
 * Set up once:  Smart Home Tools -> 📷 Photo Intake -> Set up photo intake
 * Pick reader:  Smart Home Tools -> 📷 Photo Intake -> Reader mode -> OCR / AI / Off
 * Run any time: Smart Home Tools -> 📷 Photo Intake -> Process inbox now
 *               (it also runs automatically every 15 minutes when a reader is on)
 *
 * PLUGGABLE READERS. The label reader is a swappable plugin, chosen by a
 * persisted "reader mode" (Script Property INTAKE_PROP_MODE):
 *   'ocr' — OCR via the Advanced Drive Service + Device-Database match (no key,
 *           recognises only models already in the database).
 *   'ai'  — Claude vision: the photo is sent to the Anthropic API, which reads
 *           the model / manufacturer / serial off any label (needs an API key).
 *   'off' — automation stopped: the 15-minute trigger is removed and the worker
 *           no-ops. This is the default until a reader is chosen.
 * Add a third engine by adding one entry to INTAKE_READERS — nothing else changes.
 *
 * The OCR reader requires the Advanced Drive Service ("Drive API") to be enabled.
 * The AI reader requires a Claude API key (Reader mode -> Set Claude API key);
 * the key is stored privately in Script Properties, never in the sheet or code.
 */

const INTAKE_INBOX_NAME = 'Smart Home Inventory - Inbox';
const INTAKE_DONE_NAME  = 'Smart Home Inventory - Processed';
const INTAKE_PROP_INBOX = 'INTAKE_INBOX_ID';
const INTAKE_PROP_DONE  = 'INTAKE_DONE_ID';
const INTAKE_PROP_MODE  = 'INTAKE_READER_MODE';
const INTAKE_PROP_KEY   = 'INTAKE_CLAUDE_API_KEY';
const INTAKE_MAX_PER_RUN = 10;

// Claude model used by the AI reader. Opus 4.8 reads labels best (~3¢/photo);
// switch to 'claude-haiku-4-5' to cut cost to a fraction of a cent per photo.
const INTAKE_AI_MODEL = 'claude-opus-4-8';

const INTAKE_AI_PROMPT =
  'This is a photo of a smart-home device label. Read it and return ONLY a ' +
  'JSON object (no markdown, no commentary) with the keys "model", ' +
  '"manufacturer" and "serial". Use the exact model or part number printed on ' +
  'the label for "model". Leave a field as an empty string if it is not ' +
  'visible. Example: {"model":"LCA001","manufacturer":"Philips Hue","serial":"ABC123"}';

// =============================================================================
// READER REGISTRY  — each reader takes a Drive file and returns
//   { found: { model, serial, manufacturer }, text: <raw text for the Notes cell> }
// =============================================================================

const INTAKE_READERS = {
  ocr: {
    label: 'OCR (Drive)',
    read: function (file) {
      let text = '';
      try { text = ocrImage_(file); } catch (e) { text = ''; }
      const found = extractDeviceFromText_(text);
      return { found: found, text: text };
    }
  },
  ai: {
    label: 'Claude vision',
    read: function (file) {
      return extractDeviceWithAI_(file);
    }
  }
};

// =============================================================================
// MODE  — persisted choice of reader (or 'off')
// =============================================================================

/** Returns the active reader mode: 'ocr', 'ai', or 'off' (default). */
function getIntakeReaderMode_(props) {
  props = props || PropertiesService.getScriptProperties();
  return props.getProperty(INTAKE_PROP_MODE) || 'off';
}

/**
 * Persists the reader mode and (re)installs or removes the 15-minute trigger to
 * match. 'off' removes it; 'ocr'/'ai' ensure the folders + trigger exist.
 */
function applyIntakeReaderMode_(mode) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty(INTAKE_PROP_MODE, mode);

  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'processPhotoInbox') ScriptApp.deleteTrigger(t);
  });

  if (mode !== 'off') {
    getOrCreateFolder_(props, INTAKE_PROP_INBOX, INTAKE_INBOX_NAME);
    getOrCreateFolder_(props, INTAKE_PROP_DONE,  INTAKE_DONE_NAME);
    ScriptApp.newTrigger('processPhotoInbox').timeBased().everyMinutes(15).create();
  }
}

// Menu handlers — set the reader mode.
function setIntakeReaderOCR() {
  applyIntakeReaderMode_('ocr');
  SpreadsheetApp.getActive().toast('Reader: OCR (Drive). Checked every 15 min.', 'Photo Intake', 6);
}

function setIntakeReaderAI() {
  const props = PropertiesService.getScriptProperties();
  applyIntakeReaderMode_('ai');
  const ui = SpreadsheetApp.getUi();
  if (!props.getProperty(INTAKE_PROP_KEY)) {
    ui.alert('Reader set to Claude vision',
      'No Claude API key is set yet, so processing will fail until you add one.\n\n' +
      'Run 📷 Photo Intake → Reader mode → Set Claude API key.', ui.ButtonSet.OK);
  } else {
    SpreadsheetApp.getActive().toast('Reader: Claude vision. Checked every 15 min.', 'Photo Intake', 6);
  }
}

function setIntakeReaderOff() {
  applyIntakeReaderMode_('off');
  SpreadsheetApp.getActive().toast('Photo intake is off. Folders are kept.', 'Photo Intake', 6);
}

/** Stores the Claude API key privately in Script Properties. */
function setClaudeApiKey() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('Claude API key',
    'Paste your Anthropic API key (starts with "sk-ant-"). It is stored ' +
    'privately in Script Properties — never in the sheet or the code.',
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const key = res.getResponseText().trim();
  if (!key) { ui.alert('No key entered.'); return; }
  PropertiesService.getScriptProperties().setProperty(INTAKE_PROP_KEY, key);
  ui.alert('Saved', 'Claude vision is ready. Switch Reader mode to "Claude vision" to use it.', ui.ButtonSet.OK);
}

// =============================================================================
// SETUP
// =============================================================================

/**
 * One-time setup: creates the Drive folders, switches the reader on (OCR by
 * default) and installs the 15-minute trigger. The first run asks you to
 * authorize Drive + triggers.
 */
function setupPhotoIntake() {
  const props = PropertiesService.getScriptProperties();
  const inbox = getOrCreateFolder_(props, INTAKE_PROP_INBOX, INTAKE_INBOX_NAME);
  const done  = getOrCreateFolder_(props, INTAKE_PROP_DONE,  INTAKE_DONE_NAME);

  // Default to OCR unless a reader was already chosen.
  const mode = getIntakeReaderMode_(props);
  applyIntakeReaderMode_(mode === 'off' ? 'ocr' : mode);

  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Photo intake ready',
    'Drop device-label photos into this Drive folder:\n\n' + inbox.getName() + '\n' + inbox.getUrl() +
    '\n\nIt is checked automatically every 15 minutes. Pick the reader under ' +
    '"Reader mode" (OCR or Claude vision). Processed photos are moved to "' + done.getName() + '".',
    ui.ButtonSet.OK
  );
}

/** Back-compat menu wrapper: turns the automation off. */
function disablePhotoIntake() {
  setIntakeReaderOff();
}

function getOrCreateFolder_(props, key, name) {
  const id = props.getProperty(key);
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const existing = DriveApp.getFoldersByName(name);
  const folder = existing.hasNext() ? existing.next() : DriveApp.createFolder(name);
  props.setProperty(key, folder.getId());
  return folder;
}

// =============================================================================
// WORKER
// =============================================================================

/**
 * Worker: processes new photos in the Inbox folder using the active reader.
 * Runs on the trigger and from the "Process inbox now" menu item.
 */
function processPhotoInbox() {
  const props = PropertiesService.getScriptProperties();
  const mode  = getIntakeReaderMode_(props);
  if (mode === 'off') {
    SpreadsheetApp.getActive().toast('Photo intake is off — pick a reader under Reader mode.', 'Photo Intake', 6);
    return;
  }
  const reader = INTAKE_READERS[mode];
  if (!reader) {
    SpreadsheetApp.getActive().toast('Unknown reader "' + mode + '".', 'Photo Intake', 6);
    return;
  }

  const inbox = getOrCreateFolder_(props, INTAKE_PROP_INBOX, INTAKE_INBOX_NAME);
  const done  = getOrCreateFolder_(props, INTAKE_PROP_DONE,  INTAKE_DONE_NAME);
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  const files = inbox.getFiles();
  let processed = 0, added = 0, matched = 0, seen = 0;
  const log = [];
  while (files.hasNext() && processed < INTAKE_MAX_PER_RUN && seen < 200) {
    seen++;
    const file = files.next();
    if (String(file.getMimeType()).indexOf('image/') !== 0) continue;   // only images
    processed++;

    let result;
    try {
      result = reader.read(file);
    } catch (e) {
      result = { found: { model: '', serial: '', manufacturer: '' }, text: '' };
      log.push('Reader error: ' + e.message);
    }
    const found = result.found || {};
    appendIntakeRow_(sheet, file, found, result.text || '');
    added++;
    if (found.model) matched++;

    try { file.moveTo(done); }
    catch (e) { try { done.addFile(file); inbox.removeFile(file); } catch (e2) {} }
    log.push((found.model || '(no model)'));
  }

  const msg = processed
    ? ('Photo intake (' + reader.label + '): ' + processed + ' photo(s), ' + added +
       ' row(s) added, ' + matched + ' with a model.')
    : 'No new photos found in the Inbox folder.';
  Logger.log(msg + ' ' + JSON.stringify(log));
  ss.toast(msg, 'Photo Intake', 6);
}

// =============================================================================
// READER: OCR  (Advanced Drive Service + Device-Database match)
// =============================================================================

/**
 * OCRs an image file into text using the Advanced Drive Service. Works whether
 * the enabled Drive service is v2 or v3.
 */
function ocrImage_(file) {
  const id = file.getId();
  let doc = null;
  try { doc = Drive.Files.copy({ title: 'intake-ocr-temp', mimeType: 'application/vnd.google-apps.document' }, id, { ocr: true, ocrLanguage: 'en' }); } catch (e) {}
  if (!doc) { try { doc = Drive.Files.copy({ name: 'intake-ocr-temp', mimeType: 'application/vnd.google-apps.document' }, id, { ocrLanguage: 'en' }); } catch (e) {} }
  if (!doc) throw new Error('OCR copy failed - is the Drive Advanced Service enabled?');
  let text = '';
  try { text = DocumentApp.openById(doc.id).getBody().getText(); } catch (e) {}
  try { Drive.Files.remove(doc.id); } catch (e) { try { DriveApp.getFileById(doc.id).setTrashed(true); } catch (e2) {} }
  return text || '';
}

/**
 * Find a known model number from the Device Database inside the OCR text.
 * Returns { model, serial }.
 */
function extractDeviceFromText_(text) {
  const out = { model: '', serial: '' };
  if (!text) return out;
  const compact = text.toUpperCase().replace(/[^A-Z0-9]/g, '');

  let keys = [];
  try { keys = Object.keys(DEVICE_MAP); } catch (e) {}
  try { if (typeof CUSTOM_DEVICE_MAP !== 'undefined') keys = keys.concat(Object.keys(CUSTOM_DEVICE_MAP)); } catch (e) {}
  keys.sort(function (a, b) { return String(b).length - String(a).length; });   // longest first

  for (let i = 0; i < keys.length; i++) {
    const kc = String(keys[i]).toUpperCase().replace(/[^A-Z0-9]/g, '');
    if (kc.length >= 4 && compact.indexOf(kc) !== -1) { out.model = keys[i]; break; }
  }
  const sn = text.match(/S\/?N[:\s]*([A-Z0-9\-]{5,})/i);
  if (sn) out.serial = sn[1];
  return out;
}

// =============================================================================
// READER: AI  (Claude vision via the Anthropic API)
// =============================================================================

/**
 * Sends the photo to Claude and asks it to read the label. Returns
 * { found: { model, serial, manufacturer }, text }. Throws if no key is set or
 * the API call fails (the worker logs the error and still creates a flagged row).
 */
function extractDeviceWithAI_(file) {
  const apiKey = PropertiesService.getScriptProperties().getProperty(INTAKE_PROP_KEY);
  if (!apiKey) throw new Error('No Claude API key set — run Reader mode → Set Claude API key.');

  const blob = file.getBlob();
  const payload = {
    model: INTAKE_AI_MODEL,
    max_tokens: 400,
    messages: [{
      role: 'user',
      content: [
        { type: 'image', source: { type: 'base64', media_type: blob.getContentType(), data: Utilities.base64Encode(blob.getBytes()) } },
        { type: 'text', text: INTAKE_AI_PROMPT }
      ]
    }]
  };

  const resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = resp.getResponseCode();
  const body = resp.getContentText();
  if (code !== 200) throw new Error('Claude API ' + code + ': ' + body.slice(0, 200));

  const data = JSON.parse(body);
  const textOut = (data.content || [])
    .filter(function (b) { return b.type === 'text'; })
    .map(function (b) { return b.text; })
    .join('');

  const parsed = parseAiJson_(textOut) || {};
  return {
    found: {
      model: String(parsed.model || '').trim(),
      serial: String(parsed.serial || '').trim(),
      manufacturer: String(parsed.manufacturer || '').trim()
    },
    text: textOut
  };
}

/** Tolerantly parse the JSON object out of the model's text reply. */
function parseAiJson_(text) {
  if (!text) return null;
  try { return JSON.parse(text); } catch (e) {}
  const m = text.match(/\{[\s\S]*\}/);   // strip any markdown fences / prose
  if (m) { try { return JSON.parse(m[0]); } catch (e2) {} }
  return null;
}

// =============================================================================
// ROW
// =============================================================================

/** Appends a flagged review row for one photo. */
function appendIntakeRow_(sheet, file, found, readerText) {
  found = found || {};
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const newRow = lastRow + 1;
  sheet.insertRowAfter(lastRow);

  const src = sheet.getRange(lastRow, 1, 1, lastCol);
  const dst = sheet.getRange(newRow, 1, 1, lastCol);
  src.copyTo(dst, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  src.copyTo(dst, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
  dst.clearContent();

  if (found.model) {
    sheet.getRange(newRow, COLUMN_MAP['Model']).setValue(found.model);
    try { const info = getDeviceInfo(found.model); if (info) fillDeviceInformation(sheet, newRow, info); } catch (e) {}
  }
  if (found.serial && COLUMN_MAP['Serial number']) {
    const sc = sheet.getRange(newRow, COLUMN_MAP['Serial number']);
    if (!sc.getValue()) sc.setValue(found.serial);
  }
  // The AI reader can also read the brand off the label — fill it when the
  // model wasn't in the database (so autofill didn't already supply it).
  if (found.manufacturer && COLUMN_MAP['Manufacturer']) {
    const mc = sheet.getRange(newRow, COLUMN_MAP['Manufacturer']);
    if (!mc.getValue()) mc.setValue(found.manufacturer);
  }

  const note = '📷 From photo ' + new Date().toLocaleString() +
    (found.model ? ' · model ' + found.model : ' · model not read - please fill in') +
    ' · ' + file.getUrl() +
    (readerText ? ' · read: ' + String(readerText).replace(/\s+/g, ' ').slice(0, 200) : '');
  sheet.getRange(newRow, COLUMN_MAP['Notes']).setValue(note);

  dst.setBackground(found.model ? '#fff8e1' : '#ffe0b2');   // soft yellow if matched, orange if unknown
  sheet.getRange(newRow, COLUMN_MAP['Model']).setNote('⏳ Added from a photo - please review.');
  return newRow;
}
