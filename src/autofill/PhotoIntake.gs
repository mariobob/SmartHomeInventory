/**
 * Smart Home Inventory - Photo Intake
 *
 * Drop a photo of a device label into the Drive "Inbox" folder and this
 * pipeline reads the label (OCR), matches the model against the Device
 * Database, appends a flagged row to the Items sheet (auto-filling known
 * models), then files the photo into a "Processed" folder and links it in the
 * row's Notes.
 *
 * Set up once:  Smart Home Tools -> 📷 Photo Intake -> Set up photo intake
 * Run any time: Smart Home Tools -> 📷 Photo Intake -> Process inbox now
 *               (it also runs automatically every 15 minutes once set up)
 *
 * FUTURE-PROOFING: the label reader is isolated in extractDeviceFromText_().
 * Today it is OCR + Device-Database match. To grow into the "snap-and-forget"
 * idea, swap that one function for an AI-vision call (pass the image blob to a
 * model, get back model/serial/brand) - nothing else has to change. The Drive
 * folder can also be fed automatically from a phone or a Photos Picker later.
 *
 * Requires the Advanced Drive Service ("Drive API") to be enabled for OCR.
 */

const INTAKE_INBOX_NAME = 'Smart Home Inventory - Inbox';
const INTAKE_DONE_NAME  = 'Smart Home Inventory - Processed';
const INTAKE_PROP_INBOX = 'INTAKE_INBOX_ID';
const INTAKE_PROP_DONE  = 'INTAKE_DONE_ID';
const INTAKE_MAX_PER_RUN = 10;

/**
 * One-time setup: creates the Drive folders and installs the 15-minute trigger.
 * The first run asks you to authorize Drive + triggers.
 */
function setupPhotoIntake() {
  const props = PropertiesService.getScriptProperties();
  const inbox = getOrCreateFolder_(props, INTAKE_PROP_INBOX, INTAKE_INBOX_NAME);
  const done  = getOrCreateFolder_(props, INTAKE_PROP_DONE,  INTAKE_DONE_NAME);

  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'processPhotoInbox') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('processPhotoInbox').timeBased().everyMinutes(15).create();

  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Photo intake ready',
    'Drop device-label photos into this Drive folder:\n\n' + inbox.getName() + '\n' + inbox.getUrl() +
    '\n\nIt is checked automatically every 15 minutes. You can also run ' +
    '"Process inbox now" at any time. Processed photos are moved to "' + done.getName() + '".',
    ui.ButtonSet.OK
  );
}

/** Turns off the automatic checking (folders and processed photos are kept). */
function disablePhotoIntake() {
  let removed = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'processPhotoInbox') { ScriptApp.deleteTrigger(t); removed++; }
  });
  SpreadsheetApp.getUi().alert(
    removed ? 'Automatic photo intake is now off. Folders are kept; you can still run "Process inbox now".'
            : 'Automatic photo intake was not running.'
  );
}

function getOrCreateFolder_(props, key, name) {
  const id = props.getProperty(key);
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const existing = DriveApp.getFoldersByName(name);
  const folder = existing.hasNext() ? existing.next() : DriveApp.createFolder(name);
  props.setProperty(key, folder.getId());
  return folder;
}

/**
 * Worker: processes new photos in the Inbox folder. Runs on the trigger and
 * from the "Process inbox now" menu item.
 */
function processPhotoInbox() {
  const props = PropertiesService.getScriptProperties();
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

    let text = '';
    try { text = ocrImage_(file); } catch (e) { text = ''; log.push('OCR error: ' + e.message); }
    const found = extractDeviceFromText_(text);
    appendIntakeRow_(sheet, file, found, text);
    added++;
    if (found.model) matched++;

    try { file.moveTo(done); }
    catch (e) { try { done.addFile(file); inbox.removeFile(file); } catch (e2) {} }
    log.push((found.model || '(no model matched)'));
  }

  const msg = processed
    ? ('Photo intake: ' + processed + ' photo(s), ' + added + ' row(s) added, ' + matched + ' matched a known model.')
    : 'No new photos found in the Inbox folder.';
  Logger.log(msg + ' ' + JSON.stringify(log));
  ss.toast(msg, 'Photo Intake', 6);
}

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
 * v1 label reader: find a known model number from the Device Database inside
 * the OCR text. Returns { model, serial }. UPGRADE HOOK: replace the body with
 * an AI-vision call for full extraction of any label.
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

/** Appends a flagged review row for one photo. */
function appendIntakeRow_(sheet, file, found, ocrText) {
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

  const note = '📷 From photo ' + new Date().toLocaleString() +
    (found.model ? ' · matched model ' + found.model : ' · model not recognised - please fill in') +
    ' · ' + file.getUrl() +
    (ocrText ? ' · OCR: ' + ocrText.replace(/\s+/g, ' ').slice(0, 120) : '');
  sheet.getRange(newRow, COLUMN_MAP['Notes']).setValue(note);

  dst.setBackground(found.model ? '#fff8e1' : '#ffe0b2');   // soft yellow if matched, orange if unknown
  sheet.getRange(newRow, COLUMN_MAP['Model']).setNote('⏳ Added from a photo - please review.');
  return newRow;
}
