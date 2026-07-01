# Smart Home Inventory — project context

A Google Sheet ("Smart Home Inventory") plus a **container-bound Apps Script** that catalogues every smart/home device.

## IDs
- **Spreadsheet:** `1jV_FbzzY1N-4-fmR5M9HzKcNn_ZXcWQTA7Jf-1fWd4g`
- **Apps Script project (bound):** `18ozKcnB6IspKlKNm9TMILLqG_F0hFQ35H7lbu5FxgDUsPsE7HbqCGxp9`

## Layout
- Code lives in `src/autofill/*.gs` (+ `AddModelDialog.html`).
- One logical change per commit. Keep `src/autofill/` and the live script in sync.

## Sheet structure (Items tab)
Header row 1, data from row 2 (`CONFIG`). Column indices live in `COLUMN_MAP` (Config.gs, 1-based) — the source of truth; don't re-list them here. Non-obvious bit: the sheet also has a **Retail price (€)** column (ballpark estimates) that is deliberately **not** in `COLUMN_MAP`, so no script logic depends on it.

## Conventions (important)
- **No dated comments in code** — git records dates.
- **Sort order is hardcoded** in `SortUtility.gs` (`ROOM_ORDER`, `DEVICE_TYPE_ORDER`, `MANUFACTURER_ORDER`). It is the user's **physical** room order — NOT alphabetical, NOT the dropdown order. Don't "tidy" it.
- Sorting uses **minimal whole-row moves** (LIS-based) so version history shows small moves, not a full rewrite.

## Files & features
- `AutoFillDeviceInfo.gs` — `onEdit` autofills device fields when a Model is entered; `getDeviceInfo(model)` + `fillDeviceInformation(sheet,row,info)`.
- `DeviceDatabase.gs` — `DEVICE_MAP` (model → device info).
- `Config.gs` — `COLUMN_MAP`, `CONFIG`, `FILLABLE_FIELDS`, helpers.
- `SortUtility.gs` — `sortItemsWithinRooms()`.
- `ValidationUtility.gs` — validate/fix against the database.
- `MenuManager.gs` — `onOpen` builds the "🏠 Smart Home Tools" menu.
- `DropdownAutoAdd.gs` — Installable onEdit: typing a value not in a dropdown offers to add it. Enable: Smart Home Tools → ⚙️ Setup → Enable dropdown auto-add (authorize once; it also relaxes reject→warning validation).
- `PhotoIntake.gs` — See below.

## PhotoIntake.gs — finishing the deploy
Drop a device-label photo into a Drive "Inbox" folder → a 15-min trigger reads it, matches the model against `DEVICE_MAP`, appends a **flagged** review row to Items (auto-filling known models) with the photo linked in Notes, then moves the photo to a "Processed" folder. If no model is read the row is still created (orange flag) with the photo + raw text, so you just type the model and autofill does the rest.

**Pluggable readers.** The label reader is a plugin chosen by a persisted reader mode (Script Property `INTAKE_READER_MODE`), set from the menu (📷 Photo Intake → Reader mode):
- `ocr` — OCR via the Advanced Drive Service + `DEVICE_MAP` match (no key; only recognises models already in the database). `extractDeviceFromText_()`.
- `ai` — Claude vision: the photo is POSTed to the Anthropic Messages API (`UrlFetchApp`, no SDK), which reads model/manufacturer/serial off **any** label. `extractDeviceWithAI_()`. Needs an API key — Reader mode → Set Claude API key (stored in Script Property `INTAKE_CLAUDE_API_KEY`, never in code). Model is `INTAKE_AI_MODEL` (`claude-opus-4-8`, ~3¢/photo; swap to `claude-haiku-4-5` for ~sub-cent). Returns strict JSON parsed by `parseAiJson_`.
- `off` — default; trigger removed, worker no-ops.

Add a third engine by adding one entry to `INTAKE_READERS` — the worker dispatches on mode and nothing else changes. The `ai` reader needs the `script.external_request` scope (for `UrlFetchApp`) in addition to the OCR scopes.

To deploy:
1. `clasp push` (or paste the file into the editor).
2. **Enable the Advanced Drive Service** "Drive API" (v2) — needed for OCR. In the manifest: `dependencies.enabledAdvancedServices` → `{ "userSymbol": "Drive", "serviceId": "drive", "version": "v2" }`.
3. Ensure OAuth scopes include: `spreadsheets`, `drive`, `documents`, `script.scriptapp`, `script.container.ui`.
4. Run `setupPhotoIntake` (authorize Drive + triggers). It creates the Drive folders and the 15-min trigger and shows the Inbox folder URL.
5. Drop a label photo in "Smart Home Inventory - Inbox", run "Process inbox now".

## Deploying with clasp
```
npm i -g @google/clasp
clasp login
# .clasp.json (already in repo) points at the scriptId with rootDir src/autofill
clasp pull   # first time: pulls appsscript.json (manifest) into src/autofill
# add the Drive advanced service + scopes above to appsscript.json, then:
clasp push
```
Note: `clasp pull` overwrites local with live — do it into a scratch checkout if you don't want to lose `PhotoIntake.gs` (which is in the repo but not yet live).

## Statistics tab
A live dashboard (KPI cards, Value summary, By Room + column chart, By Manufacturer, By Device Type + pie chart, Status). Numbers are live COUNTIF/SUMIF formulas; the category lists are fixed at build time, so regenerate if you add a brand-new room / brand / type.
