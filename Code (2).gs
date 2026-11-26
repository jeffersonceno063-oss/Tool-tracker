/**
 * Tool Tracker + Inventory - Code.gs
 *
 * Setup: open your Google Sheet -> Extensions -> Apps Script. Paste this file as
 * Code.gs, add Index.html and View.html (and optionally Tests.gs), then
 * Deploy -> New deployment -> Web app. The Tools, Log, Inventory and Locations
 * tabs are created automatically on first use.
 *
 * Two pages are served from the same web app link:
 *   - Index.html  (no parameters)  the manager: scan, add, edit, delete, print labels.
 *   - View.html   (?loc=... or ?id=...)  read-only list. This is what a phone's built-in
 *     camera opens when it scans a QR code.
 *
 * If this script is NOT attached to the Sheet (created on script.google.com),
 * paste the Sheet's ID below (the long string in the Sheet's URL).
 */
const SPREADSHEET_ID = '';

const TOOLS_HEADERS = ['Tool ID', 'Name', 'Status', 'Held by', 'Location'];
const LOG_HEADERS = ['Time', 'Tool ID', 'Name', 'Person', 'Location', 'Action'];
const INVENTORY_HEADERS = ['Item ID', 'Name', 'Quantity', 'Location'];
const LOCATIONS_HEADERS = ['Location', 'Location ID'];

// Column positions (1-based)
const COL = { ID: 1, NAME: 2, STATUS: 3, HELD_BY: 4, LOCATION: 5 };   // Tools tab
const INV = { ID: 1, NAME: 2, QTY: 3, LOCATION: 4 };                    // Inventory tab
const LOC = { NAME: 1, ID: 2 };                                         // Locations tab

// Which page a link opens. Scanning a QR with a phone camera always carries ?loc= or ?id=.
function pageFor_(params) {
  return params && (params.loc || params.id) ? 'View' : 'Index';
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  const t = HtmlService.createTemplateFromFile(pageFor_(p));
  t.initialId = p.id || '';
  t.initialLoc = p.loc || '';
  t.appUrl = ScriptApp.getService().getUrl();
  return t.evaluate()
    .setTitle('Tool Tracker')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ---------- Sheet helpers ----------

function getSpreadsheet_() {
  const ss = SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error(
      'This script is not attached to a Google Sheet. Open the Sheet, go to ' +
      'Extensions -> Apps Script, or set SPREADSHEET_ID at the top of Code.gs.'
    );
  }
  return ss;
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (sheet) return sheet;

  // Reuse an empty default "Sheet1" for the Tools tab instead of leaving it behind.
  if (name === 'Tools') {
    const first = ss.getSheets()[0];
    if (first && first.getName() === 'Sheet1' && first.getLastRow() === 0) {
      first.setName('Tools');
      sheet = first;
    }
  }
  if (!sheet) sheet = ss.insertSheet(name);

  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
  return sheet;
}

let sheetsCache_ = null;

function sheets_() {
  if (sheetsCache_) return sheetsCache_;
  const ss = getSpreadsheet_();
  const tools = ensureSheet_(ss, 'Tools', TOOLS_HEADERS);
  const log = ensureSheet_(ss, 'Log', LOG_HEADERS);
  const inventory = ensureSheet_(ss, 'Inventory', INVENTORY_HEADERS);
  const locations = ensureSheet_(ss, 'Locations', LOCATIONS_HEADERS);
  // Older versions created the Locations tab with only a "Location" column.
  if (String(locations.getRange(1, LOC.ID).getValue()).trim() !== 'Location ID') {
    locations.getRange(1, LOC.ID).setValue('Location ID').setFontWeight('bold');
  }
  sheetsCache_ = { tools: tools, log: log, inventory: inventory, locations: locations };
  return sheetsCache_;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

// IDs only ever go up. A counter is kept so that deleting the newest tool, item or
// location never lets its ID be handed out again (an old sticker must not point at a
// different thing).
function nextNum_(prefix, existingMax) {
  const props = PropertiesService.getScriptProperties();
  const key = 'last_' + prefix;
  const n = Math.max(Number(props.getProperty(key)) || 0, existingMax) + 1;
  props.setProperty(key, String(n));
  return n;
}

function formatId_(prefix, n) {
  return prefix + '-' + String(n).padStart(3, '0');
}

// Accepts "TL-001" or a full link containing ?id=TL-001
function cleanId_(raw) {
  raw = String(raw || '').trim();
  const m = /[?&]id=([^&#\s]+)/.exec(raw);
  return decodeURIComponent(m ? m[1] : raw).toUpperCase();
}

// ---------- Locations ----------
// Every location has a permanent ID (LOC-001...). The QR code holds the ID, so a location
// can be renamed without reprinting its label. Items and tools store the location NAME,
// which keeps the Sheet easy to read and edit by hand; renaming updates them all.

function normLoc_(s) {
  return String(s || '').trim().replace(/\s+/g, ' ');
}

function locKey_(s) {
  return normLoc_(s).toLowerCase();
}

function isLocId_(s) {
  return /^LOC-\d+$/i.test(String(s || '').trim());
}

// "LOC-001", "Rack A", or a full link containing ?loc=... -> cleaned text
function cleanLoc_(raw) {
  raw = String(raw || '').trim();
  const m = /[?&]loc=([^&#]+)/.exec(raw);
  return normLoc_(m ? decodeURIComponent(m[1].replace(/\+/g, ' ')) : raw);
}

// Reads the Locations tab and finds anything that still needs an ID or a row.
function computeRegistry_() {
  const s = sheets_();
  const entries = [];
  const byKey = Object.create(null);
  const noId = [];
  let maxNum = 0;

  const locData = s.locations.getDataRange().getValues();
  for (let i = 1; i < locData.length; i++) {
    const name = normLoc_(locData[i][LOC.NAME - 1]);
    if (!name || byKey[name.toLowerCase()]) continue;
    const id = String(locData[i][LOC.ID - 1] || '').trim().toUpperCase();
    const e = { row: i + 1, name: name, id: id };
    byKey[name.toLowerCase()] = e;
    entries.push(e);
    const m = /^LOC-(\d+)$/.exec(id);
    if (m) maxNum = Math.max(maxNum, Number(m[1])); else noId.push(e);
  }

  // Locations typed by hand into Inventory or Tools that are not registered yet.
  const newNames = [];
  const queued = Object.create(null);
  function scan(sheet, col) {
    sheet.getDataRange().getValues().slice(1).forEach(function (r) {
      const n = normLoc_(r[col - 1]);
      const k = n.toLowerCase();
      if (n && !byKey[k] && !queued[k]) { queued[k] = true; newNames.push(n); }
    });
  }
  scan(s.inventory, INV.LOCATION);
  scan(s.tools, COL.LOCATION);

  return { entries: entries, noId: noId, newNames: newNames, maxNum: maxNum };
}

// Writes missing IDs and rows. Call only while holding the lock.
function syncLocations_() {
  const reg = computeRegistry_();
  if (!reg.noId.length && !reg.newNames.length) return reg.entries;
  const sheet = sheets_().locations;
  reg.noId.forEach(function (e) {
    e.id = formatId_('LOC', nextNum_('LOC', reg.maxNum));
    sheet.getRange(e.row, LOC.ID).setValue(e.id);
  });
  reg.newNames.forEach(function (n) {
    const id = formatId_('LOC', nextNum_('LOC', reg.maxNum));
    sheet.appendRow([n, id]);
    reg.entries.push({ row: sheet.getLastRow(), name: n, id: id });
  });
  return reg.entries;
}

// For read-only callers: takes the lock only when something has to be written.
function registry_() {
  const reg = computeRegistry_();
  if (!reg.noId.length && !reg.newNames.length) return reg.entries;
  return withLock_(syncLocations_);
}

function pickLocation_(entries, raw) {
  const text = cleanLoc_(raw);
  if (!text) return null;
  if (isLocId_(text)) {
    const id = text.toUpperCase();
    for (let i = 0; i < entries.length; i++) if (entries[i].id === id) return entries[i];
    return null;
  }
  const k = text.toLowerCase();
  for (let i = 0; i < entries.length; i++) if (entries[i].name.toLowerCase() === k) return entries[i];
  return null;
}

function maxLocNum_(entries) {
  let max = 0;
  entries.forEach(function (e) {
    const m = /^LOC-(\d+)$/.exec(e.id);
    if (m) max = Math.max(max, Number(m[1]));
  });
  return max;
}

// Finds a location by ID or name, creating it if the name is new. Hold the lock.
function ensureLocation_(raw) {
  const entries = syncLocations_();
  const hit = pickLocation_(entries, raw);
  if (hit) return hit;
  const text = cleanLoc_(raw);
  if (!text || isLocId_(text)) return null;   // never invent a location from an unknown ID
  const sheet = sheets_().locations;
  const id = formatId_('LOC', nextNum_('LOC', maxLocNum_(entries)));
  sheet.appendRow([text, id]);
  return { row: sheet.getLastRow(), name: text, id: id };
}

function itemsAt_(name) {
  const key = name.toLowerCase();
  return sheets_().inventory.getDataRange().getValues().slice(1)
    .filter(function (r) { return r[0] && locKey_(r[INV.LOCATION - 1]) === key; })
    .map(function (r) { return { id: String(r[0]), name: String(r[1]), qty: Number(r[2]) || 0 }; })
    .sort(function (a, b) { return a.name.toLowerCase() < b.name.toLowerCase() ? -1 : 1; });
}

function toolsAt_(name) {
  const key = name.toLowerCase();
  return sheets_().tools.getDataRange().getValues().slice(1)
    .filter(function (r) { return r[0] && locKey_(r[COL.LOCATION - 1]) === key; })
    .map(function (r) {
      return {
        id: String(r[0]),
        name: String(r[1]),
        status: String(r[2] || 'Available'),
        borrower: String(r[3] || '')
      };
    });
}

function describeLocation_(e) {
  return { found: true, id: e.id, name: e.name, items: itemsAt_(e.name), tools: toolsAt_(e.name) };
}

// What is stored at a location (items and tools). Accepts an ID, a name or a ?loc= link.
function getLocation(raw) {
  const e = pickLocation_(registry_(), raw);
  return e ? describeLocation_(e) : { found: false };
}

// Every location with how many items and tools it holds.
function listLocations() {
  const entries = registry_();
  const s = sheets_();
  const itemCount = Object.create(null);
  const toolCount = Object.create(null);
  s.inventory.getDataRange().getValues().slice(1).forEach(function (r) {
    const k = locKey_(r[INV.LOCATION - 1]);
    if (r[0] && k) itemCount[k] = (itemCount[k] || 0) + 1;
  });
  s.tools.getDataRange().getValues().slice(1).forEach(function (r) {
    const k = locKey_(r[COL.LOCATION - 1]);
    if (r[0] && k) toolCount[k] = (toolCount[k] || 0) + 1;
  });
  return entries
    .map(function (e) {
      const k = e.name.toLowerCase();
      return { id: e.id, name: e.name, items: itemCount[k] || 0, tools: toolCount[k] || 0 };
    })
    .sort(function (a, b) {
      return a.name.toLowerCase() < b.name.toLowerCase() ? -1 : a.name.toLowerCase() > b.name.toLowerCase() ? 1 : 0;
    });
}

function addLocation(name) {
  return withLock_(function () {
    name = normLoc_(name);
    if (!name) return { ok: false, msg: 'Enter a location name.' };
    if (isLocId_(name)) return { ok: false, msg: 'Enter a name such as Rack A, not an ID.' };
    const existing = pickLocation_(syncLocations_(), name);
    const e = existing || ensureLocation_(name);
    return { ok: true, id: e.id, name: e.name, existed: !!existing };
  });
}

function renameInColumn_(sheet, col, oldKey, newName) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (locKey_(data[i][col - 1]) === oldKey) sheet.getRange(i + 1, col).setValue(newName);
  }
}

// Renames a location everywhere. Its ID, and so its printed QR code, does not change.
function renameLocation(raw, newName) {
  return withLock_(function () {
    newName = normLoc_(newName);
    if (!newName) return { ok: false, msg: 'Enter a location name.' };
    if (isLocId_(newName)) return { ok: false, msg: 'Enter a name such as Rack A, not an ID.' };
    const entries = syncLocations_();
    const e = pickLocation_(entries, raw);
    if (!e) return { ok: false, msg: 'Location not found.' };
    const clash = pickLocation_(entries, newName);
    if (clash && clash.id !== e.id) {
      return { ok: false, msg: 'A location named "' + clash.name + '" already exists.' };
    }
    if (newName === e.name) return { ok: true, id: e.id, name: e.name, msg: 'No change.' };
    const s = sheets_();
    const oldKey = e.name.toLowerCase();
    s.locations.getRange(e.row, LOC.NAME).setValue(newName);
    renameInColumn_(s.inventory, INV.LOCATION, oldKey, newName);
    renameInColumn_(s.tools, COL.LOCATION, oldKey, newName);
    return { ok: true, id: e.id, name: newName, msg: e.name + ' renamed to ' + newName + '. Its QR code still works.' };
  });
}

function countText_(items, tools) {
  const parts = [];
  if (items) parts.push(items + (items === 1 ? ' item' : ' items'));
  if (tools) parts.push(tools + (tools === 1 ? ' tool' : ' tools'));
  return parts.join(' and ');
}

// Only an empty location can be deleted, so nothing is left without a place.
function deleteLocation(raw) {
  return withLock_(function () {
    const e = pickLocation_(syncLocations_(), raw);
    if (!e) return { ok: false, msg: 'Location not found.' };
    const d = describeLocation_(e);
    if (d.items.length || d.tools.length) {
      return { ok: false, msg: e.name + ' still has ' + countText_(d.items.length, d.tools.length) + '. Move or remove them first.' };
    }
    sheets_().locations.deleteRow(e.row);
    return { ok: true, msg: e.name + ' deleted.' };
  });
}

// ---------- Tools ----------

function findRow_(id) {
  const data = sheets_().tools.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toUpperCase() === id) {
      return {
        row: i + 1,
        id: String(data[i][0]).trim(),
        name: String(data[i][1]),
        status: String(data[i][2] || ''),
        borrower: String(data[i][3] || ''),
        location: String(data[i][4] || '')
      };
    }
  }
  return null;
}

function getTool(raw) {
  const t = findRow_(cleanId_(raw));
  if (!t) return { found: false };
  const loc = t.location ? pickLocation_(registry_(), t.location) : null;
  return {
    found: true,
    id: t.id,
    name: t.name,
    status: t.status || 'Available',
    borrower: t.borrower,
    location: t.location,
    locationId: loc ? loc.id : ''
  };
}

function setLocation(raw, location) {
  return withLock_(function () {
    if (!cleanLoc_(location)) return { ok: false, msg: 'Enter a location.' };
    const t = findRow_(cleanId_(raw));
    if (!t) return { ok: false, msg: 'Tool not found.' };
    const loc = ensureLocation_(location);
    if (!loc) return { ok: false, msg: 'Location not found.' };
    const s = sheets_();
    s.tools.getRange(t.row, COL.LOCATION).setValue(loc.name);
    s.log.appendRow([new Date(), t.id, t.name, '', loc.name, 'LOCATION']);
    return { ok: true, msg: t.name + ' is stored at ' + loc.name + '. Next scan = borrow.' };
  });
}

function borrowTool(raw, borrower) {
  return withLock_(function () {
    borrower = String(borrower || '').trim();
    if (!borrower) return { ok: false, msg: 'Enter the borrower\'s name.' };
    const t = findRow_(cleanId_(raw));
    if (!t) return { ok: false, msg: 'Tool not found.' };
    if (!t.location) return { ok: false, msg: 'Set this tool\'s location first.' };
    if (t.status === 'Borrowed') {
      return { ok: false, msg: 'Already borrowed by ' + t.borrower + '.' };
    }
    const s = sheets_();
    s.tools.getRange(t.row, COL.STATUS, 1, 2).setValues([['Borrowed', borrower]]);
    s.log.appendRow([new Date(), t.id, t.name, borrower, t.location, 'OUT']);
    return { ok: true, msg: t.name + ' borrowed by ' + borrower + '.' };
  });
}

function returnTool(raw, location) {
  return withLock_(function () {
    if (!cleanLoc_(location)) return { ok: false, msg: 'Enter where the tool was put back.' };
    const t = findRow_(cleanId_(raw));
    if (!t) return { ok: false, msg: 'Tool not found.' };
    if (t.status !== 'Borrowed') return { ok: false, msg: 'This tool is not currently borrowed.' };
    const loc = ensureLocation_(location);
    if (!loc) return { ok: false, msg: 'Location not found.' };
    const s = sheets_();
    s.tools.getRange(t.row, COL.STATUS, 1, 3).setValues([['Available', '', loc.name]]);
    s.log.appendRow([new Date(), t.id, t.name, t.borrower, loc.name, 'IN']);
    return { ok: true, msg: t.name + ' returned to ' + loc.name + ' (from ' + t.borrower + ').' };
  });
}

function addTool(name) {
  return withLock_(function () {
    name = String(name || '').trim();
    if (!name) return { ok: false, msg: 'Enter a tool name.' };
    const tools = sheets_().tools;
    const data = tools.getDataRange().getValues();
    let max = 0;
    for (let i = 1; i < data.length; i++) {
      const m = /^TL-(\d+)$/i.exec(String(data[i][0]).trim());
      if (m) max = Math.max(max, Number(m[1]));
    }
    const id = formatId_('TL', nextNum_('TL', max));
    tools.appendRow([id, name, 'Available', '', '']);
    return { ok: true, id: id, name: name };
  });
}

function renameTool(raw, newName) {
  return withLock_(function () {
    newName = String(newName || '').trim();
    if (!newName) return { ok: false, msg: 'Enter a tool name.' };
    const t = findRow_(cleanId_(raw));
    if (!t) return { ok: false, msg: 'Tool not found.' };
    sheets_().tools.getRange(t.row, COL.NAME).setValue(newName);
    return { ok: true, msg: t.name + ' renamed to ' + newName + '. Its QR code still works.' };
  });
}

function deleteTool(raw) {
  return withLock_(function () {
    const t = findRow_(cleanId_(raw));
    if (!t) return { ok: false, msg: 'Tool not found.' };
    if (t.status === 'Borrowed') {
      return { ok: false, msg: t.name + ' is borrowed by ' + t.borrower + '. Return it first.' };
    }
    const s = sheets_();
    s.tools.deleteRow(t.row);
    s.log.appendRow([new Date(), t.id, t.name, '', t.location, 'DELETED']);
    return { ok: true, msg: t.name + ' deleted. Its history stays in the Log tab.' };
  });
}

function listTools() {
  return sheets_().tools.getDataRange().getValues().slice(1)
    .filter(function (r) { return r[0]; })
    .map(function (r) {
      return {
        id: String(r[0]),
        name: String(r[1]),
        status: String(r[2] || 'Available'),
        borrower: String(r[3] || ''),
        location: String(r[4] || '')
      };
    });
}

// ---------- Inventory items ----------

// Blank quantity -> dflt if given, otherwise invalid. Must be a whole number >= 0.
function parseQty_(v, dflt) {
  if (v === null || v === undefined || String(v).trim() === '') {
    return dflt === undefined ? null : dflt;
  }
  const n = Number(v);
  return (isFinite(n) && n >= 0 && Math.floor(n) === n) ? n : null;
}

// Adds an item to a location. The same name in the same location adds to its quantity.
function addItem(name, qty, location) {
  return withLock_(function () {
    name = String(name || '').trim();
    if (!name) return { ok: false, msg: 'Enter an item name.' };
    if (!cleanLoc_(location)) return { ok: false, msg: 'Choose a location.' };
    const q = parseQty_(qty, 1);
    if (q === null) return { ok: false, msg: 'Quantity must be a whole number, 0 or more.' };

    const loc = ensureLocation_(location);
    if (!loc) return { ok: false, msg: 'Location not found.' };
    const inv = sheets_().inventory;
    const data = inv.getDataRange().getValues();
    let max = 0;
    for (let i = 1; i < data.length; i++) {
      const m = /^IT-(\d+)$/i.exec(String(data[i][0]).trim());
      if (m) max = Math.max(max, Number(m[1]));
      if (String(data[i][INV.NAME - 1]).trim().toLowerCase() === name.toLowerCase() &&
          locKey_(data[i][INV.LOCATION - 1]) === loc.name.toLowerCase()) {
        const total = (Number(data[i][INV.QTY - 1]) || 0) + q;
        inv.getRange(i + 1, INV.QTY).setValue(total);
        return { ok: true, id: String(data[i][0]), msg: name + ': quantity is now ' + total + ' at ' + loc.name + '.' };
      }
    }
    const id = formatId_('IT', nextNum_('IT', max));
    inv.appendRow([id, name, q, loc.name]);
    return { ok: true, id: id, msg: 'Added ' + q + ' x ' + name + ' to ' + loc.name + '.' };
  });
}

function findItemRow_(id) {
  id = String(id || '').trim().toUpperCase();
  const data = sheets_().inventory.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toUpperCase() === id) {
      return { row: i + 1, id: String(data[i][0]), name: String(data[i][1]) };
    }
  }
  return null;
}

function setItemQuantity(id, qty) {
  return withLock_(function () {
    const q = parseQty_(qty);
    if (q === null) return { ok: false, msg: 'Quantity must be a whole number, 0 or more.' };
    const it = findItemRow_(id);
    if (!it) return { ok: false, msg: 'Item not found.' };
    sheets_().inventory.getRange(it.row, INV.QTY).setValue(q);
    return { ok: true, msg: it.name + ': quantity set to ' + q + '.' };
  });
}

// Changes an item's name, quantity and location (moving it to another rack) in one go.
function editItem(id, name, qty, location) {
  return withLock_(function () {
    name = String(name || '').trim();
    if (!name) return { ok: false, msg: 'Enter an item name.' };
    const q = parseQty_(qty);
    if (q === null) return { ok: false, msg: 'Quantity must be a whole number, 0 or more.' };
    if (!cleanLoc_(location)) return { ok: false, msg: 'Choose a location.' };
    const it = findItemRow_(id);
    if (!it) return { ok: false, msg: 'Item not found.' };

    // Do not allow two rows for the same item name in the same location.
    const known = pickLocation_(syncLocations_(), location);
    const targetKey = known ? known.name.toLowerCase() : cleanLoc_(location).toLowerCase();
    const inv = sheets_().inventory;
    const data = inv.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (i + 1 !== it.row &&
          String(data[i][INV.NAME - 1]).trim().toLowerCase() === name.toLowerCase() &&
          locKey_(data[i][INV.LOCATION - 1]) === targetKey) {
        return { ok: false, msg: '"' + name + '" already exists at that location. Edit that one instead.' };
      }
    }

    const loc = ensureLocation_(location);
    if (!loc) return { ok: false, msg: 'Location not found.' };
    inv.getRange(it.row, INV.NAME, 1, 3).setValues([[name, q, loc.name]]);
    return { ok: true, msg: name + ' saved (' + q + ' at ' + loc.name + ').' };
  });
}

function deleteItem(id) {
  return withLock_(function () {
    const it = findItemRow_(id);
    if (!it) return { ok: false, msg: 'Item not found.' };
    sheets_().inventory.deleteRow(it.row);
    return { ok: true, msg: it.name + ' removed.' };
  });
}
