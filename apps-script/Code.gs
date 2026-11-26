/**
 * The Ephemera Cabinet — Google Sheets + Drive back end
 * -----------------------------------------------------
 * Paste into Extensions → Apps Script of "Design Archive — Database".
 * Enable the advanced service "Drive API" (v3) under Services (+).
 *
 * Tabs
 *   Works      one row per work
 *   Entities   clients, with a short code in `abbr`
 *   Designers  people, print houses and studios
 *   Solutions  category › subcategory › type (one type per row), with the codes used in IDs. Add a row to add an option.
 *              A work can take several types: comma-separated in Works › types.
 *   Styles     one column per movement (header = movement), styles listed below
 *   _lists     hidden helper that feeds the per-row dropdowns (subcategory, design_style). Don't edit.
 *
 * Works › types has ONE whole-column rule listing every type, with "Allow multiple selections" ticked in the
 * Data validation panel. The script removes any type that doesn't belong to the row's subcategory.
 *
 * Work IDs are built automatically:
 *   CATEGORY - SUBCATEGORY - ENTITY - UNIQUE NUMBER - YEAR      e.g.  PE-CUR-BTR-0001-1931
 * The unique number (uid) never changes; the ID is rebuilt if category, client or year change.
 */

const CFG = {
  ORIGINALS_FOLDER_ID: '1QkSjkKvoqdWMH6tDBH9DpdkKFbdNuQp8',
  WEB_FOLDER_ID: '15yk25DcuLlHHCzliNBkYxDig8FbPDwHX',
  WEB_MAX_PX: 1600,
  THUMB_PX: 400,
  UID_DIGITS: 4,
  CACHE_SECONDS: 600,
  MAX_RUN_MS: 4.5 * 60 * 1000,
  DROPDOWN_ROWS_AHEAD: 100     // dropdowns are always ready this many rows past the last filled row
};

const TABS = {
  Works: ['work_id', 'uid', 'publish', 'title', 'year', 'country', 'entity_id', 'category', 'subcategory', 'types',
    'designer_id', 'intent', 'movement', 'design_style', 'typography_style', 'dimensions', 'credit', 'source_url', 'notes',
    'image_file', 'drive_full_id', 'drive_web_id', 'image_full_url', 'image_web_url', 'image_w', 'image_h', 'image_status'],
  Entities: ['entity_id', 'abbr', 'name', 'type', 'notes'],
  Designers: ['designer_id', 'name', 'kind', 'years_active', 'notes']
};

// Where each dependent list lives in _lists (row r of _lists = row r of Works)
const LISTS = {
  subcategory: { col: 1, width: 30 },    // A:AD  subcategories of the row's category
  design_style: { col: 61, width: 40 }   // BI:CV styles of the row's movement
};

/* ---------------- menu & edit trigger ---------------- */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Archive')
    .addItem('Rebuild all work IDs', 'assignIds')
    .addItem('Import new images as draft rows', 'importNewFiles')
    .addItem('Link images and make web copies', 'syncImages')
    .addItem('Run import + link now', 'syncAll')
    .addSeparator()
    .addItem('Repair / extend dropdowns', 'setupDropdowns')
    .addItem('Turn on automatic sync + dropdown repair', 'installTrigger')
    .addItem('Clear website cache', 'clearCache')
    .addToUi();
}

function onEdit(e) {
  clearCache();
  if (!e || !e.range) return;
  const sh = e.range.getSheet(), name = sh.getName();
  if (name === 'Works' && e.range.getLastRow() > 1) {
    const h = headerIndex_(sh);
    const top = Math.max(2, e.range.getRow()), bottom = e.range.getLastRow();
    // keep dependent choices consistent when the parent changes
    if (e.range.getNumRows() === 1 && e.range.getNumColumns() === 1) {
      const c = e.range.getColumn() - 1, row = top;
      if (c === h.category || c === h.subcategory) tidySolution_(sh, h, row);
      if (c === h.movement) tidyStyle_(sh, h, row);
      if (c === h.types) tidySolution_(sh, h, row, true);
    }
    // a row past the prepared dropdowns: prepare more
    if (!sh.getRange(bottom, h.subcategory + 1).getDataValidation()) setupDropdowns(bottom + CFG.DROPDOWN_ROWS_AHEAD);
    assignIds(top, bottom);
  }
  if (name === 'Entities' || name === 'Solutions' || name === 'Styles') assignIds();
}

function tidySolution_(sh, h, row, fromTypes) {
  const sol = readSolutions_();
  const cat = String(sh.getRange(row, h.category + 1).getValue()).trim();
  const subCell = sh.getRange(row, h.subcategory + 1), typCell = sh.getRange(row, h.types + 1);
  const sub = String(subCell.getValue()).trim();
  if (sub && sol.subCat[sub] !== cat) { subCell.clearContent(); typCell.clearContent(); return; }
  const allowed = sol.typesOf[sub] || [];
  const all = String(typCell.getValue()).split(',').map(t => t.trim()).filter(Boolean);
  const kept = all.filter(t => allowed.indexOf(t) >= 0), dropped = all.filter(t => allowed.indexOf(t) < 0);
  if (dropped.length) {
    typCell.setValue(kept.join(', '));
    SpreadsheetApp.getActive().toast('Removed ' + dropped.join(', ') + (sub ? ': not a type of ' + sub + ' in Solutions.' : ': choose a subcategory first.'), 'Types', 6);
  }
}
function tidyStyle_(sh, h, row) {
  const styles = readStyles_();
  const mv = String(sh.getRange(row, h.movement + 1).getValue()).trim();
  const cell = sh.getRange(row, h.design_style + 1), st = String(cell.getValue()).trim();
  if (st && (!styles[mv] || styles[mv].indexOf(st) < 0)) cell.clearContent();
}

/* ---------------- dependent dropdowns ---------------- */

/** Fills the _lists helper and sets one dropdown rule per row for subcategory and design_style, so each row only offers
 *  options for its own category / movement. The types column is NOT touched here: it has one whole-column rule
 *  with "Allow multiple selections" ticked by hand, which the script must not overwrite. */
function setupDropdowns(upTo) {
  const ss = SpreadsheetApp.getActive();
  const works = ss.getSheetByName('Works'), lists = ss.getSheetByName('_lists');
  const h = headerIndex_(works);
  const last = Math.min(works.getMaxRows(), Math.max(upTo || 0, works.getLastRow() + CFG.DROPDOWN_ROWS_AHEAD));
  if (works.getMaxRows() < last) works.insertRowsAfter(works.getMaxRows(), last - works.getMaxRows());
  if (lists.getMaxRows() < last) lists.insertRowsAfter(lists.getMaxRows(), last - lists.getMaxRows());
  const n = last - 1;
  // helper formulas look up their own row number, so deleting or inserting Works rows can't break them
  const F = {
    subcategory: '=LET(v,INDEX(Works!$H:$H,ROW()),IF(v="","",IFERROR(TRANSPOSE(UNIQUE(FILTER(Solutions!$C$2:$C,Solutions!$A$2:$A=v))),"")))',
    design_style: '=LET(v,INDEX(Works!$M:$M,ROW()),IF(v="","",IFERROR(TRANSPOSE(TOCOL(INDEX(Styles!$A$2:$Z,0,MATCH(v,Styles!$A$1:$Z$1,0)),1)),"")))'
  };
  lists.getRange(2, 1, lists.getMaxRows() - 1, 100).clearContent();
  Object.keys(LISTS).forEach(k => lists.getRange(2, LISTS[k].col, n, 1).setFormula(F[k]));
  lists.getRange('CW2').setFormula('=UNIQUE(FILTER(Solutions!$A$2:$A,Solutions!$A$2:$A<>""))');
  lists.getRange('CX2').setFormula('=SORT(UNIQUE(TOCOL(Styles!$A$2:$Z,1)))');
  // one rule per row
  Object.keys(LISTS).forEach(k => {
    const spec = LISTS[k], rules = [];
    for (let r = 2; r <= last; r++) {
      rules.push([SpreadsheetApp.newDataValidation()
        .requireValueInRange(lists.getRange(r, spec.col, 1, spec.width), true)
        .setAllowInvalid(false)
        .build()]);
    }
    works.getRange(2, h[k] + 1, n, 1).setDataValidations(rules);
  });
}

/* ---------------- IDs ---------------- */

function abbr_(name, len) {
  const words = String(name || '').replace(/[^A-Za-z0-9 ]/g, ' ').split(/\s+/)
    .filter(w => w && !/^(and|of|the|for|de|la|le|el|al)$/i.test(w));
  if (!words.length) return '';
  let a = words.length === 1 ? words[0].slice(0, len) : words.map(w => w[0]).join('').slice(0, len);
  if (a.length < len && words.length > 1) a = (a + words[words.length - 1].slice(1)).slice(0, len);
  return a.toUpperCase();
}
function makeId_(catCode, subCode, ent, uid, year) {
  if (!catCode || !subCode || !ent || !uid || !year) return '';
  return [catCode, subCode, ent, String(uid).padStart(CFG.UID_DIGITS, '0'), String(year).trim()].join('-');
}

/** Gives rows a uid, rebuilds work_id, and writes a warning if something doesn't match. */
function assignIds(fromRow, toRow) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Works');
  const h = headerIndex_(sh);
  const last = sh.getLastRow();
  if (last < 2) return;
  const data = sh.getRange(1, 1, last, sh.getLastColumn()).getValues();
  const sol = readSolutions_(), styles = readStyles_();
  const ents = {};
  const ev = ss.getSheetByName('Entities').getDataRange().getValues();
  const eh = {}; ev[0].forEach((k, i) => eh[k] = i);
  ev.slice(1).forEach(r => { if (r[eh.entity_id]) ents[r[eh.entity_id]] = String(r[eh.abbr] || abbr_(r[eh.name], 3)).toUpperCase(); });

  let maxUid = 0;
  data.slice(1).forEach(r => { maxUid = Math.max(maxUid, +r[h.uid] || 0); });
  const a = Math.max(2, fromRow || 2), b = Math.min(last, toRow || last);
  for (let i = a - 1; i <= b - 1; i++) {
    const r = data[i];
    const filled = [h.title, h.category, h.subcategory, h.entity_id, h.year, h.image_file].some(k => r[k] !== '');
    if (!filled) continue;
    let uid = +r[h.uid];
    if (!uid) { uid = ++maxUid; sh.getRange(i + 1, h.uid + 1).setValue(uid); }
    const cat = String(r[h.category]).trim(), sub = String(r[h.subcategory]).trim();
    const year = r[h.year] instanceof Date ? r[h.year].getFullYear() : r[h.year];
    const id = makeId_(sol.catCode[cat] || abbr_(cat, 3), sol.subCode[sub] || abbr_(sub, 3), ents[r[h.entity_id]], uid, year);
    if (id !== r[h.work_id]) sh.getRange(i + 1, h.work_id + 1).setValue(id);

    let warn = '';
    const mv = String(r[h.movement]).trim(), st = String(r[h.design_style]).trim();
    if (sub && sol.subCat[sub] && sol.subCat[sub] !== cat) warn = 'check: "' + sub + '" belongs to ' + sol.subCat[sub];
    else if (mv && st && styles[mv] && styles[mv].indexOf(st) < 0) warn = 'check: "' + st + '" is not listed under ' + mv + ' in Styles';
    else if (!id) warn = 'needs category, subcategory, entity and year for an ID';
    const status = String(r[h.image_status]);
    if (warn && status !== warn) sh.getRange(i + 1, h.image_status + 1).setValue(warn);
    if (!warn && /^(check:|needs category)/.test(status)) sh.getRange(i + 1, h.image_status + 1).setValue(r[h.drive_web_id] ? 'linked' : 'no image yet');
  }
}

/* ---------------- images ---------------- */

function importNewFiles() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Works');
  const h = headerIndex_(sh);
  const data = sh.getDataRange().getValues();
  const used = new Set();
  data.slice(1).forEach(r => {
    if (r[h.image_file]) used.add(String(r[h.image_file]));
    if (r[h.drive_full_id]) used.add('id:' + r[h.drive_full_id]);
    if (r[h.work_id]) used.add('wid:' + String(r[h.work_id]).toUpperCase());
  });
  const files = DriveApp.getFolderById(CFG.ORIGINALS_FOLDER_ID).getFiles();
  const add = [];
  while (files.hasNext()) {
    const f = files.next();
    if (!/^image\//.test(f.getMimeType())) continue;
    const name = f.getName();
    if (used.has(name) || used.has('id:' + f.getId()) || used.has('wid:' + idInName_(name))) continue;
    const row = new Array(data[0].length).fill('');
    row[h.publish] = false;
    row[h.title] = name.replace(/\.[a-z0-9]+$/i, '').replace(/[_-]+/g, ' ').trim();
    row[h.image_file] = name;
    row[h.image_status] = 'draft: fill in the record';
    add.push(row);
  }
  if (add.length) {
    const start = sh.getLastRow() + 1;
    sh.getRange(start, 1, add.length, add[0].length).setValues(add);
    if (!sh.getRange(start + add.length - 1, h.subcategory + 1).getDataValidation()) setupDropdowns();
    assignIds(start, start + add.length - 1);
  }
  return add.length;
}

function syncImages() {
  const started = Date.now();
  const sh = SpreadsheetApp.getActive().getSheetByName('Works');
  const h = headerIndex_(sh);
  const data = sh.getDataRange().getValues();
  const webFolder = DriveApp.getFolderById(CFG.WEB_FOLDER_ID);
  const byName = {}, byId = {};
  const it = DriveApp.getFolderById(CFG.ORIGINALS_FOLDER_ID).getFiles();
  while (it.hasNext()) {
    const f = it.next();
    byName[f.getName()] = f;
    const wid = idInName_(f.getName());
    if (wid && !byId[wid]) byId[wid] = f;
  }
  for (let i = 1; i < data.length; i++) {
    if (Date.now() - started > CFG.MAX_RUN_MS) break;
    const r = data[i];
    const uid = +r[h.uid];
    if (!uid) continue;
    if (r[h.drive_full_id] && r[h.drive_web_id]) continue;
    const set = (k, v) => sh.getRange(i + 1, h[k] + 1).setValue(v);
    try {
      const file = r[h.drive_full_id] ? DriveApp.getFileById(r[h.drive_full_id])
        : byName[String(r[h.image_file])] || byId[String(r[h.work_id]).toUpperCase()];
      if (!file) { if (r[h.work_id]) set('image_status', 'no image: name it "' + r[h.work_id] + ' ….jpg" or type its file name in image_file'); continue; }
      shareByLink_(file);
      const meta = Drive.Files.get(file.getId(), { fields: 'thumbnailLink,imageMediaMetadata(width,height)' });
      set('drive_full_id', file.getId());
      set('image_file', file.getName());
      set('image_full_url', 'https://drive.google.com/file/d/' + file.getId() + '/view');
      if (meta.imageMediaMetadata) { set('image_w', meta.imageMediaMetadata.width); set('image_h', meta.imageMediaMetadata.height); }
      if (!meta.thumbnailLink) { set('image_status', 'waiting for Drive to process the image, retry later'); continue; }
      const url = meta.thumbnailLink.replace(/=s\d+(-[a-z0-9-]+)?$/i, '') + '=s' + CFG.WEB_MAX_PX;
      const res = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
      if (res.getResponseCode() !== 200) { set('image_status', 'web copy failed (' + res.getResponseCode() + '), retry later'); continue; }
      const webName = String(uid).padStart(CFG.UID_DIGITS, '0') + '_web.jpg';
      const old = webFolder.getFilesByName(webName);
      while (old.hasNext()) old.next().setTrashed(true);
      const web = webFolder.createFile(res.getBlob().setName(webName));
      shareByLink_(web);
      set('drive_web_id', web.getId());
      set('image_web_url', 'https://lh3.googleusercontent.com/d/' + web.getId());
      set('image_status', 'linked ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')
        + ' · web ' + Math.round(web.getSize() / 1024) + ' KB / original ' + Math.round(file.getSize() / 1024) + ' KB');
    } catch (err) {
      set('image_status', 'error: ' + err.message);
    }
  }
  clearCache();
}

/** Rows inserted or deleted in Works shift the per-row dropdown rules, so rebuild them. */
function onStructureChange(e) {
  if (e && ['INSERT_ROW', 'REMOVE_ROW', 'INSERT_GRID', 'REMOVE_GRID'].indexOf(e.changeType) >= 0) setupDropdowns();
}

function syncAll() { assignIds(); importNewFiles(); syncImages(); }

function installTrigger() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'syncAll').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('syncAll').timeBased().everyMinutes(30).create();
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'onStructureChange').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('onStructureChange').forSpreadsheet(SpreadsheetApp.getActive()).onChange().create();
  SpreadsheetApp.getActive().toast('New images will be linked every 30 minutes.', 'Archive');
}

/* ---------------- reading the reference tabs ---------------- */

/** Solutions tab → codes, parent lookups and the tree for the website. */
function readSolutions_() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Solutions');
  return parseSolutions(sh ? sh.getDataRange().getValues() : [[]]);
}
function parseSolutions(v) {
  const head = v[0].map(x => String(x).trim()), ix = {};
  head.forEach((k, i) => ix[k] = i);
  const out = { catCode: {}, subCode: {}, subCat: {}, typesOf: {}, tree: [] };
  const catNode = {}, subNode = {};
  v.slice(1).forEach(r => {
    const cat = String(r[ix.category] || '').trim(), sub = String(r[ix.subcategory] || '').trim();
    if (!cat) return;
    const cc = String(r[ix.category_code] || '').trim().toUpperCase(), sc = String(r[ix.subcategory_code] || '').trim().toUpperCase();
    const typ = String(r[ix.type] || '').trim();
    if (cc && !out.catCode[cat]) out.catCode[cat] = cc;
    if (!catNode[cat]) { catNode[cat] = { name: cat, code: '', subs: [] }; out.tree.push(catNode[cat]); }
    if (!sub) return;
    if (sc && !out.subCode[sub]) out.subCode[sub] = sc;
    out.subCat[sub] = out.subCat[sub] || cat;
    if (!subNode[sub]) { subNode[sub] = { name: sub, code: '', types: [] }; catNode[cat].subs.push(subNode[sub]); }
    out.typesOf[sub] = out.typesOf[sub] || [];
    if (typ && out.typesOf[sub].indexOf(typ) < 0) { out.typesOf[sub].push(typ); subNode[sub].types.push(typ); }
  });
  out.tree.forEach(c => { c.code = out.catCode[c.name] || abbr_(c.name, 3); c.subs.forEach(s => s.code = out.subCode[s.name] || abbr_(s.name, 3)); });
  return out;
}

/** Styles tab: one column per movement, header = movement name. Returns { movement: [styles] }. */
function readStyles_() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Styles');
  if (!sh || sh.getLastRow() < 1) return {};
  return parseStyles(sh.getDataRange().getValues());
}
function parseStyles(v) {
  const out = {};
  v[0].forEach((m, c) => { if (String(m).trim()) out[String(m).trim()] = v.slice(1).map(r => String(r[c]).trim()).filter(Boolean); });
  return out;
}

/* ---------------- JSON feed ---------------- */

function doGet(e) {
  const cache = CacheService.getScriptCache();
  let json = (e && e.parameter && e.parameter.refresh) ? null : readCache_(cache);
  if (!json) { json = JSON.stringify(buildPayload(readTables_())); writeCache_(cache, json); }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function readTables_() {
  const ss = SpreadsheetApp.getActive(), t = {};
  Object.keys(TABS).forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) { t[name] = []; return; }
    const v = sh.getDataRange().getValues(), head = v[0].map(String);
    t[name] = v.slice(1).filter(r => r.some(c => c !== '')).map(r => {
      const o = {}; head.forEach((k, i) => o[k] = r[i] instanceof Date ? r[i].getFullYear() : r[i]); return o;
    });
  });
  t.Solutions = readSolutions_().tree;
  t.Styles = readStyles_();
  return t;
}

/** Pure: sheet rows in, site payload out. */
function buildPayload(t) {
  const s = v => (v === null || v === undefined) ? '' : String(v).trim();
  const thumb = id => id ? 'https://lh3.googleusercontent.com/d/' + id + '=w' + CFG.THUMB_PX : '';
  const works = (t.Works || []).filter(r => r.publish === true || s(r.publish).toUpperCase() === 'TRUE').map(r => ({
    id: s(r.work_id), uid: +r.uid || 0, title: s(r.title), year: s(r.year), country: s(r.country),
    entity: s(r.entity_id), designer: s(r.designer_id), intent: s(r.intent),
    category: s(r.category), subcategory: s(r.subcategory),
    types: s(r.types).split(',').map(x => x.trim()).filter(Boolean),
    movement: s(r.movement), designStyle: s(r.design_style), typeStyle: s(r.typography_style),
    dims: s(r.dimensions), credit: s(r.credit), source: s(r.source_url), notes: s(r.notes),
    img: { web: s(r.image_web_url), full: s(r.image_full_url), thumb: thumb(s(r.drive_web_id)), w: +r.image_w || 0, h: +r.image_h || 0 }
  })).filter(w => w.id && w.category && w.subcategory);
  return {
    version: 3,
    updated: new Date().toISOString(),
    structure: t.Solutions || [],
    styles: t.Styles || {},
    entities: (t.Entities || []).map(r => ({ id: s(r.entity_id), abbr: s(r.abbr), name: s(r.name), type: s(r.type), notes: s(r.notes) })).filter(x => x.id),
    designers: (t.Designers || []).map(r => ({ id: s(r.designer_id), name: s(r.name), kind: s(r.kind), years: s(r.years_active), notes: s(r.notes) })).filter(x => x.id),
    works
  };
}

/* ---------------- helpers ---------------- */

function headerIndex_(sh) {
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const o = {}; head.forEach((k, i) => o[String(k)] = i); return o;
}
function idInName_(name) {
  const m = String(name).match(/^([A-Z0-9]{2,4}-[A-Z0-9]{2,4}-[A-Z0-9]{2,6}-\d{3,}-\d{4})/i);
  return m ? m[1].toUpperCase() : '';
}
function shareByLink_(file) {
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }
  catch (e) { /* some Workspace domains forbid public links; the image will not show on the site */ }
}
function writeCache_(cache, json) {
  const size = 90000, n = Math.ceil(json.length / size), parts = {};
  if (n > 50) return;
  for (let i = 0; i < n; i++) parts['feed_' + i] = json.slice(i * size, (i + 1) * size);
  parts.feed_n = String(n);
  cache.putAll(parts, CFG.CACHE_SECONDS);
}
function readCache_(cache) {
  const n = +cache.get('feed_n');
  if (!n) return null;
  const keys = Array.from({ length: n }, (_, i) => 'feed_' + i), got = cache.getAll(keys);
  return keys.some(k => !got[k]) ? null : keys.map(k => got[k]).join('');
}
function clearCache() { CacheService.getScriptCache().remove('feed_n'); }

if (typeof module !== 'undefined') module.exports = { buildPayload, parseSolutions, parseStyles, makeId_, abbr_, idInName_ };
