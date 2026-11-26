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
 *   Images     one row per image, linked to its work by uid. A work can have any number; `order` sets the sequence.
 *   Solutions  category › subcategory › type (one type per row), with the codes used in IDs. Add a row to add an option.
 *              A work can take several types: comma-separated in Works › types.
 *   Styles     one column per movement (header = movement), styles listed below;
 *              plus a "Typography" column (with optional descriptions in the column to its right)
 *   _lists     hidden helper that feeds the per-row dropdowns (subcategory, design_style). Don't edit.
 *
 * Works › types and Works › designer_id each have ONE whole-column rule with "Allow multiple selections"
 * ticked in the Data validation panel (several values, comma-separated). The script removes any type that doesn't belong to the row's subcategory.
 *
 * Work IDs are built automatically:
 *   CATEGORY - SUBCATEGORY - ENTITY - UNIQUE NUMBER - YEAR      e.g.  PE-CUR-BTR-0001-1931
 * The unique number (uid) never changes; the ID is rebuilt if category, client or year change.
 */

const CFG = {
  ORIGINALS_FOLDER_ID: '1QkSjkKvoqdWMH6tDBH9DpdkKFbdNuQp8',
  WEB_FOLDER_ID: '15yk25DcuLlHHCzliNBkYxDig8FbPDwHX',
  WEB_MAX_PX: 1600,
  // Format of the web copy: 'same' keeps the original's format (PNG stays PNG, JPEG stays JPEG),
  // or force one: 'webp' (smallest, keeps transparency), 'jpeg', 'png'.
  WEB_FORMAT: 'same',
  THUMB_PX: 400,
  UID_DIGITS: 4,
  CACHE_SECONDS: 600,
  MAX_RUN_MS: 4.5 * 60 * 1000,
  DROPDOWN_ROWS_AHEAD: 100     // dropdowns are always ready this many rows past the last filled row
};

const TABS = {
  Works: ['work_id', 'uid', 'publish', 'title', 'year', 'country', 'entity_id', 'category', 'subcategory', 'types',
    'designer_id', 'intent', 'movement', 'design_style', 'typography_style', 'dimensions', 'credit', 'source_url', 'notes',
    'status'],
  Images: ['uid', 'work_id', 'order', 'image_file', 'caption', 'drive_full_id', 'drive_web_id', 'image_full_url',
    'image_web_url', 'image_w', 'image_h', 'status'],
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
    .addItem('Remake web copies for selected rows', 'remakeWebCopies')
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
  if (name === 'Images') syncImageIds_();
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
  // rows really in use (the publish checkboxes fill every row, so getLastRow() can't be trusted)
  const last = Math.max(upTo || 0, lastWorkRow_(works, h) + CFG.DROPDOWN_ROWS_AHEAD);
  if (works.getMaxRows() < last) works.insertRowsAfter(works.getMaxRows(), last - works.getMaxRows());
  if (lists.getMaxRows() < works.getMaxRows()) lists.insertRowsAfter(lists.getMaxRows(), works.getMaxRows() - lists.getMaxRows());
  const n = last - 1, all = lists.getMaxRows() - 1;

  // helper lists: every _lists row looks up the same row number in Works
  const F = {
    subcategory: '=LET(v,INDEX(Works!$H:$H,ROW()),IF(v="","",IFERROR(TRANSPOSE(UNIQUE(FILTER(Solutions!$C$2:$C,Solutions!$A$2:$A=v))),"")))',
    design_style: '=LET(v,INDEX(Works!$M:$M,ROW()),IF(v="","",IFERROR(TRANSPOSE(TOCOL(INDEX(Styles!$A$2:$Z,0,MATCH(v,Styles!$A$1:$Z$1,0)),1)),"")))'
  };
  lists.getRange(2, 1, all, 100).clearContent();
  Object.keys(LISTS).forEach(k => lists.getRange(2, LISTS[k].col, all, 1).setFormula(F[k]));
  lists.getRange(2, 101, all, 3).clearContent();
  lists.getRange('CW2').setFormula('=UNIQUE(FILTER(Solutions!$A$2:$A,Solutions!$A$2:$A<>""))');
  lists.getRange('CX2').setFormula('=IFERROR(TOCOL(INDEX(Styles!$A$2:$Z,0,MATCH("Typography",Styles!$A$1:$Z$1,0)),1),"")');
  lists.getRange('CY2').setFormula('=TOCOL(FILTER(Styles!$A$1:$Z$1,Styles!$A$1:$Z$1<>"",LEFT(Styles!$A$1:$Z$1,10)<>"Typography"),1)');

  // per-row rules: subcategory follows the row's category, design_style follows its movement
  Object.keys(LISTS).forEach(k => {
    const spec = LISTS[k], rules = [];
    for (let r = 2; r <= last; r++) {
      rules.push([SpreadsheetApp.newDataValidation()
        .requireValueInRange(lists.getRange(r, spec.col, 1, spec.width), true)
        .setAllowInvalid(false).build()]);
    }
    works.getRange(2, h[k] + 1, n, 1).setDataValidations(rules);
  });

  // whole-column rules. types and designer_id are left alone: their "Allow multiple selections" setting
  // can only be made by hand in the Data validation panel, and rewriting the rule here would remove it.
  const col = (k, range, strict) => works.getRange(2, h[k] + 1, works.getMaxRows() - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInRange(range, true).setAllowInvalid(!strict).build());
  col('category', lists.getRange('CW2:CW200'), true);
  col('movement', lists.getRange('CY2:CY50'), true);
  col('typography_style', lists.getRange('CX2:CX500'), true);
  col('entity_id', ss.getSheetByName('Entities').getRange('A2:A500'), true);
  ss.toast('Dropdowns ready for rows 2–' + last + '.', 'Archive', 5);
}

/** "MEDICOLT_VINYLCOV_SONO_1960_3.png" and "…_1960_1.png" share the key "medicolt_vinylcov_sono_1960":
 *  the trailing number (and the extension) is dropped, so numbered shots of one object stay together. */
function groupKey_(name) {
  return String(name).replace(/\.[a-z0-9]+$/i, '').replace(/(\s*copy)?([ _-]*\(?\d{1,3}\)?)$/i, '').trim().toLowerCase();
}

/** Last Works row that holds an actual record. */
function lastWorkRow_(sh, h) {
  const n = sh.getMaxRows() - 1;
  if (n < 1) return 1;
  const v = sh.getRange(2, 1, n, sh.getLastColumn()).getValues();
  const keys = ['work_id', 'title', 'year', 'entity_id', 'category', 'subcategory'].map(k => h[k]);
  for (let i = v.length - 1; i >= 0; i--) if (keys.some(k => v[i][k] !== '' && v[i][k] !== null)) return i + 2;
  return 1;
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
  const sol = readSolutions_(), styles = readStyles_(), imgCount = countImages_();
  const ents = {};
  const ev = ss.getSheetByName('Entities').getDataRange().getValues();
  const eh = {}; ev[0].forEach((k, i) => eh[k] = i);
  ev.slice(1).forEach(r => { if (r[eh.entity_id]) ents[r[eh.entity_id]] = String(r[eh.abbr] || abbr_(r[eh.name], 3)).toUpperCase(); });

  let maxUid = 0;
  data.slice(1).forEach(r => { maxUid = Math.max(maxUid, +r[h.uid] || 0); });
  const a = Math.max(2, fromRow || 2), b = Math.min(last, toRow || last);
  for (let i = a - 1; i <= b - 1; i++) {
    const r = data[i];
    const filled = [h.title, h.category, h.subcategory, h.entity_id, h.year].some(k => r[k] !== '');
    if (!filled) continue;
    let uid = +r[h.uid];
    if (!uid) { uid = ++maxUid; sh.getRange(i + 1, h.uid + 1).setValue(uid); }
    const cat = String(r[h.category]).trim(), sub = String(r[h.subcategory]).trim();
    const year = r[h.year] instanceof Date ? r[h.year].getFullYear() : r[h.year];
    const id = makeId_(sol.catCode[cat] || abbr_(cat, 3), sol.subCode[sub] || abbr_(sub, 3), ents[r[h.entity_id]], uid, year);
    if (id !== r[h.work_id]) sh.getRange(i + 1, h.work_id + 1).setValue(id);

    const mv = String(r[h.movement]).trim(), st = String(r[h.design_style]).trim();
    let status;
    if (sub && sol.subCat[sub] && sol.subCat[sub] !== cat) status = 'check: "' + sub + '" belongs to ' + sol.subCat[sub];
    else if (mv && st && styles[mv] && styles[mv].indexOf(st) < 0) status = 'check: "' + st + '" is not listed under ' + mv + ' in Styles';
    else if (!id) status = 'needs category, subcategory, entity and year for an ID';
    else { const n = imgCount[uid] || 0; status = n ? n + (n > 1 ? ' images' : ' image') : 'no images yet'; }
    if (String(r[h.status]) !== status) sh.getRange(i + 1, h.status + 1).setValue(status);
  }
  syncImageIds_();
}

/** uid → number of linked images */
function countImages_() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Images'), out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  const h = headerIndex_(sh);
  sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues()
    .forEach(r => { if (+r[h.uid] && r[h.drive_web_id]) out[+r[h.uid]] = (out[+r[h.uid]] || 0) + 1; });
  return out;
}

/** Keeps the Images tab in step with Works.
 *  - a work_id you type or pick wins: the row's uid is set to that work's uid
 *  - if a work's ID changed (so the old ID no longer exists), the uid brings the new ID back
 *  - a work_id is never blanked out */
function syncImageIds_() {
  const ss = SpreadsheetApp.getActive(), im = ss.getSheetByName('Images'), wk = ss.getSheetByName('Works');
  if (!im || im.getLastRow() < 2) return;
  const wh = headerIndex_(wk), ih = headerIndex_(im);
  const byUid = {}, byId = {};
  wk.getRange(2, 1, Math.max(wk.getLastRow() - 1, 1), wk.getLastColumn()).getValues().forEach(r => {
    const u = +r[wh.uid], id = String(r[wh.work_id]).trim();
    if (u) { byUid[u] = id; if (id) byId[id.toUpperCase()] = u; }
  });
  const v = im.getRange(2, 1, im.getLastRow() - 1, im.getLastColumn()).getValues();
  v.forEach((r, i) => {
    const row = i + 2, uid = +r[ih.uid], typed = String(r[ih.work_id]).trim();
    const typedUid = typed ? byId[typed.toUpperCase()] : 0;
    if (typedUid) {
      if (typedUid !== uid) im.getRange(row, ih.uid + 1).setValue(typedUid);
    } else if (uid && byUid[uid]) {
      if (byUid[uid] !== typed) im.getRange(row, ih.work_id + 1).setValue(byUid[uid]);
    } else if (typed) {
      im.getRange(row, ih.status + 1).setValue('work_id "' + typed + '" not found in Works');
    }
  });
}

/* ---------------- images ---------------- */

function importNewFiles() {
  const ss = SpreadsheetApp.getActive();
  const wk = ss.getSheetByName('Works'), im = ss.getSheetByName('Images');
  const wh = headerIndex_(wk), ih = headerIndex_(im);
  const works = wk.getRange(1, 1, Math.max(wk.getLastRow(), 1), wk.getLastColumn()).getValues();
  const imgs = im.getLastRow() > 1 ? im.getRange(2, 1, im.getLastRow() - 1, im.getLastColumn()).getValues() : [];
  const used = new Set(), uidById = {}, nextOrder = {}, uidByGroup = {};
  imgs.forEach(r => {
    if (r[ih.image_file]) used.add(String(r[ih.image_file]));
    if (r[ih.image_file] && +r[ih.uid]) uidByGroup[groupKey_(r[ih.image_file])] = +r[ih.uid];
    if (r[ih.drive_full_id]) used.add('id:' + r[ih.drive_full_id]);
    const u = +r[ih.uid]; if (u) nextOrder[u] = Math.max(nextOrder[u] || 1, (+r[ih.order] || 0) + 1);
  });
  let maxUid = 0;
  works.slice(1).forEach(r => {
    maxUid = Math.max(maxUid, +r[wh.uid] || 0);
    if (r[wh.work_id] && +r[wh.uid]) uidById[String(r[wh.work_id]).toUpperCase()] = +r[wh.uid];
  });
  // new files, in natural name order (…_1, …_2, …_10), so image order follows the file numbers
  const found = [], it = DriveApp.getFolderById(CFG.ORIGINALS_FOLDER_ID).getFiles();
  while (it.hasNext()) {
    const f = it.next();
    if (/^image\//.test(f.getMimeType()) && !used.has(f.getName()) && !used.has('id:' + f.getId())) found.push(f);
  }
  found.sort((a, b) => a.getName().localeCompare(b.getName(), undefined, { numeric: true, sensitivity: 'base' }));
  const newWorks = [], newImages = [];
  for (const f of found) {
    const name = f.getName(), group = groupKey_(name);
    // 1) name starts with a work ID  2) a sibling file (same name, other number) already belongs to a work
    let uid = uidById[idInName_(name)] || uidByGroup[group];
    if (!uid) {                                   // no work yet: start a draft work for it
      uid = ++maxUid;
      const row = new Array(works[0].length).fill('');
      row[wh.uid] = uid; row[wh.publish] = false;
      row[wh.title] = name.replace(/\.[a-z0-9]+$/i, '').replace(/[_-]+/g, ' ').trim();
      row[wh.status] = 'draft: fill in the record';
      newWorks.push(row);
    }
    uidByGroup[group] = uid;
    const order = nextOrder[uid] || 1; nextOrder[uid] = order + 1;
    const ir = new Array(TABS.Images.length).fill('');
    ir[ih.uid] = uid; ir[ih.order] = order; ir[ih.image_file] = name; ir[ih.status] = 'waiting for sync';
    newImages.push(ir);
  }
  if (newWorks.length) {
    const start = lastWorkRow_(wk, wh) + 1;                   // right under the last record, not after the empty checkbox rows
    wk.getRange(start, 1, newWorks.length, newWorks[0].length).setValues(newWorks);
    if (!wk.getRange(start + newWorks.length - 1, wh.subcategory + 1).getDataValidation()) setupDropdowns();
  }
  if (newImages.length) im.getRange(im.getLastRow() + 1, 1, newImages.length, newImages[0].length).setValues(newImages);
  return newImages.length;
}

function syncImages() {
  const started = Date.now();
  const im = SpreadsheetApp.getActive().getSheetByName('Images');
  if (!im || im.getLastRow() < 2) return;
  syncImageIds_();
  const h = headerIndex_(im);
  const data = im.getRange(1, 1, im.getLastRow(), im.getLastColumn()).getValues();
  const webFolder = DriveApp.getFolderById(CFG.WEB_FOLDER_ID);
  const byName = {};
  const it = DriveApp.getFolderById(CFG.ORIGINALS_FOLDER_ID).getFiles();
  while (it.hasNext()) { const f = it.next(); byName[f.getName()] = f; }

  for (let i = 1; i < data.length; i++) {
    if (Date.now() - started > CFG.MAX_RUN_MS) break;          // the trigger continues next run
    const r = data[i];
    const uid = +r[h.uid];
    if (!uid || (r[h.drive_full_id] && r[h.drive_web_id])) continue;
    const set = (k, v) => im.getRange(i + 1, h[k] + 1).setValue(v);
    try {
      const file = r[h.drive_full_id] ? DriveApp.getFileById(r[h.drive_full_id]) : byName[String(r[h.image_file])];
      if (!file) { set('status', 'file not found in Originals: check the name in image_file'); continue; }
      shareByLink_(file);
      const meta = Drive.Files.get(file.getId(), { fields: 'thumbnailLink,imageMediaMetadata(width,height)' });
      set('drive_full_id', file.getId());
      set('image_file', file.getName());
      set('image_full_url', 'https://drive.google.com/file/d/' + file.getId() + '/view');
      if (meta.imageMediaMetadata) { set('image_w', meta.imageMediaMetadata.width); set('image_h', meta.imageMediaMetadata.height); }
      if (!meta.thumbnailLink) { set('status', 'waiting for Drive to process the image, retry later'); continue; }
      const url = meta.thumbnailLink.replace(/=s\d+(-[a-z0-9-]+)?$/i, '') + '=s' + CFG.WEB_MAX_PX + formatFlag_(file.getMimeType());
      const res = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
      if (res.getResponseCode() !== 200) { set('status', 'web copy failed (' + res.getResponseCode() + '), retry later'); continue; }
      const blob = res.getBlob();
      const ext = { 'image/png': 'png', 'image/webp': 'webp', 'image/gif': 'gif' }[blob.getContentType()] || 'jpg';
      // named by uid and position, so it survives work ID changes: 0004-02_web.png
      const base = String(uid).padStart(CFG.UID_DIGITS, '0') + '-' + String(+r[h.order] || i).padStart(2, '0') + '_web';
      if (r[h.drive_web_id]) { try { DriveApp.getFileById(r[h.drive_web_id]).setTrashed(true); } catch (e) {} }
      const web = webFolder.createFile(blob.setName(base + '.' + ext));
      shareByLink_(web);
      set('drive_web_id', web.getId());
      set('image_web_url', 'https://lh3.googleusercontent.com/d/' + web.getId());
      set('status', 'linked ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')
        + ' · web ' + ext.toUpperCase() + ' ' + Math.round(web.getSize() / 1024) + ' KB / original ' + Math.round(file.getSize() / 1024) + ' KB');
    } catch (err) {
      set('status', 'error: ' + err.message);
    }
  }
  assignIds();          // refreshes the image counts on Works
  clearCache();
}

/** Rows inserted or deleted in Works shift the per-row dropdown rules, so rebuild them. */
function onStructureChange(e) {
  if (e && ['INSERT_ROW', 'REMOVE_ROW', 'INSERT_GRID', 'REMOVE_GRID'].indexOf(e.changeType) >= 0) setupDropdowns();
}

/** Google's image service picks the output format from a flag on the URL: -rp PNG, -rw WebP, -rj JPEG, -rg GIF. */
function formatFlag_(originalMime) {
  const f = CFG.WEB_FORMAT;
  if (f === 'webp') return '-rw';
  if (f === 'jpeg') return '-rj';
  if (f === 'png') return '-rp';
  return { 'image/png': '-rp', 'image/webp': '-rw', 'image/gif': '-rg' }[originalMime] || '-rj';
}

/** Makes the web copies again for the selected rows: rows on the Images tab, or every image of the selected works. */
function remakeWebCopies() {
  const ss = SpreadsheetApp.getActive(), sh = ss.getActiveSheet(), r = sh.getActiveRange();
  const im = ss.getSheetByName('Images'), ih = headerIndex_(im);
  if (im.getLastRow() < 2) return;
  const top = Math.max(2, r.getRow()), bottom = r.getLastRow();
  let rows = [];
  if (sh.getName() === 'Images') for (let i = top; i <= bottom; i++) rows.push(i);
  else if (sh.getName() === 'Works') {
    const wh = headerIndex_(sh);
    const uids = sh.getRange(top, wh.uid + 1, bottom - top + 1, 1).getValues().map(x => +x[0]).filter(Boolean);
    im.getRange(2, ih.uid + 1, im.getLastRow() - 1, 1).getValues().forEach((x, i) => { if (uids.indexOf(+x[0]) >= 0) rows.push(i + 2); });
  } else { ss.toast('Select rows on the Works or Images tab first.', 'Archive'); return; }
  rows.forEach(i => {
    const cell = im.getRange(i, ih.drive_web_id + 1), old = cell.getValue();
    if (old) { try { DriveApp.getFileById(old).setTrashed(true); } catch (e) {} }
    cell.clearContent();
  });
  syncImages();
  ss.toast('Web copies remade for ' + rows.length + ' image(s).', 'Archive');
}

function syncAll() { importNewFiles(); syncImages(); assignIds(); }

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
  v[0].forEach((m, c) => {
    const name = String(m).trim();
    if (name && !/^typography/i.test(name)) out[name] = v.slice(1).map(r => String(r[c]).trim()).filter(Boolean);
  });
  return out;
}
/** The Typography column of Styles, with the description from the column to its right. */
function parseTypography(v) {
  const c = v[0].findIndex(m => /^typography$/i.test(String(m).trim()));
  if (c < 0) return [];
  return v.slice(1).map(r => ({ name: String(r[c] || '').trim(), note: String(r[c + 1] || '').trim() })).filter(t => t.name);
}
function readTypography_() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Styles');
  return sh && sh.getLastRow() ? parseTypography(sh.getDataRange().getValues()) : [];
}

/* ---------------- JSON feed ---------------- */

function doGet(e) {
  const cache = CacheService.getScriptCache();
  let json = (e && e.parameter && e.parameter.refresh) ? null : readCache_(cache);
  if (!json) { json = JSON.stringify(buildPayload(readTables_())); writeCache_(cache, json); }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

/** Contact form on the website → a row on the Messages tab. Bots that fill the hidden "website" field are ignored. */
function doPost(e) {
  const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (d.kind !== 'contact') return out({ ok: false, error: 'unknown request' });
    if (d.website) return out({ ok: true });                                   // honeypot: pretend success
    const clip = (v, n) => String(v || '').replace(/^[=+\-@]/, "'$&").slice(0, n).trim();   // no formulas from strangers
    const name = clip(d.name, 120), email = clip(d.email, 160), topic = clip(d.topic, 80), msg = clip(d.message, 4000);
    if (!name || !/^\S+@\S+\.\S+$/.test(email) || !msg) return out({ ok: false, error: 'missing fields' });
    const lock = LockService.getScriptLock(); lock.waitLock(10000);
    const ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName('Messages');
    if (!sh) {
      sh = ss.insertSheet('Messages');
      sh.getRange(1, 1, 1, 6).setValues([['received', 'name', 'email', 'about', 'message', 'status']]).setFontWeight('bold').setBackground('#f2d54a');
      sh.setFrozenRows(1);
    }
    sh.appendRow([new Date(), name, email, topic, msg, 'new']);
    lock.releaseLock();
    return out({ ok: true });
  } catch (err) {
    return out({ ok: false, error: 'server error' });
  }
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
  t.Typography = readTypography_();
  return t;
}

/** Pure: sheet rows in, site payload out. */
function buildPayload(t) {
  const s = v => (v === null || v === undefined) ? '' : String(v).trim();
  const thumb = id => id ? 'https://lh3.googleusercontent.com/d/' + id + '=w' + CFG.THUMB_PX : '';
  const imgs = {};
  (t.Images || []).filter(r => +r.uid && s(r.image_web_url))
    .sort((a, b) => (+a.order || 0) - (+b.order || 0))
    .forEach(r => (imgs[+r.uid] = imgs[+r.uid] || []).push({
      web: s(r.image_web_url), full: s(r.image_full_url), thumb: thumb(s(r.drive_web_id)),
      w: +r.image_w || 0, h: +r.image_h || 0, caption: s(r.caption)
    }));
  const works = (t.Works || []).filter(r => r.publish === true || s(r.publish).toUpperCase() === 'TRUE').map(r => ({
    id: s(r.work_id), uid: +r.uid || 0, title: s(r.title), year: s(r.year), country: s(r.country),
    entity: s(r.entity_id), intent: s(r.intent),
    designers: s(r.designer_id).split(',').map(x => x.trim()).filter(Boolean),
    category: s(r.category), subcategory: s(r.subcategory),
    types: s(r.types).split(',').map(x => x.trim()).filter(Boolean),
    movement: s(r.movement), designStyle: s(r.design_style), typeStyle: s(r.typography_style),
    dims: s(r.dimensions), credit: s(r.credit), source: s(r.source_url), notes: s(r.notes),
    images: (imgs[+r.uid] || [])
  })).filter(w => w.id && w.category && w.subcategory);
  return {
    version: 5,
    updated: new Date().toISOString(),
    structure: t.Solutions || [],
    styles: t.Styles || {},
    typography: t.Typography || [],
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

if (typeof module !== 'undefined') module.exports = { groupKey_, buildPayload, parseSolutions, parseStyles, parseTypography, makeId_, abbr_, idInName_ };
