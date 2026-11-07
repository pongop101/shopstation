/**
 * ShopStation — หลังบ้าน (Google Apps Script + Google Sheets + Google Drive)
 * วิธีใช้: ดูไฟล์ "คู่มือติดตั้ง ShopStation"
 *  1) สร้าง Google Sheets ใหม่ → ส่วนขยาย → Apps Script → วางโค้ดนี้ทั้งหมด
 *  2) เลือกฟังก์ชัน setup แล้วกด "เรียกใช้" (อนุญาตสิทธิ์) → ดูรหัส API Key ในบันทึกการดำเนินการ
 *  3) การทำให้ใช้งานได้ → การทำให้ใช้งานได้รายการใหม่ → เว็บแอป
 *     ดำเนินการในฐานะ: ฉัน  |  ผู้ที่มีสิทธิ์เข้าถึง: ทุกคน
 */
const SHEET_STORES = 'Stores';
const SHEET_INSP = 'Inspections';
const PHOTO_FOLDER = 'ShopStation Photos';

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [SHEET_STORES, SHEET_INSP].forEach(n => getSheet_(n));
  const props = PropertiesService.getScriptProperties();
  let key = props.getProperty('API_KEY');
  if (!key) {
    key = Utilities.getUuid().replace(/-/g, '').slice(0, 12);
    props.setProperty('API_KEY', key);
  }
  getFolder_();
  Logger.log('ตั้งค่าเรียบร้อย  ไฟล์: ' + ss.getName());
  Logger.log('API Key (นำไปใส่ในแอป หน้า ตั้งค่า): ' + key);
}

/** เปลี่ยนรหัส API Key ใหม่ (ทุกเครื่องต้องใส่รหัสใหม่) */
function resetApiKey() {
  const key = Utilities.getUuid().replace(/-/g, '').slice(0, 12);
  PropertiesService.getScriptProperties().setProperty('API_KEY', key);
  Logger.log('API Key ใหม่: ' + key);
}

function doGet() {
  return json_({ ok: true, app: 'ShopStation', message: 'ShopStation API พร้อมใช้งาน' });
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    const key = PropertiesService.getScriptProperties().getProperty('API_KEY');
    if (key && req.key !== key) throw new Error('รหัส API Key ไม่ถูกต้อง');
    switch (req.action) {
      case 'ping': return json_({ ok: true, sheet: SpreadsheetApp.getActiveSpreadsheet().getName() });
      case 'list': return json_({ ok: true, stores: readAll_(SHEET_STORES), inspections: readAll_(SHEET_INSP) });
      case 'saveStore': return json_(withLock_(() => upsert_(SHEET_STORES, req.record, req.flat)));
      case 'saveInspection': return json_(withLock_(() => upsert_(SHEET_INSP, req.record, req.flat)));
      case 'deleteStore': return json_(withLock_(() => deleteStore_(req.id)));
      case 'deleteInspection': return json_(withLock_(() => deleteRows_(SHEET_INSP, r => r.id === req.id)));
      case 'uploadPhoto': return json_(uploadPhoto_(req));
      default: throw new Error('ไม่รู้จักคำสั่ง ' + req.action);
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

/* ---------------- helpers ---------------- */
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function getSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, 3).setValues([['id', 'updated', 'json']]);
    sh.setFrozenRows(1);
    sh.getRange('1:1').setFontWeight('bold').setBackground('#ccfbf1');
  }
  return sh;
}
function headers_(sh) {
  const lc = sh.getLastColumn();
  return lc ? sh.getRange(1, 1, 1, lc).getValues()[0].map(String) : [];
}
function readAll_(name) {
  const sh = getSheet_(name);
  const lr = sh.getLastRow();
  if (lr < 2) return [];
  const h = headers_(sh);
  const jc = h.indexOf('json');
  if (jc < 0) return [];
  return sh.getRange(2, jc + 1, lr - 1, 1).getValues()
    .map(r => { try { return JSON.parse(r[0]); } catch (e) { return null; } })
    .filter(Boolean);
}
function upsert_(name, rec, flat) {
  if (!rec || !rec.id) throw new Error('ข้อมูลไม่ครบ (ไม่มี id)');
  const sh = getSheet_(name);
  const row = Object.assign({ id: rec.id, updated: rec.updated || Date.now() }, flat || {});
  row['แก้ไขล่าสุด'] = new Date(rec.updated || Date.now());
  row.json = JSON.stringify(rec);
  const h = headers_(sh);
  let changed = false;
  Object.keys(row).forEach(k => { if (h.indexOf(k) < 0) { h.push(k); changed = true; } });
  if (changed) {
    sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#ccfbf1');
  }
  const values = h.map(k => {
    const v = row[k];
    if (v === undefined || v === null) return '';
    if (typeof v === 'object' && !(v instanceof Date)) return JSON.stringify(v);
    if (typeof v === 'string' && /^[=+\-@]/.test(v) && k !== 'json') return "'" + v; // กันสูตร
    return v;
  });
  const lr = sh.getLastRow();
  const ids = lr > 1 ? sh.getRange(2, 1, lr - 1, 2).getValues() : [];
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(rec.id)) {
      if (Number(ids[i][1]) > Number(rec.updated || 0)) return { ok: true, skipped: 'server newer' };
      sh.getRange(i + 2, 1, 1, h.length).setValues([values]);
      return { ok: true, updated: true };
    }
  }
  sh.getRange(lr + 1, 1, 1, h.length).setValues([values]);
  return { ok: true, created: true };
}
function deleteRows_(name, test) {
  const sh = getSheet_(name);
  const lr = sh.getLastRow();
  if (lr < 2) return { ok: true, deleted: 0 };
  const h = headers_(sh);
  const jc = h.indexOf('json');
  const data = sh.getRange(2, 1, lr - 1, h.length).getValues();
  let n = 0;
  for (let i = data.length - 1; i >= 0; i--) {
    let rec = { id: String(data[i][0]) };
    try { rec = JSON.parse(data[i][jc]); } catch (e) { }
    if (test(rec)) { sh.deleteRow(i + 2); n++; }
  }
  return { ok: true, deleted: n };
}
function deleteStore_(id) {
  deleteRows_(SHEET_INSP, r => r.storeId === id);
  return deleteRows_(SHEET_STORES, r => r.id === id);
}
function getFolder_() {
  const props = PropertiesService.getScriptProperties();
  const fid = props.getProperty('PHOTO_FOLDER_ID');
  if (fid) { try { return DriveApp.getFolderById(fid); } catch (e) { } }
  const ssFile = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  const parents = ssFile.getParents();
  const parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const f = parent.createFolder(PHOTO_FOLDER);
  props.setProperty('PHOTO_FOLDER_ID', f.getId());
  return f;
}
function uploadPhoto_(req) {
  if (!req.data) throw new Error('ไม่มีข้อมูลรูป');
  const blob = Utilities.newBlob(Utilities.base64Decode(req.data), req.mime || 'image/jpeg', req.name || ('photo_' + Date.now() + '.jpg'));
  const file = getFolder_().createFile(blob);
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { }
  const id = file.getId();
  return { ok: true, id: id, url: 'https://drive.google.com/thumbnail?id=' + id + '&sz=w1280' };
}
