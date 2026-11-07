/**
 * ShopStation — หลังบ้าน v3 (Google Apps Script + Google Sheets + Google Drive)
 * หลายหน่วยบริการ: ผู้ดูแลระบบสูงสุด (Super Admin) / ผู้ดูแลหน่วยบริการ / เจ้าหน้าที่ / ดูอย่างเดียว
 * เข้าสู่ระบบด้วยบัญชี Google หรือรหัส OTP ทางอีเมล / อนุมัติผู้ใช้ / ข้อมูลแยกตามหน่วยบริการ
 *
 * ขั้นตอนติดตั้ง (รายละเอียดใน "คู่มือติดตั้ง ShopStation")
 *  1) ใส่ค่า GOOGLE_CLIENT_ID และ ADMIN_EMAILS ด้านล่าง
 *  2) เลือกฟังก์ชัน setup → เรียกใช้ (อนุญาตสิทธิ์)
 *  3) การทำให้ใช้งานได้ → เว็บแอป  | ดำเนินการในฐานะ: ฉัน | ผู้ที่มีสิทธิ์เข้าถึง: ทุกคน
 *     (ถ้าแก้โค้ด: จัดการการทำให้ใช้งานได้ → แก้ไข → เวอร์ชันใหม่)
 */

// ====== ตั้งค่า ======
const GOOGLE_CLIENT_ID = '';            // OAuth Client ID จาก Google Cloud เช่น '1234-abc.apps.googleusercontent.com'
const ADMIN_EMAILS = [];                // ผู้ดูแลระบบสูงสุด (Super Admin) เช่น ['yourname@gmail.com']
const APP_NAME = 'ShopStation';
const APP_URL = 'https://pongop101.github.io/shopstation/app/';   // ลิงก์แอป (ใช้ในอีเมลแจ้งเตือน)
const SESSION_DAYS = 30;                // อายุการเข้าสู่ระบบ (วัน)
// หน่วยบริการ [รหัสหน่วยบริการ 5 หลัก, ชื่อ] — setup จะเพิ่ม/ปรับชื่อให้ตรงรายการนี้ (หน่วยที่มีอยู่แล้วจับคู่ด้วยรหัสหรือชื่อตำบล)
const DEFAULT_UNITS = [
  ['03822', 'โรงพยาบาลส่งเสริมสุขภาพตำบลอำเภอเมืองยโสธร'],
  ['03823', 'โรงพยาบาลส่งเสริมสุขภาพตำบลห้องข่า'],
  ['03824', 'โรงพยาบาลส่งเสริมสุขภาพตำบลตาดทอง'],
  ['03825', 'โรงพยาบาลส่งเสริมสุขภาพตำบลสว่าง'],
  ['03826', 'โรงพยาบาลส่งเสริมสุขภาพตำบลคำน้ำสร้าง'],
  ['03827', 'โรงพยาบาลส่งเสริมสุขภาพตำบลดอนกลอย'],
  ['03828', 'โรงพยาบาลส่งเสริมสุขภาพตำบลสามเพีย'],
  ['03829', 'โรงพยาบาลส่งเสริมสุขภาพตำบลดู่ทุ่ง'],
  ['03830', 'โรงพยาบาลส่งเสริมสุขภาพตำบลคำแดง'],
  ['03831', 'โรงพยาบาลส่งเสริมสุขภาพตำบลคำฮี'],
  ['03832', 'โรงพยาบาลส่งเสริมสุขภาพตำบลทุ่งแต้'],
  ['03833', 'โรงพยาบาลส่งเสริมสุขภาพตำบลสิงห์'],
  ['03834', 'โรงพยาบาลส่งเสริมสุขภาพตำบลนาสะไมย์'],
  ['03835', 'โรงพยาบาลส่งเสริมสุขภาพตำบลหนองหอย'],
  ['03836', 'โรงพยาบาลส่งเสริมสุขภาพตำบลหนองหิน'],
  ['03837', 'โรงพยาบาลส่งเสริมสุขภาพตำบลโนนค้อ'],
  ['03838', 'โรงพยาบาลส่งเสริมสุขภาพตำบลหนองคู'],
  ['03839', 'โรงพยาบาลส่งเสริมสุขภาพตำบลหนองบัว'],
  ['03840', 'โรงพยาบาลส่งเสริมสุขภาพตำบลขุมเงิน'],
  ['03841', 'โรงพยาบาลส่งเสริมสุขภาพตำบลทุ่งนางโอก'],
  ['03842', 'โรงพยาบาลส่งเสริมสุขภาพตำบลหนองเรือ'],
  ['10701', 'โรงพยาบาลยโสธร'],
  ['11921', 'โรงพยาบาลนายแพทย์หาญ'],
  ['14424', 'ศูนย์สุขภาพชุมชนเมืองบ้านท่าศรีธรรม'],
  ['22156', 'ศูนย์สุขภาพชุมชนเมืองยศ'],
  ['22802', 'สถานพยาบาลเรือนจำจังหวัดยโสธร'],
  ['28899', 'ศูนย์สุขภาพชุมชนตำบลเดิดค่ายบดินทรเดชา']
];

const SHEET_STORES = 'Stores';
const SHEET_INSP = 'Inspections';
const SHEET_USERS = 'Users';
const SHEET_UNITS = 'Units';
const SHEET_SELF = 'SelfReports';
const SHEET_MSG = 'Messages';
const MSG_COLS = ['id', 'created', 'from', 'fromName', 'to', 'toLabel', 'subject', 'body', 'link', 'kind', 'replyTo', 'readBy', 'hiddenBy'];
const SELF_COLS = ['id', 'created', 'type', 'status', 'unitId', 'storeId', 'storeName', 'contact', 'reviewedBy', 'reviewedAt', 'json'];
// สิทธิ์ย่อยของ อสม. (เจ้าหน้าที่/ผู้ดูแลหน่วยกำหนดได้)
const OSM_PERMS = { add: 'เพิ่มร้านค้าใหม่', info: 'แก้ไขข้อมูลร้าน', location: 'ปรับพิกัด', photo: 'ถ่ายรูปหน้าร้าน', inspect: 'ตรวจประเมินร้านค้า', close: 'แจ้งร้านปิด/เปิด' };
const OSM_DEFAULT = 'add,location,photo';
const STORE_GROUPS = { info: ['village', 'moo', 'name', 'type', 'kind', 'owner', 'caretaker', 'phone', 'address', 'license', 'note'], location: ['lat', 'lng'], photo: ['photo'], close: ['closed', 'closedAt', 'closedReason', 'closedBy'] };
const PHOTO_FOLDER = 'ShopStation Photos';
const USER_COLS = ['email', 'name', 'position', 'unit', 'phone', 'role', 'status', 'picture', 'created', 'lastLogin', 'approvedBy', 'note', 'unitId', 'transferTo', 'transferBy', 'transferAt', 'perms'];
const UNIT_COLS = ['id', 'name', 'created', 'createdBy', 'code'];
const ROLES = { superadmin: 'ผู้ดูแลระบบสูงสุด', unitadmin: 'ผู้ดูแลหน่วยบริการ', staff: 'เจ้าหน้าที่', osm: 'อสม.', viewer: 'ดูอย่างเดียว' };

function setup() {
  [SHEET_STORES, SHEET_INSP].forEach(n => getSheet_(n));
  usersSheet_();
  unitsSheet_();
  selfSheet_();
  msgSheet_();
  importUnits_();
  secret_();
  getFolder_();
  migrate_();
  Logger.log('ตั้งค่าเรียบร้อย: ' + SpreadsheetApp.getActiveSpreadsheet().getName());
  Logger.log('GOOGLE_CLIENT_ID: ' + (GOOGLE_CLIENT_ID || '(ยังไม่ได้ใส่ — ใช้ได้เฉพาะเข้าระบบด้วยรหัส OTP ทางอีเมล)'));
  Logger.log('ผู้ดูแลระบบสูงสุด: ' + (ADMIN_EMAILS.join(', ') || '(ผู้ที่เข้าระบบคนแรก)'));
  Logger.log('หน่วยบริการ: ' + readUnits_().map(u => u.name).join(', '));
  MailApp.getRemainingDailyQuota();
}

/** ออกจากระบบทุกเครื่องทันที (เปลี่ยนกุญแจ session) */
function logoutEveryone() {
  PropertiesService.getScriptProperties().setProperty('SESSION_SECRET', Utilities.getUuid() + Utilities.getUuid());
  Logger.log('ทุกเครื่องต้องเข้าสู่ระบบใหม่');
}

function doGet() {
  return json_({ ok: true, app: APP_NAME, message: 'ShopStation API พร้อมใช้งาน' });
}

function doPost(e) {
  ADMIN_CACHE_ = null;
  try {
    const req = JSON.parse(e.postData.contents);
    const a = req.action;
    // ---- ไม่ต้องเข้าระบบ ----
    if (a === 'ping') return json_({ ok: true, app: APP_NAME, sheet: SpreadsheetApp.getActiveSpreadsheet().getName(), googleClientId: GOOGLE_CLIENT_ID, otp: true,
      units: readUnits_().map(u => ({ id: u.id, name: u.name, code: u.code })) });
    if (a === 'loginGoogle') return json_(loginGoogle_(req.credential));
    if (a === 'otpSend') return json_(otpSend_(req.email));
    if (a === 'otpVerify') return json_(otpVerify_(req.email, req.code));
    // ---- ร้านค้ากรอกเอง (ไม่ต้องเข้าระบบ ใช้ลิงก์เฉพาะร้าน/หน่วย) ----
    if (a === 'selfGet') return json_(selfGet_(req.token));
    if (a === 'selfUnit') return json_(selfUnit_(req.unitId));
    if (a === 'selfSubmit') return json_(withLock_(() => selfSubmit_(req)));

    // ---- ต้องเข้าระบบ ----
    const user = auth_(req.token);
    if (a === 'me') return json_({ ok: true, user: publicUser_(user) });
    if (a === 'register') return json_(register_(user, req.profile || {}));
    if (user.status !== 'approved') throw err_('NOT_APPROVED', 'บัญชียังไม่ได้รับการอนุมัติจากผู้ดูแลระบบ');

    switch (a) {
      case 'list': return json_(list_(user));
      case 'saveStore': need_(user, ['superadmin', 'unitadmin', 'staff', 'osm']); return json_(withLock_(() => saveStore_(user, req.record, req.flat)));
      case 'saveInspection': need_(user, ['superadmin', 'unitadmin', 'staff', 'osm']); return json_(withLock_(() => saveInsp_(user, req.record, req.flat)));
      case 'deleteStore': need_(user, ['superadmin', 'unitadmin']); return json_(withLock_(() => deleteStore_(user, req.id)));
      case 'deleteInspection': need_(user, ['superadmin', 'unitadmin', 'staff', 'osm']); needPerm_(user, 'inspect'); return json_(withLock_(() => deleteInsp_(user, req.id)));
      case 'uploadPhoto': need_(user, ['superadmin', 'unitadmin', 'staff', 'osm']); return json_(uploadPhoto_(req));
      case 'users': need_(user, ['superadmin', 'unitadmin', 'staff']); return json_({ ok: true, users: visibleUsers_(user).map(publicUser_), units: readUnits_() });
      case 'userUpdate': need_(user, ['superadmin', 'unitadmin', 'staff']); return json_(withLock_(() => userUpdate_(user, req.email, req.changes || {})));
      case 'userDelete': need_(user, ['superadmin', 'unitadmin']); return json_(withLock_(() => userDelete_(user, req.email)));
      case 'transferRequest': return json_(withLock_(() => transferRequest_(user, req.email, req.unitId)));
      case 'transferCancel': return json_(withLock_(() => transferCancel_(user, req.email)));
      case 'transferDecide': need_(user, ['superadmin', 'unitadmin']); return json_(withLock_(() => transferDecide_(user, req.email, !!req.approve)));
      case 'storeLink': need_(user, ['superadmin', 'unitadmin', 'staff', 'osm']); return json_(withLock_(() => storeLink_(user, req.storeId)));
      case 'selfList': need_(user, ['superadmin', 'unitadmin', 'staff', 'osm']); return json_({ ok: true, reports: selfList_(user) });
      case 'selfReview': need_(user, ['superadmin', 'unitadmin', 'staff']); return json_(withLock_(() => selfReview_(user, req)));
      case 'msgList': return json_(msgList_(user));
      case 'msgContacts': return json_(msgContacts_(user));
      case 'msgSend': return json_(withLock_(() => msgSend_(user, req)));
      case 'msgRead': return json_(withLock_(() => msgMark_(user, req.ids || [], 'readBy')));
      case 'msgHide': return json_(withLock_(() => msgMark_(user, req.ids || [], 'hiddenBy')));
      case 'unitSave': need_(user, ['superadmin']); return json_(withLock_(() => unitSave_(user, req.unit || {})));
      case 'unitDelete': need_(user, ['superadmin']); return json_(withLock_(() => unitDelete_(req.id)));
      default: throw new Error('ไม่รู้จักคำสั่ง ' + a);
    }
  } catch (er) {
    return json_({ ok: false, code: er.code || 'ERROR', error: String(er && er.message || er) });
  }
}

/* ================= สิทธิ์ ================= */
function err_(code, msg) { const e = new Error(msg); e.code = code; return e; }
function need_(user, roles) { if (roles.indexOf(user.role) < 0) throw err_('FORBIDDEN', 'สิทธิ์ของคุณ (' + (ROLES[user.role] || user.role) + ') ทำรายการนี้ไม่ได้'); }
function norm_(email) { return String(email || '').trim().toLowerCase(); }
function isSuper_(u) { return u && u.role === 'superadmin'; }
function isSuperEmail_(email) { return ADMIN_EMAILS.map(norm_).indexOf(norm_(email)) >= 0; }
function sameUnit_(user, unitId) { return isSuper_(user) || (!!user.unitId && String(unitId || '') === String(user.unitId)); }
function permSet_(u) { const o = {}; String(u.perms || '').split(',').map(x => x.trim()).filter(Boolean).forEach(k => o[k] = true); return o; }
function hasPerm_(u, p) { return u.role !== 'osm' || !!permSet_(u)[p]; }
function needPerm_(u, p) { if (!hasPerm_(u, p)) throw err_('FORBIDDEN', 'อสม. ยังไม่ได้รับสิทธิ์: ' + (OSM_PERMS[p] || p)); }
function needUnit_(user) { if (!isSuper_(user) && !user.unitId) throw err_('NO_UNIT', 'บัญชียังไม่ได้กำหนดหน่วยบริการ กรุณาติดต่อผู้ดูแลระบบ'); }

/* ================= session ================= */
function secret_() {
  const p = PropertiesService.getScriptProperties();
  let s = p.getProperty('SESSION_SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('SESSION_SECRET', s); }
  return s;
}
function b64u_(bytesOrStr) { return Utilities.base64EncodeWebSafe(bytesOrStr).replace(/=+$/, ''); }
function makeToken_(email) {
  const exp = Date.now() + SESSION_DAYS * 864e5;
  const payload = email + '|' + exp;
  const sig = b64u_(Utilities.computeHmacSha256Signature(payload, secret_()));
  return { token: b64u_(payload) + '.' + sig, exp: exp };
}
function auth_(token) {
  if (!token) throw err_('AUTH_REQUIRED', 'กรุณาเข้าสู่ระบบ');
  const parts = String(token).split('.');
  if (parts.length !== 2) throw err_('AUTH_REQUIRED', 'กรุณาเข้าสู่ระบบใหม่');
  let payload;
  try { payload = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString(); } catch (e) { throw err_('AUTH_REQUIRED', 'กรุณาเข้าสู่ระบบใหม่'); }
  const sig = b64u_(Utilities.computeHmacSha256Signature(payload, secret_()));
  if (sig !== parts[1]) throw err_('AUTH_REQUIRED', 'การเข้าสู่ระบบไม่ถูกต้อง กรุณาเข้าสู่ระบบใหม่');
  const [email, exp] = payload.split('|');
  if (Date.now() > Number(exp)) throw err_('AUTH_REQUIRED', 'หมดเวลาการเข้าสู่ระบบ กรุณาเข้าสู่ระบบใหม่');
  const u = findUser_(email);
  if (!u) throw err_('AUTH_REQUIRED', 'ไม่พบบัญชีผู้ใช้ (อาจถูกลบแล้ว)');
  fixRole_(u);
  if (u.status === 'disabled') throw err_('DISABLED', 'บัญชีนี้ถูกระงับการใช้งาน');
  return u;
}
/** ปรับบทบาทให้ตรงกับโครงสร้างใหม่ (admin เดิม → superadmin/unitadmin) */
function fixRole_(u) {
  let ch = false;
  if (isSuperEmail_(u.email) && (u.role !== 'superadmin' || u.status !== 'approved')) { u.role = 'superadmin'; u.status = 'approved'; ch = true; }
  else if (u.role === 'admin') { u.role = isSuperEmail_(u.email) || !ADMIN_EMAILS.length ? 'superadmin' : 'unitadmin'; ch = true; }
  else if (!ROLES[u.role]) { u.role = 'staff'; ch = true; }
  if (ch) writeUser_(u, false);
}

/* ================= login ================= */
function loginGoogle_(credential) {
  if (!GOOGLE_CLIENT_ID) throw new Error('ผู้ดูแลยังไม่ได้ตั้งค่า GOOGLE_CLIENT_ID — ใช้การเข้าระบบด้วยรหัสทางอีเมลแทน');
  if (!credential) throw new Error('ไม่มีข้อมูลการเข้าสู่ระบบ');
  const r = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(credential), { muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) throw new Error('ยืนยันตัวตนกับ Google ไม่สำเร็จ');
  const t = JSON.parse(r.getContentText());
  if (t.aud !== GOOGLE_CLIENT_ID) throw new Error('Client ID ไม่ตรงกัน');
  if (['accounts.google.com', 'https://accounts.google.com'].indexOf(t.iss) < 0) throw new Error('ผู้ออกโทเค็นไม่ถูกต้อง');
  if (String(t.email_verified) !== 'true') throw new Error('อีเมลนี้ยังไม่ได้รับการยืนยันกับ Google');
  if (Number(t.exp) * 1000 < Date.now()) throw new Error('โทเค็นหมดอายุ');
  return withLock_(() => loginOk_(norm_(t.email), { name: t.name || '', picture: t.picture || '' }));
}
function otpSend_(email) {
  email = norm_(email);
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('อีเมลไม่ถูกต้อง');
  const u = findUser_(email);
  if (u && u.status === 'disabled') throw new Error('บัญชีนี้ถูกระงับการใช้งาน');
  const c = CacheService.getScriptCache();
  if (c.get('otpwait_' + email)) throw new Error('เพิ่งส่งรหัสไปแล้ว กรุณารอ 1 นาทีแล้วลองใหม่');
  const code = String(Math.floor(100000 + Math.random() * 900000));
  c.put('otp_' + email, JSON.stringify({ code: code, tries: 0 }), 600);
  c.put('otpwait_' + email, '1', 60);
  MailApp.sendEmail({
    to: email, name: APP_NAME, subject: 'รหัสเข้าสู่ระบบ ' + APP_NAME + ': ' + code,
    htmlBody: '<div style="font-family:sans-serif;font-size:16px">รหัสเข้าสู่ระบบ ' + APP_NAME + ' ของคุณคือ<div style="font-size:32px;font-weight:bold;letter-spacing:6px;margin:12px 0;color:#0f766e">' + code +
      '</div>รหัสใช้ได้ 10 นาที หากคุณไม่ได้ขอรหัสนี้ ไม่ต้องดำเนินการใด ๆ</div>'
  });
  return { ok: true };
}
function otpVerify_(email, code) {
  email = norm_(email);
  const c = CacheService.getScriptCache();
  const raw = c.get('otp_' + email);
  if (!raw) throw new Error('รหัสหมดอายุหรือยังไม่ได้ขอรหัส');
  const o = JSON.parse(raw);
  if (o.tries >= 5) { c.remove('otp_' + email); throw new Error('ใส่รหัสผิดเกินกำหนด กรุณาขอรหัสใหม่'); }
  if (String(code).trim() !== o.code) { o.tries++; c.put('otp_' + email, JSON.stringify(o), 600); throw new Error('รหัสไม่ถูกต้อง'); }
  c.remove('otp_' + email);
  return withLock_(() => loginOk_(email, {}));
}
function loginOk_(email, info) {
  let u = findUser_(email);
  const users = readUsers_();
  const noSuperYet = !ADMIN_EMAILS.length && !users.some(x => (x.role === 'superadmin' || x.role === 'admin') && x.status === 'approved');
  if (!u) {
    u = { email: email, name: info.name || '', position: '', unit: '', phone: '', role: 'staff', status: 'pending', picture: info.picture || '', created: new Date(), lastLogin: new Date(), approvedBy: '', note: '', unitId: '' };
    if (isSuperEmail_(email) || noSuperYet) { u.role = 'superadmin'; u.status = 'approved'; u.approvedBy = 'system'; }
    writeUser_(u, true);
    u = findUser_(email);
  } else {
    u.lastLogin = new Date();
    if (info.picture) u.picture = info.picture;
    if (!u.name && info.name) u.name = info.name;
    writeUser_(u, false);
  }
  fixRole_(u);
  if (u.status === 'disabled') throw new Error('บัญชีนี้ถูกระงับการใช้งาน');
  const t = makeToken_(email);
  return { ok: true, token: t.token, exp: t.exp, user: publicUser_(u) };
}
function register_(user, p) {
  return withLock_(() => {
    const u = findUser_(user.email);
    const isNew = !u.position && !u.phone;
    ['name', 'position', 'phone'].forEach(k => { if (p[k] !== undefined) u[k] = String(p[k]).slice(0, 200); });
    if (u.status === 'pending' && (p.asRole === 'osm' || p.asRole === 'staff')) { u.role = p.asRole; if (p.asRole === 'osm' && !u.perms) u.perms = OSM_DEFAULT; }
    const units = readUnits_();
    if (p.unitId !== undefined && p.unitId !== u.unitId) {
      const un = units.filter(x => x.id === p.unitId)[0];
      if (!un && units.length) throw new Error('กรุณาเลือกหน่วยบริการ');
      if (un) {
        if (u.status === 'approved' && !isSuper_(u) && u.unitId) throw new Error('ย้ายหน่วยบริการ ให้ใช้เมนู ตั้งค่า → ย้ายหน่วยบริการ');
        u.unitId = un.id; u.unit = un.name;
      }
    } else if (!units.length && p.unit !== undefined) u.unit = String(p.unit).slice(0, 200);
    writeUser_(u, false);
    if (isNew && u.status === 'pending') { notifyAdmins_(u); sysMsg_(u.unitId ? 'unit:' + u.unitId + ':admin' : 'super', '👤 ผู้สมัครใหม่รออนุมัติ: ' + (u.name || u.email),
      (u.name || '-') + ' (' + u.email + ')\nสมัครเป็น: ' + (ROLES[u.role] || u.role) + '\nตำแหน่ง: ' + (u.position || '-') + '\nหน่วยบริการ: ' + (u.unit || '-') + '\nโทร: ' + (u.phone || '-'), '#users'); }
    return { ok: true, user: publicUser_(u) };
  });
}

/* ================= แจ้งเตือน ================= */
function adminEmailsFor_(unitId) {
  const set = {};
  ADMIN_EMAILS.map(norm_).forEach(e => { if (e) set[e] = 1; });
  readUsers_().filter(x => x.status === 'approved' && (x.role === 'superadmin' || (x.role === 'unitadmin' && unitId && String(x.unitId) === String(unitId))))
    .forEach(x => set[x.email] = 1);
  return Object.keys(set);
}
function notifyAdmins_(u) {
  try {
    const admins = adminEmailsFor_(u.unitId);
    if (!admins.length) return;
    MailApp.sendEmail({
      to: admins.join(','), name: APP_NAME, subject: '🔔 ' + APP_NAME + ': ' + (/ย้าย/.test(u.note || '') ? 'มีเจ้าหน้าที่ย้ายเข้า' : 'มีผู้สมัครใหม่') + 'รออนุมัติ — ' + (u.name || u.email) + (u.unit ? ' (' + u.unit + ')' : ''),
      htmlBody: '<div style="font-family:sans-serif;font-size:15px;line-height:1.6">มีผู้สมัครใช้งาน ' + APP_NAME + ' รอการอนุมัติ<br><br>' +
        '<b>' + (u.name || '-') + '</b> (' + u.email + ')<br>' + (u.note ? '<span style="color:#b45309">' + u.note + '</span><br>' : '') + 'ตำแหน่ง: ' + (u.position || '-') + '<br>หน่วยบริการ: ' + (u.unit || '-') + '<br>โทร: ' + (u.phone || '-') +
        '<br><br><a href="' + APP_URL + '#users" style="background:#0f766e;color:#fff;padding:10px 18px;border-radius:8px;text-decoration:none;font-weight:bold">เปิดหน้าอนุมัติผู้ใช้</a>' +
        '<br><br><span style="color:#6b7280;font-size:13px">หรือเข้าแอป → ตั้งค่า → จัดการผู้ใช้</span></div>'
    });
  } catch (e) { }
}
/** ส่งอีเมลสรุปผู้สมัครที่รออนุมัติทั้งหมดถึงผู้ดูแลระบบสูงสุด (รันเองได้ หรือตั้งทริกเกอร์รายวัน) */
function notifyPending() {
  const pend = readUsers_().filter(u => u.status === 'pending');
  if (!pend.length) { Logger.log('ไม่มีผู้สมัครรออนุมัติ'); return; }
  const admins = adminEmailsFor_('');
  MailApp.sendEmail({
    to: admins.join(','), name: APP_NAME, subject: '🔔 ' + APP_NAME + ': ผู้สมัครรออนุมัติ ' + pend.length + ' คน',
    htmlBody: '<div style="font-family:sans-serif;font-size:15px;line-height:1.6">ผู้สมัครรออนุมัติ ' + pend.length + ' คน<ul>' +
      pend.map(u => '<li><b>' + (u.name || '-') + '</b> (' + u.email + ') ' + (u.position || '') + ' · ' + (u.unit || '') + ' · ' + (u.phone || '') + '</li>').join('') +
      '</ul><a href="' + APP_URL + '#users" style="background:#0f766e;color:#fff;padding:10px 18px;border-radius:8px;text-decoration:none;font-weight:bold">เปิดหน้าอนุมัติผู้ใช้</a></div>'
  });
  Logger.log('ส่งอีเมลถึง ' + admins.join(', ') + ' แล้ว (' + pend.length + ' คน)');
}

/* ================= ผู้ใช้ ================= */
function visibleUsers_(admin) {
  const all = readUsers_();
  if (admin.role === 'staff') return all.filter(u => u.role === 'osm' && u.unitId && String(u.unitId) === String(admin.unitId));
  return isSuper_(admin) ? all : all.filter(u => u.role !== 'superadmin' && ((u.unitId && String(u.unitId) === String(admin.unitId)) || (u.transferTo && String(u.transferTo) === String(admin.unitId))));
}
function canManage_(admin, u) {
  if (isSuper_(admin)) return true;
  if (admin.role === 'staff') return !!admin.unitId && u.role === 'osm' && String(u.unitId) === String(admin.unitId);
  return admin.role === 'unitadmin' && !!admin.unitId && String(u.unitId) === String(admin.unitId) && u.role !== 'superadmin' && u.role !== 'unitadmin';
}
function userUpdate_(admin, email, ch) {
  email = norm_(email);
  const u = findUser_(email);
  if (!u) throw new Error('ไม่พบผู้ใช้');
  if (email === admin.email) throw new Error('แก้ไขสิทธิ์หรือสถานะของบัญชีตัวเองไม่ได้');
  if (!canManage_(admin, u)) throw err_('FORBIDDEN', 'จัดการได้เฉพาะเจ้าหน้าที่ในหน่วยบริการของคุณ');
  if (isSuperEmail_(email)) throw new Error('แก้ไขผู้ดูแลระบบสูงสุดไม่ได้');
  if (admin.role === 'staff') {   // เจ้าหน้าที่: จัดการ อสม. ได้เฉพาะ อนุมัติ/ระงับ และสิทธิ์ย่อย — เปลี่ยนเป็นสิทธิ์อื่นไม่ได้
    if (ch.role && ch.role !== 'osm') throw err_('FORBIDDEN', 'เจ้าหน้าที่เปลี่ยนสิทธิ์ อสม. เป็นสิทธิ์อื่นไม่ได้');
    delete ch.role; delete ch.unitId;
  }
  if (ch.perms !== undefined) {
    const list = String(ch.perms).split(',').map(x => x.trim()).filter(k => OSM_PERMS[k]);
    if (list.join(',') !== String(u.perms || '') && u.status === 'approved') sysMsg_(u.email, '🔑 สิทธิ์ อสม. ของคุณถูกปรับปรุง', 'สิทธิ์ใหม่: ' + permText_(list.join(',')) + '\nโดย: ' + (admin.name || admin.email), '#settings');
    u.perms = list.join(',');
  }
  if (ch.role) {
    if (!ROLES[ch.role]) throw new Error('บทบาทไม่ถูกต้อง');
    if (ch.role === 'osm' && !u.perms) u.perms = OSM_DEFAULT;
    if (!isSuper_(admin) && ['staff', 'osm', 'viewer'].indexOf(ch.role) < 0) throw err_('FORBIDDEN', 'ผู้ดูแลหน่วยบริการกำหนดได้เฉพาะสิทธิ์ เจ้าหน้าที่ / อสม. / ดูอย่างเดียว');
    u.role = ch.role;
  }
  if (ch.unitId !== undefined && ch.unitId !== u.unitId) {
    if (!isSuper_(admin)) throw err_('FORBIDDEN', 'ผู้ดูแลหน่วยบริการ ให้ใช้ปุ่ม "ส่งคำขอย้าย" (หน่วยปลายทางต้องอนุมัติ)');
    const un = readUnits_().filter(x => x.id === ch.unitId)[0];
    if (!un) throw new Error('ไม่พบหน่วยบริการ');
    u.unitId = un.id; u.unit = un.name; u.transferTo = ''; u.transferBy = ''; u.transferAt = '';
  }
  if (ch.status) {
    if (['pending', 'approved', 'disabled'].indexOf(ch.status) < 0) throw new Error('สถานะไม่ถูกต้อง');
    if (ch.status === 'approved' && u.role !== 'superadmin' && !u.unitId) throw new Error('กรุณากำหนดหน่วยบริการให้ผู้ใช้ก่อนอนุมัติ');
    const wasPending = u.status !== 'approved' && ch.status === 'approved';
    u.status = ch.status; if (ch.status === 'approved') u.approvedBy = admin.email;
    if (wasPending) {
      sysMsg_(u.email, '🎉 บัญชีของคุณได้รับการอนุมัติแล้ว', 'ยินดีต้อนรับสู่ ' + APP_NAME + '\nหน่วยบริการ: ' + (u.unit || '-') + '\nสิทธิ์: ' + ROLES[u.role] + (u.role === 'osm' ? '\nสิทธิ์ อสม.: ' + permText_(u.perms) : '') + '\nอนุมัติโดย: ' + (admin.name || admin.email), '#map');
      try { MailApp.sendEmail({ to: u.email, name: APP_NAME, subject: APP_NAME + ': บัญชีของคุณได้รับการอนุมัติแล้ว', body: 'บัญชี ' + u.email + ' ได้รับการอนุมัติให้ใช้งาน ' + APP_NAME + ' แล้ว\nหน่วยบริการ: ' + (u.unit || '-') + '\nสิทธิ์: ' + ROLES[u.role] + '\n\nเปิดแอป: ' + APP_URL }); } catch (e) { }
    }
  }
  if (ch.note !== undefined) u.note = String(ch.note).slice(0, 300);
  writeUser_(u, false);
  return { ok: true, user: publicUser_(u) };
}
/* ---- ย้ายหน่วยบริการ ---- */
function transferRequest_(actor, email, unitId) {
  email = norm_(email || actor.email);
  const u = findUser_(email);
  if (!u) throw new Error('ไม่พบผู้ใช้');
  if (u.role === 'superadmin') throw new Error('ผู้ดูแลระบบสูงสุดไม่ต้องสังกัดหน่วยบริการ');
  const un = readUnits_().filter(x => x.id === unitId)[0];
  if (!un) throw new Error('กรุณาเลือกหน่วยบริการปลายทาง');
  if (String(u.unitId) === String(unitId)) throw new Error('อยู่ในหน่วยบริการนี้อยู่แล้ว');
  const self = email === actor.email;
  if (!self && !isSuper_(actor) && !(actor.role === 'unitadmin' && actor.unitId && actor.unitId === u.unitId))
    throw err_('FORBIDDEN', 'สั่งย้ายได้เฉพาะเจ้าหน้าที่ในหน่วยบริการของคุณ');
  if (self && u.status !== 'approved') throw new Error('บัญชีที่ยังไม่อนุมัติ ให้แก้หน่วยบริการในข้อมูลสมัครแทน');
  if (isSuper_(actor) && !self) {   // ผู้ดูแลระบบสูงสุดย้ายได้ทันที
    u.unitId = un.id; u.unit = un.name; u.transferTo = ''; u.transferBy = ''; u.transferAt = '';
    if (u.role === 'unitadmin') u.role = 'staff';
    writeUser_(u, false);
    mailUser_(u, 'ย้ายหน่วยบริการเรียบร้อย', 'บัญชีของคุณถูกย้ายไปสังกัด ' + un.name + ' แล้ว');
    return { ok: true, moved: true, user: publicUser_(u) };
  }
  u.transferTo = un.id; u.transferBy = actor.email; u.transferAt = new Date();
  writeUser_(u, false);
  sysMsg_('unit:' + un.id + ':admin', '🔁 คำขอย้ายเข้า ' + un.name + ': ' + (u.name || u.email), (u.name || '-') + ' (' + u.email + ')\nตำแหน่ง: ' + (u.position || '-') + '\nจากหน่วย: ' + (u.unit || '-') + '\nผู้ขอ: ' + actor.email, '#users');
  try {
    const to = adminEmailsFor_(un.id);
    MailApp.sendEmail({ to: to.join(','), name: APP_NAME, subject: '🔔 ' + APP_NAME + ': คำขอย้ายเข้า ' + un.name + ' — ' + (u.name || u.email),
      htmlBody: '<div style="font-family:sans-serif;font-size:15px;line-height:1.6">มีคำขอย้ายเจ้าหน้าที่เข้าหน่วยบริการ <b>' + un.name + '</b><br><br><b>' + (u.name || '-') + '</b> (' + u.email + ')<br>ตำแหน่ง: ' + (u.position || '-') +
        '<br>จากหน่วย: ' + (u.unit || '-') + '<br>ผู้ขอ: ' + actor.email + '<br><br><a href="' + APP_URL + '#users" style="background:#0f766e;color:#fff;padding:10px 18px;border-radius:8px;text-decoration:none;font-weight:bold">เปิดหน้าอนุมัติ</a></div>' });
  } catch (e) { }
  return { ok: true, requested: true, user: publicUser_(u) };
}
function transferCancel_(actor, email) {
  email = norm_(email || actor.email);
  const u = findUser_(email);
  if (!u || !u.transferTo) return { ok: true };
  if (email !== actor.email && !isSuper_(actor) && !(actor.role === 'unitadmin' && (actor.unitId === u.unitId || actor.unitId === u.transferTo)))
    throw err_('FORBIDDEN', 'ไม่มีสิทธิ์ยกเลิกคำขอนี้');
  u.transferTo = ''; u.transferBy = ''; u.transferAt = '';
  writeUser_(u, false);
  return { ok: true, user: publicUser_(u) };
}
function transferDecide_(admin, email, approve) {
  const u = findUser_(email);
  if (!u || !u.transferTo) throw new Error('ไม่พบคำขอย้าย');
  if (!isSuper_(admin) && !(admin.role === 'unitadmin' && admin.unitId === u.transferTo)) throw err_('FORBIDDEN', 'อนุมัติได้เฉพาะผู้ดูแลหน่วยบริการปลายทาง');
  const name = unitName_(u.transferTo);
  if (approve) {
    u.unitId = u.transferTo; u.unit = name;
    if (u.role === 'unitadmin') u.role = 'staff';
    u.approvedBy = admin.email;
  }
  u.transferTo = ''; u.transferBy = ''; u.transferAt = '';
  writeUser_(u, false);
  mailUser_(u, approve ? 'อนุมัติการย้ายหน่วยบริการแล้ว' : 'คำขอย้ายหน่วยบริการไม่ได้รับการอนุมัติ',
    approve ? 'คุณได้ย้ายไปสังกัด ' + name + ' แล้ว กรุณาเปิดแอปเพื่อซิงก์ข้อมูลของหน่วยบริการใหม่' : 'คำขอย้ายไป ' + name + ' ไม่ได้รับการอนุมัติ');
  return { ok: true, user: publicUser_(u) };
}
function mailUser_(u, subject, body) {
  sysMsg_(u.email, subject, body, '');
  try { MailApp.sendEmail({ to: u.email, name: APP_NAME, subject: APP_NAME + ': ' + subject, body: body + '\n\nเปิดแอป: ' + APP_URL }); } catch (e) { }
}
function userDelete_(admin, email) {
  if (admin.role === 'staff') throw err_('FORBIDDEN', 'เจ้าหน้าที่ลบบัญชีไม่ได้ (ใช้ “ระงับ” แทน)');
  email = norm_(email);
  const u = findUser_(email);
  if (!u) throw new Error('ไม่พบผู้ใช้');
  if (email === admin.email) throw new Error('ลบบัญชีของตัวเองไม่ได้');
  if (isSuperEmail_(email)) throw new Error('ลบผู้ดูแลระบบสูงสุดไม่ได้');
  if (!canManage_(admin, u)) throw err_('FORBIDDEN', 'ลบได้เฉพาะเจ้าหน้าที่ในหน่วยบริการของคุณ');
  usersSheet_().deleteRow(u._row);
  return { ok: true, deleted: email };
}
function publicUser_(u) {
  return { email: u.email, name: u.name, position: u.position, unit: u.unit, unitId: u.unitId || '', phone: u.phone, role: u.role, status: u.status, picture: u.picture,
    perms: u.role === 'osm' ? String(u.perms || '') : '', transferTo: String(u.transferTo || ''), transferBy: u.transferBy || '', transferAt: u.transferAt ? new Date(u.transferAt).getTime() : 0,
    created: u.created ? new Date(u.created).getTime() : 0, lastLogin: u.lastLogin ? new Date(u.lastLogin).getTime() : 0, approvedBy: u.approvedBy, note: u.note };
}
function usersSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_USERS);
  if (!sh) { sh = ss.insertSheet(SHEET_USERS); sh.setFrozenRows(1); }
  const lc = sh.getLastColumn();
  const h = lc ? sh.getRange(1, 1, 1, lc).getValues()[0] : [];
  if (h.length < USER_COLS.length || USER_COLS.some((k, i) => h[i] !== k)) {
    sh.getRange(1, 1, 1, USER_COLS.length).setValues([USER_COLS]).setFontWeight('bold').setBackground('#fde68a');
  }
  return sh;
}
function readUsers_() {
  const sh = usersSheet_();
  const lr = sh.getLastRow();
  if (lr < 2) return [];
  const h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  return sh.getRange(2, 1, lr - 1, h.length).getValues().map((r, i) => {
    const o = { _row: i + 2 }; h.forEach((k, j) => o[k] = r[j]); o.email = norm_(o.email); o.unitId = String(o.unitId || ''); return o;
  }).filter(o => o.email);
}
function findUser_(email) { email = norm_(email); return readUsers_().filter(u => u.email === email)[0] || null; }
function writeUser_(u, isNew) {
  const sh = usersSheet_();
  const vals = [USER_COLS.map(k => u[k] === undefined ? '' : u[k])];
  if (isNew || !u._row) sh.getRange(sh.getLastRow() + 1, 1, 1, USER_COLS.length).setValues(vals);
  else sh.getRange(u._row, 1, 1, USER_COLS.length).setValues(vals);
}

/* ================= หน่วยบริการ ================= */
function unitsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_UNITS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_UNITS);
    sh.setFrozenRows(1);
  }
  const lc = sh.getLastColumn(); const h = lc ? sh.getRange(1, 1, 1, lc).getValues()[0] : [];
  if (h.length < UNIT_COLS.length || UNIT_COLS.some((k, i) => h[i] !== k)) sh.getRange(1, 1, 1, UNIT_COLS.length).setValues([UNIT_COLS]).setFontWeight('bold').setBackground('#bfdbfe');
  return sh;
}
function newId_() { return 'u' + Utilities.getUuid().replace(/-/g, '').slice(0, 10); }
function readUnits_() {
  const sh = unitsSheet_();
  const lr = sh.getLastRow();
  if (lr < 2) return [];
  return sh.getRange(2, 1, lr - 1, UNIT_COLS.length).getValues().map((r, i) => ({ _row: i + 2, id: String(r[0]), name: String(r[1]).trim(), code: String(r[4] || '').trim() }))
    .filter(u => u.id && u.name).sort((a, b) => (a.code || '99999').localeCompare(b.code || '99999') || a.name.localeCompare(b.name, 'th'));
}
function unitSave_(admin, p) {
  const name = String(p.name || '').trim();
  const code = String(p.code || '').trim();
  if (!name) throw new Error('กรุณาใส่ชื่อหน่วยบริการ');
  if (code && !/^\d{5}$/.test(code)) throw new Error('รหัสหน่วยบริการต้องเป็นตัวเลข 5 หลัก');
  const units = readUnits_();
  if (units.some(u => u.name === name && u.id !== p.id)) throw new Error('มีหน่วยบริการชื่อนี้แล้ว');
  if (code && units.some(u => u.code === code && u.id !== p.id)) throw new Error('มีรหัสหน่วยบริการ ' + code + ' แล้ว');
  const sh = unitsSheet_();
  if (p.id) {
    const u = units.filter(x => x.id === p.id)[0];
    if (!u) throw new Error('ไม่พบหน่วยบริการ');
    sh.getRange(u._row, 2).setValue(name);
    if (p.code !== undefined) sh.getRange(u._row, 5).setValue(code ? "'" + code : '');
    readUsers_().filter(x => x.unitId === u.id).forEach(x => { x.unit = name; writeUser_(x, false); });
    return { ok: true, unit: { id: u.id, name: name } };
  }
  const id = newId_();
  sh.appendRow([id, name, new Date(), admin.email, code ? "'" + code : '']);
  return { ok: true, unit: { id: id, name: name, code: code } };
}
function unitDelete_(id) {
  const u = readUnits_().filter(x => x.id === id)[0];
  if (!u) throw new Error('ไม่พบหน่วยบริการ');
  const nUser = readUsers_().filter(x => x.unitId === id).length;
  const nStore = readAll_(SHEET_STORES).filter(s => s.unitId === id).length;
  if (nUser || nStore) throw new Error('ลบไม่ได้: ยังมีผู้ใช้ ' + nUser + ' คน และร้านค้า ' + nStore + ' ร้านในหน่วยนี้');
  unitsSheet_().deleteRow(u._row);
  return { ok: true };
}
function unitName_(id) { const u = readUnits_().filter(x => x.id === id)[0]; return u ? u.name : ''; }

/** เพิ่ม/ปรับหน่วยบริการตาม DEFAULT_UNITS (จับคู่ด้วยรหัส หรือชื่อตำบลสำหรับหน่วยเดิมที่ยังไม่มีรหัส) */
function importUnits_() {
  const key = s => String(s || '').replace(/โรงพยาบาลส่งเสริมสุขภาพตำบล|รพ\.?\s*สต\.?|รพสต\.?|บ้าน|\s/g, '');
  const sh = unitsSheet_();
  DEFAULT_UNITS.forEach(([code, name]) => {
    const units = readUnits_();
    let u = units.filter(x => x.code === code)[0];
    if (!u) u = units.filter(x => !x.code && key(x.name) && (key(x.name) === key(name)))[0];
    if (u) {
      if (u.name !== name || u.code !== code) {
        sh.getRange(u._row, 2).setValue(name); sh.getRange(u._row, 5).setValue("'" + code);
        readUsers_().filter(x => x.unitId === u.id).forEach(x => { x.unit = name; writeUser_(x, false); });
        Logger.log('ปรับหน่วย ' + u.name + ' → ' + code + ' ' + name);
      }
    } else { sh.appendRow([newId_(), name, new Date(), 'system', "'" + code]); }
  });
}
/** ย้ายข้อมูลเดิม: จับคู่หน่วยบริการของผู้ใช้จากชื่อ และกำหนดหน่วยให้ร้าน/ผลตรวจเดิมตามผู้บันทึก */
function migrate_() {
  const units = readUnits_();
  const key = s => String(s || '').replace(/โรงพยาบาลส่งเสริมสุขภาพตำบล|รพ\.?\s*สต\.?|รพสต\.?|\s/g, '');
  const users = readUsers_();
  users.forEach(u => {
    fixRole_(u);
    if (!u.unitId && u.unit) {
      const k = key(u.unit);
      const m = units.filter(x => k && (key(x.name) === k || k.indexOf(key(x.name)) >= 0 || key(x.name).indexOf(k) >= 0))[0];
      if (m) { u.unitId = m.id; u.unit = m.name; writeUser_(u, false); Logger.log('ผู้ใช้ ' + u.email + ' → ' + m.name); }
    }
  });
  const byEmail = {}; readUsers_().forEach(u => byEmail[u.email] = u);
  const stores = readAll_(SHEET_STORES); const storeUnit = {};
  stores.forEach(s => {
    if (!s.unitId) {
      const u = byEmail[norm_(s.createdBy)];
      if (u && u.unitId) { s.unitId = u.unitId; rewriteJson_(SHEET_STORES, s); Logger.log('ร้าน ' + s.name + ' → ' + u.unit); }
    }
    storeUnit[s.id] = s.unitId || '';
  });
  readAll_(SHEET_INSP).forEach(i => { if (!i.unitId && storeUnit[i.storeId]) { i.unitId = storeUnit[i.storeId]; rewriteJson_(SHEET_INSP, i); } });
}
function rewriteJson_(name, rec) {
  const sh = getSheet_(name); const h = headers_(sh); const jc = h.indexOf('json'); const lr = sh.getLastRow();
  if (lr < 2 || jc < 0) return;
  const ids = sh.getRange(2, 1, lr - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(rec.id)) {
    sh.getRange(i + 2, jc + 1).setValue(JSON.stringify(rec));
    const uc = h.indexOf('หน่วยบริการ'); if (uc >= 0) sh.getRange(i + 2, uc + 1).setValue(unitName_(rec.unitId));
    return;
  }
}

/* ================= ข้อมูลร้าน / ผลตรวจ ================= */
function list_(user) {
  const stores = readAll_(SHEET_STORES).filter(s => sameUnit_(user, s.unitId));
  const ids = {}; stores.forEach(s => ids[s.id] = 1);
  const insps = readAll_(SHEET_INSP).filter(i => ids[i.storeId]);
  const pending = isSuper_(user) ? readUsers_().filter(u => u.status === 'pending' || u.transferTo).length :
    user.role === 'unitadmin' ? readUsers_().filter(u => (u.status === 'pending' && u.unitId === user.unitId) || (u.transferTo && u.transferTo === user.unitId)).length : 0;
  const selfPending = readSelf_().filter(r => r.status === 'pending' && sameUnit_(user, r.unitId)).length;
  return { ok: true, user: publicUser_(user), stores: stores, inspections: insps, pending: pending, selfPending: selfPending, msgUnread: msgUnread_(user), units: readUnits_().map(u => ({ id: u.id, name: u.name, code: u.code })) };
}
function findRec_(name, id) { return readAll_(name).filter(r => r.id === id)[0] || null; }
function saveStore_(user, rec, flat) {
  if (!rec || !rec.id) throw new Error('ข้อมูลไม่ครบ');
  needUnit_(user);
  const old = findRec_(SHEET_STORES, rec.id);
  if (old && !sameUnit_(user, old.unitId)) throw err_('FORBIDDEN', 'ร้านนี้อยู่ในหน่วยบริการอื่น');
  if (user.role === 'osm') {
    const P = permSet_(user);
    if (!old) { if (!P.add) throw err_('FORBIDDEN', 'อสม. ยังไม่ได้รับสิทธิ์: ' + OSM_PERMS.add); }
    else {   // รับเฉพาะส่วนที่มีสิทธิ์ ส่วนอื่นคงค่าเดิม
      const merged = JSON.parse(JSON.stringify(old));
      let any = false;
      Object.keys(STORE_GROUPS).forEach(g => { if (P[g]) { any = true; STORE_GROUPS[g].forEach(k => { if (rec[k] === undefined) delete merged[k]; else merged[k] = rec[k]; }); } });
      if (!any) throw err_('FORBIDDEN', 'อสม. ยังไม่ได้รับสิทธิ์แก้ไขข้อมูลร้าน');
      ['selfToken'].forEach(k => { if (old[k]) merged[k] = old[k]; });
      merged.updated = rec.updated; merged.by = rec.by;
      rec = merged;
    }
  }
  if (old && old.selfToken && !rec.selfToken) rec.selfToken = old.selfToken;
  if (isSuper_(user)) rec.unitId = rec.unitId || (old && old.unitId) || user.unitId || '';
  else rec.unitId = user.unitId;
  flat = Object.assign({}, flat || {}, { 'หน่วยบริการ': unitName_(rec.unitId) });
  return upsert_(SHEET_STORES, rec, flat, user);
}
function saveInsp_(user, rec, flat) {
  if (!rec || !rec.id) throw new Error('ข้อมูลไม่ครบ');
  needUnit_(user);
  needPerm_(user, 'inspect');
  const st = findRec_(SHEET_STORES, rec.storeId);
  if (!st) throw new Error('ไม่พบร้านค้าของผลตรวจนี้ (อาจถูกลบแล้ว)');
  if (!sameUnit_(user, st.unitId)) throw err_('FORBIDDEN', 'ร้านนี้อยู่ในหน่วยบริการอื่น');
  rec.unitId = st.unitId || '';
  flat = Object.assign({}, flat || {}, { 'หน่วยบริการ': unitName_(rec.unitId) });
  return upsert_(SHEET_INSP, rec, flat, user);
}
function deleteStore_(user, id) {
  const st = findRec_(SHEET_STORES, id);
  if (!st) return { ok: true, deleted: 0 };
  if (!sameUnit_(user, st.unitId)) throw err_('FORBIDDEN', 'ลบได้เฉพาะร้านค้าในหน่วยบริการของคุณ');
  deleteRows_(SHEET_INSP, r => r.storeId === id);
  return deleteRows_(SHEET_STORES, r => r.id === id);
}
function deleteInsp_(user, id) {
  const r = findRec_(SHEET_INSP, id);
  if (!r) return { ok: true, deleted: 0 };
  const st = findRec_(SHEET_STORES, r.storeId) || {};
  const unitId = r.unitId || st.unitId;
  const ok = isSuper_(user) || (sameUnit_(user, unitId) && (user.role === 'unitadmin' || !r.createdBy || r.createdBy === user.email));
  if (!ok) throw err_('FORBIDDEN', 'ลบได้เฉพาะผลตรวจที่ตนเองบันทึก หรือผู้ดูแลหน่วยบริการ');
  return deleteRows_(SHEET_INSP, x => x.id === id);
}

/* ================= ร้านค้ากรอกข้อมูลเอง / ประเมินตนเอง ================= */
function selfSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_SELF);
  if (!sh) { sh = ss.insertSheet(SHEET_SELF); sh.getRange(1, 1, 1, SELF_COLS.length).setValues([SELF_COLS]).setFontWeight('bold').setBackground('#e9d5ff'); sh.setFrozenRows(1); }
  return sh;
}
function readSelf_() {
  const sh = selfSheet_(); const lr = sh.getLastRow(); if (lr < 2) return [];
  return sh.getRange(2, 1, lr - 1, SELF_COLS.length).getValues().map((r, i) => {
    const o = { _row: i + 2 }; SELF_COLS.forEach((k, j) => o[k] = r[j]);
    try { o.data = JSON.parse(o.json || '{}'); } catch (e) { o.data = {}; }
    delete o.json; o.created = o.created ? new Date(o.created).getTime() : 0; o.reviewedAt = o.reviewedAt ? new Date(o.reviewedAt).getTime() : 0; return o;
  });
}
function selfToken_() { return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8); }
function storeByToken_(token) { if (!token || String(token).length < 20) return null; return readAll_(SHEET_STORES).filter(s => s.selfToken === token)[0] || null; }
function storeLink_(user, storeId) {
  const st = findRec_(SHEET_STORES, storeId);
  if (!st) throw new Error('ไม่พบร้านค้า (ถ้าเพิ่งสร้าง กรุณาซิงก์ก่อน)');
  if (!sameUnit_(user, st.unitId)) throw err_('FORBIDDEN', 'ร้านนี้อยู่ในหน่วยบริการอื่น');
  if (!st.selfToken) { st.selfToken = selfToken_(); rewriteJson_(SHEET_STORES, st); }
  return { ok: true, token: st.selfToken, url: APP_URL + '#self/' + st.selfToken };
}
const SELF_FIELDS = ['name', 'type', 'kind', 'owner', 'caretaker', 'phone', 'address', 'village', 'moo', 'license', 'lat', 'lng'];
function selfGet_(token) {
  const st = storeByToken_(token);
  if (!st) throw new Error('ลิงก์ไม่ถูกต้องหรือหมดอายุ กรุณาขอลิงก์ใหม่จากเจ้าหน้าที่');
  const info = {}; SELF_FIELDS.concat(['photo', 'closed']).forEach(k => info[k] = st[k] === undefined ? '' : st[k]);
  const pend = readSelf_().filter(r => r.storeId === st.id && r.status === 'pending').length;
  return { ok: true, store: info, unitName: unitName_(st.unitId), pending: pend };
}
function selfUnit_(unitId) {
  const u = readUnits_().filter(x => x.id === unitId)[0];
  if (!u) throw new Error('ลิงก์ลงทะเบียนไม่ถูกต้อง');
  // รายชื่อร้านในหน่วย (เฉพาะชื่อ/หมู่บ้าน) ให้ร้านที่มีในระบบแล้วค้นหาชื่อตัวเองได้ — ไม่เปิดเผยข้อมูลส่วนตัว
  const stores = readAll_(SHEET_STORES).filter(s => String(s.unitId) === String(u.id) && !s.closed)
    .map(s => ({ id: s.id, name: s.name || '', village: s.village || '', moo: s.moo || '' }));
  return { ok: true, unit: { id: u.id, name: u.name, code: u.code }, stores: stores };
}
function cleanSelf_(d) {
  const info = {}; SELF_FIELDS.forEach(k => { if (d.info && d.info[k] !== undefined) info[k] = String(d.info[k]).slice(0, 300); });
  if (info.lat !== undefined && info.lat !== '' && isNaN(Number(info.lat))) delete info.lat;
  if (info.lng !== undefined && info.lng !== '' && isNaN(Number(info.lng))) delete info.lng;
  const a = d.assessment || null;
  const assess = a ? { answers: a.answers || {}, notes: a.notes || {}, sources: a.sources || [], sourceOther: String(a.sourceOther || '').slice(0, 200),
    training: a.training || '', trainingYear: String(a.trainingYear || '').slice(0, 10), understanding: a.understanding || '', comment: String(a.comment || '').slice(0, 1000) } : null;
  return { info: info, assessment: assess, contactName: String(d.contactName || '').slice(0, 100), contactPhone: String(d.contactPhone || '').slice(0, 30) };
}
function selfSubmit_(req) {
  const c = CacheService.getScriptCache();
  const key = 'self_' + (req.token || req.storeId || (req.unitId + '_' + String((req.info || {}).name || '').trim()));
  if (c.get(key)) throw new Error('เพิ่งส่งข้อมูลไปแล้ว กรุณารอสักครู่แล้วลองใหม่');
  let st = null, unitId = '', type = 'update';
  if (req.token) { st = storeByToken_(req.token); if (!st) throw new Error('ลิงก์ไม่ถูกต้อง'); unitId = st.unitId; }
  else if (req.unitId && req.storeId) {   // ร้านที่มีในระบบแล้ว ส่งผ่าน QR ของหน่วยบริการ
    st = findRec_(SHEET_STORES, req.storeId);
    if (!st || String(st.unitId) !== String(req.unitId)) throw new Error('ไม่พบร้านค้านี้ในหน่วยบริการ');
    unitId = st.unitId;
  }
  else { const u = readUnits_().filter(x => x.id === req.unitId)[0]; if (!u) throw new Error('ลิงก์ลงทะเบียนไม่ถูกต้อง'); unitId = u.id; type = 'new'; }
  const d = cleanSelf_(req);
  if (type === 'new' && !d.info.name) throw new Error('กรุณากรอกชื่อร้านค้า');
  if (req.photo) { if (String(req.photo).length > 4e6) throw new Error('รูปมีขนาดใหญ่เกินไป'); d.info.photo = uploadPhoto_({ data: req.photo, name: 'self_' + Date.now() + '.jpg' }).url; }
  if (!Object.keys(d.info).length && !d.assessment) throw new Error('ไม่มีข้อมูลที่จะส่ง');
  const id = 'sr' + Utilities.getUuid().replace(/-/g, '').slice(0, 12);
  selfSheet_().appendRow([id, new Date(), type, 'pending', unitId, st ? st.id : '', st ? st.name : d.info.name, (d.contactName + ' ' + d.contactPhone).trim(), '', '', JSON.stringify(d)]);
  c.put(key, '1', 30);
  sysMsg_('unit:' + unitId + ':staff', '📥 ' + (type === 'new' ? 'ร้านค้าใหม่ลงทะเบียน: ' : 'ร้านค้าส่งข้อมูล: ') + (st ? st.name : d.info.name),
    (type === 'new' ? 'มีร้านค้าใหม่ลงทะเบียนผ่านลิงก์ของหน่วยบริการ' : 'ร้านส่งข้อมูล' + (d.assessment ? 'และผลประเมินตนเอง' : '')) + '\nผู้ส่ง: ' + ((d.contactName + ' ' + d.contactPhone).trim() || '-') + '\nกรุณาตรวจรับในกล่อง “ข้อมูลจากร้านค้า”', '#inbox');
  try {
    const to = readUsers_().filter(x => x.status === 'approved' && x.unitId === unitId && ['unitadmin', 'staff'].indexOf(x.role) >= 0).map(x => x.email);
    if (to.length) MailApp.sendEmail({ to: to.join(','), name: APP_NAME, subject: '📥 ' + APP_NAME + ': ' + (type === 'new' ? 'ร้านค้าใหม่ลงทะเบียน' : 'ร้านค้าส่งข้อมูล/ประเมินตนเอง') + ' — ' + (st ? st.name : d.info.name),
      htmlBody: '<div style="font-family:sans-serif;font-size:15px">' + (type === 'new' ? 'มีร้านค้าใหม่ลงทะเบียนผ่านลิงก์ของหน่วยบริการ' : 'ร้าน <b>' + st.name + '</b> ส่งข้อมูล' + (d.assessment ? 'และผลประเมินตนเอง' : '')) +
        '<br>ผู้ส่ง: ' + (d.contactName || '-') + ' ' + (d.contactPhone || '') + '<br><br><a href="' + APP_URL + '#inbox" style="background:#0f766e;color:#fff;padding:10px 18px;border-radius:8px;text-decoration:none;font-weight:bold">เปิดกล่องข้อมูลจากร้านค้า</a></div>' });
  } catch (e) { }
  return { ok: true, id: id };
}
function selfList_(user) {
  return readSelf_().filter(r => sameUnit_(user, r.unitId)).sort((a, b) => (a.status === 'pending' ? 0 : 1) - (b.status === 'pending' ? 0 : 1) || b.created - a.created).slice(0, 200);
}
function selfReview_(user, req) {
  const r = readSelf_().filter(x => x.id === req.id)[0];
  if (!r) throw new Error('ไม่พบรายการ');
  if (!sameUnit_(user, r.unitId)) throw err_('FORBIDDEN', 'รายการนี้อยู่ในหน่วยบริการอื่น');
  if (r.status !== 'pending') throw new Error('รายการนี้ตรวจรับไปแล้ว');
  const sh = selfSheet_();
  const res = { ok: true };
  if (req.decision === 'accept') {
    const d = r.data || {}; const now = Date.now();
    let st;
    if (r.type === 'new') {
      st = Object.assign({ id: 's' + Utilities.getUuid().replace(/-/g, '').slice(0, 12), village: '', name: '' }, d.info || {});
      st.unitId = r.unitId; st.created = now; st.selfToken = selfToken_(); st.by = 'ร้านค้าลงทะเบียนเอง';
      if (st.lat !== undefined && st.lat !== '') st.lat = Number(st.lat); if (st.lng !== undefined && st.lng !== '') st.lng = Number(st.lng);
    } else {
      st = findRec_(SHEET_STORES, r.storeId); if (!st) throw new Error('ไม่พบร้านค้า (อาจถูกลบแล้ว)');
      if (req.applyInfo !== false) Object.keys(d.info || {}).forEach(k => { const v = d.info[k]; if (v !== '' && v !== undefined) st[k] = (k === 'lat' || k === 'lng') ? Number(v) : v; });
    }
    st.updated = now;
    upsert_(SHEET_STORES, st, { 'หน่วยบริการ': unitName_(st.unitId), 'ชื่อร้าน': st.name || '', 'หมู่บ้าน': st.village || '' }, user);
    res.store = st;
    if (d.assessment && req.applyAssessment !== false) {
      const ins = Object.assign({ id: 'self' + Utilities.getUuid().replace(/-/g, '').slice(0, 10), storeId: st.id, unitId: st.unitId, type: 'self',
        date: Utilities.formatDate(new Date(r.created || now), 'Asia/Bangkok', 'yyyy-MM-dd'), inspector: 'ร้านค้าประเมินตนเอง' + (d.contactName ? ' (' + d.contactName + ')' : ''),
        tests: {}, photos: [], created: now, updated: now }, d.assessment);
      upsert_(SHEET_INSP, ins, { 'รหัสร้าน': st.id, 'ชื่อร้าน': st.name || '', 'วันที่ตรวจ': ins.date, 'ผู้ตรวจ': ins.inspector, 'ประเภท': 'ประเมินตนเอง', 'หน่วยบริการ': unitName_(st.unitId) }, user);
      res.inspection = ins;
    }
  }
  sh.getRange(r._row, 4).setValue(req.decision === 'accept' ? 'accepted' : 'rejected');
  sh.getRange(r._row, 9, 1, 2).setValues([[user.email, new Date()]]);
  return res;
}

/* ================= กล่องจดหมาย ================= */
// ผู้รับ: อีเมล | 'all' | 'super' | 'unit:<id>' (ทุกคนในหน่วย) | 'unit:<id>:staff' (เจ้าหน้าที่+ผู้ดูแลหน่วย) | 'unit:<id>:osm' | 'unit:<id>:admin'
function permText_(p) { return String(p || '').split(',').filter(k => OSM_PERMS[k]).map(k => OSM_PERMS[k]).join(', ') || '-'; }
function msgSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_MSG);
  if (!sh) { sh = ss.insertSheet(SHEET_MSG); sh.getRange(1, 1, 1, MSG_COLS.length).setValues([MSG_COLS]).setFontWeight('bold').setBackground('#bfdbfe'); sh.setFrozenRows(1); }
  return sh;
}
function readMsgs_() {
  const sh = msgSheet_(); const lr = sh.getLastRow(); if (lr < 2) return [];
  const n = Math.min(lr - 1, 3000), start = lr - n + 1;   // อ่านล่าสุด 3,000 ข้อความ
  return sh.getRange(start, 1, n, MSG_COLS.length).getValues().map((r, i) => {
    const o = { _row: start + i }; MSG_COLS.forEach((k, j) => o[k] = r[j]);
    o.created = o.created ? new Date(o.created).getTime() : 0; o.from = norm_(o.from); o.to = String(o.to || ''); return o;
  });
}
function hasIn_(list, email) { return (',' + String(list || '') + ',').indexOf(',' + email + ',') >= 0; }
function isFor_(m, u) {
  const t = m.to;
  if (t === u.email || t === 'all') return true;
  if (t === 'super') return isSuper_(u);
  const p = t.split(':'); if (p[0] !== 'unit') return false;
  if (p[2] === 'admin') return isSuper_(u) && !hasAdmin_(p[1]) || (u.role === 'unitadmin' && u.unitId === p[1]);
  if (u.unitId !== p[1]) return false;
  if (!p[2]) return true;
  if (p[2] === 'staff') return u.role === 'staff' || u.role === 'unitadmin';
  if (p[2] === 'osm') return u.role === 'osm';
  return false;
}
let ADMIN_CACHE_ = null;
function hasAdmin_(unitId) {   // หน่วยที่ยังไม่มีผู้ดูแลหน่วย → ส่งถึงผู้ดูแลระบบสูงสุดแทน
  if (!ADMIN_CACHE_) { ADMIN_CACHE_ = {}; readUsers_().filter(x => x.role === 'unitadmin' && x.status === 'approved').forEach(x => ADMIN_CACHE_[x.unitId] = 1); }
  return !!ADMIN_CACHE_[unitId];
}
function toLabel_(to) {
  if (to === 'all') return 'ทุกคนในระบบ';
  if (to === 'super') return 'ผู้ดูแลระบบสูงสุด';
  const p = to.split(':');
  if (p[0] === 'unit') { const n = unitName_(p[1]) || 'หน่วยบริการ'; return ({ '': 'ทุกคนใน ', staff: 'เจ้าหน้าที่ ', osm: 'อสม. ', admin: 'ผู้ดูแล ' })[p[2] || ''] + n; }
  const u = findUser_(to); return u ? (u.name || u.email) : to;
}
function sysMsg_(to, subject, body, link) {
  try { msgSheet_().appendRow(['m' + Utilities.getUuid().replace(/-/g, '').slice(0, 12), new Date(), 'system', 'ระบบแจ้งเตือน', to, toLabel_(to), String(subject).slice(0, 200), String(body || '').slice(0, 3000), link || '', 'sys', '', '', '']); } catch (e) { }
}
function msgPublic_(m, u) {
  return { id: m.id, created: m.created, from: m.from, fromName: m.fromName, to: m.to, toLabel: m.toLabel, subject: m.subject, body: m.body, link: m.link, kind: m.kind, replyTo: m.replyTo,
    read: m.from === u.email || hasIn_(m.readBy, u.email), mine: m.from === u.email };
}
function msgList_(u) {
  const all = readMsgs_().filter(m => !hasIn_(m.hiddenBy, u.email) && (m.from === u.email || isFor_(m, u)));
  all.sort((a, b) => b.created - a.created);
  return { ok: true, messages: all.slice(0, 300).map(m => msgPublic_(m, u)), unread: all.filter(m => m.from !== u.email && !hasIn_(m.readBy, u.email)).length };
}
function msgUnread_(u) { try { return readMsgs_().filter(m => m.from !== u.email && isFor_(m, u) && !hasIn_(m.readBy, u.email) && !hasIn_(m.hiddenBy, u.email)).length; } catch (e) { return 0; } }
function msgTargets_(u) {   // กลุ่มที่ส่งถึงได้
  const t = [];
  if (isSuper_(u)) { t.push(['all', 'ทุกคนในระบบ']); readUnits_().forEach(x => { t.push(['unit:' + x.id, 'ทุกคนใน ' + x.name]); }); }
  else if (u.unitId) {
    const n = unitName_(u.unitId);
    t.push(['unit:' + u.unitId, 'ทุกคนใน ' + n]);
    t.push(['unit:' + u.unitId + ':staff', 'เจ้าหน้าที่ ' + n]);
    if (u.role !== 'osm') t.push(['unit:' + u.unitId + ':osm', 'อสม. ' + n]);
    t.push(['unit:' + u.unitId + ':admin', 'ผู้ดูแล ' + n]);
  }
  t.push(['super', 'ผู้ดูแลระบบสูงสุด']);
  return t;
}
function msgContacts_(u) {
  const people = readUsers_().filter(x => x.status === 'approved' && x.email !== u.email && (isSuper_(u) || x.role === 'superadmin' || (u.unitId && x.unitId === u.unitId)))
    .map(x => ({ email: x.email, name: x.name || x.email, role: x.role, unit: x.unit || '' }));
  return { ok: true, groups: msgTargets_(u).map(x => ({ to: x[0], label: x[1] })), people: people };
}
function msgSend_(u, req) {
  let to = String(req.to || '').trim(); if (to.indexOf('@') > 0) to = norm_(to);
  const subject = String(req.subject || '').trim(), body = String(req.body || '').trim();
  if (!to) throw new Error('กรุณาเลือกผู้รับ');
  if (!subject && !body) throw new Error('กรุณาพิมพ์ข้อความ');
  const c = CacheService.getScriptCache(); const k = 'msg_' + u.email; const n = Number(c.get(k) || 0);
  if (n >= 20) throw new Error('ส่งข้อความถี่เกินไป กรุณารอสักครู่');
  c.put(k, String(n + 1), 60);
  let ok = false, toReal = to;
  if (to.indexOf('@') > 0) {
    const r = findUser_(to);
    ok = !!r && r.status === 'approved' && (isSuper_(u) || r.role === 'superadmin' || (u.unitId && r.unitId === u.unitId));
    if (!ok && req.replyTo) { const orig = readMsgs_().filter(m => m.id === req.replyTo)[0]; ok = !!orig && orig.from === to && isFor_(orig, u); }
  } else {
    ok = msgTargets_(u).some(x => x[0] === to);
  }
  if (!ok) throw err_('FORBIDDEN', 'ส่งข้อความถึงผู้รับนี้ไม่ได้');
  const id = 'm' + Utilities.getUuid().replace(/-/g, '').slice(0, 12);
  const link = /^#(store|report)\/[\w-]+$/.test(req.link || '') ? req.link : '';
  msgSheet_().appendRow([id, new Date(), u.email, u.name || u.email, toReal, toLabel_(toReal), (subject || body.slice(0, 60)).slice(0, 200), body.slice(0, 5000), link, 'msg', String(req.replyTo || '').slice(0, 20), '', '']);
  if (req.email && toReal.indexOf('@') > 0) { try { MailApp.sendEmail({ to: toReal, name: APP_NAME, subject: '📬 ' + APP_NAME + ': ข้อความจาก ' + (u.name || u.email) + ' — ' + subject, body: body + '\n\nตอบกลับในแอป: ' + APP_URL + '#mail' }); } catch (e) { } }
  return { ok: true, id: id };
}
function msgMark_(u, ids, col) {
  if (!ids.length) return { ok: true };
  const set = {}; ids.forEach(i => set[i] = 1);
  const sh = msgSheet_(); const ci = MSG_COLS.indexOf(col) + 1;
  readMsgs_().forEach(m => {
    if (!set[m.id] || hasIn_(m[col], u.email)) return;
    if (!(m.from === u.email || isFor_(m, u))) return;
    sh.getRange(m._row, ci).setValue((m[col] ? m[col] + ',' : '') + u.email);
  });
  return { ok: true, unread: msgUnread_(u) };
}

/* ================= ตาราง ================= */
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
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
function headers_(sh) { const lc = sh.getLastColumn(); return lc ? sh.getRange(1, 1, 1, lc).getValues()[0].map(String) : []; }
function readAll_(name) {
  const sh = getSheet_(name);
  const lr = sh.getLastRow();
  if (lr < 2) return [];
  const jc = headers_(sh).indexOf('json');
  if (jc < 0) return [];
  return sh.getRange(2, jc + 1, lr - 1, 1).getValues().map(r => { try { return JSON.parse(r[0]); } catch (e) { return null; } }).filter(Boolean);
}
function upsert_(name, rec, flat, user) {
  const sh = getSheet_(name);
  const h = headers_(sh);
  const jc = h.indexOf('json');
  const lr = sh.getLastRow();
  const data = lr > 1 ? sh.getRange(2, 1, lr - 1, Math.max(h.length, 2)).getValues() : [];
  let idx = -1;
  for (let i = 0; i < data.length; i++) if (String(data[i][0]) === String(rec.id)) { idx = i; break; }
  if (idx >= 0) {
    if (Number(data[idx][1]) > Number(rec.updated || 0)) return { ok: true, skipped: 'server newer' };
    try { const old = JSON.parse(data[idx][jc]); if (old.createdBy) rec.createdBy = old.createdBy; } catch (e) { }
  }
  if (!rec.createdBy) rec.createdBy = user.email;
  rec.updatedBy = user.email;
  const row = Object.assign({ id: rec.id, updated: rec.updated || Date.now() }, flat || {});
  row['แก้ไขล่าสุด'] = new Date(rec.updated || Date.now());
  row['บันทึกโดย (อีเมล)'] = user.email;
  row.json = JSON.stringify(rec);
  let changed = false;
  Object.keys(row).forEach(k => { if (h.indexOf(k) < 0) { h.push(k); changed = true; } });
  if (changed) sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#ccfbf1');
  const values = h.map(k => {
    const v = row[k];
    if (v === undefined || v === null) return '';
    if (typeof v === 'object' && !(v instanceof Date)) return JSON.stringify(v);
    if (typeof v === 'string' && /^[=+\-@]/.test(v) && k !== 'json') return "'" + v;
    return v;
  });
  if (idx >= 0) sh.getRange(idx + 2, 1, 1, h.length).setValues([values]);
  else sh.getRange(lr + 1, 1, 1, h.length).setValues([values]);
  return { ok: true, record: rec };
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
function getFolder_() {
  const props = PropertiesService.getScriptProperties();
  const fid = props.getProperty('PHOTO_FOLDER_ID');
  if (fid) { try { return DriveApp.getFolderById(fid); } catch (e) { } }
  const parents = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()).getParents();
  const parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const f = parent.createFolder(PHOTO_FOLDER);
  props.setProperty('PHOTO_FOLDER_ID', f.getId());
  return f;
}
function uploadPhoto_(req) {
  if (!req.data) throw new Error('ไม่มีข้อมูลรูป');
  if (req.data.length > 4e6) throw new Error('รูปมีขนาดใหญ่เกินไป');
  const blob = Utilities.newBlob(Utilities.base64Decode(req.data), 'image/jpeg', String(req.name || ('photo_' + Date.now() + '.jpg')).replace(/[^\w.\-]/g, '_'));
  const file = getFolder_().createFile(blob);
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { }
  return { ok: true, id: file.getId(), url: 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1280' };
}
