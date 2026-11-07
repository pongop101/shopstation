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
const PHOTO_FOLDER = 'ShopStation Photos';
const USER_COLS = ['email', 'name', 'position', 'unit', 'phone', 'role', 'status', 'picture', 'created', 'lastLogin', 'approvedBy', 'note', 'unitId', 'transferTo', 'transferBy', 'transferAt'];
const UNIT_COLS = ['id', 'name', 'created', 'createdBy', 'code'];
const ROLES = { superadmin: 'ผู้ดูแลระบบสูงสุด', unitadmin: 'ผู้ดูแลหน่วยบริการ', staff: 'เจ้าหน้าที่', viewer: 'ดูอย่างเดียว' };

function setup() {
  [SHEET_STORES, SHEET_INSP].forEach(n => getSheet_(n));
  usersSheet_();
  unitsSheet_();
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
  try {
    const req = JSON.parse(e.postData.contents);
    const a = req.action;
    // ---- ไม่ต้องเข้าระบบ ----
    if (a === 'ping') return json_({ ok: true, app: APP_NAME, sheet: SpreadsheetApp.getActiveSpreadsheet().getName(), googleClientId: GOOGLE_CLIENT_ID, otp: true,
      units: readUnits_().map(u => ({ id: u.id, name: u.name, code: u.code })) });
    if (a === 'loginGoogle') return json_(loginGoogle_(req.credential));
    if (a === 'otpSend') return json_(otpSend_(req.email));
    if (a === 'otpVerify') return json_(otpVerify_(req.email, req.code));

    // ---- ต้องเข้าระบบ ----
    const user = auth_(req.token);
    if (a === 'me') return json_({ ok: true, user: publicUser_(user) });
    if (a === 'register') return json_(register_(user, req.profile || {}));
    if (user.status !== 'approved') throw err_('NOT_APPROVED', 'บัญชียังไม่ได้รับการอนุมัติจากผู้ดูแลระบบ');

    switch (a) {
      case 'list': return json_(list_(user));
      case 'saveStore': need_(user, ['superadmin', 'unitadmin', 'staff']); return json_(withLock_(() => saveStore_(user, req.record, req.flat)));
      case 'saveInspection': need_(user, ['superadmin', 'unitadmin', 'staff']); return json_(withLock_(() => saveInsp_(user, req.record, req.flat)));
      case 'deleteStore': need_(user, ['superadmin', 'unitadmin']); return json_(withLock_(() => deleteStore_(user, req.id)));
      case 'deleteInspection': need_(user, ['superadmin', 'unitadmin', 'staff']); return json_(withLock_(() => deleteInsp_(user, req.id)));
      case 'uploadPhoto': need_(user, ['superadmin', 'unitadmin', 'staff']); return json_(uploadPhoto_(req));
      case 'users': need_(user, ['superadmin', 'unitadmin']); return json_({ ok: true, users: visibleUsers_(user).map(publicUser_), units: readUnits_() });
      case 'userUpdate': need_(user, ['superadmin', 'unitadmin']); return json_(withLock_(() => userUpdate_(user, req.email, req.changes || {})));
      case 'userDelete': need_(user, ['superadmin', 'unitadmin']); return json_(withLock_(() => userDelete_(user, req.email)));
      case 'transferRequest': return json_(withLock_(() => transferRequest_(user, req.email, req.unitId)));
      case 'transferCancel': return json_(withLock_(() => transferCancel_(user, req.email)));
      case 'transferDecide': need_(user, ['superadmin', 'unitadmin']); return json_(withLock_(() => transferDecide_(user, req.email, !!req.approve)));
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
    if (isNew && u.status === 'pending') notifyAdmins_(u);
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
  return isSuper_(admin) ? all : all.filter(u => u.role !== 'superadmin' && ((u.unitId && String(u.unitId) === String(admin.unitId)) || (u.transferTo && String(u.transferTo) === String(admin.unitId))));
}
function canManage_(admin, u) {
  if (isSuper_(admin)) return true;
  return admin.role === 'unitadmin' && !!admin.unitId && String(u.unitId) === String(admin.unitId) && u.role !== 'superadmin' && u.role !== 'unitadmin';
}
function userUpdate_(admin, email, ch) {
  email = norm_(email);
  const u = findUser_(email);
  if (!u) throw new Error('ไม่พบผู้ใช้');
  if (email === admin.email) throw new Error('แก้ไขสิทธิ์หรือสถานะของบัญชีตัวเองไม่ได้');
  if (!canManage_(admin, u)) throw err_('FORBIDDEN', 'จัดการได้เฉพาะเจ้าหน้าที่ในหน่วยบริการของคุณ');
  if (isSuperEmail_(email)) throw new Error('แก้ไขผู้ดูแลระบบสูงสุดไม่ได้');
  if (ch.role) {
    if (!ROLES[ch.role]) throw new Error('บทบาทไม่ถูกต้อง');
    if (!isSuper_(admin) && ['staff', 'viewer'].indexOf(ch.role) < 0) throw err_('FORBIDDEN', 'ผู้ดูแลหน่วยบริการกำหนดได้เฉพาะสิทธิ์ เจ้าหน้าที่ / ดูอย่างเดียว');
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
  try { MailApp.sendEmail({ to: u.email, name: APP_NAME, subject: APP_NAME + ': ' + subject, body: body + '\n\nเปิดแอป: ' + APP_URL }); } catch (e) { }
}
function userDelete_(admin, email) {
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
    transferTo: String(u.transferTo || ''), transferBy: u.transferBy || '', transferAt: u.transferAt ? new Date(u.transferAt).getTime() : 0,
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
  return { ok: true, user: publicUser_(user), stores: stores, inspections: insps, pending: pending, units: readUnits_().map(u => ({ id: u.id, name: u.name, code: u.code })) };
}
function findRec_(name, id) { return readAll_(name).filter(r => r.id === id)[0] || null; }
function saveStore_(user, rec, flat) {
  if (!rec || !rec.id) throw new Error('ข้อมูลไม่ครบ');
  needUnit_(user);
  const old = findRec_(SHEET_STORES, rec.id);
  if (old && !sameUnit_(user, old.unitId)) throw err_('FORBIDDEN', 'ร้านนี้อยู่ในหน่วยบริการอื่น');
  if (isSuper_(user)) rec.unitId = rec.unitId || (old && old.unitId) || user.unitId || '';
  else rec.unitId = user.unitId;
  flat = Object.assign({}, flat || {}, { 'หน่วยบริการ': unitName_(rec.unitId) });
  return upsert_(SHEET_STORES, rec, flat, user);
}
function saveInsp_(user, rec, flat) {
  if (!rec || !rec.id) throw new Error('ข้อมูลไม่ครบ');
  needUnit_(user);
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
