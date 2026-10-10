/**
 * ระบบลงรับและติดตามเอกสาร (ไปราชการ/โครงการ) — Google Apps Script
 * โหมด GitHub Pages: ไฟล์นี้เป็นแค่ API (ไม่ต้องมีไฟล์ Index.html)
 * วางไฟล์นี้ในโปรเจกต์ที่ผูกกับ Google Sheet (ส่วนขยาย > Apps Script)
 * ข้อมูลเก็บในชีตซ่อน "_docs" (ห้ามแก้ด้วยมือ) และสำเนาอ่านง่ายในชีต "รายการ"
 */
const DOCS_SHEET = '_docs';
const VIEW_SHEET = 'รายการ';
const SESSION_SECONDS = 6 * 60 * 60; // เข้าสู่ระบบอยู่ได้ 6 ชั่วโมง
const TZ = 'Asia/Bangkok';

// รับคำสั่งจากหน้าเว็บบน GitHub Pages (ส่งมาแบบ POST ข้อความ JSON)
const ALLOWED = { rpc: rpc, login: login, me: me, logout: logout, changePassword: changePassword,
  listUsers: listUsers, addUser: addUser, resetPassword: resetPassword, deleteUser: deleteUser, dump: dump };

function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    const f = ALLOWED[req.fn];
    if (!f) throw new Error('ไม่อนุญาต');
    out = { ok: true, result: f.apply(null, req.args || []) };
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// เปิดลิงก์ /exec ในเบราว์เซอร์จะเห็นข้อความนี้ แปลว่า Web App ทำงานอยู่
function doGet() {
  return ContentService.createTextOutput(JSON.stringify({ ok: true, service: 'travel-docs-api' }))
    .setMimeType(ContentService.MimeType.JSON);
}

/** รันครั้งแรกด้วยมือ 1 ครั้ง เพื่ออนุญาตสิทธิ์และสร้างชีต/ผู้ใช้เริ่มต้น */
function setup() {
  docsSheet_();
  viewSheet_();
  ensureAdmin_();
  Logger.log('พร้อมใช้งาน — ผู้ใช้เริ่มต้น admin / 1234 (ให้เปลี่ยนรหัสผ่านทันที)');
}

// ---------- ที่เก็บเอกสาร (ชีต _docs: path | json) ----------
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function docsSheet_() {
  let s = ss_().getSheetByName(DOCS_SHEET);
  if (!s) {
    s = ss_().insertSheet(DOCS_SHEET);
    s.getRange('A:B').setNumberFormat('@');
    s.getRange(1, 1, 1, 2).setValues([['path', 'json']]);
    s.setFrozenRows(1);
    s.hideSheet();
  }
  return s;
}

function findRow_(path) {
  const s = docsSheet_();
  const f = s.getRange('A:A').createTextFinder(path).matchEntireCell(true).findNext();
  return f ? f.getRow() : 0;
}

function getDoc_(path) {
  const r = findRow_(path);
  return r ? JSON.parse(docsSheet_().getRange(r, 2).getValue()) : null;
}

function putDoc_(path, obj) {
  const s = docsSheet_(), r = findRow_(path), j = JSON.stringify(obj);
  if (r) s.getRange(r, 2).setValue(j); else s.appendRow([path, j]);
}

function listDocs_(col) {
  const s = docsSheet_(), n = s.getLastRow();
  if (n < 2) return [];
  const pre = col + '/';
  return s.getRange(2, 1, n - 1, 2).getValues()
    .filter(function (x) { return String(x[0]).indexOf(pre) === 0; })
    .map(function (x) { return { id: String(x[0]).slice(pre.length), data: JSON.parse(x[1]) }; });
}

function deleteDoc_(path) {
  const r = findRow_(path);
  if (r) docsSheet_().deleteRow(r);
}

function merge_(a, b) {
  if (b === null || typeof b !== 'object' || Array.isArray(b)) return b;
  const o = (a && typeof a === 'object' && !Array.isArray(a)) ? JSON.parse(JSON.stringify(a)) : {};
  Object.keys(b).forEach(function (k) { o[k] = merge_(o[k], b[k]); });
  return o;
}

// ---------- ผู้ใช้และการเข้าสู่ระบบ (ตรวจที่เซิร์ฟเวอร์) ----------
function hash_(salt, pw) {
  const d = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + pw, Utilities.Charset.UTF_8);
  return d.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

function roleOf_(d, u) {
  return (d && d.role === 'admin') || u === 'admin' ? 'admin' : 'staff';
}

function ensureAdmin_() {
  if (getDoc_('auth/u-admin')) return;
  const salt = Utilities.getUuid();
  putDoc_('auth/u-admin', { user: 'admin', role: 'admin', salt: salt, hash: hash_(salt, '1234'), must: true,
    created: new Date().toISOString() });
}

function login(u, p) {
  u = String(u || '').trim().toLowerCase();
  const c = CacheService.getScriptCache();
  if (c.get('lock:' + u)) return { ok: false, msg: 'ใส่รหัสผิดหลายครั้ง กรุณารอ 5 นาทีแล้วลองใหม่' };
  if (/^[a-z0-9_-]{1,30}$/.test(u)) {
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    let d;
    try { ensureAdmin_(); d = getDoc_('auth/u-' + u); } finally { lock.releaseLock(); }
    if (d && hash_(d.salt, String(p || '')) === d.hash) {
      c.remove('f:' + u);
      const t = Utilities.getUuid() + Utilities.getUuid();
      c.put('t:' + t, u, SESSION_SECONDS);
      return { ok: true, token: t, user: u, must: !!d.must, role: roleOf_(d, u) };
    }
  }
  const n = Number(c.get('f:' + u) || 0) + 1;
  c.put('f:' + u, String(n), 600);
  if (n >= 5) { c.put('lock:' + u, '1', 300); c.remove('f:' + u); }
  return { ok: false, msg: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };
}

function requireAuth_(token) {
  const u = token ? CacheService.getScriptCache().get('t:' + token) : null;
  if (!u || !getDoc_('auth/u-' + u)) throw new Error('SESSION_EXPIRED');
  return u;
}

function requireAdmin_(token) {
  const u = requireAuth_(token);
  if (roleOf_(getDoc_('auth/u-' + u), u) !== 'admin') throw new Error('FORBIDDEN');
  return u;
}

function me(token) {
  const u = token ? CacheService.getScriptCache().get('t:' + token) : null;
  if (!u) return null;
  const d = getDoc_('auth/u-' + u);
  return d ? { user: u, must: !!d.must, role: roleOf_(d, u) } : null;
}

function logout(token) {
  if (token) CacheService.getScriptCache().remove('t:' + token);
  return true;
}

function changePassword(token, oldPw, newPw) {
  const u = requireAuth_(token);
  oldPw = String(oldPw || ''); newPw = String(newPw || '');
  if (newPw.length < 6) return { ok: false, msg: 'รหัสผ่านใหม่ต้องมีอย่างน้อย 6 ตัวอักษร' };
  if (newPw === '1234' || newPw === oldPw) return { ok: false, msg: 'รหัสผ่านใหม่ต้องไม่ซ้ำกับรหัสเดิมหรือรหัสเริ่มต้น' };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = getDoc_('auth/u-' + u);
    if (!d || hash_(d.salt, oldPw) !== d.hash) return { ok: false, msg: 'รหัสผ่านปัจจุบันไม่ถูกต้อง' };
    const salt = Utilities.getUuid();
    d.salt = salt; d.hash = hash_(salt, newPw); d.must = false;
    putDoc_('auth/u-' + u, d);
    return { ok: true };
  } finally { lock.releaseLock(); }
}

// ---------- งานผู้ดูแลระบบ (เฉพาะ admin) ----------
function validPw_(pw) {
  if (String(pw || '').length < 6) return 'รหัสผ่านต้องมีอย่างน้อย 6 ตัวอักษร';
  if (pw === '1234') return 'ห้ามใช้รหัสผ่านเริ่มต้น';
  return '';
}

function listUsers(token) {
  requireAdmin_(token);
  return listDocs_('auth').map(function (x) {
    return { user: x.data.user, role: roleOf_(x.data, x.data.user), must: !!x.data.must, created: x.data.created || '' };
  }).sort(function (a, b) { return a.user === 'admin' ? -1 : b.user === 'admin' ? 1 : a.user < b.user ? -1 : 1; });
}

function addUser(token, u, pw) {
  requireAdmin_(token);
  u = String(u || '').trim().toLowerCase();
  if (!/^[a-z0-9_-]{3,30}$/.test(u)) return { ok: false, msg: 'ชื่อผู้ใช้ใช้ได้เฉพาะ a-z 0-9 _ - ยาว 3–30 ตัว' };
  const bad = validPw_(pw);
  if (bad) return { ok: false, msg: bad };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (getDoc_('auth/u-' + u)) return { ok: false, msg: 'มีชื่อผู้ใช้นี้อยู่แล้ว' };
    const salt = Utilities.getUuid();
    putDoc_('auth/u-' + u, { user: u, role: 'staff', salt: salt, hash: hash_(salt, String(pw)), must: true,
      created: new Date().toISOString() });
    return { ok: true };
  } finally { lock.releaseLock(); }
}

function resetPassword(token, u, pw) {
  requireAdmin_(token);
  u = String(u || '').trim().toLowerCase();
  const bad = validPw_(pw);
  if (bad) return { ok: false, msg: bad };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = getDoc_('auth/u-' + u);
    if (!d) return { ok: false, msg: 'ไม่พบผู้ใช้' };
    const salt = Utilities.getUuid();
    d.salt = salt; d.hash = hash_(salt, String(pw)); d.must = true;
    putDoc_('auth/u-' + u, d);
    CacheService.getScriptCache().remove('lock:' + u);
    return { ok: true };
  } finally { lock.releaseLock(); }
}

function deleteUser(token, u) {
  requireAdmin_(token);
  u = String(u || '').trim().toLowerCase();
  if (u === 'admin') return { ok: false, msg: 'ลบผู้ใช้ admin ไม่ได้' };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (!getDoc_('auth/u-' + u)) return { ok: false, msg: 'ไม่พบผู้ใช้' };
    deleteDoc_('auth/u-' + u);
    return { ok: true };
  } finally { lock.releaseLock(); }
}

// ไฟล์สำรองข้อมูล (ไม่รวมรหัสผ่าน)
function dump(token) {
  requireAdmin_(token);
  const out = { exportedAt: new Date().toISOString() };
  ['requests', 'public', 'meta', 'config'].forEach(function (c) { out[c] = listDocs_(c); });
  return out;
}

// ---------- จุดเรียกข้อมูลจากหน้าเว็บ ----------
// public/*, config/* อ่านได้ทุกคน | config/* เขียน และลบข้อมูล ได้เฉพาะ admin | อื่นๆ ต้องมี token เจ้าหน้าที่ | auth/* ห้ามเข้าถึงจากหน้าเว็บ
function rpc(token, op, path, data) {
  path = String(path || '');
  const col = path.split('/')[0];
  if (!/^[a-z]+(\/[A-Za-z0-9_-]+)?$/.test(path) || col === 'auth') throw new Error('ไม่อนุญาต');
  const open = (col === 'public' || col === 'config') && (op === 'get' || op === 'list');
  if (!open) {
    if (col === 'config' || op === 'delete') requireAdmin_(token); else requireAuth_(token);
  }
  if (op === 'get') return getDoc_(path);
  if (op === 'list') return listDocs_(col);
  if (op === 'delete') {
    if (col !== 'requests' && col !== 'public') throw new Error('ไม่อนุญาต');
    const lk = LockService.getScriptLock();
    lk.waitLock(20000);
    try {
      const old = getDoc_(path);
      deleteDoc_(path);
      if (col === 'requests' && old) unmirror_(old);
      return true;
    } finally { lk.releaseLock(); }
  }
  if (op !== 'set' && op !== 'update') throw new Error('คำสั่งไม่ถูกต้อง');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    let out;
    if (op === 'set') out = data;
    else {
      const cur = getDoc_(path);
      if (!cur) throw new Error('ไม่พบเอกสาร ' + path);
      out = merge_(cur, data);
    }
    putDoc_(path, out);
    if (col === 'requests') mirror_(out);
    return true;
  } finally { lock.releaseLock(); }
}

// ---------- สำเนาอ่านง่ายในชีต "รายการ" ----------
const VIEW_HEAD = ['เลขรับ', 'สถานะ', 'หมวดหมู่', 'ผู้ขอ', 'อีเมล', 'กลุ่มงาน', 'ผู้ร่วมเดินทาง', 'ประเภท/กิจกรรม',
  'ชื่อเรื่อง/โครงการ', 'สถานที่', 'ตั้งแต่วันที่', 'ถึงวันที่', 'จำนวนวัน', 'งบประมาณรวม', 'แหล่งงบ',
  'เลขยืมเงิน', 'เงินยืม', 'กำหนดคืน', 'วันที่ยื่นคืน', 'ค่าใช้จ่ายจริง (คืนเงินยืม)', 'เหตุผลยกเลิก'];

function viewSheet_() {
  let s = ss_().getSheetByName(VIEW_SHEET);
  if (!s) {
    s = ss_().insertSheet(VIEW_SHEET);
    s.getRange('A:A').setNumberFormat('@');
    s.getRange(1, 1, 1, VIEW_HEAD.length).setValues([VIEW_HEAD]).setFontWeight('bold');
    s.setFrozenRows(1);
  }
  return s;
}

function unmirror_(d) {
  try {
    const s = viewSheet_(), no = String(d.no || d.id);
    const f = s.getRange('A:A').createTextFinder(no).matchEntireCell(true).findNext();
    if (f && f.getRow() > 1) s.deleteRow(f.getRow());
  } catch (e) { /* ไม่กระทบข้อมูลหลัก */ }
}

function mirror_(d) {
  try {
    const s = viewSheet_(), no = String(d.no || d.id);
    const l = d.loanData || {}, t = d.returnData || {};
    let due = '';
    if (l.financeDate) {
      const x = new Date(l.financeDate + 'T00:00:00'); x.setDate(x.getDate() + 30);
      due = Utilities.formatDate(x, TZ, 'yyyy-MM-dd');
    }
    const c = d.cancelled && d.cancelled.active ? d.cancelled : null;
    const row = [no, d.status || '', d.category || '', d.name || '', d.email || '', d.dept || '',
      (d.companions || []).join(', '), d.type || '', d.title || '', d.place || '', d.dateFrom || '', d.dateTo || '',
      d.days || '', d.budget || 0, d.source || '', d.loan || '', l.total || '', due, t.returnDate || '',
      t.actual == null ? '' : t.actual, c ? (c.reason || '') + (c.note ? ' – ' + c.note : '') : ''];
    const f = s.getRange('A:A').createTextFinder(no).matchEntireCell(true).findNext();
    const r = f ? f.getRow() : s.getLastRow() + 1;
    s.getRange(r, 1, 1, row.length).setValues([row]);
  } catch (e) { /* สำเนาอ่านง่ายล้มเหลวได้ ไม่กระทบข้อมูลหลัก */ }
}
