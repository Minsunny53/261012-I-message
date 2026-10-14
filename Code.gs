/**
 * 나 전달법 놀이터 — Apps Script 서버 v2 (관리자 1명이 설치 → 선생님들은 앱에서 학급을 직접 만들어 써요)
 *
 * [구조]
 *  - 이 서버는 '관리자(나)'만 한 번 설치해요. 선생님은 설치할 것이 없어요.
 *  - 선생님이 앱의 [내 학급] 탭에서 [새 학급 만들기]를 누르면 8자리 학급 코드가 자동 생성되고, 선생님이 정한 비밀번호로 잠가요.
 *  - 학생은 학급 코드(8자리) + 번호만 입력해요. (이름·구글계정 저장 안 함)
 *  - 선생님은 '자기 학급' 기록만 볼 수 있고, 지울 수도 있어요. 비밀번호는 암호화(해시)해서만 저장해요.
 *  - Gemini 키는 관리자 것을 함께 쓰되, 학급별로 호출 횟수를 제한해요.
 *
 * [설치] 구글 스프레드시트 새로 만들기 → 확장 프로그램 → Apps Script → 이 코드 전체 붙여넣기
 *  → 프로젝트 설정(톱니) → 스크립트 속성에 아래 값 추가
 *      GEMINI_API_KEY : Google AI Studio에서 받은 키
 *      ADMIN_KEY      : 관리자(나)만 아는 긴 비밀번호 (전체 학급 관리용)
 *      SIGNUP_CODE    : (선택) 넣으면 '새 학급 만들기' 때 이 초대 코드를 입력해야 해요. 학교 선생님들께만 알려 주세요.
 *      MODEL          : (선택) 기본 gemini-2.5-flash
 *      MAX_CLASSES    : (선택) 최대 학급 수, 기본 300
 *      GEMINI_PER_CLASS : (선택) 학급당 6시간 AI 호출 한도, 기본 400
 *      RETENTION_DAYS : (선택) 제출 기록 자동 삭제 기간(일), 기본 180
 *  → 편집기에서 setupSheet 함수를 한 번 실행(권한 승인)
 *  → 배포 → 새 배포 → 유형 '웹 앱' → 실행 사용자: 나 / 액세스 권한: 모든 사용자 → 배포 → 웹 앱 URL(/exec) 복사
 *  → 앱 코드의 PROXY_URL_CONST에 그 URL을 넣고 GitHub에 배포하면, 선생님들은 따로 설정할 게 없어요.
 *  ※ 코드를 고친 뒤에는 '배포 관리 → 수정 → 새 버전'으로 다시 배포해야 반영돼요.
 */

const SHEET_NAME = '제출함';
const CLASS_SHEET = '학급';
const HEADERS = ['id', '제출시각', '학급코드', '번호', '캐릭터(말하는/대답하는)', '상황', '행동·상황', '감정', '바람', '인정', '사과', '약속', '자동점검 통과(6칸 중)'];
const CLASS_HEADERS = ['학급코드', '학급 이름', 'salt', 'hash', '만든 시각', '마지막 사용'];
const MAX_FIELD = 200;
const PII_RE = /\d[\d\s-]{6,}\d|@|https?:\/\//i;   // 전화번호·이메일·주소(URL)처럼 보이면 저장하지 않아요

function props_() { return PropertiesService.getScriptProperties(); }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function digits_(v) { return String(v == null ? '' : v).replace(/\D/g, ''); }
function now_() { return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss'); }
function clip_(v, n) { return String(v == null ? '' : v).replace(/\s+/g, ' ').trim().slice(0, n || MAX_FIELD); }

function doGet() { return out_({ ok: true, msg: '나 전달법 놀이터 서버가 켜져 있어요.', version: 2 }); }

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'BAD_REQUEST' }); }
  try {
    switch (req.action) {
      case 'createClass':    return out_(createClass_(req));
      case 'login':          return out_(login_(req));
      case 'checkClass':     return out_(checkClass_(req));
      case 'changePassword': return out_(changePassword_(req));
      case 'deleteClass':    return out_(deleteClass_(req));
      case 'gemini':         return out_(handleGemini_(req));
      case 'submit':         return out_(handleSubmit_(req));
      case 'read':           return out_(handleRead_(req));
      case 'purge':          return out_(handlePurge_(req));
      case 'adminList':      return out_(adminList_(req));
      case 'adminReset':     return out_(adminReset_(req));
      case 'adminDelete':    return out_(adminDelete_(req));
      default:               return out_({ ok: false, error: 'BAD_ACTION' });
    }
  } catch (err) {
    return out_({ ok: false, error: 'SERVER', detail: String(err).slice(0, 200) });
  }
}

/* ---------- 공통: 속도 제한 / 시트 ---------- */
function tooFast_(bucket, id, limit, seconds) {
  const cache = CacheService.getScriptCache();
  const k = 'rl:' + bucket + ':' + id;
  const n = Number(cache.get(k) || 0);
  if (n >= limit) return true;
  cache.put(k, String(n + 1), seconds || 60);
  return false;
}
function sheet_(name, headers) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) { sh.appendRow(headers); sh.setFrozenRows(1); }
  return sh;
}
function subSheet_() { return sheet_(SHEET_NAME, HEADERS); }
function classSheet_() { return sheet_(CLASS_SHEET, CLASS_HEADERS); }

/* ---------- 학급 / 비밀번호 ---------- */
function hash_(salt, pw) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + pw, Utilities.Charset.UTF_8);
  return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}
/** 학급 행 찾기 → { row, code, label, salt, hash } 또는 null */
function findClass_(code) {
  code = digits_(code);
  if (code.length !== 8) return null;
  const sh = classSheet_(), last = sh.getLastRow();
  if (last < 2) return null;
  const vals = sh.getRange(2, 1, last - 1, CLASS_HEADERS.length).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (digits_(vals[i][0]) === code) return { row: i + 2, code: code, label: String(vals[i][1] || ''), salt: String(vals[i][2]), hash: String(vals[i][3]) };
  }
  return null;
}
function classExists_(code) {
  code = digits_(code);
  const cache = CacheService.getScriptCache(), k = 'cx:' + code;
  if (cache.get(k) === '1') return true;
  const c = findClass_(code);
  if (c) cache.put(k, '1', 600);
  return !!c;
}
/** 코드+비밀번호 확인. 5번 틀리면 10분 잠금 */
function authClass_(code, pw) {
  code = digits_(code);
  const cache = CacheService.getScriptCache(), fk = 'fail:' + code;
  if (Number(cache.get(fk) || 0) >= 5) return { ok: false, error: 'LOCKED' };
  const c = findClass_(code);
  if (!c) return { ok: false, error: 'NO_CLASS' };
  if (!pw || hash_(c.salt, String(pw)) !== c.hash) {
    cache.put(fk, String(Number(cache.get(fk) || 0) + 1), 600);
    return { ok: false, error: 'BAD_PW' };
  }
  cache.remove(fk);
  return { ok: true, cls: c };
}
function touch_(row) { try { classSheet_().getRange(row, 6).setValue(now_()); } catch (e) {} }

function createClass_(req) {
  const signup = props_().getProperty('SIGNUP_CODE');
  if (signup && String(req.invite || '').trim() !== signup) return { ok: false, error: 'BAD_INVITE' };
  const pw = String(req.password || '');
  if (pw.length < 4 || pw.length > 40) return { ok: false, error: 'WEAK_PW' };
  if (tooFast_('mk', 'all', 30, 3600)) return { ok: false, error: 'RATE' };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = classSheet_();
    const max = Number(props_().getProperty('MAX_CLASSES') || 300);
    if (sh.getLastRow() - 1 >= max) return { ok: false, error: 'FULL' };
    let code = '';
    for (let i = 0; i < 30; i++) {
      const c = String(10000000 + Math.floor(Math.random() * 90000000));
      if (!findClass_(c)) { code = c; break; }
    }
    if (!code) return { ok: false, error: 'SERVER' };
    const salt = Utilities.getUuid();
    sh.appendRow([code, clip_(req.label, 30), salt, hash_(salt, pw), now_(), '']);
    return { ok: true, code: code };
  } finally { lock.releaseLock(); }
}
/** 학생 입장 확인: 학급 코드가 있는지만 알려줘요 (비밀번호 필요 없음, 너무 자주 시도하면 막아요) */
function checkClass_(req) {
  if (tooFast_('chk', 'all', 120, 60)) return { ok: false, error: 'RATE' };
  return classExists_(req.code) ? { ok: true } : { ok: false, error: 'NO_CLASS' };
}
function login_(req) {
  const a = authClass_(req.code, req.password);
  if (!a.ok) return a;
  touch_(a.cls.row);
  return { ok: true, code: a.cls.code, label: a.cls.label };
}
function changePassword_(req) {
  const a = authClass_(req.code, req.password);
  if (!a.ok) return a;
  const np = String(req.newPassword || '');
  if (np.length < 4 || np.length > 40) return { ok: false, error: 'WEAK_PW' };
  setPassword_(a.cls.row, np);
  return { ok: true };
}
function setPassword_(row, pw) {
  const salt = Utilities.getUuid();
  classSheet_().getRange(row, 3, 1, 2).setValues([[salt, hash_(salt, pw)]]);
}
function deleteClass_(req) {
  const a = authClass_(req.code, req.password);
  if (!a.ok) return a;
  removeRows_(a.cls.code);
  classSheet_().deleteRow(a.cls.row);
  CacheService.getScriptCache().remove('cx:' + a.cls.code);
  return { ok: true };
}

/* ---------- 1) Gemini 중계 ---------- */
function handleGemini_(req) {
  const key = props_().getProperty('GEMINI_API_KEY');
  if (!key) return { ok: false, error: 'NO_KEY' };
  const code = digits_(req.code);
  if (!classExists_(code)) return { ok: false, error: 'BAD_CODE' };
  if (tooFast_('g', code + ':' + String(req.number || 'x'), 8, 60)) return { ok: false, error: 'RATE' };
  const cap = Number(props_().getProperty('GEMINI_PER_CLASS') || 400);
  if (tooFast_('gc', code, cap, 21600)) return { ok: false, error: 'CLASS_QUOTA' };
  const model = /^gemini-[\w.\-]+$/.test(req.model || '') ? req.model : (props_().getProperty('MODEL') || 'gemini-2.5-flash');
  const body = req.body || {};
  try {   // 학생의 마지막 입력에 전화번호·이메일·주소 형태가 있으면 AI로 보내지 않아요
    const cs = body.contents || [], last = cs[cs.length - 1];
    const lastText = last && last.parts ? last.parts.map(function (p) { return p.text || ''; }).join(' ') : '';
    if (PII_RE.test(lastText)) return { ok: false, error: 'PII' };
  } catch (e) {}
  if (!body.safetySettings) {
    body.safetySettings = ['HARM_CATEGORY_HARASSMENT', 'HARM_CATEGORY_HATE_SPEECH', 'HARM_CATEGORY_SEXUALLY_EXPLICIT', 'HARM_CATEGORY_DANGEROUS_CONTENT']
      .map(function (c) { return { category: c, threshold: 'BLOCK_MEDIUM_AND_ABOVE' }; });
  }
  const payload = JSON.stringify(body);
  if (payload.length > 30000) return { ok: false, error: 'TOO_LONG' };
  const res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(model) + ':generateContent', {
    method: 'post', contentType: 'application/json', headers: { 'x-goog-api-key': key }, payload: payload, muteHttpExceptions: true
  });
  const status = res.getResponseCode();
  if (status === 429) return { ok: false, error: 'QUOTA' };
  if (status !== 200) return { ok: false, error: 'HTTP_' + status };
  const data = JSON.parse(res.getContentText());
  const parts = (data.candidates && data.candidates[0] && data.candidates[0].content && data.candidates[0].content.parts) || [];
  const text = parts.map(function (p) { return p.text || ''; }).join('').trim();
  if (!text) return { ok: false, error: (data.promptFeedback && data.promptFeedback.blockReason) ? 'BLOCKED' : 'EMPTY' };
  return { ok: true, text: text };
}

/* ---------- 2) 학생 제출 ---------- */
function handleSubmit_(req) {
  purgeOldOnce_();
  const code = digits_(req.code);
  if (!classExists_(code)) return { ok: false, error: 'BAD_CODE' };
  const no = String(req.number || '').trim();
  if (!/^\d{1,2}$/.test(no)) return { ok: false, error: 'BAD_NUMBER' };
  if (tooFast_('s', code + ':' + no, 20, 60)) return { ok: false, error: 'RATE' };
  const f = ['situation', 'a', 'e', 'w', 'ack', 'sorry', 'prom'].map(function (k) { return clip_(req[k]); });
  const chars = clip_(req.chars);
  if (PII_RE.test(f.join(' ') + ' ' + chars)) return { ok: false, error: 'PII' };
  const pass = Math.max(0, Math.min(6, Number(req.pass) || 0));
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    subSheet_().appendRow([Utilities.getUuid().slice(0, 8), now_(), code, Number(no), chars, f[0], f[1], f[2], f[3], f[4], f[5], f[6], pass]);
  } finally { lock.releaseLock(); }
  return { ok: true };
}

/* ---------- 3) 교사: 자기 학급 기록 열람 ---------- */
function handleRead_(req) {
  const a = authClass_(req.code, req.password);
  if (!a.ok) return a;
  touch_(a.cls.row);
  const sh = subSheet_(), last = sh.getLastRow();
  if (last < 2) return { ok: true, rows: [], total: 0 };
  const vals = sh.getRange(2, 1, last - 1, HEADERS.length).getDisplayValues();
  const mine = vals.filter(function (r) { return digits_(r[2]) === a.cls.code; });
  const rows = mine.slice(-500).map(function (r) { const o = {}; HEADERS.forEach(function (h, i) { o[h] = r[i]; }); return o; }).reverse();
  return { ok: true, rows: rows, total: mine.length };
}

/* ---------- 4) 교사: 자기 학급 기록 삭제 ---------- */
function removeRows_(code) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = subSheet_(), last = sh.getLastRow();
    if (last < 2) return 0;
    const vals = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
    const keep = vals.filter(function (r) { return digits_(r[2]) !== code; });
    const removed = vals.length - keep.length;
    if (!removed) return 0;
    sh.getRange(2, 1, vals.length, HEADERS.length).clearContent();
    if (keep.length) sh.getRange(2, 1, keep.length, HEADERS.length).setValues(keep);
    return removed;
  } finally { lock.releaseLock(); }
}
/** 오래된 제출 기록 자동 삭제 (하루에 한 번만 확인) — RETENTION_DAYS(기본 180일) */
function purgeOldOnce_() {
  const cache = CacheService.getScriptCache();
  if (cache.get('purgeday')) return;
  cache.put('purgeday', '1', 21600);
  try { purgeOld_(); } catch (e) {}
}
function purgeOld_() {
  const days = Number(props_().getProperty('RETENTION_DAYS') || 180);
  const cutoff = Utilities.formatDate(new Date(Date.now() - days * 86400000), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = subSheet_(), last = sh.getLastRow();
    if (last < 2) return 0;
    const vals = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
    const keep = vals.filter(function (r) { return String(r[1]) >= cutoff; });
    const removed = vals.length - keep.length;
    if (!removed) return 0;
    sh.getRange(2, 1, vals.length, HEADERS.length).clearContent();
    if (keep.length) sh.getRange(2, 1, keep.length, HEADERS.length).setValues(keep);
    return removed;
  } finally { lock.releaseLock(); }
}
function handlePurge_(req) {
  const a = authClass_(req.code, req.password);
  if (!a.ok) return a;
  return { ok: true, removed: removeRows_(a.cls.code) };
}

/* ---------- 5) 관리자(나) 전용 ---------- */
function adminGuard_(req) {
  const k = props_().getProperty('ADMIN_KEY');
  const cache = CacheService.getScriptCache();
  if (Number(cache.get('fail:admin') || 0) >= 5) return { ok: false, error: 'LOCKED' };
  if (!k || String(req.adminKey || '') !== k) {
    cache.put('fail:admin', String(Number(cache.get('fail:admin') || 0) + 1), 600);
    return { ok: false, error: 'BAD_ADMIN' };
  }
  cache.remove('fail:admin');
  return null;
}
function adminList_(req) {
  const g = adminGuard_(req); if (g) return g;
  const sh = classSheet_(), last = sh.getLastRow();
  const counts = {};
  const ss = subSheet_(), sl = ss.getLastRow();
  if (sl >= 2) ss.getRange(2, 3, sl - 1, 1).getDisplayValues().forEach(function (r) { const c = digits_(r[0]); counts[c] = (counts[c] || 0) + 1; });
  const rows = last < 2 ? [] : sh.getRange(2, 1, last - 1, CLASS_HEADERS.length).getDisplayValues().map(function (r) {
    const c = digits_(r[0]);
    return { code: c, label: r[1], created: r[4], lastUsed: r[5], submissions: counts[c] || 0 };
  });
  return { ok: true, classes: rows };
}
function adminReset_(req) {
  const g = adminGuard_(req); if (g) return g;
  const c = findClass_(req.code); if (!c) return { ok: false, error: 'NO_CLASS' };
  const np = String(req.newPassword || '');
  if (np.length < 4 || np.length > 40) return { ok: false, error: 'WEAK_PW' };
  setPassword_(c.row, np);
  CacheService.getScriptCache().remove('fail:' + c.code);
  return { ok: true };
}
function adminDelete_(req) {
  const g = adminGuard_(req); if (g) return g;
  const c = findClass_(req.code); if (!c) return { ok: false, error: 'NO_CLASS' };
  removeRows_(c.code);
  classSheet_().deleteRow(c.row);
  CacheService.getScriptCache().remove('cx:' + c.code);
  return { ok: true };
}

/* ---------- (선택) 편집기에서 직접 실행해 보는 점검용 ---------- */
function setupSheet() { subSheet_(); classSheet_(); }   // 처음에 한 번 실행하면 시트 2개와 권한 승인이 준비돼요
