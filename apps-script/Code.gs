/**
 * 고객 미팅 준비판 · 구글 시트 연결 (Apps Script 웹 앱)
 *
 * 읽기  GET  …/exec                 모든 탭을 JSON으로 돌려줌 (주소를 아는 누구나)
 * 쓰기  POST …/exec  {key, op, …}   스크립트 속성 WRITE_KEY 와 key 가 같을 때만
 *       op: ping | add {col,id,data} | update {col,id,data} | del {col,id} | batch {ops:[…]}
 *
 * 설치
 *  1. 이 시트에서 확장 프로그램 → Apps Script → 이 코드를 Code.gs 에 붙여 넣고 저장
 *  2. 프로젝트 설정(톱니바퀴) → 스크립트 속성 → 속성 추가: WRITE_KEY = 팀 편집 비밀번호
 *  3. 배포 → 새 배포 → 유형 「웹 앱」 → 실행 계정 「나」 → 액세스 권한 「모든 사용자」 → 배포
 *  4. 나오는 웹 앱 주소(…/exec)를 웹 화면의 config.js 의 API_URL 에 넣음
 *  코드를 고친 뒤에는 배포 → 배포 관리 → 연필 → 버전 「새 버전」 → 배포 (주소는 그대로)
 *
 * 탭마다 1행 머리글 이름(아래 SCHEMA 두 번째 값)은 바꾸지 마세요. 열 순서는 바꿔도 됩니다.
 */
const SCHEMA = {
  meetings: {tab: "미팅", cols: [["id", "ID"], ["title", "미팅 이름"], ["company", "고객 회사"], ["customerIds", "참석 고객ID", "json"], ["contact", "고객 담당자"], ["kind", "미팅 구분"], ["start", "시작일"], ["end", "끝 날짜"], ["time", "시간"], ["place", "장소"], ["lead", "준비 총괄"], ["status", "상태"], ["goals", "목표", "json"], ["created", "만든 시각"]]},
  customers: {tab: "고객", cols: [["id", "ID"], ["company", "회사"], ["name", "성함"], ["title", "직급"], ["dept", "부서"], ["phone", "연락처"], ["email", "이메일"], ["country", "국가"], ["memo", "메모"], ["created", "만든 시각"]]},
  requests: {tab: "팀 할 일", cols: [["id", "ID"], ["meetingId", "미팅ID"], ["no", "번호", "num"], ["group", "준비 항목"], ["title", "할 일"], ["owner", "담당"], ["asked", "요청일"], ["due", "기한"], ["done", "회신일"], ["status", "상태"], ["via", "경로"], ["note", "비고"], ["prevStatus", "이전 상태"], ["src", "출처"], ["created", "만든 시각"]]},
  records: {tab: "기록", cols: [["id", "ID"], ["meetingId", "미팅ID"], ["seq", "순번", "num"], ["kind", "종류"], ["title", "제목"], ["at", "시각"], ["durationSec", "길이(초)", "num"], ["lang", "언어"], ["text", "내용"], ["summary", "일반 요약"], ["notice", "공지"], ["report", "보고"], ["key5", "중요 요약", "json"], ["extract", "뽑은 항목", "json"], ["applied", "반영 여부", "bool"], ["aiAt", "AI 정리 시각"], ["created", "만든 시각"]]}
};
const PFX = {customers: 'c', meetings: 'm', requests: 'r', questions: 'q', schedule: 's', followups: 'f', records: 'rc'};
const TIME_KEYS = ['created', 'at', 'aiAt'];

function doGet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tz = ss.getSpreadsheetTimeZone();
    const data = {};
    Object.keys(SCHEMA).forEach(function (col) { data[col] = readTab_(ss, col, tz); });
    return out_({ok: true, data: data, at: new Date().toISOString()});
  } catch (err) {
    return out_({ok: false, code: 'server', message: String(err)});
  }
}

function doPost(e) {
  var req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return out_({ok: false, code: 'bad_json'}); }
  const key = PropertiesService.getScriptProperties().getProperty('WRITE_KEY');
  if (!key) return out_({ok: false, code: 'no_key_set'});
  if (String(req.key || '') !== key) { Utilities.sleep(700); return out_({ok: false, code: 'bad_key'}); }
  if (req.op === 'ping') return out_({ok: true});
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return out_({ok: false, code: 'busy'});
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cache = {};
    const ops = req.op === 'batch' ? (req.ops || []) : [req];
    const results = ops.map(function (op) {
      try { return apply_(ss, cache, op); } catch (err) { return {ok: false, code: 'op_failed', message: String(err)}; }
    });
    SpreadsheetApp.flush();
    return out_({ok: results.every(function (r) { return r.ok; }), results: results});
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 쓰기 ---------- */
function apply_(ss, cache, op) {
  const t = tab_(ss, cache, op.col);
  const id = String(op.id || '');
  if (op.op === 'add') {
    const newId = id && t.ids.indexOf(id) < 0 ? id : makeId_(op.col);
    const data = Object.assign({}, op.data || {}, {id: newId});
    Object.keys(data).forEach(function (k) { colFor_(t, k); });
    const row = t.ids.length + 2;
    const vals = [], fmts = [];
    for (var c = 1; c <= t.width; c++) { vals.push(''); fmts.push('@'); }
    Object.keys(data).forEach(function (k) {
      const c = t.byKey[k];
      vals[c - 1] = enc_(t, k, data[k]);
      fmts[c - 1] = type_(t, k) === 'num' ? '0' : '@';
    });
    const rg = t.sh.getRange(row, 1, 1, t.width);
    rg.setNumberFormats([fmts]);
    rg.setValues([vals]);
    t.ids.push(newId);
    return {ok: true, id: newId};
  }
  const idx = t.ids.indexOf(id);
  if (idx < 0) return {ok: true, missing: true};
  const row = idx + 2;
  if (op.op === 'update') {
    const data = op.data || {};
    Object.keys(data).forEach(function (k) {
      if (k === 'id' || data[k] === undefined) return;
      const c = colFor_(t, k);
      t.sh.getRange(row, c).setNumberFormat(type_(t, k) === 'num' ? '0' : '@').setValue(enc_(t, k, data[k]));
    });
    return {ok: true};
  }
  if (op.op === 'del') {
    t.sh.deleteRow(row);
    t.ids.splice(idx, 1);
    return {ok: true};
  }
  return {ok: false, code: 'bad_op'};
}

/* ---------- 읽기 ---------- */
function readTab_(ss, col, tz) {
  const t = tab_(ss, {}, col);
  const n = t.sh.getLastRow() - 1;
  if (n < 1) return [];
  const vals = t.sh.getRange(2, 1, n, t.width).getValues();
  const rows = [];
  vals.forEach(function (r, i) {
    if (r.every(function (v) { return v === '' || v === null; })) return;
    var id = String(r[t.byKey.id - 1] || '').trim();
    if (!id) {  // 시트에서 직접 넣은 줄은 ID를 채움
      id = makeId_(col);
      t.sh.getRange(i + 2, t.byKey.id).setNumberFormat('@').setValue(id);
    }
    const o = {id: id};
    Object.keys(t.byKey).forEach(function (k) {
      if (k === 'id') return;
      const v = dec_(t, k, r[t.byKey[k] - 1], tz);
      if (v !== undefined) o[k] = v;
    });
    rows.push(o);
  });
  return rows;
}

/* ---------- 도움 함수 ---------- */
function tab_(ss, cache, col) {
  if (cache[col]) return cache[col];
  const def = SCHEMA[col];
  if (!def) throw new Error('알 수 없는 목록: ' + col);
  var sh = ss.getSheetByName(def.tab);
  if (!sh) {
    sh = ss.insertSheet(def.tab);
    sh.getRange(1, 1, 1, def.cols.length).setValues([def.cols.map(function (c) { return c[1]; })]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  const width = Math.max(sh.getLastColumn(), 1);
  const head = sh.getRange(1, 1, 1, width).getDisplayValues()[0].map(function (h) { return String(h).trim(); });
  const byKey = {};
  def.cols.forEach(function (c) { var i = head.indexOf(c[1]); if (i < 0) i = head.indexOf(c[0]); if (i >= 0) byKey[c[0]] = i + 1; });
  head.forEach(function (h, i) { if (h && !def.cols.some(function (c) { return c[1] === h || c[0] === h; }) && !byKey[h]) byKey[h] = i + 1; });
  if (!byKey.id) throw new Error(def.tab + ' 탭 1행에 「ID」 머리글이 없습니다');
  const n = sh.getLastRow() - 1;
  const ids = n > 0 ? sh.getRange(2, byKey.id, n, 1).getDisplayValues().map(function (r) { return String(r[0]).trim(); }) : [];
  return cache[col] = {sh: sh, def: def, byKey: byKey, ids: ids, width: width};
}

function colFor_(t, k) {
  if (t.byKey[k]) return t.byKey[k];
  const c = t.def.cols.filter(function (x) { return x[0] === k; })[0];
  t.width += 1;
  t.sh.getRange(1, t.width).setValue(c ? c[1] : k).setFontWeight('bold');
  t.byKey[k] = t.width;
  return t.width;
}

function type_(t, k) {
  const c = t.def.cols.filter(function (x) { return x[0] === k; })[0];
  return c && c[2] ? c[2] : '';
}

function enc_(t, k, v) {
  if (v === undefined || v === null) return '';
  const ty = type_(t, k);
  if (ty === 'json' || typeof v === 'object') return JSON.stringify(v);
  if (ty === 'num') { const n = Number(v); return v === '' || isNaN(n) ? '' : n; }
  if (ty === 'bool') return v ? 'TRUE' : 'FALSE';
  return String(v);
}

function dec_(t, k, v, tz) {
  const ty = type_(t, k);
  if (v instanceof Date) {
    if (k === 'time') return Utilities.formatDate(v, tz, 'HH:mm');
    if (TIME_KEYS.indexOf(k) >= 0) return Utilities.formatDate(v, tz, "yyyy-MM-dd'T'HH:mm:ssXXX");
    return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  }
  if (ty === 'json') {
    if (v === '' || v === null) return undefined;
    try { return JSON.parse(v); } catch (e) { return undefined; }
  }
  if (ty === 'num') {
    if (v === '' || v === null) return undefined;
    const n = Number(v);
    return isNaN(n) ? undefined : n;
  }
  if (ty === 'bool') return v === true || String(v).toUpperCase() === 'TRUE';
  if (!ty && typeof v === 'string' && /^[\[{]/.test(v)) { try { return JSON.parse(v); } catch (e) { /* 글 그대로 */ } }
  return v === null ? '' : String(v);
}

function makeId_(col) {
  return (PFX[col] || 'x') + Date.now().toString(36) + Utilities.getUuid().replace(/-/g, '').slice(0, 4);
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
