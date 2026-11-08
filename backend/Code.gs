/**
 * 대신 말해줘 — 저장 서버 (Google Apps Script, 구글 시트에 붙여서 사용)
 *
 * 시트에 탭 3개가 자동으로 생긴다.
 *   문장     : 자주 쓰는 말 (순서, 문장)
 *   대본     : 현재 대본 (id, 제목, 내용, 수정시각, 삭제됨)
 *   버전기록 : 대본을 저장할 때마다 한 줄씩 쌓이는 기록 (저장시각, 대본id, 제목, 내용, 글자수)
 *
 * 비밀키(SECRET)를 아는 요청만 읽고 쓸 수 있다. 공개 저장소에는 비밀키를 넣지 않는다.
 */
const SECRET = '__SECRET__';
const TZ = 'Asia/Seoul';
const MAX_VERSIONS_RETURNED = 50;

function doGet(e) {
  return handle_((e && e.parameter) || {});
}

function doPost(e) {
  let p = {};
  try {
    p = JSON.parse(e.postData.contents);
  } catch (err) {
    return out_({ ok: false, error: 'bad_json' });
  }
  return handle_(p);
}

function handle_(p) {
  if (!p.k || p.k !== SECRET) return out_({ ok: false, error: 'unauthorized' });
  if (p.a === 'ping') return out_({ ok: true, pong: true });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    switch (p.a) {
      case 'load':
        return out_(Object.assign({ ok: true }, load_()));
      case 'init':
        return out_(Object.assign({ ok: true }, init_(p.phrases || [], p.scripts || [])));
      case 'savePhrases':
        savePhrases_(p.phrases || []);
        return out_({ ok: true });
      case 'saveScript':
        return out_(Object.assign({ ok: true }, saveScript_(p.script, !!p.draft)));
      case 'deleteScript':
        deleteScript_(String(p.id || ''));
        return out_({ ok: true });
      case 'versions':
        return out_({ ok: true, versions: versions_(String(p.id || '')) });
      default:
        return out_({ ok: false, error: 'unknown_action' });
    }
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// ---------- 시트 ----------
function sheet_(name, header, textCols) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
    sh.setFrozenRows(1);
    // '='로 시작하는 문장이 수식으로 바뀌지 않게 글자 형식으로 고정
    textCols.forEach(function (c) { sh.getRange(1, c, sh.getMaxRows(), 1).setNumberFormat('@'); });
  }
  return sh;
}
function shPhrases_() { return sheet_('문장', ['순서', '문장'], [2]); }
function shScripts_() { return sheet_('대본', ['id', '제목', '내용', '수정시각', '삭제됨'], [1, 2, 3, 4]); }
function shVersions_() { return sheet_('버전기록', ['저장시각', '대본id', '제목', '내용', '글자수'], [1, 2, 3, 4]); }

function rows_(sh, ncol) {
  const n = sh.getLastRow() - 1;
  if (n <= 0) return [];
  return sh.getRange(2, 1, n, ncol).getValues();
}
function now_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }

// ---------- 기능 ----------
function load_() {
  const phrases = rows_(shPhrases_(), 2)
    .filter(function (r) { return String(r[1]).trim() !== ''; })
    .sort(function (a, b) { return Number(a[0]) - Number(b[0]); })
    .map(function (r) { return String(r[1]); });
  const allScripts = rows_(shScripts_(), 5);
  const scripts = allScripts
    .filter(function (r) { return String(r[0]) !== '' && String(r[4]).toUpperCase() !== 'TRUE'; })
    .map(function (r) { return { id: String(r[0]), title: String(r[1]), text: String(r[2]), updated: String(r[3]) }; });
  return { initialized: phrases.length > 0 || allScripts.length > 0, phrases: phrases, scripts: scripts };
}

function init_(phrases, scripts) {
  const cur = load_();
  if (cur.initialized) return cur; // 이미 다른 기기가 시작했으면 덮어쓰지 않는다
  savePhrases_(phrases);
  scripts.forEach(function (s) { saveScript_(s, false); });
  return load_();
}

function savePhrases_(phrases) {
  const sh = shPhrases_();
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, 2).clearContent();
  const vals = phrases
    .map(function (t) { return String(t).trim(); })
    .filter(function (t) { return t !== ''; })
    .map(function (t, i) { return [i + 1, t]; });
  if (vals.length) sh.getRange(2, 1, vals.length, 2).setValues(vals);
}

function saveScript_(s, draft) {
  if (!s || !s.id) throw new Error('no_script_id');
  const id = String(s.id), title = String(s.title || ''), text = String(s.text || '');
  const sh = shScripts_();
  const rows = rows_(sh, 5);
  const t = now_();
  let rowIdx = -1;
  for (let i = 0; i < rows.length; i++) if (String(rows[i][0]) === id) { rowIdx = i + 2; break; }
  if (rowIdx > 0) sh.getRange(rowIdx, 1, 1, 5).setValues([[id, title, text, t, 'FALSE']]);
  else sh.appendRow([id, title, text, t, 'FALSE']);

  let versioned = false;
  if (!draft) {
    // 마지막 버전과 내용이 같으면 새 버전을 만들지 않는다
    const vs = shVersions_();
    const vrows = rows_(vs, 4);
    let lastV = null;
    for (let j = vrows.length - 1; j >= 0; j--) if (String(vrows[j][1]) === id) { lastV = vrows[j]; break; }
    if (!lastV || String(lastV[2]) !== title || String(lastV[3]) !== text) {
      vs.appendRow([t, id, title, text, text.length]);
      versioned = true;
    }
  }
  return { updated: t, versioned: versioned };
}

function deleteScript_(id) {
  const sh = shScripts_();
  const rows = rows_(sh, 5);
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === id) { sh.getRange(i + 2, 5).setValue('TRUE'); return; }
  }
}

function versions_(id) {
  const out = [];
  const vrows = rows_(shVersions_(), 5);
  for (let j = vrows.length - 1; j >= 0 && out.length < MAX_VERSIONS_RETURNED; j--) {
    if (String(vrows[j][1]) === id) out.push({ time: String(vrows[j][0]), title: String(vrows[j][2]), text: String(vrows[j][3]) });
  }
  return out;
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
