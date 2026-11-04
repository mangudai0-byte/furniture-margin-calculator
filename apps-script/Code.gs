/**
 * 가구공사 마진계산기 — 공용 저장소 (구글 스프레드시트 + Apps Script 웹앱)
 *
 * 1) 아래 TEAM_CODE를 직원들과 공유할 비밀번호로 바꾸세요. (예: 'hanbit-2026')
 * 2) 배포 → 새 배포 → 유형: 웹 앱, 실행 사용자: 나, 액세스 권한: 모든 사용자
 */
const TEAM_CODE = '여기에-팀-코드를-적으세요';

const SHEET_PROJECTS = '현장';
const SHEET_SETTINGS = '설정';

// 시트 열 순서 (key, 머리글)
const COLS = [
  ['id', 'ID'], ['date', '날짜'], ['company', '업체'], ['site', '현장명'],
  ['contract', '총공사금액'],
  ['q15', '15T 소요량'], ['q18', '18T 소요량'], ['qura', '우라 소요량'],
  ['p15', '15T 단가'], ['p18', '18T 단가'], ['pura', '우라 단가'],
  ['pb', 'PB 금액'], ['rate', '기본기물 비율(%)'], ['hw', '기본기물'],
  ['md', 'MD'], ['opt', '선택기물'], ['marble', '대리석'], ['labor', '시공비'],
  ['cost', '비용 총합계'], ['margin', '순마진'], ['marginRate', '이익률(%)'],
  ['author', '작성자'], ['updatedAt', '저장일시']
];
const NUM_KEYS = ['contract','q15','q18','qura','p15','p18','pura','pb','rate','hw','md','opt','marble','labor','cost','margin','marginRate'];
const DEFAULT_PRICES = { p15: 70000, p18: 80000, pura: 39000, rate: 5 };

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.code !== TEAM_CODE) return out({ ok: false, error: 'bad_code' });
  return out({ ok: true, projects: readProjects(), prices: readPrices() });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return out({ ok: false, error: 'bad_request' }); }
  if (!body || body.code !== TEAM_CODE) return out({ ok: false, error: 'bad_code' });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (body.action === 'save') {
      const list = Array.isArray(body.projects) ? body.projects : [body.project];
      list.forEach(saveProject);
    } else if (body.action === 'delete') {
      deleteProject(String(body.id || ''));
    } else if (body.action === 'prices') {
      writePrices(body.prices || {});
    } else {
      return out({ ok: false, error: 'bad_action' });
    }
    return out({ ok: true, projects: readProjects(), prices: readPrices() });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 현장 ---------- */
function projectSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_PROJECTS);
  if (!sh) sh = ss.insertSheet(SHEET_PROJECTS);
  if (sh.getLastRow() === 0) {
    sh.appendRow(COLS.map(c => c[1]));
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, COLS.length).setFontWeight('bold');
  }
  return sh;
}

function readProjects() {
  const sh = projectSheet();
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const rows = sh.getRange(2, 1, n, COLS.length).getValues();
  return rows.filter(r => r[0]).map(r => {
    const o = {};
    COLS.forEach((c, i) => {
      let v = r[i];
      if (v instanceof Date) v = Utilities.formatDate(v, Session.getScriptTimeZone(), c[0] === 'date' ? 'yyyy-MM-dd' : "yyyy-MM-dd'T'HH:mm:ss");
      o[c[0]] = NUM_KEYS.indexOf(c[0]) >= 0 ? Number(v) || 0 : String(v);
    });
    return o;
  });
}

function saveProject(p) {
  if (!p || !p.id) return;
  const sh = projectSheet();
  p.marginRate = Number(p.contract) > 0 ? Math.round(Number(p.margin) / Number(p.contract) * 1000) / 10 : 0;
  const row = COLS.map(c => {
    const v = p[c[0]];
    if (NUM_KEYS.indexOf(c[0]) >= 0) return Number(v) || 0;
    // 글자는 앞에 ' 를 붙여 그대로 저장 (날짜 자동변환·수식 실행 방지)
    return v == null || v === '' ? '' : "'" + String(v);
  });
  const idx = findRow(sh, p.id);
  if (idx > 0) {
    // 같은 현장이 이미 있고 시트 쪽이 더 최근이면 덮어쓰지 않음
    const cur = String(sh.getRange(idx, COLS.length).getValue() || '');
    if (cur && p.updatedAt && cur > p.updatedAt && p.onlyIfNewer) return;
    sh.getRange(idx, 1, 1, COLS.length).setValues([row]);
  } else {
    sh.appendRow(row);
  }
}

function deleteProject(id) {
  if (!id) return;
  const sh = projectSheet();
  const idx = findRow(sh, id);
  if (idx > 0) sh.deleteRow(idx);
}

function findRow(sh, id) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return -1;
  const ids = sh.getRange(2, 1, n, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}

/* ---------- 단가 ---------- */
function settingsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_SETTINGS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_SETTINGS);
    sh.getRange(1, 1, 5, 2).setValues([
      ['항목', '값'],
      ['15T 장당 단가', DEFAULT_PRICES.p15],
      ['18T 장당 단가', DEFAULT_PRICES.p18],
      ['우라 장당 단가', DEFAULT_PRICES.pura],
      ['기본기물 비율(%)', DEFAULT_PRICES.rate]
    ]);
    sh.getRange(1, 1, 1, 2).setFontWeight('bold');
  }
  return sh;
}

function readPrices() {
  const v = settingsSheet().getRange(2, 2, 4, 1).getValues();
  return {
    p15: Number(v[0][0]) || DEFAULT_PRICES.p15,
    p18: Number(v[1][0]) || DEFAULT_PRICES.p18,
    pura: Number(v[2][0]) || DEFAULT_PRICES.pura,
    rate: v[3][0] === '' ? DEFAULT_PRICES.rate : Number(v[3][0])
  };
}

function writePrices(p) {
  const cur = readPrices();
  const next = {
    p15: Number(p.p15) || cur.p15, p18: Number(p.p18) || cur.p18,
    pura: Number(p.pura) || cur.pura, rate: p.rate === undefined ? cur.rate : Number(p.rate) || 0
  };
  settingsSheet().getRange(2, 2, 4, 1).setValues([[next.p15], [next.p18], [next.pura], [next.rate]]);
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
