// 가계부 구글시트에 붙어 있는 Apps Script입니다.
// 웹 앱으로 배포하면 GitHub Pages의 가계부 화면이 이 주소로 읽기/쓰기를 요청합니다.
// 웹 앱 주소 자체가 열쇠 역할을 하므로, 주소를 다른 사람에게 공유하지 마세요.

const APP_URL = 'https://remindforce-ops.github.io/gagyebu/';
const SHEET = '내역';
const HEADER = ['ID', '날짜', '구분', '금액', '카테고리', '결제수단', '사용처', '메모', '등록시각'];

// 브라우저로 직접 열면 연결용 페이지(버튼 + QR 코드)를 보여 줍니다.
function doGet() {
  sheet_(); // 처음 실행 시 '내역' 탭과 머리글 준비
  const t = HtmlService.createTemplateFromFile('setup');
  t.setupUrl = APP_URL + '#setup=' + encodeURIComponent(ScriptApp.getService().getUrl());
  return t.evaluate()
    .setTitle('가계부 연결')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 가계부 화면에서 오는 요청: { action, entry?, id? }
function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    out = { ok: true, data: handle_(req) };
  } catch (err) {
    out = { ok: false, error: String(err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function handle_(req) {
  if (req.action === 'list') return { entries: list_(), sheetUrl: SpreadsheetApp.getActiveSpreadsheet().getUrl() };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (req.action === 'add') { sheet_().appendRow(toRow_(req.entry)); return null; }
    if (req.action === 'update') { const r = rowOf_(req.entry.id); sheet_().getRange(r, 1, 1, HEADER.length).setValues([toRow_(req.entry)]); return null; }
    if (req.action === 'remove') { sheet_().deleteRow(rowOf_(req.id)); return null; }
  } finally {
    lock.releaseLock();
  }
  throw new Error('알 수 없는 요청입니다.');
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET, 0);
    sh.getRange(1, 1, 1, HEADER.length).setValues([HEADER]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('B:B').setNumberFormat('@'); // 날짜를 글자 그대로 저장
  }
  return sh;
}

function list_() {
  const sh = sheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const tz = Session.getScriptTimeZone();
  return sh.getRange(2, 1, n, HEADER.length).getValues()
    .filter((v) => v[0] !== '')
    .map((v) => ({
      id: String(v[0]),
      date: v[1] instanceof Date ? Utilities.formatDate(v[1], tz, 'yyyy-MM-dd') : String(v[1]),
      type: v[2] === '수입' ? '수입' : '지출',
      amount: Number(v[3]) || 0,
      category: String(v[4]), method: String(v[5]), place: String(v[6]), memo: String(v[7]),
      createdAt: v[8] instanceof Date ? v[8].toISOString() : String(v[8]),
    }));
}

// =, +, -, @ 로 시작하는 글자는 수식으로 해석되지 않도록 앞에 ' 를 붙입니다.
function text_(s) {
  s = String(s == null ? '' : s).slice(0, 200);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
function toRow_(e) {
  if (!e || !e.id || !/^\d{4}-\d{2}-\d{2}$/.test(e.date) || !(Number(e.amount) > 0)) throw new Error('입력값이 올바르지 않습니다.');
  return [text_(e.id), e.date, e.type === '수입' ? '수입' : '지출', Number(e.amount),
    text_(e.category), text_(e.method), text_(e.place), text_(e.memo), text_(e.createdAt)];
}

// 시트를 직접 고쳤을 수도 있으니 매번 ID로 행을 찾습니다.
function rowOf_(id) {
  const sh = sheet_();
  const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  const i = ids.findIndex((v) => String(v[0]) === String(id));
  if (i < 1) throw new Error('시트에서 해당 내역을 찾지 못했습니다. 새로고침해 주세요.');
  return i + 1;
}
