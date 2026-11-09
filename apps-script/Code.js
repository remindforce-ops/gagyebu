// 가계부 구글시트에 붙어 있는 Apps Script입니다.
// 화면과 데이터 처리를 모두 이 웹 앱이 담당합니다.
// 웹 앱 접근 권한이 '나만(MYSELF)'이라서, 내 구글 계정으로 로그인한 경우에만 열립니다.

const SHEET = '내역';
const HEADER = ['ID', '날짜', '구분', '금액', '카테고리', '결제수단', '가맹점명', '메모', '등록시각'];
const ITEM_SHEET = '품목';
const ITEM_HEADER = ['내역ID', '순번', '품명', '수량', '단가', '금액'];

function doGet() {
  sheet_(); // 처음 실행 시 '내역' 탭과 머리글 준비
  const t = HtmlService.createTemplateFromFile('index');
  t.appUrl = ScriptApp.getService().getUrl();
  t.sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  return t.evaluate()
    .setTitle('가계부')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

// ---------- 화면에서 google.script.run으로 호출하는 함수 ----------
function apiList() { return list_(); }
function apiAdd(entry) {
  return withLock_(() => {
    const items = toItemRows_(entry);
    sheet_().appendRow(toRow_(entry));
    appendItems_(items);
  });
}
function apiUpdate(entry) {
  return withLock_(() => {
    const items = toItemRows_(entry);
    sheet_().getRange(rowOf_(entry.id), 1, 1, HEADER.length).setValues([toRow_(entry)]);
    removeItems_(entry.id);
    appendItems_(items);
  });
}
function apiRemove(id) {
  return withLock_(() => {
    sheet_().deleteRow(rowOf_(id));
    removeItems_(id);
  });
}

// 영수증 사진(JPEG base64)을 구글 드라이브 OCR로 읽어 글자만 돌려줍니다.
// 사진은 OCR용 임시 문서로만 쓰고 바로 영구 삭제합니다(드라이브에 남지 않음).
function apiOcr(base64) {
  if (typeof base64 !== 'string' || base64.length > 15 * 1024 * 1024) throw new Error('사진이 너무 커요.');
  const blob = Utilities.newBlob(Utilities.base64Decode(base64), 'image/jpeg', 'receipt.jpg');
  const doc = Drive.Files.create({ name: '가계부-영수증-임시', mimeType: MimeType.GOOGLE_DOCS }, blob, { ocrLanguage: 'ko', fields: 'id' });
  const auth = { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() };
  try {
    const res = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + doc.id + '/export?mimeType=text/plain',
      { headers: auth, muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) throw new Error('영수증 글자를 읽지 못했어요. 다시 찍어 주세요.');
    return res.getContentText('UTF-8');
  } finally {
    UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + doc.id, { method: 'delete', headers: auth, muteHttpExceptions: true });
  }
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { fn(); return null; } finally { lock.releaseLock(); }
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET, 0);
    sh.getRange(1, 1, 1, HEADER.length).setValues([HEADER]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('B:B').setNumberFormat('@'); // 날짜를 글자 그대로 저장
  } else if (sh.getRange('G1').getValue() === '사용처') {
    sh.getRange('G1').setValue('가맹점명'); // 이전 버전 머리글 이름 변경
  }
  return sh;
}

function itemSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ITEM_SHEET);
  if (!sh) {
    sh = ss.insertSheet(ITEM_SHEET);
    sh.getRange(1, 1, 1, ITEM_HEADER.length).setValues([ITEM_HEADER]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

// 품목은 '품목' 탭에 내역ID로 연결해 저장합니다.
function toItemRows_(e) {
  const items = Array.isArray(e && e.items) ? e.items : [];
  if (items.length > 100) throw new Error('품목은 100개까지 저장할 수 있어요.');
  return items.map((it, i) => {
    if (!it || !String(it.name || '').trim()) throw new Error('품명이 비어 있는 품목이 있어요.');
    return [text_(e.id), i + 1, text_(String(it.name).trim().slice(0, 60)),
      Number(it.qty) || '', Number(it.price) || '', Number(it.amount) || 0];
  });
}
function appendItems_(rows) {
  if (!rows.length) return;
  const sh = itemSheet_();
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, ITEM_HEADER.length).setValues(rows);
}
function removeItems_(id) {
  const sh = itemSheet_();
  const last = sh.getLastRow();
  if (last < 2) return;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]) === String(id)) sh.deleteRow(i + 2);
  }
}
function itemsById_() {
  const sh = itemSheet_();
  const last = sh.getLastRow();
  const map = {};
  if (last < 2) return map;
  sh.getRange(2, 1, last - 1, ITEM_HEADER.length).getValues().forEach((v) => {
    if (v[0] === '') return;
    (map[String(v[0])] = map[String(v[0])] || []).push({
      seq: Number(v[1]) || 0, name: String(v[2]), qty: Number(v[3]) || 0, price: Number(v[4]) || 0, amount: Number(v[5]) || 0,
    });
  });
  Object.values(map).forEach((list) => list.sort((a, b) => a.seq - b.seq));
  return map;
}

function list_() {
  const sh = sheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const tz = Session.getScriptTimeZone();
  const items = itemsById_();
  return sh.getRange(2, 1, n, HEADER.length).getValues()
    .filter((v) => v[0] !== '')
    .map((v) => ({
      id: String(v[0]),
      date: v[1] instanceof Date ? Utilities.formatDate(v[1], tz, 'yyyy-MM-dd') : String(v[1]),
      type: v[2] === '수입' ? '수입' : '지출',
      amount: Number(v[3]) || 0,
      category: String(v[4]), method: String(v[5]), place: String(v[6]), memo: String(v[7]),
      createdAt: v[8] instanceof Date ? v[8].toISOString() : String(v[8]),
      items: items[String(v[0])] || [],
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
