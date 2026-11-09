// Google 로그인(OAuth)과 Google Sheets API 호출을 담당합니다.
// Client ID와 시트 ID는 코드에 넣지 않고, 사용자의 브라우저(localStorage)에만 저장합니다.

const SCOPE = 'https://www.googleapis.com/auth/spreadsheets';
const SHEET = '내역';
const HEADER = ['ID', '날짜', '구분', '금액', '카테고리', '결제수단', '사용처', '메모', '등록시각'];
const API = 'https://sheets.googleapis.com/v4/spreadsheets';

const store = {
  get(key) { try { return JSON.parse(localStorage.getItem(key)); } catch { return null; } },
  set(key, val) { try { localStorage.setItem(key, JSON.stringify(val)); } catch {} },
  del(key) { try { localStorage.removeItem(key); } catch {} },
};

const Sheets = {
  settings() { return store.get('gb.settings') || { clientId: '', sheetId: '' }; },
  saveSettings(s) { store.set('gb.settings', s); },

  token() {
    const t = store.get('gb.token');
    return t && t.exp > Date.now() + 60_000 ? t.value : null;
  },
  isReady() { const s = this.settings(); return !!(s.clientId && s.sheetId && this.token()); },

  // 반드시 버튼 클릭 같은 사용자 동작 안에서 호출해야 팝업이 차단되지 않습니다.
  signIn() {
    return new Promise((resolve, reject) => {
      const { clientId } = this.settings();
      if (!clientId) return reject(new Error('설정에서 Client ID를 먼저 입력해 주세요.'));
      if (!window.google?.accounts?.oauth2) return reject(new Error('구글 로그인 모듈을 불러오지 못했습니다. 새로고침해 주세요.'));
      const client = google.accounts.oauth2.initTokenClient({
        client_id: clientId,
        scope: SCOPE,
        callback: (res) => {
          if (res.error) return reject(new Error(res.error_description || res.error));
          store.set('gb.token', { value: res.access_token, exp: Date.now() + res.expires_in * 1000 });
          resolve();
        },
        error_callback: (err) => reject(new Error(err.message || '로그인이 취소되었습니다.')),
      });
      client.requestAccessToken({ prompt: '' });
    });
  },
  signOut() {
    const t = store.get('gb.token');
    if (t && window.google?.accounts?.oauth2) google.accounts.oauth2.revoke(t.value, () => {});
    store.del('gb.token');
  },

  async call(path, opts = {}) {
    const token = this.token();
    if (!token) throw Object.assign(new Error('구글 연결이 필요합니다.'), { auth: true });
    const res = await fetch(API + path, {
      ...opts,
      headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
    });
    if (res.status === 401) { store.del('gb.token'); throw Object.assign(new Error('구글 연결이 만료되었습니다.'), { auth: true }); }
    const body = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(body.error?.message || '요청 실패 (' + res.status + ')');
    return body;
  },
  sid() { return encodeURIComponent(this.settings().sheetId); },
  range(r) { return encodeURIComponent(`'${SHEET}'!${r}`); },

  // 새 스프레드시트를 만들고 ID를 저장합니다.
  async create() {
    const ss = await this.call('', {
      method: 'POST',
      body: JSON.stringify({ properties: { title: '가계부' }, sheets: [{ properties: { title: SHEET } }] }),
    });
    this.saveSettings({ ...this.settings(), sheetId: ss.spreadsheetId });
    await this.writeHeader();
    return ss.spreadsheetUrl;
  },

  // 기존 시트에 '내역' 탭이 없으면 만들고, 머리글을 씁니다.
  async ensureSheet() {
    const ss = await this.call(`/${this.sid()}?fields=sheets.properties`);
    const tab = ss.sheets.find((s) => s.properties.title === SHEET);
    if (tab) { this._gid = tab.properties.sheetId; return; }
    const r = await this.call(`/${this.sid()}:batchUpdate`, {
      method: 'POST',
      body: JSON.stringify({ requests: [{ addSheet: { properties: { title: SHEET } } }] }),
    });
    this._gid = r.replies[0].addSheet.properties.sheetId;
    await this.writeHeader();
  },
  async writeHeader() {
    await this.call(`/${this.sid()}/values/${this.range('A1:I1')}?valueInputOption=RAW`, {
      method: 'PUT', body: JSON.stringify({ values: [HEADER] }),
    });
  },

  async list() {
    const r = await this.call(`/${this.sid()}/values/${this.range('A2:I')}?valueRenderOption=UNFORMATTED_VALUE`);
    return (r.values || [])
      .filter((v) => v[0])
      .map((v) => ({
        id: String(v[0]), date: String(v[1] || ''), type: v[2] === '수입' ? '수입' : '지출',
        amount: Number(v[3]) || 0, category: v[4] || '', method: v[5] || '',
        place: v[6] || '', memo: v[7] || '', createdAt: v[8] || '',
      }));
  },

  toRow(e) { return [e.id, e.date, e.type, e.amount, e.category, e.method, e.place, e.memo, e.createdAt]; },

  async add(e) {
    await this.call(`/${this.sid()}/values/${this.range('A:I')}:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`, {
      method: 'POST', body: JSON.stringify({ values: [this.toRow(e)] }),
    });
  },

  // 시트를 직접 편집했을 수도 있으니, 수정/삭제 직전에 ID로 행 번호를 다시 찾습니다.
  async rowOf(id) {
    const r = await this.call(`/${this.sid()}/values/${this.range('A:A')}`);
    const i = (r.values || []).findIndex((v) => String(v[0]) === id);
    if (i < 1) throw new Error('시트에서 해당 내역을 찾지 못했습니다. 새로고침해 주세요.');
    return i + 1; // 1부터 시작하는 행 번호
  },
  async update(e) {
    const row = await this.rowOf(e.id);
    await this.call(`/${this.sid()}/values/${this.range(`A${row}:I${row}`)}?valueInputOption=RAW`, {
      method: 'PUT', body: JSON.stringify({ values: [this.toRow(e)] }),
    });
  },
  async remove(id) {
    if (this._gid === undefined) await this.ensureSheet();
    const row = await this.rowOf(id);
    await this.call(`/${this.sid()}:batchUpdate`, {
      method: 'POST',
      body: JSON.stringify({ requests: [{ deleteDimension: { range: { sheetId: this._gid, dimension: 'ROWS', startIndex: row - 1, endIndex: row } } }] }),
    });
  },
};

// 시트 URL을 붙여넣어도 ID만 뽑아냅니다.
function parseSheetId(input) {
  const m = String(input).match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  return m ? m[1] : String(input).trim();
}
