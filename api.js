// 구글시트에 붙어 있는 Apps Script 웹 앱과 통신합니다.
// 웹 앱 주소는 코드에 넣지 않고, 연결 페이지(#setup=...)로 받아 이 기기 브라우저에만 저장합니다.

const store = {
  get(key) { try { return JSON.parse(localStorage.getItem(key)); } catch { return null; } },
  set(key, val) { try { localStorage.setItem(key, JSON.stringify(val)); } catch {} },
  del(key) { try { localStorage.removeItem(key); } catch {} },
};

const Api = {
  url() { return store.get('gb.api') || ''; },
  isReady() { return !!this.url(); },
  disconnect() { store.del('gb.api'); },

  // 연결 페이지에서 넘어온 주소를 저장하고, 주소창에서는 지웁니다.
  takeSetupFromHash() {
    const m = location.hash.match(/^#setup=(.+)$/);
    if (!m) return false;
    history.replaceState(null, '', location.pathname + location.search);
    const url = decodeURIComponent(m[1]);
    if (!/^https:\/\/script\.google\.com\/macros\/s\/[\w-]+\/exec$/.test(url)) return false;
    store.set('gb.api', url);
    return true;
  },

  async call(action, payload = {}) {
    if (!this.isReady()) throw new Error('아직 연결되지 않았어요.');
    let res;
    try {
      // text/plain으로 보내야 브라우저가 사전 확인(CORS preflight) 없이 바로 보냅니다.
      res = await fetch(this.url(), {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({ action, ...payload }),
      });
    } catch {
      throw new Error('인터넷 연결을 확인해 주세요.');
    }
    const body = await res.json().catch(() => null);
    if (!body) throw new Error('구글시트 응답을 읽지 못했어요. (' + res.status + ')');
    if (!body.ok) throw new Error(body.error);
    return body.data;
  },

  list() { return this.call('list'); },
  add(entry) { return this.call('add', { entry }); },
  update(entry) { return this.call('update', { entry }); },
  remove(id) { return this.call('remove', { id }); },
};
