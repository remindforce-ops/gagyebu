const CATEGORIES = {
  지출: ['식비', '카페·간식', '교통', '주거·통신', '생활용품', '쇼핑', '의료·건강', '문화·여가', '교육', '경조사·선물', '기타'],
  수입: ['급여', '부수입', '용돈', '금융수입', '기타'],
};
const METHODS = ['현금', '체크카드', '신용카드', '계좌이체'];

const $ = (id) => document.getElementById(id);
const won = (n) => n.toLocaleString('ko-KR') + '원';
const pad = (n) => String(n).padStart(2, '0');
const today = () => { const d = new Date(); return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; };
const DAYS = ['일', '월', '화', '수', '목', '금', '토'];

const state = { entries: [], month: today().slice(0, 7), editing: null, form: { category: '', method: '' } };

function toast(msg) {
  const t = $('toast');
  t.textContent = msg;
  t.classList.add('show');
  clearTimeout(toast._t);
  toast._t = setTimeout(() => t.classList.remove('show'), 2200);
}

// ---------- 화면 그리기 ----------
function render() {
  const [y, m] = state.month.split('-').map(Number);
  $('monthLabel').textContent = `${y}년 ${m}월`;

  const items = state.entries.filter((e) => e.date.startsWith(state.month));
  let income = 0, expense = 0;
  const byCat = {};
  for (const e of items) {
    if (e.type === '수입') income += e.amount;
    else { expense += e.amount; byCat[e.category || '기타'] = (byCat[e.category || '기타'] || 0) + e.amount; }
  }
  $('incomeSum').textContent = won(income);
  $('expenseSum').textContent = won(expense);
  $('balance').textContent = (income - expense < 0 ? '-' : '') + won(Math.abs(income - expense));

  // 카테고리별 지출 (상위 5개)
  const cats = Object.entries(byCat).sort((a, b) => b[1] - a[1]).slice(0, 5);
  $('categoryCard').hidden = cats.length === 0;
  $('categoryBars').replaceChildren(...cats.map(([name, v]) => {
    const li = document.createElement('li');
    li.innerHTML = '<span class="name"></span><span class="track"><span class="fill"></span></span><span class="val"></span>';
    li.querySelector('.name').textContent = name;
    li.querySelector('.fill').style.width = (v / cats[0][1]) * 100 + '%';
    li.querySelector('.val').textContent = `${won(v)} · ${Math.round((v / expense) * 100)}%`;
    return li;
  }));

  // 날짜별 목록 (최신순)
  const groups = {};
  for (const e of items) (groups[e.date] ||= []).push(e);
  const days = Object.keys(groups).sort().reverse();
  $('list').replaceChildren(...days.map((date) => {
    const list = groups[date].sort((a, b) => String(b.createdAt).localeCompare(String(a.createdAt)));
    const net = list.reduce((s, e) => s + (e.type === '수입' ? e.amount : -e.amount), 0);
    const sec = document.createElement('div');
    sec.className = 'day';
    const d = new Date(date + 'T00:00:00');
    sec.innerHTML = '<div class="day-head"><span></span><span></span></div><div class="day-items"></div>';
    sec.querySelector('.day-head span:first-child').textContent = `${d.getMonth() + 1}월 ${d.getDate()}일 (${DAYS[d.getDay()]})`;
    sec.querySelector('.day-head span:last-child').textContent = (net > 0 ? '+' : net < 0 ? '-' : '') + won(Math.abs(net));
    sec.querySelector('.day-items').append(...list.map(itemEl));
    return sec;
  }));
  $('emptyMsg').hidden = items.length > 0 || !Sheets.isReady();
}

function itemEl(e) {
  const b = document.createElement('button');
  b.className = 'item';
  b.type = 'button';
  b.innerHTML = '<span class="badge"></span><span class="main"><div class="t1"></div><div class="t2"></div></span><span class="amt"></span>';
  b.querySelector('.badge').textContent = (e.category || '기타').slice(0, 2);
  b.querySelector('.t1').textContent = e.place || e.memo || e.category || e.type;
  b.querySelector('.t2').textContent = [e.category, e.method, e.place && e.memo].filter(Boolean).join(' · ');
  const amt = b.querySelector('.amt');
  amt.textContent = (e.type === '수입' ? '+' : '-') + won(e.amount);
  amt.classList.add(e.type === '수입' ? 'income' : 'expense');
  b.addEventListener('click', () => openEntry(e));
  return b;
}

function showConnect(msg) {
  $('connectBanner').hidden = false;
  if (msg) $('connectMsg').textContent = msg;
}

// ---------- 데이터 불러오기 ----------
async function load() {
  const s = Sheets.settings();
  if (!s.clientId || !s.sheetId) {
    showConnect('처음이시군요. 설정에서 Client ID와 구글시트를 등록해 주세요.');
    $('connectBtn').textContent = '설정 열기';
    render();
    return;
  }
  $('connectBtn').textContent = '구글 연결';
  if (!Sheets.token()) { showConnect('구글 연결이 필요해요. (보안상 1시간마다 다시 연결해요)'); render(); return; }
  $('connectBanner').hidden = true;
  $('status').textContent = '불러오는 중…';
  try {
    await Sheets.ensureSheet();
    state.entries = await Sheets.list();
    $('status').textContent = '';
  } catch (err) {
    $('status').textContent = '';
    if (err.auth) showConnect(err.message);
    else toast(err.message);
  }
  render();
}

async function connect() {
  const s = Sheets.settings();
  if (!s.clientId || !s.sheetId) return openSettings();
  try { await Sheets.signIn(); await load(); }
  catch (err) { toast(err.message); }
}

// ---------- 입력/수정 ----------
function chips(container, options, selected, onPick) {
  container.replaceChildren(...options.map((opt) => {
    const c = document.createElement('button');
    c.type = 'button';
    c.className = 'chip';
    c.textContent = opt;
    c.setAttribute('aria-pressed', String(opt === selected));
    c.addEventListener('click', () => {
      const next = opt === selected ? '' : opt;
      onPick(next);
    });
    return c;
  }));
}
function currentType() { return $('entryForm').querySelector('input[name="type"]:checked').value; }
function drawChips() {
  chips($('categoryChips'), CATEGORIES[currentType()], state.form.category, (v) => { state.form.category = v; drawChips(); });
  chips($('methodChips'), METHODS, state.form.method, (v) => { state.form.method = v; drawChips(); });
}

function openEntry(e = null) {
  state.editing = e;
  $('entryTitle').textContent = e ? '내역 수정' : '내역 추가';
  $('deleteBtn').hidden = !e;
  $('entryError').textContent = '';
  $('entryForm').querySelector(`input[name="type"][value="${e ? e.type : '지출'}"]`).checked = true;
  $('amount').value = e ? e.amount.toLocaleString('ko-KR') : '';
  // 다른 달을 보고 있을 때 새 내역은 그 달 1일이 아니라 오늘 날짜가 기본값
  $('date').value = e ? e.date : today();
  $('place').value = e ? e.place : '';
  $('memo').value = e ? e.memo : '';
  state.form = { category: e ? e.category : '', method: e ? e.method : '' };
  drawChips();
  $('entryDialog').showModal();
  if (!e) $('amount').focus();
}

$('amount').addEventListener('input', (ev) => {
  const digits = ev.target.value.replace(/\D/g, '').replace(/^0+/, '').slice(0, 12);
  ev.target.value = digits ? Number(digits).toLocaleString('ko-KR') : '';
});
$('entryForm').querySelectorAll('input[name="type"]').forEach((r) => r.addEventListener('change', () => {
  if (!CATEGORIES[currentType()].includes(state.form.category)) state.form.category = '';
  drawChips();
}));

$('entryForm').addEventListener('submit', async (ev) => {
  ev.preventDefault();
  const amount = Number($('amount').value.replace(/\D/g, ''));
  if (!amount) { $('entryError').textContent = '금액을 입력해 주세요.'; return; }
  if (!state.form.category) { $('entryError').textContent = '카테고리를 골라 주세요.'; return; }
  if (!Sheets.isReady()) { $('entryError').textContent = '구글 연결이 필요해요. 닫고 상단의 [구글 연결]을 눌러 주세요.'; return; }

  const e = {
    id: state.editing ? state.editing.id : (crypto.randomUUID ? crypto.randomUUID() : String(Date.now()) + Math.random().toString(16).slice(2)),
    date: $('date').value,
    type: currentType(),
    amount,
    category: state.form.category,
    method: state.form.method,
    place: $('place').value.trim(),
    memo: $('memo').value.trim(),
    createdAt: state.editing ? state.editing.createdAt : new Date().toISOString(),
  };
  $('saveBtn').disabled = true;
  try {
    if (state.editing) {
      await Sheets.update(e);
      state.entries = state.entries.map((x) => (x.id === e.id ? e : x));
    } else {
      await Sheets.add(e);
      state.entries.push(e);
    }
    $('entryDialog').close();
    state.month = e.date.slice(0, 7);
    render();
    toast(state.editing ? '수정했어요' : '저장했어요');
  } catch (err) {
    $('entryError').textContent = err.message;
    if (err.auth) showConnect(err.message);
  } finally {
    $('saveBtn').disabled = false;
  }
});

$('deleteBtn').addEventListener('click', async () => {
  const e = state.editing;
  if (!e || !confirm('이 내역을 삭제할까요?')) return;
  $('deleteBtn').disabled = true;
  try {
    await Sheets.remove(e.id);
    state.entries = state.entries.filter((x) => x.id !== e.id);
    $('entryDialog').close();
    render();
    toast('삭제했어요');
  } catch (err) {
    $('entryError').textContent = err.message;
  } finally {
    $('deleteBtn').disabled = false;
  }
});

// ---------- 설정 ----------
function openSettings() {
  const s = Sheets.settings();
  $('clientId').value = s.clientId;
  $('sheetInput').value = s.sheetId;
  $('settingsError').textContent = '';
  const link = $('openSheetLink');
  link.hidden = !s.sheetId;
  if (s.sheetId) link.href = `https://docs.google.com/spreadsheets/d/${encodeURIComponent(s.sheetId)}/edit`;
  $('settingsDialog').showModal();
}
function saveSettingsFromForm() {
  Sheets.saveSettings({ clientId: $('clientId').value.trim(), sheetId: parseSheetId($('sheetInput').value) });
}

$('settingsForm').addEventListener('submit', async (ev) => {
  ev.preventDefault();
  saveSettingsFromForm();
  const s = Sheets.settings();
  if (!s.clientId) { $('settingsError').textContent = 'Client ID를 입력해 주세요.'; return; }
  if (!s.sheetId) { $('settingsError').textContent = '시트 주소를 넣거나 [새 시트 만들기]를 눌러 주세요.'; return; }
  try {
    if (!Sheets.token()) await Sheets.signIn();
    $('settingsDialog').close();
    await load();
  } catch (err) { $('settingsError').textContent = err.message; }
});

$('createSheetBtn').addEventListener('click', async () => {
  saveSettingsFromForm();
  if (!Sheets.settings().clientId) { $('settingsError').textContent = 'Client ID를 먼저 입력해 주세요.'; return; }
  if (Sheets.settings().sheetId && !confirm('이미 등록된 시트가 있어요. 새 시트로 바꿀까요? (기존 시트는 지워지지 않아요)')) return;
  try {
    if (!Sheets.token()) await Sheets.signIn();
    await Sheets.create();
    $('settingsDialog').close();
    toast('내 구글 드라이브에 \'가계부\' 시트를 만들었어요');
    await load();
  } catch (err) { $('settingsError').textContent = err.message; }
});

$('signOutBtn').addEventListener('click', () => {
  Sheets.signOut();
  state.entries = [];
  $('settingsDialog').close();
  load();
  toast('구글 연결을 해제했어요');
});

// ---------- 공통 ----------
document.querySelectorAll('[data-close]').forEach((b) => b.addEventListener('click', () => b.closest('dialog').close()));
document.querySelectorAll('dialog').forEach((d) => d.addEventListener('click', (ev) => { if (ev.target === d) d.close(); }));

function shiftMonth(delta) {
  const [y, m] = state.month.split('-').map(Number);
  const d = new Date(y, m - 1 + delta, 1);
  state.month = `${d.getFullYear()}-${pad(d.getMonth() + 1)}`;
  render();
}
$('prevMonth').addEventListener('click', () => shiftMonth(-1));
$('nextMonth').addEventListener('click', () => shiftMonth(1));
$('openSettings').addEventListener('click', openSettings);
$('connectBtn').addEventListener('click', connect);
$('addBtn').addEventListener('click', () => openEntry());

load();
