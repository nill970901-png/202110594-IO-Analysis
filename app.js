'use strict';
/* 산업연관표 분석 웹앱 — 화면 동작. 계산은 worker.js(Pyodide) → model.py 에서 수행 */
const $ = id => document.getElementById(id);
const LEVELS = [
  { key: 'large', label: '대분류', n: 33, file: 'io-2023-large.xlsx', size: '0.2MB', desc: '농림수산품·화학제품 등 큰 산업 묶음' },
  { key: 'medium', label: '중분류', n: 83, file: 'io-2023-medium.xlsx', size: '0.7MB', desc: '철강제품·반도체·보험서비스 수준' },
  { key: 'small', label: '소분류', n: 165, file: 'io-2023-small.xlsx', size: '2.4MB', desc: '열간압연강재·자동차 부품 수준' },
  { key: 'basic', label: '기본부문', n: 380, file: 'io-2023-basic.xlsx', size: '12MB', desc: '한국은행 원표 · 가장 세밀 (로딩이 김)' },
];
const FD = [['9111', '민간소비지출'], ['9112', '정부소비지출'], ['9121', '민간고정자본형성'], ['9122', '정부고정자본형성'],
  ['9131', '재고증감'], ['9132', '귀중품순취득'], ['9140', '수출']];
const VA_NAME = { '9610': '피용자보수', '9621': '영업잉여', '9622': '고정자본소모', '9630': '생산세(보조금공제)' };

const nf = d => new Intl.NumberFormat('ko-KR', { maximumFractionDigits: d, minimumFractionDigits: 0 });
const fmt = (x, d = 3) => x === null || x === undefined ? '—' : nf(d).format(x);
const fmtM = x => fmt(x, 0);                                   // 백만원 정수
const fmtP = x => fmt(x, 2);                                    // %
const fmtC = x => fmt(x, 4);                                    // 계수
function human(m) {                                             // 백만원 → 읽기 쉬운 단위
  const a = Math.abs(m);
  if (a >= 1e6) return fmt(m / 1e6, 2) + '조원';
  if (a >= 100) return fmt(m / 100, 1) + '억원';
  return fmt(m, 0) + '백만원';
}
function el(tag, attrs = {}, ...kids) {
  const e = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (k === 'class') e.className = v; else if (k === 'text') e.textContent = v;
    else if (k.startsWith('on')) e[k] = v; else e.setAttribute(k, v);
  }
  for (const k of kids) e.append(k);
  return e;
}
const esc = t => String(t).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const match = (q, ...s) => !q || s.join(' ').toLowerCase().includes(q.toLowerCase());
function options(sel, list, extra) {
  sel.replaceChildren(...(extra || []).map(([v, t]) => el('option', { value: v, text: t })),
    ...list.map(x => el('option', { value: x.code, text: x.code + ' ' + x.name })));
}

/* ---------------------------------------------------------------- worker 통신 */
let worker, nextId = 0;
const pending = new Map();
function startWorker() {
  worker = new Worker('./worker.js');
  worker.onmessage = ({ data: d }) => {
    if (d.type === 'status') { setStatus(d.text); return; }
    const p = pending.get(d.id); if (!p) return; pending.delete(d.id);
    d.ok ? p.resolve(d.result) : p.reject(Error(d.error));
  };
  worker.onerror = e => { for (const p of pending.values()) p.reject(Error(e.message || 'Worker 실행 오류')); pending.clear(); };
}
function call(type, payload, transfer) {
  return new Promise((resolve, reject) => {
    const id = ++nextId; pending.set(id, { resolve, reject });
    worker.postMessage({ id, type, payload }, transfer || []);
  });
}
const api = (fn, ...args) => call('call', { fn, args });
const setStatus = t => { $('status').textContent = t; };

/* ---------------------------------------------------------------- 상태 */
const S = { loaded: {}, active: null, overview: {}, linkage: {}, basket: [], pbasket: [], sim: null, price: null, classes: [] };
const cur = () => S.loaded[S.active];
const sectors = () => cur() ? cur().sectors : [];

/* ---------------------------------------------------------------- 0. 자료 불러오기 */
function renderLevels() {
  $('levels').replaceChildren(...LEVELS.map(L => {
    const st = S.loaded[L.key] ? 'done' : L.busy ? 'busy' : 'idle';
    const btn = el('button', { type: 'button', class: st === 'done' ? 'ghost' : '', text: st === 'done' ? '불러옴 ✓' : st === 'busy' ? '불러오는 중…' : '불러오기' });
    btn.disabled = st !== 'idle' || LEVELS.some(x => x.busy);
    btn.onclick = () => loadLevel(L.key);
    const input = el('input', { type: 'file', accept: '.xlsx' });
    input.disabled = LEVELS.some(x => x.busy);
    input.onchange = () => { if (input.files[0]) loadLevel(L.key, input.files[0]); };
    const pick = el('label', { class: 'filepick', title: '가지고 있는 ' + L.label + ' Excel로 바꿔 불러오기' }, '내 Excel 선택', input);
    return el('div', { class: 'level ' + st },
      el('div', { class: 'level-top' }, el('strong', { text: L.label }), el('span', { class: 'pill', text: L.n + '부문' })),
      el('p', { text: L.desc }),
      el('p', { class: 'note', text: L.source ? '불러온 파일: ' + L.source : 'Excel ' + L.size + ' · data/' + L.file }),
      el('div', { class: 'level-btns' }, btn, pick));
  }));
}
async function loadLevel(key, file) {
  const L = LEVELS.find(x => x.key === key);
  if ((S.loaded[key] && !file) || L.busy) return;
  L.busy = true; renderLevels();
  try {
    let buffer;
    if (file) {
      setStatus(L.label + ' — 선택한 파일(' + file.name + ') 읽는 중…');
      buffer = await file.arrayBuffer();
    } else {
      setStatus(L.label + ' Excel 내려받는 중…');
      const res = await fetch('./data/' + L.file);
      if (!res.ok) throw Error('Excel 다운로드 실패 (HTTP ' + res.status + ')');
      buffer = await res.arrayBuffer();
    }
    const info = await call('load', { level: key, label: L.label, buffer }, [buffer]);
    const replaced = !!S.loaded[key];
    S.loaded[key] = info; L.source = file ? file.name : 'data/' + L.file;
    delete S.overview[key]; delete S.linkage[key];
    if (replaced && S.active === key) {
      await activate(key);
      setStatus(L.label + ' 자료를 ' + L.source + '(으)로 바꿨습니다 · ' + info.sectors.length + '개 부문 · 계수 검증 통과');
      renderValidation(); updateCompareHint();
      return;
    }
    renderValidation();
    const first = !S.active;
    const sel = $('active-level');
    sel.replaceChildren(...LEVELS.filter(x => S.loaded[x.key]).map(x => el('option', { value: x.key, text: x.label + ' (' + x.n + '부문)' })));
    sel.disabled = false;
    if (first) { S.active = key; }
    sel.value = S.active;
    setStatus(L.label + ' 준비 완료 · ' + info.sectors.length + '개 부문 · ' + info.sheets.length + '개 시트 · 계수 검증 통과' +
      (first ? '' : ' — 위 "분석 기준 분류"에서 바꿔 볼 수 있습니다.'));
    if (first) await activate(key);
    $('cmp-run').disabled = false;
    updateCompareHint();
  } catch (e) {
    setStatus(L.label + ' 준비 실패: ' + e.message + ' — 새로고침 후 다시 시도하세요.');
  } finally { L.busy = false; renderLevels(); }
}
function renderValidation() {
  const out = {};
  for (const L of LEVELS) if (S.loaded[L.key]) out[L.label] = S.loaded[L.key].residuals;
  $('validation').textContent = JSON.stringify(out, null, 2) + '\n\n모든 값이 허용 오차 1e-6 미만이어야 통과합니다.';
}
$('load-all').onclick = async () => { $('load-all').disabled = true; for (const L of LEVELS) await loadLevel(L.key); $('load-all').disabled = false; };
$('active-level').onchange = e => activate(e.target.value);

async function activate(key) {
  S.active = key;
  const info = cur(), label = LEVELS.find(x => x.key === key).label;
  document.querySelectorAll('.needs-data').forEach(s => s.dataset.level = label + ' ' + info.sectors.length + '부문 기준');
  for (const id of ['sheet', 'row', 'row-search', 'col-search', 'metric', 'topn', 'structure-search', 'structure-csv',
    'sim-search', 'sim-code', 'sim-amount', 'sim-add', 'item-code', 'item-sector', 'item-pct', 'item-add', 'sim-run',
    'quad-filter', 'link-search', 'pf-search', 'pf-code', 'pf-run', 'pi-search', 'pi-code', 'pi-pct', 'pi-add', 'pv-comp', 'pv-code', 'pv-pct', 'pv-add', 'price-run']) $(id).disabled = false;
  // 시트 목록
  $('sheet').replaceChildren(...info.sheets.map(s => el('option', { value: s, text: s })));
  $('row-search').value = ''; $('col-search').value = '';
  // 산업 선택 목록
  options($('sim-code'), info.sectors); options($('pi-code'), info.sectors); options($('pf-code'), info.sectors);
  $('pf-search').value = ''; $('pf-out').hidden = true;
  const semi = info.sectors.find(s => s.name === '반도체') || info.sectors.find(s => /집적회로|반도체|컴퓨터, 전자/.test(s.name));
  if (semi) $('pf-code').value = semi.code;
  options($('item-sector'), info.sectors, [['ALL', '전체 산업']]); options($('pv-code'), info.sectors, [['ALL', '전체 산업']]);
  options($('item-code'), FD.map(([code, name]) => ({ code, name })));
  $('item-code').value = '9140';
  const oil = info.sectors.find(s => /원유/.test(s.name)) || info.sectors.find(s => /광산|석탄, 원유/.test(s.name));
  if (oil) $('pi-code').value = oil.code;
  // 분류가 바뀌면 산업 코드가 달라지므로 시나리오 초기화
  if (S.basket.length || S.pbasket.length) $('sim-status').textContent = '분석 기준 분류가 바뀌어 시나리오를 비웠습니다.';
  S.basket = []; S.pbasket = []; renderBasket(); renderPBasket(); S.sim = null; S.price = null;
  $('sim-results').hidden = true; $('price-results').hidden = true;
  await Promise.all([explore(), loadStructure(), loadLinkage()]);
}

/* ---------------------------------------------------------------- 1. 원자료 탐색 */
let exploreSeq = 0, rowChoices = [], cells = [];
async function explore(rowCode) {
  const seq = ++exploreSeq;
  try {
    const d = await api('explore', S.active, $('sheet').value, rowCode || '');
    if (seq !== exploreSeq) return;
    rowChoices = d.rows; filterRows(d.selected); cells = d.cells;
    $('unit').textContent = d.unit;
    $('explore-caption').textContent = d.title + ' — 행: ' + d.selected + ' ' + d.name + ' (' + d.cells.length + '개 열)';
    renderCells();
  } catch (e) { $('explore-caption').textContent = '탐색 오류: ' + e.message; }
}
function filterRows(selected) {
  const q = $('row-search').value.trim();
  options($('row'), rowChoices.filter(r => match(q, r.code, r.name)));
  if ([...$('row').options].some(o => o.value === selected)) $('row').value = selected;
}
function renderCells() {
  const q = $('col-search').value.trim();
  $('cells').replaceChildren(...cells.filter(c => match(q, c.code, c.name)).map(c => el('tr', {},
    el('td', { text: c.code }), el('td', { text: c.name }),
    el('td', { class: 'r', text: c.value === null ? '빈 셀' : typeof c.value === 'number' ? fmt(c.value, Math.abs(c.value) >= 100 ? 0 : 6) : c.value }))));
}
$('sheet').onchange = () => { $('row-search').value = ''; explore(); };
$('row').onchange = () => explore($('row').value);
$('row-search').oninput = () => {
  filterRows(); if ($('row').value) explore($('row').value);
  else { $('cells').replaceChildren(); $('explore-caption').textContent = '일치하는 행이 없습니다.'; }
};
$('col-search').oninput = renderCells;

/* ---------------------------------------------------------------- 공통: 막대그래프·정렬표·CSV */
function bars(target, items, unitFn) {
  const max = Math.max(...items.map(i => Math.abs(i.value)), 1e-12);
  const signed = items.some(i => i.value < 0);
  target.replaceChildren(...items.map(i => {
    const w = Math.abs(i.value) / max * (signed ? 50 : 100);
    const fill = el('div', { class: 'fill' + (i.value < 0 ? ' neg' : '') });
    fill.style.width = w + '%';
    if (signed) fill.style.marginLeft = i.value < 0 ? (50 - w) + '%' : '50%';
    return el('div', { class: 'bar' },
      el('div', { class: 'bar-label' }, el('span', { text: i.label }), el('b', { text: unitFn(i.value) })),
      el('div', { class: 'track' + (signed ? ' signed' : '') }, fill));
  }));
}
function sortableTable(head, body, cols, rows, state, rerender) {
  head.replaceChildren(el('tr', {}, ...cols.map(c => {
    const th = el('th', { class: (c.num ? 'r ' : '') + 'sort' + (state.key === c.key ? (state.asc ? ' asc' : ' desc') : ''), text: c.label, title: c.title || '정렬' });
    th.onclick = () => { state.asc = state.key === c.key ? !state.asc : !c.num; state.key = c.key; rerender(); };
    return th;
  })));
  const s = [...rows].sort((a, b) => {
    const x = a[state.key], y = b[state.key];
    const r = typeof x === 'number' ? x - y : String(x).localeCompare(String(y), 'ko');
    return state.asc ? r : -r;
  });
  body.replaceChildren(...s.map(r => el('tr', {}, ...cols.map(c => el('td', { class: c.num ? 'r' : '', text: c.fmt ? c.fmt(r[c.key]) : r[c.key] })))));
}
function csvCell(v) { let s = String(v ?? ''); if (/^[=+@-]/.test(s) && typeof v !== 'number') s = "'" + s; return '"' + s.replaceAll('"', '""') + '"'; }
function downloadCSV(name, header, rows) {
  const text = '﻿' + [header, ...rows].map(r => r.map(csvCell).join(',')).join('\r\n');
  const url = URL.createObjectURL(new Blob([text], { type: 'text/csv;charset=utf-8' }));
  const a = el('a', { href: url, download: name }); document.body.append(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
function cards(target, list) {
  target.replaceChildren(...list.map(([t, v, sub]) => el('div', { class: 'card' }, el('span', { text: t }), el('strong', { text: v }), el('small', { text: sub || '' }))));
}

/* ---------------------------------------------------------------- 2. 산업구조 */
const METRICS = [
  { key: 'output', label: '총산출', unit: '백만원', fmt: fmtM, desc: '산업이 1년간 생산한 총액' },
  { key: 'valueAdded', label: '부가가치', unit: '백만원', fmt: fmtM, desc: '피용자보수+영업잉여+고정자본소모+순생산세' },
  { key: 'vaRate', label: '부가가치율', unit: '%', fmt: fmtP, desc: '부가가치 ÷ 총산출' },
  { key: 'laborShare', label: '피용자보수 비중', unit: '%', fmt: fmtP, desc: '피용자보수 ÷ 부가가치' },
  { key: 'exportRate', label: '수출 비중', unit: '%', fmt: fmtP, desc: '수출 ÷ 총산출' },
  { key: 'importDependence', label: '수입 의존도', unit: '%', fmt: fmtP, desc: '수입 ÷ 총공급' },
  { key: 'importInputRate', label: '수입중간투입률', unit: '%', fmt: fmtP, desc: '수입 중간재 ÷ 총투입' },
  { key: 'prodMultiplier', label: '생산유발계수(열합)', unit: '배', fmt: fmtC, desc: '최종수요 1원당 전 산업 생산' },
  { key: 'vaMultiplier', label: '부가가치유발계수', unit: '배', fmt: fmtC, desc: '최종수요 1원당 전 산업 부가가치' },
  { key: 'importMultiplier', label: '수입유발계수', unit: '배', fmt: fmtC, desc: '최종수요 1원당 유발 수입' },
];
$('metric').replaceChildren(...METRICS.map(m => el('option', { value: m.key, text: m.label + ' (' + m.unit + ')' })));
const structState = { key: 'output', asc: false };
async function loadStructure() {
  if (!S.overview[S.active]) S.overview[S.active] = await api('overview', S.active);
  const o = S.overview[S.active], t = o.total;
  cards($('structure-cards'), [
    ['총산출', human(t.output), fmtM(t.output) + '백만원'], ['부가가치', human(t.valueAdded), '부가가치율 ' + fmtP(t.vaRate) + '%'],
    ['수출', human(t.exports), '총산출 대비 ' + fmtP(t.exports / t.output * 100) + '%'], ['수입', human(t.imports), '피용자보수 비중 ' + fmtP(t.laborShare) + '%'],
  ]);
  renderStructure();
}
function renderStructure() {
  const o = S.overview[S.active]; if (!o) return;
  const m = METRICS.find(x => x.key === $('metric').value), n = +$('topn').value, q = $('structure-search').value.trim();
  const rows = o.rows.filter(r => match(q, r.code, r.name));
  const top = [...rows].sort((a, b) => b[m.key] - a[m.key]).slice(0, n);
  $('metric-title').textContent = m.label + ' 상위 ' + top.length + '개 산업 (' + m.unit + ') — ' + m.desc;
  bars($('metric-bars'), top.map(r => ({ label: r.code + ' ' + r.name, value: r[m.key] })), v => m.fmt(v) + (m.unit === '배' ? '' : ' ' + m.unit));
  sortableTable($('structure-head'), $('structure-rows'), [
    { key: 'code', label: '코드' }, { key: 'name', label: '산업' },
    ...METRICS.map(x => ({ key: x.key, label: x.label + (x.unit === '배' ? '' : '(' + x.unit + ')'), num: true, fmt: x.fmt, title: x.desc })),
  ], rows, structState, renderStructure);
}
for (const id of ['metric', 'topn']) $(id).onchange = renderStructure;
$('structure-search').oninput = renderStructure;
$('structure-csv').onclick = () => {
  const o = S.overview[S.active];
  downloadCSV('io-2023-' + S.active + '-structure.csv', ['코드', '산업', ...METRICS.map(m => m.label + '_' + m.unit), '영향력계수', '감응도계수'],
    o.rows.map(r => [r.code, r.name, ...METRICS.map(m => r[m.key]), r.BL, r.FL]));
};

/* ---------------------------------------------------------------- 3. 산업 공급망 리포트 */
$('pf-search').oninput = () => options($('pf-code'), sectors().filter(s => match($('pf-search').value.trim(), s.code, s.name)));
$('pf-code').onchange = () => { if (!$('pf-out').hidden) runProfile(); };
$('pf-run').onclick = () => runProfile();
async function runProfile() {
  const code = $('pf-code').value; if (!code) return;
  $('pf-run').disabled = true;
  try {
    const p = await api('profile', S.active, code);
    $('pf-out').hidden = false;
    $('pf-title').textContent = p.code + ' ' + p.name + ' — ' + p.level + ' ' + p.n + '개 부문 중';
    cards($('pf-cards'), [
      ['총산출', human(p.output), '전체의 ' + fmtP(p.outputShare) + '% · ' + p.outputRank + '위'],
      ['부가가치율', fmtP(p.vaRate) + '%', '영업잉여 ' + human(p.profit)],
      ['수출 비중', fmtP(p.exportRate) + '%', '총산출 대비'],
      ['수입 의존도', fmtP(p.importDependence) + '%', '총공급 대비 수입'],
      ['생산유발계수', fmt(p.prodMultiplier, 3) + '배', '평균 ' + fmt(p.prodMultiplierAvg, 3) + ' · ' + p.prodRank + '위'],
      ['영향력 / 감응도', fmt(p.BL, 2) + ' / ' + fmt(p.FL, 2), '1보다 크면 평균 이상'],
    ]);
    // 자동 해석 문장
    const msgs = [];
    if (p.exportRate >= 40) msgs.push('생산의 ' + fmt(p.exportRate, 0) + '%를 수출하는 <b>수출 주도형</b> 산업이라 해외 경기와 환율에 민감합니다.');
    else if (p.final.find(f => f.code === '9111' && f.share >= 40)) msgs.push('국산 판매의 상당 부분이 <b>민간소비</b>로 가므로 내수 경기에 민감합니다.');
    if (p.impInputRate >= 25) msgs.push('산출 1원당 수입 중간재가 ' + fmt(p.impInputRate / 100, 2) + '원으로 <b>수입 원자재 의존</b>이 커서 환율·원자재 가격 상승 시 원가 부담이 큽니다' + (p.imported[0] ? ' (최대 수입 투입: ' + esc(p.imported[0].name) + ').' : '.'));
    if (p.BL >= 1 && p.FL >= 1) msgs.push('영향력·감응도계수가 모두 1 이상인 <b>핵심 산업</b>입니다.');
    else if (p.BL >= 1) msgs.push('영향력계수가 1 이상이라 이 산업이 성장하면 <b>다른 산업 생산을 많이 끌어올립니다</b>.');
    else if (p.FL >= 1) msgs.push('감응도계수가 1 이상인 <b>중간재 공급형</b> 산업이라 전체 경기에 함께 움직입니다.');
    else msgs.push('생산유발계수가 평균보다 낮아 수요 증가의 상당 부분이 수입이나 자기 산업에 머뭅니다.');
    if (p.profit < 0) msgs.push('2023년 영업잉여가 <b>마이너스</b>로, 해당 연도 업황이 부진했음을 보여줍니다.');
    $('pf-summary').innerHTML = msgs.join(' ');
    const lab = r => r.code + ' ' + r.name + (r.self ? ' (자기 산업)' : '');
    bars($('pf-sup'), p.suppliers.map(r => ({ label: lab(r), value: r.coef })), v => fmt(v, 4));
    bars($('pf-imp'), p.imported.map(r => ({ label: lab(r), value: r.coef })), v => fmt(v, 4));
    bars($('pf-buy'), p.buyers.map(r => ({ label: r.code + ' ' + r.name, value: r.share })), v => fmtP(v) + '%');
    bars($('pf-fd'), p.final.map(r => ({ label: r.name, value: r.share })), v => fmtP(v) + '%');
    bars($('pf-spill'), p.spill.map(r => ({ label: r.code + ' ' + r.name, value: r.amount })), v => fmtM(v));
  } catch (e) { $('pf-title').textContent = '리포트 오류: ' + e.message; $('pf-out').hidden = false; }
  finally { $('pf-run').disabled = false; }
}

/* ---------------------------------------------------------------- 4. 수요충격 */
document.querySelectorAll('[data-tab]').forEach(b => b.onclick = () => {
  document.querySelectorAll('[data-tab]').forEach(x => x.classList.toggle('on', x === b));
  $('sim-sector').hidden = b.dataset.tab !== 'sector'; $('sim-item').hidden = b.dataset.tab !== 'item';
});
$('sim-search').oninput = () => options($('sim-code'), sectors().filter(s => match($('sim-search').value.trim(), s.code, s.name)));
function shockLabel(s) {
  if (s.type === 'item') {
    const item = FD.find(f => f[0] === s.item)[1];
    const sec = s.code === 'ALL' ? '전체 산업' : s.code + ' ' + sectors().find(x => x.code === s.code).name;
    return item + ' · ' + sec + ' · ' + (s.pct > 0 ? '+' : '') + fmt(s.pct, 2) + '%';
  }
  return s.code + ' ' + s.name + ' · ' + (s.amount > 0 ? '+' : '') + fmtM(s.amount) + '백만원 (' + human(s.amount) + ')';
}
function renderBasket() {
  $('basket').replaceChildren(...S.basket.map((s, i) => el('li', {}, el('span', { text: shockLabel(s) }),
    el('button', { type: 'button', class: 'mini', text: '삭제', onclick: () => { S.basket.splice(i, 1); invalidateSim(); renderBasket(); } }))));
  $('sim-status').textContent = S.basket.length ? S.basket.length + '개 입력 준비. "파급효과 계산"을 누르세요.' : '시나리오를 추가하세요.';
}
const invalidateSim = () => { S.sim = null; $('sim-results').hidden = true; };
$('sim-sector').onsubmit = e => {
  e.preventDefault();
  const amount = Number($('sim-amount').value), s = sectors().find(x => x.code === $('sim-code').value);
  if (!s || !Number.isFinite(amount) || amount === 0) { $('sim-status').textContent = '산업과 0이 아닌 변화액을 확인하세요.'; return; }
  S.basket.push({ type: 'sector', code: s.code, name: s.name, amount }); invalidateSim(); renderBasket();
};
$('sim-item').onsubmit = e => {
  e.preventDefault();
  const pct = Number($('item-pct').value);
  if (!Number.isFinite(pct) || pct === 0) { $('sim-status').textContent = '0이 아닌 변화율을 입력하세요.'; return; }
  S.basket.push({ type: 'item', item: $('item-code').value, code: $('item-sector').value, pct }); invalidateSim(); renderBasket();
};
$('sim-clear').onclick = () => { S.basket = []; invalidateSim(); renderBasket(); };
$('sim-run').onclick = async () => {
  if (!S.basket.length) { $('sim-status').textContent = '시나리오를 먼저 추가하세요.'; return; }
  const snap = JSON.stringify(S.basket), level = S.active;
  $('sim-run').disabled = true; $('sim-status').textContent = '계산 중…';
  try {
    const r = await api('simulate', level, snap);
    if (snap !== JSON.stringify(S.basket) || level !== S.active) { $('sim-status').textContent = '입력이 바뀌었습니다. 다시 계산하세요.'; return; }
    S.sim = r; renderSim();
    $('sim-status').textContent = '계산 완료 (' + LEVELS.find(x => x.key === level).label + ' 기준). 금액 단위는 백만원입니다.';
  } catch (e) { invalidateSim(); $('sim-status').textContent = '계산 실패: ' + e.message; }
  finally { $('sim-run').disabled = false; }
};
function renderSim() {
  const r = S.sim; $('sim-results').hidden = false;
  cards($('sim-cards'), [
    ['국산품 최종수요 변화', human(r.demand), fmtM(r.demand) + '백만원'],
    ['생산유발액', human(r.production), '승수 ' + fmt(r.multiplier, 3) + '배'],
    ['부가가치유발액', human(r.valueAdded), '2023년 부가가치 대비 ' + fmt(r.gdpShare, 3) + '%'],
    ['수입유발액', human(r.imports), r.directImports ? '최종수요 직접수입 ' + human(r.directImports) + ' 별도' : '중간재 수입'],
    ['피용자보수 유발액', human(r.wages), '부가가치 중 ' + fmtP(r.wages / r.valueAdded * 100) + '%'],
  ]);
  bars($('sim-bars'), r.rows.slice(0, 15).map(x => ({ label: x.code + ' ' + x.name, value: x.production })), v => fmtM(v) + '백만원');
  $('sim-rows').replaceChildren(...r.rows.map(x => el('tr', {}, ...[x.code, x.name, fmtM(x.demand), fmtM(x.production), fmt(x.pct, 3), fmtM(x.valueAdded), fmtM(x.imports), fmtM(x.wages)]
    .map((v, i) => el('td', { class: i > 1 ? 'r' : '', text: v })))));
}
$('sim-csv').onclick = () => {
  if (!S.sim) return;
  downloadCSV('io-2023-' + S.active + '-simulation.csv', ['산업코드', '산업명', '직접수요_백만원', '생산유발_백만원', '생산변화율_%', '부가가치유발_백만원', '수입유발_백만원', '피용자보수유발_백만원'],
    S.sim.rows.map(x => [x.code, x.name, x.demand, x.production, x.pct, x.valueAdded, x.imports, x.wages]));
};

/* ---------------------------------------------------------------- 4. 전후방 연관효과 */
const QCOLOR = { I: 'var(--q1)', II: 'var(--q2)', III: 'var(--q3)', IV: 'var(--q4)' };
const linkState = { key: 'BL', asc: false };
async function loadLinkage() {
  if (!S.linkage[S.active]) S.linkage[S.active] = await api('linkage', S.active);
  const d = S.linkage[S.active];
  $('quad-filter').replaceChildren(el('option', { value: '', text: '전체' }), ...Object.entries(d.quadNames).map(([k, v]) => el('option', { value: k, text: k + ' ' + v })));
  $('quad-cards').replaceChildren(...['II', 'I', 'III', 'IV'].map(k => {
    const n = d.rows.filter(r => r.quad === k).length;
    const c = el('button', { type: 'button', class: 'quad', onclick: () => { $('quad-filter').value = $('quad-filter').value === k ? '' : k; renderLinkage(); } },
      el('span', { class: 'dot' }), el('b', { text: '제' + k + '사분면 · ' + n + '개' }), el('small', { text: d.quadNames[k] }));
    c.querySelector('.dot').style.background = QCOLOR[k];
    return c;
  }));
  renderLinkage();
}
function renderLinkage() {
  const d = S.linkage[S.active]; if (!d) return;
  const qf = $('quad-filter').value, q = $('link-search').value.trim();
  const rows = d.rows.filter(r => (!qf || r.quad === qf) && match(q, r.code, r.name));
  drawScatter(d.rows, new Set(rows.map(r => r.code)));
  sortableTable($('link-head'), $('link-rows'), [
    { key: 'code', label: '코드' }, { key: 'name', label: '산업' },
    { key: 'BL', label: '영향력계수', num: true, fmt: fmtC }, { key: 'FL', label: '감응도계수', num: true, fmt: fmtC },
    { key: 'output', label: '총산출(백만원)', num: true, fmt: fmtM }, { key: 'quadName', label: '유형' },
  ], rows, linkState, renderLinkage);
}
$('quad-filter').onchange = renderLinkage; $('link-search').oninput = renderLinkage;
function drawScatter(all, keep) {
  const svg = $('scatter'), W = 640, H = 460, P = { l: 56, r: 18, t: 18, b: 48 };
  const xs = all.map(r => r.BL), ys = all.map(r => r.FL);
  // 가로축(영향력계수)은 선형, 세로축(감응도계수)은 분포가 넓어 로그 눈금
  const x0 = Math.min(0.5, ...xs) - 0.05, x1 = Math.max(1.5, ...xs) + 0.05, y0 = Math.min(0.5, ...ys) * 0.92, y1 = Math.max(1.5, ...ys) * 1.12;
  const X = v => P.l + (v - x0) / (x1 - x0) * (W - P.l - P.r);
  const Y = v => H - P.b - (Math.log(v) - Math.log(y0)) / (Math.log(y1) - Math.log(y0)) * (H - P.t - P.b);
  const maxO = Math.max(...all.map(r => r.output));
  const ns = 'http://www.w3.org/2000/svg', mk = (t, a) => { const e = document.createElementNS(ns, t); for (const k in a) e.setAttribute(k, a[k]); return e; };
  const g = [];
  const ticks = (a, b) => { const st = b - a > 1.6 ? 0.5 : 0.25, out = []; for (let v = Math.ceil(a / st) * st; v <= b; v += st) out.push(+v.toFixed(2)); return out; };
  const yticks = [0.25, 0.3, 0.4, 0.5, 0.6, 0.75, 1, 1.25, 1.5, 2, 3, 4, 5, 6, 8, 10, 15].filter(v => v >= y0 && v <= y1);
  for (const v of ticks(x0, x1)) { g.push(mk('line', { x1: X(v), x2: X(v), y1: P.t, y2: H - P.b, class: 'grid' })); const t = mk('text', { x: X(v), y: H - P.b + 18, class: 'tick', 'text-anchor': 'middle' }); t.textContent = v.toFixed(2); g.push(t); }
  for (const v of yticks) { g.push(mk('line', { x1: P.l, x2: W - P.r, y1: Y(v), y2: Y(v), class: 'grid' })); const t = mk('text', { x: P.l - 8, y: Y(v) + 4, class: 'tick', 'text-anchor': 'end' }); t.textContent = v.toFixed(2); g.push(t); }
  g.push(mk('line', { x1: X(1), x2: X(1), y1: P.t, y2: H - P.b, class: 'axis1' }), mk('line', { x1: P.l, x2: W - P.r, y1: Y(1), y2: Y(1), class: 'axis1' }));
  const lab = (x, y, s, a, cls) => { const t = mk('text', { x, y, 'text-anchor': a || 'start', class: cls || 'axis-label' }); t.textContent = s; return t; };
  g.push(lab((P.l + W - P.r) / 2, H - 8, '영향력계수 (후방연관) →', 'middle'));
  const yl = lab(0, 0, '감응도계수 (전방연관, 로그 눈금) →', 'middle'); yl.setAttribute('transform', `translate(16 ${(P.t + H - P.b) / 2}) rotate(-90)`); g.push(yl);
  g.push(lab(W - P.r - 6, P.t + 14, 'I 핵심', 'end', 'qlab'), lab(P.l + 6, P.t + 14, 'IV 중간재 공급형', 'start', 'qlab'),
    lab(P.l + 6, H - P.b - 8, 'III 독립형', 'start', 'qlab'), lab(W - P.r - 6, H - P.b - 8, 'II 원자재 수요형', 'end', 'qlab'));
  const sorted = [...all].sort((a, b) => b.output - a.output);
  const labelSet = new Set(sorted.slice(0, all.length <= 33 ? 8 : 6).map(r => r.code));
  for (const r of sorted) {
    const c = mk('circle', { cx: X(r.BL), cy: Y(r.FL), r: 3 + 16 * Math.sqrt(r.output / maxO), fill: QCOLOR[r.quad], class: 'pt' + (keep.has(r.code) ? '' : ' faded') });
    const tt = mk('title', {}); tt.textContent = `${r.code} ${r.name}\n영향력 ${fmtC(r.BL)} · 감응도 ${fmtC(r.FL)}\n총산출 ${human(r.output)}`; c.append(tt);
    c.onmouseenter = () => { $('scatter-tip').textContent = `${r.code} ${r.name} — 영향력계수 ${fmtC(r.BL)}, 감응도계수 ${fmtC(r.FL)}, 총산출 ${human(r.output)} (${r.quadName})`; };
    g.push(c);
  }
  for (const r of sorted.filter(r => labelSet.has(r.code) && keep.has(r.code))) g.push(lab(X(r.BL) + 6, Y(r.FL) - 6, r.name, 'start', 'pt-label'));
  svg.replaceChildren(...g);
}

/* ---------------------------------------------------------------- 5. 가격파급 */
document.querySelectorAll('[data-ptab]').forEach(b => b.onclick = () => {
  document.querySelectorAll('[data-ptab]').forEach(x => x.classList.toggle('on', x === b));
  $('price-import').hidden = b.dataset.ptab !== 'import'; $('price-va').hidden = b.dataset.ptab !== 'va';
});
$('pi-search').oninput = () => options($('pi-code'), sectors().filter(s => match($('pi-search').value.trim(), s.code, s.name)));
const secName = c => c === 'ALL' ? '전체 산업' : c + ' ' + (sectors().find(s => s.code === c) || {}).name;
function renderPBasket() {
  $('pbasket').replaceChildren(...S.pbasket.map((s, i) => el('li', {},
    el('span', { text: (s.kind === 'import' ? '수입품 가격 · ' + secName(s.code) : VA_NAME[s.component] + ' · ' + secName(s.code)) + ' · ' + (s.pct > 0 ? '+' : '') + fmt(s.pct, 2) + '%' }),
    el('button', { type: 'button', class: 'mini', text: '삭제', onclick: () => { S.pbasket.splice(i, 1); $('price-results').hidden = true; renderPBasket(); } }))));
  $('price-status').textContent = S.pbasket.length ? S.pbasket.length + '개 충격 준비. "가격파급 계산"을 누르세요.' : '충격을 추가하세요.';
}
$('price-import').onsubmit = e => {
  e.preventDefault(); const pct = Number($('pi-pct').value);
  if (!$('pi-code').value || !Number.isFinite(pct) || pct === 0) { $('price-status').textContent = '수입품과 0이 아닌 상승률을 확인하세요.'; return; }
  S.pbasket.push({ kind: 'import', code: $('pi-code').value, pct }); $('price-results').hidden = true; renderPBasket();
};
$('price-va').onsubmit = e => {
  e.preventDefault(); const pct = Number($('pv-pct').value);
  if (!Number.isFinite(pct) || pct === 0) { $('price-status').textContent = '0이 아닌 상승률을 입력하세요.'; return; }
  S.pbasket.push({ kind: 'va', component: $('pv-comp').value, code: $('pv-code').value, pct }); $('price-results').hidden = true; renderPBasket();
};
$('price-clear').onclick = () => { S.pbasket = []; $('price-results').hidden = true; renderPBasket(); };
$('price-run').onclick = async () => {
  if (!S.pbasket.length) { $('price-status').textContent = '충격을 먼저 추가하세요.'; return; }
  const snap = JSON.stringify(S.pbasket), level = S.active;
  $('price-run').disabled = true; $('price-status').textContent = '계산 중…';
  try {
    const r = await api('price', level, snap);
    if (snap !== JSON.stringify(S.pbasket) || level !== S.active) { $('price-status').textContent = '입력이 바뀌었습니다. 다시 계산하세요.'; return; }
    S.price = r; $('price-results').hidden = false;
    cards($('price-cards'), [
      ['소비자물가 근사 상승률', fmt(r.cpi, 3) + '%', '민간소비지출 가중 (국산+수입)'],
      ['생산자물가 근사 상승률', fmt(r.ppi, 3) + '%', '총산출 가중 평균'],
      ['가장 크게 오르는 산업', r.rows[0].name, fmt(r.rows[0].total, 3) + '%'],
    ]);
    bars($('price-bars'), r.rows.slice(0, 15).map(x => ({ label: x.code + ' ' + x.name, value: x.total })), v => fmt(v, 3) + '%');
    $('price-rows').replaceChildren(...r.rows.map(x => el('tr', {}, el('td', { text: x.code }), el('td', { text: x.name }),
      el('td', { class: 'r', text: fmt(x.direct, 4) }), el('td', { class: 'r', text: fmt(x.total, 4) }))));
    $('price-status').textContent = '계산 완료 (' + LEVELS.find(x => x.key === level).label + ' 기준). 단위는 % 입니다.';
  } catch (e) { $('price-status').textContent = '계산 실패: ' + e.message; }
  finally { $('price-run').disabled = false; }
};

/* ---------------------------------------------------------------- 6. 분류 수준 비교 */
function parseCSV(text) {
  const rows = []; let row = [], f = '', q = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (q) { if (c === '"') { if (text[i + 1] === '"') { f += '"'; i++; } else q = false; } else f += c; }
    else if (c === '"') q = true; else if (c === ',') { row.push(f); f = ''; }
    else if (c === '\n' || c === '\r') { if (c === '\r' && text[i + 1] === '\n') i++; row.push(f); rows.push(row); row = []; f = ''; }
    else f += c;
  }
  if (f || row.length) { row.push(f); rows.push(row); }
  return rows;
}
async function loadClasses() {
  try {
    const res = await fetch('./data/classification.csv'); if (!res.ok) throw Error('HTTP ' + res.status);
    const rows = parseCSV((await res.text()).replace(/^﻿/, '')).slice(1).filter(r => r.length >= 8);
    S.classes = rows.map(r => ({ code: r[0], name: r[1], small: r[2] + ' ' + r[3], medium: r[4] + ' ' + r[5], large: r[6] + ' ' + r[7] }));
    filterCmp(); const ic = S.classes.find(c => c.code === '3102'); if (ic) $('cmp-code').value = '3102';
  } catch (e) { $('cmp-status').textContent = '부문분류 대응표를 불러오지 못했습니다: ' + e.message; }
}
function filterCmp() {
  const q = $('cmp-search').value.trim();
  $('cmp-code').replaceChildren(...S.classes.filter(c => match(q, c.code, c.name, c.small, c.medium, c.large))
    .map(c => el('option', { value: c.code, text: c.code + ' ' + c.name + '  ⟵ ' + c.large.slice(3) })));
}
$('cmp-search').oninput = filterCmp;
function updateCompareHint() {
  const n = Object.keys(S.loaded).length;
  $('cmp-status').textContent = n >= 2 ? n + '개 분류를 불러왔습니다. 기본부문과 금액을 고르고 비교하세요.' : '한 개 분류만 불러왔습니다. 두 개 이상 불러오면 차이를 비교할 수 있습니다.';
}
$('cmp-run').onclick = async () => {
  const code = $('cmp-code').value, amount = Number($('cmp-amount').value);
  if (!code || !Number.isFinite(amount) || amount === 0) { $('cmp-status').textContent = '기본부문과 0이 아닌 금액을 확인하세요.'; return; }
  try {
    const r = await api('compare', code, amount);
    const base = r.rows[0];
    $('cmp-rows').replaceChildren(...r.rows.map(x => el('tr', {}, el('td', { text: x.label }), el('td', { text: x.sectors }), el('td', { text: x.code + ' ' + x.name }),
      el('td', { class: 'r', text: fmt(x.production, 2) + (x === base ? '' : ' (' + (x.production >= base.production ? '+' : '') + fmt((x.production / base.production - 1) * 100, 1) + '%)') }),
      el('td', { class: 'r', text: fmt(x.valueAdded, 2) }), el('td', { class: 'r', text: fmt(x.imports, 2) }),
      el('td', { class: 'r', text: fmtC(x.BL) }), el('td', { class: 'r', text: fmtC(x.FL) }))));
    $('cmp-status').textContent = r.rows.length > 1 ? '괄호 안은 가장 세밀한 분류(' + base.label + ') 대비 생산유발 차이입니다.' : '비교하려면 다른 분류도 불러오세요.';
  } catch (e) { $('cmp-status').textContent = '비교 실패: ' + e.message; }
};

/* ---------------------------------------------------------------- 시작 */
startWorker(); renderLevels(); loadClasses();
