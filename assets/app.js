/* Subsidios habitacionales otorgados – dashboard CChC O'Higgins
   Datos: data/subsidios.json (generado por etl/build_data.py desde el Excel del MINVU) */
'use strict';

/* ================= Constantes ================= */
const SEL = '#E4572E';      // región seleccionada (teja)
const PAIS = '#1E88D6';     // total país / contexto (mar)
const NAVY = '#16324F';     // tinta
const CATS = [
  { key: 'SECTORES VULNERABLES', name: 'Sectores vulnerables', color: '#1E88D6' },
  { key: 'SECTORES MEDIOS Y EMERGENTES', name: 'Sectores medios y emergentes', color: '#2FA56A' },
  { key: 'MEJORAMIENTO Y MANTENCIÓN DE VIVIENDA', name: 'Mejoramiento y mantención', color: '#7C5CE0' },
  { key: 'ARRIENDO', name: 'Arriendo', color: '#D08A00' },
];
const CMP_COLORS = ['#E4572E', '#1E88D6', '#2FA56A', '#7C5CE0'];
const SCOPES = { T: 'Total', R: 'Programa regular', X: 'Reconstrucción' };
const REG_LABEL = {
  'País': 'el país', "O'Higgins": "la Región de O'Higgins", 'Maule': 'la Región del Maule', 'Biobío': 'la Región del Biobío',
  'Metropolitana': 'la Región Metropolitana',
};

let D = null;
const S = { pA: 2025, pB: 2024, pAll: false, region: "O'Higgins", scope: 'T', metric: 'u', from: 2015, to: 2026, cmp: [], cmpMode: 'idx' };
const TABLES = {};
let firstRender = true;

/* ================= Utilidades ================= */
const $ = (s, el = document) => el.querySelector(s);
const $$ = (s, el = document) => [...el.querySelectorAll(s)];
const fmt = n => Math.round(n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');
const fmtD = (n, d = 1) => {
  const s = Math.abs(n).toFixed(d); const [i, f] = s.split('.');
  return (n < 0 ? '−' : '') + i.replace(/\B(?=(\d{3})+(?!\d))/g, '.') + (f ? ',' + f : '');
};
const pct = (n, d = 1) => fmtD(n, d) + '%';
const esc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const slug = s => s.normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/['’]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '');
const fmtUF = v => v >= 1e6 ? fmtD(v / 1e6, v >= 1e7 ? 1 : 2) + ' MM UF' : fmt(v) + ' UF';
const fmtMM = (v, d) => fmtD(v / 1e6, d);

const M = {
  u: { name: 'Subsidios', short: 'subsidios', tip: v => fmt(v) + ' subsidios' },
  uf: { name: 'Monto', short: 'monto', tip: v => fmtUF(v) },
};
const vfmt = (v, m = S.metric) => M[m].tip(v);

function axisFmt(m) {
  return (v, step) => {
    if (m === 'uf') return fmtD(v / 1e6, step >= 1e6 ? 0 : step >= 1e5 ? 1 : step >= 1e4 ? 2 : 3);
    if (v < 10000) return fmt(v);
    if (v < 1e6) return fmt(v / 1000) + ' mil';
    return fmtD(v / 1e6, 1) + ' M';
  };
}
function niceTicks(max, target = 4) {
  if (!(max > 0)) return { max: 1, step: 1, ticks: [0, 1] };
  const raw = max / target, mag = Math.pow(10, Math.floor(Math.log10(raw))), n = raw / mag;
  const step = (n <= 1 ? 1 : n <= 2 ? 2 : n <= 2.5 ? 2.5 : n <= 5 ? 5 : 10) * mag;
  const top = Math.ceil(max / step - 1e-9) * step; const ticks = [];
  for (let v = 0; v <= top + step * 1e-6; v += step) ticks.push(v);
  return { max: top, step, ticks };
}

/* ================= Acceso a datos ================= */
const ANIOS = () => D.meta.anios;
const ix = y => y - ANIOS()[0];
const LAST = () => ANIOS()[ANIOS().length - 1];
const FULLY = () => LAST() - 1;
const sec = r => (r === 'País' ? D.pais : D.reg[r]);
const TOT = (r, sc = S.scope) => sec(r)[sc].total;
const PROG = (r, sc = S.scope) => sec(r)[sc].prog;
const val = (r, y, m = S.metric, sc = S.scope) => TOT(r, sc)[m][ix(y)] || 0;
const sumR = (arr, a = S.from, b = S.to) => { let s = 0; for (let y = a; y <= b; y++) s += arr[ix(y)] || 0; return s; };
const yrs = (a = S.from, b = S.to) => { const o = []; for (let y = a; y <= b; y++) o.push(y); return o; };
const catOf = name => (D._cat[name] || '');
const catIdx = name => CATS.findIndex(c => c.key === catOf(name));
const regLabel = r => REG_LABEL[r] || (r === 'País' ? 'el país' : 'la Región de ' + r);
const regName = r => (r === 'País' ? 'Total país' : r);
const CORTE = () => D.meta.corte_nombre || 'octubre';
const CABR = () => CORTE().slice(0, 3);
const partial = y => y === LAST();

function short(n) {
  return n
    .replace(/Sistema Integrado de Subsidios\s*DS 01\s*/, 'Integrado DS 01 · ')
    .replace(/Fondo Solidario de Elección de Vivienda \(FSEV\)\s*/, 'FSEV ')
    .replace(/Fondo Solidario de Vivienda \(FSV\)\s*/, 'FSV ')
    .replace(/Programa Integración Social y Territorial\s*/, 'Integración Social y Territorial ')
    .replace(/Programa Extraordinario de Reactivación e Integración Social\s*/, 'Reactivación e Integración Social ')
    .replace(/Protección Patrimonio Familiar \(PPPF\)\s*/, 'PPPF ')
    .replace(/Programa de Mejoramiento Viviendas y Barrios \(DS 27\)/, 'Mejoramiento Viv. y Barrios DS 27')
    .replace(/Subsidio Habitacional Extraordinario para Adquisición de Vivienda Económica DS 4/, 'Adquisición Vivienda Económica DS 4')
    .replace(/Habitabilidad Rural DS 10 /, 'Habitabilidad Rural DS 10 · ');
}

/* ================= Tooltip ================= */
const tip = $('#tip');
function showTip(html, x, y) {
  tip.innerHTML = html; tip.style.display = 'block';
  const w = tip.offsetWidth, h = tip.offsetHeight;
  let left = x + 16; if (left + w > innerWidth - 8) left = x - w - 16; if (left < 8) left = 8;
  let top = y - h - 12; if (top < 8) top = y + 18;
  tip.style.left = left + 'px'; tip.style.top = top + 'px';
}
const hideTip = () => { tip.style.display = 'none'; };
function bindHover(el, fn) {
  el.onpointermove = e => {
    const h = e.target.closest('[data-i]');
    if (!h || !el.contains(h)) { hideTip(); return; }
    $$('.on', el).forEach(n => n.classList.remove('on')); h.classList.add('on');
    const html = fn(h);
    if (html) showTip(html, e.clientX, e.clientY);
  };
  el.onpointerleave = () => { hideTip(); $$('.on', el).forEach(n => n.classList.remove('on')); el.dispatchEvent(new CustomEvent('hoverend')); };
}
const dChg = (cur, prev) => (prev > 0 ? (cur / prev - 1) * 100 : null);
const dTxt = d => (d === null ? '' : (d >= 0 ? '+' : '') + fmtD(d, 1) + '% vs año anterior');

/* ================= Tablas (vista alternativa de cada gráfico) ================= */
function setTable(cardId, data) {
  TABLES[cardId] = data;
  const card = $('#' + cardId);
  if (card && card.dataset.mode === 'table') drawTable(card);
}
function drawTable(card) {
  const data = TABLES[card.id]; const wrap = $('.tablewrap', card);
  if (!data) { wrap.innerHTML = ''; return; }
  wrap.innerHTML = '<div class="scrolly"><table class="t"><thead><tr>' + data.head.map(h => `<th scope="col">${esc(h)}</th>`).join('') +
    '</tr></thead><tbody>' + data.rows.map(r => '<tr>' + r.map(c => `<td>${esc(c)}</td>`).join('') + '</tr>').join('') + '</tbody></table></div>';
}
function initToggles() {
  $$('[data-toggle]').forEach(btn => {
    btn.addEventListener('click', () => {
      const card = btn.closest('.card'); const tbl = card.dataset.mode === 'table';
      card.dataset.mode = tbl ? 'chart' : 'table';
      $$('.chart', card).forEach(c => (c.hidden = !tbl));
      $$('.legend', card).forEach(c => (c.hidden = !tbl));
      $('.tablewrap', card).hidden = tbl;
      btn.textContent = tbl ? 'Ver tabla' : 'Ver gráfico';
      if (!tbl) drawTable(card);
    });
  });
}

/* ================= Gráficos ================= */
function chartBox(el, hDesk = 300, hMob = 250) {
  const W = Math.max(280, Math.floor(el.clientWidth || 600));
  return { W, H: W < 520 ? hMob : hDesk };
}
function yearTicks(n, iw) { const maxLabels = Math.max(2, Math.floor(iw / 34)); return Math.max(1, Math.ceil(n / maxLabels)); }
function colPath(x, y, w, h, r) {
  r = Math.min(r, w / 2, h);
  return `M${x},${y + h}V${y + r}Q${x},${y} ${x + r},${y}H${x + w - r}Q${x + w},${y} ${x + w},${y + r}V${y + h}Z`;
}
function gridAxes(m, W, H, tk, fmtTick, unitLabel) {
  const ih = H - m.t - m.b; let s = '<g class="grid">';
  tk.ticks.forEach(v => {
    const y = m.t + ih - (v / tk.max) * ih;
    s += `<line x1="${m.l}" x2="${W - m.r}" y1="${y}" y2="${y}" stroke-width="1"/>`;
    s += `<text x="${m.l - 8}" y="${y + 4}" text-anchor="end">${fmtTick(v, tk.step)}</text>`;
  });
  s += '</g>';
  if (unitLabel) s += `<text x="${m.l - 8}" y="${m.t - 9}" text-anchor="end">${unitLabel}</text>`;
  return s;
}
function xLabels(m, W, H, years, slotX, step) {
  let s = '';
  years.forEach((y, i) => { if (i % step === 0 || i === years.length - 1 && (years.length - 1) % step === 0) s += `<text x="${slotX(i)}" y="${H - 8}" text-anchor="middle">${y}${partial(y) ? '*' : ''}</text>`; });
  return s;
}

/* Columnas simples (un valor por año) */
function columns(el, o) {
  const { W, H } = chartBox(el, o.h || 300, 250);
  const m = { l: o.kind === 'uf' ? 50 : 52, r: 10, t: 26, b: 28 };
  const iw = W - m.l - m.r, ih = H - m.t - m.b, n = o.years.length;
  const slot = iw / n, bw = Math.min(24, slot * 0.72);
  const tk = niceTicks(Math.max(...o.vals, 0));
  const uid = 'h' + Math.random().toString(36).slice(2, 7);
  const sx = i => m.l + slot * i + slot / 2;
  const yv = v => m.t + ih - (v / tk.max) * ih;
  let s = `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="${esc(o.aria)}"><defs><pattern id="${uid}" width="6" height="6" patternUnits="userSpaceOnUse" patternTransform="rotate(45)"><rect width="6" height="6" fill="${o.color}" fill-opacity=".28"/><rect width="2.5" height="6" fill="${o.color}"/></pattern></defs>`;
  s += gridAxes(m, W, H, tk, axisFmt(o.kind), o.kind === 'uf' ? 'MM UF' : '');
  o.years.forEach((y, i) => {
    const h = Math.max(0, (o.vals[i] / tk.max) * ih);
    if (h > 0) s += `<path class="col" style="animation-delay:${Math.min(i * 14, 520)}ms" d="${colPath(sx(i) - bw / 2, yv(o.vals[i]), bw, h, 4)}" fill="${partial(y) ? `url(#${uid})` : o.color}"/>`;
  });
  s += `<line class="axisline" x1="${m.l}" x2="${W - m.r}" y1="${m.t + ih}" y2="${m.t + ih}"/>`;
  s += xLabels(m, W, H, o.years, sx, yearTicks(n, iw));
  // etiquetas directas: máximo y último año completo
  const lab = new Set();
  let mx = 0; o.vals.forEach((v, i) => { if (v > o.vals[mx] && !partial(o.years[i])) mx = i; });
  lab.add(mx);
  const lastFull = o.years.map((y, i) => i).filter(i => !partial(o.years[i])).pop();
  if (lastFull !== undefined && Math.abs(sx(lastFull) - sx(mx)) > 46) lab.add(lastFull);
  lab.forEach(i => { if (o.vals[i] > 0) s += `<text class="vlab" x="${sx(i)}" y="${yv(o.vals[i]) - 6}" text-anchor="middle">${o.lab(o.vals[i])}</text>`; });
  o.years.forEach((y, i) => { s += `<rect class="hit" data-i="${i}" x="${m.l + slot * i}" y="${m.t}" width="${slot}" height="${ih + 4}"/>`; });
  s += '</svg>';
  el.innerHTML = s;
  bindHover(el, h => {
    const i = +h.dataset.i, y = o.years[i], p = i > 0 ? o.vals[i - 1] : null;
    const d = p ? dChg(o.vals[i], p) : null;
    return `<b>${y}${partial(y) ? ' (ene–' + CABR() + ')' : ''}</b><div>${o.tipv(o.vals[i])}</div>${d !== null && !partial(y) ? `<div class="m">${dTxt(d)}</div>` : ''}`;
  });
}

/* Columnas apiladas con separación de 2px */
function stacked(el, o) {
  const { W, H } = chartBox(el, 320, 270);
  const m = { l: o.kind === 'uf' ? 50 : 52, r: 10, t: 22, b: 28 };
  const iw = W - m.l - m.r, ih = H - m.t - m.b, n = o.years.length;
  const slot = iw / n, bw = Math.min(24, slot * 0.72);
  const totals = o.years.map((_, i) => o.series.reduce((a, s) => a + s.vals[i], 0));
  const tk = niceTicks(Math.max(...totals, 0));
  const sx = i => m.l + slot * i + slot / 2;
  let s = `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="${esc(o.aria)}">`;
  s += gridAxes(m, W, H, tk, axisFmt(o.kind), o.kind === 'uf' ? 'MM UF' : '');
  o.years.forEach((y, i) => {
    let acc = 0; const tops = o.series.map((se, k) => (se.vals[i] > 0 ? k : -1)).filter(k => k >= 0); const topK = tops.length ? tops[tops.length - 1] : -1;
    o.series.forEach((se, k) => {
      const v = se.vals[i]; if (v <= 0) return;
      const h = (v / tk.max) * ih, yb = m.t + ih - (acc / tk.max) * ih; acc += v;
      const gap = k === tops[0] ? 0 : 2; const hh = Math.max(1, h - gap);
      s += `<path class="col" style="animation-delay:${Math.min(i * 14, 520)}ms" d="${colPath(sx(i) - bw / 2, yb - h, bw, hh, k === topK ? 4 : 0)}" fill="${se.color}" opacity="${partial(y) ? .62 : 1}"/>`;
    });
  });
  s += `<line class="axisline" x1="${m.l}" x2="${W - m.r}" y1="${m.t + ih}" y2="${m.t + ih}"/>`;
  s += xLabels(m, W, H, o.years, sx, yearTicks(n, iw));
  o.years.forEach((y, i) => { s += `<rect class="hit" data-i="${i}" x="${m.l + slot * i}" y="${m.t}" width="${slot}" height="${ih + 4}"/>`; });
  s += '</svg>';
  el.innerHTML = s;
  bindHover(el, h => {
    const i = +h.dataset.i, y = o.years[i], t = totals[i];
    let html = `<b>${y}${partial(y) ? ' (ene–' + CABR() + ')' : ''}</b>`;
    [...o.series].reverse().forEach(se => { if (se.vals[i] > 0) html += `<div class="r"><span><i style="background:${se.color}"></i>${esc(se.name)}</span><span>${o.tipv(se.vals[i])}${t ? ' · ' + pct(se.vals[i] / t * 100, 0) : ''}</span></div>`; });
    html += `<div class="r m"><span>Total</span><span>${o.tipv(t)}</span></div>`;
    return html;
  });
}

/* Líneas con cruce de ejes en hover */
function lines(el, o) {
  const { W, H } = chartBox(el, o.h || 300, 250);
  const m = { l: o.mleft || 52, r: 14, t: 22, b: 28 };
  const iw = W - m.l - m.r, ih = H - m.t - m.b, n = o.years.length;
  const slot = iw / n; const sx = i => m.l + slot * i + slot / 2;
  const all = o.series.flatMap(se => se.vals.filter(v => v !== null));
  const tk = niceTicks(Math.max(...all, 0)); const yv = v => m.t + ih - (v / tk.max) * ih;
  let s = `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="${esc(o.aria)}">`;
  s += gridAxes(m, W, H, tk, o.axis, o.unit || '');
  s += `<line class="axisline" x1="${m.l}" x2="${W - m.r}" y1="${m.t + ih}" y2="${m.t + ih}"/>`;
  s += xLabels(m, W, H, o.years, sx, yearTicks(n, iw));
  o.series.forEach(se => {
    let solid = '', dash = ''; let started = false;
    se.vals.forEach((v, i) => {
      if (v === null) { started = false; return; }
      const pt = `${sx(i).toFixed(1)},${yv(v).toFixed(1)}`;
      const isPartialLink = partial(o.years[i]) && i > 0 && se.vals[i - 1] !== null;
      if (isPartialLink) { dash = `M${sx(i - 1).toFixed(1)},${yv(se.vals[i - 1]).toFixed(1)}L${pt}`; }
      else { solid += (started ? 'L' : 'M') + pt; started = true; }
    });
    s += `<path class="ln" pathLength="1" d="${solid}" fill="none" stroke="${se.color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>`;
    if (dash) s += `<path d="${dash}" fill="none" stroke="${se.color}" stroke-width="2" stroke-linecap="round" stroke-dasharray="4 4"/>`;
    // punto final (último año completo)
    let li = se.vals.length - 1; while (li >= 0 && (se.vals[li] === null || partial(o.years[li]))) li--;
    if (li >= 0) s += `<circle cx="${sx(li)}" cy="${yv(se.vals[li])}" r="4.5" fill="${se.color}" stroke="#fff" stroke-width="2"/>`;
  });
  s += `<line id="xh" x1="0" x2="0" y1="${m.t}" y2="${m.t + ih}" stroke="${NAVY}" stroke-opacity=".35" visibility="hidden"/>`;
  s += '<g id="dots"></g>';
  o.years.forEach((y, i) => { s += `<rect class="hit" style="fill:transparent" data-i="${i}" x="${m.l + slot * i}" y="${m.t}" width="${slot}" height="${ih + 4}"/>`; });
  s += '</svg>';
  el.innerHTML = s;
  const xh = $('#xh', el), dots = $('#dots', el);
  el.addEventListener('hoverend', () => { xh.setAttribute('visibility', 'hidden'); dots.innerHTML = ''; });
  bindHover(el, h => {
    const i = +h.dataset.i, y = o.years[i];
    xh.setAttribute('x1', sx(i)); xh.setAttribute('x2', sx(i)); xh.setAttribute('visibility', 'visible');
    dots.innerHTML = o.series.filter(se => se.vals[i] !== null).map(se => `<circle cx="${sx(i)}" cy="${yv(se.vals[i])}" r="5" fill="${se.color}" stroke="#fff" stroke-width="2"/>`).join('');
    const rows = o.series.filter(se => se.vals[i] !== null).sort((a, b) => b.vals[i] - a.vals[i]);
    return `<b>${y}${partial(y) ? ' (ene–' + CABR() + ')' : ''}</b>` + rows.map(se => `<div class="r"><span><i style="background:${se.color}"></i>${esc(se.name)}</span><span>${o.tipv(se.vals[i], se, i)}</span></div>`).join('');
  });
}

/* Barras horizontales (ranking) */
function hbars(el, o) {
  const rowH = 24, m = { l: 118, r: 92, t: 4, b: 4 };
  const W = Math.max(300, Math.floor(el.clientWidth || 500)); const H = m.t + m.b + rowH * o.items.length;
  const max = Math.max(...o.items.map(d => d.v), 1); const iw = W - m.l - m.r;
  let s = `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="${esc(o.aria)}">`;
  o.items.forEach((d, i) => {
    const y = m.t + i * rowH, w = Math.max(2, (d.v / max) * iw), sel = d.sel;
    s += `<g class="clickable" data-i="${i}" data-name="${esc(d.name)}" tabindex="0" role="button" aria-label="${esc(d.name)}: ${esc(o.lab(d.v))}. Seleccionar región">`;
    s += `<rect class="hit" x="0" y="${y}" width="${W}" height="${rowH}"/>`;
    s += `<text x="${m.l - 8}" y="${y + rowH / 2 + 4}" text-anchor="end" style="${sel ? 'fill:#16324F;font-weight:700' : ''}">${esc(d.name)}</text>`;
    s += `<path class="hb" style="animation-delay:${i * 28}ms" d="M${m.l},${y + 4}H${m.l + w - 4}Q${m.l + w},${y + 4} ${m.l + w},${y + 8}V${y + rowH - 8}Q${m.l + w},${y + rowH - 4} ${m.l + w - 4},${y + rowH - 4}H${m.l}Z" fill="${sel ? SEL : '#7DB8EA'}"/>`;
    s += `<text class="vlab" x="${m.l + w + 6}" y="${y + rowH / 2 + 4}">${o.lab(d.v)}</text></g>`;
  });
  s += '</svg>';
  el.innerHTML = s;
  bindHover(el, h => o.tip(+h.dataset.i));
  $$('[data-name]', el).forEach(g => {
    const go = () => { selectRegion(g.dataset.name); };
    g.addEventListener('click', go);
    g.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); go(); } });
  });
}

/* Barras 100% apiladas por fila (distribución por grupo) */
function share100(el, o) {
  const rowH = 26, m = { l: 118, r: 8, t: 4, b: 4 };
  const W = Math.max(300, Math.floor(el.clientWidth || 500)); const H = m.t + m.b + rowH * o.rows.length;
  const iw = W - m.l - m.r;
  let s = `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="${esc(o.aria)}">`;
  o.rows.forEach((r, i) => {
    const y = m.t + i * rowH; const t = r.vals.reduce((a, b) => a + b, 0);
    s += `<g data-i="${i}"><rect class="hit" x="0" y="${y}" width="${W}" height="${rowH}"/>`;
    s += `<text x="${m.l - 8}" y="${y + rowH / 2 + 4}" text-anchor="end" style="${r.sel ? 'fill:#16324F;font-weight:700' : ''}">${esc(r.name)}</text>`;
    if (r.sel) s += `<rect x="${m.l - 112}" y="${y + 6}" width="3" height="${rowH - 12}" rx="1.5" fill="${SEL}"/>`;
    let x = m.l;
    r.vals.forEach((v, k) => {
      if (v <= 0 || t <= 0) return; const w = v / t * iw; const gap = x > m.l ? 2 : 0;
      s += `<rect x="${x + gap}" y="${y + 5}" width="${Math.max(1, w - gap)}" height="${rowH - 10}" rx="3" fill="${CATS[k].color}"/>`;
      if (w > 36) s += `<text x="${x + gap + 6}" y="${y + rowH / 2 + 4}" style="fill:#fff;font-weight:600;font-size:11px">${fmtD(v / t * 100, 0)}%</text>`;
      x += w;
    });
    s += '</g>';
  });
  s += '</svg>';
  el.innerHTML = s;
  bindHover(el, h => {
    const r = o.rows[+h.dataset.i]; const t = r.vals.reduce((a, b) => a + b, 0);
    return `<b>${esc(r.name)}</b>` + CATS.map((c, k) => `<div class="r"><span><i style="background:${c.color}"></i>${esc(c.name)}</span><span>${o.tipv(r.vals[k])}${t ? ' · ' + pct(r.vals[k] / t * 100, 1) : ''}</span></div>`).join('');
  });
}

/* Mapa de calor programas × años */
function heatmap(el, o) {
  const labW = 250, rowH = 22, mt = 20; const n = o.years.length;
  const W = Math.max(el.clientWidth || 600, labW + n * 16 + 10); const cw = (W - labW - 8) / n;
  const H = mt + o.rows.length * rowH + 4; const max = Math.max(...o.rows.flatMap(r => r.vals), 1);
  const col = v => { if (v <= 0) return '#EEF4F8'; const t = Math.sqrt(v / max); const a = [232, 244, 252], b = [22, 50, 79]; return `rgb(${a.map((x, i) => Math.round(x + (b[i] - x) * t)).join(',')})`; };
  let s = `<svg viewBox="0 0 ${W} ${H}" style="min-width:${W}px" role="img" aria-label="${esc(o.aria)}">`;
  const step = Math.max(1, Math.ceil(n / Math.floor((W - labW) / 34)));
  o.years.forEach((y, c) => { if (c % step === 0) s += `<text x="${labW + c * cw + cw / 2}" y="12" text-anchor="middle">${y}${partial(y) ? '*' : ''}</text>`; });
  o.rows.forEach((r, i) => {
    const y = mt + i * rowH; const nm = short(r.n);
    s += `<text x="${labW - 8}" y="${y + rowH / 2 + 4}" text-anchor="end" style="fill:#16324F">${esc(nm.length > 40 ? nm.slice(0, 39) + '…' : nm)}</text>`;
    r.vals.forEach((v, c) => { s += `<rect class="hit" data-i="${i}" data-c="${c}" x="${labW + c * cw + 1}" y="${y + 1}" width="${cw - 2}" height="${rowH - 2}" rx="3" style="fill:${col(v)}"/>`; });
  });
  s += '</svg>';
  el.innerHTML = s;
  $$('.hit', el).forEach(r => { r.style.cursor = 'default'; });
  bindHover(el, h => {
    const r = o.rows[+h.dataset.i], c = +h.dataset.c; const y = o.years[c];
    return `<b>${esc(short(r.n))}</b><div>${y}${partial(y) ? ' (ene–' + CABR() + ')' : ''}: ${r.vals[c] > 0 ? o.tipv(r.vals[c]) : 'sin subsidios'}</div><div class="m">${esc(catName(r.n))}</div>`;
  });
}
const catName = n => { const c = CATS.find(c => c.key === catOf(n)); return c ? c.name : ''; };

/* ================= Cálculos reutilizables ================= */
function regionTotals(sc = S.scope, m = S.metric, a = S.from, b = S.to) {
  return D.regiones.map(r => ({ name: r, v: sumR(TOT(r, sc)[m], a, b) }));
}
function catSeries(r, m = S.metric, sc = S.scope, years = yrs()) {
  const out = CATS.map(c => ({ name: c.name, color: c.color, vals: years.map(() => 0) }));
  Object.entries(PROG(r, sc)).forEach(([name, p]) => {
    const k = catIdx(name); if (k < 0) return;
    years.forEach((y, i) => { out[k].vals[i] += p[m][ix(y)] || 0; });
  });
  return out;
}
function rankOf(r, m = S.metric, sc = S.scope, a = S.from, b = S.to) {
  const arr = regionTotals(sc, m, a, b).sort((x, y) => y.v - x.v); return arr.findIndex(d => d.name === r) + 1;
}

/* ================= Render ================= */
function periodTxt() { return S.from === S.to ? String(S.from) : `${S.from}–${S.to}`; }
function scopeTxt() { return S.scope === 'T' ? '' : ` · ${SCOPES[S.scope]}`; }
const ICON = {
  house: '<path d="M3 11.5 12 4l9 7.5M5.5 10v9h13v-9M10 19v-5h4v5"/>',
  cal: '<rect x="4" y="5" width="16" height="15" rx="2"/><path d="M8 3v4M16 3v4M4 10h16"/>',
  clock: '<circle cx="12" cy="12" r="8.5"/><path d="M12 7.5V12l3 2"/>',
  pin: '<path d="M12 21s-6.5-5.4-6.5-10.5a6.5 6.5 0 0 1 13 0C18.5 15.6 12 21 12 21z"/><circle cx="12" cy="10.5" r="2.3"/>',
  coin: '<circle cx="12" cy="12" r="8.5"/><path d="M14.8 9.2c-.6-.8-1.6-1.2-2.8-1.2-1.6 0-2.8.8-2.8 2s1 1.7 2.8 2 2.8.8 2.8 2-1.2 2-2.8 2c-1.2 0-2.2-.4-2.8-1.2M12 6.5V8m0 8v1.5"/>',
  star: '<path d="m12 4 2.4 5 5.4.7-3.9 3.8.9 5.4L12 16.3 7.2 18.9l.9-5.4L4.2 9.7 9.6 9z"/>',
};
const ico = k => `<span class="ico"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${ICON[k]}</svg></span>`;
const deltaHTML = d => (d === null ? '' : `<span class="delta ${d >= 0 ? 'up' : 'down'}">${d >= 0 ? '▲' : '▼'} ${fmtD(Math.abs(d), 1)}%</span>`);

function renderLead() {
  const r = S.region, t = TOT(r, 'T'), f = FULLY(), v = t.u[ix(f)], p = t.u[ix(f - 1)];
  const d = dChg(v, p); const cmp = d === null ? '' : `, ${fmtD(Math.abs(d), 1)}% ${d < 0 ? 'menos' : 'más'} que en ${f - 1}`;
  let extra = '';
  if (r !== 'País') {
    const sh = v / TOT('País', 'T').u[ix(f)] * 100;
    const rk = D.regiones.map(x => ({ x, v: TOT(x, 'T').u[ix(f)] })).sort((a, b) => b.v - a.v).findIndex(o => o.x === r) + 1;
    extra = ` Equivale al ${pct(sh)} del país y la deja en el lugar ${rk} entre las 16 regiones.`;
  }
  const cur = t.u[ix(LAST())];
  $('#lead').innerHTML = `En ${f} se otorgaron <b>${fmt(v)} subsidios</b> en ${regLabel(r)}${cmp}.${extra} En lo que va de ${LAST()} (enero–${CORTE()}) van ${fmt(cur)}.`;
}

function renderKPIs(animate) {
  const r = S.region, m = S.metric, f = FULLY(), L = LAST(); const t = TOT(r)[m];
  const mv = v => (m === 'uf' ? fmtUF(v) : fmt(v));
  const total = sumR(t); const fullYears = yrs().filter(y => !partial(y)); const avg = fullYears.length ? sumR(t, fullYears[0], fullYears[fullYears.length - 1]) / fullYears.length : null;
  const k = [];
  k.push(`<div class="kpi hero-kpi">${ico('house')}<div class="k">${M[m].name === 'Monto' ? 'Monto otorgado' : 'Subsidios otorgados'}, ${periodTxt()}${scopeTxt()}</div><div class="v">${mv(total)}</div><div class="d">${avg !== null ? `${mv(avg)} por año en promedio (años completos)` : 'Período con solo el año en curso'}</div></div>`);
  const vf = t[ix(f)], vp = t[ix(f - 1)];
  k.push(`<div class="kpi">${ico('cal')}<div class="k">Último año completo (${f})</div><div class="v">${mv(vf)}</div><div class="d">${deltaHTML(dChg(vf, vp))} vs ${f - 1}</div></div>`);
  const vc = t[ix(L)];
  k.push(`<div class="kpi">${ico('clock')}<div class="k">${L}, enero–${CORTE()}</div><div class="v">${mv(vc)}</div><div class="d">${vf > 0 ? `${fmtD(vc / vf * 100, 0)}% de lo otorgado en todo ${f}` : ''}</div></div>`);
  if (r === 'País') {
    const top = regionTotals().sort((a, b) => b.v - a.v)[0]; const all = sumR(t);
    k.push(`<div class="kpi w4">${ico('pin')}<div class="k">Región con mayor participación</div><div class="v">${esc(top.name)}</div><div class="d">${pct(top.v / all * 100)} del total del período</div></div>`);
  } else {
    const sh = sumR(t) / sumR(TOT('País')[m]) * 100;
    k.push(`<div class="kpi w4">${ico('pin')}<div class="k">Participación en el país</div><div class="v">${pct(sh)}</div><div class="d">Lugar ${rankOf(r)} entre las 16 regiones</div></div>`);
  }
  const au = sumR(TOT(r).uf) / Math.max(1, sumR(TOT(r).u)); const ap = sumR(TOT('País').uf) / Math.max(1, sumR(TOT('País').u));
  k.push(`<div class="kpi w4">${ico('coin')}<div class="k">Monto promedio por subsidio</div><div class="v">${fmt(au)} UF</div><div class="d">${r === 'País' ? 'Promedio del período' : `País: ${fmt(ap)} UF`}</div></div>`);
  const pr = Object.entries(PROG(r)).map(([n, p]) => [n, sumR(p[m])]).sort((a, b) => b[1] - a[1])[0];
  k.push(`<div class="kpi w4">${ico('star')}<div class="k">Programa con mayor peso</div><div class="v" style="font-size:22px">${pr && pr[1] > 0 ? esc(short(pr[0])) : '—'}</div><div class="d">${pr && pr[1] > 0 ? pct(pr[1] / total * 100) + ' del período' : ''}</div></div>`);
  $('#kpis').innerHTML = k.join('');
  if (animate) $$('#kpis .v').forEach((el, i) => { if (i < 5) countUp(el); });
  $('#resumen-sub').textContent = `${regName(r)} · ${SCOPES[S.scope]} · ${periodTxt()}${S.to === L ? ` (${L} parcial: enero–${CORTE()})` : ''}`;
}

function renderAnnual() {
  const r = S.region, m = S.metric, years = yrs(), t = TOT(r)[m], vals = years.map(y => t[ix(y)]);
  $('#t-annual').textContent = `${M[m].name === 'Monto' ? 'Monto otorgado' : 'Subsidios otorgados'} por año · ${regName(r)}`;
  $('#s-annual').textContent = `${SCOPES[S.scope]}${m === 'uf' ? ' · en millones de UF' : ''} · el último año es parcial (ene–${CABR()})`;
  columns($('#ch-annual'), {
    years, vals, color: r === 'País' ? PAIS : SEL, kind: m,
    lab: v => (m === 'uf' ? fmtMM(v, 2) : fmt(v)), tipv: v => vfmt(v, m),
    aria: `Columnas de ${M[m].short} otorgados por año, ${regName(r)}, ${periodTxt()}`,
  });
  setTable('card-annual', { head: ['Año', m === 'uf' ? 'Monto (UF)' : 'Subsidios', 'Var. anual'], rows: years.map((y, i) => [y + (partial(y) ? '*' : ''), fmt(vals[i]), i > 0 && !partial(y) ? (dChg(vals[i], vals[i - 1]) === null ? '' : fmtD(dChg(vals[i], vals[i - 1]), 1) + '%') : '']) });

  // participación
  const sh = (rg, y) => { const pv = val('País', y, m); return pv > 0 ? val(rg, y, m) / pv * 100 : null; };
  let series;
  if (r === 'País') {
    const top = regionTotals().sort((a, b) => b.v - a.v).slice(0, 4).map(d => d.name);
    $('#s-share').textContent = 'Las cuatro regiones con más ' + (m === 'uf' ? 'monto' : 'subsidios') + ' en el período';
    series = top.map((n, k) => ({ name: n, color: CMP_COLORS[k], vals: years.map(y => sh(n, y)) }));
  } else {
    $('#s-share').textContent = `${regName(r)} sobre el total país · ${SCOPES[S.scope]}`;
    series = [{ name: regName(r), color: SEL, vals: years.map(y => sh(r, y)) }];
  }
  $('#lg-share') && ($('#lg-share').innerHTML = '');
  lines($('#ch-share'), { years, series, axis: v => fmtD(v, v < 10 ? 1 : 0) + '%', unit: '', aria: 'Participación en el total país por año', tipv: v => pct(v, 1), mleft: 44 });
  setTable('card-share', { head: ['Año', ...series.map(s => s.name)], rows: years.map((y, i) => [y + (partial(y) ? '*' : ''), ...series.map(s => (s.vals[i] === null ? '' : fmtD(s.vals[i], 1) + '%'))]) });
}

function renderCats() {
  const r = S.region, m = S.metric, years = yrs(); const series = catSeries(r);
  $('#t-cat').textContent = `${M[m].name === 'Monto' ? 'Monto' : 'Subsidios'} por grupo de programas · ${regName(r)}`;
  $('#s-cat').textContent = `${SCOPES[S.scope]}${m === 'uf' ? ' · en millones de UF' : ''}`;
  $('#lg-cat').innerHTML = CATS.map(c => `<span><i style="background:${c.color}"></i>${esc(c.name)}</span>`).join('');
  stacked($('#ch-cat'), { years, series, kind: m, tipv: v => vfmt(v, m), aria: `Columnas apiladas por grupo de programas, ${regName(r)}` });
  setTable('card-cat', { head: ['Año', ...CATS.map(c => c.name), 'Total'], rows: years.map((y, i) => [y + (partial(y) ? '*' : ''), ...series.map(s => fmt(s.vals[i])), fmt(series.reduce((a, s) => a + s.vals[i], 0))]) });
}

let progSort = { k: 'v', dir: -1 };
function renderProgTable() {
  const r = S.region, m = S.metric; const pg = PROG(r);
  let rows = Object.entries(pg).map(([n, p]) => ({ n, u: sumR(p.u), uf: sumR(p.uf) })).filter(x => x.u > 0 || x.uf > 0);
  const tu = rows.reduce((a, x) => a + x.u, 0), tuf = rows.reduce((a, x) => a + x.uf, 0);
  rows.forEach(x => { x.v = x[m]; x.avg = x.u ? x.uf / x.u : 0; x.share = (m === 'uf' ? tuf : tu) ? x.v / (m === 'uf' ? tuf : tu) * 100 : 0; });
  rows.sort((a, b) => (progSort.k === 'n' ? a.n.localeCompare(b.n) * progSort.dir : (a[progSort.k] - b[progSort.k]) * progSort.dir));
  $('#t-prog').textContent = `Detalle por programa · ${regName(r)}`;
  $('#s-prog').textContent = `${periodTxt()} · ${SCOPES[S.scope]} · haz clic en un encabezado para ordenar`;
  const mx = Math.max(...rows.map(x => x.share), 1);
  const th = (k, t, s = true) => `<th scope="col" ${s ? `class="sortable" data-k="${k}" aria-sort="${progSort.k === k ? (progSort.dir < 0 ? 'descending' : 'ascending') : 'none'}"` : ''}>${t}${progSort.k === k ? (progSort.dir < 0 ? ' ▾' : ' ▴') : ''}</th>`;
  let h = `<table class="t"><thead><tr>${th('n', 'Programa')}${th('u', 'Subsidios')}${th('uf', 'Monto (MM UF)')}${th('share', '% del período')}${th('avg', 'UF por subsidio')}</tr></thead><tbody>`;
  rows.forEach(x => {
    const c = CATS.find(c => c.key === catOf(x.n));
    h += `<tr><td><div class="pname"><i style="background:${c ? c.color : '#999'}"></i><span title="${esc(x.n)}">${esc(short(x.n))}</span></div></td><td>${fmt(x.u)}</td><td>${fmtMM(x.uf, 2)}</td><td><span class="minibar" style="width:${Math.max(2, x.share / mx * 70)}px"></span> ${pct(x.share)}</td><td>${x.u ? fmt(x.avg) : '—'}</td></tr>`;
  });
  h += `<tr style="font-weight:700"><td>Total</td><td>${fmt(tu)}</td><td>${fmtMM(tuf, 2)}</td><td>100%</td><td>${tu ? fmt(tuf / tu) : '—'}</td></tr></tbody></table>`;
  $('#tb-prog').innerHTML = rows.length ? h : '<p class="note">No hay subsidios registrados en este alcance y período.</p>';
  $$('#tb-prog th.sortable').forEach(el => el.addEventListener('click', () => {
    const k = el.dataset.k; progSort = progSort.k === k ? { k, dir: -progSort.dir } : { k, dir: k === 'n' ? 1 : -1 }; renderProgTable();
  }));
}

function renderHeat() {
  const r = S.region, m = S.metric, years = yrs(); const pg = PROG(r);
  const order = D.programas.map(p => p.nombre);
  let rows = Object.entries(pg).map(([n, p]) => ({ n, vals: years.map(y => p[m][ix(y)] || 0), first: p[m].findIndex(v => v > 0) })).filter(x => x.vals.some(v => v > 0));
  rows.sort((a, b) => a.first - b.first || order.indexOf(a.n) - order.indexOf(b.n));
  $('#s-heat').textContent = `${regName(r)} · ${SCOPES[S.scope]} · más oscuro = más ${m === 'uf' ? 'monto' : 'subsidios'} · ordenado por año de inicio`;
  if (!rows.length) { $('#ch-heat').innerHTML = '<p class="note">Sin subsidios registrados en este alcance y período.</p>'; setTable('card-heat', null); return; }
  heatmap($('#ch-heat'), { years, rows, tipv: v => vfmt(v, m), aria: `Mapa de calor de programas por año, ${regName(r)}` });
  setTable('card-heat', { head: ['Programa', ...years.map(y => y + (partial(y) ? '*' : ''))], rows: rows.map(x => [short(x.n), ...x.vals.map(v => (v ? fmt(v) : ''))]) });
}

function renderAmounts() {
  const r = S.region, years = yrs();
  const uf = years.map(y => val(r, y, 'uf'));
  $('#t-uf').textContent = `Monto otorgado por año · ${regName(r)}`;
  $('#s-uf').textContent = `${SCOPES[S.scope]} · millones de UF (MM UF)`;
  columns($('#ch-uf'), { years, vals: uf, color: r === 'País' ? PAIS : SEL, kind: 'uf', lab: v => fmtMM(v, 2), tipv: v => fmtUF(v), aria: `Monto otorgado por año en UF, ${regName(r)}` });
  setTable('card-uf', { head: ['Año', 'Monto (UF)', 'Var. anual'], rows: years.map((y, i) => [y + (partial(y) ? '*' : ''), fmt(uf[i]), i > 0 && !partial(y) && dChg(uf[i], uf[i - 1]) !== null ? fmtD(dChg(uf[i], uf[i - 1]), 1) + '%' : '']) });

  const avg = (rg, y) => { const u = val(rg, y, 'u'); return u > 0 ? val(rg, y, 'uf') / u : null; };
  const series = [{ name: regName(r), color: r === 'País' ? PAIS : SEL, vals: years.map(y => avg(r, y)) }];
  if (r !== 'País') series.push({ name: 'Total país', color: PAIS, vals: years.map(y => avg('País', y)) });
  $('#t-avg').textContent = 'Monto promedio por subsidio (UF)';
  $('#s-avg').textContent = `${SCOPES[S.scope]} · UF por subsidio otorgado${r !== 'País' ? ' · región vs. país' : ''}`;
  lines($('#ch-avg'), { years, series, axis: (v) => fmt(v), unit: 'UF', aria: 'Monto promedio por subsidio en UF', tipv: v => fmt(v) + ' UF', mleft: 50 });
  setTable('card-avg', { head: ['Año', ...series.map(s => s.name + ' (UF/subsidio)')], rows: years.map((y, i) => [y + (partial(y) ? '*' : ''), ...series.map(s => (s.vals[i] === null ? '' : fmt(s.vals[i])))]) });
}

function renderRegions() {
  const m = S.metric; const arr = regionTotals().sort((a, b) => b.v - a.v); const tot = arr.reduce((a, d) => a + d.v, 0);
  $('#t-rank').textContent = `${M[m].name === 'Monto' ? 'Monto otorgado' : 'Subsidios otorgados'} por región · ${periodTxt()}`;
  $('#s-rank').textContent = `${SCOPES[S.scope]} · ordenadas de mayor a menor${S.to === LAST() ? ' · incluye ' + LAST() + ' parcial' : ''}`;
  const items = arr.map(d => ({ ...d, sel: d.name === S.region }));
  hbars($('#ch-rank'), {
    items, aria: `Ranking de regiones por ${M[m].short}`, lab: v => (m === 'uf' ? fmtMM(v, 1) + ' MM UF' : fmt(v)),
    tip: i => `<b>${esc(items[i].name)}</b><div>${vfmt(items[i].v, m)}</div><div class="m">${pct(items[i].v / tot * 100)} del total · lugar ${i + 1}</div><div class="m">Clic para analizar esta región</div>`,
  });
  setTable('card-rank', { head: ['#', 'Región', m === 'uf' ? 'Monto (UF)' : 'Subsidios', '% del total'], rows: arr.map((d, i) => [i + 1, d.name, fmt(d.v), fmtD(d.v / tot * 100, 1) + '%']) });

  $('#s-dist').textContent = `${periodTxt()} · ${SCOPES[S.scope]} · ordenadas de norte a sur · porcentaje de ${m === 'uf' ? 'monto' : 'subsidios'}`;
  $('#lg-dist').innerHTML = CATS.map(c => `<span><i style="background:${c.color}"></i>${esc(c.name)}</span>`).join('');
  const rows = ['País', ...D.regiones].map(rg => ({ name: regName(rg), key: rg, sel: rg === S.region, vals: catSeries(rg).map(s => s.vals.reduce((a, b) => a + b, 0)) }));
  share100($('#ch-dist'), { rows, tipv: v => vfmt(v, m), aria: 'Distribución por grupo de programas en cada región' });
  setTable('card-dist', { head: ['Territorio', ...CATS.map(c => c.name + ' (%)')], rows: rows.map(r => { const t = r.vals.reduce((a, b) => a + b, 0); return [r.name, ...r.vals.map(v => (t ? fmtD(v / t * 100, 1) : '0,0'))]; }) });
}

function renderCompare() {
  const m = S.metric; let years = yrs(); const hadPartial = years.some(partial); if (years.filter(y => !partial(y)).length >= 2) years = years.filter(y => !partial(y));
  const chips = ['País', ...D.regiones].map(rg => {
    const on = S.cmp.find(c => c.name === rg);
    return `<button type="button" class="chip" data-r="${esc(rg)}" aria-pressed="${!!on}" style="${on ? `border-color:${on.color}` : ''}">${on ? `<span class="dot" style="background:${on.color}"></span>` : ''}${esc(regName(rg))}</button>`;
  }).join('');
  $('#cmp-chips').innerHTML = chips;
  $$('#cmp-chips .chip').forEach(b => b.addEventListener('click', () => {
    const rg = b.dataset.r; const i = S.cmp.findIndex(c => c.name === rg);
    if (i >= 0) { if (S.cmp.length > 1) S.cmp.splice(i, 1); }
    else if (S.cmp.length < 4) { const free = CMP_COLORS.find(c => !S.cmp.some(x => x.color === c)); S.cmp.push({ name: rg, color: free }); }
    renderCompare();
  }));
  $('#seg-cmp').querySelectorAll('button').forEach(b => b.setAttribute('aria-pressed', String(b.dataset.v === S.cmpMode)));
  const base = years[0];
  const series = S.cmp.map(c => {
    const raw = years.map(y => val(c.name, y, m)); const b0 = raw[0];
    return { name: regName(c.name), color: c.color, vals: S.cmpMode === 'idx' ? raw.map(v => (b0 > 0 ? v / b0 * 100 : null)) : raw };
  });
  $('#t-cmp').textContent = S.cmpMode === 'idx' ? `Índice de ${M[m].short} otorgados (${base} = 100)` : `${M[m].name === 'Monto' ? 'Monto otorgado' : 'Subsidios otorgados'} por año`;
  $('#s-cmp').textContent = `${SCOPES[S.scope]} · ${years[0]}–${years[years.length - 1]}${hadPartial && !years.some(partial) ? ' · solo años completos' : ''}${m === 'uf' && S.cmpMode === 'abs' ? ' · UF' : ''}`;
  $('#lg-cmp').innerHTML = series.map(s => `<span><i class="line" style="background:${s.color}"></i>${esc(s.name)}</span>`).join('');
  const ax = S.cmpMode === 'idx' ? (v => fmt(v)) : axisFmt(m);
  lines($('#ch-cmp'), {
    years, series, axis: ax, unit: S.cmpMode === 'abs' && m === 'uf' ? 'MM UF' : '', aria: 'Comparación entre territorios',
    tipv: v => (S.cmpMode === 'idx' ? fmtD(v, 0) : vfmt(v, m)), mleft: 50,
  });
  setTable('card-cmp', { head: ['Año', ...series.map(s => s.name)], rows: years.map((y, i) => [y + (partial(y) ? '*' : ''), ...series.map(s => (s.vals[i] === null ? '' : fmt(s.vals[i])))]) });
}

function renderNotes() {
  const L = LAST(), f = FULLY();
  const noReg = Object.entries(D.pais.T.prog).filter(([n, p]) => p.u[ix(L)] === 0 && p.u[ix(f)] > 0).sort((a, b) => b[1].u[ix(f)] - a[1].u[ix(f)]).slice(0, 3);
  const ex = noReg.length ? ` Por ejemplo, ${noReg.map(([n, p]) => `${esc(short(n))} (${fmt(p.u[ix(f)])} subsidios en ${f})`).join(', ')} aún no registra${noReg.length > 1 ? 'n' : ''} subsidios en ${L}.` : '';
  $('#notes').innerHTML = [
    `<p><b>Fuente.</b> MINVU, sección Estadísticas, «Viviendas contratadas y subsidios otorgados» (reporte mensual Metas Subsidios Asignados, DPH a DIFIN). Archivo: ${esc(D.meta.archivo)}${D.meta.publicado ? `, publicado el ${esc(D.meta.publicado)}` : ''}. <a href="${esc(D.meta.url_fuente)}" target="_blank" rel="noopener">Ir a la página del MINVU</a>.</p>`,
    `<p><b>${L} es un año parcial.</b> Corresponde a enero–${CORTE()}. La entrega de subsidios depende de los llamados de cada programa, así que ${L} no es comparable con un año completo.${ex}</p>`,
    `<p><b>Montos en UF.</b> ${esc(D.meta.nota_uf)} Por eso este panel muestra los montos en UF y no los convierte a pesos.</p>`,
    `<p><b>Regiones nuevas.</b> Arica y Parinacota y Los Ríos aparecen con datos desde 2008, y Ñuble desde 2019. Antes de eso sus subsidios están contados dentro de Tarapacá, Los Lagos y Biobío, respectivamente.</p>`,
    `<p><b>Alcance.</b> «Total» suma el programa regular y la reconstrucción. Las líneas «FET» y «Reconstrucción FET» se mantienen con el nombre que usa el MINVU en su archivo. Hasta 2012 se contrataban viviendas SERVIU; desde 2013 solo se otorgan subsidios.</p>`,
    `<p><b>Validación.</b> Al cargar los datos se verificó que la suma de programas coincida con el total de cada región, que las 16 regiones sumen el total país y que cada hoja coincida con la tabla resumen del MINVU (${esc(D.meta.validacion)}).</p>`,
  ].join('');
  $('#foot').innerHTML = `Datos del MINVU procesados por el equipo de Estudios, CChC Sede Región de O'Higgins. Actualizado el ${esc(D.meta.generado.split('-').reverse().join('-'))}.`;
}

function renderAll(animate = false) {
  const go = animate && !REDUCE;
  $$('.chart').forEach(c => c.classList.toggle('animate', go));
  if (go) $$('.card').forEach(c => c.classList.remove('seen'));
  renderLead(); renderKPIs(go); renderAnnual(); renderCats(); renderProgTable(); renderHeat(); renderAmounts(); renderRegions(); renderCompare(); renderBarrio(); renderDetalle();
  if (go || firstRender) watchCards();
  $('#minuta-lbl').textContent = `Minuta de ${regName(S.region)} (Word)`;
  $('#btn-minuta').href = `minutas/Minuta_subsidios_${slug(S.region === 'País' ? 'pais' : S.region)}.docx`;
  firstRender = false;
}

/* ================= Movimiento ================= */
const REDUCE = matchMedia('(prefers-reduced-motion: reduce)').matches;
let io = null;
function watchCards() {
  if (!('IntersectionObserver' in window)) { $$('.card').forEach(c => { c.classList.add('seen'); c.classList.remove('pre'); }); return; }
  if (!io) io = new IntersectionObserver(es => es.forEach(e => { if (e.isIntersecting) { e.target.classList.add('seen'); e.target.classList.remove('pre'); io.unobserve(e.target); } }), { threshold: .12 });
  $$('.card').forEach(c => { if (!c.classList.contains('seen')) io.observe(c); });
}
function countUp(el) {
  const txt = el.textContent; const m = txt.match(/\d[\d.]*(?:,\d+)?/); if (!m || REDUCE) return;
  const raw = m[0]; const dec = raw.includes(',') ? raw.split(',')[1].length : 0;
  const target = parseFloat(raw.replace(/\./g, '').replace(',', '.')); if (!isFinite(target) || target === 0) return;
  const pre = txt.slice(0, m.index), post = txt.slice(m.index + raw.length), t0 = performance.now(), dur = 900;
  const step = now => { const k = Math.min(1, (now - t0) / dur), e = 1 - Math.pow(1 - k, 3); el.textContent = k < 1 ? pre + (dec ? fmtD(target * e, dec) : fmt(target * e)) + post : txt; if (k < 1) requestAnimationFrame(step); };
  requestAnimationFrame(step);
}
function parallax() {
  if (REDUCE) return; const hero = $('.hero'); let tick = false;
  addEventListener('scroll', () => { if (tick) return; tick = true; requestAnimationFrame(() => { hero.style.setProperty('--py', Math.min(scrollY, 700)); tick = false; }); }, { passive: true });
}

/* ================= Subsidios por nombre (año elegido y comparación) ================= */
function progRows(a, b, all) {
  const pg = PROG(S.region);
  return CATS.map(c => {
    const rows = D.programas.filter(p => p.categoria === c.key).map(p => {
      const d = pg[p.nombre]; const g = (m, y) => (d && y ? d[m][ix(y)] || 0 : 0);
      const act = d ? ANIOS().filter((y, i) => d.u[i] > 0 || d.uf[i] > 0) : [];
      return { n: p.nombre, u: g('u', a), uf: g('uf', a), ub: b ? g('u', b) : 0, ufb: b ? g('uf', b) : 0, first: act[0], last: act[act.length - 1], gaps: act.length > 1 && act.length < act[act.length - 1] - act[0] + 1 };
    }).filter(r => all || r.u > 0 || r.uf > 0 || r.ub > 0 || r.ufb > 0);
    return { cat: c, rows };
  });
}
const yl = y => y + (partial(y) ? ' (ene–' + CABR() + ')' : '');
function renderDetalle() {
  const a = S.pA, b = S.pB, m = S.metric, all = S.pAll; const groups = progRows(a, b, all);
  const key = m === 'uf' ? 'uf' : 'u', keyB = m === 'uf' ? 'ufb' : 'ub', oth = m === 'uf' ? 'u' : 'uf';
  const fv = v => (m === 'uf' ? fmtMM(v, 2) : fmt(v)); const unit = m === 'uf' ? 'MM UF' : 'subsidios';
  const totA = groups.reduce((s0, g) => s0 + g.rows.reduce((x, r) => x + r[key], 0), 0);
  const all_ = groups.flatMap(g => g.rows);
  const nuevos = b ? all_.filter(r => r[key] > 0 && r[keyB] === 0) : [], fuera = b ? all_.filter(r => r[key] === 0 && r[keyB] > 0) : [];
  const lista = arr => arr.map(r => esc(r.n)).join('; ');
  $('#det-sum').innerHTML = `<p><b>${all_.filter(r => r[key] > 0).length} programas</b> con subsidios en ${yl(a)}${b ? ` y <b>${all_.filter(r => r[keyB] > 0).length}</b> en ${yl(b)}` : ''} · ${regName(S.region)} · ${SCOPES[S.scope]}.</p>` +
    (b && nuevos.length ? `<p class="up">Con subsidios en ${a} y sin registros en ${b}: ${lista(nuevos)}.</p>` : '') +
    (b && fuera.length ? `<p class="down">Con subsidios en ${b} y sin registros en ${a}: ${lista(fuera)}.</p>` : '');
  const vr = (x, y) => { if (!b) return ''; const d = x - y; if (x === 0 && y === 0) return '<td>—</td>'; if (y === 0) return `<td><span class="tag nuevo">Solo en ${a}</span></td>`; if (x === 0) return `<td><span class="tag fin">Solo en ${b}</span></td>`; const pc = Math.abs(d / y * 100); return `<td class="${d >= 0 ? 'pos' : 'neg'}">${d >= 0 ? '+' : '−'}${fv(Math.abs(d))}${pc < 1000 ? ` (${d >= 0 ? '+' : '−'}${fmtD(pc, 1)}%)` : ''}</td>`; };
  let h = `<table class="t det"><thead><tr><th scope="col">Programa</th><th scope="col">Con registros</th><th scope="col">${yl(a)}</th>${b ? `<th scope="col">${yl(b)}</th><th scope="col">Variación</th>` : ''}<th scope="col">${m === 'uf' ? 'Subsidios' : 'Monto MM UF'} ${a}</th><th scope="col">% del año</th></tr></thead><tbody>`;
  const rowsOut = [];
  groups.forEach(g => {
    if (!g.rows.length) return;
    const sa = g.rows.reduce((x, r) => x + r[key], 0), sb = g.rows.reduce((x, r) => x + r[keyB], 0), so = g.rows.reduce((x, r) => x + r[oth], 0);
    h += `<tr class="grp" style="--c:${g.cat.color}"><td colspan="2"><i></i>${esc(g.cat.name)}</td><td>${fv(sa)}</td>${b ? `<td>${fv(sb)}</td>${vr(sa, sb)}` : ''}<td>${oth === 'uf' ? fmtMM(so, 2) : fmt(so)}</td><td>${totA ? pct(sa / totA * 100, 1) : '—'}</td></tr>`;
    g.rows.forEach(r => {
      const vig = r.first ? (r.first === r.last ? `${r.first}` : `${r.first}–${r.last}`) + (r.gaps ? ' (con pausas)' : '') : 'Sin registros';
      h += `<tr><td class="pn" style="--c:${g.cat.color}">${esc(r.n)}</td><td class="vig">${vig}</td><td>${r[key] ? fv(r[key]) : '—'}</td>${b ? `<td>${r[keyB] ? fv(r[keyB]) : '—'}</td>${vr(r[key], r[keyB])}` : ''}<td>${r[oth] ? (oth === 'uf' ? fmtMM(r[oth], 2) : fmt(r[oth])) : '—'}</td><td>${r[key] && totA ? `<span class="minibar" style="width:${Math.max(2, r[key] / totA * 90)}px;background:${g.cat.color}"></span> ${pct(r[key] / totA * 100, 1)}` : '—'}</td></tr>`;
      rowsOut.push([g.cat.name, r.n, r.first ? `${r.first}–${r.last}` : '', r.u, Math.round(r.uf), b ? r.ub : '', b ? Math.round(r.ufb) : '']);
    });
  });
  h += `<tr class="tot"><td colspan="2">Total ${regName(S.region)}</td><td>${fv(totA)}</td>${b ? `<td>${fv(groups.reduce((x, g) => x + g.rows.reduce((y, r) => y + r[keyB], 0), 0))}</td><td></td>` : ''}<td></td><td>100%</td></tr></tbody></table>`;
  $('#tb-det').innerHTML = all_.length ? h : '<p class="note">No hay subsidios registrados en este alcance y año.</p>';
}

/* ================= Barrio: cada casita = N subsidios ================= */
const CASA = [
  '<rect x="10" y="14" width="4" height="7" rx="1" fill="#fff"/>',
  '<rect x="15.5" y="4" width="3" height="5" fill="currentColor"/><rect x="10" y="14" width="4" height="7" rx="1" fill="#fff"/>',
  '<path d="M12 12v6M9 15h6" stroke="#fff" stroke-width="2" stroke-linecap="round"/>',
  '<circle cx="9.6" cy="15" r="2" fill="none" stroke="#fff" stroke-width="1.6"/><path d="M11.6 15H17m-2 0v2.2" stroke="#fff" stroke-width="1.6" stroke-linecap="round"/>',
];
const casaSVG = k => `<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M2.5 11.6 12 3.4l9.5 8.2V21h-19z" fill="currentColor"/>${CASA[k]}</svg>`;
const UNIDADES = [10, 25, 50, 100, 250, 500, 1000, 2500, 5000, 10000, 25000];
function renderBarrio() {
  const r = S.region, f = S.barrio || FULLY(), sc = S.scope;
  const ser = catSeries(r, 'u', sc, [f]).map((x, k) => ({ k, name: x.name, color: x.color, v: x.vals[0] }));
  const total = ser.reduce((a, x) => a + x.v, 0);
  $('#t-barrio').textContent = total > 0 ? `${fmt(total)} subsidios de ${f}, convertidos en un barrio` : `Sin subsidios en ${f}`;
  $('#s-barrio').textContent = `${regName(r)} · ${SCOPES[sc]} · cada casita representa una cantidad fija de subsidios`;
  if (!(total > 0)) { $('#ch-barrio').innerHTML = '<p class="note">No hay subsidios registrados en este alcance y año.</p>'; $('#lg-barrio').innerHTML = ''; setTable('card-barrio', null); return; }
  const U = UNIDADES.find(u => total / u <= 120) || UNIDADES[UNIDADES.length - 1];
  const target = Math.max(1, Math.round(total / U));
  let cnt = ser.map(x => Math.floor(x.v / U)); let rest = target - cnt.reduce((a, b) => a + b, 0);
  ser.map((x, k) => [k, x.v / U - Math.floor(x.v / U)]).sort((a, b) => b[1] - a[1]).forEach(([k]) => { if (rest > 0) { cnt[k]++; rest--; } });
  let idx = 0;
  $('#ch-barrio').innerHTML = `<div class="barrio" role="img" aria-label="Barrio de ${target} casitas; cada una representa ${fmt(U)} subsidios">` +
    ser.map((x, k) => Array.from({ length: cnt[k] }, () => `<span class="h" data-i="${k}" style="color:${x.color};--i:${idx++}">${casaSVG(k)}</span>`).join('')).join('') + '</div>';
  $('#lg-barrio').innerHTML = ser.map(x => `<span class="lgcasa" style="color:${x.color}">${casaSVG(x.k)}<span><b>${esc(x.name)}</b> · ${fmt(x.v)} <em>(${pct(x.v / total * 100, 0)})</em></span></span>`).join('') + `<span class="unidad">1 casita = ${fmt(U)} subsidios</span>`;
  bindHover($('#ch-barrio'), h => { const x = ser[+h.dataset.i]; return `<b>${esc(x.name)}</b><div>${fmt(x.v)} subsidios · ${pct(x.v / total * 100, 1)}</div><div class="m">1 casita = ${fmt(U)} subsidios</div>`; });
  setTable('card-barrio', { head: ['Grupo de programas', 'Subsidios', '% del año', 'Casitas'], rows: ser.map((x, k) => [x.name, fmt(x.v), fmtD(x.v / total * 100, 1) + '%', cnt[k]]) });
}

/* ================= Controles ================= */
function fitRegionSelect() {
  const sel = $('#region'); const wrap = sel.parentElement;
  let ms = $('.measure', wrap);
  if (!ms) { ms = document.createElement('span'); ms.className = 'measure'; ms.setAttribute('aria-hidden', 'true'); wrap.appendChild(ms); }
  ms.textContent = sel.options[sel.selectedIndex].text;
  sel.style.width = Math.ceil(ms.getBoundingClientRect().width + 62) + 'px';
}
function selectRegion(r) {
  S.region = r; $('#region').value = r; fitRegionSelect();
  const go = () => renderAll(true);
  if (!REDUCE && document.startViewTransition) document.startViewTransition(go); else go();
}
function syncPresets() {
  const a = ANIOS()[0], L = LAST(); const f = FULLY();
  const cur = S.from === a && S.to === L ? 'all' : S.from === 2010 && S.to === L ? '2010' : S.from === 2015 && S.to === L ? '2015' : S.from === 2015 && S.to === f ? 'full' : '';
  $$('#presets .chip').forEach(b => b.setAttribute('aria-pressed', String(b.dataset.p === cur)));
}
function initControls() {
  const reg = $('#region'); reg.innerHTML = ['País', ...D.regiones].map(r => `<option value="${esc(r)}">${esc(r === 'País' ? 'Total país' : r)}</option>`).join(''); reg.value = S.region; fitRegionSelect(); if (document.fonts) { document.fonts.ready.then(fitRegionSelect); document.fonts.load('800 60px "Plus Jakarta Sans"').then(fitRegionSelect).catch(() => {}); addEventListener('load', () => setTimeout(fitRegionSelect, 300)); }
  reg.addEventListener('change', () => selectRegion(reg.value));
  const ys = ANIOS(); const opts = ys.map(y => `<option value="${y}">${y}${partial(y) ? ' (ene–' + CABR() + ')' : ''}</option>`).join('');
  $('#from').innerHTML = opts; $('#to').innerHTML = opts; $('#from').value = S.from; $('#to').value = S.to;
  const onPeriod = () => { let a = +$('#from').value, b = +$('#to').value; if (a > b) { [a, b] = [b, a]; } S.from = a; S.to = b; $('#from').value = a; $('#to').value = b; syncPresets(); renderAll(); };
  $('#from').addEventListener('change', onPeriod); $('#to').addEventListener('change', onPeriod);
  $$('#presets .chip').forEach(b => b.addEventListener('click', () => {
    const p = b.dataset.p, L = LAST();
    S.from = p === 'all' ? ys[0] : p === '2010' ? 2010 : 2015; S.to = p === 'full' ? FULLY() : L;
    $('#from').value = S.from; $('#to').value = S.to; syncPresets(); renderAll();
  }));
  const seg = (id, key) => $$(`#${id} button`).forEach(b => b.addEventListener('click', () => { S[key] = b.dataset.v; $$(`#${id} button`).forEach(x => x.setAttribute('aria-pressed', String(x === b))); renderAll(); }));
  seg('seg-scope', 'scope'); seg('seg-metric', 'metric');
  const yo = [...ys].reverse().map(y => `<option value="${y}">${yl(y)}</option>`).join('');
  $('#pa').innerHTML = yo; $('#pb').innerHTML = '<option value="">Sin comparar</option>' + yo;
  S.pA = FULLY(); S.pB = FULLY() - 1; $('#pa').value = S.pA; $('#pb').value = S.pB;
  $('#pa').addEventListener('change', () => { S.pA = +$('#pa').value; renderDetalle(); });
  $('#pb').addEventListener('change', () => { S.pB = $('#pb').value ? +$('#pb').value : null; renderDetalle(); });
  $('#pall').addEventListener('change', () => { S.pAll = $('#pall').checked; renderDetalle(); });
  const by = $('#barrio-year'); by.innerHTML = ys.filter(y => !partial(y)).reverse().map(y => `<option value="${y}">${y}</option>`).join(''); by.value = FULLY();
  by.addEventListener('change', () => { S.barrio = +by.value; renderBarrio(); });
  $$('#seg-cmp button').forEach(b => b.addEventListener('click', () => { S.cmpMode = b.dataset.v; renderCompare(); }));
  $('#btn-xlsx').addEventListener('click', exportExcel);
  initToggles();
  let rt; addEventListener('resize', () => { clearTimeout(rt); rt = setTimeout(() => { if (D) renderAll(); }, 160); });
}

/* ================= Excel ================= */
function exportExcel() {
  if (typeof XLSX === 'undefined') { alert('No se pudo cargar la librería de Excel. Revisa tu conexión e inténtalo de nuevo.'); return; }
  const r = S.region, L = LAST(), wb = XLSX.utils.book_new(); const years = ANIOS();
  const add = (name, aoa, widths, fm = {}) => {
    const ws = XLSX.utils.aoa_to_sheet(aoa); if (widths) ws['!cols'] = widths.map(w => ({ wch: w }));
    const rng = XLSX.utils.decode_range(ws['!ref']);
    for (let R = 1; R <= rng.e.r; R++) for (const c in fm) { const cell = ws[XLSX.utils.encode_cell({ r: R, c: +c })]; if (cell && typeof cell.v === 'number') cell.z = fm[c]; }
    XLSX.utils.book_append_sheet(wb, ws, name);
  };
  const N = '#,##0', P = '0.0%';
  const tT = TOT(r, S.scope), pT = TOT('País', S.scope);

  add('Resumen', [
    ['Subsidios habitacionales otorgados · CChC O\'Higgins'], [],
    ['Territorio', regName(r)], ['Alcance', SCOPES[S.scope]], ['Período', periodTxt()],
    ['Subsidios otorgados en el período', Math.round(sumR(tT.u))], ['Monto otorgado en el período (UF)', Math.round(sumR(tT.uf))],
    ['UF por subsidio (promedio del período)', Math.round(sumR(tT.uf) / Math.max(1, sumR(tT.u)))],
    ['Participación en el país (subsidios)', r === 'País' ? 1 : sumR(tT.u) / sumR(pT.u)], [],
    ['Notas'], [`${L} es un año parcial (enero–${CORTE()}).`], [D.meta.nota_uf], ['Fuente: ' + D.meta.fuente], [D.meta.url_fuente],
    ['Archivo: ' + D.meta.archivo + (D.meta.publicado ? ' · publicado el ' + D.meta.publicado : '')],
  ], [44, 40], { 1: N });
  { const ws = wb.Sheets['Resumen']; const c = ws['B9']; if (c && typeof c.v === 'number') c.z = P; }

  add('Serie anual', [['Año', 'Subsidios ' + regName(r), 'Monto UF ' + regName(r), 'UF por subsidio', 'Subsidios total país', 'Monto UF total país', 'Participación (subsidios)', 'Observación'],
    ...years.map(y => { const i = ix(y); const u = tT.u[i], uf = tT.uf[i], pu = pT.u[i]; return [y, u, uf, u ? Math.round(uf / u) : '', pu, pT.uf[i], pu ? u / pu : '', partial(y) ? 'Año parcial (ene–' + CABR() + ')' : '']; })], [8, 20, 20, 16, 20, 20, 22, 22], { 1: N, 2: N, 3: N, 4: N, 5: N, 6: P });

  const pg = PROG(r, S.scope); const prog = Object.entries(pg).map(([n, p]) => ({ n, u: sumR(p.u), uf: sumR(p.uf) })).filter(x => x.u || x.uf).sort((a, b) => b.u - a.u);
  const tu = prog.reduce((a, x) => a + x.u, 0);
  add('Programas', [['Programa', 'Grupo', 'Subsidios', '% subsidios', 'Monto UF', 'UF por subsidio'], ...prog.map(x => [x.n, catName(x.n), x.u, tu ? x.u / tu : 0, Math.round(x.uf), x.u ? Math.round(x.uf / x.u) : ''])], [70, 32, 12, 12, 16, 16], { 2: N, 3: P, 4: N, 5: N });
  const names = Object.keys(pg).filter(n => pg[n].u.some(v => v > 0));
  add('Programas por año (subsidios)', [['Programa', ...years.map(y => y + (partial(y) ? '*' : ''))], ...names.map(n => [n, ...pg[n].u])], [70]);
  add('Programas por año (UF)', [['Programa', ...years.map(y => y + (partial(y) ? '*' : ''))], ...names.map(n => [n, ...pg[n].uf.map(v => Math.round(v))])], [70]);

  const rk = D.regiones.map(g => ({ g, u: sumR(TOT(g, S.scope).u), uf: sumR(TOT(g, S.scope).uf) })).sort((a, b) => b.u - a.u); const ru = rk.reduce((a, x) => a + x.u, 0);
  add('Ranking regiones', [['Lugar', 'Región', 'Subsidios', '% del total', 'Monto UF', 'UF por subsidio'], ...rk.map((x, i) => [i + 1, x.g, x.u, ru ? x.u / ru : 0, Math.round(x.uf), x.u ? Math.round(x.uf / x.u) : ''])], [8, 22, 14, 12, 18, 16], { 2: N, 3: P, 4: N, 5: N });
  add('Grupo por región', [['Territorio', ...CATS.flatMap(c => [c.name + ' (subsidios)', c.name + ' (UF)'])],
    ...['País', ...D.regiones].map(g => [regName(g), ...catSeries(g, 'u').flatMap((s, k) => [s.vals.reduce((a, b) => a + b, 0), Math.round(catSeries(g, 'uf')[k].vals.reduce((a, b) => a + b, 0))])])], [22, 22, 20, 22, 20, 22, 20, 22, 20].slice(0, 9));

  { const a = S.pA, b = S.pB; const rows = [];
    progRows(a, b, false).forEach(g => g.rows.forEach(r => rows.push([g.cat.name, r.n, r.first ? `${r.first}–${r.last}` : '', r.u, Math.round(r.uf), ...(b ? [r.ub, Math.round(r.ufb), r.ub ? (r.u - r.ub) / r.ub : ''] : [])])));
    add('Subsidios por nombre', [['Grupo', 'Programa', 'Años con registros', `Subsidios ${a}`, `Monto UF ${a}`, ...(b ? [`Subsidios ${b}`, `Monto UF ${b}`, 'Var. subsidios'] : [])], ...rows], [32, 70, 18, 14, 16, 14, 16, 14], { 3: N, 4: N, 5: N, 6: N, 7: P }); }
  const base = [['Territorio', 'Alcance', 'Año', 'Grupo', 'Programa', 'Subsidios', 'Monto UF']];
  ['País', ...D.regiones].forEach(g => ['R', 'X'].forEach(sc => { Object.entries(PROG(g, sc)).forEach(([n, p]) => years.forEach((y, i) => { if (p.u[i] || p.uf[i]) base.push([regName(g), SCOPES[sc], y, catName(n), n, p.u[i], Math.round(p.uf[i])]); })); }));
  add('Base completa', base, [20, 18, 8, 34, 70, 12, 16]);

  XLSX.writeFile(wb, `Subsidios_${slug(r === 'País' ? 'pais' : r)}_${S.from}-${S.to}.xlsx`);
}

/* ================= Arranque ================= */
async function init() {
  try {
    const res = await fetch('data/subsidios.json', { cache: 'no-cache' });
    if (!res.ok) throw new Error('HTTP ' + res.status);
    D = await res.json();
  } catch (e) {
    $('#loading').outerHTML = `<div class="err"><b>No se pudieron cargar los datos.</b> Si abriste el archivo directamente desde tu computador, publícalo en GitHub Pages o ábrelo con un servidor local (el navegador bloquea la lectura de data/subsidios.json). Detalle: ${esc(e.message)}</div>`;
    return;
  }
  D._cat = {}; D.programas.forEach(p => { D._cat[p.nombre] = p.categoria; });
  S.to = LAST(); S.cmp = ["O'Higgins", 'Maule', 'Valparaíso', 'Biobío'].map((n, i) => ({ name: n, color: CMP_COLORS[i] }));
  const qs = new URLSearchParams(location.search); const qr = qs.get('region'); if (qr && (qr === 'País' || D.reg[qr])) S.region = qr;
  initControls(); syncPresets();
  $('#seg-cmp').querySelectorAll('button').forEach(b => b.setAttribute('aria-pressed', String(b.dataset.v === S.cmpMode)));
  $('#loading').hidden = true; $('#content').hidden = false;
  renderNotes(); $$('.card').forEach(c => c.classList.add('pre')); renderAll(true); parallax();
  fetch('data/ultima_revision.txt', { cache: 'no-cache' }).then(r => (r.ok ? r.text() : '')).then(t => { t = (t || '').trim(); if (/^\d{4}-\d{2}-\d{2}$/.test(t)) $('#foot').insertAdjacentHTML('beforeend', ` Última revisión automática: ${t.split('-').reverse().join('-')}.`); }).catch(() => {});
}
init();
