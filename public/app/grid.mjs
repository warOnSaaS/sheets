// The spreadsheet editor. It keeps a copy of the workbook (a Yjs document) and works formulas out in the
// browser with the same engine as the server. Every change goes through a tool (ctx.callTool), exactly as
// an agent would make it; the server's change comes back over the live feed (or sheets.sync) and merges in.
// The grid is drawn as a window of cells over a scrolling area, so a tab of 100,000 rows scrolls smoothly.
import { Y, tabList, findTab, getRaw, getFmt, cellMap, toB64, fromB64, usedRange, comments as allComments, charts as allCharts } from '../../lib/book.mjs';
import { Engine, toHf } from '../../lib/engine.mjs';
import { addr, colName, colIndex, key, unkey, parseRange, rangeText, shiftFormula } from '../../lib/a1.mjs';
import { parseInput, display, editText, FORMATS, plainValue, isDateFormat } from '../../lib/input.mjs';
import * as ops from '../../lib/ops.mjs';
import { renderChart } from './chart.mjs';
import { h, none, tool, I, dialog, menu, closeMenu, popover, closePop, toast, initials, ago, copyText, showError } from './ui.mjs';

const HW = 46, HH = 26;
const ERR = /^#[A-Z0-9/!?]+[!?]?$/;
const COLORS = ['#0b0b0b', '#454545', '#8a8a8a', '#c9c9c9', '#ffffff', '#b42318', '#a26a16', '#1a7f37', '#0e7490', '#1d4ed8', '#7c3aed', '#be185d'];
const FILLS = [null, '#f1f3f5', '#e7e7e4', '#fff4cc', '#fde2e1', '#dcf3e4', '#dbeafe', '#ede9fe', '#fce7f3', '#e0f2f1', '#ffe8cc', '#0b0b0b'];

export function openEditor(root, ctx, { sheet, view = null }) {
  const readOnly = !!view;
  const phone = () => root.clientWidth < 720;
  const S = {
    sheet, doc: new Y.Doc(), engine: null, tab: null, title: '', share: null,
    sel: { r: 0, c: 0, r2: 0, c2: 0 }, anchor: { r: 0, c: 0 },
    editing: null, undo: [], redo: [], clip: null, others: new Map(), destroyed: false, dirty: true,
    geo: null, hidden: new Set(), sx: 0, sy: 0, refs: [],
  };
  root.innerHTML = `<div class="ed${readOnly ? ' is-view' : ''}">
  <header class="ed-top">
    ${readOnly ? `<a class="ed-brand" href="/" aria-label="Sheets">${I.grid}</a>` : `<a class="ed-back ui-btn is-ghost is-sm" href="${ctx.standalone ? '#/' : '/a/sheets'}" aria-label="All spreadsheets">${I.back}</a>`}
    <form class="ed-title" ${readOnly ? none('shows the title') : tool('sheets.rename_sheet')}><input name="title" aria-label="Spreadsheet title" autocomplete="off" spellcheck="false" ${readOnly ? 'readonly' : ''}></form>
    <div class="ed-who" aria-label="Who else is here"></div>
    <div class="ed-acts">${readOnly ? `<span class="ui-chip is-outline">View only</span><button type="button" class="ui-btn is-accent is-sm" data-act="copy" ${tool('sheets.copy_sheet')}>Make a copy to edit</button>` : `
      <button type="button" class="ui-btn is-ghost is-sm ed-ib" data-act="comments" ${tool('sheets.list_comments')} aria-label="Comments" title="Comments">${I.comment}<span class="ed-n" hidden></span></button>
      <button type="button" class="ui-btn is-ghost is-sm ed-ib" data-act="history" ${tool('sheets.list_versions')} aria-label="Version history" title="Version history">${I.history}</button>
      <button type="button" class="ui-btn is-ghost is-sm ed-ib" data-act="download" ${tool('sheets.export')} aria-label="Download" title="Download">${I.download}</button>
      <button type="button" class="ui-btn is-accent is-sm" data-act="share" ${tool('sheets.share')}>${I.share}<span class="ed-hide-s">Share</span></button>`}
    </div>
  </header>
  ${readOnly ? '' : '<div class="ed-tools" role="toolbar" aria-label="Formatting"></div>'}
  <div class="ed-fx">
    <form class="ed-name" ${none('moves the selection to the cell or range typed')}><input name="ref" aria-label="Cell" autocomplete="off" spellcheck="false"></form>
    <span class="ed-fxi" aria-hidden="true">fx</span>
    <form class="ed-formula" ${readOnly ? none('shows the cell contents') : tool('sheets.write_range')}><input name="v" aria-label="Cell contents" autocomplete="off" spellcheck="false" ${readOnly ? 'readonly' : ''}></form>
  </div>
  <div class="ed-grid">
    <div class="sg-scroll" tabindex="-1"><div class="sg-space"></div></div>
    <div class="sg-view" aria-hidden="true"><div class="sg-q q-body"></div><div class="sg-q q-top"></div><div class="sg-q q-left"></div><div class="sg-q q-corner"></div><div class="sg-q q-ch"></div><div class="sg-q q-chf"></div><div class="sg-q q-rh"></div><div class="sg-q q-rhf"></div><div class="sg-q q-all"></div></div>
    <div class="sg-ui"><div class="sg-charts"></div><textarea class="sg-input" aria-label="Edit cell" autocomplete="off" autocapitalize="off" spellcheck="false" rows="1" ${readOnly ? 'readonly' : ''}></textarea><div class="sg-ac" role="listbox" hidden></div></div>
    <div class="sg-load">Opening…</div>
  </div>
  <footer class="ed-foot"><nav class="ed-tabs" aria-label="Tabs"></nav><div class="ed-stat" aria-live="polite"></div></footer>
</div>`;
  const $ = (s) => root.querySelector(s);
  // Inside the suite, links inside the app move with ctx.navigate instead of loading a page.
  root.addEventListener('click', (e) => { const a = e.target.closest('a.ed-back, a[data-go]'); if (a && !ctx.standalone) { e.preventDefault(); ctx.navigate(a.dataset.go ?? '/'); } });
  const grid = $('.ed-grid'), scroll = $('.sg-scroll'), space = $('.sg-space'), viewEl = $('.sg-view'), uiEl = $('.sg-ui'), input = $('.sg-input'), ac = $('.sg-ac');
  const Q = Object.fromEntries(['body', 'top', 'left', 'corner', 'ch', 'chf', 'rh', 'rhf', 'all'].map((k) => [k, $(`.q-${k}`)]));
  const callTool = (name, inp = {}) => ctx.callTool(name, { sheet: S.sheet, ...inp });

  // ---------- loading and live changes ----------

  async function sync() {
    const sv = S.engine ? toB64(Y.encodeStateVector(S.doc)) : undefined;
    const r = await callTool('sheets.sync', sv ? { state_vector: sv } : {});
    Y.applyUpdate(S.doc, fromB64(r.update), 'remote');
    S.share = r.share;
    return r;
  }

  async function load() {
    try { await sync(); } catch (e) {
      $('.sg-load').innerHTML = `<div class="ui-empty"><b>${h(e.message)}</b><p><a href="${readOnly ? '/' : '#/'}">Back to the list</a></p></div>`;
      return;
    }
    if (S.destroyed) return;
    S.engine = new Engine(S.doc);
    S.doc.on('afterTransaction', (tr) => { if (tr.origin !== 'local') { S.dirty = true; schedule(); } });
    const want = new URLSearchParams((ctx.path.split('?')[1] ?? '')).get('tab');
    S.tab = (findTab(S.doc, want) ?? tabList(S.doc)[0])?.id;
    $('.sg-load').remove();
    if (!readOnly) buildToolbar();
    layout(true);
    selectCell(0, 0);
    if (!readOnly && matchMedia('(pointer:fine)').matches) focusGrid();
  }

  const onEvent = (e) => {
    const d = e.data ?? {};
    if (d.workbook !== S.sheet) return;
    if (e.type === 'sheets.workbook.changed') {
      if (d.update) { try { Y.applyUpdate(S.doc, fromB64(d.update), 'remote'); } catch { sync().catch(() => {}); } }
      else sync().catch(() => {});
    } else if (e.type === 'sheets.presence.changed' && d.person?.id !== ctx.me?.id) {
      if (d.left) S.others.delete(d.person.id); else S.others.set(d.person.id, { ...d, at: Date.now() });
      drawPeople(); schedule();
    } else if (e.type === 'sheets.workbook.deleted') {
      toast('This spreadsheet was deleted.');
      ctx.navigate('/');
    } else if (e.type === 'resync') sync().catch(() => {});
  };
  const offs = ['sheets.workbook.changed', 'sheets.presence.changed', 'sheets.workbook.deleted', 'resync'].map((n) => ctx.on(n, (e) => onEvent({ ...e, type: e.type ?? e.name ?? n, data: e.data ?? e })));
  const offLive = () => offs.forEach((f) => f?.());

  // A change: optimistic in this copy's engine, then the tool, then catch up with the server's version.
  async function run(name, inp, { quiet = false, undo = null } = {}) {
    try {
      const out = await callTool(name, inp);
      if (out?.pending) { toast('Waiting for a person\'s yes.'); return out; }
      if (undo) { S.undo.push(undo); if (S.undo.length > 100) S.undo.shift(); S.redo = []; }
      await sync();
      if (out?.errors?.length && !quiet) toast(`Formula error: ${out.errors[0]}`);
      return out;
    } catch (e) {
      S.engine?.build();
      S.dirty = true; schedule();
      toast(e.message);
      throw e;
    }
  }

  // ---------- geometry ----------

  const tab = () => findTab(S.doc, S.tab) ?? tabList(S.doc)[0];
  const RH = () => (phone() ? 32 : 28);
  const DW = () => (phone() ? 92 : 108);

  function layout(force = false) {
    const t = tab();
    if (!t) return;
    if (S.tab !== t.id) S.tab = t.id;
    S.hidden = ops.hiddenRows(S.doc, S.engine, t);
    const vis = [];
    const used = usedRange(S.doc, t.id);
    const rows = Math.max(t.rows, (used?.r2 ?? 0) + 1);
    for (let r = 0; r < rows; r++) if (!S.hidden.has(r)) vis.push(r);
    const pos = new Int32Array(rows).fill(-1);
    vis.forEach((r, i) => { pos[r] = i; });
    const cols = Math.max(t.cols, (used?.c2 ?? 0) + 1);
    const X = new Float64Array(cols + 1);
    for (let c = 0; c < cols; c++) X[c + 1] = X[c] + (t.widths[c] ?? DW());
    const rh = RH();
    const fr = Math.min(t.frozen_rows, rows), fc = Math.min(t.frozen_cols, cols);
    let frozenH = 0;
    for (let r = 0; r < fr; r++) if (pos[r] >= 0) frozenH += rh;
    S.geo = { t, rows, cols, vis, pos, X, rh, fr, fc, frozenW: X[fc], frozenH, W: X[cols], H: vis.length * rh };
    space.style.width = `${HW + S.geo.W + 60}px`;
    space.style.height = `${HH + S.geo.H + 120}px`;
    S.dirty = true;
    draw();
    drawTabs();
    if (force) syncBars();
  }

  const ry = (r) => { const g = S.geo; const p = g.pos[r]; return p < 0 ? -1 : p * g.rh; };
  const colAt = (x) => { const X = S.geo.X; let lo = 0, hi = S.geo.cols - 1; while (lo < hi) { const m = (lo + hi + 1) >> 1; if (X[m] <= x) lo = m; else hi = m - 1; } return lo; };
  // Screen point (inside the grid) -> { kind, r, c }.
  function hit(px, py) {
    const g = S.geo;
    const cx = px < HW + g.frozenW ? px - HW : px - HW + S.sx;
    const cyContent = py < HH + g.frozenH ? py - HH : py - HH + S.sy;
    const c = Math.max(0, Math.min(g.cols - 1, colAt(Math.max(0, cx))));
    const idx = Math.max(0, Math.min(g.vis.length - 1, Math.floor(Math.max(0, cyContent) / g.rh)));
    const r = g.vis[idx] ?? 0;
    if (px < HW && py < HH) return { kind: 'all', r, c };
    if (py < HH) return { kind: 'col', r, c, edge: edgeAt(px) };
    if (px < HW) return { kind: 'row', r, c };
    return { kind: 'cell', r, c };
  }
  function edgeAt(px) {
    const g = S.geo;
    for (let c = 0; c < g.cols; c++) {
      const x = HW + (c < g.fc ? g.X[c + 1] : g.X[c + 1] - S.sx);
      if (Math.abs(px - x) <= 4 && (c < g.fc || x > HW + g.frozenW)) return c;
      if (x > px + 8) break;
    }
    return null;
  }
  // Cell -> screen rectangle (inside the grid).
  function rect(r, c, r2 = r, c2 = c) {
    const g = S.geo;
    const x1 = HW + (c < g.fc ? g.X[c] : g.X[c] - S.sx);
    const x2 = HW + (c2 < g.fc ? g.X[c2 + 1] : g.X[c2 + 1] - S.sx);
    const y = (rr) => HH + (rr < g.fr ? ry(rr) : ry(rr) - S.sy);
    let top = y(r), bot = y(r2) + g.rh;
    if (ry(r) < 0) top = y(nextVisible(r));
    if (ry(r2) < 0) bot = y(prevVisible(r2)) + g.rh;
    return { x: x1, y: top, w: x2 - x1, h: bot - top };
  }
  const nextVisible = (r) => { while (r < S.geo.rows - 1 && S.geo.pos[r] < 0) r++; return r; };
  const prevVisible = (r) => { while (r > 0 && S.geo.pos[r] < 0) r--; return r; };

  // ---------- drawing ----------

  let frame = 0;
  function schedule() { if (!frame) frame = requestAnimationFrame(() => { frame = 0; if (S.dirty) { S.dirty = false; layout(); } else draw(); }); }

  function draw() {
    const g = S.geo;
    if (!g) return;
    const vw = scroll.clientWidth, vh = scroll.clientHeight;
    viewEl.style.width = uiEl.style.width = `${vw}px`;
    viewEl.style.height = uiEl.style.height = `${vh}px`;
    S.sx = scroll.scrollLeft; S.sy = scroll.scrollTop;
    const bodyX = HW + g.frozenW, bodyY = HH + g.frozenH;
    const place = (el, l, t, w, hh) => { el.style.left = `${l}px`; el.style.top = `${t}px`; el.style.width = `${Math.max(0, w)}px`; el.style.height = `${Math.max(0, hh)}px`; };
    place(Q.body, bodyX, bodyY, vw - bodyX, vh - bodyY);
    place(Q.top, bodyX, HH, vw - bodyX, g.frozenH);
    place(Q.left, HW, bodyY, g.frozenW, vh - bodyY);
    place(Q.corner, HW, HH, g.frozenW, g.frozenH);
    place(Q.ch, bodyX, 0, vw - bodyX, HH);
    place(Q.chf, HW, 0, g.frozenW, HH);
    place(Q.rh, 0, bodyY, HW, vh - bodyY);
    place(Q.rhf, 0, HH, HW, g.frozenH);
    place(Q.all, 0, 0, HW, HH);
    // Visible columns and rows of the scrolling part.
    const cs = [];
    for (let c = Math.max(g.fc, colAt(S.sx + g.frozenW)); c < g.cols && g.X[c] - S.sx < vw - HW; c++) cs.push(c);
    const csF = Array.from({ length: g.fc }, (_, i) => i);
    const i0 = Math.floor((S.sy + g.frozenH) / g.rh);
    const rs = [];
    for (let i = Math.max(i0, 0); i < g.vis.length && i * g.rh - S.sy < vh - HH; i++) if (g.vis[i] >= g.fr) rs.push(g.vis[i]);
    const rsF = g.vis.filter((r) => r < g.fr);
    // Each part draws with its own origin: the cells' content position minus where the part starts.
    const ox = (c) => (c < g.fc ? g.X[c] : g.X[c] - S.sx - g.frozenW);
    const oy = (r) => (r < g.fr ? ry(r) : ry(r) - S.sy - g.frozenH);
    Q.body.innerHTML = cells(rs, cs, ox, oy) + selection(rs, cs, ox, oy);
    Q.top.innerHTML = cells(rsF, cs, ox, oy) + selection(rsF, cs, ox, oy);
    Q.left.innerHTML = cells(rs, csF, ox, oy) + selection(rs, csF, ox, oy);
    Q.corner.innerHTML = cells(rsF, csF, ox, oy) + selection(rsF, csF, ox, oy);
    const s = norm(S.sel);
    const colHead = (list) => list.map((c) => `<div class="hc${c >= s.c1 && c <= s.c2 ? ' is-sel' : ''}${g.t.filter && inFilter(c) ? ' is-f' : ''}" style="left:${ox(c)}px;width:${g.X[c + 1] - g.X[c]}px">${colName(c)}</div>`).join('');
    const rowHead = (list) => list.map((r) => `<div class="hr${r >= s.r1 && r <= s.r2 ? ' is-sel' : ''}" style="top:${oy(r)}px;height:${g.rh}px">${r + 1}</div>`).join('');
    Q.ch.innerHTML = colHead(cs);
    Q.chf.innerHTML = colHead(csF);
    Q.rh.innerHTML = rowHead(rs);
    Q.rhf.innerHTML = rowHead(rsF);
    root.querySelector('.ed').classList.toggle('is-dark', getComputedStyle(document.documentElement).colorScheme.includes('dark'));
    viewEl.classList.toggle('has-fr', g.fr > 0);
    viewEl.classList.toggle('has-fc', g.fc > 0);
    Q.top.style.boxShadow = g.fr ? '' : 'none';
    drawCharts();
    placeInput();
    drawStatus();
  }

  // In dark mode, fills are toned down and near-black text uses the theme's ink, so both stay readable.
  const lum = (hex) => { const n = parseInt(String(hex).slice(1), 16); return (0.2126 * ((n >> 16) & 255) + 0.7152 * ((n >> 8) & 255) + 0.0722 * (n & 255)) / 255; };
  const inFilter = (c) => { try { const f = parseRange(S.geo.t.filter.range); return c >= f.c1 && c <= f.c2; } catch { return false; } };

  function cells(rs, cs, ox, oy) {
    if (!rs.length || !cs.length) return '';
    const g = S.geo, t = g.t, e = S.engine;
    const cm = commentCells();
    const filterHead = t.filter ? (() => { try { return parseRange(t.filter.range); } catch { return null; } })() : null;
    let out = '';
    for (const r of rs) {
      const y = oy(r);
      let skipUntil = -1;
      for (let i = 0; i < cs.length; i++) {
        const c = cs[i];
        if (c <= skipUntil) continue;
        const raw = getRaw(S.doc, t.id, r, c);
        const f = getFmt(S.doc, t.id, r, c);
        const v = raw === null && !e.isSpill(t.id, r, c) ? null : e.value(t.id, r, c);
        const type = v === null ? null : e.type(t.id, r, c);
        let text = v === null ? '' : display(v, f?.nf, type);
        const isErr = typeof v === 'object' && v !== null;
        const isNum = typeof v === 'number';
        let w = g.X[c + 1] - g.X[c];
        // Text runs on into empty cells to its right, as in every spreadsheet.
        const align = f?.ha ?? (isNum || isErr ? 'right' : typeof v === 'boolean' ? 'center' : 'left');
        if (text && align === 'left' && !f?.wrap && typeof v === 'string') {
          let j = i + 1;
          while (j < cs.length && getRaw(S.doc, t.id, r, cs[j]) === null && !e.isSpill(t.id, r, cs[j]) && cs[j] === cs[j - 1] + 1 && text.length * 7.4 > w) { w += g.X[cs[j] + 1] - g.X[cs[j]]; skipUntil = cs[j]; j++; }
        }
        let st = `left:${ox(c)}px;top:${y}px;width:${w}px;height:${g.rh}px`;
        if (f) {
          if (f.b) st += ';font-weight:600';
          if (f.i) st += ';font-style:italic';
          if (f.u || f.s) st += `;text-decoration:${[f.u && 'underline', f.s && 'line-through'].filter(Boolean).join(' ')}`;
          if (f.fc) st += `;color:${f.fc}`;
          if (f.bg) st += `;--bg:${f.bg}`;
          if (f.fs) st += `;font-size:${Math.min(28, f.fs * 1.18)}px`;
        }
        const cls = `c${f?.bg ? ' has-bg' : ''}${f?.fc && lum(f.fc) < 0.25 ? ' is-darkink' : ''} a-${align}${isErr ? ' is-err' : ''}${isNum ? ' is-num' : ''}${f?.wrap ? ' is-wrap' : ''}${f?.va ? ` v-${f.va}` : ''}${cm.has(key(r, c)) ? ' has-cm' : ''}`;
        const tip = isErr ? ` title="${h(v.message || v.value)}"` : '';
        const fbtn = filterHead && r === filterHead.r1 && c >= filterHead.c1 && c <= filterHead.c2 ? `<i class="c-f${t.filter.criteria?.[c] ? ' is-on' : ''}"></i>` : '';
        out += `<div class="${cls}" style="${st}"${tip}>${text ? `<span>${h(text)}</span>` : ''}${fbtn}</div>`;
      }
    }
    return out;
  }

  function commentCells() {
    const m = new Set();
    for (const c of allComments(S.doc)) if (c.tab === S.tab && !c.resolved) m.add(key(c.r, c.c));
    return m;
  }

  const norm = (s) => ({ r1: Math.min(s.r, s.r2), r2: Math.max(s.r, s.r2), c1: Math.min(s.c, s.c2), c2: Math.max(s.c, s.c2) });

  function selection(rs, cs, ox, oy) {
    if (!rs.length || !cs.length) return '';
    const g = S.geo;
    const s = norm(S.sel);
    const box = (r1, c1, r2, c2, cls, extra = '') => {
      const vr = rs.filter((r) => r >= r1 && r <= r2), vc = cs.filter((c) => c >= c1 && c <= c2);
      if (!vr.length || !vc.length) return '';
      const x = ox(vc[0]), y = oy(vr[0]);
      const w = ox(vc[vc.length - 1]) + (g.X[vc[vc.length - 1] + 1] - g.X[vc[vc.length - 1]]) - x;
      const hh = oy(vr[vr.length - 1]) + g.rh - y;
      return `<div class="${cls}" style="left:${x}px;top:${y}px;width:${w}px;height:${hh}px">${extra}</div>`;
    };
    let out = '';
    for (const [i, ref] of S.refs.entries()) out += box(ref.r1, ref.c1, ref.r2, ref.c2, `sel-ref r${i % 6}`);
    for (const o of S.others.values()) {
      if (o.tab !== S.tab || !o.cell) continue;
      try { const p = parseRange(o.cell); out += box(p.r1, p.c1, p.r2, p.c2, 'sel-other', `<b>${h(o.person?.name ?? '')}</b>`); } catch {}
    }
    if (s.r1 !== s.r2 || s.c1 !== s.c2) out += box(s.r1, s.c1, s.r2, s.c2, 'sel-range');
    out += box(S.sel.r, S.sel.c, S.sel.r, S.sel.c, 'sel-cell');
    if (!readOnly && !S.editing && rs.includes(s.r2) && cs.includes(s.c2)) out += `<div class="sel-fill" style="left:${ox(s.c2) + (g.X[s.c2 + 1] - g.X[s.c2]) - 4}px;top:${oy(s.r2) + g.rh - 4}px"></div>`;
    return out;
  }

  // ---------- charts ----------

  function drawCharts() {
    const box = $('.sg-charts');
    const list = allCharts(S.doc).filter((c) => c.tab === S.tab);
    const g = S.geo;
    const seen = new Set();
    for (const c of list) {
      seen.add(c.id);
      let el = box.querySelector(`[data-chart="${c.id}"]`);
      const sig = JSON.stringify(c) + S.engine.version + (phone() ? 'p' : 'd');
      const w = Math.min(c.w, phone() ? root.clientWidth - 24 : 2000), hh = c.h;
      if (!el) {
        el = document.createElement('figure');
        el.className = 'sg-chart';
        el.dataset.chart = c.id;
        box.appendChild(el);
      }
      if (el.dataset.sig !== sig) {
        el.dataset.sig = sig;
        const data = ops.chartData(S.doc, S.engine, c);
        el.innerHTML = `<figcaption><span>${h(c.title ?? 'Chart')}</span>${readOnly ? '' : `<button type="button" class="ui-btn is-ghost is-sm" data-chart-menu="${c.id}" ${none('opens the chart menu')} aria-label="Chart options">${I.more}</button>`}</figcaption><div class="sg-chart-b">${renderChart({ ...c, data }, w - 2, hh - 44)}</div>${readOnly ? '' : '<span class="sg-chart-rs" aria-hidden="true"></span>'}`;
        el.style.width = `${w}px`;
        el.style.height = `${hh}px`;
      }
      const x = HW + (c.c < g.fc ? g.X[c.c] : g.X[c.c] - S.sx);
      const y = HH + (c.r < g.fr ? ry(c.r) : ry(Math.max(c.r, 0)) - S.sy);
      el.style.transform = `translate(${x + 4}px, ${y + 4}px)`;
      const hiddenUnder = x + w < HW + g.frozenW || y + hh < HH + g.frozenH;
      el.style.visibility = hiddenUnder ? 'hidden' : '';
    }
    for (const el of [...box.children]) if (!seen.has(el.dataset.chart)) el.remove();
  }

  // ---------- selection and the cell editor ----------

  function selectCell(r, c, { extend = false, keep = false } = {}) {
    const g = S.geo;
    r = Math.max(0, Math.min(g.rows - 1, r));
    c = Math.max(0, Math.min(g.cols - 1, c));
    if (extend) { S.sel = { r: S.anchor.r, c: S.anchor.c, r2: r, c2: c }; }
    else { S.sel = { r, c, r2: r, c2: c }; S.anchor = { r, c }; }
    if (!keep) reveal(extend ? r : S.sel.r, extend ? c : S.sel.c);
    syncBars();
    draw();
    presence();
    showCommentHint();
  }

  function reveal(r, c) {
    const g = S.geo;
    const vw = scroll.clientWidth - HW - g.frozenW, vh = scroll.clientHeight - HH - g.frozenH;
    if (c >= g.fc) {
      const x1 = g.X[c] - g.frozenW, x2 = g.X[c + 1] - g.frozenW;
      if (x1 < scroll.scrollLeft) scroll.scrollLeft = x1;
      else if (x2 > scroll.scrollLeft + vw) scroll.scrollLeft = x2 - vw;
    }
    if (r >= g.fr && ry(r) >= 0) {
      const y1 = ry(r) - g.frozenH, y2 = y1 + g.rh;
      if (y1 < scroll.scrollTop) scroll.scrollTop = y1;
      else if (y2 > scroll.scrollTop + vh) scroll.scrollTop = y2 - vh;
    }
    S.sx = scroll.scrollLeft; S.sy = scroll.scrollTop;
  }

  function syncBars() {
    const s = norm(S.sel);
    const nameIn = $('.ed-name input');
    if (document.activeElement !== nameIn) nameIn.value = s.r1 === s.r2 && s.c1 === s.c2 ? addr(s.r1, s.c1) : rangeText(s);
    const fx = $('.ed-formula input');
    if (S.editing) fx.value = input.value;
    else if (document.activeElement !== fx) fx.value = editText(getRaw(S.doc, S.tab, S.sel.r, S.sel.c), getFmt(S.doc, S.tab, S.sel.r, S.sel.c)?.nf);
    $('.ed-title input').value = document.activeElement === $('.ed-title input') ? $('.ed-title input').value : S.doc.getMap('meta').get('title') ?? '';
    const n = allComments(S.doc).filter((c) => !c.resolved).length;
    const badge = root.querySelector('.ed-n');
    if (badge) { badge.hidden = !n; badge.textContent = n; }
    updateToolbarState();
  }

  function placeInput() {
    const r = rect(S.sel.r, S.sel.c);
    input.style.transform = `translate(${r.x}px, ${r.y}px)`;
    input.style.minWidth = `${r.w}px`;
    input.style.height = `${r.h}px`;
    if (S.editing) {
      const f = getFmt(S.doc, S.tab, S.sel.r, S.sel.c);
      input.style.fontWeight = f?.b ? '600' : '';
      input.style.fontStyle = f?.i ? 'italic' : '';
    }
  }

  function focusGrid() { if (!S.editing) { input.value = ''; } input.focus({ preventScroll: true }); }

  function startEdit(text = null, { caretEnd = true } = {}) {
    if (readOnly) return;
    const raw = getRaw(S.doc, S.tab, S.sel.r, S.sel.c);
    S.editing = { r: S.sel.r, c: S.sel.c, tab: S.tab, typed: text !== null };
    input.value = text ?? editText(raw, getFmt(S.doc, S.tab, S.sel.r, S.sel.c)?.nf);
    input.classList.add('is-on');
    placeInput();
    input.focus({ preventScroll: true });
    if (caretEnd) input.setSelectionRange(input.value.length, input.value.length);
    grow();
    onEditInput();
    draw();
  }

  function grow() {
    input.style.width = 'auto';
    const r = rect(S.sel.r, S.sel.c);
    input.style.width = `${Math.max(r.w, Math.min(input.scrollWidth + 12, scroll.clientWidth - r.x - 8))}px`;
  }

  async function commit(move = null) {
    const ed = S.editing;
    if (!ed) return;
    const text = input.value;
    S.editing = null;
    S.refs = [];
    input.classList.remove('is-on');
    ac.hidden = true;
    input.value = '';
    const before = getRaw(S.doc, ed.tab, ed.r, ed.c);
    const beforeText = editText(before, getFmt(S.doc, ed.tab, ed.r, ed.c)?.nf);
    if (move) moveSel(...move);
    draw();
    if (text === beforeText) return;
    await writeCells(ed.tab, ed.r, ed.c, [[text === '' ? null : text]]);
  }

  function cancelEdit() {
    S.editing = null; S.refs = [];
    input.classList.remove('is-on');
    ac.hidden = true;
    input.value = '';
    syncBars(); draw();
  }

  // Writes typed cells: shows them at once in this copy's engine, then sends sheets.write_range.
  async function writeCells(tabId, r, c, values, { input: mode = 'user', formats = null, label = 'Edit' } = {}) {
    const t = findTab(S.doc, tabId);
    const h1 = values.length, w1 = Math.max(...values.map((x) => x.length));
    const undo = snapshot(tabId, r, c, r + h1 - 1, c + w1 - 1, label);
    try {
      const sid = S.engine.sheet(tabId);
      S.engine.hf.batch(() => values.forEach((row, i) => row.forEach((v, j) => {
        if (v === undefined) return;
        const raw = mode === 'exact' ? v : parseInput(v, { nf: getFmt(S.doc, tabId, r + i, c + j)?.nf }).raw;
        S.engine.hf.setCellContents({ sheet: sid, row: r + i, col: c + j }, [[toHf(raw)]]);
      })));
      draw();
    } catch {}
    await run('sheets.write_range', { tab: t.id, range: addr(r, c), values, input: mode === 'exact' ? 'exact' : 'user', ...(formats ? { formats } : {}) }, { undo });
  }

  // What a range holds now, so the change can be taken back.
  function snapshot(tabId, r1, c1, r2, c2, label) {
    const raws = [], fmts = [];
    for (let r = r1; r <= r2; r++) {
      const a = [], b = [];
      for (let c = c1; c <= c2; c++) { a.push(getRaw(S.doc, tabId, r, c)); b.push(getFmt(S.doc, tabId, r, c)); }
      raws.push(a); fmts.push(b);
    }
    return { kind: 'cells', tab: tabId, r: r1, c: c1, raws, fmts, label };
  }

  async function undoRedo(redo = false) {
    const from = redo ? S.redo : S.undo, to = redo ? S.undo : S.redo;
    const u = from.pop();
    if (!u) return toast(redo ? 'Nothing to redo.' : 'Nothing to undo. Version history has older states.');
    if (!findTab(S.doc, u.tab)) return toast('That tab is gone; use version history.');
    const back = snapshot(u.tab, u.r, u.c, u.r + u.raws.length - 1, u.c + u.raws[0].length - 1, u.label);
    to.push(back);
    if (S.tab !== u.tab) { S.tab = u.tab; layout(); }
    S.sel = { r: u.r, c: u.c, r2: u.r + u.raws.length - 1, c2: u.c + u.raws[0].length - 1 };
    try {
      await callTool('sheets.write_range', { tab: u.tab, range: addr(u.r, u.c), values: u.raws, input: 'exact', formats: u.fmts });
      await sync();
      toast(`${redo ? 'Redid' : 'Undid'}: ${u.label.toLowerCase()}`);
    } catch (e) { toast(e.message); }
  }

  function moveSel(dr, dc, extend = false) {
    const g = S.geo;
    const from = extend ? { r: S.sel.r2, c: S.sel.c2 } : { r: S.sel.r, c: S.sel.c };
    let r = from.r, c = from.c;
    const stepRow = (d) => { let x = r + d; while (x >= 0 && x < g.rows && g.pos[x] < 0) x += d; return x < 0 || x >= g.rows ? r : x; };
    if (dr) r = stepRow(Math.sign(dr));
    c = Math.max(0, Math.min(g.cols - 1, c + dc));
    selectCell(r, c, { extend });
  }

  // Ctrl or Cmd with an arrow: to the edge of the data, as in Sheets and Excel.
  function jump(dr, dc, extend) {
    const g = S.geo;
    let { r, c } = extend ? { r: S.sel.r2, c: S.sel.c2 } : S.sel;
    const has = (rr, cc) => rr >= 0 && cc >= 0 && rr < g.rows && cc < g.cols && getRaw(S.doc, S.tab, rr, cc) !== null;
    const filled = has(r, c) && has(r + dr, c + dc);
    do { r += dr; c += dc; } while (r >= 0 && c >= 0 && r < g.rows && c < g.cols && (filled ? has(r + dr, c + dc) : !has(r, c)));
    selectCell(Math.max(0, Math.min(g.rows - 1, r)), Math.max(0, Math.min(g.cols - 1, c)), { extend });
  }

  // Formula help: function names as you type, and coloured boxes round the ranges a formula uses.
  const FN = () => (S.fnNames ??= S.engine.hf.getRegisteredFunctionNames().sort());
  function onEditInput() {
    const v = input.value;
    if (S.editing) $('.ed-formula input').value = v;
    S.refs = [];
    if (v.startsWith('=')) {
      for (const m of v.matchAll(/(?:(?:'[^']+'|[A-Za-z_][\w.]*)!)?\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?/g)) {
        try { const p = parseRange(m[0]); const t = p.tab ? findTab(S.doc, p.tab) : tab(); if (t?.id === S.tab) S.refs.push(p); } catch {}
      }
      const word = /([A-Za-z][A-Za-z0-9.]*)$/.exec(v.slice(0, input.selectionStart ?? v.length))?.[1];
      const list = word && word.length >= 2 ? FN().filter((n) => n.startsWith(word.toUpperCase())).slice(0, 7) : [];
      if (list.length && !(list.length === 1 && list[0] === word.toUpperCase())) {
        ac.innerHTML = list.map((n, i) => `<div role="option" class="${i ? '' : 'is-on'}" data-fn="${h(n)}">${h(n)}<span>(</span></div>`).join('');
        ac.hidden = false;
        const r = rect(S.sel.r, S.sel.c);
        ac.style.transform = `translate(${r.x}px, ${r.y + r.h + 2}px)`;
      } else ac.hidden = true;
    } else ac.hidden = true;
    grow();
    draw();
  }
  function acceptFn(name) {
    const pos = input.selectionStart ?? input.value.length;
    const before = input.value.slice(0, pos).replace(/[A-Za-z][A-Za-z0-9.]*$/, `${name}(`);
    input.value = before + input.value.slice(pos);
    input.setSelectionRange(before.length, before.length);
    ac.hidden = true;
    onEditInput();
  }
  ac.addEventListener('pointerdown', (e) => { const o = e.target.closest('[data-fn]'); if (o) { e.preventDefault(); acceptFn(o.dataset.fn); } });

  // While writing a formula, clicking a cell (or dragging a range) puts its address in.
  const pickingRef = () => S.editing && input.value.startsWith('=') && /[=(,+\-*/^&<>:;\s]$/.test(input.value.slice(0, input.selectionStart ?? 0));

  // ---------- pointer ----------

  let drag = null;
  scroll.addEventListener('pointerdown', (e) => {
    if (e.button === 2) return;
    const b = grid.getBoundingClientRect();
    const px = e.clientX - b.left, py = e.clientY - b.top;
    if (px > scroll.clientWidth || py > scroll.clientHeight) return;
    const x = hit(px, py);
    closePop();
    if (e.pointerType === 'touch') { drag = { touch: true, x: e.clientX, y: e.clientY, at: x }; return; }
    // Fill handle.
    const s = norm(S.sel);
    const fr = rect(s.r2, s.c2);
    if (!readOnly && x.kind === 'cell' && Math.abs(px - (fr.x + fr.w)) < 6 && Math.abs(py - (fr.y + fr.h)) < 6) {
      drag = { kind: 'fill', from: s, to: { ...s } };
      scroll.setPointerCapture(e.pointerId);
      e.preventDefault();
      return;
    }
    if (x.kind === 'col' && x.edge !== null && !readOnly) {
      drag = { kind: 'resize', c: x.edge, x0: px, w0: S.geo.X[x.edge + 1] - S.geo.X[x.edge] };
      scroll.setPointerCapture(e.pointerId);
      e.preventDefault();
      return;
    }
    if (pickingRef() && x.kind === 'cell') {
      e.preventDefault();
      drag = { kind: 'ref', start: { r: x.r, c: x.c }, pos: input.selectionStart, len: 0 };
      insertRef(drag, x);
      scroll.setPointerCapture(e.pointerId);
      return;
    }
    if (S.editing) commit();
    if (x.kind === 'all') { S.sel = { r: 0, c: 0, r2: S.geo.rows - 1, c2: S.geo.cols - 1 }; syncBars(); draw(); return; }
    if (x.kind === 'col') { S.anchor = { r: 0, c: e.shiftKey ? S.anchor.c : x.c }; S.sel = { r: 0, c: S.anchor.c, r2: S.geo.rows - 1, c2: x.c }; drag = { kind: 'cols' }; }
    else if (x.kind === 'row') { S.anchor = { r: e.shiftKey ? S.anchor.r : x.r, c: 0 }; S.sel = { r: S.anchor.r, c: 0, r2: x.r, c2: S.geo.cols - 1 }; drag = { kind: 'rows' }; }
    else {
      // A filter button in the header row.
      if (S.geo.t.filter && onFilterButton(x, px)) { e.preventDefault(); return openFilter(x.c); }
      selectCell(x.r, x.c, { extend: e.shiftKey, keep: true });
      drag = { kind: 'cells' };
      if (e.detail === 2 && !readOnly) { startEdit(); drag = null; return; }
    }
    scroll.setPointerCapture(e.pointerId);
    syncBars(); draw();
    if (!readOnly) setTimeout(focusGrid, 0);
  });
  scroll.addEventListener('pointermove', (e) => {
    const b = grid.getBoundingClientRect();
    const px = e.clientX - b.left, py = e.clientY - b.top;
    if (!drag) {
      const x = hit(px, py);
      scroll.style.cursor = !readOnly && x.kind === 'col' && x.edge !== null ? 'col-resize' : x.kind === 'cell' ? 'cell' : '';
      return;
    }
    if (drag.touch) return;
    const x = hit(Math.max(HW + 1, Math.min(px, scroll.clientWidth - 1)), Math.max(HH + 1, Math.min(py, scroll.clientHeight - 1)));
    if (drag.kind === 'resize') {
      const w = Math.max(24, Math.round(drag.w0 + px - drag.x0));
      drag.w = w;
      S.geo.t.widths[drag.c] = w;
      const X = S.geo.X;
      const d = w - (X[drag.c + 1] - X[drag.c]);
      for (let c = drag.c + 1; c < X.length; c++) X[c] += d;
      draw();
    } else if (drag.kind === 'cells') { S.sel.r2 = x.r; S.sel.c2 = x.c; edgeScroll(px, py); syncBars(); draw(); }
    else if (drag.kind === 'cols') { S.sel.c2 = x.c; draw(); }
    else if (drag.kind === 'rows') { S.sel.r2 = x.r; draw(); }
    else if (drag.kind === 'ref') insertRef(drag, x);
    else if (drag.kind === 'fill') {
      const f = drag.from;
      const down = Math.abs(x.r - f.r2) >= Math.abs(x.c - f.c2);
      drag.to = down ? { ...f, r2: Math.max(f.r2, x.r), r1: Math.min(f.r1, x.r) } : { ...f, c2: Math.max(f.c2, x.c), c1: Math.min(f.c1, x.c) };
      S.sel = { r: drag.to.r1, c: drag.to.c1, r2: drag.to.r2, c2: drag.to.c2 };
      draw();
    }
  });
  scroll.addEventListener('pointerup', async (e) => {
    const d = drag;
    drag = null;
    if (!d) return;
    if (d.touch) {
      // A tap selects; a tap on the selected cell edits (the keyboard comes up then, not before).
      if (Math.hypot(e.clientX - d.x, e.clientY - d.y) > 8) return;
      const x = d.at;
      if (x.kind !== 'cell') { if (x.kind === 'col') { S.sel = { r: 0, c: x.c, r2: S.geo.rows - 1, c2: x.c }; } else if (x.kind === 'row') { S.sel = { r: x.r, c: 0, r2: x.r, c2: S.geo.cols - 1 }; } syncBars(); draw(); return; }
      if (S.geo.t.filter && onFilterButton(x, e.clientX - grid.getBoundingClientRect().left)) return openFilter(x.c);
      if (S.editing) await commit();
      if (!readOnly && S.sel.r === x.r && S.sel.c === x.c && S.sel.r2 === x.r && S.sel.c2 === x.c) startEdit();
      else selectCell(x.r, x.c, { keep: true });
      return;
    }
    if (d.kind === 'resize' && d.w) await run('sheets.set_column_width', { tab: S.tab, columns: colName(d.c), width: d.w }, { quiet: true });
    if (d.kind === 'fill') await fill(d.from, d.to);
    if (d.kind === 'ref') input.focus({ preventScroll: true });
    else if (!readOnly && !S.editing && e.pointerType !== 'touch') focusGrid();
  });
  scroll.addEventListener('scroll', () => { closePop(); draw(); }, { passive: true });
  scroll.addEventListener('contextmenu', (e) => {
    e.preventDefault();
    const b = grid.getBoundingClientRect();
    const x = hit(e.clientX - b.left, e.clientY - b.top);
    const s = norm(S.sel);
    const inside = x.r >= s.r1 && x.r <= s.r2 && x.c >= s.c1 && x.c <= s.c2;
    if (!inside) { if (x.kind === 'col') S.sel = { r: 0, c: x.c, r2: S.geo.rows - 1, c2: x.c }; else if (x.kind === 'row') S.sel = { r: x.r, c: 0, r2: x.r, c2: S.geo.cols - 1 }; else selectCell(x.r, x.c, { keep: true }); draw(); }
    cellMenu({ x: e.clientX, y: e.clientY }, x.kind);
  });
  let edgeTimer = null;
  function edgeScroll(px, py) {
    clearTimeout(edgeTimer);
    const dx = px > scroll.clientWidth - 20 ? 24 : px < HW + 8 ? -24 : 0;
    const dy = py > scroll.clientHeight - 20 ? 24 : py < HH + 8 ? -24 : 0;
    if (dx || dy) { scroll.scrollLeft += dx; scroll.scrollTop += dy; }
  }
  function insertRef(d, x) {
    const p = { r1: Math.min(d.start.r, x.r), c1: Math.min(d.start.c, x.c), r2: Math.max(d.start.r, x.r), c2: Math.max(d.start.c, x.c) };
    const t = tab();
    const txt = (S.editing.tab !== t.id ? `${t.name.includes(' ') ? `'${t.name}'` : t.name}!` : '') + rangeText(p);
    const v = input.value;
    input.value = v.slice(0, d.pos) + txt + v.slice(d.pos + d.len);
    d.len = txt.length;
    input.setSelectionRange(d.pos + d.len, d.pos + d.len);
    onEditInput();
  }
  function onFilterButton(x, px) {
    try {
      const f = parseRange(S.geo.t.filter.range);
      if (x.r !== f.r1 || x.c < f.c1 || x.c > f.c2) return false;
      const r = rect(x.r, x.c);
      return px > r.x + r.w - 22;
    } catch { return false; }
  }

  // ---------- keyboard ----------

  input.addEventListener('keydown', async (e) => {
    const mod = e.metaKey || e.ctrlKey;
    if (S.editing) {
      if (!ac.hidden && (e.key === 'Tab' || e.key === 'Enter')) { e.preventDefault(); acceptFn(ac.querySelector('.is-on')?.dataset.fn); return; }
      if (!ac.hidden && (e.key === 'ArrowDown' || e.key === 'ArrowUp')) {
        e.preventDefault();
        const os = [...ac.children]; const i = os.findIndex((o) => o.classList.contains('is-on'));
        os[i]?.classList.remove('is-on'); os[(i + (e.key === 'ArrowDown' ? 1 : -1) + os.length) % os.length].classList.add('is-on');
        return;
      }
      if (e.key === 'Enter' && !e.altKey) { e.preventDefault(); await commit([e.shiftKey ? -1 : 1, 0]); return; }
      if (e.key === 'Enter' && e.altKey) { e.preventDefault(); const p = input.selectionStart; input.value = `${input.value.slice(0, p)}\n${input.value.slice(input.selectionEnd)}`; input.setSelectionRange(p + 1, p + 1); return; }
      if (e.key === 'Tab') { e.preventDefault(); await commit([0, e.shiftKey ? -1 : 1]); return; }
      if (e.key === 'Escape') { e.preventDefault(); cancelEdit(); return; }
      if (S.editing.typed && ['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(e.key) && !pickingRef()) {
        e.preventDefault();
        await commit({ ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[e.key]);
      }
      return;
    }
    const k = e.key;
    const arrows = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
    if (arrows[k]) { e.preventDefault(); if (mod) jump(...arrows[k], e.shiftKey); else moveSel(...arrows[k], e.shiftKey); return; }
    if (k === 'Tab') { e.preventDefault(); moveSel(0, e.shiftKey ? -1 : 1); return; }
    if (k === 'Enter') { e.preventDefault(); if (readOnly) moveSel(1, 0); else startEdit(); return; }
    if (k === 'F2') { e.preventDefault(); startEdit(); return; }
    if (k === 'Home') { e.preventDefault(); selectCell(mod ? 0 : S.sel.r, 0, { extend: e.shiftKey }); return; }
    if (k === 'End' && mod) { e.preventDefault(); const u = usedRange(S.doc, S.tab); selectCell(u?.r2 ?? 0, u?.c2 ?? 0, { extend: e.shiftKey }); return; }
    if (k === 'PageDown' || k === 'PageUp') { e.preventDefault(); moveSel((k === 'PageDown' ? 1 : -1) * Math.floor(scroll.clientHeight / S.geo.rh - 2), 0, e.shiftKey); return; }
    if (mod && k.toLowerCase() === 'a') { e.preventDefault(); S.sel = { r: 0, c: 0, r2: S.geo.rows - 1, c2: S.geo.cols - 1 }; syncBars(); draw(); return; }
    if (mod && k.toLowerCase() === 'f') { e.preventDefault(); return findDialog(); }
    if (readOnly) return;
    if ((k === 'Delete' || k === 'Backspace') && !mod) { e.preventDefault(); return clearSel(); }
    if (mod && k.toLowerCase() === 'z') { e.preventDefault(); return undoRedo(e.shiftKey); }
    if (mod && k.toLowerCase() === 'y') { e.preventDefault(); return undoRedo(true); }
    if (mod && k.toLowerCase() === 'b') { e.preventDefault(); return toggleFmt('b'); }
    if (mod && k.toLowerCase() === 'i') { e.preventDefault(); return toggleFmt('i'); }
    if (mod && k.toLowerCase() === 'u') { e.preventDefault(); return toggleFmt('u'); }
    if (mod && k.toLowerCase() === 'd') { e.preventDefault(); const s = norm(S.sel); return fill({ ...s, r2: s.r1 }, s); }
    if (mod && k.toLowerCase() === 'r') { e.preventDefault(); const s = norm(S.sel); return fill({ ...s, c2: s.c1 }, s); }
  });
  input.addEventListener('input', () => {
    if (!S.editing && input.value && !readOnly) {
      const v = input.value;
      startEdit(v);
      return;
    }
    if (S.editing) onEditInput();
  });
  input.addEventListener('blur', () => { if (S.editing && !document.activeElement?.closest?.('.ed-formula')) setTimeout(() => { if (S.editing && document.activeElement !== input && !document.activeElement?.closest?.('.ed-formula')) commit(); }, 120); });

  // Copy, cut and paste: plain text (tab-separated) for other apps, and the exact contents for pasting here.
  input.addEventListener('copy', (e) => { if (S.editing) return; e.preventDefault(); copySel(e.clipboardData, false); });
  input.addEventListener('cut', (e) => { if (S.editing || readOnly) return; e.preventDefault(); copySel(e.clipboardData, true); });
  input.addEventListener('paste', (e) => { if (S.editing || readOnly) return; e.preventDefault(); paste(e.clipboardData.getData('text/plain')); });

  function copySel(cd, cut) {
    const s = norm(S.sel);
    const u = usedRange(S.doc, S.tab);
    const r2 = Math.min(s.r2, u?.r2 ?? s.r1), c2 = Math.min(s.c2, u?.c2 ?? s.c1);
    const raws = [], text = [];
    for (let r = s.r1; r <= Math.max(s.r1, r2); r++) {
      const a = [], t = [];
      for (let c = s.c1; c <= Math.max(s.c1, c2); c++) {
        a.push({ raw: getRaw(S.doc, S.tab, r, c), fmt: getFmt(S.doc, S.tab, r, c) });
        const v = S.engine.value(S.tab, r, c);
        t.push(v === null ? '' : display(v, getFmt(S.doc, S.tab, r, c)?.nf, S.engine.type(S.tab, r, c)));
      }
      raws.push(a); text.push(t.join('\t'));
    }
    const tsv = text.join('\n');
    S.clip = { tab: S.tab, r: s.r1, c: s.c1, raws, tsv, cut };
    if (cd) cd.setData('text/plain', tsv); else copyText(tsv);
    toast(cut ? 'Cut. Paste where it should go.' : 'Copied');
  }

  async function paste(text) {
    const s = norm(S.sel);
    const clip = S.clip;
    if (clip && text.replace(/\r/g, '') === clip.tsv) {
      const values = clip.raws.map((row, i) => row.map((x, j) => (typeof x.raw === 'string' && x.raw.startsWith('=') ? shiftFormula(x.raw, s.r1 - clip.r + 0 * i, s.c1 - clip.c + 0 * j) : x.raw)));
      const formats = clip.raws.map((row) => row.map((x) => x.fmt));
      await writeCells(S.tab, s.r1, s.c1, values, { input: 'exact', formats, label: 'Paste' });
      if (clip.cut) {
        await run('sheets.clear_range', { tab: clip.tab, range: rangeText({ r1: clip.r, c1: clip.c, r2: clip.r + clip.raws.length - 1, c2: clip.c + clip.raws[0].length - 1 }), what: 'all' });
        S.clip = null;
      }
    } else {
      const rows = text.replace(/\r\n?/g, '\n').replace(/\n$/, '').split('\n').map((l) => l.split('\t'));
      if (!rows.length) return;
      await writeCells(S.tab, s.r1, s.c1, rows, { label: 'Paste' });
    }
    S.sel = { r: s.r1, c: s.c1, r2: s.r1 + (clip?.raws?.length ?? 1) - 1, c2: s.c1 + (clip?.raws?.[0]?.length ?? 1) - 1 };
    draw();
  }

  async function clearSel() {
    const s = norm(S.sel);
    const u = usedRange(S.doc, S.tab);
    if (!u) return;
    const r2 = Math.min(s.r2, u.r2), c2 = Math.min(s.c2, u.c2);
    if (r2 < s.r1 || c2 < s.c1) return;
    const values = Array.from({ length: r2 - s.r1 + 1 }, () => new Array(c2 - s.c1 + 1).fill(null));
    await writeCells(S.tab, s.r1, s.c1, values, { input: 'exact', label: 'Clear' });
  }

  // The fill handle and Ctrl+D: a series when the cells count up evenly, otherwise the pattern repeated,
  // with formulas moved as they go.
  async function fill(from, to) {
    if (to.r1 === from.r1 && to.r2 === from.r2 && to.c1 === from.c1 && to.c2 === from.c2) return;
    const down = to.c1 === from.c1 && to.c2 === from.c2;
    // Only the new cells are written (the source stays as it is).
    const reg = down
      ? (to.r2 > from.r2 ? { r1: from.r2 + 1, r2: to.r2 } : { r1: to.r1, r2: from.r1 - 1 })
      : (to.c2 > from.c2 ? { c1: from.c2 + 1, c2: to.c2 } : { c1: to.c1, c2: from.c1 - 1 });
    const R = down ? { r1: reg.r1, r2: reg.r2, c1: to.c1, c2: to.c2 } : { r1: to.r1, r2: to.r2, c1: reg.c1, c2: reg.c2 };
    const n = down ? from.r2 - from.r1 + 1 : from.c2 - from.c1 + 1;
    const src = (r, c) => {
      const off = down ? r - from.r1 : c - from.c1;
      const k = ((off % n) + n) % n;
      return { sr: down ? from.r1 + k : r, sc: down ? c : from.c1 + k, off };
    };
    const values = [], formats = [];
    for (let r = R.r1; r <= R.r2; r++) {
      const row = [], frow = [];
      for (let c = R.c1; c <= R.c2; c++) {
        const { sr, sc, off } = src(r, c);
        const raw = getRaw(S.doc, S.tab, sr, sc);
        const seq = series(down ? c : r, down);
        if (seq && typeof raw === 'number') row.push(seq.start + seq.step * off);
        else row.push(typeof raw === 'string' && raw.startsWith('=') ? shiftFormula(raw, r - sr, c - sc) : raw);
        frow.push(getFmt(S.doc, S.tab, sr, sc));
      }
      values.push(row); formats.push(frow);
    }
    S.sel = { r: to.r1, c: to.c1, r2: to.r2, c2: to.c2 };
    await writeCells(S.tab, R.r1, R.c1, values, { input: 'exact', formats, label: 'Fill' });
    function series(line, isDown) {
      const vals = [];
      if (isDown) for (let r = from.r1; r <= from.r2; r++) vals.push(getRaw(S.doc, S.tab, r, line));
      else for (let c = from.c1; c <= from.c2; c++) vals.push(getRaw(S.doc, S.tab, line, c));
      if (vals.length < 2 || !vals.every((v) => typeof v === 'number')) return null;
      const step = vals[1] - vals[0];
      if (!vals.every((v, i) => Math.abs(v - (vals[0] + step * i)) < 1e-9)) return null;
      return { start: vals[0], step };
    }
  }

  // ---------- formula bar and name box ----------

  $('.ed-formula').addEventListener('submit', async (e) => {
    e.preventDefault();
    if (readOnly) return;
    const v = $('.ed-formula input').value;
    if (!S.editing) S.editing = { r: S.sel.r, c: S.sel.c, tab: S.tab };
    input.value = v;
    await commit([1, 0]);
    focusGrid();
  });
  $('.ed-formula input').addEventListener('input', () => {
    if (readOnly) return;
    if (!S.editing) { S.editing = { r: S.sel.r, c: S.sel.c, tab: S.tab }; input.classList.add('is-on'); }
    input.value = $('.ed-formula input').value;
    placeInput(); grow();
    S.refs = [];
    if (input.value.startsWith('=')) for (const m of input.value.matchAll(/\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?/g)) { try { S.refs.push(parseRange(m[0])); } catch {} }
    draw();
  });
  $('.ed-formula input').addEventListener('keydown', (e) => { if (e.key === 'Escape') { cancelEdit(); focusGrid(); } });
  $('.ed-name').addEventListener('submit', (e) => {
    e.preventDefault();
    const v = $('.ed-name input').value.trim();
    try {
      const p = parseRange(v);
      if (p.tab) { const t = findTab(S.doc, p.tab); if (t) { S.tab = t.id; layout(); } }
      S.anchor = { r: p.r1, c: p.c1 };
      selectCell(p.r1, p.c1);
      if (p.r2 !== p.r1 || p.c2 !== p.c1) { S.sel = { r: p.r1, c: p.c1, r2: Math.min(p.r2, S.geo.rows - 1), c2: Math.min(p.c2, S.geo.cols - 1) }; draw(); }
      focusGrid();
    } catch { toast('Type a cell like B12 or a range like A1:D20.'); }
  });
  $('.ed-title').addEventListener('submit', async (e) => {
    e.preventDefault();
    const v = $('.ed-title input').value.trim();
    if (!v || v === S.doc.getMap('meta').get('title')) return;
    await run('sheets.rename_sheet', { title: v });
    $('.ed-title input').blur();
  });
  $('.ed-title input').addEventListener('blur', () => { const f = $('.ed-title'); if (!readOnly && f.querySelector('input').value.trim() !== (S.doc.getMap('meta').get('title') ?? '')) f.requestSubmit(); });

  // ---------- the status line: sum, average and count of a selection ----------

  function drawStatus() {
    const el = $('.ed-stat');
    const s = norm(S.sel);
    const u = usedRange(S.doc, S.tab);
    if (!u || (s.r1 === s.r2 && s.c1 === s.c2)) { el.textContent = ''; return; }
    let sum = 0, n = 0, count = 0, cells = 0;
    let nf = null;
    for (let r = s.r1; r <= Math.min(s.r2, u.r2); r++) {
      if (S.hidden.has(r)) continue;
      for (let c = s.c1; c <= Math.min(s.c2, u.c2); c++) {
        if (++cells > 50000) break;
        const v = plainValue(S.engine.value(S.tab, r, c));
        if (v === null || v === '') continue;
        count++;
        if (typeof v === 'number') { sum += v; n++; nf ??= getFmt(S.doc, S.tab, r, c)?.nf ?? null; }
      }
    }
    const f = nf && !isDateFormat(nf) ? nf : null;
    el.innerHTML = n ? `<span>Sum <b>${h(display(sum, f))}</b></span><span>Average <b>${h(display(sum / n, f))}</b></span><span>Count <b>${count}</b></span>` : count ? `<span>Count <b>${count}</b></span>` : '';
  }

  // ---------- tabs ----------

  function drawTabs() {
    const nav = $('.ed-tabs');
    const tabs = tabList(S.doc);
    nav.innerHTML = `${readOnly ? '' : `<button type="button" class="ed-tab-add" ${tool('sheets.add_sheet_tab')} aria-label="Add a tab" title="Add a tab">${I.plus}</button>`}${tabs.map((t) => `<span class="ed-tab${t.id === S.tab ? ' is-on' : ''}"><button type="button" data-tab="${t.id}" ${none('shows this tab')} aria-current="${t.id === S.tab}">${t.link ? `<i class="ed-tab-l" title="${t.link.kind === 'crm' ? 'Linked to the CRM' : 'Pivot table'}">${t.link.kind === 'crm' ? I.crm : I.pivot}</i>` : ''}${h(t.name)}</button>${readOnly ? '' : `<button type="button" class="ed-tab-m" data-tab-menu="${t.id}" ${none('opens the tab menu')} aria-label="${h(t.name)} options">${I.caret}</button>`}</span>`).join('')}`;
    nav.querySelector('.is-on')?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  }
  $('.ed-tabs').addEventListener('click', async (e) => {
    const b = e.target.closest('button');
    if (!b) return;
    if (b.dataset.tab) { if (S.editing && !pickingRef()) await commit(); S.tab = b.dataset.tab; scroll.scrollTo(0, 0); layout(); selectCell(0, 0); return; }
    if (b.classList.contains('ed-tab-add')) {
      const out = await run('sheets.add_sheet_tab', {});
      S.tab = out.tab.id; layout(); selectCell(0, 0);
      return;
    }
    if (b.dataset.tabMenu) tabMenu(b, b.dataset.tabMenu);
  });
  function tabMenu(at, id) {
    const t = findTab(S.doc, id);
    const tabs = tabList(S.doc);
    const i = tabs.findIndex((x) => x.id === id);
    menu(document.body, at, [
      { label: 'Rename', tool: 'sheets.rename_tab', icon: I.clear, run: () => dialog(document.body, { title: 'Rename tab', toolName: 'sheets.rename_tab', body: `<label class="ui-field"><span>Name</span><input class="ui-input" name="name" value="${h(t.name)}" required maxlength="100"></label><p class="ui-hint">Formulas on other tabs that use this one are updated.</p>`, onSubmit: async (v) => { await run('sheets.rename_tab', { tab: id, name: v.name }); } }) },
      t.link && { label: t.link.kind === 'crm' ? 'Refresh from the CRM' : 'Rebuild the pivot', tool: 'sheets.refresh_link', icon: I.refresh, run: () => refreshTab(id) },
      { label: 'Move left', tool: 'sheets.move_tab', icon: I.back, disabled: i === 0, run: () => run('sheets.move_tab', { tab: id, index: i - 1 }) },
      { label: 'Move right', tool: 'sheets.move_tab', icon: `<span style="transform:scaleX(-1);display:inline-flex">${I.back}</span>`, disabled: i === tabs.length - 1, run: () => run('sheets.move_tab', { tab: id, index: i + 1 }) },
      { sep: true },
      { label: 'Delete tab', tool: 'sheets.delete_tab', icon: I.x, danger: true, disabled: tabs.length === 1, run: () => dialog(document.body, { title: `Delete ${t.name}?`, toolName: 'sheets.delete_tab', submit: 'Delete tab', body: '<p>Everything on it goes. Formulas that use it show #REF!. Version history keeps a copy.</p>', onSubmit: async () => { await run('sheets.delete_tab', { tab: id }); S.tab = tabList(S.doc)[0]?.id; layout(); } }) },
    ], { label: 'Tab' });
  }
  async function refreshTab(id) {
    toast('Refreshing…');
    const out = await run('sheets.refresh_link', { tab: id });
    const r = out?.refreshed?.[0];
    if (r) toast(r.kind === 'crm' ? `${r.rows} records from the CRM` : `Pivot rebuilt: ${r.rows} rows`);
  }

  // ---------- toolbar ----------

  function buildToolbar() {
    const b = (act, icon, label, t, extra = '') => `<button type="button" class="tb" data-act="${act}" ${tool(t)} aria-label="${h(label)}" title="${h(label)}"${extra}>${icon}</button>`;
    const sep = '<span class="tb-sep" aria-hidden="true"></span>';
    $('.ed-tools').innerHTML = [
      b('undo', I.undo, 'Undo (Ctrl+Z)', 'sheets.write_range'), b('redo', I.redo, 'Redo (Ctrl+Shift+Z)', 'sheets.write_range'), sep,
      `<select class="tb-sel" data-act="nf" ${tool('sheets.format_range')} aria-label="Number format">${FORMATS.map((f) => `<option value="${f.id}">${h(f.label)}</option>`).join('')}<option value="custom" hidden>Custom</option></select>`,
      b('cur', I.dollar, 'Currency', 'sheets.format_range'), b('pct', I.percent, 'Percent', 'sheets.format_range'), b('decl', I.decLess, 'Fewer decimals', 'sheets.format_range', ' data-wide'), b('decm', I.decMore, 'More decimals', 'sheets.format_range', ' data-wide'), sep,
      b('b', I.bold, 'Bold (Ctrl+B)', 'sheets.format_range'), b('i', I.italic, 'Italic (Ctrl+I)', 'sheets.format_range'), b('u', I.underline, 'Underline (Ctrl+U)', 'sheets.format_range'), b('s', I.strike, 'Strikethrough', 'sheets.format_range'),
      b('color', I.color, 'Text colour', 'sheets.format_range'), b('fill', I.fill, 'Fill colour', 'sheets.format_range'), sep,
      b('left', I.left, 'Align left', 'sheets.format_range'), b('center', I.center, 'Align centre', 'sheets.format_range'), b('right', I.right, 'Align right', 'sheets.format_range'), b('wrap', I.wrap, 'Wrap text', 'sheets.format_range'), sep,
      b('freeze', I.freeze, 'Freeze rows and columns', 'sheets.freeze'), b('sortUp', I.sortUp, 'Sort A to Z by this column', 'sheets.sort'), b('sortDown', I.sortDown, 'Sort Z to A by this column', 'sheets.sort'), b('filter', I.filter, 'Filter', 'sheets.filter'), sep,
      b('chart', I.chart, 'Add a chart', 'sheets.create_chart'), b('pivot', I.pivot, 'Pivot table', 'sheets.create_pivot'), b('comment', I.comment, 'Comment', 'sheets.add_comment'), b('crm', I.crm, 'Pull from the CRM', 'sheets.import_from_crm'), sep,
      b('clearfmt', I.clear, 'Clear formatting', 'sheets.format_range'), b('find', I.search, 'Find (Ctrl+F)', 'sheets.find'),
    ].join('');
  }
  function updateToolbarState() {
    if (readOnly) return;
    const f = getFmt(S.doc, S.tab, S.sel.r, S.sel.c) ?? {};
    for (const k of ['b', 'i', 'u', 's', 'wrap']) root.querySelector(`.ed-tools [data-act="${k}"]`)?.setAttribute('aria-pressed', String(!!f[k]));
    for (const k of ['left', 'center', 'right']) root.querySelector(`.ed-tools [data-act="${k}"]`)?.setAttribute('aria-pressed', String(f.ha === k));
    root.querySelector('.ed-tools [data-act="filter"]')?.setAttribute('aria-pressed', String(!!tab()?.filter));
    const sel = root.querySelector('.ed-tools [data-act="nf"]');
    if (sel) { const p = FORMATS.find((x) => x.nf === (f.nf ?? null)); sel.value = p ? p.id : 'custom'; }
  }
  const selRange = () => { const s = norm(S.sel); const u = usedRange(S.doc, S.tab); return rangeText({ r1: s.r1, c1: s.c1, r2: s.r2 >= S.geo.rows - 1 && u ? Math.max(s.r1, u.r2) : s.r2, c2: s.c2 >= S.geo.cols - 1 && u ? Math.max(s.c1, u.c2) : s.c2 }); };
  async function format(patch, label = 'Format') {
    const s = norm(S.sel);
    const range = selRange();
    const p = parseRange(range);
    const undo = snapshot(S.tab, p.r1, p.c1, p.r2, p.c2, label);
    undo.formatOnly = true;
    await run('sheets.format_range', { tab: S.tab, range, ...patch }, { undo });
    void s;
  }
  function toggleFmt(k) {
    const f = getFmt(S.doc, S.tab, S.sel.r, S.sel.c) ?? {};
    const name = { b: 'bold', i: 'italic', u: 'underline', s: 'strike' }[k];
    return format({ [name]: !f[k] }, `${name[0].toUpperCase()}${name.slice(1)}`);
  }
  function decimals(more) {
    const f = getFmt(S.doc, S.tab, S.sel.r, S.sel.c) ?? {};
    let nf = f.nf && f.nf !== '@' ? f.nf : '0';
    if (isDateFormat(nf)) return;
    const m = /0(\.0*)?/.exec(nf);
    const d = m ? (m[1] ? m[1].length - 1 : 0) : 0;
    const nd = Math.max(0, Math.min(10, d + (more ? 1 : -1)));
    nf = nf.replace(/0(\.0*)?(?!.*0)/, `0${nd ? `.${'0'.repeat(nd)}` : ''}`);
    if (!/0/.test(nf)) nf = nd ? `0.${'0'.repeat(nd)}` : '0';
    return format({ number_format: nf }, 'Decimals');
  }
  if (!readOnly) {
    root.addEventListener('change', (e) => { if (e.target.matches('.ed-tools [data-act="nf"]') && e.target.value !== 'custom') { format({ number_format: e.target.value === 'auto' ? null : e.target.value }, 'Number format'); focusGrid(); } });
  }
  root.addEventListener('click', async (e) => {
    const b = e.target.closest('[data-act], [data-chart-menu]');
    if (!b || !root.contains(b) || b.tagName === 'SELECT') return;
    if (b.dataset.chartMenu) return chartMenu(b, b.dataset.chartMenu);
    const act = b.dataset.act;
    if (S.editing && !['undo', 'redo'].includes(act)) await commit();
    const col = colName(S.sel.c);
    const acts = {
      undo: () => undoRedo(false), redo: () => undoRedo(true),
      cur: () => format({ number_format: 'currency' }, 'Currency'), pct: () => format({ number_format: 'percent' }, 'Percent'),
      decl: () => decimals(false), decm: () => decimals(true),
      b: () => toggleFmt('b'), i: () => toggleFmt('i'), u: () => toggleFmt('u'), s: () => toggleFmt('s'),
      left: () => format({ align: 'left' }, 'Align'), center: () => format({ align: 'center' }, 'Align'), right: () => format({ align: 'right' }, 'Align'),
      wrap: () => format({ wrap: !(getFmt(S.doc, S.tab, S.sel.r, S.sel.c)?.wrap) }, 'Wrap'),
      color: () => swatches(b, COLORS, (c) => format({ color: c }, 'Text colour'), 'Text colour'),
      fill: () => swatches(b, FILLS, (c) => format({ fill: c }, 'Fill colour'), 'Fill colour'),
      clearfmt: () => format({ clear: true }, 'Clear formatting'),
      freeze: () => freezeMenu(b),
      sortUp: () => sortBy(col, 'asc'), sortDown: () => sortBy(col, 'desc'),
      filter: () => toggleFilter(),
      chart: () => chartDialog(), pivot: () => pivotDialog(), comment: () => commentNew(), crm: () => crmDialog(), find: () => findDialog(),
      comments: () => commentsPanel(), history: () => historyPanel(), download: () => downloadMenu(b), share: () => shareDialog(),
      copy: () => copyShared(),
    };
    await acts[act]?.();
    if (!S.editing && !b.closest('.ed-acts') && !['color', 'fill', 'freeze', 'chart', 'pivot', 'comment', 'crm', 'find'].includes(act)) focusGrid();
  });

  function swatches(at, list, pick, label) {
    const r = at.getBoundingClientRect();
    const p = popover(document.body, { x: r.left, y: r.bottom + 4 }, `<div class="sw-grid" role="group" aria-label="${h(label)}">${list.map((c) => `<button type="button" class="sw${c ? '' : ' is-none'}" ${tool('sheets.format_range')} data-c="${c ?? ''}" style="${c ? `--sw:${c}` : ''}" aria-label="${c ?? 'None'}" title="${c ?? 'None'}"></button>`).join('')}</div>`);
    p.addEventListener('click', (e) => { const s = e.target.closest('[data-c]'); if (s) { closePop(); pick(s.dataset.c || null); focusGrid(); } });
  }

  function freezeMenu(at) {
    const t = tab();
    menu(document.body, at, [
      { head: 'Rows' },
      ...[0, 1, 2].map((n) => ({ label: n ? `${n} row${n > 1 ? 's' : ''}` : 'No rows', tool: 'sheets.freeze', checked: t.frozen_rows === n, run: () => run('sheets.freeze', { tab: S.tab, rows: n }) })),
      { label: `Up to row ${S.sel.r + 1}`, tool: 'sheets.freeze', run: () => run('sheets.freeze', { tab: S.tab, rows: S.sel.r + 1 }) },
      { sep: true }, { head: 'Columns' },
      ...[0, 1, 2].map((n) => ({ label: n ? `${n} column${n > 1 ? 's' : ''}` : 'No columns', tool: 'sheets.freeze', checked: t.frozen_cols === n, run: () => run('sheets.freeze', { tab: S.tab, columns: n }) })),
      { label: `Up to column ${colName(S.sel.c)}`, tool: 'sheets.freeze', run: () => run('sheets.freeze', { tab: S.tab, columns: S.sel.c + 1 }) },
    ], { label: 'Freeze' });
  }

  async function sortBy(col, order) {
    const s = norm(S.sel);
    const single = s.r1 === s.r2 && s.c1 === s.c2;
    const u = usedRange(S.doc, S.tab);
    if (!u) return;
    const range = single || s.r2 - s.r1 >= S.geo.rows - 2 ? null : rangeText({ ...s, c2: Math.min(s.c2, u.c2), r2: Math.min(s.r2, u.r2) });
    const g = range ? parseRange(range) : { r1: 0, c1: 0, r2: u.r2, c2: u.c2 };
    const undo = snapshot(S.tab, g.r1, g.c1, g.r2, g.c2, 'Sort');
    const out = await run('sheets.sort', { tab: S.tab, ...(range ? { range } : {}), by: { column: col, order } }, { undo });
    if (out) toast(`Sorted by ${out.sorted_by?.[0] ?? col}${out.header ? ', header row kept on top' : ''}`);
  }

  async function toggleFilter() {
    const t = tab();
    if (t.filter) return run('sheets.filter', { tab: S.tab, clear: true });
    const s = norm(S.sel);
    const single = s.r1 === s.r2 && s.c1 === s.c2;
    const u = usedRange(S.doc, S.tab);
    await run('sheets.filter', { tab: S.tab, ...(single || !u ? {} : { range: rangeText({ ...s, r2: Math.min(s.r2, u.r2), c2: Math.min(s.c2, u.c2) }) }) });
  }

  function openFilter(c) {
    const t = tab();
    const f = parseRange(t.filter.range);
    const cur = t.filter.criteria?.[c] ?? null;
    const u = usedRange(S.doc, S.tab);
    const seen = new Map();
    for (let r = f.r1 + 1; r <= Math.min(f.r2, u?.r2 ?? f.r1); r++) {
      const v = S.engine.value(S.tab, r, c);
      const shown = v === null ? '' : display(v, getFmt(S.doc, S.tab, r, c)?.nf, S.engine.type(S.tab, r, c));
      const raw = plainValue(v);
      const k = String(raw ?? '').toLowerCase();
      if (!seen.has(k)) seen.set(k, { raw, shown });
      if (seen.size > 300) break;
    }
    const vals = [...seen.values()].sort((a, b) => ops.compareValues(a.raw, b.raw));
    const on = (x) => (cur?.op === 'in' ? (cur.values ?? []).map((y) => String(y ?? '').toLowerCase()).includes(String(x.raw ?? '').toLowerCase()) : !cur);
    const r = rect(f.r1, c);
    const g = grid.getBoundingClientRect();
    const head = plainValue(S.engine.value(S.tab, f.r1, c)) ?? colName(c);
    const p = popover(document.body, { x: g.left + r.x, y: g.top + r.y + r.h + 2 }, `<form class="flt" ${tool('sheets.filter')}><p class="flt-h"><b>${h(head)}</b><span class="flt-sort"><button type="button" class="ui-btn is-ghost is-sm" data-sort="asc" ${tool('sheets.sort')}>${I.sortUp}A to Z</button><button type="button" class="ui-btn is-ghost is-sm" data-sort="desc" ${tool('sheets.sort')}>${I.sortDown}Z to A</button></span></p>
      <label class="ui-field"><span>Show rows where</span><select class="ui-select" name="op"><option value="in"${!cur || cur.op === 'in' ? ' selected' : ''}>The value is ticked below</option>${[['contains', 'Text contains'], ['eq', 'Is equal to'], ['neq', 'Is not equal to'], ['gt', 'Greater than'], ['lt', 'Less than'], ['not_empty', 'Is not empty'], ['empty', 'Is empty']].map(([o, l]) => `<option value="${o}"${cur?.op === o ? ' selected' : ''}>${l}</option>`).join('')}</select></label>
      <input class="ui-input flt-v" name="value" placeholder="Value" value="${h(cur && cur.op !== 'in' ? cur.value ?? '' : '')}">
      <div class="flt-list">${vals.map((x, i) => `<label class="flt-i"><input type="checkbox" name="v${i}" value="${i}"${on(x) ? ' checked' : ''}><span>${h(x.shown || '(empty)')}</span></label>`).join('')}</div>
      <div class="sh-pop-a"><button type="button" class="ui-btn is-quiet is-sm" data-clear ${tool('sheets.filter')}>Show all</button><button type="submit" class="ui-btn is-accent is-sm" ${tool('sheets.filter')}>Apply</button></div></form>`, { className: 'is-filter' });
    const form = p.querySelector('form');
    const sync1 = () => { const op = form.op.value; form.querySelector('.flt-v').hidden = op === 'in' || op === 'empty' || op === 'not_empty'; form.querySelector('.flt-list').hidden = op !== 'in'; };
    form.op.addEventListener('change', sync1); sync1();
    form.addEventListener('submit', async (e) => {
      e.preventDefault();
      const op = form.op.value;
      let condition;
      if (op === 'in') {
        const ticked = vals.filter((_, i) => form[`v${i}`].checked);
        condition = ticked.length === vals.length ? { op: 'all' } : { op: 'in', values: ticked.map((x) => x.raw) };
      } else condition = { op, value: form.value.value };
      closePop();
      await run('sheets.filter', { tab: S.tab, column: colName(c), condition });
    });
    p.querySelector('[data-clear]').addEventListener('click', async () => { closePop(); await run('sheets.filter', { tab: S.tab, column: colName(c), condition: { op: 'all' } }); });
    p.querySelectorAll('[data-sort]').forEach((x) => x.addEventListener('click', async () => {
      closePop();
      const range = rangeText({ r1: f.r1, c1: f.c1, r2: Math.min(f.r2, u.r2), c2: f.c2 });
      const undo = snapshot(S.tab, f.r1, f.c1, Math.min(f.r2, u.r2), f.c2, 'Sort');
      await run('sheets.sort', { tab: S.tab, range, by: { column: colName(c), order: x.dataset.sort }, header: true }, { undo });
    }));
  }

  // ---------- right-click menu ----------

  function cellMenu(at, kind) {
    const s = norm(S.sel);
    const rows = s.r2 - s.r1 + 1, cols = s.c2 - s.c1 + 1;
    const fullCols = kind === 'col', fullRows = kind === 'row';
    menu(document.body, at, readOnly ? [
      { label: 'Copy', why: 'copies the cells', icon: I.copy, hint: 'Ctrl+C', run: () => copySel(null, false) },
    ] : [
      { label: 'Cut', why: 'cuts the cells (pasting moves them)', icon: I.clear, hint: 'Ctrl+X', run: () => copySel(null, true) },
      { label: 'Copy', why: 'copies the cells', icon: I.copy, hint: 'Ctrl+C', run: () => copySel(null, false) },
      { label: 'Paste', tool: 'sheets.write_range', icon: I.download, hint: 'Ctrl+V', run: async () => { try { await paste(await navigator.clipboard.readText()); } catch { toast('Use Ctrl+V (or Cmd+V) to paste.'); } } },
      { sep: true },
      !fullCols && { label: `Insert ${rows > 1 ? `${rows} rows` : 'a row'} above`, tool: 'sheets.insert_rows', icon: I.rows, run: () => insert('row', s.r1 + 1, rows) },
      !fullCols && { label: `Insert ${rows > 1 ? `${rows} rows` : 'a row'} below`, tool: 'sheets.insert_rows', icon: I.rows, run: () => insert('row', s.r2 + 2, rows) },
      !fullRows && { label: `Insert ${cols > 1 ? `${cols} columns` : 'a column'} left`, tool: 'sheets.insert_columns', icon: I.rows, run: () => insert('col', s.c1 + 1, cols) },
      !fullRows && { label: `Insert ${cols > 1 ? `${cols} columns` : 'a column'} right`, tool: 'sheets.insert_columns', icon: I.rows, run: () => insert('col', s.c2 + 2, cols) },
      !fullCols && { label: rows > 1 ? `Delete rows ${s.r1 + 1} to ${s.r2 + 1}` : `Delete row ${s.r1 + 1}`, tool: 'sheets.delete_rows', icon: I.x, run: () => remove('row', s.r1 + 1, rows) },
      !fullRows && { label: cols > 1 ? `Delete columns ${colName(s.c1)} to ${colName(s.c2)}` : `Delete column ${colName(s.c1)}`, tool: 'sheets.delete_columns', icon: I.x, run: () => remove('col', s.c1 + 1, cols) },
      { sep: true },
      { label: 'Sort A to Z', tool: 'sheets.sort', icon: I.sortUp, run: () => sortBy(colName(S.sel.c), 'asc') },
      { label: 'Sort Z to A', tool: 'sheets.sort', icon: I.sortDown, run: () => sortBy(colName(S.sel.c), 'desc') },
      { label: 'Comment', tool: 'sheets.add_comment', icon: I.comment, hint: '', run: () => commentNew() },
      { label: 'Clear', tool: 'sheets.write_range', icon: I.clear, hint: 'Del', run: () => clearSel() },
    ], { label: 'Cell' });
  }
  async function insert(axis, at, count) {
    await run(axis === 'row' ? 'sheets.insert_rows' : 'sheets.insert_columns', { tab: S.tab, at, count });
    toast(`${count} ${axis === 'row' ? 'row' : 'column'}${count > 1 ? 's' : ''} added`);
  }
  async function remove(axis, at, count) {
    await run(axis === 'row' ? 'sheets.delete_rows' : 'sheets.delete_columns', { tab: S.tab, at, count });
    toast(`${count} ${axis === 'row' ? 'row' : 'column'}${count > 1 ? 's' : ''} deleted. Version history has them.`);
  }

  // ---------- comments ----------

  function showCommentHint() {
    if (S.editing || drag) return;
    const list = allComments(S.doc).filter((c) => c.tab === S.tab && c.r === S.sel.r && c.c === S.sel.c && !c.resolved);
    if (!list.length) return;
    openComment(list[0]);
  }
  function commentHtml(c) {
    return `<div class="cm" data-cm="${c.id}"><div class="cm-h"><span class="ui-avatar is-sm">${h(initials(c.by?.name))}</span><b>${h(c.by?.name ?? 'Someone')}</b><time>${h(ago(c.at))}</time>${readOnly ? '' : `<button type="button" class="ui-btn is-ghost is-sm" data-resolve="${c.id}" ${tool('sheets.resolve_comment')} title="Resolve">${I.check}</button>`}</div><p>${h(c.body)}</p>${(c.replies ?? []).map((x) => `<div class="cm-r"><b>${h(x.by?.name ?? 'Someone')}</b> <time>${h(ago(x.at))}</time><p>${h(x.body)}</p></div>`).join('')}${readOnly ? '' : `<form class="cm-f" ${tool('sheets.reply_comment')}><input class="ui-input" name="body" placeholder="Reply" required maxlength="5000"><button class="ui-btn is-quiet is-sm" type="submit" ${tool('sheets.reply_comment')}>Reply</button></form>`}</div>`;
  }
  function openComment(c) {
    const r = rect(c.r, c.c);
    const g = grid.getBoundingClientRect();
    const p = popover(document.body, { x: g.left + r.x + r.w + 6, y: g.top + r.y }, commentHtml(c), { className: 'is-comment' });
    wireComment(p);
  }
  function wireComment(p) {
    p.addEventListener('submit', async (e) => {
      const f = e.target.closest('.cm-f');
      if (!f) return;
      e.preventDefault();
      const id = f.closest('[data-cm]').dataset.cm;
      await run('sheets.reply_comment', { comment: id, body: f.body.value });
      f.closest('[data-cm]').outerHTML = commentHtml(allComments(S.doc).find((x) => x.id === id));
    });
    p.addEventListener('click', async (e) => {
      const b = e.target.closest('[data-resolve]');
      if (!b) return;
      await run('sheets.resolve_comment', { comment: b.dataset.resolve, resolved: true });
      b.closest('[data-cm]').remove();
      if (p.classList.contains('is-comment')) closePop();
      toast('Resolved');
    });
  }
  function commentNew() {
    const r = rect(S.sel.r, S.sel.c);
    const g = grid.getBoundingClientRect();
    const p = popover(document.body, { x: g.left + r.x + r.w + 6, y: g.top + r.y }, `<form class="cm-new" ${tool('sheets.add_comment')}><p class="ui-label">Comment on ${addr(S.sel.r, S.sel.c)}</p><textarea class="ui-textarea" name="body" rows="3" required maxlength="5000" placeholder="Write a comment"></textarea><div class="sh-pop-a"><button type="button" class="ui-btn is-quiet is-sm" data-x ${none('closes the comment box')}>Cancel</button><button class="ui-btn is-accent is-sm" type="submit" ${tool('sheets.add_comment')}>Comment</button></div></form>`, { className: 'is-comment' });
    const f = p.querySelector('form');
    f.body.focus();
    p.querySelector('[data-x]').addEventListener('click', () => { closePop(); focusGrid(); });
    f.addEventListener('submit', async (e) => { e.preventDefault(); await run('sheets.add_comment', { tab: S.tab, cell: addr(S.sel.r, S.sel.c), body: f.body.value }); closePop(); focusGrid(); });
  }
  function commentsPanel() {
    const list = allComments(S.doc).filter((c) => !c.resolved);
    const d = dialog(document.body, { title: `Comments${list.length ? ` (${list.length})` : ''}`, cancel: null, submit: 'Done', wide: true, plain: true, body: list.length ? `<div class="cm-list">${list.map((c) => `<div class="cm-wrap"><button type="button" class="cm-go" data-go="${c.id}" ${none('shows the cell')}>${h(findTab(S.doc, c.tab)?.name ?? '')}!${addr(c.r, c.c)}</button>${commentHtml(c)}</div>`).join('')}</div>` : '<p class="ui-empty">No open comments. Select a cell and press the comment button to add one.</p>' });
    wireComment(d);
    d.addEventListener('click', (e) => {
      const go = e.target.closest('[data-go]');
      if (!go) return;
      const c = allComments(S.doc).find((x) => x.id === go.dataset.go);
      d.close();
      if (c) { S.tab = c.tab; layout(); selectCell(c.r, c.c); }
    });
  }

  // ---------- charts, pivots, the CRM, find ----------

  function chartForm(c = {}) {
    const s = norm(S.sel);
    const range = c.range ?? (s.r1 === s.r2 && s.c1 === s.c2 ? (usedRange(S.doc, S.tab) ? rangeText(usedRange(S.doc, S.tab)) : 'A1:B10') : selRange());
    return `<div class="ui-seg ch-types" role="group" aria-label="Type">${ops.CHART_TYPES.map((t) => `<button type="button" data-type="${t}" ${none('picks the chart type')} aria-pressed="${(c.type ?? 'column') === t}">${t[0].toUpperCase()}${t.slice(1)}</button>`).join('')}</div><input type="hidden" name="type" value="${c.type ?? 'column'}">
      <label class="ui-field"><span>Data <small>first column labels, first row headers</small></span><input class="ui-input" name="range" value="${h(range)}" required></label>
      <label class="ui-field"><span>Title</span><input class="ui-input" name="title" value="${h(c.title ?? '')}" maxlength="120" placeholder="Optional"></label>
      <label class="ui-field"><span>Series</span><select class="ui-select" name="series_by"><option value="columns"${c.series_by !== 'rows' ? ' selected' : ''}>One per column</option><option value="rows"${c.series_by === 'rows' ? ' selected' : ''}>One per row</option></select></label>
      <div class="ch-prev"></div>`;
  }
  function wireChartForm(d, existing = null) {
    const form = d.querySelector('form');
    const prev = () => {
      try {
        const t = tab();
        const spec = { tab: t.id, type: form.type.value, range: rangeText(parseRange(form.range.value)), by: form.series_by.value, series: existing?.series ?? null, title: form.title.value };
        d.querySelector('.ch-prev').innerHTML = renderChart({ ...spec, data: ops.chartData(S.doc, S.engine, spec) }, Math.min(440, d.clientWidth - 60), 200);
      } catch { d.querySelector('.ch-prev').innerHTML = '<p class="ui-hint">Type a range like A1:C10.</p>'; }
    };
    d.querySelectorAll('[data-type]').forEach((b) => b.addEventListener('click', () => { d.querySelectorAll('[data-type]').forEach((x) => x.setAttribute('aria-pressed', String(x === b))); form.type.value = b.dataset.type; prev(); }));
    form.addEventListener('input', prev);
    prev();
  }
  function chartDialog() {
    const d = dialog(document.body, { title: 'Add a chart', toolName: 'sheets.create_chart', submit: 'Add chart', wide: true, body: chartForm(), onSubmit: async (v) => {
      const s = norm(S.sel);
      const out = await run('sheets.create_chart', { tab: S.tab, type: v.type, range: v.range, title: v.title || undefined, series_by: v.series_by, at: addr(s.r1, Math.min(S.geo.cols - 1, (usedRange(S.doc, S.tab)?.c2 ?? s.c2) + 2)) });
      if (out) toast('Chart added. Drag its title to move it.');
    } });
    wireChartForm(d);
  }
  function chartMenu(at, id) {
    const c = allCharts(S.doc).find((x) => x.id === id);
    menu(document.body, at, [
      { label: 'Edit chart', tool: 'sheets.update_chart', icon: I.chart, run: () => {
        const d = dialog(document.body, { title: 'Edit chart', toolName: 'sheets.update_chart', wide: true, body: chartForm({ ...c, series_by: c.by }), onSubmit: async (v) => { await run('sheets.update_chart', { chart: id, type: v.type, range: v.range, title: v.title || null, series_by: v.series_by }); } });
        wireChartForm(d, c);
      } },
      { label: 'Delete chart', tool: 'sheets.delete_chart', icon: I.x, danger: true, run: () => run('sheets.delete_chart', { chart: id }) },
    ], { label: 'Chart' });
  }
  // Move a chart by its title, resize it by its corner.
  $('.sg-charts').addEventListener('pointerdown', (e) => {
    if (readOnly || e.target.closest('button')) return;
    const el = e.target.closest('.sg-chart');
    if (!el) return;
    const resize = e.target.classList.contains('sg-chart-rs');
    if (!resize && !e.target.closest('figcaption')) return;
    e.preventDefault();
    const c = allCharts(S.doc).find((x) => x.id === el.dataset.chart);
    const m = /translate\(([-\d.]+)px, ([-\d.]+)px\)/.exec(el.style.transform);
    const st = { x: e.clientX, y: e.clientY, tx: Number(m?.[1] ?? 0), ty: Number(m?.[2] ?? 0), w: el.offsetWidth, h: el.offsetHeight };
    el.setPointerCapture(e.pointerId);
    el.classList.add('is-drag');
    const mv = (ev) => {
      if (resize) { el.style.width = `${Math.max(240, st.w + ev.clientX - st.x)}px`; el.style.height = `${Math.max(180, st.h + ev.clientY - st.y)}px`; }
      else el.style.transform = `translate(${st.tx + ev.clientX - st.x}px, ${st.ty + ev.clientY - st.y}px)`;
    };
    const up = async (ev) => {
      el.removeEventListener('pointermove', mv); el.removeEventListener('pointerup', up);
      el.classList.remove('is-drag');
      if (Math.hypot(ev.clientX - st.x, ev.clientY - st.y) < 3) return;
      if (resize) await run('sheets.update_chart', { chart: c.id, width: el.offsetWidth, height: el.offsetHeight }, { quiet: true });
      else {
        const g = grid.getBoundingClientRect();
        const x = hit(Math.max(HW + 1, el.getBoundingClientRect().left - g.left + 2), Math.max(HH + 1, el.getBoundingClientRect().top - g.top + 2));
        await run('sheets.update_chart', { chart: c.id, at: addr(x.r, x.c) }, { quiet: true });
      }
      el.dataset.sig = '';
    };
    el.addEventListener('pointermove', mv);
    el.addEventListener('pointerup', up);
  });

  function pivotDialog() {
    const u = usedRange(S.doc, S.tab);
    if (!u) return toast('This tab is empty. A pivot needs a table with a header row.');
    const s = norm(S.sel);
    const g = s.r1 === s.r2 && s.c1 === s.c2 ? u : { r1: s.r1, c1: s.c1, r2: Math.min(s.r2, u.r2), c2: Math.min(s.c2, u.c2) };
    const heads = [];
    for (let c = g.c1; c <= g.c2; c++) heads.push({ c, name: String(plainValue(S.engine.value(S.tab, g.r1, c)) ?? colName(c)) });
    const opts = (sel) => heads.map((x) => `<option value="${colName(x.c)}"${x.c === sel ? ' selected' : ''}>${h(x.name)}</option>`).join('');
    const numCol = heads.find((x) => typeof plainValue(S.engine.value(S.tab, g.r1 + 1, x.c)) === 'number')?.c ?? g.c1;
    dialog(document.body, { title: 'Pivot table', toolName: 'sheets.create_pivot', submit: 'Make pivot table', body: `<p class="ui-hint">From ${h(tab().name)}!${rangeText(g)}. It goes on a new tab as live formulas.</p>
      <label class="ui-field"><span>Rows</span><select class="ui-select" name="rows">${opts(g.c1)}</select></label>
      <label class="ui-field"><span>Columns <small>optional</small></span><select class="ui-select" name="columns"><option value="">None</option>${opts(-1)}</select></label>
      <div class="ui-fields pv-v"><label class="ui-field"><span>Values</span><select class="ui-select" name="how"><option value="sum">Sum of</option><option value="count">Count of rows</option><option value="average">Average of</option><option value="min">Smallest</option><option value="max">Largest</option></select></label><label class="ui-field"><span>Field</span><select class="ui-select" name="field">${opts(numCol)}</select></label></div>`,
    onSubmit: async (v) => {
      const out = await run('sheets.create_pivot', { tab: S.tab, range: rangeText(g), rows: v.rows, ...(v.columns ? { columns: v.columns } : {}), values: [v.how === 'count' ? { summarize: 'count' } : { field: v.field, summarize: v.how }] });
      if (out) { const t = findTab(S.doc, out.tab); if (t) { S.tab = t.id; layout(); selectCell(0, 0); } toast(`Pivot table on ${out.tab}`); }
    } });
  }

  async function crmDialog() {
    let settings = null;
    try { settings = await ctx.callTool('sheets.get_settings', {}); } catch {}
    const ok = settings?.crm?.connected;
    dialog(document.body, { title: 'Pull from the CRM', toolName: 'sheets.import_from_crm', submit: ok ? 'Pull records' : null, cancel: 'Close', body: ok ? `
      <p class="ui-hint">Records land on a new tab, linked: Refresh brings them up to date and formulas beside them keep working.</p>
      <label class="ui-field"><span>What</span><select class="ui-select" name="kind"><option value="deals">Deals</option><option value="contacts">Contacts</option><option value="leads">Leads</option><option value="organizations">Organizations</option><option value="activities">Activities</option><option value="soql">A query (SOQL)</option></select></label>
      <label class="ui-field crm-q" hidden><span>Query</span><textarea class="ui-textarea" name="soql" rows="3" placeholder="SELECT Name, StageName, Amount FROM Opportunity WHERE Amount > 5000"></textarea></label>
      <label class="ui-field"><span>Only where <small>optional</small></span><input class="ui-input" name="where" placeholder="Amount > 5000"></label>
      <label class="ui-field"><span>Tab name</span><input class="ui-input" name="name" placeholder="CRM deals" maxlength="100"></label>` : `<p>No CRM is connected to Sheets yet. ${['owner', 'admin'].includes(ctx.me?.role) ? `Connect one in <a href="${ctx.standalone ? '#/settings' : '/a/sheets/settings'}" data-go="/settings">Settings</a>.` : 'Ask an admin to connect one in Settings.'}</p>`,
    onSubmit: async (v) => {
      const out = await run('sheets.import_from_crm', v.kind === 'soql' ? { soql: v.soql, name: v.name || undefined } : { kind: v.kind, ...(v.where ? { where: v.where } : {}), name: v.name || undefined });
      if (out) { const t = findTab(S.doc, out.tab); if (t) { S.tab = t.id; layout(); selectCell(0, 0); } toast(`${out.rows} records from the CRM`); }
    } });
    const d = document.querySelector('dialog[open] form');
    d?.kind?.addEventListener('change', () => { d.querySelector('.crm-q').hidden = d.kind.value !== 'soql'; });
  }

  function findDialog() {
    const d = dialog(document.body, { title: 'Find', toolName: 'sheets.find', submit: 'Find', cancel: 'Close', body: `<label class="ui-field"><span>Look for</span><input class="ui-input" name="query" required></label><label class="flt-i"><input type="checkbox" name="formulas" value="1"><span>Search inside formulas</span></label><div class="fd-res"></div>`,
      onSubmit: async (v, form) => {
        const out = await callTool('sheets.find', { query: v.query, formulas: !!v.formulas, limit: 200 });
        form.querySelector('.fd-res').innerHTML = out.matches.length ? `<p class="ui-hint">${out.total} found</p><div class="fd-list">${out.matches.map((m) => `<button type="button" data-go="${h(m.tab)}!${h(m.cell)}" ${none('goes to the cell')}><b>${h(m.tab)}!${h(m.cell)}</b><span>${h(m.formula ?? m.text)}</span></button>`).join('')}</div>` : '<p class="ui-empty">Nothing matches.</p>';
        return false;
      } });
    d.addEventListener('click', (e) => {
      const b = e.target.closest('[data-go]');
      if (!b) return;
      const p = parseRange(b.dataset.go);
      const t = findTab(S.doc, p.tab);
      if (t && t.id !== S.tab) { S.tab = t.id; layout(); }
      selectCell(p.r1, p.c1);
    });
  }

  // ---------- history, download, share ----------

  async function historyPanel() {
    const out = await callTool('sheets.list_versions', {});
    const d = dialog(document.body, { title: 'Version history', cancel: null, submit: 'Done', wide: true, plain: true, body: `<form class="vh-save" ${tool('sheets.save_version')}><input class="ui-input" name="label" placeholder="Name this version, like Before the Q4 update" required maxlength="120"><button class="ui-btn is-quiet" type="submit" ${tool('sheets.save_version')}>Save</button></form>
      <div class="vh-list">${out.versions.map((v) => `<div class="vh-i"><div><b>${h(v.label ?? new Date(v.created_at).toLocaleString(undefined, { month: 'short', day: 'numeric', hour: 'numeric', minute: '2-digit' }))}</b><span>${h(v.created_by_name ?? '')}${v.created_by_name ? ' · ' : ''}${h(ago(v.created_at))}${v.auto ? ' · saved automatically' : ''}</span></div><button type="button" class="ui-btn is-ghost is-sm" data-peek="${v.id}" ${tool('sheets.read_range')}>Look</button><button type="button" class="ui-btn is-quiet is-sm" data-restore="${v.id}" ${tool('sheets.restore_version')}>Restore</button></div>`).join('')}</div><div class="vh-peek"></div>` });
    d.querySelector('.vh-save').addEventListener('submit', async (e) => {
      e.preventDefault(); e.stopPropagation();
      await callTool('sheets.save_version', { label: e.target.label.value });
      d.close(); toast('Version saved'); historyPanel();
    });
    d.addEventListener('click', async (e) => {
      const r = e.target.closest('[data-restore]');
      if (r) {
        await run('sheets.restore_version', { version: r.dataset.restore });
        d.close(); S.tab = findTab(S.doc, S.tab)?.id ?? tabList(S.doc)[0].id; layout(); toast('Restored. The version before is in the history too.');
      }
      const p = e.target.closest('[data-peek]');
      if (p) {
        const t = tab();
        const v = await callTool('sheets.read_range', { tab: t.name, version: p.dataset.peek, text: true }).catch((err) => ({ error: err.message }));
        const box = d.querySelector('.vh-peek');
        box.innerHTML = v.error ? `<p class="ui-hint">${h(t.name)} did not exist in that version.</p>` : `<p class="ui-label">${h(t.name)} in that version</p><div class="ui-dtable-wrap"><table class="ui-dtable vh-t"><tbody>${(v.text ?? []).slice(0, 40).map((row) => `<tr>${row.slice(0, 12).map((x) => `<td>${h(x)}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`;
      }
    });
  }

  function downloadMenu(at) {
    const go = async (fmt) => {
      const out = await callTool('sheets.export', { format: fmt, tab: S.tab });
      if (out.not_carried?.length) toast(`Downloaded. Not in the file: ${out.not_carried[0]}`);
      const a = document.createElement('a');
      a.href = out.file.url; a.download = out.file.name;
      document.body.appendChild(a); a.click(); a.remove();
    };
    menu(document.body, at, [
      { label: 'Excel (.xlsx), every tab', tool: 'sheets.export', icon: I.download, run: () => go('xlsx') },
      { label: `CSV (.csv), ${tab().name} only`, tool: 'sheets.export', icon: I.download, run: () => go('csv') },
    ], { label: 'Download' });
  }

  async function shareDialog() {
    const on = !!S.share;
    const d = dialog(document.body, { title: 'Share', cancel: null, submit: 'Done', plain: true, body: `<p>Everyone on your team can already open and edit this spreadsheet.</p>
      <div class="sh-share"><div><b>View-only link</b><span>Anyone with the link can look, without an account. They cannot edit or see comments.</span></div>
      <button type="button" class="ui-btn ${on ? 'is-quiet' : 'is-accent'} is-sm" data-link="${on ? 'off' : 'view'}" ${tool('sheets.share')}>${on ? 'Turn off' : 'Turn on'}</button></div>
      ${on ? `<div class="ui-copy sh-url"><code>${h(S.share.url)}</code><button type="button" class="ui-btn is-quiet is-sm" data-copy="${h(S.share.url)}" ${none('copies the link')}>Copy</button></div>` : ''}` });
    d.addEventListener('click', async (e) => {
      const b = e.target.closest('[data-link]');
      if (b) { await run('sheets.share', { link: b.dataset.link }); d.close(); shareDialog(); }
      const c = e.target.closest('[data-copy]');
      if (c) { await copyText(c.dataset.copy); toast('Link copied'); }
    });
  }

  async function copyShared() {
    if (!ctx.signedIn) { location.href = `/login?next=${encodeURIComponent(`/#/copy/${view.token}`)}`; return; }
    try { const out = await ctx.callTool('sheets.copy_sheet', { share_token: view.token }); location.href = `/#/s/${out.id}`; } catch (e) { toast(e.message); }
  }

  // ---------- presence ----------

  let presT = null;
  function presence() {
    if (readOnly || !ctx.live) return;
    clearTimeout(presT);
    presT = setTimeout(() => { ctx.callTool('sheets.set_presence', { sheet: S.sheet, tab: S.tab, cell: addr(S.sel.r, S.sel.c) }).catch(() => {}); }, 400);
  }
  function drawPeople() {
    const el = $('.ed-who');
    const now = Date.now();
    for (const [k, o] of S.others) if (now - o.at > 5 * 60e3) S.others.delete(k);
    el.innerHTML = `<span class="ui-avatars">${[...S.others.values()].slice(0, 5).map((o) => `<span class="ui-avatar is-sm" title="${h(o.person?.name)} is here">${h(initials(o.person?.name))}</span>`).join('')}</span>`;
  }

  // ---------- size ----------

  const ro = new ResizeObserver(() => { S.dirty = true; schedule(); });
  ro.observe(grid);
  load();

  return {
    destroy() {
      S.destroyed = true;
      offLive?.();
      ro.disconnect();
      closeMenu(); closePop();
      if (!readOnly && ctx.live) ctx.callTool('sheets.set_presence', { sheet: S.sheet, left: true }).catch(() => {});
      S.engine?.destroy();
    },
    update(path) {
      const want = new URLSearchParams(path.split('?')[1] ?? '').get('tab');
      const t = want && findTab(S.doc, want);
      if (t && t.id !== S.tab) { S.tab = t.id; layout(); }
    },
    state: S,
  };
}

export { showError };
