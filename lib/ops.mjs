// What every tool does to a workbook. Each function takes the Yjs document and its formula engine and
// returns plain data. Values an agent reads are exact: computed values from the engine, and the formulas
// as written.
import { parseRange, rangeText, addr, colName, colIndex, key, unkey, shiftFormula, quoteTab } from './a1.mjs';
import { parseInput, display, plainValue, textRaw, FORMATS, isDateFormat } from './input.mjs';
import {
  Y, BookError, tabList, needTab, findTab, addTab, removeTab, tabMap, cellMap, fmtMap, setTabProp, getRaw, setRaw, getFmt, setFmt,
  usedRange, checkTabName, comments as allComments, charts as allCharts, newCommentId, newChartId, MAX_ROWS, MAX_COLS, FMT_KEYS,
} from './book.mjs';

const MAX_CELLS_READ = 20000;
const MAX_CELLS_WRITE = 50000;

// ---------- ranges ----------

// A range on a tab. The range may name its own tab ('Pipeline'!A1:C9), which wins over tab.
// Open ends (A:C, 2:5) and a missing range stop at the last cell in use.
export function resolve(doc, tabRef, range, { open = 'used', allowEmpty = true } = {}) {
  let p = null;
  if (range) {
    try { p = parseRange(range); } catch (e) { throw new BookError(e.message); }
  }
  const tab = needTab(doc, p?.tab ?? tabRef);
  const used = usedRange(doc, tab.id) ?? { r1: 0, c1: 0, r2: 0, c2: 0 };
  if (!p) {
    if (!allowEmpty) throw new BookError('Give a range, like A1:D20.');
    return { tab, r1: 0, c1: 0, r2: used.r2, c2: used.c2 };
  }
  const lastR = open === 'tab' ? tab.rows - 1 : used.r2;
  const lastC = open === 'tab' ? tab.cols - 1 : used.c2;
  return { tab, r1: p.r1, c1: p.c1, r2: p.r2 === Infinity ? Math.max(p.r1, lastR) : p.r2, c2: p.c2 === Infinity ? Math.max(p.c1, lastC) : p.c2 };
}

const size = (g) => (g.r2 - g.r1 + 1) * (g.c2 - g.c1 + 1);

function grow(doc, tab, r, c) {
  if (r >= MAX_ROWS || c >= MAX_COLS) throw new BookError(`A tab holds at most ${MAX_ROWS} rows and ${MAX_COLS} columns (A to ZZ).`);
  const m = tabMap(doc, tab.id);
  if (r >= (m.get('rows') ?? 1000)) m.set('rows', r + 1);
  if (c >= (m.get('cols') ?? 26)) m.set('cols', c + 1);
}

// ---------- reading ----------

export function cellView(doc, engine, tabId, r, c, { formats = false } = {}) {
  const raw = getRaw(doc, tabId, r, c);
  const fmt = getFmt(doc, tabId, r, c);
  const v = engine.value(tabId, r, c);
  const out = { cell: addr(r, c), value: plainValue(v), text: display(v, fmt?.nf, engine.type(tabId, r, c)) };
  if (typeof raw === 'string' && raw.startsWith('=')) out.formula = raw;
  if (formats && fmt) out.format = fmt;
  return out;
}

export function readRange(doc, engine, { tab, range, formats = false, text = false } = {}) {
  const g = resolve(doc, tab, range);
  if (size(g) > MAX_CELLS_READ) throw new BookError(`That is ${size(g)} cells; read at most ${MAX_CELLS_READ} at a time (for example ${rangeText({ r1: g.r1, c1: g.c1, r2: Math.min(g.r2, g.r1 + Math.floor(MAX_CELLS_READ / (g.c2 - g.c1 + 1)) - 1), c2: g.c2 })}).`);
  const values = [], formulas = [], texts = [], fmts = [];
  let anyFormula = false;
  for (let r = g.r1; r <= g.r2; r++) {
    const vr = [], fr = [], tr = [], mr = [];
    for (let c = g.c1; c <= g.c2; c++) {
      const v = engine.value(g.tab.id, r, c);
      const raw = getRaw(doc, g.tab.id, r, c);
      const isF = typeof raw === 'string' && raw.startsWith('=');
      anyFormula ||= isF;
      vr.push(plainValue(v));
      fr.push(isF ? raw : null);
      const f = getFmt(doc, g.tab.id, r, c);
      if (text) tr.push(display(v, f?.nf, engine.type(g.tab.id, r, c)));
      if (formats) mr.push(f ?? null);
    }
    values.push(vr); formulas.push(fr); texts.push(tr); fmts.push(mr);
  }
  const out = { tab: g.tab.name, range: rangeText(g), values };
  if (anyFormula) out.formulas = formulas;
  if (text) out.text = texts;
  if (formats) out.formats = fmts;
  return out;
}

// ---------- writing ----------

// values: rows of cells. Each cell is what a person would type ("$1,200", "=SUM(B2:B9)", "Acme") unless
// input is 'raw', which stores strings exactly as text. null clears a cell; undefined (a hole) leaves it.
export function writeRange(doc, engine, { tab, range, values, input = 'user', formats = null }) {
  if (!Array.isArray(values) || !values.every(Array.isArray)) throw new BookError('values must be rows of cells, like [["Name", "Amount"], ["Acme Dental", 1200]].');
  const start = range ? resolve(doc, tab, range, { open: 'tab' }) : resolve(doc, tab, 'A1');
  const h = values.length, w = Math.max(0, ...values.map((r) => r.length));
  if (h * w > MAX_CELLS_WRITE) throw new BookError(`Write at most ${MAX_CELLS_WRITE} cells at a time.`);
  const t = start.tab;
  const r2 = start.r1 + h - 1, c2 = start.c1 + w - 1;
  let n = 0;
  doc.transact(() => {
    if (h && w) grow(doc, t, r2, c2);
    for (let i = 0; i < h; i++) {
      for (let j = 0; j < values[i].length; j++) {
        const v = values[i][j];
        if (v === undefined) continue;
        const r = start.r1 + i, c = start.c1 + j;
        const cur = getFmt(doc, t.id, r, c);
        if (input === 'raw') setRaw(doc, t.id, r, c, typeof v === 'string' && !v.startsWith('=') ? textRaw(v) : v);
        else {
          const p = parseInput(v, { nf: cur?.nf ?? null });
          setRaw(doc, t.id, r, c, p.raw);
          if (p.nf && !cur?.nf) setFmt(doc, t.id, r, c, { ...(cur ?? {}), nf: p.nf });
        }
        n++;
      }
      if (formats?.[i]) for (let j = 0; j < formats[i].length; j++) if (formats[i][j] !== undefined) setFmt(doc, t.id, start.r1 + i, start.c1 + j, formats[i][j]);
    }
  });
  const g = { r1: start.r1, c1: start.c1, r2: Math.max(start.r1, r2), c2: Math.max(start.c1, c2) };
  return { tab: t.name, range: rangeText(g), cells: n, ...after(doc, engine, t, g) };
}

// The values in a range after a change, so an agent sees what its formulas came to without a second call.
function after(doc, engine, tab, g) {
  if (size(g) > 2000) return {};
  const r = readRange(doc, engine, { tab: tab.id, range: rangeText(g) });
  const errors = [];
  r.values.forEach((row, i) => row.forEach((v, j) => { if (typeof v === 'string' && /^#[A-Z0-9/!?]+[!?]?$/.test(v) && r.formulas?.[i]?.[j]) errors.push(`${addr(g.r1 + i, g.c1 + j)} ${r.formulas[i][j]} -> ${v}`); }));
  return { values: r.values, ...(r.formulas ? { formulas: r.formulas } : {}), ...(errors.length ? { errors } : {}) };
}

// One formula, optionally filled across a range the way dragging the fill handle does.
export function setFormula(doc, engine, { tab, cell, formula, fill = null }) {
  let f = String(formula ?? '').trim();
  if (!f.startsWith('=')) f = `=${f}`;
  if (!engine.validFormula(f)) throw new BookError(`${f} is not a formula the engine can read. Check the brackets, commas and quotes.`);
  const at = resolve(doc, tab, cell, { open: 'tab' });
  const t = at.tab;
  const g = fill ? resolve(doc, t.id, fill.includes('!') ? fill : `${quoteTab(t.name)}!${fill}`, { open: 'tab' }) : { r1: at.r1, c1: at.c1, r2: at.r1, c2: at.c1 };
  if (fill && (g.r1 > at.r1 || g.c1 > at.c1 || g.r2 < at.r1 || g.c2 < at.c1)) throw new BookError(`The fill range ${fill} must include ${addr(at.r1, at.c1)}.`);
  if (size(g) > MAX_CELLS_WRITE) throw new BookError(`Fill at most ${MAX_CELLS_WRITE} cells at a time.`);
  doc.transact(() => {
    grow(doc, t, g.r2, g.c2);
    for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) setRaw(doc, t.id, r, c, shiftFormula(f, r - at.r1, c - at.c1));
  });
  return { tab: t.name, range: rangeText(g), ...after(doc, engine, t, g) };
}

// Rows added under the last row in use (or under the table that starts at range).
export function appendRows(doc, engine, { tab, rows, input = 'user', range = null }) {
  if (!Array.isArray(rows) || !rows.every(Array.isArray)) throw new BookError('rows must be rows of cells, like [["Acme Dental", 1200]].');
  const t = needTab(doc, range ? parseRange(range).tab ?? tab : tab);
  let c1 = 0, last = -1;
  if (range) {
    const g = resolve(doc, t.id, range.includes('!') ? range : `${quoteTab(t.name)}!${range}`);
    c1 = g.c1;
    // The table is the block of rows below its first row with something in its columns.
    last = g.r1 - 1;
    for (let r = g.r1; ; r++) {
      let any = false;
      for (let c = g.c1; c <= Math.max(g.c2, g.c1); c++) if (getRaw(doc, t.id, r, c) !== null) { any = true; break; }
      if (!any) break;
      last = r;
    }
  } else {
    last = usedRange(doc, t.id)?.r2 ?? -1;
  }
  return writeRange(doc, engine, { tab: t.id, range: addr(last + 1, c1), values: rows, input });
}

export function clearRange(doc, engine, { tab, range, what = 'values' }) {
  const g = resolve(doc, tab, range, { allowEmpty: false });
  let n = 0;
  doc.transact(() => {
    const cells = cellMap(doc, g.tab.id), fmts = fmtMap(doc, g.tab.id);
    for (const m of [what !== 'formats' && cells, what !== 'values' && fmts].filter(Boolean)) {
      for (const k of [...m.keys()]) { const [r, c] = unkey(k); if (r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2) { m.delete(k); n++; } }
    }
  });
  return { tab: g.tab.name, range: rangeText(g), cleared: n };
}

// ---------- find ----------

export function find(doc, engine, { query, tab = null, match_case = false, whole_cell = false, formulas = false, regex = false, limit = 100 }) {
  if (query === undefined || query === null || query === '') throw new BookError('Say what to look for.');
  let re;
  try {
    const src = regex ? String(query) : String(query).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    re = new RegExp(whole_cell ? `^(?:${src})$` : src, match_case ? '' : 'i');
  } catch { throw new BookError('That is not a valid regular expression.'); }
  const tabs = tab ? [needTab(doc, tab)] : tabList(doc);
  const matches = [];
  let total = 0;
  for (const t of tabs) {
    const keys = [...cellMap(doc, t.id).keys()].map(unkey).sort((a, b) => a[0] - b[0] || a[1] - b[1]);
    for (const [r, c] of keys) {
      const raw = getRaw(doc, t.id, r, c);
      const v = engine.value(t.id, r, c);
      const shown = display(v, getFmt(doc, t.id, r, c)?.nf, engine.type(t.id, r, c));
      const hay = formulas && typeof raw === 'string' && raw.startsWith('=') ? raw : shown;
      if (!re.test(hay)) continue;
      total++;
      if (matches.length < limit) matches.push({ tab: t.name, cell: addr(r, c), value: plainValue(v), text: shown, ...(typeof raw === 'string' && raw.startsWith('=') ? { formula: raw } : {}) });
    }
  }
  return { query: String(query), matches, total };
}

// ---------- sort and filter ----------

// Which column, by letter (B), by number (2, counting from 1) or by its header text ("Amount").
function columnOf(doc, engine, g, ref, header) {
  if (typeof ref === 'number') return g.c1 + ref - 1;
  const s = String(ref ?? '').trim();
  if (/^[A-Za-z]{1,3}$/.test(s) && colIndex(s) >= g.c1 && colIndex(s) <= g.c2 && !headerMatch(doc, engine, g, s, header)) return colIndex(s);
  const byHeader = headerMatch(doc, engine, g, s, header);
  if (byHeader !== null) return byHeader;
  if (/^[A-Za-z]{1,3}$/.test(s)) return colIndex(s);
  throw new BookError(`No column "${s}" in ${rangeText(g, g.tab.name)}. Use a letter (B) or a header from the first row.`);
}
function headerMatch(doc, engine, g, s, header) {
  if (!header) return null;
  for (let c = g.c1; c <= g.c2; c++) {
    const v = engine.value(g.tab.id, g.r1, c);
    if (v !== null && String(plainValue(v)).trim().toLowerCase() === s.toLowerCase()) return c;
  }
  return null;
}

// Looks like a header row: the first row is all text and some column below it is not.
export function looksLikeHeader(doc, engine, g) {
  if (g.r2 <= g.r1) return false;
  let text = 0, filled = 0;
  for (let c = g.c1; c <= g.c2; c++) {
    const v = engine.value(g.tab.id, g.r1, c);
    if (v === null || v === '') continue;
    filled++;
    if (typeof v === 'string') text++;
  }
  if (!filled || text !== filled) return false;
  for (let c = g.c1; c <= g.c2; c++) if (typeof engine.value(g.tab.id, g.r1 + 1, c) === 'number') return true;
  return text >= 2;
}

const rankOf = (v) => (v === null || v === '' ? 4 : typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : typeof v === 'boolean' ? 2 : 3);
export function compareValues(a, b) {
  const ra = rankOf(a), rb = rankOf(b);
  if (ra !== rb) return ra - rb;
  if (ra === 0) return a - b;
  if (ra === 1) return a.localeCompare(b, undefined, { sensitivity: 'base', numeric: true });
  if (ra === 2) return Number(a) - Number(b);
  return 0;
}

// Sorts the rows of a range by one or more columns. Cells outside the range's columns stay where they
// are. Each moved formula keeps its relative references, as in Excel and Google Sheets.
export function sortRange(doc, engine, { tab, range = null, by, header = null }) {
  const g0 = resolve(doc, tab, range);
  const hasHeader = header ?? looksLikeHeader(doc, engine, g0);
  const keys = (Array.isArray(by) ? by : [by]).filter(Boolean).map((b) => (typeof b === 'object' ? b : { column: b }));
  if (!keys.length) throw new BookError('Say which column to sort by.');
  const cols = keys.map((k) => ({ c: columnOf(doc, engine, g0, k.column, hasHeader), desc: /^desc/i.test(k.order ?? 'asc') }));
  const g = { ...g0, r1: g0.r1 + (hasHeader ? 1 : 0) };
  if (g.r2 < g.r1) return { tab: g.tab.name, range: rangeText(g0), moved: 0, header: hasHeader };
  const rows = [];
  for (let r = g.r1; r <= g.r2; r++) rows.push({ r, k: cols.map(({ c }) => plainValue(engine.value(g.tab.id, r, c))) });
  rows.sort((a, b) => {
    for (let i = 0; i < cols.length; i++) {
      const x = a.k[i], y = b.k[i];
      const empty = (v) => v === null || v === '';
      if (empty(x) || empty(y)) { if (empty(x) && empty(y)) continue; return empty(x) ? 1 : -1; }
      const d = compareValues(x, y);
      if (d) return cols[i].desc ? -d : d;
    }
    return a.r - b.r;
  });
  const snap = new Map();
  for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) snap.set(key(r, c), [getRaw(doc, g.tab.id, r, c), getFmt(doc, g.tab.id, r, c)]);
  let moved = 0;
  doc.transact(() => {
    rows.forEach((row, i) => {
      const to = g.r1 + i;
      if (row.r !== to) moved++;
      for (let c = g.c1; c <= g.c2; c++) {
        const [raw, fmt] = snap.get(key(row.r, c));
        setRaw(doc, g.tab.id, to, c, typeof raw === 'string' && raw.startsWith('=') ? shiftFormula(raw, to - row.r, 0) : raw);
        setFmt(doc, g.tab.id, to, c, fmt);
      }
    });
  });
  return { tab: g.tab.name, range: rangeText(g0), header: hasHeader, sorted_by: cols.map((x) => `${colName(x.c)} ${x.desc ? 'descending' : 'ascending'}`), moved };
}

export const FILTER_OPS = ['eq', 'neq', 'gt', 'gte', 'lt', 'lte', 'contains', 'not_contains', 'starts_with', 'empty', 'not_empty', 'in'];

export function matchCondition(v, cond) {
  const val = plainValue(v);
  const s = val === null ? '' : String(val).toLowerCase();
  const target = cond.value;
  const num = (x) => (typeof x === 'number' ? x : x !== null && x !== '' && !Number.isNaN(Number(x)) ? Number(x) : null);
  switch (cond.op) {
    case 'eq': return num(val) !== null && num(target) !== null ? num(val) === num(target) : s === String(target ?? '').toLowerCase();
    case 'neq': return !matchCondition(v, { ...cond, op: 'eq' });
    case 'gt': return num(val) !== null && num(val) > num(target);
    case 'gte': return num(val) !== null && num(val) >= num(target);
    case 'lt': return num(val) !== null && num(val) < num(target);
    case 'lte': return num(val) !== null && num(val) <= num(target);
    case 'contains': return s.includes(String(target ?? '').toLowerCase());
    case 'not_contains': return !s.includes(String(target ?? '').toLowerCase());
    case 'starts_with': return s.startsWith(String(target ?? '').toLowerCase());
    case 'empty': return s === '';
    case 'not_empty': return s !== '';
    case 'in': return (cond.values ?? []).map((x) => String(x ?? '').toLowerCase()).includes(s);
    default: return true;
  }
}

// The rows a tab's filter hides (0-based), for the screen and for readers.
export function hiddenRows(doc, engine, tab) {
  const f = tab.filter;
  if (!f?.range || !f.criteria || !Object.keys(f.criteria).length) return new Set();
  let g;
  try { g = { tab, ...parseRange(f.range) }; } catch { return new Set(); }
  const hidden = new Set();
  const last = Math.min(g.r2 === Infinity ? (usedRange(doc, tab.id)?.r2 ?? 0) : g.r2, (usedRange(doc, tab.id)?.r2 ?? 0));
  for (let r = g.r1 + 1; r <= last; r++) {
    for (const [c, cond] of Object.entries(f.criteria)) {
      if (!matchCondition(engine.value(tab.id, r, Number(c)), cond)) { hidden.add(r); break; }
    }
  }
  return hidden;
}

// Turns on the tab's filter (everyone sees the same filter, like a filter in Google Sheets), adds or
// replaces the condition on one column, and returns the rows that show.
export function setFilter(doc, engine, { tab, range = null, column = null, condition = null, clear = false }) {
  const t0 = needTab(doc, range && parseRange(range).tab ? parseRange(range).tab : tab);
  if (clear) { setTabProp(doc, t0.id, 'filter', null); return { tab: t0.name, filter: null, shown: null }; }
  const g = t0.filter?.range && !range ? { tab: t0, ...parseRange(t0.filter.range) } : resolve(doc, t0.id, range);
  if (g.r2 === Infinity) g.r2 = usedRange(doc, t0.id)?.r2 ?? g.r1;
  const f = { range: rangeText(g), criteria: { ...(t0.filter?.range === rangeText(g) ? t0.filter.criteria : {}) } };
  if (column !== null && column !== undefined) {
    const c = columnOf(doc, engine, g, column, true);
    if (!condition || condition.op === 'all') delete f.criteria[c];
    else {
      if (!FILTER_OPS.includes(condition.op)) throw new BookError(`op is one of ${FILTER_OPS.join(', ')}.`);
      f.criteria[c] = condition;
    }
  }
  setTabProp(doc, t0.id, 'filter', JSON.stringify(f));
  const t = needTab(doc, t0.id);
  const hidden = hiddenRows(doc, engine, t);
  const shown = [];
  for (let r = g.r1; r <= g.r2 && shown.length < 500; r++) {
    if (hidden.has(r)) continue;
    const row = [];
    for (let c = g.c1; c <= g.c2; c++) row.push(plainValue(engine.value(t.id, r, c)));
    shown.push({ row: r + 1, values: row });
  }
  return { tab: t.name, filter: { range: f.range, criteria: Object.fromEntries(Object.entries(f.criteria).map(([c, v]) => [colName(Number(c)), v])) }, shown, hidden: hidden.size };
}

// ---------- formatting ----------

const COLOR = /^#[0-9a-f]{6}$/i;
export function formatRange(doc, engine, { tab, range, bold, italic, underline, strike, color, fill, align, valign, wrap, number_format, font_size, clear = false }) {
  const g = resolve(doc, tab, range, { open: 'used', allowEmpty: false });
  if (size(g) > MAX_CELLS_WRITE) throw new BookError(`Format at most ${MAX_CELLS_WRITE} cells at a time.`);
  for (const [n, v] of [['color', color], ['fill', fill]]) if (v && !COLOR.test(v)) throw new BookError(`${n} is a colour like #1f6feb.`);
  let nf;
  if (number_format !== undefined) {
    const preset = FORMATS.find((f) => f.id === number_format);
    nf = preset ? preset.nf : number_format;
  }
  const patch = { b: bold, i: italic, u: underline, s: strike, fc: color, bg: fill, ha: align, va: valign, wrap, nf, fs: font_size };
  doc.transact(() => {
    for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) {
      if (clear) { setFmt(doc, g.tab.id, r, c, null); continue; }
      const cur = { ...(getFmt(doc, g.tab.id, r, c) ?? {}) };
      for (const [k, v] of Object.entries(patch)) if (v !== undefined) { if (v === null || v === false || v === '') delete cur[k]; else cur[k] = v; }
      setFmt(doc, g.tab.id, r, c, cur);
    }
  });
  return { tab: g.tab.name, range: rangeText(g), cells: size(g) };
}

export function setColumnWidth(doc, { tab, columns, width }) {
  const g = resolve(doc, tab, /^[A-Za-z]+$/.test(columns) ? `${columns}:${columns}` : columns, { open: 'tab' });
  doc.transact(() => { for (let c = g.c1; c <= g.c2; c++) setTabProp(doc, g.tab.id, `w:${c}`, width ? Math.max(24, Math.min(800, Math.round(width))) : null); });
  return { tab: g.tab.name, columns: `${colName(g.c1)}:${colName(g.c2)}`, width: width ?? null };
}

export function setFreeze(doc, { tab, rows, columns }) {
  const t = needTab(doc, tab);
  doc.transact(() => {
    if (rows !== undefined) setTabProp(doc, t.id, 'fr', Math.max(0, Math.min(rows, 50)));
    if (columns !== undefined) setTabProp(doc, t.id, 'fc', Math.max(0, Math.min(columns, 26)));
  });
  const n = needTab(doc, t.id);
  return { tab: n.name, frozen_rows: n.frozen_rows, frozen_cols: n.frozen_cols };
}

// ---------- tabs ----------

export function addSheetTab(doc, { name, index, rows, cols }) {
  const id = addTab(doc, { name, index, rows, cols });
  return needTab(doc, id);
}

export function renameTab(doc, engine, { tab, name }) {
  const t = needTab(doc, tab);
  const n = checkTabName(doc, name, t.id);
  if (n === t.name) return needTab(doc, t.id);
  // HyperFormula rewrites every formula that names the tab; those formulas are written back with the new name.
  engine.structural((hf) => hf.renameSheet(engine.sheet(t.id), n));
  doc.transact(() => tabMap(doc, t.id).set('name', n), 'engine');
  engine.build();
  return needTab(doc, t.id);
}

export function deleteTab(doc, engine, { tab }) {
  const t = needTab(doc, tab);
  if (tabList(doc).length === 1) throw new BookError('A spreadsheet keeps at least one tab.');
  engine.structural((hf) => hf.removeSheet(engine.sheet(t.id)));
  removeTab(doc, t.id);
  return { deleted: t.name, tabs: tabList(doc).map((x) => x.name) };
}

export function moveTab(doc, { tab, index }) {
  const t = needTab(doc, tab);
  doc.transact(() => {
    const order = doc.getArray('order');
    const i = order.toArray().indexOf(t.id);
    order.delete(i, 1);
    order.insert(Math.max(0, Math.min(index, order.length)), [t.id]);
  });
  return { tabs: tabList(doc).map((x) => x.name) };
}

// ---------- inserting and deleting rows and columns ----------

function shiftKeyed(map, axis, at, by) {
  const moves = [];
  for (const [k, v] of map.entries()) {
    const [r, c] = unkey(k);
    const n = axis === 'row' ? r : c;
    if (n < at) continue;
    if (by < 0 && n < at - by) { moves.push([k, null, null]); continue; }
    moves.push([k, axis === 'row' ? key(r + by, c) : key(r, c + by), v]);
  }
  for (const [k] of moves) map.delete(k);
  for (const [, nk, v] of moves) if (nk) map.set(nk, v);
}

export function insertDelete(doc, engine, { tab, axis, at, count, del = false }) {
  const t = needTab(doc, tab);
  if (!Number.isInteger(at) || at < 1) throw new BookError(`${axis === 'row' ? 'Rows' : 'Columns'} count from 1.`);
  if (!Number.isInteger(count) || count < 1 || count > 5000) throw new BookError('count is 1 to 5000.');
  const i = at - 1;
  const sheet = engine.sheet(t.id);
  engine.structural((hf) => {
    if (axis === 'row') del ? hf.removeRows(sheet, [i, count]) : hf.addRows(sheet, [i, count]);
    else del ? hf.removeColumns(sheet, [i, count]) : hf.addColumns(sheet, [i, count]);
  });
  doc.transact(() => {
    const by = del ? -count : count;
    shiftKeyed(fmtMap(doc, t.id), axis, i, by);
    const m = tabMap(doc, t.id);
    if (axis === 'row') m.set('rows', Math.max(1, Math.min(MAX_ROWS, (m.get('rows') ?? 1000) + by)));
    else m.set('cols', Math.max(1, Math.min(MAX_COLS, (m.get('cols') ?? 26) + by)));
    // Widths follow their columns.
    if (axis === 'col') {
      const ws = [...m.entries()].filter(([k]) => k.startsWith('w:')).map(([k, v]) => [Number(k.slice(2)), v]);
      for (const [c] of ws) m.delete(`w:${c}`);
      for (const [c, w] of ws) { const n = c < i ? c : del ? (c < i + count ? null : c - count) : c + count; if (n !== null) m.set(`w:${n}`, w); }
    }
    for (const [id, cm] of doc.getMap('comments').entries()) {
      if (cm.tab !== t.id) continue;
      const n = axis === 'row' ? cm.r : cm.c;
      if (n < i) continue;
      if (del && n < i + count) { doc.getMap('comments').delete(id); continue; }
      doc.getMap('comments').set(id, { ...cm, [axis === 'row' ? 'r' : 'c']: n + (del ? -count : count) });
    }
  });
  const n = needTab(doc, t.id);
  return { tab: n.name, [del ? 'deleted' : 'inserted']: count, at, rows: n.rows, cols: n.cols };
}

// ---------- comments ----------

export function addComment(doc, me, { tab, cell, body }) {
  const g = resolve(doc, tab, cell, { open: 'tab' });
  const text = String(body ?? '').trim();
  if (!text) throw new BookError('A comment needs some words.');
  const id = newCommentId();
  const c = { tab: g.tab.id, r: g.r1, c: g.c1, by: { id: me.id, name: me.name }, body: text.slice(0, 5000), at: new Date().toISOString(), resolved: false, replies: [] };
  doc.getMap('comments').set(id, c);
  return commentView(doc, id, c);
}

export function replyComment(doc, me, { comment, body }) {
  const c = doc.getMap('comments').get(comment);
  if (!c) throw new BookError(`No comment ${comment}.`, 404, 'not_found');
  const text = String(body ?? '').trim();
  if (!text) throw new BookError('A reply needs some words.');
  const n = { ...c, replies: [...(c.replies ?? []), { by: { id: me.id, name: me.name }, body: text.slice(0, 5000), at: new Date().toISOString() }] };
  doc.getMap('comments').set(comment, n);
  return commentView(doc, comment, n);
}

export function resolveComment(doc, { comment, resolved = true }) {
  const c = doc.getMap('comments').get(comment);
  if (!c) throw new BookError(`No comment ${comment}.`, 404, 'not_found');
  const n = { ...c, resolved: !!resolved };
  doc.getMap('comments').set(comment, n);
  return commentView(doc, comment, n);
}

export function deleteComment(doc, me, { comment }) {
  const c = doc.getMap('comments').get(comment);
  if (!c) throw new BookError(`No comment ${comment}.`, 404, 'not_found');
  if (c.by?.id !== me.id && !['owner', 'admin'].includes(me.role)) throw new BookError('Only the person who wrote a comment (or an admin) can delete it.', 403, 'forbidden');
  doc.getMap('comments').delete(comment);
  return { deleted: comment };
}

export function commentView(doc, id, c) {
  const t = findTab(doc, c.tab);
  return { id, tab: t?.name ?? null, cell: addr(c.r, c.c), by: c.by, body: c.body, at: c.at, resolved: !!c.resolved, replies: c.replies ?? [] };
}

export function listComments(doc, { tab = null, include_resolved = false } = {}) {
  const t = tab ? needTab(doc, tab) : null;
  return allComments(doc).filter((c) => (!t || c.tab === t.id) && (include_resolved || !c.resolved)).map((c) => commentView(doc, c.id, c));
}

// ---------- charts ----------

export const CHART_TYPES = ['column', 'bar', 'line', 'area', 'pie', 'scatter'];

export function createChart(doc, engine, { tab, type = 'column', range, title = null, at = null, stacked = false }) {
  if (!CHART_TYPES.includes(type)) throw new BookError(`type is one of ${CHART_TYPES.join(', ')}.`);
  const g = resolve(doc, tab, range, { allowEmpty: false });
  if (g.c2 === g.c1 && type !== 'pie' && g.r2 === g.r1) throw new BookError('A chart needs at least two cells of data.');
  const place = at ? resolve(doc, g.tab.id, at, { open: 'tab' }) : { r1: g.r1, c1: g.c2 + 2 };
  const id = newChartId();
  const c = { tab: g.tab.id, type, range: rangeText(g), title: title ? String(title).slice(0, 120) : null, r: place.r1, c: place.c1, w: 480, h: 300, stacked: !!stacked };
  doc.getMap('charts').set(id, c);
  return chartView(doc, engine, id, c);
}

export function updateChart(doc, engine, { chart, type, range, title, at, stacked, width, height }) {
  const c = doc.getMap('charts').get(chart);
  if (!c) throw new BookError(`No chart ${chart}.`, 404, 'not_found');
  const n = { ...c };
  if (type !== undefined) { if (!CHART_TYPES.includes(type)) throw new BookError(`type is one of ${CHART_TYPES.join(', ')}.`); n.type = type; }
  if (range !== undefined) n.range = rangeText(resolve(doc, c.tab, range, { allowEmpty: false }));
  if (title !== undefined) n.title = title ? String(title).slice(0, 120) : null;
  if (stacked !== undefined) n.stacked = !!stacked;
  if (at !== undefined && at !== null) { const p = resolve(doc, c.tab, at, { open: 'tab' }); n.r = p.r1; n.c = p.c1; }
  if (width) n.w = Math.max(200, Math.min(1400, width));
  if (height) n.h = Math.max(160, Math.min(900, height));
  doc.getMap('charts').set(chart, n);
  return chartView(doc, engine, chart, n);
}

export function deleteChart(doc, { chart }) {
  if (!doc.getMap('charts').get(chart)) throw new BookError(`No chart ${chart}.`, 404, 'not_found');
  doc.getMap('charts').delete(chart);
  return { deleted: chart };
}

// The chart's data as series, worked out from its range: the first column is the labels, each other
// column a series named by its header (when the first row is a header).
export function chartData(doc, engine, c) {
  const t = findTab(doc, c.tab);
  if (!t) return { labels: [], series: [] };
  let g;
  try { g = { tab: t, ...parseRange(c.range) }; } catch { return { labels: [], series: [] }; }
  if (g.r2 === Infinity) g.r2 = usedRange(doc, t.id)?.r2 ?? g.r1;
  const hasHeader = looksLikeHeader(doc, engine, g) || (g.r2 > g.r1 && typeof plainValue(engine.value(t.id, g.r1, g.c2)) === 'string');
  const r0 = g.r1 + (hasHeader ? 1 : 0);
  const single = g.c1 === g.c2;
  const hidden = hiddenRows(doc, engine, t);
  const labels = [], series = [];
  for (let c = single ? g.c1 : g.c1 + 1; c <= g.c2; c++) series.push({ name: hasHeader ? String(plainValue(engine.value(t.id, g.r1, c)) ?? colName(c)) : colName(c), values: [] });
  for (let r = r0; r <= g.r2; r++) {
    if (hidden.has(r)) continue;
    const lab = single ? String(r + 1) : display(engine.value(t.id, r, g.c1), getFmt(doc, t.id, r, g.c1)?.nf, engine.type(t.id, r, g.c1));
    labels.push(lab);
    series.forEach((s, i) => { const v = plainValue(engine.value(t.id, r, (single ? g.c1 : g.c1 + 1) + i)); s.values.push(typeof v === 'number' ? v : null); });
  }
  const nf = getFmt(doc, t.id, r0, single ? g.c1 : g.c1 + 1)?.nf ?? null;
  return { labels, series, format: nf };
}

export function chartView(doc, engine, id, c) {
  const t = findTab(doc, c.tab);
  return { id, tab: t?.name ?? null, type: c.type, range: c.range, title: c.title, at: addr(c.r, c.c), width: c.w, height: c.h, stacked: !!c.stacked, data: chartData(doc, engine, c) };
}

export const listCharts = (doc, engine, { tab = null } = {}) => {
  const t = tab ? needTab(doc, tab) : null;
  return allCharts(doc).filter((c) => !t || c.tab === t.id).map((c) => chartView(doc, engine, c.id, c));
};

// ---------- pivot tables ----------

const SUMMARIES = { sum: 'SUMIFS', count: 'COUNTIFS', average: 'AVERAGEIFS', min: 'MINIFS', max: 'MAXIFS' };

// A pivot table on its own tab, written as live formulas (SUMIFS, COUNTIFS, ...) over the source, so it
// stays right as the source changes and survives export to Excel. New categories appear on refresh
// (sheets.refresh_link), which rebuilds it from the same definition.
export function createPivot(doc, engine, { tab, range = null, rows, columns = null, values = [{ summarize: 'count' }], target = null, def = null }) {
  const spec = def ?? { tab: needTab(doc, tab).id, range, rows, columns, values };
  const src = resolve(doc, spec.tab, spec.range);
  if (!looksLikeHeader(doc, engine, src) && src.r2 > src.r1) {
    // Still allow it: the first row is taken as the header either way.
  }
  const headerName = (c) => String(plainValue(engine.value(src.tab.id, src.r1, c)) ?? colName(c));
  const rowCol = columnOf(doc, engine, src, spec.rows, true);
  const colCol = spec.columns ? columnOf(doc, engine, src, spec.columns, true) : null;
  const vals = (spec.values?.length ? spec.values : [{ summarize: 'count' }]).map((v) => {
    const how = String(v.summarize ?? 'sum').toLowerCase();
    if (!SUMMARIES[how]) throw new BookError(`summarize is one of ${Object.keys(SUMMARIES).join(', ')}.`);
    return { how, c: v.field !== undefined && v.field !== null ? columnOf(doc, engine, src, v.field, true) : rowCol };
  });
  const uniq = (c) => {
    const seen = new Map();
    for (let r = src.r1 + 1; r <= src.r2; r++) {
      const v = plainValue(engine.value(src.tab.id, r, c));
      if (v === null || v === '') continue;
      const k = typeof v === 'string' ? v.toLowerCase() : String(v);
      if (!seen.has(k)) seen.set(k, v);
    }
    return [...seen.values()].sort(compareValues);
  };
  const rowKeys = uniq(rowCol);
  const colKeys = colCol !== null ? uniq(colCol) : [null];
  const abs = (c, r1, r2) => `${quoteTab(src.tab.name)}!$${colName(c)}$${r1 + 1}:$${colName(c)}$${r2 + 1}`;
  const R1 = src.r1 + 1, R2 = Math.max(src.r1 + 1, src.r2);
  const lit = (v) => (typeof v === 'number' ? String(v) : `"${String(v).replace(/"/g, '""')}"`);
  const crit = (c, ref) => `,${abs(c, R1, R2)},${ref}`;
  // One formula per cell: the label cells are referenced, so a renamed category label still adds up.
  const formula = (v, rowRef, colRef) => {
    const fn = SUMMARIES[v.how];
    const range = v.how === 'count' ? '' : `${abs(v.c, R1, R2)}`;
    const conds = `${rowRef ? crit(rowCol, rowRef) : ''}${colRef ? crit(colCol, colRef) : ''}`;
    if (v.how === 'count') return conds ? `=COUNTIFS(${conds.slice(1)})` : `=COUNTA(${abs(rowCol, R1, R2)})`;
    if (!conds) return `=${{ sum: 'SUM', average: 'AVERAGE', min: 'MIN', max: 'MAX' }[v.how]}(${range})`;
    return `=${fn}(${range}${conds})`;
  };
  // Lay it out: labels down column A, one block of value columns per column category.
  const out = [];
  const head = [headerName(rowCol)];
  for (const ck of colKeys) for (const v of vals) head.push(`${v.how === 'count' ? 'Count' : `${v.how[0].toUpperCase()}${v.how.slice(1)} of ${headerName(v.c)}`}${ck !== null ? ` · ${ck}` : ''}`);
  out.push(head);
  if (colCol !== null) {
    // A hidden-by-convention row of column keys is not used: the keys are written into the formulas.
  }
  rowKeys.forEach((rk, i) => {
    const row = [typeof rk === 'string' ? textRaw(rk) : rk];
    for (const ck of colKeys) for (const v of vals) row.push(formula(v, `$A${i + 2}`, ck !== null ? lit(ck) : null));
    out.push(row);
  });
  const total = ['Grand total'];
  for (const ck of colKeys) for (const v of vals) total.push(formula(v, null, ck !== null ? lit(ck) : null));
  out.push(total);

  let t = def?.target ? findTab(doc, def.target) : null;
  doc.transact(() => {
    if (!t) {
      const name = target ? checkTabName(doc, target) : uniquePivotName(doc, src.tab.name);
      t = needTab(doc, addTab(doc, { name, rows: Math.max(100, out.length + 10), cols: Math.max(10, out[0].length + 2) }));
    } else {
      cellMap(doc, t.id).clear();
      fmtMap(doc, t.id).clear();
    }
    out.forEach((row, r) => row.forEach((v, c) => setRaw(doc, t.id, r, c, v)));
    for (let c = 0; c < out[0].length; c++) setFmt(doc, t.id, 0, c, { b: true });
    for (let c = 0; c < out[0].length; c++) setFmt(doc, t.id, out.length - 1, c, { b: true });
    const nf = getFmt(doc, src.tab.id, src.r1 + 1, vals[0].c)?.nf;
    if (nf) for (let r = 1; r < out.length; r++) for (let c = 1; c < out[0].length; c++) if (vals[(c - 1) % vals.length].how !== 'count') setFmt(doc, t.id, r, c, { ...(getFmt(doc, t.id, r, c) ?? {}), nf });
    setTabProp(doc, t.id, 'fr', 1);
    setTabProp(doc, t.id, 'link', JSON.stringify({ kind: 'pivot', def: { ...spec, target: t.id } }));
  });
  return { tab: needTab(doc, t.id).name, range: rangeText({ r1: 0, c1: 0, r2: out.length - 1, c2: out[0].length - 1 }), rows: rowKeys.length, columns: head.length - 1, ...after(doc, engine, needTab(doc, t.id), { r1: 0, c1: 0, r2: out.length - 1, c2: out[0].length - 1 }) };
}

function uniquePivotName(doc, base) {
  const names = new Set(tabList(doc).map((t) => t.name.toLowerCase()));
  for (let i = 1; ; i++) { const n = `Pivot of ${base}${i > 1 ? ` ${i}` : ''}`.slice(0, 100); if (!names.has(n.toLowerCase())) return n; }
}

// ---------- the whole workbook ----------

export function bookSummary(doc, engine) {
  return {
    title: doc.getMap('meta').get('title') ?? 'Untitled spreadsheet',
    tabs: tabList(doc).map((t) => {
      const u = usedRange(doc, t.id);
      return {
        id: t.id, name: t.name, rows: t.rows, cols: t.cols, used: u ? rangeText(u) : null,
        frozen_rows: t.frozen_rows, frozen_cols: t.frozen_cols,
        filter: t.filter ? { range: t.filter.range, criteria: Object.fromEntries(Object.entries(t.filter.criteria ?? {}).map(([c, v]) => [colName(Number(c)), v])) } : null,
        link: t.link ? { kind: t.link.kind, ...(t.link.kind === 'crm' ? { tool: t.link.tool, input: t.link.input, refreshed_at: t.link.refreshed_at ?? null } : { source: findTab(doc, t.link.def?.tab)?.name ?? null }) } : null,
      };
    }),
    charts: allCharts(doc).map((c) => ({ id: c.id, tab: findTab(doc, c.tab)?.name ?? null, type: c.type, range: c.range, title: c.title })),
    comments: allComments(doc).filter((c) => !c.resolved).length,
  };
}

export { isDateFormat, FMT_KEYS };
