// A workbook is one Yjs document. The same helpers read and change it on the server and in the browser.
//
//   meta      Y.Map    title
//   order     Y.Array  tab ids, left to right
//   tabs      Y.Map    tab id -> Y.Map { name, rows, cols, fr (frozen rows), fc (frozen columns), color,
//                                         filter (JSON), link (JSON), w:<col> widths, h:<row> heights }
//   cells     Y.Map    tab id -> Y.Map { "r,c" -> raw }   raw: number, boolean, "=formula", text, "'text"
//   fmt       Y.Map    tab id -> Y.Map { "r,c" -> { b, i, u, s, fc, bg, ha, wrap, nf, fs } }
//   comments  Y.Map    comment id -> { tab, r, c, by, body, at, resolved, replies: [{ by, body, at }] }
//   charts    Y.Map    chart id -> { tab, type, range, title, r, c, w, h }
import * as Y from 'yjs';
import { key, unkey } from './a1.mjs';

export { Y };
export const DEFAULT_ROWS = 1000;
export const DEFAULT_COLS = 26;
export const MAX_ROWS = 100000;
export const MAX_COLS = 702; // A to ZZ

const rid = (p) => `${p}${Math.random().toString(36).slice(2, 9)}`;
export const newTabId = () => rid('t');
export const newCommentId = () => rid('cm');
export const newChartId = () => rid('ch');

export const FMT_KEYS = ['b', 'i', 'u', 's', 'fc', 'bg', 'ha', 'va', 'wrap', 'nf', 'fs'];

export class BookError extends Error {
  constructor(message, status = 400, code = 'bad_input') { super(message); this.status = status; this.code = code; }
}

export function initBook(doc, { title = 'Untitled spreadsheet', tabs = [{ name: 'Sheet1' }] } = {}) {
  doc.transact(() => {
    doc.getMap('meta').set('title', title);
    for (const t of tabs) addTab(doc, t);
  });
}

export const title = (doc) => doc.getMap('meta').get('title') ?? 'Untitled spreadsheet';

export function tabList(doc) {
  const tabs = doc.getMap('tabs');
  return doc.getArray('order').toArray().map((id) => tabView(id, tabs.get(id))).filter(Boolean);
}

function tabView(id, m) {
  if (!m) return null;
  const widths = {}, heights = {};
  for (const [k, v] of m.entries()) {
    if (k.startsWith('w:')) widths[k.slice(2)] = v;
    else if (k.startsWith('h:')) heights[k.slice(2)] = v;
  }
  const parse = (v) => { try { return v ? JSON.parse(v) : null; } catch { return null; } };
  return {
    id, name: m.get('name'), rows: m.get('rows') ?? DEFAULT_ROWS, cols: m.get('cols') ?? DEFAULT_COLS,
    frozen_rows: m.get('fr') ?? 0, frozen_cols: m.get('fc') ?? 0, color: m.get('color') ?? null,
    filter: parse(m.get('filter')), link: parse(m.get('link')), widths, heights,
  };
}

// A tab by id or by name (any case). Without one, the first tab.
export function findTab(doc, ref) {
  const all = tabList(doc);
  if (ref === undefined || ref === null || ref === '') return all[0] ?? null;
  const s = String(ref);
  return all.find((t) => t.id === s) ?? all.find((t) => t.name === s) ?? all.find((t) => t.name.toLowerCase() === s.toLowerCase()) ?? null;
}

export function needTab(doc, ref) {
  const t = findTab(doc, ref);
  if (!t) throw new BookError(`No tab called "${ref}". Tabs: ${tabList(doc).map((x) => x.name).join(', ')}.`, 404, 'not_found');
  return t;
}

export function uniqueTabName(doc, base = 'Sheet') {
  const names = new Set(tabList(doc).map((t) => t.name.toLowerCase()));
  if (base !== 'Sheet' && !names.has(base.toLowerCase())) return base;
  for (let i = base === 'Sheet' ? 1 : 2; ; i++) { const n = `${base}${base === 'Sheet' ? '' : ' '}${i}`; if (!names.has(n.toLowerCase())) return n; }
}

export function checkTabName(doc, name, except = null) {
  const n = String(name ?? '').trim();
  if (!n) throw new BookError('A tab needs a name.');
  if (n.length > 100) throw new BookError('Tab names are at most 100 characters.');
  if (/[\[\]*?:/\\]/.test(n)) throw new BookError('Tab names cannot have [ ] * ? : / or \\ in them.');
  if (tabList(doc).some((t) => t.id !== except && t.name.toLowerCase() === n.toLowerCase())) throw new BookError(`There is already a tab called "${n}".`, 409, 'conflict');
  return n;
}

export function addTab(doc, { name, rows = DEFAULT_ROWS, cols = DEFAULT_COLS, index = null, id = newTabId(), frozen_rows = 0, frozen_cols = 0 } = {}) {
  const n = name ? checkTabName(doc, name) : uniqueTabName(doc);
  doc.transact(() => {
    const m = new Y.Map();
    doc.getMap('tabs').set(id, m);
    m.set('name', n);
    m.set('rows', Math.min(Math.max(rows, 1), MAX_ROWS));
    m.set('cols', Math.min(Math.max(cols, 1), MAX_COLS));
    if (frozen_rows) m.set('fr', frozen_rows);
    if (frozen_cols) m.set('fc', frozen_cols);
    doc.getMap('cells').set(id, new Y.Map());
    doc.getMap('fmt').set(id, new Y.Map());
    const order = doc.getArray('order');
    const at = index === null || index === undefined ? order.length : Math.max(0, Math.min(index, order.length));
    order.insert(at, [id]);
  });
  return id;
}

export const tabMap = (doc, id) => doc.getMap('tabs').get(id);
export const cellMap = (doc, id) => doc.getMap('cells').get(id);
export const fmtMap = (doc, id) => doc.getMap('fmt').get(id);

export function setTabProp(doc, id, prop, value) {
  const m = tabMap(doc, id);
  if (value === null || value === undefined || value === 0 && (prop === 'fr' || prop === 'fc')) m.delete(prop);
  else m.set(prop, value);
}

export function removeTab(doc, id) {
  doc.transact(() => {
    const order = doc.getArray('order');
    const i = order.toArray().indexOf(id);
    if (i >= 0) order.delete(i, 1);
    doc.getMap('tabs').delete(id);
    doc.getMap('cells').delete(id);
    doc.getMap('fmt').delete(id);
    for (const [cid, c] of doc.getMap('comments').entries()) if (c.tab === id) doc.getMap('comments').delete(cid);
    for (const [cid, c] of doc.getMap('charts').entries()) if (c.tab === id) doc.getMap('charts').delete(cid);
  });
}

export const getRaw = (doc, tab, r, c) => cellMap(doc, tab)?.get(key(r, c)) ?? null;

export function setRaw(doc, tab, r, c, raw) {
  const m = cellMap(doc, tab);
  const k = key(r, c);
  if (raw === null || raw === undefined || raw === '') { if (m.has(k)) m.delete(k); }
  else if (m.get(k) !== raw) m.set(k, raw);
}

export const getFmt = (doc, tab, r, c) => fmtMap(doc, tab)?.get(key(r, c)) ?? null;

export function setFmt(doc, tab, r, c, fmt) {
  const m = fmtMap(doc, tab);
  const k = key(r, c);
  const clean = fmt ? Object.fromEntries(Object.entries(fmt).filter(([kk, v]) => FMT_KEYS.includes(kk) && v !== null && v !== undefined && v !== false && v !== '')) : null;
  if (!clean || !Object.keys(clean).length) { if (m.has(k)) m.delete(k); return; }
  const cur = m.get(k);
  if (!cur || JSON.stringify(cur) !== JSON.stringify(clean)) m.set(k, clean);
}

// The smallest range holding every cell with something in it (values only), or null for an empty tab.
export function usedRange(doc, tab) {
  let r2 = -1, c2 = -1;
  for (const k of cellMap(doc, tab)?.keys() ?? []) { const [r, c] = unkey(k); if (r > r2) r2 = r; if (c > c2) c2 = c; }
  return r2 < 0 ? null : { r1: 0, c1: 0, r2, c2 };
}

// Every cell of a tab as a 2D array of raw contents, padded to the used size.
export function tabGrid(doc, tab) {
  const u = usedRange(doc, tab);
  if (!u) return [];
  const g = Array.from({ length: u.r2 + 1 }, () => new Array(u.c2 + 1).fill(null));
  for (const [k, v] of cellMap(doc, tab).entries()) { const [r, c] = unkey(k); g[r][c] = v; }
  return g;
}

export const comments = (doc) => [...doc.getMap('comments').entries()].map(([id, c]) => ({ id, ...c })).sort((a, b) => String(a.at).localeCompare(String(b.at)));
export const charts = (doc) => [...doc.getMap('charts').entries()].map(([id, c]) => ({ id, ...c }));

// Base64 helpers that work in Node and the browser.
export function toB64(u8) {
  if (typeof Buffer !== 'undefined') return Buffer.from(u8).toString('base64');
  let s = '';
  for (let i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode.apply(null, u8.subarray(i, i + 0x8000));
  return btoa(s);
}
export function fromB64(s) {
  if (typeof Buffer !== 'undefined') return new Uint8Array(Buffer.from(s, 'base64'));
  const bin = atob(s);
  const u8 = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
  return u8;
}
