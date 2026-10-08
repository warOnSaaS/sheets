// Excel (.xlsx) and CSV in and out, with an honest report of what did not carry over.
// ExcelJS (MIT) reads and writes the workbook; fflate (MIT) opens the zip first so we can count what
// ExcelJS does not read (charts, pivot tables, images, macros) and read cell notes ourselves.
import ExcelJS from 'exceljs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { Y, initBook, addTab, tabList, cellMap, fmtMap, setRaw, setFmt, tabMap, BookError, usedRange, getRaw, getFmt, charts as allCharts, comments as allComments, newCommentId } from './book.mjs';
import { Engine } from './engine.mjs';
import { addr, key, unkey, colName, parseRange } from './a1.mjs';
import { textRaw, parseInput, display, plainValue, dateSerial } from './input.mjs';

const MAX_IMPORT_CELLS = 500000;

// Functions newer than Excel 2007 are stored with a prefix in the file (_xlfn.XLOOKUP). We read them
// without it and write them with it, so Excel recognises them.
const XLFN = ['CONCAT', 'IFS', 'MAXIFS', 'MINIFS', 'SWITCH', 'TEXTJOIN', 'XLOOKUP', 'XMATCH', 'UNIQUE', 'SEQUENCE', 'RANK.EQ', 'RANK.AVG', 'STDEV.S', 'STDEV.P', 'VAR.S', 'VAR.P', 'IFNA', 'DAYS', 'ISOWEEKNUM', 'NORM.DIST', 'NORM.INV', 'NORM.S.DIST', 'PERCENTILE.INC', 'PERCENTILE.EXC', 'QUARTILE.INC', 'QUARTILE.EXC', 'MODE.SNGL', 'CEILING.MATH', 'FLOOR.MATH', 'COVARIANCE.P', 'COVARIANCE.S', 'BITAND', 'BITOR', 'BITXOR', 'BITLSHIFT', 'BITRSHIFT', 'FORMULATEXT', 'SHEET', 'SHEETS', 'XOR', 'ARABIC', 'BASE', 'DECIMAL', 'COT', 'COTH', 'CSC', 'CSCH', 'SEC', 'SECH', 'ACOT', 'ACOTH', 'DAYS360', 'GAMMA', 'EXPON.DIST', 'T.DIST', 'T.INV', 'CHISQ.DIST', 'F.DIST', 'BINOM.DIST', 'POISSON.DIST', 'WEIBULL.DIST', 'LOGNORM.DIST', 'GAMMA.DIST', 'BETA.DIST', 'HYPGEOM.DIST', 'NEGBINOM.DIST', 'CONFIDENCE.NORM', 'CONFIDENCE.T', 'Z.TEST', 'T.TEST', 'F.TEST', 'CHISQ.TEST', 'ERF.PRECISE', 'ERFC.PRECISE', 'NETWORKDAYS.INTL', 'WORKDAY.INTL', 'AGGREGATE'];
const XLWS = ['FILTER', 'SORT', 'SORTBY'];
const OURS_ONLY = ['REGEXMATCH', 'REGEXEXTRACT', 'REGEXREPLACE', 'SPLIT', 'COUNTUNIQUE', 'ARRAYFORMULA'];

export const fromExcelFormula = (f) => `=${String(f).replace(/_xlfn\._xlws\.|_xlfn\.|_xlws\./gi, '').replace(/^=/, '')}`;
export function toExcelFormula(f) {
  let s = String(f).slice(1);
  s = s.replace(/(^|[^A-Za-z0-9_.])([A-Za-z][A-Za-z0-9.]*)\(/g, (m, pre, name) => {
    const up = name.toUpperCase();
    if (XLWS.includes(up)) return `${pre}_xlfn._xlws.${up}(`;
    if (XLFN.includes(up)) return `${pre}_xlfn.${up}(`;
    return m;
  });
  return s;
}

const argbToHex = (c) => {
  const a = c?.argb;
  if (!a || typeof a !== 'string') return null;
  const h = a.length === 8 ? a.slice(2) : a.length === 6 ? a : null;
  return h ? `#${h.toLowerCase()}` : null;
};
const hexToArgb = (h) => `FF${h.slice(1).toUpperCase()}`;
const pxFromChars = (w) => Math.round(w * 7 + 5);
const charsFromPx = (px) => Math.max(1, Math.round(((px - 5) / 7) * 100) / 100);

// ---------- import ----------

// Opens the zip, counts what ExcelJS skips, reads the notes, and removes the parts ExcelJS trips on.
function prepare(buffer) {
  let files;
  try { files = unzipSync(new Uint8Array(buffer)); } catch { throw new BookError('That file is not an .xlsx workbook (it is not a zip). Old .xls files need saving as .xlsx first.'); }
  if (!files['xl/workbook.xml']) throw new BookError('That file is not an .xlsx workbook (no xl/workbook.xml).');
  const count = (re) => Object.keys(files).filter((n) => re.test(n)).length;
  const extras = {
    charts: count(/^xl\/charts\/chart\d*\.xml$/),
    pivots: count(/^xl\/pivotTables\/[^/]+\.xml$/),
    images: count(/^xl\/media\//),
    macros: count(/vbaProject\.bin$/),
    externalLinks: count(/^xl\/externalLinks\/[^/]+\.xml$/),
    slicers: count(/^xl\/slicers?\//),
    tables: count(/^xl\/tables\/[^/]+\.xml$/),
  };
  // Which worksheet file is which tab.
  const wbXml = strFromU8(files['xl/workbook.xml']);
  const wbRels = strFromU8(files['xl/_rels/workbook.xml.rels'] ?? new Uint8Array());
  const relTarget = {};
  for (const m of wbRels.matchAll(/<Relationship\b[^>]*>/g)) {
    const id = /Id="([^"]+)"/.exec(m[0])?.[1], t = /Target="([^"]+)"/.exec(m[0])?.[1];
    if (id && t) relTarget[id] = t.startsWith('/') ? t.slice(1) : `xl/${t.replace(/^\.\//, '')}`;
  }
  const sheetFile = {};
  for (const m of wbXml.matchAll(/<sheet\b[^>]*>/g)) {
    const name = /name="([^"]*)"/.exec(m[0])?.[1];
    const rid = /r:id="([^"]+)"/.exec(m[0])?.[1] ?? /\bid="([^"]+)"/.exec(m[0])?.[1];
    if (name && rid && relTarget[rid]) sheetFile[relTarget[rid]] = xmlText(name);
  }
  // Notes: read from each sheet's comments part, then take the comments relationship out of the package.
  const notes = {};
  for (const [file, tab] of Object.entries(sheetFile)) {
    const relsName = file.replace(/([^/]+)$/, '_rels/$1.rels');
    if (!files[relsName]) continue;
    let rels = strFromU8(files[relsName]);
    const dir = file.replace(/[^/]+$/, '');
    for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) {
      const type = /Type="([^"]+)"/.exec(m[0])?.[1] ?? '';
      const target = /Target="([^"]+)"/.exec(m[0])?.[1] ?? '';
      const path = target.startsWith('/') ? target.slice(1) : normalise(dir + target);
      if (/\/comments$/.test(type) && files[path]) {
        const xml = strFromU8(files[path]);
        const authors = [...xml.matchAll(/<author>([\s\S]*?)<\/author>/g)].map((a) => xmlText(a[1]));
        for (const c of xml.matchAll(/<comment\b([^>]*)>([\s\S]*?)<\/comment>/g)) {
          const ref = /ref="([^"]+)"/.exec(c[1])?.[1];
          const author = authors[Number(/authorId="(\d+)"/.exec(c[1])?.[1] ?? 0)] ?? null;
          const text = [...c[2].matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)].map((t) => xmlText(t[1])).join('');
          if (ref) (notes[tab] ??= []).push({ ref, author, text: text.replace(/^[^:\n]{1,60}:\n/, (h) => (author && h.startsWith(author) ? '' : h)) });
        }
      }
    }
    rels = rels.replace(/<Relationship\b[^>]*Type="[^"]*\/(comments|vmlDrawing)"[^>]*\/>/g, '');
    files[relsName] = strToU8(rels);
    // The sheet's <legacyDrawing> points at the removed drawing; drop it too.
    if (files[file]) files[file] = strToU8(strFromU8(files[file]).replace(/<legacyDrawing\b[^>]*\/>/g, ''));
  }
  return { buffer: Buffer.from(zipSync(files)), extras, notes };
}

const normalise = (p) => { const out = []; for (const s of p.split('/')) { if (s === '..') out.pop(); else if (s && s !== '.') out.push(s); } return out.join('/'); };
const xmlText = (s) => String(s).replace(/<[^>]+>/g, '').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#(\d+);/g, (m, n) => String.fromCharCode(Number(n))).replace(/&amp;/g, '&');

const dateToSerial = (d) => dateSerial(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate(), d.getUTCHours(), d.getUTCMinutes(), d.getUTCSeconds());

export async function importXlsx(buffer, { title = null, by = null } = {}) {
  const { buffer: clean, extras, notes } = prepare(buffer);
  const wb = new ExcelJS.Workbook();
  try { await wb.xlsx.load(clean); } catch (e) { throw new BookError(`This workbook could not be read: ${e.message}. Try saving it again from Excel or Google Sheets.`); }
  const doc = new Y.Doc();
  const report = { tabs: 0, cells: 0, formulas: 0, comments: 0, carried: [], dropped: [], formula_problems: [] };
  const drop = new Map();
  const dropped = (what, where, why) => { const d = drop.get(what) ?? { what, count: 0, examples: [], why }; d.count++; if (d.examples.length < 5 && where) d.examples.push(where); drop.set(what, d); };
  const cached = [];
  const carried = new Set(['values']);
  let cells = 0;
  doc.transact(() => {
    doc.getMap('meta').set('title', title || wb.title || 'Imported spreadsheet');
    for (const ws of wb.worksheets) {
      let name = String(ws.name || 'Sheet').replace(/[\[\]*?:/\\]/g, ' ').slice(0, 100).trim() || 'Sheet';
      while (tabList(doc).some((t) => t.name.toLowerCase() === name.toLowerCase())) name = `${name} 2`;
      if (ws.state && ws.state !== 'visible') dropped('Hidden tabs (shown here)', name);
      const id = addTab(doc, { name, rows: Math.max(1000, ws.rowCount + 50), cols: Math.max(26, ws.columnCount + 2) });
      report.tabs++;
      const tm = tabMap(doc, id);
      const view = (ws.views ?? [])[0];
      if (view?.state === 'frozen') {
        if (view.ySplit) tm.set('fr', Math.min(50, view.ySplit));
        if (view.xSplit) tm.set('fc', Math.min(26, view.xSplit));
        carried.add('frozen rows and columns');
      }
      if (ws.properties?.tabColor?.argb) tm.set('color', argbToHex(ws.properties.tabColor));
      for (let c = 1; c <= ws.columnCount; c++) {
        const col = ws.getColumn(c);
        if (col.width) { tm.set(`w:${c - 1}`, pxFromChars(col.width)); carried.add('column widths'); }
        if (col.hidden) dropped('Hidden columns (shown here)', `${name}!${colName(c - 1)}`);
      }
      for (const m of ws.model.merges ?? []) dropped('Merged cells (unmerged here; the value stays in the top-left cell)', `${name}!${m}`);
      if (ws.autoFilter) dropped('Filters (the filter range is not turned on here)', `${name}!${typeof ws.autoFilter === 'string' ? ws.autoFilter : 'a range'}`);
      for (const cf of ws.conditionalFormattings ?? []) dropped('Conditional formatting', `${name}!${cf.ref}`);
      const dv = ws.dataValidations?.model ?? {};
      for (const ref of Object.keys(dv)) dropped('Data validation (drop-down lists and input rules)', `${name}!${ref}`);
      if (ws.getImages?.().length) for (const img of ws.getImages()) dropped('Images', name);
      ws.eachRow({ includeEmpty: false }, (row, rn) => {
        if (row.hidden) dropped('Hidden rows (shown here)', `${name}!${rn}`);
        if (row.height && Math.abs(row.height - 15) > 3) dropped('Custom row heights', `${name}!${rn}`);
        row.eachCell({ includeEmpty: false }, (cell, cn) => {
          if (++cells > MAX_IMPORT_CELLS) throw new BookError(`This workbook has more than ${MAX_IMPORT_CELLS} cells; Sheets takes up to that many for now.`);
          const r = rn - 1, c = cn - 1;
          const at = `${name}!${addr(r, c)}`;
          if (cell.isMerged && cell.master !== cell) return;
          let v = cell.value;
          let raw = null;
          if (v && typeof v === 'object' && !(v instanceof Date)) {
            if (v.formula || v.sharedFormula) {
              const f = cell.formula ?? v.formula;
              if (!f) { dropped('Formulas that could not be read', at); raw = null; }
              else {
                raw = fromExcelFormula(f);
                report.formulas++;
                if (/\[\d+\]/.test(f)) report.formula_problems.push({ cell: at, formula: raw, problem: 'Links to another workbook, which Sheets cannot open' });
                if (v.shareType === 'array' && v.ref) dropped('Legacy array formulas (Ctrl+Shift+Enter): kept as a normal formula', at);
                let res = v.result;
                if (res instanceof Date) res = dateToSerial(res);
                if (res && typeof res === 'object' && res.error) res = res.error;
                cached.push({ tab: id, r, c, at, formula: raw, result: res ?? null });
              }
            } else if (v.richText) {
              raw = textRaw(v.richText.map((t) => t.text).join(''));
              if (v.richText.length > 1) dropped('Mixed formatting inside one cell (the text is kept)', at);
            } else if (v.hyperlink) {
              raw = textRaw(String(v.text?.richText ? v.text.richText.map((t) => t.text).join('') : v.text ?? v.hyperlink));
              dropped('Links in cells (the text is kept)', at);
            } else if (v.error) {
              raw = v.error;
            } else {
              raw = textRaw(JSON.stringify(v));
            }
          } else if (v instanceof Date) {
            raw = dateToSerial(v);
          } else if (typeof v === 'string') {
            raw = textRaw(v);
          } else {
            raw = v;
          }
          if (raw !== null && raw !== undefined) { setRaw(doc, id, r, c, raw); report.cells++; }
          const f = {};
          const font = cell.font ?? {};
          if (font.bold) f.b = true;
          if (font.italic) f.i = true;
          if (font.underline) f.u = true;
          if (font.strike) f.s = true;
          const fc = argbToHex(font.color);
          if (fc && fc !== '#000000') f.fc = fc;
          if (font.size && font.size !== 11) f.fs = font.size;
          if (font.color?.theme !== undefined && font.color.theme > 1 && !fc) dropped('Theme colours (shown in the default colour)', at);
          const fill = cell.fill;
          if (fill?.type === 'pattern' && fill.pattern === 'solid') { const bg = argbToHex(fill.fgColor); if (bg) f.bg = bg; else if (fill.fgColor?.theme !== undefined) dropped('Theme colours (shown in the default colour)', at); }
          else if (fill?.type === 'gradient') dropped('Gradient fills', at);
          const al = cell.alignment ?? {};
          if (['left', 'center', 'right'].includes(al.horizontal)) f.ha = al.horizontal;
          if (['top', 'middle', 'bottom'].includes(al.vertical) && al.vertical !== 'bottom') f.va = al.vertical;
          if (al.wrapText) f.wrap = true;
          if (al.textRotation) dropped('Rotated text', at);
          if (cell.numFmt && cell.numFmt !== 'General') f.nf = cell.numFmt;
          if (cell.border && Object.values(cell.border).some((b) => b && b.style)) dropped('Borders', at);
          if (font.name && !/^(calibri|arial|aptos|helvetica|aptos narrow)$/i.test(font.name)) dropped('Font families (shown in the app font)', at);
          if (Object.keys(f).length) { setFmt(doc, id, r, c, f); if (f.nf) carried.add('number formats'); if (f.b || f.i || f.u || f.s) carried.add('bold, italic, underline and strikethrough'); if (f.fc || f.bg) carried.add('text and fill colours'); if (f.ha || f.va || f.wrap) carried.add('alignment and wrapping'); if (f.fs) carried.add('font sizes'); }
        });
      });
      for (const n of notes[ws.name] ?? []) {
        let p;
        try { p = parseRange(n.ref); } catch { continue; }
        doc.getMap('comments').set(newCommentId(), { tab: id, r: p.r1, c: p.c1, by: { id: null, name: n.author || 'Excel note' }, body: n.text.trim().slice(0, 5000), at: new Date().toISOString(), resolved: false, replies: [] });
        report.comments++;
        carried.add('notes, as comments');
      }
    }
    // Named ranges.
    const names = doc.getMap('names');
    for (const dn of wb.definedNames?.model ?? []) {
      if (!dn.name || /^_xlnm\./.test(dn.name)) continue;
      const ranges = dn.ranges ?? [];
      if (ranges.length !== 1 || dn.localSheetId !== undefined) { dropped('Named ranges with several parts or one tab only', dn.name); continue; }
      names.set(dn.name, `=${ranges[0]}`);
      carried.add('named ranges');
    }
  });
  if (!report.tabs) throw new BookError('That workbook has no tabs.');
  if (report.formulas) carried.add('formulas');
  for (const [what, n] of Object.entries({ Charts: extras.charts, 'Pivot tables (their last values are kept as plain cells)': extras.pivots, 'Macros (VBA)': extras.macros, 'Links to other workbooks': extras.externalLinks, 'Slicers': extras.slicers, 'Excel tables (the cells are kept; the table styling and names are not)': extras.tables })) {
    if (n) drop.set(what, { what, count: n, examples: [], why: null });
  }
  if (extras.images && !drop.has('Images')) drop.set('Images', { what: 'Images', count: extras.images, examples: [] });

  // Work the formulas out here and compare with what Excel saved, so the report says where we differ.
  const engine = new Engine(doc, { live: false });
  const unknown = new Map();
  let same = 0;
  for (const x of cached) {
    const v = plainValue(engine.value(x.tab, x.r, x.c));
    if (typeof v === 'string' && v === '#NAME?') {
      for (const m of x.formula.matchAll(/([A-Za-z][A-Za-z0-9.]*)\(/g)) {
        const fn = m[1].toUpperCase();
        if (!engine.hf.getRegisteredFunctionNames().includes(fn)) unknown.set(fn, (unknown.get(fn) ?? 0) + 1);
      }
      report.formula_problems.push({ cell: x.at, formula: x.formula, problem: 'Uses a function Sheets does not have yet' });
      continue;
    }
    if (x.result === null || x.result === undefined) { same++; continue; }
    if (agree(v, x.result)) same++;
    else if (report.formula_problems.length < 200) report.formula_problems.push({ cell: x.at, formula: x.formula, problem: `Worked out as ${fmtVal(v)}; the file saved ${fmtVal(x.result)}` });
  }
  engine.destroy();
  report.formulas_matching_excel = same;
  report.unsupported_functions = [...unknown.entries()].map(([name, count]) => ({ name, count }));
  report.carried = [...carried];
  report.dropped = [...drop.values()];
  report.summary = summary(report);
  return { doc, report };
}

const fmtVal = (v) => (typeof v === 'string' ? `"${v}"` : String(v));
function agree(a, b) {
  if (typeof a === 'number' && typeof b === 'number') return Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(a), Math.abs(b));
  if (typeof b === 'string' && typeof a === 'string') return a === b;
  if (typeof a === 'boolean' || typeof b === 'boolean') return Boolean(a) === Boolean(b) && typeof a === typeof b;
  if ((a === null || a === '') && (b === '' || b === null || b === 0)) return true;
  return String(a) === String(b);
}

function summary(r) {
  const lost = r.dropped.map((d) => `${d.what.split(' (')[0].toLowerCase()} (${d.count})`);
  return `${r.tabs} tab${r.tabs === 1 ? '' : 's'}, ${r.cells} cells, ${r.formulas} formulas${r.formulas ? ` (${r.formulas_matching_excel} give the same result as the file)` : ''}.${lost.length ? ` Not carried over: ${lost.join(', ')}.` : ' Nothing was left out.'}${r.unsupported_functions.length ? ` Functions Sheets does not have yet: ${r.unsupported_functions.map((f) => f.name).join(', ')}.` : ''}`;
}

// ---------- export ----------

export async function exportXlsx(doc, engine) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'wOS Sheets';
  wb.created = new Date();
  const notCarried = [];
  const tabs = tabList(doc);
  for (const t of tabs) {
    const views = t.frozen_rows || t.frozen_cols ? [{ state: 'frozen', xSplit: t.frozen_cols, ySplit: t.frozen_rows }] : [];
    const ws = wb.addWorksheet(t.name, { views, properties: t.color ? { tabColor: { argb: hexToArgb(t.color) } } : {} });
    for (const [c, w] of Object.entries(t.widths)) ws.getColumn(Number(c) + 1).width = charsFromPx(w);
    const cells = cellMap(doc, t.id), fmts = fmtMap(doc, t.id);
    const keys = new Set([...cells.keys(), ...fmts.keys()]);
    for (const k of keys) {
      const [r, c] = unkey(k);
      const cell = ws.getCell(r + 1, c + 1);
      const raw = cells.get(k) ?? null;
      if (typeof raw === 'string' && raw.startsWith('=')) {
        let res = plainValue(engine.value(t.id, r, c));
        if (typeof res === 'string' && /^#[A-Z0-9/!?]+[!?]?$/.test(res)) res = { error: res };
        cell.value = { formula: toExcelFormula(raw), result: res ?? undefined };
        for (const fn of OURS_ONLY) if (new RegExp(`\\b${fn.replace('.', '\\.')}\\(`, 'i').test(raw)) notCarried.push(`${t.name}!${addr(r, c)} uses ${fn}, which Excel does not have`);
      } else if (typeof raw === 'string') {
        cell.value = raw.startsWith("'") ? raw.slice(1) : raw;
      } else if (raw !== null) {
        cell.value = raw;
      }
      const f = fmts.get(k);
      if (f) {
        const font = {};
        if (f.b) font.bold = true;
        if (f.i) font.italic = true;
        if (f.u) font.underline = true;
        if (f.s) font.strike = true;
        if (f.fc) font.color = { argb: hexToArgb(f.fc) };
        if (f.fs) font.size = f.fs;
        if (Object.keys(font).length) cell.font = font;
        if (f.bg) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToArgb(f.bg) } };
        if (f.ha || f.va || f.wrap) cell.alignment = { ...(f.ha ? { horizontal: f.ha } : {}), ...(f.va ? { vertical: f.va } : {}), ...(f.wrap ? { wrapText: true } : {}) };
        if (f.nf) cell.numFmt = f.nf;
      }
    }
    if (t.filter?.range) {
      try { ws.autoFilter = t.filter.range.replace(/^.*!/, ''); } catch {}
      if (Object.keys(t.filter.criteria ?? {}).length) notCarried.push(`The filter conditions on ${t.name} (the filter range is kept, the conditions are not)`);
    }
    for (const cm of allComments(doc).filter((x) => x.tab === t.id)) {
      const text = [`${cm.by?.name ?? 'Someone'}: ${cm.body}`, ...(cm.replies ?? []).map((x) => `${x.by?.name ?? 'Someone'}: ${x.body}`)].join('\n');
      ws.getCell(cm.r + 1, cm.c + 1).note = text;
    }
  }
  const names = doc.getMap('names');
  for (const [n, expr] of names.entries()) { try { wb.definedNames.add(String(expr).replace(/^=/, ''), n); } catch {} }
  const nCharts = allCharts(doc).length;
  if (nCharts) notCarried.push(`${nCharts} chart${nCharts === 1 ? '' : 's'} (Excel files from Sheets do not include charts yet)`);
  if (allComments(doc).some((c) => c.resolved)) notCarried.push('Resolved comments are written as plain notes');
  const buf = Buffer.from(await wb.xlsx.writeBuffer());
  return { buffer: buf, not_carried: notCarried };
}

// ---------- CSV ----------

export function parseCsv(text, delimiter = null) {
  const s = String(text).replace(/^﻿/, '');
  const first = s.split(/\r?\n/, 1)[0] ?? '';
  const d = delimiter ?? ([',', ';', '\t'].map((x) => [x, first.split(x).length]).sort((a, b) => b[1] - a[1])[0][0]);
  const rows = [];
  let row = [], cell = '', q = false;
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (q) {
      if (ch === '"') { if (s[i + 1] === '"') { cell += '"'; i++; } else q = false; }
      else cell += ch;
    } else if (ch === '"' && cell === '') q = true;
    else if (ch === d) { row.push(cell); cell = ''; }
    else if (ch === '\n' || ch === '\r') { if (ch === '\r' && s[i + 1] === '\n') i++; row.push(cell); rows.push(row); row = []; cell = ''; }
    else cell += ch;
  }
  if (cell !== '' || row.length) { row.push(cell); rows.push(row); }
  return { rows, delimiter: d };
}

export function toCsv(doc, engine, tab, { values = 'formatted' } = {}) {
  const u = usedRange(doc, tab.id);
  if (!u) return '';
  const lines = [];
  for (let r = 0; r <= u.r2; r++) {
    const row = [];
    for (let c = 0; c <= u.c2; c++) {
      const v = engine.value(tab.id, r, c);
      const s = values === 'raw' ? (plainValue(v) ?? '') : display(v, getFmt(doc, tab.id, r, c)?.nf, engine.type(tab.id, r, c));
      const t = String(s);
      row.push(/[",\n\r]/.test(t) ? `"${t.replace(/"/g, '""')}"` : t);
    }
    lines.push(row.join(','));
  }
  return `${lines.join('\r\n')}\r\n`;
}
