// The formula engine for one workbook: HyperFormula, kept in step with the Yjs document. Built once from
// the document, then every change to the document (from this copy or from someone else's) is applied as it
// lands. The server reads values from it for agents and exports; the screen reads values from it to draw.
import { HyperFormula } from 'hyperformula';
import { registerFunctions } from './functions.mjs';
import { tabList, tabGrid, cellMap, Y } from './book.mjs';
import { unkey, key } from './a1.mjs';
import { textRaw } from './input.mjs';

registerFunctions();

export const HF_CONFIG = {
  licenseKey: 'gpl-v3',
  dateFormats: ['YYYY-MM-DD', 'MM/DD/YYYY', 'YYYY/MM/DD'],
  timeFormats: ['hh:mm', 'hh:mm:ss.sss'],
  currencySymbol: ['$'],
  useArrayArithmetic: true,
  evaluateNullToZero: true,
  maxRows: 100000,
  maxColumns: 702,
  precisionRounding: 10,
};

// Stored contents -> what HyperFormula is given. Text always goes in as text (a leading '), so a word
// that looks like a date or a number to HyperFormula stays exactly what was typed.
export function toHf(raw) {
  if (raw === null || raw === undefined) return null;
  if (typeof raw === 'string' && !raw.startsWith('=') && !raw.startsWith("'")) return `'${raw}`;
  return raw;
}

// HyperFormula's serialized contents -> what is stored.
export function fromHf(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'string' && v.startsWith("'")) return textRaw(v.slice(1));
  return v;
}

export class Engine {
  constructor(doc, { live = true } = {}) {
    this.doc = doc;
    this.build();
    if (live) {
      this.onTx = (tr) => this.#afterTransaction(tr);
      doc.on('afterTransaction', this.onTx);
    }
  }

  destroy() {
    if (this.onTx) this.doc.off('afterTransaction', this.onTx);
    this.hf?.destroy();
  }

  build() {
    this.hf?.destroy();
    const tabs = tabList(this.doc);
    const sheets = {};
    for (const t of tabs) sheets[t.name] = tabGrid(this.doc, t.id).map((row) => row.map(toHf));
    const names = [...this.doc.getMap('names').entries()].map(([name, expression]) => ({ name, expression }));
    try { this.hf = HyperFormula.buildFromSheets(sheets, HF_CONFIG, names); } catch { this.hf = HyperFormula.buildFromSheets(sheets, HF_CONFIG); }
    this.ids = new Map(tabs.map((t) => [t.id, this.hf.getSheetId(t.name)]));
    this.maps = new Map(tabs.map((t) => [cellMap(this.doc, t.id), t.id]));
    this.version = (this.version ?? 0) + 1;
  }

  #afterTransaction(tr) {
    if (tr.origin === 'engine') return;
    const order = this.doc.getArray('order');
    const tabs = this.doc.getMap('tabs');
    const cells = this.doc.getMap('cells');
    let rebuild = tr.changed.has(order) || tr.changed.has(tabs) || tr.changed.has(cells) || tr.changed.has(this.doc.getMap('names'));
    if (!rebuild) {
      for (const [type, keys] of tr.changed) if (type instanceof Y.Map && type.parent === tabs && keys.has('name')) { rebuild = true; break; }
    }
    if (rebuild) { this.build(); return; }
    const sets = [];
    for (const [type, keys] of tr.changed) {
      const tab = this.maps.get(type);
      if (!tab) continue;
      const sheet = this.ids.get(tab);
      for (const k of keys) { const [r, c] = unkey(k); sets.push([{ sheet, row: r, col: c }, toHf(type.get(k) ?? null)]); }
    }
    if (!sets.length) return;
    try {
      this.hf.batch(() => { for (const [a, v] of sets) this.hf.setCellContents(a, [[v]]); });
    } catch {
      this.build();
    }
    this.version++;
  }

  sheet(tab) { return this.ids.get(tab); }
  value(tab, r, c) { const s = this.ids.get(tab); return s === undefined ? null : this.hf.getCellValue({ sheet: s, row: r, col: c }); }
  type(tab, r, c) { const s = this.ids.get(tab); return s === undefined ? null : this.hf.getCellValueDetailedType({ sheet: s, row: r, col: c }); }
  // Spilled results of array formulas (FILTER, SORT, UNIQUE) show in cells that hold nothing themselves.
  isSpill(tab, r, c) { const s = this.ids.get(tab); return s !== undefined && this.hf.isCellPartOfArray({ sheet: s, row: r, col: c }) && !this.hf.doesCellHaveFormula({ sheet: s, row: r, col: c }); }
  size(tab) { const s = this.ids.get(tab); return s === undefined ? { height: 0, width: 0 } : this.hf.getSheetDimensions(s); }

  // Runs a HyperFormula operation that moves cells (insert or delete rows and columns), then writes every
  // cell whose contents changed back into the document, in one change.
  structural(fn) {
    fn(this.hf);
    const writes = [];
    for (const t of tabList(this.doc)) {
      const sheet = this.ids.get(t.id);
      const ser = this.hf.getSheetSerialized(sheet);
      const m = cellMap(this.doc, t.id);
      const want = new Map();
      ser.forEach((row, r) => row.forEach((v, c) => { const raw = fromHf(v); if (raw !== null) want.set(key(r, c), raw); }));
      for (const k of m.keys()) if (!want.has(k)) writes.push([m, k, null]);
      for (const [k, v] of want) if (m.get(k) !== v) writes.push([m, k, v]);
    }
    this.doc.transact(() => { for (const [m, k, v] of writes) { if (v === null) m.delete(k); else m.set(k, v); } }, 'engine');
    this.version++;
    return writes.length;
  }

  // Works out a formula without putting it in a cell (used to check a formula before saving it).
  calculate(formula, tab) {
    const s = this.ids.get(tab) ?? 0;
    try { return this.hf.calculateFormula(formula, s); } catch (e) { return { value: '#ERROR!', message: e.message }; }
  }

  validFormula(formula) { return this.hf.validateFormula(formula); }
}
