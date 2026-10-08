// Cell addresses in A1 style, used the same way on the server and in the browser.
// Rows and columns are 0-based inside the code (row 0 is "1", column 0 is "A").

export function colName(c) {
  let s = '';
  for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

export function colIndex(s) {
  let n = 0;
  for (const ch of String(s).toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

export const addr = (r, c) => `${colName(c)}${r + 1}`;
export const key = (r, c) => `${r},${c}`;
export const unkey = (k) => { const i = k.indexOf(','); return [Number(k.slice(0, i)), Number(k.slice(i + 1))]; };

// A tab name as a formula writes it: plain when it is a simple word, quoted otherwise.
export const quoteTab = (name) => (/^[A-Za-z_][A-Za-z0-9_]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name) ? name : `'${String(name).replace(/'/g, "''")}'`);

export class RangeError2 extends Error {}

// "B3", "A1:C10", "A:C" (whole columns), "2:5" (whole rows), "'Q3 pipeline'!A1:B2", "Data!A1".
// Returns { tab, r1, c1, r2, c2 }, with Infinity for the open end of a whole row or column.
export function parseRange(text) {
  let s = String(text ?? '').trim();
  let tab = null;
  const bang = s.lastIndexOf('!');
  if (bang > 0) {
    tab = s.slice(0, bang);
    if (tab.startsWith("'") && tab.endsWith("'")) tab = tab.slice(1, -1).replace(/''/g, "'");
    s = s.slice(bang + 1);
  }
  s = s.replace(/\$/g, '').toUpperCase();
  let m = /^([A-Z]{1,3})(\d+)(?::([A-Z]{1,3})(\d+))?$/.exec(s);
  if (m) {
    const a = { r: Number(m[2]) - 1, c: colIndex(m[1]) };
    const b = m[3] ? { r: Number(m[4]) - 1, c: colIndex(m[3]) } : a;
    if (a.r < 0 || b.r < 0) throw new RangeError2(`${text} is not a cell range.`);
    return { tab, r1: Math.min(a.r, b.r), c1: Math.min(a.c, b.c), r2: Math.max(a.r, b.r), c2: Math.max(a.c, b.c) };
  }
  m = /^([A-Z]{1,3}):([A-Z]{1,3})$/.exec(s);
  if (m) { const a = colIndex(m[1]), b = colIndex(m[2]); return { tab, r1: 0, c1: Math.min(a, b), r2: Infinity, c2: Math.max(a, b) }; }
  m = /^(\d+):(\d+)$/.exec(s);
  if (m) { const a = Number(m[1]) - 1, b = Number(m[2]) - 1; if (a < 0 || b < 0) throw new RangeError2(`${text} is not a range.`); return { tab, r1: Math.min(a, b), c1: 0, r2: Math.max(a, b), c2: Infinity }; }
  throw new RangeError2(`"${text}" is not a cell or range. Write it like B3, A1:C10, A:C or 'Tab name'!A1:B2.`);
}

export function rangeText({ r1, c1, r2, c2 }, tab = null) {
  const head = tab ? `${quoteTab(tab)}!` : '';
  if (r1 === r2 && c1 === c2) return `${head}${addr(r1, c1)}`;
  return `${head}${addr(r1, c1)}:${addr(r2, c2)}`;
}

// Moves a formula by dr rows and dc columns the way copy and paste does: relative references move,
// $absolute ones stay. A reference pushed off the sheet becomes #REF!, as in Excel and Google Sheets.
// Strings in quotes are left alone, and so are function names (LOG10( is not a cell).
export function shiftFormula(formula, dr, dc) {
  if (typeof formula !== 'string' || !formula.startsWith('=') || (!dr && !dc)) return formula;
  let out = '';
  let i = 0;
  const f = formula;
  const cellRe = /(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?::(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7}))?/y;
  const colsRe = /(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![A-Za-z0-9_(])/y;
  const rowsRe = /(\$?)(\d{1,7}):(\$?)(\d{1,7})(?![0-9.])/y;
  const isWordChar = (ch) => !!ch && /[A-Za-z0-9_.]/.test(ch);
  const moveRow = (abs, n) => { if (abs) return `$${n}`; const v = Number(n) + dr; return v < 1 ? null : String(v); };
  const moveCol = (abs, s) => { if (abs) return `$${s.toUpperCase()}`; const v = colIndex(s) + dc; return v < 0 ? null : colName(v); };
  while (i < f.length) {
    const ch = f[i];
    if (ch === '"') {
      let j = i + 1;
      while (j < f.length) { if (f[j] === '"') { if (f[j + 1] === '"') { j += 2; continue; } break; } j++; }
      out += f.slice(i, j + 1);
      i = j + 1;
      continue;
    }
    if (ch === "'") {
      // A quoted tab name: copy it through, then the reference after the ! is handled on the next turn.
      let j = i + 1;
      while (j < f.length) { if (f[j] === "'") { if (f[j + 1] === "'") { j += 2; continue; } break; } j++; }
      out += f.slice(i, j + 1);
      i = j + 1;
      continue;
    }
    const prev = f[i - 1];
    if (!isWordChar(prev) || prev === undefined) {
      cellRe.lastIndex = i;
      let m = cellRe.exec(f);
      if (m) {
        const end = cellRe.lastIndex;
        const next = f[end];
        if (!(next && /[A-Za-z0-9_(!]/.test(next))) {
          const a = [moveCol(m[1], m[2]), moveRow(m[3], m[4])];
          let txt = a.includes(null) ? null : `${a[0]}${a[1]}`;
          if (m[6] && txt) {
            const b = [moveCol(m[5], m[6]), moveRow(m[7], m[8])];
            txt = b.includes(null) ? null : `${txt}:${b[0]}${b[1]}`;
          }
          out += txt ?? '#REF!';
          i = end;
          continue;
        }
      }
      colsRe.lastIndex = i;
      m = colsRe.exec(f);
      if (m) {
        const a = moveCol(m[1], m[2]), b = moveCol(m[3], m[4]);
        out += a === null || b === null ? '#REF!' : `${a}:${b}`;
        i = colsRe.lastIndex;
        continue;
      }
      rowsRe.lastIndex = i;
      m = rowsRe.exec(f);
      if (m && !/[0-9.]/.test(prev ?? '')) {
        const a = moveRow(m[1], m[2]), b = moveRow(m[3], m[4]);
        out += a === null || b === null ? '#REF!' : `${a}:${b}`;
        i = rowsRe.lastIndex;
        continue;
      }
    }
    // Skip over a whole word (a function or a name) so its tail is never read as a cell.
    if (/[A-Za-z_]/.test(ch)) {
      let j = i;
      while (j < f.length && /[A-Za-z0-9_.]/.test(f[j])) j++;
      out += f.slice(i, j);
      i = j;
      continue;
    }
    out += ch;
    i++;
  }
  return out;
}
