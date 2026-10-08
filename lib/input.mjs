// What a typed value means, and how a value is shown. Shared by the server (agents, export) and the screen.
// Typing works like Google Sheets: "$1,200" is the number 1200 shown as currency, "12%" is 0.12 shown as a
// percent, "2026-10-07" is a date, "=SUM(A1:A3)" is a formula, and a leading ' keeps anything as text.
import SSF from 'ssf';

const EPOCH = Date.UTC(1899, 11, 30);
const DAY = 86400000;

export const dateSerial = (y, m, d, hh = 0, mm = 0, ss = 0) => (Date.UTC(y, m - 1, d, hh, mm, ss) - EPOCH) / DAY;
export const serialToDate = (n) => new Date(EPOCH + Math.round(n * DAY));

// Number formats offered in the toolbar (Excel format codes, so they survive an xlsx round trip).
export const FORMATS = [
  { id: 'auto', label: 'Automatic', nf: null },
  { id: 'number', label: 'Number', nf: '#,##0.00', example: '1,234.56' },
  { id: 'integer', label: 'Whole number', nf: '#,##0', example: '1,235' },
  { id: 'currency', label: 'Currency', nf: '$#,##0.00', example: '$1,234.56' },
  { id: 'currency0', label: 'Currency, rounded', nf: '$#,##0', example: '$1,235' },
  { id: 'percent', label: 'Percent', nf: '0.0%', example: '12.3%' },
  { id: 'date', label: 'Date', nf: 'yyyy-mm-dd', example: '2026-10-07' },
  { id: 'date_us', label: 'Date (US)', nf: 'm/d/yyyy', example: '10/7/2026' },
  { id: 'datetime', label: 'Date and time', nf: 'yyyy-mm-dd hh:mm', example: '2026-10-07 14:30' },
  { id: 'text', label: 'Plain text', nf: '@', example: 'as typed' },
];

const NUM = /^[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?$/i;
const THOUSANDS = /^[+-]?\d{1,3}(,\d{3})+(\.\d+)?$/;

// Text typed into a cell (or given by an agent) -> { raw, nf }. raw is what is stored: a number, a
// boolean, a formula string starting with =, text, or null for empty. nf is a number format to apply
// when the cell has none yet (typing "$5" makes a currency cell).
export function parseInput(input, { nf: current = null } = {}) {
  if (input === null || input === undefined) return { raw: null };
  if (typeof input === 'number') return Number.isFinite(input) ? { raw: input } : { raw: null };
  if (typeof input === 'boolean') return { raw: input };
  const s = String(input);
  if (s === '') return { raw: null };
  if (current === '@') return { raw: textRaw(s) };
  if (s.startsWith('=') && s.length > 1) return { raw: s };
  if (s.startsWith("'")) return { raw: s };
  const t = s.trim();
  if (/^(true|false)$/i.test(t)) return { raw: t.toLowerCase() === 'true' };
  if (NUM.test(t)) return { raw: Number(t) };
  if (THOUSANDS.test(t)) { const n = Number(t.replace(/,/g, '')); return { raw: n, nf: t.includes('.') ? '#,##0.00' : '#,##0' }; }
  let m = /^(-)?\$\s?(-)?([\d,]*\.?\d+)$/.exec(t);
  if (m && (THOUSANDS.test(m[3]) || NUM.test(m[3]))) {
    const n = Number(m[3].replace(/,/g, '')) * (m[1] || m[2] ? -1 : 1);
    return { raw: n, nf: m[3].includes('.') ? '$#,##0.00' : '$#,##0' };
  }
  m = /^([+-]?(\d+\.?\d*|\.\d+))\s?%$/.exec(t);
  if (m) { const dec = (m[1].split('.')[1] ?? '').length; return { raw: round(Number(m[1]) / 100, 12), nf: dec ? `0.${'0'.repeat(dec)}%` : '0%' }; }
  m = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(t);
  if (m && validDate(+m[1], +m[2], +m[3])) return { raw: dateSerial(+m[1], +m[2], +m[3], +(m[4] ?? 0), +(m[5] ?? 0), +(m[6] ?? 0)), nf: m[4] ? 'yyyy-mm-dd hh:mm' : 'yyyy-mm-dd' };
  m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(t);
  if (m && validDate(+m[3], +m[1], +m[2])) return { raw: dateSerial(+m[3], +m[1], +m[2]), nf: 'm/d/yyyy' };
  return { raw: s };
}

const validDate = (y, m, d) => m >= 1 && m <= 12 && d >= 1 && d <= 31 && y >= 1900 && y <= 9999 && new Date(Date.UTC(y, m - 1, d)).getUTCDate() === d;
const round = (n, p) => Number(n.toPrecision(p));

// Stored text that would otherwise read as a number, date, boolean or formula keeps a leading '.
export function textRaw(s) {
  s = String(s);
  if (s === '') return null;
  if (s.startsWith("'") || s.startsWith('=')) return `'${s}`;
  const p = parseInput(s);
  return typeof p.raw === 'string' && !p.raw.startsWith('=') ? s : `'${s}`;
}

// What a person typed, back as editable text (the formula bar and the cell editor).
export function editText(raw, nf = null) {
  if (raw === null || raw === undefined) return '';
  if (typeof raw === 'boolean') return raw ? 'TRUE' : 'FALSE';
  if (typeof raw === 'number') {
    if (nf && isDateFormat(nf)) return SSF.format(/h/.test(nf) ? 'yyyy-mm-dd hh:mm' : 'yyyy-mm-dd', raw);
    if (nf && nf.includes('%')) return `${round(raw * 100, 12)}%`;
    return String(raw);
  }
  return String(raw);
}

export const isDateFormat = (nf) => !!nf && /(^|[^"\\])[ymd]/i.test(nf.replace(/"[^"]*"/g, '')) && !/^[#0,.$%]+$/.test(nf);

// HyperFormula's guess at a formula result's type, when the cell has no format of its own.
const TYPE_FORMAT = { NUMBER_DATE: 'yyyy-mm-dd', NUMBER_DATETIME: 'yyyy-mm-dd hh:mm', NUMBER_TIME: 'hh:mm:ss', NUMBER_PERCENT: '0%', NUMBER_CURRENCY: '$#,##0.00' };

// A value (from the engine) -> the text shown in the cell.
export function display(value, nf = null, type = null) {
  if (value === null || value === undefined || value === '') return '';
  if (typeof value === 'object' && value.value !== undefined) return String(value.value); // an error such as #DIV/0!
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (typeof value === 'number') {
    const f = nf && nf !== '@' ? nf : TYPE_FORMAT[type] ?? 'General';
    try { return SSF.format(f, value); } catch { return String(value); }
  }
  const s = String(value);
  return s;
}

// The plain value an agent reads: numbers stay numbers, errors are their code (#DIV/0!), empty is null.
export function plainValue(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'object' && v.value !== undefined) return String(v.value);
  return v;
}
