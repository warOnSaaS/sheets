// The tool catalogue: everything a person can do in Sheets, as tools. Screens call them at
// /api/tools/<name>, agents call the same ones over MCP at /mcp, and tools.json is generated from this list
// (npm run tools:json). Names follow the suite contract: sheets.verb_noun here and in tools.json, and
// sheets_verb_noun on the wire (MCP and OpenAPI). Each has a plain description, a scope (read, write,
// delete, admin), confirm (none or human) and the events it emits.
import { z } from 'zod';
import { zodToJsonSchema } from 'zod-to-json-schema';
import { Y, BookError, title as bookTitle, toB64, fromB64, tabList, needTab, addTab, cellMap, fmtMap, tabMap, setRaw, initBook } from './book.mjs';
import { Engine } from './engine.mjs';
import * as ops from './ops.mjs';
import { importXlsx, exportXlsx, parseCsv, toCsv } from './xlsx.mjs';
import { fetchCrm, writeCrmTab, crmConnection, callCrm, CRM_KINDS } from './crm.mjs';
import { buildTemplate, TEMPLATES } from './templates.mjs';
import { personView } from './team.mjs';
import { FORMATS } from './input.mjs';
import { newId, nowIso } from './ids.mjs';
import { rowView } from './store.mjs';
import { zipSync, strToU8 } from 'fflate';

const SHEET = z.string().min(1).describe('The spreadsheet: its id (wb_...) or its exact title');
const TAB = z.string().optional().describe('A tab, by name or id (default: the first tab). A range like \'Deals\'!A1:C9 names its own tab.');
const RANGE = z.string().describe('A1 notation: B3, A1:D20, A:C (whole columns), 2:5 (whole rows), or \'Tab name\'!A1:B9');
const CELL = z.union([z.string(), z.number(), z.boolean(), z.null()]);
const Values = z.array(z.array(CELL));
const Fmt = z.object({ b: z.boolean(), i: z.boolean(), u: z.boolean(), s: z.boolean(), fc: z.string(), bg: z.string(), ha: z.enum(['left', 'center', 'right']), va: z.enum(['top', 'middle', 'bottom']), wrap: z.boolean(), nf: z.string(), fs: z.number() }).partial().nullable();
const Out = z.object({}).passthrough();
const FileRef = z.object({ id: z.string(), name: z.string(), type: z.string(), size: z.number(), url: z.string() });

const TOOLS = [];
const tool = (t) => TOOLS.push({ confirm: 'none', emits: [], ...t });

// ---------- helpers ----------

const find = (app, me, ref) => app.store.find(me.team_id, ref);

async function reading(app, me, ref, fn) {
  const row = await find(app, me, ref);
  const e = await app.store.open(row);
  return fn(e.doc, app.store.engine(e), row);
}

async function writing(app, me, ref, fn, opts) {
  const row = await find(app, me, ref);
  const { result, seq } = await app.store.change(me, row, fn, opts);
  return { sheet: row.id, ...(result ?? {}), ...(seq ? { seq } : {}) };
}

const sheetView = (app, row, doc, engine) => ({ ...rowView(row), url: app.urlFor(row.id), ...ops.bookSummary(doc, engine) });

// ---------- spreadsheets ----------

tool({
  name: 'sheets.list_sheets', title: 'List spreadsheets', scope: 'read',
  description: 'The team\'s spreadsheets, most recently changed first. q narrows by words in the title.',
  input: { q: z.string().optional().describe('Words in the title') },
  output: z.object({ sheets: z.array(Out) }),
  run: async ({ app, me }, a) => ({ sheets: (await app.store.list(me.team_id, a)).map((s) => ({ ...s, url: app.urlFor(s.id) })) }),
});

tool({
  name: 'sheets.create_sheet', title: 'New spreadsheet', scope: 'write', emits: ['sheets.workbook.created'],
  description: `Make a spreadsheet. Give tabs to name them, values to fill the first tab (rows of cells, typed as a person would: "$1,200", "=SUM(B2:B9)"), or template to start from an example: ${Object.entries(TEMPLATES).map(([k, t]) => `${k} (${t.about})`).join('; ')}.`,
  input: { title: z.string().max(200).optional(), tabs: z.array(z.string()).max(50).optional().describe('Tab names, left to right'), values: Values.optional().describe('Rows of cells for the first tab, from A1'), template: z.enum(Object.keys(TEMPLATES)).optional() },
  output: Out,
  run: async ({ app, me }, a) => {
    let row;
    if (a.template) {
      const doc = buildTemplate(a.template);
      row = await app.store.create(me, { title: a.title || TEMPLATES[a.template].title, state: Y.encodeStateAsUpdate(doc) });
    } else {
      row = await app.store.create(me, { title: a.title || 'Untitled spreadsheet', tabs: (a.tabs?.length ? a.tabs : ['Sheet1']).map((name) => ({ name })) });
    }
    if (a.values?.length) await app.store.change(me, row, (doc, engine) => ops.writeRange(doc, engine, { tab: null, range: 'A1', values: a.values }));
    return reading(app, me, row.id, (doc, engine, r) => sheetView(app, r, doc, engine));
  },
});

tool({
  name: 'sheets.get_sheet', title: 'Open a spreadsheet', scope: 'read',
  description: 'One spreadsheet: its tabs (with the range in use, frozen rows, filter and any CRM or pivot link), charts, open comments, sharing and address. Read cells with sheets.read_range.',
  input: { sheet: SHEET },
  output: Out,
  run: async ({ app, me }, a) => reading(app, me, a.sheet, async (doc, engine, row) => ({ ...sheetView(app, row, doc, engine), share: row.share_mode !== 'off' ? { mode: row.share_mode, url: app.shareUrl(row.share_token) } : null })),
});

tool({
  name: 'sheets.rename_sheet', title: 'Rename a spreadsheet', scope: 'write', emits: ['sheets.workbook.changed'],
  description: 'Change a spreadsheet\'s title.',
  input: { sheet: SHEET, title: z.string().min(1).max(200) },
  output: Out,
  run: async ({ app, me }, a) => writing(app, me, a.sheet, (doc) => { doc.getMap('meta').set('title', a.title.trim()); return { title: a.title.trim() }; }),
});

tool({
  name: 'sheets.copy_sheet', title: 'Make a copy', scope: 'write', emits: ['sheets.workbook.created'],
  description: 'Copy a spreadsheet with everything in it (tabs, formulas, formats, charts, comments). The copy is not shared.',
  input: { sheet: SHEET, title: z.string().max(200).optional() },
  output: Out,
  run: async ({ app, me }, a) => {
    const src = await find(app, me, a.sheet);
    const e = await app.store.open(src);
    const row = await app.store.create(me, { title: a.title || `Copy of ${bookTitle(e.doc)}`, state: Y.encodeStateAsUpdate(e.doc) });
    return reading(app, me, row.id, (doc, engine, r) => sheetView(app, r, doc, engine));
  },
});

tool({
  name: 'sheets.delete_sheet', title: 'Delete a spreadsheet', scope: 'delete', confirm: 'human', emits: ['sheets.workbook.deleted'],
  description: 'Delete a spreadsheet for everyone on the team. It disappears from the list and its share link stops working. An agent calling this asks a person first.',
  input: { sheet: SHEET },
  output: Out,
  run: async ({ app, me }, a) => { const row = await find(app, me, a.sheet); await app.store.remove(row); return { deleted: row.id, title: row.title }; },
});

// ---------- cells ----------

tool({
  name: 'sheets.read_range', title: 'Read cells', scope: 'read',
  description: 'Exact cell contents: values (numbers stay numbers, errors are their code like #DIV/0!) and, where a cell has one, its formula. Leave range out to read everything in use on the tab. text: true adds what each cell shows (formatted, like $1,200.00); formats: true adds each cell\'s format. version reads an earlier version (sheets.list_versions).',
  input: { sheet: SHEET, tab: TAB, range: RANGE.optional(), text: z.boolean().optional(), formats: z.boolean().optional(), version: z.string().optional().describe('A version id, to read the spreadsheet as it was then') },
  output: Out,
  run: async ({ app, me }, a) => {
    if (a.version) {
      const row = await find(app, me, a.sheet);
      const { doc, version } = await app.store.versionDoc(row, a.version);
      const engine = new Engine(doc, { live: false });
      try { return { sheet: row.id, version: { id: version.id, label: version.label, created_at: version.created_at }, ...ops.readRange(doc, engine, a) }; } finally { engine.destroy(); }
    }
    return reading(app, me, a.sheet, (doc, engine, row) => ({ sheet: row.id, ...ops.readRange(doc, engine, a) }));
  },
});

tool({
  name: 'sheets.write_range', title: 'Write cells', scope: 'write', emits: ['sheets.workbook.changed'],
  description: 'Write rows of cells starting at the top-left of range. Cells are read as a person typing them: "1200" and "$1,200" are numbers, "12%" is 0.12, "2026-10-07" is a date, "=SUM(B2:B9)" is a formula, "\'007" stays text. input: "raw" stores strings exactly as text instead. null clears a cell. Returns the values the cells came to, and any formula errors.',
  input: { sheet: SHEET, tab: TAB, range: RANGE.optional().describe('Where to start (default A1), or the whole range'), values: Values, input: z.enum(['user', 'raw']).optional(), formats: z.array(z.array(Fmt.optional())).optional().describe('Formats for the same cells: b, i, u, s, fc (text colour), bg (fill), ha (align), wrap, nf (number format code), fs (size). null clears.') },
  output: Out,
  run: async ({ app, me }, a) => writing(app, me, a.sheet, (doc, engine) => ops.writeRange(doc, engine, a)),
});

tool({
  name: 'sheets.set_formula', title: 'Set a formula', scope: 'write', emits: ['sheets.workbook.changed'],
  description: 'Put a formula in a cell, like =SUMIF(B2:B20,"Won",C2:C20). fill copies it across a range the way dragging the fill handle does: relative references move, $absolute ones stay (fill: "D2:D20" from cell D2). Checks the formula first and returns what it came to.',
  input: { sheet: SHEET, tab: TAB, cell: z.string().describe('The cell, like D2'), formula: z.string().describe('The formula, with or without the leading ='), fill: z.string().optional().describe('A range that includes the cell, to fill the formula into') },
  output: Out,
  run: async ({ app, me }, a) => writing(app, me, a.sheet, (doc, engine) => ops.setFormula(doc, engine, a)),
});

tool({
  name: 'sheets.append_rows', title: 'Add rows', scope: 'write', emits: ['sheets.workbook.changed'],
  description: 'Add rows under the last row in use (or under the table that starts at range, like A1). Cells are read as typed, as in sheets.write_range.',
  input: { sheet: SHEET, tab: TAB, rows: Values, range: z.string().optional().describe('The table\'s first cell or range, when the tab holds more than one table'), input: z.enum(['user', 'raw']).optional() },
  output: Out,
  run: async ({ app, me }, a) => writing(app, me, a.sheet, (doc, engine) => ops.appendRows(doc, engine, a)),
});

tool({
  name: 'sheets.clear_range', title: 'Clear cells', scope: 'write', emits: ['sheets.workbook.changed'],
  description: 'Empty a range: what: values (default), formats or all.',
  input: { sheet: SHEET, tab: TAB, range: RANGE, what: z.enum(['values', 'formats', 'all']).optional() },
  output: Out,
  run: async ({ app, me }, a) => writing(app, me, a.sheet, (doc, engine) => ops.clearRange(doc, engine, a)),
});

tool({
  name: 'sheets.find', title: 'Find', scope: 'read',
  description: 'Find cells by what they show (or by their formula, with formulas: true), across every tab or one. Returns each match\'s tab, cell, value and formula.',
  input: { sheet: SHEET, query: z.string(), tab: TAB, match_case: z.boolean().optional(), whole_cell: z.boolean().optional(), formulas: z.boolean().optional(), regex: z.boolean().optional(), limit: z.number().int().min(1).max(500).optional() },
  output: Out,
  run: async ({ app, me }, a) => reading(app, me, a.sheet, (doc, engine, row) => ({ sheet: row.id, ...ops.find(doc, engine, a) })),
});

// ---------- sort, filter, rows and columns ----------

const SortKey = z.object({ column: z.union([z.string(), z.number()]).describe('A letter (C), a header from the first row ("Amount"), or a number counting from 1 within the range'), order: z.enum(['asc', 'desc']).optional() });

tool({
  name: 'sheets.sort', title: 'Sort', scope: 'write', emits: ['sheets.workbook.changed'],
  description: 'Sort the rows of a range (default: everything in use) by one or more columns. A header row is kept on top (found automatically; header: true or false to say). Formulas in moved rows keep pointing at their own row, as in Excel. Empty cells go last.',
  input: { sheet: SHEET, tab: TAB, range: RANGE.optional(), by: z.union([SortKey, z.array(SortKey).min(1), z.string()]), header: z.boolean().optional() },
  output: Out,
  run: async ({ app, me }, a) => writing(app, me, a.sheet, (doc, engine) => ops.sortRange(doc, engine, { ...a, by: typeof a.by === 'string' ? { column: a.by } : a.by })),
});

const Condition = z.object({ op: z.enum([...ops.FILTER_OPS, 'all']).describe('eq, neq, gt, gte, lt, lte, contains, not_contains, starts_with, empty, not_empty, in (with values), or all (no condition)'), value: CELL.optional(), values: z.array(CELL).optional() });

tool({
  name: 'sheets.filter', title: 'Filter', scope: 'write', emits: ['sheets.workbook.changed'],
  description: 'Turn on the tab\'s filter (everyone sees it, as in Google Sheets) and set a condition on one column; rows that do not match are hidden. Call again for more columns. clear: true removes the filter. Returns the rows that still show. To look without changing what others see, use sheets.read_range or sheets.find.',
  input: { sheet: SHEET, tab: TAB, range: RANGE.optional().describe('The table, header row first (default: everything in use)'), column: z.union([z.string(), z.number()]).optional().describe('A letter or a header'), condition: Condition.optional(), clear: z.boolean().optional() },
  output: Out,
  run: async ({ app, me }, a) => writing(app, me, a.sheet, (doc, engine) => ops.setFilter(doc, engine, a)),
});

for (const [name, axis, del, title] of [['insert_rows', 'row', false, 'Insert rows'], ['delete_rows', 'row', true, 'Delete rows'], ['insert_columns', 'col', false, 'Insert columns'], ['delete_columns', 'col', true, 'Delete columns']]) {
  tool({
    name: `sheets.${name}`, title, scope: 'write', emits: ['sheets.workbook.changed'],
    description: `${title} ${del ? 'starting at' : 'before'} ${axis === 'row' ? 'row' : 'column'} at (counting from 1${axis === 'col' ? ', or a letter' : ''}). Every formula that points past them moves${del ? '; one that pointed into them shows #REF!' : ''}.`,
    input: { sheet: SHEET, tab: TAB, at: axis === 'row' ? z.number().int().min(1) : z.union([z.number().int().min(1), z.string()]), count: z.number().int().min(1).max(5000).optional() },
    output: Out,
    run: async ({ app, me }, a) => writing(app, me, a.sheet, (doc, engine) => ops.insertDelete(doc, engine, { tab: a.tab, axis, del, count: a.count ?? 1, at: typeof a.at === 'string' ? (/^\d+$/.test(a.at) ? Number(a.at) : colIndexOf(a.at) + 1) : a.at })),
  });
}
const colIndexOf = (s) => { let n = 0; for (const ch of String(s).toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64); return n - 1; };

// ---------- formatting ----------

tool({
  name: 'sheets.format_range', title: 'Format cells', scope: 'write', emits: ['sheets.workbook.changed'],
  description: `Format a range: bold, italic, underline, strike, color and fill (#rrggbb), align (left, center, right), valign, wrap, font_size, number_format (${FORMATS.map((f) => f.id).join(', ')}, or an Excel format code like "$#,##0.00" or "0.0%"). false or null turns one off; clear: true removes all formatting.`,
  input: { sheet: SHEET, tab: TAB, range: RANGE, bold: z.boolean().optional(), italic: z.boolean().optional(), underline: z.boolean().optional(), strike: z.boolean().optional(), color: z.string().nullable().optional(), fill: z.string().nullable().optional(), align: z.enum(['left', 'center', 'right']).nullable().optional(), valign: z.enum(['top', 'middle', 'bottom']).nullable().optional(), wrap: z.boolean().optional(), font_size: z.number().min(6).max(48).nullable().optional(), number_format: z.string().nullable().optional(), clear: z.boolean().optional() },
  output: Out,
  run: async ({ app, me }, a) => writing(app, me, a.sheet, (doc, engine) => ops.formatRange(doc, engine, a)),
});

tool({
  name: 'sheets.set_column_width', title: 'Column width', scope: 'write', emits: ['sheets.workbook.changed'],
  description: 'Set the width of columns in pixels (24 to 800), like columns: "B" or "B:D". width: null goes back to the default.',
  input: { sheet: SHEET, tab: TAB, columns: z.string(), width: z.number().nullable() },
  output: Out,
  run: async ({ app, me }, a) => writing(app, me, a.sheet, (doc) => ops.setColumnWidth(doc, a)),
});

tool({
  name: 'sheets.freeze', title: 'Freeze rows and columns', scope: 'write', emits: ['sheets.workbook.changed'],
  description: 'Keep the top rows and left columns in view while scrolling. 0 unfreezes.',
  input: { sheet: SHEET, tab: TAB, rows: z.number().int().min(0).max(50).optional(), columns: z.number().int().min(0).max(26).optional() },
  output: Out,
  run: async ({ app, me }, a) => writing(app, me, a.sheet, (doc) => ops.setFreeze(doc, a)),
});

// ---------- tabs ----------

tool({
  name: 'sheets.add_sheet_tab', title: 'Add a tab', scope: 'write', emits: ['sheets.workbook.changed'],
  description: 'Add a tab to a spreadsheet, at the end or at index (0 is first).',
  input: { sheet: SHEET, name: z.string().max(100).optional(), index: z.number().int().min(0).optional() },
  output: Out,
  run: async ({ app, me }, a) => writing(app, me, a.sheet, (doc) => ({ tab: ops.addSheetTab(doc, a), tabs: tabList(doc).map((t) => t.name) })),
});

tool({
  name: 'sheets.rename_tab', title: 'Rename a tab', scope: 'write', emits: ['sheets.workbook.changed'],
  description: 'Rename a tab. Formulas on other tabs that point at it are rewritten to the new name.',
  input: { sheet: SHEET, tab: z.string(), name: z.string().min(1).max(100) },
  output: Out,
  run: async ({ app, me }, a) => writing(app, me, a.sheet, (doc, engine) => ({ tab: ops.renameTab(doc, engine, a) })),
});

tool({
  name: 'sheets.move_tab', title: 'Move a tab', scope: 'write', emits: ['sheets.workbook.changed'],
  description: 'Move a tab to a new place (index 0 is first).',
  input: { sheet: SHEET, tab: z.string(), index: z.number().int().min(0) },
  output: Out,
  run: async ({ app, me }, a) => writing(app, me, a.sheet, (doc) => ops.moveTab(doc, a)),
});

tool({
  name: 'sheets.delete_tab', title: 'Delete a tab', scope: 'delete', emits: ['sheets.workbook.changed'],
  description: 'Delete a tab and everything on it. Formulas elsewhere that pointed at it show #REF!. Version history keeps the old tab (sheets.restore_version).',
  input: { sheet: SHEET, tab: z.string() },
  output: Out,
  run: async ({ app, me }, a) => writing(app, me, a.sheet, (doc, engine) => ops.deleteTab(doc, engine, a)),
});

// ---------- charts and pivots ----------

tool({
  name: 'sheets.create_chart', title: 'Add a chart', scope: 'write', emits: ['sheets.workbook.changed'],
  description: `A chart from a range, drawn on the tab and kept up to date. type: ${ops.CHART_TYPES.join(', ')}. By columns (default) the first column is the labels and each other column a series named by its header; series_by: "rows" turns that around. series picks which columns (letters) or rows (numbers) to plot. at is the cell where its top-left corner sits. Returns the chart with the data it shows.`,
  input: { sheet: SHEET, tab: TAB, range: RANGE, type: z.enum(ops.CHART_TYPES).optional(), title: z.string().max(120).optional(), at: z.string().optional(), stacked: z.boolean().optional(), series_by: z.enum(['columns', 'rows']).optional(), series: z.array(z.union([z.string(), z.number()])).optional() },
  output: Out,
  run: async ({ app, me }, a) => writing(app, me, a.sheet, (doc, engine) => ops.createChart(doc, engine, a)),
});

tool({
  name: 'sheets.update_chart', title: 'Change a chart', scope: 'write', emits: ['sheets.workbook.changed'],
  description: 'Change a chart\'s type, range, title, series, place (at) or size (width and height in pixels).',
  input: { sheet: SHEET, chart: z.string(), type: z.enum(ops.CHART_TYPES).optional(), range: z.string().optional(), title: z.string().max(120).nullable().optional(), at: z.string().optional(), stacked: z.boolean().optional(), series_by: z.enum(['columns', 'rows']).optional(), series: z.array(z.union([z.string(), z.number()])).nullable().optional(), width: z.number().optional(), height: z.number().optional() },
  output: Out,
  run: async ({ app, me }, a) => writing(app, me, a.sheet, (doc, engine) => ops.updateChart(doc, engine, a)),
});

tool({
  name: 'sheets.delete_chart', title: 'Delete a chart', scope: 'write', emits: ['sheets.workbook.changed'],
  description: 'Remove a chart. Its data stays.',
  input: { sheet: SHEET, chart: z.string() },
  output: Out,
  run: async ({ app, me }, a) => writing(app, me, a.sheet, (doc) => ops.deleteChart(doc, a)),
});

tool({
  name: 'sheets.list_charts', title: 'List charts', scope: 'read',
  description: 'The charts in a spreadsheet (or on one tab), each with the labels and numbers it shows right now.',
  input: { sheet: SHEET, tab: TAB },
  output: Out,
  run: async ({ app, me }, a) => reading(app, me, a.sheet, (doc, engine, row) => ({ sheet: row.id, charts: ops.listCharts(doc, engine, a) })),
});

tool({
  name: 'sheets.create_pivot', title: 'Pivot table', scope: 'write', emits: ['sheets.workbook.changed'],
  description: 'Summarise a table on a new tab: one row per value of rows (a header or letter), optional columns to split across, and values to work out (summarize: sum, count, average, min, max of a field). It is written as live formulas (SUMIFS, COUNTIFS, ...), so it updates as the table changes and opens in Excel. sheets.refresh_link picks up new categories.',
  input: { sheet: SHEET, tab: TAB.describe('The tab with the table'), range: RANGE.optional().describe('The table, header row first (default: everything in use)'), rows: z.union([z.string(), z.number()]), columns: z.union([z.string(), z.number()]).optional(), values: z.array(z.object({ field: z.union([z.string(), z.number()]).optional(), summarize: z.enum(['sum', 'count', 'average', 'min', 'max']) })).optional(), target: z.string().max(100).optional().describe('Name for the new tab') },
  output: Out,
  run: async ({ app, me }, a) => writing(app, me, a.sheet, (doc, engine) => ops.createPivot(doc, engine, a)),
});

// ---------- comments ----------

tool({
  name: 'sheets.list_comments', title: 'Comments', scope: 'read',
  description: 'Open comments (include_resolved: true for all), with the cell each is on and its replies.',
  input: { sheet: SHEET, tab: TAB, include_resolved: z.boolean().optional() },
  output: Out,
  run: async ({ app, me }, a) => reading(app, me, a.sheet, (doc, engine, row) => ({ sheet: row.id, comments: ops.listComments(doc, a) })),
});

tool({
  name: 'sheets.add_comment', title: 'Comment on a cell', scope: 'write', emits: ['sheets.workbook.changed', 'sheets.comment.added'],
  description: 'Add a comment to a cell.',
  input: { sheet: SHEET, tab: TAB, cell: z.string(), body: z.string().min(1).max(5000) },
  output: Out,
  run: async ({ app, me }, a) => writing(app, me, a.sheet, (doc) => ({ comment: ops.addComment(doc, me, a) })),
});

tool({
  name: 'sheets.reply_comment', title: 'Reply to a comment', scope: 'write', emits: ['sheets.workbook.changed', 'sheets.comment.added'],
  description: 'Reply in a comment\'s thread.',
  input: { sheet: SHEET, comment: z.string(), body: z.string().min(1).max(5000) },
  output: Out,
  run: async ({ app, me }, a) => writing(app, me, a.sheet, (doc) => ({ comment: ops.replyComment(doc, me, a) })),
});

tool({
  name: 'sheets.resolve_comment', title: 'Resolve a comment', scope: 'write', emits: ['sheets.workbook.changed'],
  description: 'Mark a comment done (resolved: false opens it again).',
  input: { sheet: SHEET, comment: z.string(), resolved: z.boolean().optional() },
  output: Out,
  run: async ({ app, me }, a) => writing(app, me, a.sheet, (doc) => ({ comment: ops.resolveComment(doc, a) })),
});

tool({
  name: 'sheets.delete_comment', title: 'Delete a comment', scope: 'write', emits: ['sheets.workbook.changed'],
  description: 'Delete a comment and its replies (your own, or anyone\'s if you are an admin).',
  input: { sheet: SHEET, comment: z.string() },
  output: Out,
  run: async ({ app, me }, a) => writing(app, me, a.sheet, (doc) => ops.deleteComment(doc, me, a)),
});

// ---------- import and export ----------

tool({
  name: 'sheets.import_xlsx', title: 'Import an Excel file', scope: 'write', emits: ['sheets.workbook.created'],
  description: 'Bring in an .xlsx file (as base64) as a new spreadsheet: values, formulas, formats, frozen panes, widths, notes and named ranges. Returns an honest report: what carried over, what did not (merged cells, conditional formatting, charts, macros, ...), and every formula whose result differs from what the file saved.',
  input: { content_base64: z.string().describe('The .xlsx file, base64'), title: z.string().max(200).optional(), name: z.string().optional().describe('The file name, used as the title when there is none') },
  output: Out,
  run: async ({ app, me }, a) => {
    const buf = Buffer.from(a.content_base64, 'base64');
    if (buf.length > 25 * 1024 * 1024) throw new BookError('Files up to 25 MB for now.');
    const { doc, report } = await importXlsx(buf, { title: a.title || (a.name ? a.name.replace(/\.xlsx$/i, '') : null) });
    const row = await app.store.create(me, { title: bookTitle(doc), state: Y.encodeStateAsUpdate(doc) });
    return { sheet: row.id, title: row.title, url: app.urlFor(row.id), report };
  },
});

tool({
  name: 'sheets.import_csv', title: 'Import CSV', scope: 'write', emits: ['sheets.workbook.changed'],
  description: 'Bring in CSV text (or base64). With sheet, it goes on a new tab of that spreadsheet (or replaces tab); without, it becomes a new spreadsheet. Numbers, money, percents and dates are read as such. The delimiter (comma, semicolon or tab) is found automatically.',
  input: { content: z.string().optional().describe('The CSV text'), content_base64: z.string().optional(), sheet: SHEET.optional(), tab: z.string().optional().describe('Replace this tab\'s contents'), name: z.string().max(100).optional().describe('Name for the new tab or spreadsheet'), delimiter: z.enum([',', ';', '\t']).optional() },
  output: Out,
  run: async ({ app, me }, a) => {
    const text = a.content ?? (a.content_base64 ? Buffer.from(a.content_base64, 'base64').toString('utf8') : null);
    if (text === null) throw new BookError('Give content (the CSV text) or content_base64.');
    const { rows, delimiter } = parseCsv(text, a.delimiter);
    if (rows.length * Math.max(1, ...rows.map((r) => r.length)) > 500000) throw new BookError('CSV files up to 500,000 cells for now.');
    const name = (a.name || 'Imported').replace(/\.csv$/i, '').slice(0, 100);
    const report = { rows: rows.length, columns: Math.max(0, ...rows.map((r) => r.length)), delimiter: delimiter === '\t' ? 'tab' : delimiter };
    if (!a.sheet) {
      const row = await app.store.create(me, { title: name, tabs: [{ name: 'Sheet1' }] });
      const out = await writing(app, me, row.id, (doc, engine) => { const w = ops.writeRange(doc, engine, { tab: null, range: 'A1', values: rows }); return { tab: w.tab, range: w.range }; });
      return { ...out, title: name, url: app.urlFor(row.id), report };
    }
    return writing(app, me, a.sheet, (doc, engine) => {
      let tab;
      if (a.tab) { tab = needTab(doc, a.tab); doc.transact(() => { cellMap(doc, tab.id).clear(); fmtMap(doc, tab.id).clear(); }); }
      else tab = ops.addSheetTab(doc, { name: tabList(doc).some((t) => t.name.toLowerCase() === name.toLowerCase()) ? undefined : name });
      const w = ops.writeRange(doc, engine, { tab: tab.id, range: 'A1', values: rows });
      return { tab: w.tab, range: w.range, report };
    });
  },
});

tool({
  name: 'sheets.export', title: 'Download', scope: 'read',
  description: 'Download a spreadsheet as .xlsx (every tab, with formulas and their results, formats, frozen panes and comments as notes) or one tab as .csv (what the cells show; values: "raw" for plain numbers). Returns a download link; content: true also returns the file as base64. Says what an .xlsx cannot carry (charts, for now).',
  input: { sheet: SHEET, format: z.enum(['xlsx', 'csv']).optional(), tab: TAB, values: z.enum(['formatted', 'raw']).optional(), content: z.boolean().optional().describe('Also return the file as base64') },
  output: Out,
  run: async ({ app, me }, a) => reading(app, me, a.sheet, async (doc, engine, row) => {
    const format = a.format ?? 'xlsx';
    const base = bookTitle(doc).replace(/[^\w .-]+/g, '').trim() || 'spreadsheet';
    let data, name, type, notCarried = [];
    if (format === 'csv') {
      const t = needTab(doc, a.tab);
      data = Buffer.from(toCsv(doc, engine, t, a), 'utf8');
      name = `${base} - ${t.name}.csv`;
      type = 'text/csv';
    } else {
      const x = await exportXlsx(doc, engine);
      data = x.buffer; notCarried = x.not_carried;
      name = `${base}.xlsx`;
      type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    }
    const file = await app.files.put(me, { name, type, data });
    return { sheet: row.id, file, not_carried: notCarried, ...(a.content ? { content_base64: data.toString('base64') } : {}) };
  }),
});

tool({
  name: 'sheets.export_data', title: 'Export everything', scope: 'admin',
  description: 'Every spreadsheet on the team in one zip: each as .xlsx, plus sheets.json with every tab, cell, format, chart and comment exactly as stored. Your data is never stuck here.',
  input: {},
  output: z.object({ file: FileRef, counts: Out }),
  run: async ({ app, me }) => {
    const list = await app.store.list(me.team_id);
    const files = {};
    const all = [];
    for (const s of list) {
      const row = await app.store.row(me.team_id, s.id);
      const e = await app.store.open(row);
      const engine = app.store.engine(e);
      const x = await exportXlsx(e.doc, engine);
      const safe = `${s.title.replace(/[^\w .-]+/g, '').trim() || 'spreadsheet'} (${s.id}).xlsx`;
      files[`xlsx/${safe}`] = new Uint8Array(x.buffer);
      all.push({ ...s, document: e.doc.toJSON() });
    }
    files['sheets.json'] = strToU8(JSON.stringify({ exported_at: nowIso(), team: me.team_id, sheets: all }, null, 1));
    const zip = Buffer.from(zipSync(files));
    const file = await app.files.put(me, { name: `sheets-export-${nowIso().slice(0, 10)}.zip`, type: 'application/zip', data: zip });
    return { file, counts: { sheets: list.length } };
  },
});

// ---------- the CRM ----------

tool({
  name: 'sheets.import_from_crm', title: 'Pull from the CRM', scope: 'write', emits: ['sheets.workbook.changed'],
  description: `Pull CRM records onto a tab, kept linked so sheets.refresh_link brings it up to date. kind: ${Object.keys(CRM_KINDS).join(', ')} (with the usual fields), or soql for any read-only query, like SELECT Name, StageName, Amount FROM Opportunity WHERE StageName != 'Lost'. Uses the CRM's own tools (crm.query), as you. Formulas beside the table keep working after a refresh.`,
  input: { sheet: SHEET, kind: z.enum(Object.keys(CRM_KINDS)).optional(), fields: z.array(z.string()).optional().describe('CRM field names, like Name, Amount, CloseDate, Account.Name'), where: z.string().optional().describe('A SOQL condition, like Amount > 5000'), soql: z.string().optional(), tab: z.string().optional().describe('Replace this tab instead of adding one'), name: z.string().max(100).optional().describe('Name for the new tab') },
  output: Out,
  run: async ({ app, me, call }, a) => {
    if (!a.kind && !a.soql) throw new BookError('Give kind (deals, contacts, ...) or soql.');
    const source = a.soql ? { soql: a.soql } : { kind: a.kind, fields: a.fields ?? null, where: a.where ?? null };
    await find(app, me, a.sheet);
    const got = await fetchCrm(app, me, call, source);
    return writing(app, me, a.sheet, (doc) => writeCrmTab(doc, { tab: a.tab ?? null, name: a.name ?? null, source, ...got }));
  },
});

tool({
  name: 'sheets.refresh_link', title: 'Refresh', scope: 'write', emits: ['sheets.workbook.changed', 'sheets.link.refreshed'],
  description: 'Bring a linked tab up to date: a CRM tab pulls its records again; a pivot tab is rebuilt from its table (new categories appear). Without tab, every linked tab in the spreadsheet.',
  input: { sheet: SHEET, tab: z.string().optional() },
  output: Out,
  run: async ({ app, me, call }, a) => {
    const tabs = await reading(app, me, a.sheet, (doc) => {
      const ts = a.tab ? [needTab(doc, a.tab)] : tabList(doc).filter((t) => t.link);
      if (a.tab && !ts[0].link) throw new BookError(`${ts[0].name} is not linked to anything. CRM tabs come from sheets.import_from_crm, pivot tabs from sheets.create_pivot.`);
      return ts;
    });
    const pulled = new Map();
    for (const t of tabs) if (t.link.kind === 'crm') { const source = t.link.source ?? { soql: t.link.input?.soql }; pulled.set(t.id, { source, ...(await fetchCrm(app, me, call, source)) }); }
    return writing(app, me, a.sheet, (doc, engine) => {
      const done = [];
      for (const t of tabs) {
        if (t.link.kind === 'crm') done.push({ ...writeCrmTab(doc, { tab: t.id, ...pulled.get(t.id) }), kind: 'crm' });
        else if (t.link.kind === 'pivot') { const r = ops.createPivot(doc, engine, { def: t.link.def }); done.push({ tab: r.tab, rows: r.rows, kind: 'pivot' }); }
      }
      return { refreshed: done };
    });
  },
});

tool({
  name: 'sheets.connect_crm', title: 'Connect the CRM', scope: 'admin',
  description: 'Standalone servers: where the wOS CRM is (its address) and an access token from it (crm.create_access_token, read only is enough). disconnect: true forgets it. Inside the wOS suite nothing is needed: Sheets uses the CRM app directly.',
  input: { url: z.string().url().optional(), token: z.string().optional(), disconnect: z.boolean().optional() },
  output: Out,
  run: async ({ app, me, call }, a) => {
    if (a.disconnect) { await app.team.setSetting(me.team_id, 'crm', null); return { connected: false }; }
    if (!a.url) throw new BookError('Give the CRM\'s address (url).');
    if (!/^https:\/\//.test(a.url) && !/^http:\/\/(localhost|127\.0\.0\.1)(:|\/|$)/.test(a.url)) throw new BookError('The CRM address must start with https://.');
    await app.team.setSetting(me.team_id, 'crm', { url: a.url.replace(/\/$/, ''), token: a.token ? app.seal(a.token) : null, at: nowIso() });
    try {
      const d = await callCrm(app, me, null, 'crm.query', { soql: 'SELECT Name FROM Account LIMIT 1' });
      return { connected: true, url: a.url, check: d?.totalSize !== undefined ? `It answered: ${d.totalSize} organizations.` : 'It answered.' };
    } catch (e) {
      await app.team.setSetting(me.team_id, 'crm', null);
      throw new BookError(`Saved nothing: ${e.message}`, e.status ?? 400);
    }
  },
});

// ---------- sharing and versions ----------

tool({
  name: 'sheets.share', title: 'Share a link', scope: 'write', confirm: 'human', emits: ['sheets.workbook.shared'],
  description: 'link: "view" makes a read-only link anyone can open without an account (it shows values and charts, never lets them edit). link: "off" stops it. Everyone on the team can already open and edit every spreadsheet. An agent turning a link on asks a person first.',
  input: { sheet: SHEET, link: z.enum(['view', 'off']) },
  output: Out,
  run: async ({ app, me }, a) => {
    const row = await find(app, me, a.sheet);
    const s = await app.store.setShare(row, a.link);
    app.bus?.publish({ team: me.team_id, type: 'sheets.workbook.shared', data: { workbook: row.id, mode: s.mode } });
    return { sheet: row.id, link: s.mode, url: s.token ? app.shareUrl(s.token) : null };
  },
});

tool({
  name: 'sheets.list_versions', title: 'Version history', scope: 'read',
  description: 'Earlier versions of a spreadsheet, newest first: one is kept automatically every ten minutes of editing, plus every named version. Read one with sheets.read_range version, or bring it back with sheets.restore_version.',
  input: { sheet: SHEET },
  output: Out,
  run: async ({ app, me }, a) => {
    const row = await find(app, me, a.sheet);
    const people = new Map((await app.team.people(me.team_id)).map((p) => [p.id, p.name]));
    return { sheet: row.id, versions: (await app.store.versions(row)).map((v) => ({ ...v, created_by_name: people.get(v.created_by) ?? null })) };
  },
});

tool({
  name: 'sheets.save_version', title: 'Name this version', scope: 'write', emits: ['sheets.version.saved'],
  description: 'Save the spreadsheet as it is now, with a name, so it is easy to find in the history.',
  input: { sheet: SHEET, label: z.string().min(1).max(120) },
  output: Out,
  run: async ({ app, me }, a) => {
    const row = await find(app, me, a.sheet);
    const e = await app.store.open(row);
    return { sheet: row.id, version: await app.store.saveVersion(me, row, e.doc, { label: a.label }) };
  },
});

tool({
  name: 'sheets.restore_version', title: 'Restore a version', scope: 'write', emits: ['sheets.workbook.changed'],
  description: 'Bring back an earlier version for everyone. Nothing is lost: the version before the restore is saved first, so it can be restored too.',
  input: { sheet: SHEET, version: z.string() },
  output: Out,
  run: async ({ app, me }, a) => {
    const row = await find(app, me, a.sheet);
    const { doc: old, version } = await app.store.versionDoc(row, a.version);
    const cur = await app.store.open(row);
    await app.store.saveVersion(me, row, cur.doc, { label: `Before restoring ${version.label ?? version.created_at.slice(0, 16).replace('T', ' ')}` });
    return writing(app, me, row.id, (doc, engine) => { restoreInto(doc, old); engine.build(); return { restored: version.id, title: bookTitle(doc), tabs: tabList(doc).map((t) => t.name) }; });
  },
});

// Makes doc hold what old holds, as one change everyone receives.
function restoreInto(doc, old) {
  doc.transact(() => {
    doc.getMap('meta').set('title', bookTitle(old));
    const want = tabList(old);
    const wantIds = new Set(want.map((t) => t.id));
    for (const t of tabList(doc)) if (!wantIds.has(t.id)) { doc.getMap('tabs').delete(t.id); doc.getMap('cells').delete(t.id); doc.getMap('fmt').delete(t.id); }
    const order = doc.getArray('order');
    order.delete(0, order.length);
    order.insert(0, want.map((t) => t.id));
    for (const t of want) {
      let m = doc.getMap('tabs').get(t.id);
      if (!m) { m = new Y.Map(); doc.getMap('tabs').set(t.id, m); doc.getMap('cells').set(t.id, new Y.Map()); doc.getMap('fmt').set(t.id, new Y.Map()); }
      const src = old.getMap('tabs').get(t.id);
      for (const k of [...m.keys()]) if (!src.has(k)) m.delete(k);
      for (const [k, v] of src.entries()) if (m.get(k) !== v) m.set(k, v);
      for (const which of ['cells', 'fmt']) {
        const d = doc.getMap(which).get(t.id), s = old.getMap(which).get(t.id);
        for (const k of [...d.keys()]) if (!s.has(k)) d.delete(k);
        for (const [k, v] of s.entries()) if (JSON.stringify(d.get(k)) !== JSON.stringify(v)) d.set(k, v);
      }
    }
    for (const which of ['comments', 'charts', 'names']) {
      const d = doc.getMap(which), s = old.getMap(which);
      for (const k of [...d.keys()]) if (!s.has(k)) d.delete(k);
      for (const [k, v] of s.entries()) if (JSON.stringify(d.get(k)) !== JSON.stringify(v)) d.set(k, v);
    }
  });
}

// ---------- live editing (screens) ----------

tool({
  name: 'sheets.sync', title: 'Catch up', scope: 'read',
  description: 'For screens: the changes to a spreadsheet that a copy at state_vector (base64 Yjs state vector; empty for everything) has not seen, as one base64 Yjs update. Agents use sheets.read_range instead.',
  input: { sheet: SHEET, state_vector: z.string().optional() },
  output: Out,
  run: async ({ app, me }, a) => {
    const row = await find(app, me, a.sheet);
    const e = await app.store.open(row);
    const update = a.state_vector ? Y.encodeStateAsUpdate(e.doc, fromB64(a.state_vector)) : Y.encodeStateAsUpdate(e.doc);
    return { sheet: row.id, title: bookTitle(e.doc), update: toB64(update), seq: e.seq, share: row.share_mode !== 'off' ? { mode: row.share_mode, url: app.shareUrl(row.share_token) } : null, updated_at: row.updated_at };
  },
});

tool({
  name: 'sheets.set_presence', title: 'Show where I am', scope: 'read', emits: ['sheets.presence.changed'],
  description: 'For screens: tell others with the spreadsheet open which cell you are on. Nothing is stored.',
  input: { sheet: SHEET, tab: z.string().optional(), cell: z.string().optional(), left: z.boolean().optional() },
  output: z.object({ ok: z.boolean() }),
  run: async ({ app, me }, a) => {
    const row = await find(app, me, a.sheet);
    app.bus?.publish({ team: me.team_id, type: 'sheets.presence.changed', ephemeral: true, data: { workbook: row.id, person: { id: me.id, name: me.name }, tab: a.tab ?? null, cell: a.cell ?? null, left: !!a.left } });
    return { ok: true };
  },
});

// ---------- people and settings ----------

tool({
  name: 'sheets.list_people', title: 'People', scope: 'read',
  description: 'Everyone on the team. me is you.',
  input: {},
  output: Out,
  run: async ({ app, me }) => ({ me: personView(me), people: await app.team.people(me.team_id) }),
});

tool({
  name: 'sheets.add_person', title: 'Add someone', scope: 'admin', emits: ['sheets.person.added'],
  description: 'Let someone sign in and edit the team\'s spreadsheets: by email (for the email link or the warOnSaaS account), by GitHub username, or both.',
  input: { name: z.string().min(1).max(100), email: z.string().email().optional(), github: z.string().optional(), role: z.enum(['admin', 'member']).optional() },
  output: Out,
  run: async ({ app, me }, a) => {
    if (!a.email && !a.github) throw new BookError('Give an email or a GitHub username, so they can sign in.');
    return personView(await app.team.add(me.team_id, { ...a, role: a.role ?? 'member' }));
  },
});

tool({
  name: 'sheets.remove_person', title: 'Remove someone', scope: 'admin', confirm: 'human', emits: ['sheets.person.removed'],
  description: 'Take someone off the team: they are signed out everywhere. What they wrote stays. An agent calling this asks a person first.',
  input: { person: z.string().describe('Name, email or GitHub username') },
  output: Out,
  run: async ({ app, me }, a) => app.team.remove(me.team_id, a.person),
});

tool({
  name: 'sheets.get_settings', title: 'Settings', scope: 'read',
  description: 'You, your preferences, the team, where data is stored, how sign-in works, the CRM connection and the address to connect an AI.',
  input: {},
  output: Out,
  run: async ({ app, me, call }) => {
    const crm = call?.callTool ? { connected: true, via: 'suite' } : await crmConnection(app, me.team_id).then((c) => (c ? { connected: true, via: c.from, url: c.url } : { connected: false }));
    return {
      me: personView(me), prefs: await app.team.prefs(me),
      team: { id: me.team_id, name: (await app.team.get(me.team_id))?.name ?? null },
      storage: app.db.kind, auth: app.authProvider, demo: !!app.demo, open_signup: !!app.openSignup,
      crm, mcp_url: `${app.publicUrl()}/mcp`, version: app.version,
    };
  },
});

tool({
  name: 'sheets.set_preferences', title: 'Preferences', scope: 'write',
  description: 'Your own preferences: theme (auto, light or dark).',
  input: { theme: z.enum(['auto', 'light', 'dark']).optional() },
  output: Out,
  run: async ({ app, me }, a) => app.team.setPrefs(me, a),
});

// ---------- approvals (confirm: human) ----------

tool({
  name: 'sheets.list_approvals', title: 'Waiting for your yes', scope: 'read',
  description: 'Things an agent asked to do that need your yes first (deleting a spreadsheet, turning on a public link, removing someone).',
  input: {},
  output: Out,
  run: async ({ app, me }) => ({ approvals: (await app.db.all(`select * from sheets_approvals where person_id = $1 and status = 'waiting' order by created_at desc`, [me.id])).map(approvalView) }),
});

tool({
  name: 'sheets.decide_approval', title: 'Approve or decline', scope: 'write',
  description: 'Say yes or no to something an agent asked to do. Yes runs it as you. Only a person in the app can answer.',
  input: { approval: z.string(), approve: z.boolean() },
  output: Out,
  run: async ({ app, me, via }, a) => {
    if (via !== 'web') throw new BookError('Only a person can approve, from the app.', 403, 'forbidden');
    const row = await app.db.get('select * from sheets_approvals where id = $1 and person_id = $2', [a.approval, me.id]);
    if (!row || row.status !== 'waiting') throw new BookError('Nothing waiting with that id.', 404, 'not_found');
    let status = 'declined', result = null;
    if (a.approve) {
      try { result = await runTool(app, me, row.tool, JSON.parse(row.input), { via: 'web', approved: true }); status = 'done'; } catch (e) { result = { error: e.message }; status = 'failed'; }
    }
    await app.db.run('update sheets_approvals set status = $2, result = $3, decided_at = $4 where id = $1', [row.id, status, JSON.stringify(result), nowIso()]);
    return { ...approvalView({ ...row, status }), result };
  },
});

// ---------- running tools ----------

export function listTools() {
  return TOOLS.map((t) => ({ ...t, inputJson: jsonSchema(z.object(t.input)), outputJson: jsonSchema(t.output) }));
}
export const getTool = (name) => TOOLS.find((t) => t.name === name);
export const toWire = (n) => n.replace(/\./g, '_');
export const fromWire = (n) => (getTool(n) ? n : TOOLS.find((t) => toWire(t.name) === n)?.name ?? n);

function jsonSchema(s) {
  const j = zodToJsonSchema(s, { target: 'jsonSchema7', $refStrategy: 'none' });
  delete j.$schema;
  return j;
}

const SCOPES = ['read', 'write', 'delete', 'admin'];
export const CALLED = new Set();

// via: 'web' (a person clicked), 'mcp' or 'rest' (an app acting for a person). call: the suite's call, when inside it.
export async function runTool(app, me, name, input = {}, { via = 'web', scopes = SCOPES, approved = false, client = null, call = null } = {}) {
  const t = getTool(fromWire(name));
  if (!t) throw new BookError(`No tool called ${name}.`, 404, 'no_tool');
  const parsed = z.object(t.input).strict().safeParse(input ?? {});
  if (!parsed.success) throw new BookError(parsed.error.issues.map((i) => `${i.path.join('.') || 'input'}: ${i.message}`).join('; '), 400);
  if (!scopes.includes(t.scope)) throw new BookError(`This connection may not ${t.scope} (it has ${scopes.join(', ')}).`, 403, 'scope');
  if (t.scope === 'admin' && !['owner', 'admin'].includes(me.role)) throw new BookError('Only team owners and admins can do that.', 403, 'forbidden');
  if (me.role === 'viewer' && t.scope !== 'read') throw new BookError('This is a view-only link. Sign in to edit.', 403, 'forbidden');
  if (t.confirm === 'human' && via !== 'web' && !approved) {
    const id = newId('ap');
    await app.db.run('insert into sheets_approvals (id, team_id, person_id, requested_by, tool, input, created_at) values ($1, $2, $3, $4, $5, $6, $7)', [id, me.team_id, me.id, client || via, t.name, JSON.stringify(parsed.data), nowIso()]);
    app.bus?.publish({ team: me.team_id, type: 'sheets.approval.requested', audience: [me.id], data: { approval: approvalView({ id, tool: t.name, input: JSON.stringify(parsed.data), requested_by: client || via, status: 'waiting', created_at: nowIso() }) } });
    CALLED.add(t.name);
    return { pending: { approval_id: id, message: `${t.title} needs a person's yes. ${me.name} has been asked in the app; nothing happened yet.` } };
  }
  const out = await t.run({ app, me, via, call }, parsed.data);
  CALLED.add(t.name);
  return out;
}

function approvalView(r) {
  let input = {};
  try { input = JSON.parse(r.input); } catch {}
  return { id: r.id, tool: r.tool, title: getTool(r.tool)?.title ?? r.tool, input, requested_by: r.requested_by, status: r.status, created_at: r.created_at };
}

export { initBook, setRaw, addTab, tabMap };
