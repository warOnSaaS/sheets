import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StreamableHTTPClientTransport } from '@modelcontextprotocol/sdk/client/streamableHttp.js';

// The agent journey, over MCP only: build a pipeline sheet with formulas and a chart, read it back exactly,
// download it as .xlsx, and check the file with openpyxl (an independent reader). Then a workbook made by
// openpyxl, with formulas and no saved results, goes in and out, and every computed value is checked.
// The openpyxl part needs Python with openpyxl (python3 -m venv .venv && .venv/bin/pip install openpyxl).

process.env.OAUTH_SECRET = 'agent-test';
process.env.SQLITE_FILE = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'sheets-agent-')), 'sheets.db');
const PY = ['.venv/bin/python', 'python3'].find((p) => { try { execFileSync(p, ['-c', 'import openpyxl'], { stdio: 'ignore' }); return true; } catch { return false; } });
const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'sheets-xlsx-'));

let server, app, client;
const tool = async (name, args) => {
  const r = await client.callTool({ name, arguments: args });
  if (r.isError) throw new Error(r.content[0].text);
  return r.structuredContent;
};

before(async () => {
  const { createServer } = await import('../server.mjs');
  const { issueTokens } = await import('../lib/auth.mjs');
  server = createServer();
  await new Promise((r) => server.listen(0, r));
  app = await server.ready;
  const sam = await app.team.add(app.teamId, { name: 'Sam Rivera', email: 'sam@acme-dental.example', role: 'owner' });
  client = new Client({ name: 'pipeline-agent', version: '1' });
  await client.connect(new StreamableHTTPClientTransport(new URL(`http://localhost:${server.address().port}/mcp`), { requestInit: { headers: { authorization: `Bearer ${issueTokens(sam, { app: 'pipeline-agent' }).access_token}` } } }));
});
after(async () => { await client?.close(); server?.closeAllConnections?.(); server?.close(); await app?.close(); });

const deals = [
  ['Staff dental plan', 'Birch Law', 'Negotiation', 18000, '2026-10-16'],
  ['Family plan add-on', 'Harbor Fitness', 'Proposal', 12000, '2026-10-21'],
  ['School screening day', 'Pinecrest School', 'Qualified', 6200, '2026-11-12'],
  ['Driver checkups', 'Summit Logistics', 'Proposal', 22400, '2026-11-03'],
  ['Cleaning days', 'Maple Street Cafe', 'Won', 2600, '2026-08-28'],
];
const prob = { Qualified: 0.25, Proposal: 0.5, Negotiation: 0.75, Won: 1 };

let sheetId, exported;

test('an agent builds a pipeline sheet with formulas and a chart, over MCP only', async () => {
  const s = await tool('sheets_create_sheet', { title: 'Q4 pipeline (agent)', tabs: ['Deals', 'Stages', 'Summary'] });
  sheetId = s.id;
  await tool('sheets_write_range', { sheet: s.id, tab: 'Stages', values: [['Stage', 'Probability'], ...Object.entries(prob).map(([k, v]) => [k, `${v * 100}%`])] });
  await tool('sheets_write_range', { sheet: s.id, tab: 'Deals', values: [['Deal', 'Company', 'Stage', 'Amount', 'Close date', 'Probability', 'Weighted'], ...deals] });
  const p = await tool('sheets_set_formula', { sheet: s.id, tab: 'Deals', cell: 'F2', formula: '=VLOOKUP(C2,Stages!$A$2:$B$5,2,FALSE)', fill: 'F2:F6' });
  assert.deepEqual(p.values.map((r) => r[0]), deals.map((d) => prob[d[2]]));
  const w = await tool('sheets_set_formula', { sheet: s.id, tab: 'Deals', cell: 'G2', formula: '=D2*F2', fill: 'G2:G6' });
  assert.deepEqual(w.values.map((r) => r[0]), deals.map((d) => d[3] * prob[d[2]]));
  await tool('sheets_append_rows', { sheet: s.id, tab: 'Deals', rows: [['Total', null, null, '=SUM(D2:D6)', null, null, '=SUM(G2:G6)']] });
  await tool('sheets_format_range', { sheet: s.id, tab: 'Deals', range: 'D2:D7', number_format: 'currency0' });
  await tool('sheets_format_range', { sheet: s.id, tab: 'Deals', range: 'G2:G7', number_format: 'currency0' });
  await tool('sheets_format_range', { sheet: s.id, tab: 'Deals', range: 'A1:G1', bold: true });
  await tool('sheets_freeze', { sheet: s.id, tab: 'Deals', rows: 1 });
  await tool('sheets_write_range', { sheet: s.id, tab: 'Summary', values: [['Stage', 'Amount', 'Weighted'], ...Object.keys(prob).map((k) => [k])] });
  await tool('sheets_set_formula', { sheet: s.id, tab: 'Summary', cell: 'B2', formula: '=SUMIF(Deals!$C$2:$C$6,A2,Deals!$D$2:$D$6)', fill: 'B2:B5' });
  const sum = await tool('sheets_set_formula', { sheet: s.id, tab: 'Summary', cell: 'C2', formula: '=SUMIF(Deals!$C$2:$C$6,A2,Deals!$G$2:$G$6)', fill: 'C2:C5' });
  assert.deepEqual(sum.values.map((r) => r[0]), [1550, 17200, 13500, 2600]);
  const chart = await tool('sheets_create_chart', { sheet: s.id, tab: 'Summary', range: 'A1:C5', type: 'column', title: 'Pipeline by stage' });
  assert.deepEqual(chart.data.labels, ['Qualified', 'Proposal', 'Negotiation', 'Won']);
  assert.deepEqual(chart.data.series.map((x) => x.name), ['Amount', 'Weighted']);
  assert.deepEqual(chart.data.series[0].values, [6200, 34400, 18000, 2600]);
  // Reading back is exact: values and the formulas that made them.
  const r = await tool('sheets_read_range', { sheet: s.id, tab: 'Deals', range: 'A7:G7', text: true });
  assert.deepEqual(r.values[0], ['Total', null, null, 61200, null, null, 34850]);
  assert.deepEqual(r.formulas[0], [null, null, null, '=SUM(D2:D6)', null, null, '=SUM(G2:G6)']);
  assert.equal(r.text[0][6], '$34,850');
  const g = await tool('sheets_get_sheet', { sheet: s.id });
  assert.equal(g.charts.length, 1);
  assert.equal(g.tabs.find((t) => t.name === 'Deals').frozen_rows, 1);
});

test('the agent downloads it as .xlsx; openpyxl reads the same formulas and values', { skip: !PY && 'no Python with openpyxl' }, async () => {
  const x = await tool('sheets_export', { sheet: sheetId, format: 'xlsx', content: true });
  exported = path.join(dir, 'pipeline.xlsx');
  fs.writeFileSync(exported, Buffer.from(x.content_base64, 'base64'));
  const cells = JSON.parse(execFileSync(PY, ['scripts/xlsx_check.py', 'read', exported], { encoding: 'utf8' }));
  assert.equal(cells['Deals!G7'].formula, '=SUM(G2:G6)');
  assert.equal(cells['Deals!G7'].value, 34850);
  assert.equal(cells['Deals!F3'].formula, '=VLOOKUP(C3,Stages!$A$2:$B$5,2,FALSE)');
  assert.equal(cells['Deals!F3'].value, 0.5);
  assert.equal(cells['Deals!E2'].value, '2026-10-16');
  assert.equal(cells['Summary!C3'].value, 17200);
  // Every formula's saved value matches what Sheets reads over MCP.
  for (const tab of ['Deals', 'Summary']) {
    const r = await tool('sheets_read_range', { sheet: sheetId, tab });
    r.values.forEach((row, i) => row.forEach((v, j) => {
      const f = r.formulas?.[i]?.[j];
      if (!f) return;
      const c = cells[`${tab}!${String.fromCharCode(65 + j)}${i + 1}`];
      assert.equal(c.formula, f, `${tab} ${i},${j} formula`);
      assert.ok(Math.abs(c.value - v) < 1e-9 || c.value === v, `${tab} ${i},${j}: file ${c.value}, Sheets ${v}`);
    }));
  }
});

test('an xlsx made elsewhere, with formulas and no saved results, round-trips with the right values', { skip: !PY && 'no Python with openpyxl' }, async () => {
  const src = path.join(dir, 'made-by-openpyxl.xlsx');
  const expected = JSON.parse(execFileSync(PY, ['scripts/xlsx_check.py', 'make', src], { encoding: 'utf8' }));
  const imp = await tool('sheets_import_xlsx', { content_base64: fs.readFileSync(src).toString('base64'), name: 'made-by-openpyxl.xlsx' });
  assert.equal(imp.report.tabs, 2);
  assert.equal(imp.report.unsupported_functions.length, 0);
  assert.equal(imp.report.formula_problems.length, 0, JSON.stringify(imp.report.formula_problems));
  for (const [ref, want] of Object.entries(expected)) {
    const [tab, cell] = ref.split('!');
    const r = await tool('sheets_read_range', { sheet: imp.sheet, tab, range: cell });
    assert.equal(r.values[0][0], want, `${ref} in Sheets`);
  }
  const x = await tool('sheets_export', { sheet: imp.sheet, content: true });
  const out = path.join(dir, 'round-trip.xlsx');
  fs.writeFileSync(out, Buffer.from(x.content_base64, 'base64'));
  const cells = JSON.parse(execFileSync(PY, ['scripts/xlsx_check.py', 'read', out], { encoding: 'utf8' }));
  const orig = JSON.parse(execFileSync(PY, ['scripts/xlsx_check.py', 'read', src], { encoding: 'utf8' }));
  for (const [ref, want] of Object.entries(expected)) {
    assert.equal(cells[ref].value, want, `${ref} value after export`);
    assert.equal(cells[ref].formula, orig[ref].formula, `${ref} formula after export`);
  }
  // Plain values and dates come back as they went in.
  for (const [ref, c] of Object.entries(orig)) if (!c.formula) assert.deepEqual(cells[ref]?.value, c.value, ref);
  fs.copyFileSync(out, path.join(dir, 'kept.xlsx'));
});
