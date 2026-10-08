import { test } from 'node:test';
import assert from 'node:assert/strict';
import { makeApp, fakeCrm } from './helpers.mjs';
import { listTools, CALLED } from '../lib/tools.mjs';

// Every tool, the way an agent or a screen calls it. The last test fails if any tool was never called.

test('make a spreadsheet, write cells as typed, read exact values and formulas', async () => {
  const { sam, run } = await makeApp();
  const s = await run(sam, 'sheets.create_sheet', { title: 'Q4 pipeline', tabs: ['Deals', 'Summary'], values: [['Deal', 'Amount', 'Stage'], ['Birch Law', '$18,000', 'Won'], ['Harbor Fitness', '12000', 'Open']] });
  assert.equal(s.title, 'Q4 pipeline');
  assert.deepEqual(s.tabs.map((t) => t.name), ['Deals', 'Summary']);
  const w = await run(sam, 'sheets.write_range', { sheet: s.id, tab: 'Deals', range: 'A4', values: [['Pinecrest School', '6.5%', 'Open'], ['Total', '=SUM(B2:B4)', null]] });
  assert.equal(w.values[1][1], 30000.065);
  const r = await run(sam, 'sheets.read_range', { sheet: 'q4 PIPELINE', tab: 'Deals', text: true, formats: true });
  assert.equal(r.range, 'A1:C5');
  assert.equal(r.values[1][1], 18000);
  assert.equal(r.text[1][1], '$18,000');
  assert.equal(r.formulas[4][1], '=SUM(B2:B4)');
  assert.equal(r.formats[1][1].nf, '$#,##0');
  // A formula error comes back with the write, so an agent sees it at once.
  const bad = await run(sam, 'sheets.set_formula', { sheet: s.id, tab: 'Deals', cell: 'D2', formula: 'B2/0' });
  assert.equal(bad.values[0][0], '#DIV/0!');
  assert.match(bad.errors[0], /D2/);
  await assert.rejects(run(sam, 'sheets.set_formula', { sheet: s.id, cell: 'D2', formula: '=SUM(B2:B4' }), /not a formula/);
  // Fill a formula down; relative references move.
  const f = await run(sam, 'sheets.set_formula', { sheet: s.id, tab: 'Deals', cell: 'D2', formula: '=B2*2', fill: 'D2:D4' });
  assert.deepEqual(f.formulas.map((x) => x[0]), ['=B2*2', '=B3*2', '=B4*2']);
  assert.deepEqual(f.values.map((x) => x[0]), [36000, 24000, 0.13]);
  // Cross-tab formulas, and a rename that rewrites them.
  await run(sam, 'sheets.write_range', { sheet: s.id, tab: 'Summary', values: [['Won', '=SUMIF(Deals!C2:C4,"Won",Deals!B2:B4)']] });
  const ren = await run(sam, 'sheets.rename_tab', { sheet: s.id, tab: 'Deals', name: 'All deals' });
  assert.equal(ren.tab.name, 'All deals');
  const sum = await run(sam, 'sheets.read_range', { sheet: s.id, tab: 'Summary', range: 'B1' });
  assert.equal(sum.formulas[0][0], "=SUMIF('All deals'!C2:C4,\"Won\",'All deals'!B2:B4)");
  assert.equal(sum.values[0][0], 18000);
  const list = await run(sam, 'sheets.list_sheets', { q: 'q4' });
  assert.equal(list.sheets.length, 1);
  const got = await run(sam, 'sheets.get_sheet', { sheet: s.id });
  assert.equal(got.tabs[0].used, 'A1:D5');
  await run(sam, 'sheets.rename_sheet', { sheet: s.id, title: 'Q4 deals' });
  assert.equal((await run(sam, 'sheets.get_sheet', { sheet: s.id })).title, 'Q4 deals');
});

test('append, clear, find, sort keeping each row\'s formulas, filter, insert and delete rows and columns', async () => {
  const { sam, run } = await makeApp();
  const s = await run(sam, 'sheets.create_sheet', { values: [['Name', 'Amount', 'Double'], ['Casey', 30, '=B2*2'], ['Avery', 10, '=B3*2'], ['Riley', 20, '=B4*2']] });
  const a = await run(sam, 'sheets.append_rows', { sheet: s.id, rows: [['Jordan', 25, '=B5*2']] });
  assert.equal(a.range, 'A5:C5');
  const sorted = await run(sam, 'sheets.sort', { sheet: s.id, by: { column: 'Amount', order: 'desc' } });
  assert.equal(sorted.header, true);
  const r = await run(sam, 'sheets.read_range', { sheet: s.id, range: 'A2:C5' });
  assert.deepEqual(r.values.map((x) => x[0]), ['Casey', 'Jordan', 'Riley', 'Avery']);
  assert.deepEqual(r.formulas.map((x) => x[2]), ['=B2*2', '=B3*2', '=B4*2', '=B5*2']);
  assert.deepEqual(r.values.map((x) => x[2]), [60, 50, 40, 20]);
  const found = await run(sam, 'sheets.find', { sheet: s.id, query: 'ri' });
  assert.deepEqual(found.matches.map((m) => m.cell), ['A4']);
  const ff = await run(sam, 'sheets.find', { sheet: s.id, query: 'B5', formulas: true });
  assert.equal(ff.matches[0].cell, 'C5');
  const fl = await run(sam, 'sheets.filter', { sheet: s.id, column: 'Amount', condition: { op: 'gte', value: 25 } });
  assert.deepEqual(fl.shown.map((x) => x.values[0]), ['Name', 'Casey', 'Jordan']);
  assert.equal(fl.hidden, 2);
  assert.equal((await run(sam, 'sheets.filter', { sheet: s.id, clear: true })).filter, null);
  await run(sam, 'sheets.write_range', { sheet: s.id, range: 'E1', values: [['=SUM(B2:B5)']] });
  await run(sam, 'sheets.insert_rows', { sheet: s.id, at: 2, count: 2 });
  let t = await run(sam, 'sheets.read_range', { sheet: s.id, range: 'E1' });
  assert.equal(t.formulas[0][0], '=SUM(B4:B7)');
  assert.equal(t.values[0][0], 85);
  await run(sam, 'sheets.delete_rows', { sheet: s.id, at: 2, count: 2 });
  await run(sam, 'sheets.insert_columns', { sheet: s.id, at: 'B' });
  t = await run(sam, 'sheets.read_range', { sheet: s.id, range: 'F1' });
  assert.equal(t.formulas[0][0], '=SUM(C2:C5)');
  await run(sam, 'sheets.delete_columns', { sheet: s.id, at: 2 });
  t = await run(sam, 'sheets.read_range', { sheet: s.id, range: 'E1' });
  assert.equal(t.formulas[0][0], '=SUM(B2:B5)');
  const c = await run(sam, 'sheets.clear_range', { sheet: s.id, range: 'C2:C5' });
  assert.equal(c.cleared, 4);
});

test('format, widths, freeze, tabs', async () => {
  const { sam, run } = await makeApp();
  const s = await run(sam, 'sheets.create_sheet', { values: [['Amount'], [0.256]] });
  await run(sam, 'sheets.format_range', { sheet: s.id, range: 'A1', bold: true, fill: '#f1f3f5', align: 'center' });
  await run(sam, 'sheets.format_range', { sheet: s.id, range: 'A2', number_format: 'percent' });
  const r = await run(sam, 'sheets.read_range', { sheet: s.id, text: true, formats: true });
  assert.deepEqual(r.formats[0][0], { b: true, bg: '#f1f3f5', ha: 'center' });
  assert.equal(r.text[1][0], '25.6%');
  await assert.rejects(run(sam, 'sheets.format_range', { sheet: s.id, range: 'A1', color: 'red' }), /colour/);
  await run(sam, 'sheets.set_column_width', { sheet: s.id, columns: 'A:B', width: 180 });
  await run(sam, 'sheets.freeze', { sheet: s.id, rows: 1, columns: 1 });
  await run(sam, 'sheets.add_sheet_tab', { sheet: s.id, name: 'Notes' });
  await run(sam, 'sheets.move_tab', { sheet: s.id, tab: 'Notes', index: 0 });
  let g = await run(sam, 'sheets.get_sheet', { sheet: s.id });
  assert.deepEqual(g.tabs.map((t) => t.name), ['Notes', 'Sheet1']);
  assert.equal(g.tabs[1].frozen_rows, 1);
  await run(sam, 'sheets.delete_tab', { sheet: s.id, tab: 'Notes' });
  g = await run(sam, 'sheets.get_sheet', { sheet: s.id });
  assert.deepEqual(g.tabs.map((t) => t.name), ['Sheet1']);
  await assert.rejects(run(sam, 'sheets.delete_tab', { sheet: s.id, tab: 'Sheet1' }), /at least one tab/);
});

test('charts, pivots and comments', async () => {
  const { sam, jordan, run } = await makeApp();
  const s = await run(sam, 'sheets.create_sheet', { values: [['Stage', 'Owner', 'Amount'], ['Won', 'Sam', 100], ['Open', 'Sam', 50], ['Won', 'Jordan', 70], ['Open', 'Jordan', 30]] });
  const ch = await run(sam, 'sheets.create_chart', { sheet: s.id, range: 'A1:C5', series: ['C'], title: 'By deal' });
  assert.deepEqual(ch.data.series[0].values, [100, 50, 70, 30]);
  const up = await run(sam, 'sheets.update_chart', { sheet: s.id, chart: ch.id, type: 'bar', title: 'Amounts' });
  assert.equal(up.type, 'bar');
  assert.equal((await run(sam, 'sheets.list_charts', { sheet: s.id })).charts.length, 1);
  await run(sam, 'sheets.delete_chart', { sheet: s.id, chart: ch.id });
  const pv = await run(sam, 'sheets.create_pivot', { sheet: s.id, rows: 'Stage', columns: 'Owner', values: [{ field: 'Amount', summarize: 'sum' }] });
  assert.equal(pv.tab, 'Pivot of Sheet1');
  assert.deepEqual(pv.values, [['Stage', 'Sum of Amount · Jordan', 'Sum of Amount · Sam'], ['Open', 30, 50], ['Won', 70, 100], ['Grand total', 100, 150]]);
  assert.match(pv.formulas[1][1], /^=SUMIFS\(/);
  // The pivot is live: change the table and it follows; a new stage appears on refresh.
  await run(sam, 'sheets.write_range', { sheet: s.id, range: 'C2', values: [[200]] });
  assert.equal((await run(sam, 'sheets.read_range', { sheet: s.id, tab: pv.tab, range: 'C3' })).values[0][0], 200);
  await run(sam, 'sheets.append_rows', { sheet: s.id, rows: [['Lost', 'Sam', 5]] });
  await run(sam, 'sheets.refresh_link', { sheet: s.id, tab: pv.tab });
  const after = await run(sam, 'sheets.read_range', { sheet: s.id, tab: pv.tab });
  assert.deepEqual(after.values.map((x) => x[0]), ['Stage', 'Lost', 'Open', 'Won', 'Grand total']);
  const cm = await run(sam, 'sheets.add_comment', { sheet: s.id, cell: 'C2', body: 'Check this with Jordan' });
  await run(jordan, 'sheets.reply_comment', { sheet: s.id, comment: cm.comment.id, body: 'Confirmed' });
  let list = await run(sam, 'sheets.list_comments', { sheet: s.id });
  assert.equal(list.comments[0].replies[0].by.name, 'Jordan Lee');
  await run(sam, 'sheets.resolve_comment', { sheet: s.id, comment: cm.comment.id });
  list = await run(sam, 'sheets.list_comments', { sheet: s.id });
  assert.equal(list.comments.length, 0);
  await assert.rejects(run(jordan, 'sheets.delete_comment', { sheet: s.id, comment: cm.comment.id }), /Only the person/);
  await run(sam, 'sheets.delete_comment', { sheet: s.id, comment: cm.comment.id });
});

test('versions: name one, read it, restore it', async () => {
  const { sam, run } = await makeApp();
  const s = await run(sam, 'sheets.create_sheet', { values: [['a', 1]] });
  const v = await run(sam, 'sheets.save_version', { sheet: s.id, label: 'Before' });
  await run(sam, 'sheets.write_range', { sheet: s.id, values: [['b', 2]] });
  await run(sam, 'sheets.add_sheet_tab', { sheet: s.id, name: 'Extra' });
  const old = await run(sam, 'sheets.read_range', { sheet: s.id, version: v.version.id });
  assert.deepEqual(old.values, [['a', 1]]);
  const list = await run(sam, 'sheets.list_versions', { sheet: s.id });
  assert.ok(list.versions.some((x) => x.label === 'Before'));
  await run(sam, 'sheets.restore_version', { sheet: s.id, version: v.version.id });
  const now = await run(sam, 'sheets.read_range', { sheet: s.id });
  assert.deepEqual(now.values, [['a', 1]]);
  assert.deepEqual((await run(sam, 'sheets.get_sheet', { sheet: s.id })).tabs.map((t) => t.name), ['Sheet1']);
  assert.ok((await run(sam, 'sheets.list_versions', { sheet: s.id })).versions.some((x) => /^Before restoring/.test(x.label ?? '')));
});

test('xlsx and csv in and out', async () => {
  const { sam, run } = await makeApp();
  const s = await run(sam, 'sheets.create_sheet', { template: 'pipeline' });
  const x = await run(sam, 'sheets.export', { sheet: s.id, content: true });
  assert.match(x.file.name, /\.xlsx$/);
  assert.match(x.not_carried[0], /chart/);
  const back = await run(sam, 'sheets.import_xlsx', { content_base64: x.content_base64, name: 'pipeline.xlsx' });
  assert.equal(back.report.tabs, 3);
  assert.equal(back.report.formulas, back.report.formulas_matching_excel);
  assert.equal(back.report.formula_problems.length, 0);
  const a = await run(sam, 'sheets.read_range', { sheet: s.id, tab: 'Summary' });
  const b = await run(sam, 'sheets.read_range', { sheet: back.sheet, tab: 'Summary' });
  assert.deepEqual(b.values, a.values);
  assert.deepEqual(b.formulas, a.formulas);
  const csv = await run(sam, 'sheets.export', { sheet: s.id, format: 'csv', tab: 'Summary', content: true });
  const text = Buffer.from(csv.content_base64, 'base64').toString();
  assert.match(text, /^Stage,Deals,Amount,Weighted\r\nLead,2,"\$14,500","\$1,450"/);
  const c1 = await run(sam, 'sheets.import_csv', { content: 'Name;Amount;Date\nAcme;"1,200";2026-10-07\n', name: 'semi.csv' });
  assert.equal(c1.report.delimiter, ';');
  const r = await run(sam, 'sheets.read_range', { sheet: c1.sheet, text: true });
  assert.deepEqual(r.values[1], ['Acme', 1200, 46302]);
  const c2 = await run(sam, 'sheets.import_csv', { sheet: s.id, content_base64: Buffer.from('x,y\n1,2\n').toString('base64'), name: 'More' });
  assert.equal(c2.tab, 'More');
  await assert.rejects(run(sam, 'sheets.import_xlsx', { content_base64: Buffer.from('not a zip').toString('base64') }), /not an \.xlsx/);
  const all = await run(sam, 'sheets.export_data', {});
  assert.equal(all.counts.sheets, 3);
});

test('the CRM: pull deals onto a linked tab, then refresh', async () => {
  const { app, sam, run } = await makeApp();
  let recs = [{ id: 'd1', Name: 'Staff plan', 'Account.Name': 'Birch Law', StageName: 'won', Amount: 18000, CloseDate: '2026-10-16' }];
  const crm = await fakeCrm(() => recs);
  await assert.rejects(run(sam, 'sheets.import_from_crm', { sheet: 'x', kind: 'deals' }), /No spreadsheet/);
  const s = await run(sam, 'sheets.create_sheet', { title: 'From the CRM' });
  await assert.rejects(run(sam, 'sheets.import_from_crm', { sheet: s.id, kind: 'deals' }), /No CRM is connected/);
  const c = await run(sam, 'sheets.connect_crm', { url: crm.url, token: 'crm-token' });
  assert.equal(c.connected, true);
  const p = await run(sam, 'sheets.import_from_crm', { sheet: s.id, kind: 'deals', where: 'Amount > 1000' });
  assert.equal(p.rows, 1);
  assert.equal(crm.calls.at(-1).auth, 'Bearer crm-token');
  assert.match(crm.calls.at(-1).body.soql, /FROM Opportunity WHERE Amount > 1000 LIMIT 2000/);
  // A formula beside the table keeps working after a refresh that adds rows.
  await run(sam, 'sheets.write_range', { sheet: s.id, tab: p.tab, range: 'H1', values: [['=SUM(D2:D100)']] });
  recs = [...recs, { id: 'd2', Name: 'Family plan', 'Account.Name': 'Harbor Fitness', StageName: 'proposal', Amount: 12000, CloseDate: '2026-10-21' }];
  const rf = await run(sam, 'sheets.refresh_link', { sheet: s.id });
  assert.equal(rf.refreshed[0].rows, 2);
  const r = await run(sam, 'sheets.read_range', { sheet: s.id, tab: p.tab, text: true });
  assert.deepEqual(r.values[0], ['Name', 'Organization', 'Stage', 'Amount', 'Close date', 'CRM id', null, 30000]);
  assert.equal(r.text[2][4], '2026-10-21');
  const st = await run(sam, 'sheets.get_settings', {});
  assert.equal(st.crm.connected, true);
  await run(sam, 'sheets.connect_crm', { disconnect: true });
  assert.equal((await run(sam, 'sheets.get_settings', {})).crm.connected, false);
  void app;
});

test('confirm: human tools wait for a person; viewers and scopes are held to their limits', async () => {
  const { sam, jordan, run } = await makeApp();
  const s = await run(sam, 'sheets.create_sheet', { title: 'Keep' });
  const p = await run(sam, 'sheets.delete_sheet', { sheet: s.id }, { via: 'mcp', client: 'Claude' });
  assert.ok(p.pending.approval_id);
  assert.equal((await run(sam, 'sheets.list_sheets', {})).sheets.length, 1);
  const w = await run(sam, 'sheets.list_approvals', {});
  assert.equal(w.approvals[0].tool, 'sheets.delete_sheet');
  await assert.rejects(run(sam, 'sheets.decide_approval', { approval: p.pending.approval_id, approve: true }, { via: 'mcp' }), /Only a person/);
  const d = await run(sam, 'sheets.decide_approval', { approval: p.pending.approval_id, approve: true });
  assert.equal(d.status, 'done');
  assert.equal((await run(sam, 'sheets.list_sheets', {})).sheets.length, 0);
  // Sharing a public link is the same.
  const s2 = await run(sam, 'sheets.create_sheet', { title: 'Shared' });
  assert.ok((await run(sam, 'sheets.share', { sheet: s2.id, link: 'view' }, { via: 'mcp' })).pending);
  const on = await run(sam, 'sheets.share', { sheet: s2.id, link: 'view' });
  assert.match(on.url, /\/v\/[\w-]{20,}$/);
  const token = on.url.split('/v/')[1];
  const copy = await run(jordan, 'sheets.copy_sheet', { share_token: token });
  assert.equal(copy.title, 'Shared');
  await run(sam, 'sheets.copy_sheet', { sheet: s2.id });
  await run(sam, 'sheets.share', { sheet: s2.id, link: 'off' });
  await assert.rejects(run(jordan, 'sheets.connect_crm', { url: 'https://crm.example.com' }), /owners and admins/);
  await assert.rejects(run(sam, 'sheets.write_range', { sheet: s2.id, values: [[1]] }, { via: 'mcp', scopes: ['read'] }), /may not write/);
  // People.
  const added = await run(sam, 'sheets.add_person', { name: 'Riley Chen', email: 'riley@acme-dental.example' });
  assert.equal(added.role, 'member');
  assert.equal((await run(sam, 'sheets.list_people', {})).people.length, 4);
  assert.ok((await run(sam, 'sheets.remove_person', { person: 'riley@acme-dental.example' }, { via: 'mcp' })).pending);
  await run(sam, 'sheets.remove_person', { person: 'riley@acme-dental.example' });
  assert.equal((await run(sam, 'sheets.list_people', {})).people.length, 3);
  assert.equal((await run(sam, 'sheets.set_preferences', { theme: 'dark' })).theme, 'dark');
});

test('screens catch up with sync and show where people are', async () => {
  const { sam, run } = await makeApp();
  const s = await run(sam, 'sheets.create_sheet', { values: [[1]] });
  const all = await run(sam, 'sheets.sync', { sheet: s.id });
  assert.ok(all.update.length > 10);
  assert.deepEqual(await run(sam, 'sheets.set_presence', { sheet: s.id, tab: 'Sheet1', cell: 'A1' }), { ok: true });
});

test('every tool in the catalogue was called', () => {
  const missing = listTools().map((t) => t.name).filter((n) => !CALLED.has(n));
  assert.deepEqual(missing, []);
});
