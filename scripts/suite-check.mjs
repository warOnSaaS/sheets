// Loads Sheets into a real wOS suite core (a checkout of warOnSaaS/suite) and drives it there: turn it on,
// build a sheet with formulas, a pivot and a chart as two members, export, then turn it off.
// Needs Node 22.6 or newer (the suite is TypeScript run directly) and the suite next to this repo.
//   node scripts/suite-check.mjs [path-to-suite]      (default ~/wos-suite)
import path from 'node:path';
import os from 'node:os';
import assert from 'node:assert/strict';

const suite = path.resolve(process.argv[2] ?? path.join(os.homedir(), 'wos-suite'));
const appDir = path.resolve(new URL('..', import.meta.url).pathname);
const { makeCore, person } = await import(path.join(suite, 'test/unit/helpers.ts'));

const core = await makeCore({ WOS_APPS: appDir });
const sam = await person(core, 'Sam');
const jordan = await person(core, 'Jordan', 'member', sam.team);
await sam.call('apps.enable', { app: 'sheets' });
const s = await sam.call('sheets.create_sheet', { title: 'Inside the suite', values: [['Stage', 'Amount'], ['Won', 100], ['Open', 50], ['Won', 70]] });
assert.match(s.url, /\/a\/sheets\/s\//);
const f = await jordan.call('sheets.set_formula', { sheet: s.id, cell: 'C2', formula: '=B2*2', fill: 'C2:C4' });
assert.deepEqual(f.values.map((r) => r[0]), [200, 100, 140]);
const pv = await sam.call('sheets.create_pivot', { sheet: s.id, rows: 'Stage', values: [{ field: 'Amount', summarize: 'sum' }] });
assert.deepEqual(pv.values.at(-1), ['Grand total', 220]);
const ch = await jordan.call('sheets.create_chart', { sheet: s.id, range: 'A1:B4' });
assert.deepEqual(ch.data.series[0].values, [100, 50, 70]);
const x = await sam.call('sheets.export', { sheet: s.id });
assert.match(x.file.url, /^\/files\/sheets\//);
const people = await sam.call('sheets.list_people');
assert.ok(people.people.some((p) => p.name === 'Jordan'), 'members come from the suite');
const tools = [...core.catalogue.tools.keys()].filter((n) => n.startsWith('sheets.'));
console.log(`suite check: Sheets loaded into the suite, ${tools.length} tools in the catalogue; formulas, fill, pivot, chart and export work as two members.`);
await sam.call('apps.disable', { app: 'sheets' });
await assert.rejects(sam.call('sheets.list_sheets'));
console.log('suite check: turned off, its tools are gone.');
await core.stop();
process.exit(0);
