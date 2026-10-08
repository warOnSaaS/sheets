import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createApp } from '../lib/app.mjs';

// On Postgres (SHEETS_TEST_PG=postgres://...): migrations, two server copies sharing one database that hear
// each other's changes over LISTEN/NOTIFY, and folding many changes into the stored state.
const url = process.env.SHEETS_TEST_PG;
process.env.OAUTH_SECRET = 'pg-test';

test('two server copies on one Postgres: a change in one reaches the other live, and both agree', { skip: !url && 'set SHEETS_TEST_PG' }, async () => {
  const team = `t${Math.random().toString(36).slice(2, 8)}`;
  const env = { DATABASE_URL: url, SHEETS_TEAM_ID: team, SHEETS_EXAMPLES: '0' };
  const a = await createApp(env), b = await createApp(env);
  try {
    const sam = await a.team.add(team, { name: 'Sam Rivera', email: `sam-${team}@acme-dental.example`, role: 'owner' });
    const s = await a.run(sam, 'sheets.create_sheet', { values: [['Amount'], [10]] });
    const heard = new Promise((ok) => b.bus.on('event', (e) => { if (e.type === 'sheets.workbook.changed' && e.data.workbook === s.id && e.data.update) ok(e); }));
    await a.run(sam, 'sheets.write_range', { sheet: s.id, range: 'A3', values: [[32]] });
    assert.ok(await heard);
    await b.run(sam, 'sheets.set_formula', { sheet: s.id, cell: 'A4', formula: '=SUM(A2:A3)' });
    assert.equal((await a.run(sam, 'sheets.read_range', { sheet: s.id, range: 'A4' })).values[0][0], 42);
    // Many changes from both copies: they fold into the stored state and nothing is lost.
    for (let i = 0; i < 170; i++) await (i % 2 ? a : b).run(sam, 'sheets.write_range', { sheet: s.id, range: `B${i + 1}`, values: [[i]] });
    const row = await a.db.get('select state_seq from sheets_workbooks where id = $1', [s.id]);
    assert.ok(Number(row.state_seq) > 0, 'the changes were folded');
    const fresh = await createApp(env);
    const r = await fresh.run(sam, 'sheets.read_range', { sheet: s.id, range: 'A1:B170' });
    assert.equal(r.values[169][1], 169);
    assert.equal(r.values[3][0], 42);
    await fresh.close();
  } finally { await a.close(); await b.close(); }
});
