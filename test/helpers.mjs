import http from 'node:http';
import { after } from 'node:test';
import { createApp } from '../lib/app.mjs';

process.env.OAUTH_SECRET = 'test-secret';
const open = [];
after(async () => { for (const a of open) await a.close().catch(() => {}); });

// A fresh app on an in-memory SQLite database (or SHEETS_TEST_PG when given), with a team of three.
export async function makeApp(extra = {}) {
  const env = { ...(process.env.SHEETS_TEST_PG ? { DATABASE_URL: process.env.SHEETS_TEST_PG } : {}), SQLITE_FILE: ':memory:', SHEETS_TEAM_ID: `t${Math.random().toString(36).slice(2, 8)}`, SHEETS_TEAM_NAME: 'Acme Dental', SHEETS_EXAMPLES: '0', ...extra };
  const app = await createApp(env);
  open.push(app);
  const sam = await app.team.add(app.teamId, { name: 'Sam Rivera', email: 'sam@acme-dental.example', role: 'owner' });
  const jordan = await app.team.add(app.teamId, { name: 'Jordan Lee', email: 'jordan@acme-dental.example', role: 'member' });
  const casey = await app.team.add(app.teamId, { name: 'Casey Morgan', github: 'casey-gh', role: 'member' });
  return { app, sam, jordan, casey, run: (me, name, input, o) => app.run(me, name, input, o) };
}

// A stand-in for the wOS CRM: answers crm.query like the real one does.
export async function fakeCrm(records) {
  const calls = [];
  const s = http.createServer(async (req, res) => {
    let body = '';
    for await (const c of req) body += c;
    calls.push({ path: req.url, auth: req.headers.authorization ?? null, body: JSON.parse(body || '{}') });
    if (req.url !== '/api/tools/crm.query') return res.writeHead(404).end('{}');
    const cols = ['Name', 'Account.Name', 'StageName', 'Amount', 'CloseDate'];
    res.writeHead(200, { 'content-type': 'application/json' }).end(JSON.stringify({ result: `${records().length} rows`, data: { kind: 'deals', object: 'Opportunity', columns: cols, records: records(), totalSize: records().length, done: true, grouped: false } }));
  });
  await new Promise((r) => s.listen(0, r));
  open.push({ close: async () => s.close() });
  return { url: `http://localhost:${s.address().port}`, calls };
}
