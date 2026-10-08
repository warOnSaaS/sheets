import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import WebSocket from 'ws';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StreamableHTTPClientTransport } from '@modelcontextprotocol/sdk/client/streamableHttp.js';
import * as Y from 'yjs';

process.env.OAUTH_SECRET = 'server-test';
process.env.SQLITE_FILE = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'sheets-srv-')), 'sheets.db');
process.env.SHEETS_EXAMPLES = '1';

let server, base, app, sam, jordan, issueTokens;
before(async () => {
  const mod = await import('../server.mjs');
  ({ issueTokens } = await import('../lib/auth.mjs'));
  server = mod.createServer();
  await new Promise((r) => server.listen(0, r));
  app = await server.ready;
  base = `http://localhost:${server.address().port}`;
  sam = await app.team.add(app.teamId, { name: 'Sam Rivera', email: 'sam@acme-dental.example', role: 'owner' });
  jordan = await app.team.add(app.teamId, { name: 'Jordan Lee', email: 'jordan@acme-dental.example' });
});
after(async () => { server.closeAllConnections?.(); server.close(); await app.close(); });

const bearer = (p, scopes) => `Bearer ${issueTokens(p, scopes ? { scopes } : {}).access_token}`;
const call = async (name, input, auth = bearer(sam), extra = {}) => {
  const r = await fetch(`${base}/api/tools/${name}`, { method: 'POST', headers: { 'content-type': 'application/json', ...(auth ? { authorization: auth } : {}), ...extra }, body: JSON.stringify(input) });
  return { status: r.status, body: await r.json() };
};

test('pages: look freely, sign in to use', async () => {
  const home = await fetch(`${base}/`).then((r) => r.text());
  assert.match(home, /Host it yourself, free/);
  assert.match(home, /Acme Dental pipeline/);
  const token = /href="\/v\/([^"]+)"/.exec(home)[1];
  const view = await fetch(`${base}/v/${token}`);
  assert.equal(view.status, 200);
  assert.match(await view.text(), /"view":\{"token"/);
  assert.equal((await call('sheets.list_sheets', {}, null)).status, 401);
  const r = await fetch(`${base}/api/tools/sheets.list_sheets`, { method: 'POST' });
  assert.match(r.headers.get('www-authenticate'), /resource_metadata/);
  assert.equal((await fetch(`${base}/health`).then((x) => x.json())).ok, true);
  const meta = await fetch(`${base}/.well-known/oauth-authorization-server`).then((x) => x.json());
  assert.equal(meta.token_endpoint, `${base}/oauth/token`);
  const script = await fetch(`${base}/connect/claude`).then((x) => x.text());
  assert.match(script, /claude mcp add --transport http --scope user/);
  const openapi = await fetch(`${base}/openapi.json`).then((x) => x.json());
  assert.ok(openapi.paths['/api/tools/sheets_read_range']);
});

test('a view-only link reads one spreadsheet and nothing else', async () => {
  const token = app.exampleList[0].token;
  const h = { 'x-share-token': token };
  const synced = await call('sheets.sync', {}, null, h);
  assert.equal(synced.status, 200);
  assert.equal(synced.body.result.title, 'Acme Dental pipeline');
  const read = await call('sheets.read_range', { tab: 'Summary', range: 'C8' }, null, h);
  assert.equal(read.body.result.values[0][0], 122050);
  assert.equal((await call('sheets.write_range', { values: [[1]] }, null, h)).status, 403);
  assert.equal((await call('sheets.list_sheets', {}, null, h)).status, 403);
  // It can only ever be the shared spreadsheet, whatever it asks for.
  const mine = (await call('sheets.create_sheet', { title: 'Private' })).body.result;
  const other = await call('sheets.read_range', { sheet: mine.id }, null, h);
  assert.equal(other.body.result.sheet, app.exampleList[0].id);
});

test('REST and MCP serve the same tools, with wire names', async () => {
  const s = (await call('sheets_create_sheet', { title: 'Wire names' })).body.result;
  assert.equal(s.title, 'Wire names');
  const client = new Client({ name: 'test-agent', version: '1' });
  await client.connect(new StreamableHTTPClientTransport(new URL(`${base}/mcp`), { requestInit: { headers: { authorization: bearer(sam) } } }));
  const { tools } = await client.listTools();
  assert.ok(tools.every((t) => /^sheets_[a-z_]+$/.test(t.name) && t.name.length <= 64));
  assert.ok(tools.some((t) => t.name === 'sheets_read_range'));
  assert.ok(!tools.some((t) => t.name === 'sheets_sync'));
  const w = await client.callTool({ name: 'sheets_write_range', arguments: { sheet: s.id, values: [['=6*7']] } });
  assert.equal(w.structuredContent.values[0][0], 42);
  const del = await client.callTool({ name: 'sheets_delete_sheet', arguments: { sheet: s.id } });
  assert.ok(del.structuredContent.pending, 'deleting from an agent waits for a person');
  await client.close();
  // A read-only connection sees only reading tools.
  const ro = new Client({ name: 'reader', version: '1' });
  await ro.connect(new StreamableHTTPClientTransport(new URL(`${base}/mcp`), { requestInit: { headers: { authorization: bearer(sam, ['read']) } } }));
  const list = (await ro.listTools()).tools;
  assert.ok(list.length > 5 && list.every((t) => t.annotations.readOnlyHint));
  await ro.close();
});

test('two people edit at once: both changes merge and reach both screens live', async () => {
  const s = (await call('sheets.create_sheet', { title: 'Together' })).body.result;
  const docs = [new Y.Doc(), new Y.Doc()];
  const auths = [bearer(sam), bearer(jordan)];
  for (const [i, d] of docs.entries()) Y.applyUpdate(d, Buffer.from((await call('sheets.sync', { sheet: s.id }, auths[i])).body.result.update, 'base64'));
  const sockets = await Promise.all(auths.map((a, i) => new Promise((ok, bad) => {
    const ws = new WebSocket(`${base.replace('http', 'ws')}/ws`, { headers: { authorization: a } });
    ws.on('message', (m) => { const e = JSON.parse(m); if (e.type === 'sheets.workbook.changed' && e.data.workbook === s.id) Y.applyUpdate(docs[i], Buffer.from(e.data.update, 'base64')); });
    ws.on('open', () => ok(ws));
    ws.on('error', bad);
  })));
  await Promise.all([
    call('sheets.write_range', { sheet: s.id, range: 'A1', values: [['Sam was here', 1]] }, auths[0]),
    call('sheets.write_range', { sheet: s.id, range: 'A2', values: [['Jordan too', 2]] }, auths[1]),
    call('sheets.set_formula', { sheet: s.id, cell: 'B3', formula: '=B1+B2' }, auths[1]),
  ]);
  await new Promise((r) => setTimeout(r, 300));
  for (const d of docs) {
    const cells = d.getMap('cells').get(d.getArray('order').get(0)).toJSON();
    assert.deepEqual(cells, { '0,0': 'Sam was here', '0,1': 1, '1,0': 'Jordan too', '1,1': 2, '2,1': '=B1+B2' });
  }
  assert.equal((await call('sheets.read_range', { sheet: s.id, range: 'B3' })).body.result.values[0][0], 3);
  // Presence reaches the other person.
  const seen = new Promise((ok) => sockets[1].on('message', (m) => { const e = JSON.parse(m); if (e.type === 'sheets.presence.changed') ok(e.data); }));
  await call('sheets.set_presence', { sheet: s.id, cell: 'C4' }, auths[0]);
  assert.equal((await seen).person.name, 'Sam Rivera');
  sockets.forEach((w) => w.close());
});

test('downloads need sign-in and the same team', async () => {
  const s = (await call('sheets.create_sheet', { template: 'budget' })).body.result;
  const x = (await call('sheets.export', { sheet: s.id })).body.result;
  const r = await fetch(`${base}${x.file.url}`, { headers: { authorization: bearer(sam) } });
  assert.equal(r.status, 200);
  assert.equal(r.headers.get('content-type'), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  assert.equal((await fetch(`${base}${x.file.url}`)).status, 401);
});

test('MCP OAuth: register, authorize after sign-in, swap the code for a token', async () => {
  const crypto = await import('node:crypto');
  const reg = await fetch(`${base}/oauth/register`, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ client_name: 'Claude', redirect_uris: ['http://localhost:9999/cb'] }) }).then((r) => r.json());
  const verifier = 'v'.repeat(50);
  const challenge = crypto.createHash('sha256').update(verifier).digest('base64url');
  const q = new URLSearchParams({ client_id: reg.client_id, redirect_uri: 'http://localhost:9999/cb', code_challenge: challenge, code_challenge_method: 'S256', response_type: 'code', state: 's1' });
  const anon = await fetch(`${base}/oauth/authorize?${q}`, { redirect: 'manual' });
  assert.match(anon.headers.get('location'), /^\/login\?next=/);
  const cookie = `sheets_session=${issueTokens(sam).access_token}`;
  const auth = await fetch(`${base}/oauth/authorize?${q}`, { redirect: 'manual', headers: { cookie } });
  const code = new URL(auth.headers.get('location')).searchParams.get('code');
  const tok = await fetch(`${base}/oauth/token`, { method: 'POST', headers: { 'content-type': 'application/x-www-form-urlencoded' }, body: new URLSearchParams({ grant_type: 'authorization_code', code, client_id: reg.client_id, redirect_uri: 'http://localhost:9999/cb', code_verifier: verifier }) }).then((r) => r.json());
  assert.ok(tok.access_token);
  assert.equal((await call('sheets.list_people', {}, `Bearer ${tok.access_token}`)).body.result.me.name, 'Sam Rivera');
});
