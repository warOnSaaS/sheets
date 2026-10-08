// wOS Sheets server: pages, the tools (REST at /api/tools/<name> and MCP at /mcp, one set of handlers),
// downloads, sign-in and live editing over WebSockets at /ws. Runs as a long-lived Node server (npm start,
// Docker) and as one Vercel function (api/index.mjs exports the same server).
//   npm run dev                              examples on, SQLite in ./data, email links in the log
//   DATABASE_URL=postgres://... npm start    your team, on any Postgres
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import { WebSocketServer } from 'ws';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js';
import { createApp, VERSION } from './lib/app.mjs';
import { listTools, runTool, toWire, fromWire } from './lib/tools.mjs';
import { canSee } from './lib/bus.mjs';
import { appShell, landing, esc } from './lib/html.mjs';
import { BookError } from './lib/book.mjs';
import { serveFile } from './lib/files.mjs';
import { connectScript, APPS, safeHost } from './lib/connect.mjs';
import { register } from './lib/suite.mjs';
import {
  identify, setCookie, challenge, json, page, bodyObject, loginPage, githubRedirect, accountStart, handleAccountCallback, account,
  handleGithubCallback, handleEmailStart, handleEmailVerify, handleAuthorize, handleToken, handleRegister, resourceMetadata, serverMetadata, COOKIE,
} from './lib/auth.mjs';

const ROOT = path.dirname(new URL(import.meta.url).pathname);
const PUBLIC = path.join(ROOT, 'public');
const TYPES = { '.css': 'text/css; charset=utf-8', '.mjs': 'text/javascript; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.map': 'application/json', '.woff2': 'font/woff2', '.txt': 'text/plain; charset=utf-8', '.svg': 'image/svg+xml', '.png': 'image/png', '.json': 'application/json', '.webmanifest': 'application/manifest+json' };

// What someone with a view-only link may call: only reading, only that one spreadsheet.
const VIEW_TOOLS = new Set(['sheets.sync', 'sheets.read_range', 'sheets.list_charts', 'sheets.get_sheet']);

export const hostOf = (req) => {
  const h = req.headers['x-forwarded-host'] ?? req.headers.host;
  const proto = req.headers['x-forwarded-proto'] ?? (/^(localhost|127\.0\.0\.1|\[::1\])(:|$)/.test(h ?? '') ? 'http' : 'https');
  return `${proto}://${h}`;
};
// On Vercel every request is rewritten to the one function with the original path in ?__p.
const pathOf = (req) => { const u = new URL(req.url, 'http://x'); const p = u.searchParams.get('__p'); if (p) u.searchParams.delete('__p'); return { url: u, p: p ?? u.pathname }; };

export function createServer(appOrPromise) {
  let failedAt = 0;
  const start = (x) => {
    const p = Promise.resolve(x ?? createApp()).then((app) => { attachLive(app, wss); return app; });
    p.catch((e) => { console.error('sheets: could not start:', e); failedAt = Date.now(); });
    return p;
  };
  let ready = start(appOrPromise);
  const getReady = () => {
    if (failedAt && !appOrPromise && Date.now() - failedAt > 5000) { failedAt = 0; ready = start(); server.ready = ready; }
    return ready;
  };
  const server = http.createServer(async (req, res) => {
    let app;
    try { app = await getReady(); } catch { return json(res, 503, { error: { code: 'starting', message: 'Sheets could not start. Check DATABASE_URL and the server log.' } }); }
    try { await route(app, req, res); } catch (e) {
      if (!(e instanceof BookError)) console.error(e);
      if (!res.headersSent) json(res, e.status ?? 500, { error: e instanceof BookError ? { code: e.code, message: e.message } : { code: 'server', message: 'Something went wrong on our side. Try again.' } });
    }
  });
  const wss = new WebSocketServer({ noServer: true, maxPayload: 64 * 1024 });
  server.on('upgrade', async (req, socket, head) => {
    const { p, url } = pathOf(req);
    if (p !== '/ws') return socket.destroy();
    let app;
    try { app = await getReady(); } catch { return socket.destroy(); }
    const share = url.searchParams.get('share');
    const who = await identify(app, { headers: { ...req.headers, ...(share ? { 'x-share-token': share } : {}) } }).catch(() => null);
    if (!who) { socket.write('HTTP/1.1 401 Unauthorized\r\n\r\n'); return socket.destroy(); }
    wss.handleUpgrade(req, socket, head, (ws) => wss.emit('connection', ws, req, who, app));
  });
  server.ready = ready;
  return server;
}

// ---------- live: WebSockets fed by the bus (which Postgres LISTEN/NOTIFY feeds from other copies) ----------

function attachLive(app, wss) {
  if (app.live) return;
  app.live = new Set();
  wss.on('connection', (ws, req, who) => {
    const client = { ws, who };
    app.live.add(client);
    ws.send(JSON.stringify({ type: 'hello', server: app.bus.id }));
    ws.on('message', (raw) => { if (String(raw) === 'ping') ws.send('{"type":"pong"}'); });
    ws.on('close', () => app.live.delete(client));
    ws.on('error', () => app.live.delete(client));
  });
  app.bus.on('event', (e) => {
    for (const c of app.live) {
      const { me, workbook } = c.who;
      if (c.ws.readyState !== 1 || me.team_id !== e.team || !canSee(e, me.id)) continue;
      if (workbook && e.data?.workbook !== workbook) continue;
      if (workbook && !['sheets.workbook.changed', 'sheets.workbook.deleted', 'sheets.workbook.shared'].includes(e.type)) continue;
      const { audience, team, ...out } = e;
      c.ws.send(JSON.stringify(out));
    }
  });
  const timer = setInterval(() => { for (const c of app.live) if (c.ws.readyState === 1) c.ws.ping(); }, 25000);
  timer.unref?.();
}

// ---------- routes ----------

async function route(app, req, res) {
  const { url, p } = pathOf(req);
  const host = hostOf(req);
  app.seenHost(safeHost(host));
  if (req.method === 'GET' && (p.startsWith('/ui/') || p.startsWith('/app/') || ['/manifest.webmanifest', '/icon.svg', '/icon-192.png', '/icon-512.png', '/robots.txt'].includes(p))) return serveStatic(res, p);
  if (p === '/wos-app.json' || p === '/tools.json') return serveFile1(res, path.join(ROOT, p.slice(1)), 'application/json');
  if (p === '/health') return json(res, 200, { ok: true, storage: app.db.kind, auth: app.authProvider, live: app.db.kind === 'postgres' ? 'websockets + postgres notify' : 'websockets', version: VERSION });
  if (p === '/mcp') return handleMcp(app, req, res, host);
  if (p.startsWith('/.well-known/oauth-protected-resource')) return json(res, 200, resourceMetadata(host), { 'access-control-allow-origin': '*' });
  if (p.startsWith('/.well-known/oauth-authorization-server')) return json(res, 200, serverMetadata(host), { 'access-control-allow-origin': '*' });
  if (p === '/oauth/register') return handleRegister(req, res);
  if (p === '/oauth/token') return handleToken(app, req, res);
  if (p === '/oauth/authorize') return handleAuthorize(app, req, res, host);
  if (p === '/oauth/github/callback') return handleGithubCallback(app, req, res, host);
  if (p === '/auth/waronsaas') return accountStart(req, res, host, { next: url.searchParams.get('next') ?? '/' });
  if (p === '/auth/waronsaas/callback') return handleAccountCallback(app, req, res, host);
  if (p === '/login') return app.authProvider === 'waronsaas' ? accountStart(req, res, host, { next: url.searchParams.get('next') ?? '/' }) : loginPage(app, res, url.searchParams.get('next') ?? '/');
  if (p === '/login/github') return process.env.GITHUB_OAUTH_CLIENT_ID ? githubRedirect(res, host, url.searchParams.get('next')) : loginPage(app, res, '/', 'GitHub sign-in is not set up on this server (GITHUB_OAUTH_CLIENT_ID). Use the email link.');
  if (p === '/auth/email' && req.method === 'POST') return handleEmailStart(app, req, res, host);
  if (p === '/auth/email/verify') return handleEmailVerify(app, req, res, host);
  // Sign out here, then out of the warOnSaaS account too (it comes back to the front page).
  if (p === '/logout') return res.writeHead(302, { location: app.authProvider === 'waronsaas' && account() ? account().endSessionUrl(`${host}/`) : '/', 'set-cookie': setCookie(host, COOKIE, '', 0), 'cache-control': 'no-store' }).end();
  if (p === '/api/tools' && req.method === 'GET') return json(res, 200, { tools: catalogue() });
  if (p === '/openapi.json') return json(res, 200, openApi(host), { 'access-control-allow-origin': '*' });
  if (p.startsWith('/api/tools/')) return handleTool(app, req, res, decodeURIComponent(p.slice('/api/tools/'.length)));
  if (p.startsWith('/files/sheets/')) {
    const who = await identify(app, req);
    if (!who || who.via === 'share') return json(res, 401, { error: { code: 'sign_in', message: 'Sign in first.' } });
    return serveFile(app, who.me, res, p);
  }
  if (p === '/connect' || p.startsWith('/connect/')) {
    const a = p.split('/')[2] ?? null;
    const h = safeHost(host);
    if (!h || (a && !APPS.includes(a))) return res.writeHead(404, { 'content-type': 'text/plain; charset=utf-8' }).end('Not found. Try /connect/claude or /connect/codex.\n');
    return res.writeHead(200, { 'content-type': 'text/plain; charset=utf-8', 'cache-control': 'no-store', 'x-content-type-options': 'nosniff' }).end(connectScript({ host: h, app: a, name: 'Sheets' }));
  }
  if (req.method !== 'GET') return json(res, 404, { error: { code: 'not_found', message: 'Not found' } });
  const headers = { 'content-type': 'text/html; charset=utf-8', 'cache-control': 'no-store' };
  // A view-only link: the same screen, read-only, for that one spreadsheet. No account needed.
  if (p.startsWith('/v/')) {
    const row = await app.store.byShareToken(p.slice(3));
    if (!row) return page(res, 404, '<h1>This link does not work any more</h1><p>Whoever shared it turned it off, or the spreadsheet was deleted.</p><a class="ui-btn is-quiet" href="/">Go to Sheets</a>');
    return res.writeHead(200, { ...headers, 'x-robots-tag': 'noindex' }).end(appShell({ title: `${row.title} · Sheets`, version: VERSION, view: { token: p.slice(3), sheet: row.id, title: row.title }, account: app.authProvider === 'waronsaas', me: (await identify(app, req))?.me ? { signedIn: true } : null }));
  }
  const who = await identify(app, req);
  if (!who) {
    if (p === '/' || p === '') return res.writeHead(200, headers).end(landing({ examples: app.exampleList, version: VERSION, openSignup: app.openSignup, account: app.authProvider === 'waronsaas' }));
    return loginPage(app, res, p + url.search);
  }
  const prefs = await app.team.prefs(who.me);
  res.writeHead(200, { ...headers, 'x-robots-tag': 'noindex' }).end(appShell({ title: 'Sheets', theme: prefs.theme, version: VERSION, me: { id: who.me.id, name: who.me.name, role: who.me.role }, account: app.authProvider === 'waronsaas' }));
}

export function catalogue() {
  return listTools().map((t) => ({ name: t.name, title: t.title, description: t.description, input: t.inputJson, output: t.outputJson, scope: t.scope, confirm: t.confirm, emits: t.emits }));
}

async function handleTool(app, req, res, name) {
  if (req.method !== 'POST') return json(res, 405, { error: { code: 'method', message: 'Use POST' } }, { allow: 'POST' });
  const who = await identify(app, req);
  if (!who) return json(res, 401, { error: { code: 'sign_in', message: 'Sign in first.' } }, { 'www-authenticate': challenge(hostOf(req)) });
  // A browser call must come from this site (the cookie is SameSite=Lax; this closes the rest).
  if (who.via === 'web' && req.headers.origin && req.headers.origin !== hostOf(req)) return json(res, 403, { error: { code: 'forbidden', message: 'Wrong origin.' } });
  const input = await bodyObject(req);
  const tool = fromWire(name);
  if (who.via === 'share') {
    if (!VIEW_TOOLS.has(tool)) return json(res, 403, { error: { code: 'sign_in', message: 'This is a view-only link. Sign in to edit.' } });
    input.sheet = who.workbook;
  }
  const result = await runTool(app, who.me, tool, input, { via: who.via === 'mcp' ? 'rest' : who.via === 'share' ? 'web' : 'web', scopes: who.scopes, client: who.client });
  if (result?.pending) return json(res, 202, result);
  json(res, 200, { result });
}

async function handleMcp(app, req, res, host) {
  if (req.method !== 'POST') return json(res, 405, { error: 'Use POST (this is an MCP endpoint)' }, { allow: 'POST' });
  const who = await identify(app, { headers: { authorization: req.headers.authorization } });
  if (!who) return json(res, 401, { error: 'Sign in to use Sheets.' }, { 'www-authenticate': challenge(host) });
  const server = new McpServer({ name: 'wos-sheets', version: VERSION }, { instructions: INSTRUCTIONS });
  for (const t of listTools()) {
    if (!who.scopes.includes(t.scope) || t.name === 'sheets.sync' || t.name === 'sheets.set_presence') continue;
    server.registerTool(toWire(t.name), {
      title: t.title,
      description: t.description + (t.confirm === 'human' ? ' Needs a person\'s yes: it asks them in the app first.' : ''),
      inputSchema: t.input,
      annotations: { readOnlyHint: t.scope === 'read', destructiveHint: t.scope === 'delete' || t.confirm === 'human', openWorldHint: t.name === 'sheets.import_from_crm' || t.name === 'sheets.refresh_link' },
    }, async (args) => {
      try {
        const out = await runTool(app, who.me, t.name, args ?? {}, { via: 'mcp', scopes: who.scopes, client: who.client });
        return { content: [{ type: 'text', text: JSON.stringify(out) }], structuredContent: out };
      } catch (e) {
        return { isError: true, content: [{ type: 'text', text: e.message }] };
      }
    });
  }
  const transport = new StreamableHTTPServerTransport({ sessionIdGenerator: undefined, enableJsonResponse: true });
  res.on('close', () => { transport.close(); server.close(); });
  await server.connect(transport);
  await transport.handleRequest(req, res, await bodyObject(req));
}

const INSTRUCTIONS = `This is wOS Sheets, the team's spreadsheets. Values you read are exact: sheets_read_range gives each cell's value (numbers stay numbers, errors like #DIV/0! are strings) and its formula.
Start with sheets_list_sheets, then sheets_get_sheet to see the tabs and the range in use. Write with sheets_write_range (cells are read as typed: "$1,200", "12%", "2026-10-07", "=SUM(B2:B9)") or sheets_set_formula with fill to copy a formula down a column. Check the values and errors that writes return.
Build summaries with formulas, sheets_create_pivot (live SUMIFS formulas) and sheets_create_chart. Pull CRM records with sheets_import_from_crm and keep them fresh with sheets_refresh_link.
Never delete or share without being asked. Never invent numbers: read them.`;

function openApi(host) {
  const paths = {};
  for (const t of listTools()) {
    if (['sheets.sync', 'sheets.set_presence', 'sheets.decide_approval'].includes(t.name)) continue;
    paths[`/api/tools/${toWire(t.name)}`] = { post: { operationId: toWire(t.name), summary: t.title, description: t.description.slice(0, 300), 'x-openai-isConsequential': t.scope === 'delete' || t.confirm === 'human', requestBody: { required: true, content: { 'application/json': { schema: t.inputJson } } }, responses: { 200: { description: 'Done', content: { 'application/json': { schema: { type: 'object', properties: { result: t.outputJson } } } } } } } };
  }
  return { openapi: '3.1.0', info: { title: 'wOS Sheets', version: VERSION, description: 'The team\'s spreadsheets: read exact values and formulas, write cells, charts, pivots.' }, servers: [{ url: host }], paths, components: { securitySchemes: { oauth: { type: 'oauth2', flows: { authorizationCode: { authorizationUrl: `${host}/oauth/authorize`, tokenUrl: `${host}/oauth/token`, scopes: { read: 'Read', write: 'Write', delete: 'Delete', admin: 'Admin' } } } } } }, security: [{ oauth: ['read', 'write'] }] };
}

function serveStatic(res, p) {
  const file = path.join(PUBLIC, path.normalize(p).replace(/^(\.\.[/\\])+/, ''));
  if (!file.startsWith(PUBLIC)) return res.writeHead(404).end();
  return serveFile1(res, file);
}

function serveFile1(res, file, type) {
  if (!fs.existsSync(file) || fs.statSync(file).isDirectory()) return res.writeHead(404).end();
  res.writeHead(200, { 'content-type': type ?? TYPES[path.extname(file)] ?? 'application/octet-stream', 'cache-control': 'no-cache' }).end(fs.readFileSync(file));
}

if (import.meta.url === `file://${process.argv[1]}`) {
  const server = createServer();
  const port = Number(process.env.PORT || 3994);
  server.listen(port, async () => {
    const app = await server.ready;
    console.log(`wOS Sheets on http://localhost:${port} (${app.db.kind}, sign-in: ${app.authProvider}${app.openSignup ? ', open sign-up' : ''}). Agents connect to /mcp.`);
  });
  const stop = async () => { server.close(); (await server.ready).close?.(); process.exit(0); };
  process.on('SIGTERM', stop);
  process.on('SIGINT', stop);
}

export { esc };

// In the suite, this file is the app's server part: register(ctx) returns the tool handlers (lib/suite.mjs).
export default register;
