import crypto from 'node:crypto';
import { esc } from './html.mjs';
import { newId } from './ids.mjs';
import { WosAccount } from './account-client.mjs';

// Sign-in. AUTH_PROVIDER picks how people sign in:
//   waronsaas  the shared warOnSaaS account (WOS_ACCOUNT_CLIENT_ID and WOS_ACCOUNT_CLIENT_SECRET; WOS_ACCOUNT_URL to self-host it)
//   github     GitHub, plus an email link
//   local      an email link only
// Agents (Claude, ChatGPT, Claude Code, Codex) connect to /mcp with standard MCP OAuth (discovery, dynamic
// client registration, PKCE); the person signs in the same way. Tokens are signed and stateless; who is on
// the team comes from the database, so removing someone signs them out everywhere. Ported from wOS Chat.

const SECRET = () => process.env.OAUTH_SECRET || process.env.SESSION_SECRET || 'dev-secret-change-me';
const GH_WEB = () => process.env.GITHUB_WEB_BASE || 'https://github.com';
const GH_API = () => process.env.GITHUB_API_BASE || 'https://api.github.com';
const now = () => Math.floor(Date.now() / 1000);
const DAY = 24 * 3600;
export const COOKIE = 'sheets_session';
export const ALL_SCOPES = ['read', 'write', 'delete', 'admin'];

export function authProvider(env = process.env) {
  const p = String(env.AUTH_PROVIDER ?? '').toLowerCase();
  if (['waronsaas', 'github', 'local'].includes(p)) return p;
  if (env.WOS_ACCOUNT_CLIENT_ID && env.WOS_ACCOUNT_CLIENT_SECRET) return 'waronsaas';
  return env.GITHUB_OAUTH_CLIENT_ID ? 'github' : 'local';
}

export function sign(payload) {
  const body = Buffer.from(JSON.stringify(payload)).toString('base64url');
  return `${body}.${crypto.createHmac('sha256', SECRET()).update(body).digest('base64url')}`;
}

export function verify(token, kind) {
  const [body, mac] = String(token ?? '').split('.');
  if (!body || !mac) return null;
  const want = crypto.createHmac('sha256', SECRET()).update(body).digest('base64url');
  if (want.length !== mac.length || !crypto.timingSafeEqual(Buffer.from(want), Buffer.from(mac))) return null;
  let p;
  try { p = JSON.parse(Buffer.from(body, 'base64url').toString()); } catch { return null; }
  if (p.k !== kind || (p.exp && p.exp < now())) return null;
  return p;
}

export const cookieOf = (req, name) => {
  const v = new RegExp(`(?:^|;\\s*)${name}=([^;]+)`).exec(req.headers.cookie ?? '')?.[1];
  return v ? decodeURIComponent(v) : null;
};
const secureFlag = (host) => (host.startsWith('https://') ? '; Secure' : '');
export const setCookie = (host, name, value, maxAge) => `${name}=${encodeURIComponent(value)}; Path=/; HttpOnly${secureFlag(host)}; SameSite=Lax; Max-Age=${maxAge}`;

// Who is asking, and how: { me, via, scopes, client }, or a viewer of one shared spreadsheet.
export async function identify(app, req) {
  const bearer = (req.headers.authorization ?? '').replace(/^Bearer\s+/i, '');
  if (bearer) {
    const t = verify(bearer, 'access');
    const me = t && (await app.team.active(t.id)) && (await live(t)) ? await app.team.active(t.id) : null;
    return me ? { me, via: 'mcp', scopes: t.sc ?? ALL_SCOPES, client: t.a ?? null } : null;
  }
  const s = verify(cookieOf(req, COOKIE), 'access');
  const me = s && (await live(s)) ? await app.team.active(s.id) : null;
  if (me) return { me, via: 'web', scopes: ALL_SCOPES, client: null };
  const share = req.headers['x-share-token'];
  if (share) {
    const row = await app.store.byShareToken(String(share));
    if (row) return { me: { id: `viewer_${row.id}`, team_id: row.team_id, role: 'viewer', name: 'Guest' }, via: 'share', scopes: ['read'], client: null, workbook: row.id };
  }
  return null;
}

// A token from an account sign-in carries the account's session id; sign out everywhere ends it here too.
const live = async (t) => !t.sid || !account() || account().isLive(t.sid);

export function issueTokens(person, { scopes = ALL_SCOPES, app: client, sid = null } = {}) {
  const base = { id: person.id, sc: scopes, ...(client ? { a: String(client).slice(0, 60) } : {}), ...(sid ? { sid } : {}) };
  return {
    access_token: sign({ k: 'access', ...base, exp: now() + 30 * DAY }),
    refresh_token: sign({ k: 'refresh', ...base, exp: now() + 365 * DAY }),
    token_type: 'bearer', expires_in: 30 * DAY, scope: scopes.join(' '),
  };
}

// ---------- MCP OAuth discovery ----------

export const resourceMetadata = (host) => ({ resource: `${host}/mcp`, authorization_servers: [host], bearer_methods_supported: ['header'], resource_name: 'wOS Sheets', scopes_supported: ALL_SCOPES });
export const serverMetadata = (host) => ({
  issuer: host, authorization_endpoint: `${host}/oauth/authorize`, token_endpoint: `${host}/oauth/token`, registration_endpoint: `${host}/oauth/register`,
  response_types_supported: ['code'], grant_types_supported: ['authorization_code', 'refresh_token'], code_challenge_methods_supported: ['S256'],
  token_endpoint_auth_methods_supported: ['none'], scopes_supported: ALL_SCOPES,
});
export const challenge = (host) => `Bearer resource_metadata="${host}/.well-known/oauth-protected-resource"`;

export async function handleRegister(req, res) {
  const b = await bodyObject(req);
  const uris = Array.isArray(b.redirect_uris) ? b.redirect_uris.filter(okRedirect) : [];
  if (!uris.length) return json(res, 400, { error: 'invalid_redirect_uri' });
  const client_id = sign({ k: 'client', r: uris, n: String(b.client_name ?? '').slice(0, 80) });
  json(res, 201, { client_id, client_name: b.client_name, redirect_uris: uris, grant_types: ['authorization_code', 'refresh_token'], response_types: ['code'], token_endpoint_auth_method: 'none', client_id_issued_at: now() });
}

function okRedirect(uri) {
  try {
    const u = new URL(uri);
    return u.protocol === 'https:' || (u.protocol === 'http:' && ['localhost', '127.0.0.1', '[::1]'].includes(u.hostname));
  } catch { return false; }
}

const scopesFrom = (s) => { const want = String(s ?? '').split(/[\s,]+/).filter((x) => ALL_SCOPES.includes(x)); return want.length ? want : ALL_SCOPES; };

export async function handleAuthorize(app, req, res, host) {
  const url = new URL(req.url, 'http://x');
  const q = Object.fromEntries(url.searchParams);
  const c = verify(q.client_id, 'client');
  if (!c || !c.r.includes(q.redirect_uri)) return page(res, 400, '<h1>This sign-in link is not valid</h1><p>Start again from your app.</p>');
  if (!q.code_challenge || (q.code_challenge_method ?? 'S256') !== 'S256') return page(res, 400, '<h1>This app must use PKCE</h1>');
  const who = await identify(app, { headers: { cookie: req.headers.cookie } });
  if (!who || who.via !== 'web') {
    const next = `/oauth/authorize?${url.searchParams}`.replace(/[?&]__p=[^&]*/, '');
    if (app.authProvider === 'waronsaas') return accountStart(req, res, host, { next, connection: `${c.n || 'An AI app'} via Sheets` });
    return res.writeHead(302, { location: `/login?next=${encodeURIComponent(next)}`, 'cache-control': 'no-store' }).end();
  }
  const sess = verify(cookieOf(req, COOKIE), 'access');
  const code = sign({ k: 'code', id: who.me.id, c: q.client_id, r: q.redirect_uri, cc: q.code_challenge, sc: scopesFrom(q.scope), a: c.n, sid: sess?.sid ?? null, exp: now() + 300 });
  const to = new URL(q.redirect_uri);
  to.searchParams.set('code', code);
  if (q.state) to.searchParams.set('state', q.state);
  res.writeHead(302, { location: to.toString(), 'cache-control': 'no-store' }).end();
}

export async function handleToken(app, req, res) {
  const q = await bodyObject(req);
  const client = verify(q.client_id, 'client');
  if (!client) return json(res, 401, { error: 'invalid_client' });
  let grant = null;
  if (q.grant_type === 'authorization_code') {
    grant = verify(q.code, 'code');
    if (grant && (grant.c !== q.client_id || (q.redirect_uri && grant.r !== q.redirect_uri))) grant = null;
    if (grant && crypto.createHash('sha256').update(String(q.code_verifier ?? '')).digest('base64url') !== grant.cc) grant = null;
  } else if (q.grant_type === 'refresh_token') {
    grant = verify(q.refresh_token, 'refresh');
  }
  const me = grant && (await app.team.active(grant.id));
  if (!me) return json(res, 400, { error: 'invalid_grant' });
  if (!(await live(grant))) return json(res, 400, { error: 'invalid_grant' });
  json(res, 200, issueTokens(me, { scopes: grant.sc ?? ALL_SCOPES, app: grant.a, sid: grant.sid ?? null }));
}

// ---------- browser sign-in ----------

const safeNext = (n) => (n && n.startsWith('/') && !n.startsWith('//') ? n : '/');

export function loginPage(app, res, next = '/', note = '') {
  const p = app.authProvider;
  const gh = p === 'github' && !!process.env.GITHUB_OAUTH_CLIENT_ID;
  const acct = p === 'waronsaas';
  const mail = p !== 'waronsaas';
  page(res, 200, `<a class="gate-back" href="/">Back to the examples</a><span class="gate-mark" aria-hidden="true">${MARK}</span><h1>Sign in to Sheets</h1><p>Spreadsheets your team owns, that your own AI can work in.</p>${note ? `<p class="gate-note">${note}</p>` : ''}
  ${acct ? `<a class="ui-btn is-accent is-block is-lg" data-auth href="/auth/waronsaas?next=${encodeURIComponent(next)}">Sign in with your warOnSaaS account</a><p class="gate-small">GitHub, Google or an email link. One free account for every warOnSaaS app.</p>` : ''}
  ${gh ? `<a class="ui-btn is-accent is-block is-lg" data-auth href="/login/github?next=${encodeURIComponent(next)}">Sign in with GitHub</a>` : ''}
  ${mail ? `<form method="post" action="/auth/email" class="gate-form" data-auth><input type="hidden" name="next" value="${esc(next)}"><label class="ui-field"><span>${gh ? 'Or get a sign-in link by email' : 'Get a sign-in link by email'}</span><input class="ui-input" type="email" name="email" required placeholder="you@company.example" autocomplete="email"></label><button class="ui-btn ${gh ? 'is-quiet' : 'is-accent'} is-block" type="submit">Email me a link</button></form>` : ''}
  <p class="gate-small">${app.openSignup ? 'New here? Signing in makes your own workspace with the examples in it. Free.' : 'Not on the team yet? Ask whoever runs this server to add you.'}</p>`);
}

export const MARK = '<svg viewBox="0 0 24 24" width="28" height="28" aria-hidden="true"><rect x="3" y="3" width="18" height="18" rx="4" fill="none" stroke="currentColor" stroke-width="1.8"/><path d="M3 9h18M3 15h18M9 3v18" stroke="currentColor" stroke-width="1.8"/></svg>';

export function githubRedirect(res, host, next) {
  const state = sign({ k: 'gh', web: safeNext(next), exp: now() + 900 });
  const u = new URL(`${GH_WEB()}/login/oauth/authorize`);
  u.searchParams.set('client_id', process.env.GITHUB_OAUTH_CLIENT_ID ?? '');
  u.searchParams.set('redirect_uri', `${host}/oauth/github/callback`);
  u.searchParams.set('scope', 'read:user user:email');
  u.searchParams.set('state', state);
  res.writeHead(302, { location: u.toString(), 'cache-control': 'no-store' }).end();
}

export async function handleGithubCallback(app, req, res, host) {
  const q = Object.fromEntries(new URL(req.url, 'http://x').searchParams);
  const st = verify(q.state, 'gh');
  if (!st || !q.code) return page(res, 400, '<h1>Sign-in expired</h1><p><a href="/login">Try again</a>.</p>');
  const tok = await fetch(`${GH_WEB()}/login/oauth/access_token`, {
    method: 'POST', headers: { accept: 'application/json', 'content-type': 'application/json' },
    body: JSON.stringify({ client_id: process.env.GITHUB_OAUTH_CLIENT_ID, client_secret: process.env.GITHUB_OAUTH_CLIENT_SECRET, code: q.code, redirect_uri: `${host}/oauth/github/callback` }),
  }).then((r) => r.json()).catch(() => ({}));
  if (!tok.access_token) return page(res, 400, '<h1>GitHub sign-in failed</h1><p><a href="/login">Try again</a>.</p>');
  const gh = (p) => fetch(`${GH_API()}${p}`, { headers: { authorization: `Bearer ${tok.access_token}`, accept: 'application/vnd.github+json', 'user-agent': 'wos-sheets' } }).then((r) => (r.ok ? r.json() : null));
  const [user, emails] = await Promise.all([gh('/user'), gh('/user/emails')]);
  if (!user?.login) return page(res, 400, '<h1>GitHub sign-in failed</h1><p><a href="/login">Try again</a>.</p>');
  const verified = (emails ?? []).filter((e) => e.verified).map((e) => e.email.toLowerCase());
  const me = await personForSignIn(app, { github: user.login, emails: verified, name: user.name || user.login });
  if (!me) return page(res, 403, `<h1>Hi @${esc(user.login)}</h1><p>You're signed in to GitHub, but you're not on this team yet. Ask a team admin to add <b>${esc(user.login)}</b>, then sign in again.</p>`);
  signedIn(res, host, me, st.web);
}

// The warOnSaaS account (account.waronsaas.com), through its client library (lib/account-client.mjs, copied
// from warOnSaaS/account). The account's session id rides in our own session token, and isLive() makes
// "sign out everywhere" reach Sheets within a minute.
let accountClient;
export function account(host) {
  if (accountClient === undefined) accountClient = WosAccount.fromEnv(process.env, { secret: SECRET() });
  return accountClient;
}

export function accountStart(req, res, host, { next = '/', connection = null } = {}) {
  const a = account();
  if (!a) return page(res, 503, '<h1>The warOnSaaS account is not set up here</h1><p>This server needs WOS_ACCOUNT_CLIENT_ID and WOS_ACCOUNT_CLIENT_SECRET.</p>');
  const q = new URL(req.url, 'http://x').searchParams;
  const { location, cookie } = a.start({ next: safeNext(next), connection, prompt: q.get('prompt') || null, provider: q.get('provider') || null, redirectUri: `${host}/auth/waronsaas/callback`, secure: host.startsWith('https://') });
  res.writeHead(302, { location, 'set-cookie': cookie, 'cache-control': 'no-store' }).end();
}

export async function handleAccountCallback(app, req, res, host) {
  const a = account();
  if (!a) return page(res, 503, '<h1>The warOnSaaS account is not set up here</h1>');
  const r = await a.finish(req);
  // Cancelled, or a silent try that found nobody: back to the page, still signed out, still open.
  if (r.error) return res.writeHead(302, { location: safeNext(r.next), 'set-cookie': r.clear, 'cache-control': 'no-store' }).end();
  const p = r.profile;
  const emails = p.email && p.email_verified !== false ? [String(p.email).toLowerCase()] : [];
  const me = await personForSignIn(app, { account_sub: p.sub, github: p.github_login ?? null, emails, name: p.name || p.email?.split('@')[0] || 'Someone' });
  if (!me) return page(res, 403, `<h1>Not on this team yet</h1><p>Ask a team admin to add ${esc(p.email ?? 'you')}, then sign in again.</p>`);
  signedIn(res, host, me, r.next, [r.clear], p.sid);
}

// Email link: always the same answer, so the form never tells a stranger who is on the team.
export async function handleEmailStart(app, req, res, host) {
  const b = await bodyObject(req);
  const email = String(b.email ?? '').trim().toLowerCase();
  const next = safeNext(b.next);
  if (app.authProvider === 'waronsaas') return loginPage(app, res, next);
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return loginPage(app, res, next, 'That email address does not look right.');
  const known = await app.db.get('select id from sheets_people where lower(email) = $1 and deactivated_at is null', [email]);
  const empty = !(await app.db.get('select id from sheets_people where team_id = $1', [app.teamId]));
  if (known || empty || app.openSignup) {
    const t = sign({ k: 'email', e: email, web: next, exp: now() + 900 });
    const link = `${host}/auth/email/verify?t=${encodeURIComponent(t)}`;
    await app.mailer.send({ to: email, subject: 'Your sign-in link for Sheets', text: `Here is your sign-in link for wOS Sheets. It works for 15 minutes:\n\n${link}\n\nIf you did not ask for it, ignore this email.` });
  }
  page(res, 200, `<span class="gate-mark" aria-hidden="true">${MARK}</span><h1>Check your email</h1><p>If ${esc(email)} can sign in here, a link is on its way. It works for 15 minutes.</p>${app.mailer.ready ? '' : '<p class="gate-note">This server has no email set up (SMTP_URL or RESEND_API_KEY), so the link was written to the server log.</p>'}`);
}

export async function handleEmailVerify(app, req, res, host) {
  const t = verify(new URL(req.url, 'http://x').searchParams.get('t'), 'email');
  if (!t) return page(res, 400, '<h1>That link has expired</h1><p><a href="/login">Get a new one</a>.</p>');
  const me = await personForSignIn(app, { emails: [t.e], name: t.e.split('@')[0] });
  if (!me) return page(res, 403, '<h1>Not on this team</h1><p>Ask a team admin to add your email.</p>');
  signedIn(res, host, me, t.web);
}

function signedIn(res, host, me, next, extra = [], sid = null) {
  const t = issueTokens(me, { sid });
  res.writeHead(302, { location: safeNext(next), 'set-cookie': [setCookie(host, COOKIE, t.access_token, 30 * DAY), ...extra], 'cache-control': 'no-store' }).end();
}

// Someone already on a team, matched by account, GitHub login or a verified email. Otherwise: on a brand-new
// server the first person becomes the owner; with open sign-up (the hosted demo) everyone gets a workspace.
async function personForSignIn(app, { github = null, account_sub = null, emails = [], name }) {
  const db = app.db;
  const live = 'deactivated_at is null';
  if (account_sub) { const p = await db.get(`select * from sheets_people where account_sub = $1 and ${live}`, [account_sub]); if (p) return p; }
  if (github) { const p = await db.get(`select * from sheets_people where lower(github) = $1 and ${live}`, [github.toLowerCase()]); if (p) return p; }
  for (const e of emails) {
    const p = await db.get(`select * from sheets_people where lower(email) = $1 and ${live} order by created_at`, [e]);
    if (p) {
      if (github && !p.github) await db.run('update sheets_people set github = $2 where id = $1', [p.id, github]);
      if (account_sub && !p.account_sub) await db.run('update sheets_people set account_sub = $2 where id = $1', [p.id, account_sub]);
      return app.team.person(p.id);
    }
  }
  const first = !(await db.get('select id from sheets_people where team_id = $1', [app.teamId]));
  if (first && !app.openSignup) {
    await app.team.ensure(app.teamId, app.teamName);
    return app.team.add(app.teamId, { name, email: emails[0] ?? null, github, account_sub, role: 'owner' });
  }
  if (app.openSignup) {
    const teamId = newId('team');
    await app.team.ensure(teamId, `${name}'s sheets`);
    const p = await app.team.add(teamId, { name, email: emails[0] ?? null, github, account_sub, role: 'owner' });
    await app.seedWorkspace?.(p);
    return p;
  }
  return null;
}

// ---------- bits ----------

export async function bodyObject(req, limit = 40e6) {
  if (req.body && typeof req.body === 'object' && !Buffer.isBuffer(req.body)) return req.body;
  let s = typeof req.body === 'string' ? req.body : Buffer.isBuffer(req.body) ? req.body.toString('utf8') : '';
  if (!s) { const parts = []; let n = 0; for await (const c of req) { parts.push(c); n += c.length; if (n > limit) break; } s = Buffer.concat(parts).toString('utf8'); }
  if (!s) return {};
  try { return (req.headers['content-type'] ?? '').includes('json') ? JSON.parse(s) : Object.fromEntries(new URLSearchParams(s)); } catch { return {}; }
}

export const json = (res, status, obj, headers = {}) => res.writeHead(status, { 'content-type': 'application/json', 'cache-control': 'no-store', ...headers }).end(JSON.stringify(obj));

export function page(res, status, inner) {
  res.writeHead(status, { 'content-type': 'text/html; charset=utf-8', 'cache-control': 'no-store', 'x-robots-tag': 'noindex' }).end(`<!doctype html><html lang="en" data-scheme="ops" data-mode="auto" data-shape="soft" data-type="grotesk" data-surface="bordered"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Sign in · Sheets</title><meta name="robots" content="noindex"><link rel="icon" href="/icon.svg" type="image/svg+xml">
<link rel="stylesheet" href="/ui/src/ui.css"><link rel="stylesheet" href="/ui/src/tokens.css"><link rel="stylesheet" href="/app/sheets.css">
</head><body class="gate"><main class="gate-card">${inner}</main></body></html>`);
}
