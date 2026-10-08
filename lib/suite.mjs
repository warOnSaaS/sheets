// Sheets inside the wOS suite (warOnSaaS/suite CONTRACTS.md). The suite calls register(ctx) once, runs
// migrations/ itself, checks sign-in, team, scope and input before any handler, gates confirm: human tools,
// and serves /api/tools/<name> and /mcp from these handlers. The Sheets code is the same as standalone;
// this file only translates the suite's ctx and call into it.
import { EventEmitter } from 'node:events';
import { Store } from './store.mjs';
import { Team } from './team.mjs';
import { Files, serveFile } from './files.mjs';
import { listTools, runTool } from './tools.mjs';
import { seal, unseal } from './seal.mjs';
import { VERSION } from './app.mjs';
import { json } from './auth.mjs';

// The suite's db speaks ? placeholders; Sheets writes $1, $2. Convert, repeating values used twice.
export function fromSuiteDb(sdb) {
  const conv = (sql, params = []) => {
    const out = [];
    const text = sql.replace(/\$(\d+)/g, (_, n) => { out.push(params[Number(n) - 1]); return '?'; });
    return [text, /\$\d/.test(sql) ? out : params];
  };
  const wrap = (q, dialect) => ({
    dialect,
    async all(sql, params) { const [t, p] = conv(sql, params); return q.query(t, p); },
    async get(sql, params) { const [t, p] = conv(sql, params); return (await q.get(t, p)) ?? null; },
    async run(sql, params) {
      const [t, p] = conv(sql, params);
      if (/\breturning\b/i.test(t)) { const rows = await q.query(t, p); return { changes: rows.length, rows }; }
      const r = await q.run(t, p);
      return { changes: r?.changes ?? 0, rows: [] };
    },
  });
  const dialect = sdb.dialect === 'postgres' ? 'pg' : 'sqlite';
  return {
    ...wrap(sdb, dialect),
    kind: sdb.dialect === 'postgres' ? 'postgres' : 'sqlite',
    tx: (fn) => sdb.tx((t) => fn(wrap(t, dialect))),
    async notify() {}, async listen() { return async () => {}; }, async close() {},
  };
}

// The suite carries events to every member's screen and between server copies.
class SuiteBus extends EventEmitter {
  constructor(events) { super(); this.events = events; this.id = 'suite'; }
  publish({ team, type, data = {}, audience = 'team' }) {
    this.emit('event', { team, type, data, audience });
    try { this.events?.publish(team, type, data); } catch {}
  }
}

export async function register(ctx) {
  const env = (name) => ctx.env?.(name) ?? undefined;
  const envObj = new Proxy({}, { get: (_, k) => (typeof k === 'string' ? env(k) : undefined) });
  const db = fromSuiteDb(ctx.db);
  const bus = new SuiteBus(ctx.events);
  const secret = env('OAUTH_SECRET') || env('WOS_SECRET_KEY') || 'suite';
  const base = String(ctx.publicUrl ?? '').replace(/\/$/, '');
  const app = {
    env: envObj, db, bus, version: VERSION, suite: true,
    store: new Store({ db, bus }), team: new Team(db), files: new Files(db, { prefix: '/files/sheets' }),
    authProvider: 'suite', openSignup: false, demo: false,
    seal: (s) => seal(s, secret), unseal: (s) => unseal(s, secret),
    publicUrl: () => base,
    urlFor: (id) => `${base}/a/sheets/s/${id}`,
    // View-only links are served by the standalone server for now; inside the suite they stay off.
    shareUrl: () => null,
  };

  // The suite's member list is the source of truth; people are copied in as they act, refreshed each minute.
  const synced = new Map();
  async function personFor(call) {
    const team = call.team;
    if (Date.now() - (synced.get(team.id) ?? 0) > 60000) {
      synced.set(team.id, Date.now());
      await app.team.ensure(team.id, team.name);
      let members = [];
      try { members = (await ctx.people?.members(team.id)) ?? []; } catch {}
      for (const m of members) {
        const role = m.role === 'owner' ? 'owner' : m.role === 'admin' ? 'admin' : m.role === 'guest' ? 'member' : 'member';
        const have = await app.team.person(m.id);
        if (!have) await db.run('insert into sheets_people (id, team_id, name, email, github, role, created_at) values ($1, $2, $3, $4, $5, $6, $7) on conflict (id) do nothing', [m.id, team.id, m.name || m.email || 'Someone', m.email ?? null, m.github_login ?? null, role, new Date().toISOString()]);
        else await db.run('update sheets_people set role = $2, name = $3, deactivated_at = null where id = $1', [m.id, role, m.name || have.name]);
      }
    }
    const id = call.actor.kind === 'agent' && call.actor.personId ? call.actor.personId : call.actor.id;
    let p = await app.team.person(id);
    if (!p) {
      await db.run('insert into sheets_people (id, team_id, name, role, created_at) values ($1, $2, $3, $4, $5) on conflict (id) do nothing', [id, call.team.id, call.actor.name || 'Someone', call.scopes.includes('admin') ? 'admin' : 'member', new Date().toISOString()]);
      p = await app.team.person(id);
    }
    return { ...p, team_id: call.team.id };
  }

  const VIA = { screen: 'web', rest: 'rest', mcp: 'mcp', agent: 'mcp', email: 'rest', system: 'rest' };
  const handlers = {};
  for (const t of listTools()) {
    // The suite asked for the person's yes before calling a confirm: human tool, so it runs straight away.
    handlers[t.name] = async (input, call) => runTool(app, await personFor(call), t.name, input ?? {}, { via: VIA[call.via] ?? 'rest', scopes: call.scopes, approved: true, client: call.actor.kind === 'agent' ? call.actor.name : null, call });
  }

  return {
    handlers,
    async routes(req, res, url, call) {
      if (!url.pathname.startsWith('/files/sheets/')) return false;
      if (!call) { json(res, 401, { error: { code: 'sign_in', message: 'Sign in first.' } }); return true; }
      await serveFile(app, await personFor(call), res, url.pathname);
      return true;
    },
    async exportTeam(team) {
      const rows = (sql) => db.all(sql, [team.id]);
      return {
        workbooks: await rows('select id, title, created_by, created_at, updated_at, state, state_seq, share_mode from sheets_workbooks where team_id = $1 and deleted_at is null'),
        updates: await rows('select * from sheets_updates where team_id = $1 order by id'),
        versions: await rows('select id, workbook_id, label, auto, created_by, created_at from sheets_versions where team_id = $1'),
        people: await rows('select * from sheets_people where team_id = $1'),
      };
    },
  };
}
