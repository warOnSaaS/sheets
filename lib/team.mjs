// Teams and the people on them. Standalone, Sheets keeps its own list; inside the suite the suite's
// member list is the source and people are added here as they arrive.
import { newId, nowIso } from './ids.mjs';
import { BookError } from './book.mjs';

export class Team {
  constructor(db) { this.db = db; }

  async ensure(id, name) {
    await this.db.run('insert into sheets_teams (id, name, created_at) values ($1, $2, $3) on conflict (id) do nothing', [id, name, nowIso()]);
    return this.get(id);
  }

  get(id) { return this.db.get('select * from sheets_teams where id = $1', [id]); }

  async add(teamId, { id = newId('p'), name, email = null, github = null, account_sub = null, role = 'member' }) {
    const e = email ? String(email).trim().toLowerCase() : null;
    if (e && (await this.db.get('select id from sheets_people where team_id = $1 and lower(email) = $2 and deactivated_at is null', [teamId, e]))) throw new BookError(`${e} is already on the team.`, 409, 'conflict');
    if (github && (await this.db.get('select id from sheets_people where team_id = $1 and lower(github) = $2 and deactivated_at is null', [teamId, String(github).toLowerCase()]))) throw new BookError(`@${github} is already on the team.`, 409, 'conflict');
    await this.db.run('insert into sheets_people (id, team_id, name, email, github, account_sub, role, created_at) values ($1, $2, $3, $4, $5, $6, $7, $8) on conflict (id) do nothing',
      [id, teamId, String(name || e || github || 'Someone').slice(0, 100), e, github, account_sub, role, nowIso()]);
    return this.person(id);
  }

  person(id) { return id ? this.db.get('select * from sheets_people where id = $1', [id]) : null; }

  async active(id) { const p = await this.person(id); return p && !p.deactivated_at ? p : null; }

  async people(teamId) {
    const rows = await this.db.all('select * from sheets_people where team_id = $1 and deactivated_at is null order by created_at', [teamId]);
    return rows.map(personView);
  }

  async find(teamId, ref) {
    const s = String(ref ?? '').trim().replace(/^@/, '').toLowerCase();
    const all = await this.db.all('select * from sheets_people where team_id = $1 and deactivated_at is null', [teamId]);
    return all.find((p) => p.id === ref) ?? all.find((p) => p.email?.toLowerCase() === s) ?? all.find((p) => p.github?.toLowerCase() === s) ?? all.find((p) => p.name.toLowerCase() === s) ?? null;
  }

  async remove(teamId, ref) {
    const p = await this.find(teamId, ref);
    if (!p) throw new BookError(`Nobody called ${ref} on the team.`, 404, 'not_found');
    if (p.role === 'owner') throw new BookError('The owner cannot be removed.', 403, 'forbidden');
    await this.db.run('update sheets_people set deactivated_at = $2 where id = $1', [p.id, nowIso()]);
    return personView({ ...p, deactivated_at: nowIso() });
  }

  async prefs(p) { try { return { theme: 'auto', ...(JSON.parse(p.prefs ?? '{}')) }; } catch { return { theme: 'auto' }; } }

  async setPrefs(p, patch) {
    const cur = await this.prefs(p);
    const next = { ...cur, ...Object.fromEntries(Object.entries(patch).filter(([, v]) => v !== undefined)) };
    await this.db.run('update sheets_people set prefs = $2 where id = $1', [p.id, JSON.stringify(next)]);
    return next;
  }

  async setting(teamId, k) { const r = await this.db.get('select value from sheets_settings where team_id = $1 and key = $2', [teamId, k]); return r ? JSON.parse(r.value) : null; }
  async setSetting(teamId, k, v) {
    if (v === null || v === undefined) return this.db.run('delete from sheets_settings where team_id = $1 and key = $2', [teamId, k]);
    const s = JSON.stringify(v);
    await this.db.run('insert into sheets_settings (team_id, key, value) values ($1, $2, $3) on conflict (team_id, key) do update set value = excluded.value', [teamId, k, s]);
  }
}

export const personView = (p) => ({ id: p.id, name: p.name, email: p.email ?? null, github: p.github ?? null, role: p.role, joined_at: p.created_at });
