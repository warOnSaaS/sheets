// Workbooks in the database. Each workbook is a Yjs document: a folded state in sheets_workbooks.state,
// plus the changes since then in sheets_updates. Any number of server copies can add changes at once;
// Yjs merges them. A copy keeps recently used documents in memory and reads only the changes it has not
// seen. Every change is also sent out live (the bus), and a version is saved now and then for history.
import crypto from 'node:crypto';
import { Y, initBook, title as bookTitle, toB64, fromB64, BookError } from './book.mjs';
import { Engine } from './engine.mjs';
import { newId, nowIso } from './ids.mjs';

const FOLD_AFTER = 150;          // fold changes into the state after this many
const AUTO_VERSION_MS = 10 * 60 * 1000;
const KEEP_AUTO_VERSIONS = 200;
const CACHE = 40;

export class Store {
  constructor({ db, bus }) {
    this.db = db;
    this.bus = bus;
    this.cache = new Map();
    this.locks = new Map();
  }

  // Changes to one workbook in this server copy run one after another (other copies merge through Yjs).
  async #locked(id, fn) {
    const prev = this.locks.get(id) ?? Promise.resolve();
    const run = prev.then(fn, fn);
    const tail = run.catch(() => {});
    this.locks.set(id, tail);
    try { return await run; } finally { if (this.locks.get(id) === tail) this.locks.delete(id); }
  }

  // ---------- reading ----------

  async list(team, { q = null, include_examples = true } = {}) {
    const rows = await this.db.all(`select id, title, created_by, created_at, updated_at, updated_by, share_mode, example from sheets_workbooks where team_id = $1 and deleted_at is null order by updated_at desc`, [team]);
    const words = q ? String(q).toLowerCase().split(/\s+/).filter(Boolean) : [];
    return rows.filter((r) => (include_examples || !r.example) && words.every((w) => r.title.toLowerCase().includes(w))).map(rowView);
  }

  async row(team, id) {
    const r = await this.db.get('select * from sheets_workbooks where id = $1 and team_id = $2 and deleted_at is null', [id, team]);
    if (!r) throw new BookError(`No spreadsheet with id ${id}. sheets_list_sheets has them.`, 404, 'not_found');
    return r;
  }

  // A workbook by id or by its exact title (any case).
  async find(team, ref) {
    const s = String(ref ?? '').trim();
    if (!s) throw new BookError('Say which spreadsheet: its id (wb_...) or its title.');
    let r = await this.db.get('select * from sheets_workbooks where id = $1 and team_id = $2 and deleted_at is null', [s, team]);
    if (!r) {
      const all = await this.db.all('select * from sheets_workbooks where team_id = $1 and deleted_at is null', [team]);
      const hits = all.filter((x) => x.title.toLowerCase() === s.toLowerCase());
      if (hits.length > 1) throw new BookError(`${hits.length} spreadsheets are called "${s}". Use the id: ${hits.map((h) => h.id).join(', ')}.`, 409, 'ambiguous');
      r = hits[0];
    }
    if (!r) throw new BookError(`No spreadsheet "${s}". sheets_list_sheets has them.`, 404, 'not_found');
    return r;
  }

  async byShareToken(token) {
    if (!token || String(token).length < 16) return null;
    return this.db.get(`select * from sheets_workbooks where share_token = $1 and share_mode <> 'off' and deleted_at is null`, [String(token)]);
  }

  // The document (and its formula engine), brought up to date with every change in the database.
  async open(row) {
    let e = this.cache.get(row.id);
    const fresh = !e || Number(row.state_seq) > e.seq;
    if (fresh) {
      e?.engine?.destroy();
      const doc = new Y.Doc({ gc: true });
      if (row.state) Y.applyUpdate(doc, fromB64(row.state), 'load');
      e = { id: row.id, team: row.team_id, doc, seq: Number(row.state_seq), engine: null };
      this.cache.set(row.id, e);
      if (this.cache.size > CACHE) {
        const [oldest] = this.cache.keys();
        this.cache.get(oldest)?.engine?.destroy();
        this.cache.delete(oldest);
      }
    }
    const ups = await this.db.all('select id, data from sheets_updates where workbook_id = $1 and id > $2 order by id', [row.id, e.seq]);
    if (ups.length) {
      Y.applyUpdate(e.doc, Y.mergeUpdates(ups.map((u) => fromB64(u.data))), 'remote');
      e.seq = Number(ups[ups.length - 1].id);
    }
    // Move it to the back of the cache order.
    this.cache.delete(row.id);
    this.cache.set(row.id, e);
    return e;
  }

  engine(e) { return (e.engine ??= new Engine(e.doc)); }

  // ---------- changing ----------

  async create(me, { title = 'Untitled spreadsheet', tabs, state = null, example = false } = {}) {
    const doc = new Y.Doc();
    if (state) Y.applyUpdate(doc, state);
    else initBook(doc, { title, tabs: tabs?.length ? tabs : [{ name: 'Sheet1' }] });
    if (state) doc.getMap('meta').set('title', title);
    const id = newId('wb');
    const at = nowIso();
    await this.db.run(`insert into sheets_workbooks (id, team_id, title, created_by, created_at, updated_at, updated_by, state, state_seq, example) values ($1, $2, $3, $4, $5, $5, $4, $6, 0, $7)`,
      [id, me.team_id, bookTitle(doc), me.id, at, toB64(Y.encodeStateAsUpdate(doc)), example ? 1 : 0]);
    await this.saveVersion(me, { id, team_id: me.team_id }, doc, { label: 'Created', auto: true });
    this.bus?.publish({ team: me.team_id, type: 'sheets.workbook.created', data: { workbook: id, title: bookTitle(doc) } });
    return this.row(me.team_id, id);
  }

  // Runs fn(doc, engine, entry) as one change. Returns { result, changed }.
  change(me, row, fn, opts) { return this.#locked(row.id, () => this.#change(me, row, fn, opts)); }

  async #change(me, row, fn, { origin = 'tool' } = {}) {
    const e = await this.open(row);
    const parts = [];
    const onUpdate = (u, o) => { if (o !== 'remote' && o !== 'load') parts.push(u); };
    e.doc.on('update', onUpdate);
    let result;
    try {
      const engine = this.engine(e);
      result = await fn(e.doc, engine, e);
    } finally {
      e.doc.off('update', onUpdate);
    }
    if (!parts.length) return { result, changed: false };
    const update = Y.mergeUpdates(parts);
    const at = nowIso();
    const ins = await this.db.run('insert into sheets_updates (team_id, workbook_id, data, by_id, created_at) values ($1, $2, $3, $4, $5) returning id', [row.team_id, row.id, toB64(update), me.id, at]);
    const seq = Number(ins.rows[0]?.id ?? 0);
    const t = bookTitle(e.doc);
    await this.db.run('update sheets_workbooks set updated_at = $2, updated_by = $3, title = $4 where id = $1', [row.id, at, me.id, t]);
    this.bus?.publish({ team: row.team_id, type: 'sheets.workbook.changed', data: { workbook: row.id, seq, update: toB64(update), by: { id: me.id, name: me.name }, origin } });
    await this.#maybeFold(row.id);
    await this.#maybeAutoVersion(me, row, e.doc);
    return { result, changed: true, seq };
  }

  async #maybeFold(id) {
    const n = await this.db.get('select count(*) as n from sheets_updates where workbook_id = $1', [id]);
    if (Number(n?.n ?? 0) < FOLD_AFTER) return;
    await this.db.tx(async (t) => {
      const row = await t.get(`select * from sheets_workbooks where id = $1${t.dialect === 'pg' ? ' for update' : ''}`, [id]);
      const ups = await t.all('select id, data from sheets_updates where workbook_id = $1 and id > $2 order by id', [id, row.state_seq]);
      if (ups.length < FOLD_AFTER) return;
      const merged = Y.mergeUpdates([...(row.state ? [fromB64(row.state)] : []), ...ups.map((u) => fromB64(u.data))]);
      const last = Number(ups[ups.length - 1].id);
      await t.run('update sheets_workbooks set state = $2, state_seq = $3 where id = $1', [id, toB64(merged), last]);
      await t.run('delete from sheets_updates where workbook_id = $1 and id <= $2', [id, last]);
    });
  }

  // ---------- versions ----------

  async #maybeAutoVersion(me, row, doc) {
    const last = await this.db.get('select created_at from sheets_versions where workbook_id = $1 order by created_at desc limit 1', [row.id]);
    if (last && Date.now() - Date.parse(last.created_at) < AUTO_VERSION_MS) return;
    await this.saveVersion(me, row, doc, { auto: true });
  }

  async saveVersion(me, row, doc, { label = null, auto = false } = {}) {
    const id = newId('v');
    await this.db.run('insert into sheets_versions (id, team_id, workbook_id, label, auto, state, created_by, created_at) values ($1, $2, $3, $4, $5, $6, $7, $8)',
      [id, row.team_id, row.id, label, auto ? 1 : 0, toB64(Y.encodeStateAsUpdate(doc)), me?.id ?? null, nowIso()]);
    if (auto) {
      const old = await this.db.all('select id from sheets_versions where workbook_id = $1 and auto = 1 and label is null order by created_at desc', [row.id]);
      for (const v of old.slice(KEEP_AUTO_VERSIONS)) await this.db.run('delete from sheets_versions where id = $1', [v.id]);
    }
    return { id, label, auto, created_at: nowIso(), created_by: me?.id ?? null };
  }

  async versions(row) {
    const rows = await this.db.all('select id, label, auto, created_by, created_at from sheets_versions where workbook_id = $1 order by created_at desc, id desc', [row.id]);
    return rows.map((v) => ({ id: v.id, label: v.label, auto: !!Number(v.auto), created_by: v.created_by, created_at: v.created_at }));
  }

  async versionDoc(row, versionId) {
    const v = await this.db.get('select * from sheets_versions where id = $1 and workbook_id = $2', [versionId, row.id]);
    if (!v) throw new BookError(`No version ${versionId} of this spreadsheet. sheets_list_versions has them.`, 404, 'not_found');
    const doc = new Y.Doc();
    Y.applyUpdate(doc, fromB64(v.state));
    return { doc, version: v };
  }

  // ---------- sharing and deleting ----------

  async setShare(row, mode) {
    const token = mode === 'off' ? row.share_token : row.share_token || crypto.randomBytes(18).toString('base64url');
    await this.db.run('update sheets_workbooks set share_mode = $2, share_token = $3 where id = $1', [row.id, mode, token]);
    return { mode, token: mode === 'off' ? null : token };
  }

  async remove(row) {
    await this.db.run('update sheets_workbooks set deleted_at = $2 where id = $1', [row.id, nowIso()]);
    this.cache.get(row.id)?.engine?.destroy();
    this.cache.delete(row.id);
    this.bus?.publish({ team: row.team_id, type: 'sheets.workbook.deleted', data: { workbook: row.id } });
  }
}

export const rowView = (r) => ({ id: r.id, title: r.title, created_by: r.created_by, created_at: r.created_at, updated_at: r.updated_at, updated_by: r.updated_by, shared: r.share_mode !== 'off' && !!r.share_mode, example: !!Number(r.example) });
