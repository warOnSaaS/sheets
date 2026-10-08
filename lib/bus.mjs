import { EventEmitter } from 'node:events';
import crypto from 'node:crypto';

// Live updates. A change is sent to this server's open sockets, and with Postgres also by NOTIFY, so every
// other copy of the server tells its own sockets. Nothing is stored here: a screen that was away catches up
// from the workbook itself (sheets.sync with its Yjs state vector).
const CHANNEL = 'sheets_events';

export class Bus extends EventEmitter {
  constructor(db) {
    super();
    this.db = db;
    this.id = crypto.randomBytes(6).toString('hex');
    this.setMaxListeners(0);
  }

  async start() { this.stopListening = await this.db.listen(CHANNEL, (p) => this.#heard(p)); }
  async stop() { await this.stopListening?.(); }

  async #heard(payload) {
    let e;
    try { e = JSON.parse(payload); } catch { return; }
    if (e.origin === this.id) return;
    // A big change travels as a reference to its row in sheets_updates.
    if (e.ref) {
      const row = await this.db.get('select data from sheets_updates where id = $1', [e.ref]).catch(() => null);
      if (!row) return;
      e.data.update = row.data;
    }
    delete e.origin;
    delete e.ref;
    this.emit('event', e);
  }

  // audience: 'team' or a list of person ids.
  publish({ team, type, audience = 'team', data = {}, ephemeral = false }) {
    const e = { team, type, audience, data, at: new Date().toISOString(), ephemeral };
    this.emit('event', e);
    let payload = JSON.stringify({ ...e, origin: this.id });
    if (payload.length > 7000) {
      if (!data.seq) return e;
      payload = JSON.stringify({ ...e, data: { ...data, update: null }, ref: data.seq, origin: this.id });
    }
    this.db.notify(CHANNEL, payload).catch((err) => console.error('notify:', err.message));
    return e;
  }
}

export const canSee = (e, personId) => e.audience === 'team' || (Array.isArray(e.audience) && e.audience.includes(personId));
