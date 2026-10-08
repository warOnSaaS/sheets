// Files Sheets makes (downloads: .xlsx, .csv, the full export). Kept in the database for a day, so any
// server copy can hand them out. Uploads come in as base64 tool input instead, so there is no upload route.
import { newId, nowIso } from './ids.mjs';

const KEEP_MS = 24 * 3600 * 1000;

export class Files {
  constructor(db, { prefix = '/files/sheets' } = {}) { this.db = db; this.prefix = prefix; }

  async put(me, { name, type, data }) {
    const id = newId('f');
    await this.db.run('delete from sheets_files where created_at < $1', [new Date(Date.now() - KEEP_MS).toISOString()]);
    await this.db.run('insert into sheets_files (id, team_id, uploader_id, name, type, size, data, created_at) values ($1, $2, $3, $4, $5, $6, $7, $8)',
      [id, me.team_id, me.id, name, type, data.length, Buffer.from(data).toString('base64'), nowIso()]);
    return { id, name, type, size: data.length, url: `${this.prefix}/${id}/${encodeURIComponent(name)}` };
  }

  async get(teamId, id) {
    const f = await this.db.get('select * from sheets_files where id = $1 and team_id = $2', [id, teamId]);
    return f ? { ...f, data: Buffer.from(f.data, 'base64') } : null;
  }
}

export async function serveFile(app, me, res, p) {
  const id = p.split('/')[3];
  const f = id ? await app.files.get(me.team_id, id) : null;
  if (!f) return res.writeHead(404, { 'content-type': 'text/plain' }).end('That download has expired. Make it again.');
  res.writeHead(200, { 'content-type': f.type, 'content-length': f.data.length, 'content-disposition': `attachment; filename="${f.name.replace(/[^\w .()-]/g, '_')}"`, 'cache-control': 'private, no-store' }).end(f.data);
}
