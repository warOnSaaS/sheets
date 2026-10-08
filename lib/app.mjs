import fs from 'node:fs';
import path from 'node:path';
import { openDb, migrate } from './db.mjs';
import { Bus } from './bus.mjs';
import { Store } from './store.mjs';
import { Team } from './team.mjs';
import { Files } from './files.mjs';
import { createMailer } from './mail.mjs';
import { authProvider } from './auth.mjs';
import { seal, unseal } from './seal.mjs';
import { runTool } from './tools.mjs';
import { seedExamples, seedWorkspace } from './demo.mjs';

const ROOT = path.resolve(new URL('..', import.meta.url).pathname);
export const VERSION = JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8')).version;

// Everything the server needs, wired once: the database (migrated on start), live events, workbooks,
// people, files and mail. server.mjs serves it; tests make their own with an in-memory database.
export async function createApp(env = process.env) {
  const db = await openDb({ url: env.DATABASE_URL, file: env.SQLITE_FILE });
  const applied = await migrate(db);
  if (applied.length) console.log(`sheets: database updated (${applied.join(', ')})`);
  const bus = new Bus(db);
  await bus.start();
  const secret = env.OAUTH_SECRET || env.SESSION_SECRET || 'dev-secret-change-me';
  let lastHost = env.PUBLIC_URL?.replace(/\/$/, '') ?? null;
  const app = {
    env, db, bus, version: VERSION,
    store: new Store({ db, bus }),
    team: new Team(db),
    files: new Files(db),
    mailer: await createMailer(env),
    authProvider: authProvider(env),
    openSignup: env.SHEETS_OPEN_SIGNUP === '1',
    examples: env.SHEETS_EXAMPLES !== '0',
    teamId: env.SHEETS_TEAM_ID || 'default',
    teamName: env.SHEETS_TEAM_NAME || 'Our team',
    seal: (s) => seal(s, secret),
    unseal: (s) => unseal(s, secret),
    seenHost(h) { if (!env.PUBLIC_URL && h) lastHost = h; },
    publicUrl: () => lastHost ?? `http://localhost:${env.PORT || 3994}`,
    urlFor: (id) => `${app.publicUrl()}/#/s/${id}`,
    shareUrl: (token) => (token ? `${app.publicUrl()}/v/${token}` : null),
  };
  app.run = (me, name, input, o) => runTool(app, me, name, input, o);
  app.seedWorkspace = (p) => seedWorkspace(app, p);
  await app.team.ensure(app.teamId, app.teamName);
  app.exampleList = app.examples ? await seedExamples(app) : [];
  app.close = async () => { await bus.stop(); await db.close(); };
  return app;
}
