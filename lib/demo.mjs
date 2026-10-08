// The examples. A team called "examples" holds the three fictional example spreadsheets, each with a
// view-only link, so anyone can look at them before signing in. With open sign-up (the hosted demo),
// everyone who signs in gets their own copies to edit.
import { Y } from './book.mjs';
import { buildTemplate, TEMPLATES } from './templates.mjs';

export const EXAMPLES_TEAM = 'examples';
const ORDER = ['pipeline', 'budget', 'hiring'];

export async function seedExamples(app) {
  await app.team.ensure(EXAMPLES_TEAM, 'Examples');
  let owner = await app.db.get(`select * from sheets_people where team_id = $1 and role = 'owner'`, [EXAMPLES_TEAM]);
  if (!owner) owner = await app.team.add(EXAMPLES_TEAM, { id: 'p_examples', name: 'Sam Rivera', role: 'owner' });
  const have = await app.db.all('select id, title, share_token from sheets_workbooks where team_id = $1 and deleted_at is null and example = 1', [EXAMPLES_TEAM]);
  for (const name of ORDER) {
    if (have.some((h) => h.title === TEMPLATES[name].title)) continue;
    const row = await app.store.create(owner, { title: TEMPLATES[name].title, state: Y.encodeStateAsUpdate(buildTemplate(name)), example: true });
    await app.store.setShare(row, 'view');
  }
  const rows = await app.db.all('select id, title, share_token, updated_at from sheets_workbooks where team_id = $1 and deleted_at is null and example = 1', [EXAMPLES_TEAM]);
  return ORDER.map((n) => rows.find((r) => r.title === TEMPLATES[n].title)).filter(Boolean).map((r) => ({ id: r.id, title: r.title, about: TEMPLATES[ORDER.find((n) => TEMPLATES[n].title === r.title)].about, token: r.share_token }));
}

export async function seedWorkspace(app, person) {
  for (const name of ORDER) await app.store.create(person, { title: TEMPLATES[name].title, state: Y.encodeStateAsUpdate(buildTemplate(name)), example: true });
}
