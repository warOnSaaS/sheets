// Screenshots at 1440 and 390, light and dark, into .shots/ (git-ignored). Starts its own server on a
// temp SQLite file with the examples, signs in as Sam, and opens every screen. The browser is muted.
//   node scripts/shots.mjs [only-matching-name]
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { chromium } from 'playwright';

process.env.OAUTH_SECRET = 'shots';
process.env.SQLITE_FILE = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'sheets-shots-')), 'sheets.db');
const only = process.argv[2] ?? '';
const { createServer } = await import('../server.mjs');
const { issueTokens } = await import('../lib/auth.mjs');
const server = createServer();
await new Promise((r) => server.listen(0, r));
const app = await server.ready;
const base = `http://localhost:${server.address().port}`;
const sam = await app.team.add(app.teamId, { name: 'Sam Rivera', email: 'sam@acme-dental.example', role: 'owner' });
await app.team.add(app.teamId, { name: 'Jordan Lee', email: 'jordan@acme-dental.example' });
const pipe = await app.run(sam, 'sheets.create_sheet', { template: 'pipeline' });
const budget = await app.run(sam, 'sheets.create_sheet', { template: 'budget' });
const hiring = await app.run(sam, 'sheets.create_sheet', { template: 'hiring' });
const token = issueTokens(sam).access_token;
fs.mkdirSync('.shots', { recursive: true });
const browser = await chromium.launch({ args: ['--mute-audio'] });
const shots = [];
for (const [w, hgt, tag] of [[1440, 900, 'desk'], [390, 844, 'phone']]) {
  for (const mode of ['light', 'dark']) {
    const ctx = await browser.newContext({ viewport: { width: w, height: hgt }, colorScheme: mode, deviceScaleFactor: 2, hasTouch: tag === 'phone', isMobile: tag === 'phone' });
    const page = await ctx.newPage();
    page.on('pageerror', (e) => console.error(`page error (${tag} ${mode}):`, e.message));
    const shot = async (name) => { if (only && !name.includes(only)) return; const f = `.shots/${tag}-${mode}-${name}.png`; await page.screenshot({ path: f }); shots.push(f); };
    await page.goto(`${base}/`); await page.waitForSelector('.ex-card'); await shot('landing');
    await page.goto(`${base}/v/${app.exampleList[0].token}`); await page.waitForSelector('.c'); await page.waitForTimeout(300); await shot('view-only');
    await ctx.addCookies([{ name: 'sheets_session', value: token, url: base }]);
    await page.goto(`${base}/#/`); await page.waitForSelector('.hm-t'); await shot('home');
    await page.goto(`${base}/#/s/${pipe.id}`); await page.waitForSelector('.c'); await page.waitForTimeout(400); await shot('pipeline');
    await page.goto(`${base}/#/s/${pipe.id}?tab=Summary`); await page.waitForSelector('.sg-chart svg'); await page.waitForTimeout(300); await shot('pipeline-summary');
    await page.goto(`${base}/#/s/${budget.id}`); await page.waitForSelector('.sg-chart svg'); await page.waitForTimeout(300); await shot('budget');
    await page.goto(`${base}/#/s/${hiring.id}?tab=By%20role`); await page.waitForSelector('.c'); await page.waitForTimeout(300); await shot('hiring-pivot');
    if (tag === 'desk') {
      await page.goto(`${base}/#/s/${pipe.id}`); await page.waitForSelector('.c'); await page.waitForTimeout(300);
      const g = await page.locator('.ed-grid').boundingBox();
      await page.mouse.click(g.x + 46 + 108 * 4 + 40, g.y + 26 + 28 * 2 + 10);
      await page.keyboard.type('=SUM(E');
      await page.waitForTimeout(150); await shot('editing-formula');
      await page.keyboard.press('Escape');
      await page.click('[data-act="chart"]'); await page.waitForSelector('.ch-prev svg'); await shot('chart-dialog');
      await page.keyboard.press('Escape');
    }
    await page.goto(`${base}/#/connect`); await page.waitForSelector('.ct'); await shot('connect');
    await page.goto(`${base}/#/settings`); await page.waitForSelector('.st-s'); await shot('settings');
    await ctx.close();
  }
}
await browser.close();
server.close(); await app.close();
console.log(`${shots.length} screenshots in .shots/`);
process.exit(0);
