import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { chromium } from 'playwright';

// Agent parity (ROADMAP 3.2), enforced:
// 1. Screen-to-tool: open every screen, menu, dialog and popover, at desk and phone width, and collect every
//    button, menu item, form, select and file picker. Each must name a tool from the catalogue (data-tool),
//    be the submit button or a field of a form that does, or say data-tool="none" with a reason in data-why
//    when it only moves around the page. Links are navigation.
// 2. No side doors: screen code only talks to /api/tools/*, /files/sheets/* downloads and the /ws live feed.
// 3. A parity report in .shots/parity-report.json.
// It also edits through the screen, so a typed formula is checked against the server's value.

process.env.OAUTH_SECRET = 'parity';
process.env.SQLITE_FILE = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'sheets-parity-')), 'sheets.db');

let server, base, app, browser, catalogue, sam, token, sheet;
before(async () => {
  const mod = await import('../server.mjs');
  catalogue = new Set(mod.catalogue().map((t) => t.name));
  server = mod.createServer();
  await new Promise((r) => server.listen(0, r));
  app = await server.ready;
  base = `http://localhost:${server.address().port}`;
  sam = await app.team.add(app.teamId, { name: 'Sam Rivera', email: 'sam@acme-dental.example', role: 'owner' });
  await app.team.add(app.teamId, { name: 'Jordan Lee', email: 'jordan@acme-dental.example' });
  sheet = await app.run(sam, 'sheets.create_sheet', { template: 'pipeline' });
  await app.run(sam, 'sheets.filter', { sheet: sheet.id, tab: 'Deals', range: 'A1:I13' });
  await app.run(sam, 'sheets.share', { sheet: sheet.id, link: 'view' });
  await app.run(sam, 'sheets.delete_sheet', { sheet: sheet.id }, { via: 'mcp', client: 'Claude' }).catch(() => {});
  const { issueTokens } = await import('../lib/auth.mjs');
  token = issueTokens(sam).access_token;
  browser = await chromium.launch({ args: ['--mute-audio'] });
});
after(async () => { await browser?.close(); server?.closeAllConnections?.(); server?.close(); await app?.close(); });

const collect = (page) => page.evaluate(() => {
  const out = [];
  const visible = (el) => !!(el.offsetWidth || el.offsetHeight || el.getClientRects().length) || el.type === 'file';
  for (const el of document.querySelectorAll('button, [role=menuitem], form, input[type=file], select')) {
    if (!visible(el) && el.tagName !== 'FORM') continue;
    if (el.closest('[data-auth]')) continue;
    const form = el.closest('form');
    let tool = el.getAttribute('data-tool');
    let how = tool ? 'tool' : null;
    if (tool === 'none') { how = el.getAttribute('data-why') ? `page helper: ${el.getAttribute('data-why')}` : null; tool = null; }
    if (!how && el.tagName === 'BUTTON' && el.type === 'submit' && form?.dataset.tool) { tool = form.dataset.tool === 'none' ? null : form.dataset.tool; how = 'submit'; }
    if (!how && (el.tagName === 'SELECT' || el.tagName === 'INPUT') && form?.dataset.tool) { tool = form.dataset.tool === 'none' ? null : form.dataset.tool; how = 'form field'; }
    out.push({ tag: el.tagName.toLowerCase(), text: (el.getAttribute('aria-label') || el.textContent || '').trim().slice(0, 50), tool, how });
  }
  return out;
});

test('every action on every screen has a tool, at desk and phone width', async () => {
  const report = { screens: {}, toolsOnScreens: new Set(), problems: [] };
  for (const [w, h, tag] of [[1440, 900, 'desk'], [390, 844, 'phone']]) {
    const ctx = await browser.newContext({ viewport: { width: w, height: h }, hasTouch: tag === 'phone', isMobile: tag === 'phone' });
    const page = await ctx.newPage();
    const errors = [];
    page.on('pageerror', (e) => errors.push(e.message));
    const check = async (name) => {
      await page.waitForTimeout(200);
      const found = await collect(page);
      report.screens[`${tag} ${name}`] = found.length;
      for (const f of found) {
        if (f.tool) report.toolsOnScreens.add(f.tool);
        if (!f.how) report.problems.push(`${tag} ${name}: <${f.tag}> "${f.text}" names no tool`);
        else if (f.tool && !catalogue.has(f.tool)) report.problems.push(`${tag} ${name}: <${f.tag}> "${f.text}" names ${f.tool}, not in the catalogue`);
      }
    };
    const closeAll = async () => { await page.keyboard.press('Escape'); await page.evaluate(() => { document.querySelectorAll('dialog[open]').forEach((d) => d.close()); document.querySelectorAll('.sh-menu,.sh-pop').forEach((m) => m.remove()); }); };
    const open = async (sel, name) => {
      const el = page.locator(sel).first();
      if (!(await el.count()) || !(await el.isVisible())) { report.problems.push(`${tag}: could not open ${name} (${sel})`); return; }
      await el.click();
      await check(name);
      await closeAll();
    };
    // Signed out: the view-only link.
    await page.goto(`${base}/v/${app.exampleList[0]?.token ?? (await app.store.row(app.teamId, sheet.id)).share_token}`);
    await page.waitForSelector('.c');
    await check('view-only');
    await ctx.addCookies([{ name: 'sheets_session', value: token, url: base }]);
    await page.goto(`${base}/#/`); await page.waitForSelector('.hm-t'); await check('home');
    await open('[data-row]', 'home row menu');
    await page.goto(`${base}/#/connect`); await page.waitForSelector('.ct'); await check('connect');
    await page.goto(`${base}/#/settings`); await page.waitForSelector('.st-s'); await check('settings');
    await page.goto(`${base}/#/s/${sheet.id}`); await page.waitForSelector('.c'); await page.waitForTimeout(300);
    await check('editor');
    if (tag === 'desk') {
      for (const act of ['color', 'fill', 'freeze', 'chart', 'pivot', 'comment', 'crm', 'find', 'comments', 'history', 'download', 'share']) await open(`[data-act="${act}"]`, `toolbar ${act}`);
      await open('[data-tab-menu]', 'tab menu');
      // The filter button in the header row, and the right-click menu.
      const g = await page.locator('.ed-grid').boundingBox();
      const fb = await page.locator('.c-f').first().boundingBox();
      await page.mouse.click(fb.x + fb.width / 2, fb.y + fb.height / 2);
      await page.waitForSelector('.sh-pop.is-filter');
      await check('filter popover'); await closeAll();
      await page.mouse.click(g.x + 300, g.y + 200, { button: 'right' });
      await check('cell menu'); await closeAll();
    } else {
      for (const act of ['comments', 'history', 'download', 'share']) await open(`[data-act="${act}"]`, `phone ${act}`);
    }
    // The Summary tab has a chart.
    await page.goto(`${base}/#/s/${sheet.id}?tab=Summary`); await page.waitForSelector('.sg-chart svg');
    await check('editor with chart');
    if (tag === 'desk') await open('.sg-chart [data-chart-menu]', 'chart menu');
    assert.deepEqual(errors, [], `${tag}: page errors`);
    await ctx.close();
  }
  const used = [...report.toolsOnScreens].sort();
  const unused = [...catalogue].filter((t) => !report.toolsOnScreens.has(t)).sort();
  const actions = Object.values(report.screens).reduce((a, b) => a + b, 0);
  const summary = { screens: report.screens, actions, covered: actions - report.problems.length, parity: `${Math.round(((actions - report.problems.length) / actions) * 100)}%`, tools_on_screens: used, tools_without_a_screen: unused, problems: report.problems };
  fs.mkdirSync('.shots', { recursive: true });
  fs.writeFileSync('.shots/parity-report.json', JSON.stringify(summary, null, 2));
  console.log(`parity ${summary.parity}: ${actions} screen actions on ${Object.keys(report.screens).length} screens, ${used.length} tools used by screens, ${unused.length} tools with no screen (${unused.join(', ')})`);
  assert.deepEqual(report.problems, []);
});

test('editing through the screen: a typed formula is worked out here and on the server', async () => {
  const ctx = await browser.newContext({ viewport: { width: 1440, height: 900 } });
  await ctx.addCookies([{ name: 'sheets_session', value: token, url: base }]);
  const page = await ctx.newPage();
  const s = await app.run(sam, 'sheets.create_sheet', { values: [['Amount'], [100], [250]] });
  await page.goto(`${base}/#/s/${s.id}`);
  await page.waitForSelector('.c');
  const g = await page.locator('.ed-grid').boundingBox();
  const cell = (r, c) => ({ x: g.x + 46 + 108 * c + 30, y: g.y + 26 + 28 * r + 14 });
  await page.mouse.click(cell(3, 0).x, cell(3, 0).y);
  await page.keyboard.type('=SUM(A2:A3)');
  await page.keyboard.press('Enter');
  await page.waitForFunction(() => [...document.querySelectorAll('.c')].some((x) => x.textContent === '350'));
  await page.waitForTimeout(300);
  const r = await app.run(sam, 'sheets.read_range', { sheet: s.id, range: 'A4' });
  assert.deepEqual([r.values[0][0], r.formulas[0][0]], [350, '=SUM(A2:A3)']);
  // Typing money makes a currency cell; undo takes it back.
  await page.keyboard.type('$1,200');
  await page.keyboard.press('Enter');
  await page.waitForFunction(() => [...document.querySelectorAll('.c')].some((x) => x.textContent === '$1,200'));
  await page.waitForTimeout(200);
  await page.keyboard.press('Control+z');
  await page.waitForTimeout(500);
  assert.equal((await app.run(sam, 'sheets.read_range', { sheet: s.id, range: 'A5' })).values[0][0], null);
  // Bold from the toolbar.
  await page.mouse.click(cell(1, 0).x, cell(1, 0).y);
  await page.click('[data-act="b"]');
  await page.waitForTimeout(400);
  assert.equal((await app.run(sam, 'sheets.read_range', { sheet: s.id, range: 'A2', formats: true })).formats[0][0].b, true);
  // Another person's change shows up without a reload.
  await app.run(sam, 'sheets.write_range', { sheet: s.id, range: 'B1', values: [['From an agent']] });
  await page.waitForFunction(() => [...document.querySelectorAll('.c')].some((x) => x.textContent === 'From an agent'), null, { timeout: 8000 });
  await ctx.close();
});

test('no side doors: screen code only calls tools, downloads and the live feed', () => {
  const dir = new URL('../public/app/', import.meta.url);
  for (const f of fs.readdirSync(dir).filter((x) => /\.mjs$/.test(x))) {
    const src = fs.readFileSync(new URL(f, dir), 'utf8');
    for (const m of src.matchAll(/\bfetch\(\s*(`[^`]*`|'[^']*'|"[^"]*")/g)) {
      const target = m[1].slice(1, -1);
      assert.ok(/^\/api\/tools\//.test(target), `${f}: fetch(${m[1]}) is a side door`);
      assert.equal(f, 'page.mjs', `${f}: only the page bootstrap fetches; screens use ctx.callTool`);
    }
    assert.ok(!/\bfetch\(\s*[a-zA-Z_$]/.test(src), `${f}: fetch with a computed address`);
    assert.ok(!/XMLHttpRequest|sendBeacon|EventSource/.test(src), `${f}: another way to the server`);
    for (const m of src.matchAll(/new WebSocket\(([^)]*)\)/g)) assert.match(m[1], /\/ws/, `${f}: a socket to somewhere other than /ws`);
  }
});
