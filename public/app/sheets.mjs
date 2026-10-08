// wOS Sheets screens: the list of spreadsheets, the editor (grid.mjs), Connect your AI, and Settings.
// The same module runs standalone (public/app/page.mjs makes the context) and inside the wOS suite
// (screens.mjs, built from this file). ctx.callTool is the only way to the server.
import { openEditor } from './grid.mjs';
import { h, none, tool, I, dialog, menu, toast, initials, ago, copyText, fileToBase64, showError } from './ui.mjs';

const TEMPLATES = [
  { id: null, title: 'Blank', about: 'An empty spreadsheet' },
  { id: 'pipeline', title: 'Sales pipeline', about: 'Deals, weighted value, a chart' },
  { id: 'budget', title: 'Budget', about: 'Months, totals, net and margin' },
  { id: 'hiring', title: 'Hiring tracker', about: 'Candidates and a pivot by role' },
];

export function mount(el, ctx) {
  el.classList.add('sheets-root');
  if (!ctx.standalone) el.classList.add('in-suite');
  let current = null;
  let token = 0;

  function route(path) {
    const my = ++token;
    current?.destroy?.();
    current = null;
    const p = (path || '/').split('?')[0];
    if (ctx.view) { current = openEditor(el, ctx, { sheet: ctx.view.sheet, view: ctx.view }); return; }
    let m;
    if ((m = /^\/s\/([^/]+)/.exec(p))) current = openEditor(el, ctx, { sheet: decodeURIComponent(m[1]) });
    else if (p === '/connect') current = connectPage(el, ctx);
    else if (p === '/settings') current = settingsPage(el, ctx, () => my === token);
    else if ((m = /^\/copy\/([^/]+)/.exec(p))) current = copyPage(el, ctx, m[1]);
    else current = homePage(el, ctx, () => my === token);
  }
  route(ctx.path);
  return { unmount() { current?.destroy?.(); el.innerHTML = ''; }, update(p) { ctx.path = p; route(p); } };
}

export default { title: 'Sheets', mount };

// ---------- the frame around list pages ----------

function frame(ctx, active, body) {
  const nav = [['/', 'Spreadsheets', I.grid], ['/connect', 'Connect your AI', I.plug], ['/settings', 'Settings', I.gear]];
  const link = (p) => (ctx.standalone ? `#${p}` : `/a/sheets${p === '/' ? '' : p}`);
  return `<div class="pg">
  ${ctx.standalone ? `<header class="pg-top"><a class="pg-brand" href="#/"><span class="pg-mark">${I.grid}</span>Sheets</a>
    <nav class="pg-nav">${nav.map(([p, l, ic]) => `<a href="${link(p)}" data-nav="${p}"${active === p ? ' aria-current="page"' : ''}>${ic}<span>${h(l)}</span></a>`).join('')}</nav>
    <div class="pg-me"><span class="ui-avatar is-sm" title="${h(ctx.me?.name ?? '')}">${h(initials(ctx.me?.name))}</span><a class="ui-btn is-ghost is-sm" href="/logout">Sign out</a></div></header>`
    : `<nav class="pg-nav is-inapp">${nav.map(([p, l, ic]) => `<a href="${link(p)}" data-nav="${p}"${active === p ? ' aria-current="page"' : ''}>${ic}<span>${h(l)}</span></a>`).join('')}</nav>`}
  <main class="ui-page pg-main">${body}</main></div>`;
}

function wireNav(el, ctx) {
  el.addEventListener('click', (e) => {
    const a = e.target.closest('a[data-nav], a[data-go]');
    if (!a || ctx.standalone && a.dataset.nav) return;
    if (a.dataset.go || !ctx.standalone) { e.preventDefault(); ctx.navigate(a.dataset.go ?? a.dataset.nav); }
  });
}

// ---------- home: the spreadsheets ----------

function homePage(el, ctx, alive) {
  el.innerHTML = frame(ctx, '/', `
  <div class="ui-ph"><div><h1>Spreadsheets</h1><p>Your team's sheets. Your own AI can work in them too.</p></div>
    <div class="ph-acts"><label class="ui-btn is-quiet" title="Import an .xlsx or .csv file">${I.upload}<span>Import</span><input type="file" class="sr-file" accept=".xlsx,.csv,text/csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" ${tool('sheets.import_xlsx')} aria-label="Import a file"></label>
    <button type="button" class="ui-btn is-accent" data-new="" ${tool('sheets.create_sheet')}>${I.plus}<span>New spreadsheet</span></button></div></div>
  <section class="hm-start" aria-label="Start from"><h2 class="ui-label">Start from</h2><div class="hm-tpl">${TEMPLATES.map((t) => `<button type="button" class="hm-t" data-new="${t.id ?? ''}" ${tool('sheets.create_sheet')}><span class="hm-thumb${t.id ? ` is-${t.id}` : ''}" aria-hidden="true">${thumb(t.id)}</span><b>${h(t.title)}</b><span>${h(t.about)}</span></button>`).join('')}</div></section>
  <a class="hm-ai" href="${ctx.standalone ? '#/connect' : '/a/sheets/connect'}" data-go="/connect"><span class="hm-ai-ic">${I.plug}</span><span><b>Let your AI work in these sheets</b><span>Connect Claude, ChatGPT, Claude Code or Codex. It reads exact values and formulas, builds tables and charts. Your subscription, no credits here.</span></span><span class="hm-ai-go">Connect</span></a>
  <section aria-label="All spreadsheets"><div class="hm-bar"><h2 class="ui-label">All spreadsheets</h2><form class="ui-search hm-q" role="search" ${tool('sheets.list_sheets')}>${I.search}<input name="q" placeholder="Search titles" aria-label="Search titles"></form></div>
  <div class="hm-list" aria-busy="true"><p class="ui-empty">Loading…</p></div></section>`);
  wireNav(el, ctx);
  const list = el.querySelector('.hm-list');
  async function load(q = '') {
    let out;
    try { out = await ctx.callTool('sheets.list_sheets', q ? { q } : {}); } catch (e) { list.innerHTML = `<p class="ui-empty">${h(e.message)}</p>`; return; }
    if (!alive()) return;
    list.removeAttribute('aria-busy');
    if (!out.sheets.length) { list.innerHTML = `<div class="hm-empty">${q ? '<p>Nothing matches.</p>' : '<p><b>No spreadsheets yet.</b></p><p class="ui-hint">Start from one above, import a file, or ask your AI to make one.</p>'}</div>`; return; }
    list.innerHTML = `<div class="ui-dtable-wrap"><table class="ui-dtable hm-t"><thead><tr><th>Title</th><th class="hm-when">Last change</th><th class="end"><span class="sr-only">Actions</span></th></tr></thead><tbody>${out.sheets.map((s) => `<tr><td><a class="hm-link" href="${ctx.standalone ? `#/s/${s.id}` : `/a/sheets/s/${s.id}`}" data-go="/s/${s.id}"><span class="hm-ic">${I.grid}</span><span><b>${h(s.title)}</b>${s.shared ? ' <span class="ui-chip is-outline">Link on</span>' : ''}</span></a></td><td class="hm-when">${h(ago(s.updated_at))}</td><td class="end"><button type="button" class="ui-btn is-ghost is-sm" data-row="${s.id}" ${none('opens the menu for this spreadsheet')} aria-label="More for ${h(s.title)}">${I.more}</button></td></tr>`).join('')}</tbody></table></div>`;
    list._sheets = out.sheets;
  }
  load();
  el.querySelector('.hm-q').addEventListener('submit', (e) => { e.preventDefault(); load(e.target.q.value.trim()); });
  el.querySelector('.hm-q input').addEventListener('input', (e) => { clearTimeout(load.t); load.t = setTimeout(() => load(e.target.value.trim()), 250); });
  el.addEventListener('click', async (e) => {
    const n = e.target.closest('[data-new]');
    if (n) {
      n.disabled = true;
      try {
        const out = await ctx.callTool('sheets.create_sheet', n.dataset.new ? { template: n.dataset.new } : { title: 'Untitled spreadsheet' });
        ctx.navigate(`/s/${out.id}`);
      } catch (err) { toast(err.message); n.disabled = false; }
      return;
    }
    const r = e.target.closest('[data-row]');
    if (r) {
      const s = list._sheets.find((x) => x.id === r.dataset.row);
      menu(document.body, r, [
        { label: 'Open', why: 'opens the spreadsheet', icon: I.grid, run: () => ctx.navigate(`/s/${s.id}`) },
        { label: 'Make a copy', tool: 'sheets.copy_sheet', icon: I.copy, run: async () => { const out = await ctx.callTool('sheets.copy_sheet', { sheet: s.id }); toast(`Copied: ${out.title}`); load(); } },
        { label: 'Download as Excel', tool: 'sheets.export', icon: I.download, run: async () => { const out = await ctx.callTool('sheets.export', { sheet: s.id, format: 'xlsx' }); const a = document.createElement('a'); a.href = out.file.url; a.download = out.file.name; document.body.appendChild(a); a.click(); a.remove(); } },
        { sep: true },
        { label: 'Delete', tool: 'sheets.delete_sheet', icon: I.x, danger: true, run: () => dialog(document.body, { title: `Delete ${s.title}?`, toolName: 'sheets.delete_sheet', submit: 'Delete', body: '<p>It goes for everyone on the team, and its share link stops working.</p>', onSubmit: async () => { await ctx.callTool('sheets.delete_sheet', { sheet: s.id }); toast('Deleted'); load(); } }) },
      ], { label: 'Spreadsheet' });
    }
  });
  el.querySelector('.sr-file').addEventListener('change', async (e) => {
    const f = e.target.files?.[0];
    e.target.value = '';
    if (!f) return;
    if (f.size > 3.2 * 1024 * 1024) { toast('Files up to about 3 MB here for now. Larger ones import through the API or on your own server.'); return; }
    toast(`Importing ${f.name}…`);
    try {
      const b64 = await fileToBase64(f);
      const csv = /\.csv$/i.test(f.name) || f.type === 'text/csv';
      const out = csv ? await ctx.callTool('sheets.import_csv', { content_base64: b64, name: f.name }) : await ctx.callTool('sheets.import_xlsx', { content_base64: b64, name: f.name });
      importReport(ctx, out, csv, f.name);
    } catch (err) { toast(err.message); }
  });
  return { destroy() {} };
}

function importReport(ctx, out, csv, name) {
  const r = out.report ?? {};
  const body = csv
    ? `<p><b>${h(name)}</b>: ${r.rows} rows and ${r.columns} columns, split on ${r.delimiter === 'tab' ? 'tabs' : `"${h(r.delimiter)}"`}. Numbers, money, percents and dates were read as such.</p>`
    : `<p>${h(r.summary ?? '')}</p>
      ${r.carried?.length ? `<p class="ui-label">Carried over</p><ul class="rp-l is-ok">${r.carried.map((x) => `<li>${I.check}${h(x[0].toUpperCase() + x.slice(1))}</li>`).join('')}</ul>` : ''}
      ${r.dropped?.length ? `<p class="ui-label">Not carried over</p><ul class="rp-l">${r.dropped.map((d) => `<li>${I.x}<span>${h(d.what)} <b>${d.count}</b>${d.examples?.length ? `<small>${h(d.examples.join(', '))}</small>` : ''}</span></li>`).join('')}</ul>` : ''}
      ${r.formula_problems?.length ? `<p class="ui-label">Formulas to check</p><ul class="rp-l">${r.formula_problems.slice(0, 12).map((p) => `<li>${I.x}<span><code>${h(p.cell)}</code> ${h(p.problem)}<small>${h(p.formula)}</small></span></li>`).join('')}${r.formula_problems.length > 12 ? `<li>…and ${r.formula_problems.length - 12} more</li>` : ''}</ul>` : ''}`;
  dialog(document.body, { title: 'Import report', body, cancel: 'Close', submit: 'Open it', toolName: 'sheets.get_sheet', wide: true, onSubmit: () => { ctx.navigate(`/s/${out.sheet}`); } });
}

function thumb(id) {
  const bars = { pipeline: [40, 62, 78, 55], budget: [30, 45, 52, 70], hiring: [70, 50, 38, 22] }[id];
  if (!bars) return `<svg viewBox="0 0 120 64" width="120" height="64"><g class="th-g">${[0, 1, 2, 3].map((i) => `<line x1="0" x2="120" y1="${14 + i * 14}" y2="${14 + i * 14}"/>`).join('')}${[0, 1, 2].map((i) => `<line y1="0" y2="64" x1="${30 + i * 30}" x2="${30 + i * 30}"/>`).join('')}</g><path class="th-plus" d="M60 24v16M52 32h16"/></svg>`;
  return `<svg viewBox="0 0 120 64" width="120" height="64"><g class="th-g">${[0, 1, 2, 3].map((i) => `<line x1="0" x2="58" y1="${14 + i * 14}" y2="${14 + i * 14}"/>`).join('')}<line y1="0" y2="64" x1="22" x2="22"/></g>${bars.map((b, i) => `<rect class="th-b" x="${68 + i * 12}" y="${58 - b * 0.7}" width="8" height="${b * 0.7}" rx="1.5"/>`).join('')}</svg>`;
}

// ---------- connect your AI ----------

const CLAUDE_ADD = (name, url) => `https://claude.ai/customize/connectors?modal=add-custom-connector&connectorName=${encodeURIComponent(name)}&connectorUrl=${encodeURIComponent(url)}`;

function connectPage(el, ctx) {
  const host = location.origin;
  const mcp = ctx.standalone ? `${host}/mcp` : `${host}/mcp`;
  const cmd = (app) => `curl -fsSL ${host}/connect/${app} | sh`;
  const tile = ({ ic, title, where, action, hint, term = false }) => `<section class="ct"><div class="ct-top"><span class="ct-logo">${ic}${term ? '<i class="ct-term" aria-hidden="true">&gt;_</i>' : ''}</span><div><h3>${h(title)}</h3><p>${h(where)}</p></div></div><div class="ct-act">${action}</div><p class="ct-hint">${h(hint)}</p></section>`;
  const copyBtn = (text, label = 'Copy') => `<button type="button" class="ui-btn is-quiet is-sm" data-copy="${h(text)}" ${none('copies text to paste somewhere else')}>${h(label)}</button>`;
  el.innerHTML = frame(ctx, '/connect', `
  <div class="ui-ph"><div><h1>Connect your AI</h1><p>Sheets has no AI of its own. Pick the app you already use, click once, and it can work in your sheets with your subscription. You never pay us for AI.</p></div></div>
  <div class="ct-group"><h2 class="ui-label">Chat apps</h2><div class="ct-grid">
    ${tile({ ic: I.chat, title: 'Claude', where: 'Web, desktop and phone', action: `<a class="ui-btn is-accent is-lg" href="${h(CLAUDE_ADD('wOS Sheets', mcp))}" target="_blank" rel="noopener" data-copy-also="${h(mcp)}">Add to Claude</a>`, hint: 'Click Add, then Connect, then sign in here once.' })}
    ${tile({ ic: I.chat, title: 'ChatGPT', where: 'Web, with developer mode on', action: `<a class="ui-btn is-accent is-lg" href="https://chatgpt.com/#settings/Connectors" target="_blank" rel="noopener" data-copy-also="${h(mcp)}">Open ChatGPT</a>`, hint: 'The address is copied. Paste it in Settings, Apps, Create.' })}
  </div></div>
  <div class="ct-group"><h2 class="ui-label">Terminal</h2><div class="ct-grid is-term">
    ${tile({ ic: I.terminal, term: true, title: 'Claude Code', where: 'Terminal', action: `<div class="ui-copy ct-cmd"><code>${h(cmd('claude'))}</code>${copyBtn(cmd('claude'))}</div>`, hint: 'Paste it in your terminal. A browser opens to sign in.' })}
    ${tile({ ic: I.terminal, term: true, title: 'Codex', where: 'Terminal', action: `<div class="ui-copy ct-cmd"><code>${h(cmd('codex'))}</code>${copyBtn(cmd('codex'))}</div>`, hint: 'Paste it in your terminal. A browser opens to sign in.' })}
  </div></div>
  <div class="ct-group"><h2 class="ui-label">Any app that speaks MCP</h2><div class="ui-copy ct-cmd is-wide"><code>${h(mcp)}</code>${copyBtn(mcp, 'Copy address')}</div><p class="ui-hint">Also as plain HTTPS for GPT Actions and scripts: <a href="/openapi.json">openapi.json</a>. Every button in Sheets is one of these tools, so your AI can do anything you can.</p></div>
  <div class="ct-group"><h2 class="ui-label">Things to ask it</h2><ul class="ct-ask">
    <li>Build a pipeline sheet from my CRM deals, with weighted value by stage and a column chart.</li>
    <li>Read the Budget sheet and tell me which month has the lowest margin. Show the formula.</li>
    <li>Add a Total row under the table on Deals and freeze the header.</li>
    <li>Make a pivot of candidates by role with the average score.</li>
  </ul></div>`);
  wireNav(el, ctx);
  el.addEventListener('click', (e) => {
    const c = e.target.closest('[data-copy]');
    if (c) { copyText(c.dataset.copy).then(() => { c.textContent = 'Copied'; toast('Copied'); setTimeout(() => { c.textContent = c.dataset.copy.startsWith('http') && !c.dataset.copy.includes('curl') ? 'Copy address' : 'Copy'; }, 1800); }); }
    const a = e.target.closest('[data-copy-also]');
    if (a) copyText(a.dataset.copyAlso).then(() => toast('The address is copied too.'), () => {});
  });
  return { destroy() {} };
}

// ---------- settings ----------

function settingsPage(el, ctx, alive) {
  el.innerHTML = frame(ctx, '/settings', `<div class="ui-ph"><div><h1>Settings</h1></div></div><div class="st" aria-busy="true"><p class="ui-empty">Loading…</p></div>`);
  wireNav(el, ctx);
  const box = el.querySelector('.st');
  async function load() {
    let s, people, approvals;
    try { [s, people, approvals] = await Promise.all([ctx.callTool('sheets.get_settings', {}), ctx.callTool('sheets.list_people', {}), ctx.callTool('sheets.list_approvals', {})]); } catch (e) { box.innerHTML = `<p class="ui-empty">${h(e.message)}</p>`; return; }
    if (!alive()) return;
    const admin = ['owner', 'admin'].includes(s.me.role);
    box.removeAttribute('aria-busy');
    box.innerHTML = `
    ${approvals.approvals.length ? `<section class="st-s"><h2>Waiting for your yes</h2><p class="ui-hint">An AI asked to do these. Nothing happens until you say yes.</p>${approvals.approvals.map((a) => `<div class="st-row"><div><b>${h(a.title)}</b><span>${h(a.requested_by)} · ${h(ago(a.created_at))} · ${h(JSON.stringify(a.input).slice(0, 120))}</span></div><button type="button" class="ui-btn is-quiet is-sm" data-decide="${a.id}" data-yes="0" ${tool('sheets.decide_approval')}>Decline</button><button type="button" class="ui-btn is-accent is-sm" data-decide="${a.id}" data-yes="1" ${tool('sheets.decide_approval')}>Approve</button></div>`).join('')}</section>` : ''}
    <section class="st-s"><h2>Appearance</h2><div class="ui-seg" role="group" aria-label="Theme">${['auto', 'light', 'dark'].map((t) => `<button type="button" data-theme="${t}" ${tool('sheets.set_preferences')} aria-pressed="${s.prefs.theme === t}">${{ auto: 'Match my device', light: 'Light', dark: 'Dark' }[t]}</button>`).join('')}</div></section>
    <section class="st-s"><h2>People</h2><p class="ui-hint">Everyone here can open and edit every spreadsheet on the team.</p>
      <div class="st-people">${people.people.map((p) => `<div class="st-row"><span class="ui-avatar is-sm">${h(initials(p.name))}</span><div><b>${h(p.name)}${p.id === s.me.id ? ' (you)' : ''}</b><span>${h([p.email, p.github && `@${p.github}`].filter(Boolean).join(' · '))}</span></div><span class="ui-chip is-outline">${h(p.role)}</span>${admin && p.role !== 'owner' && p.id !== s.me.id ? `<button type="button" class="ui-btn is-ghost is-sm" data-remove="${h(p.email ?? p.github ?? p.name)}" ${tool('sheets.remove_person')}>Remove</button>` : ''}</div>`).join('')}</div>
      ${admin ? `<form class="st-add" ${tool('sheets.add_person')}><input class="ui-input" name="name" placeholder="Name" required maxlength="100"><input class="ui-input" name="email" type="email" placeholder="Email" required><button class="ui-btn is-quiet" type="submit">Add</button></form>` : ''}</section>
    <section class="st-s"><h2>CRM</h2>${s.crm.connected ? `<p>Connected${s.crm.url ? ` to <code>${h(s.crm.url)}</code>` : ' through the wOS suite'}${s.crm.via === 'server' ? ' (set by whoever runs this server)' : ''}. Pull records from the toolbar in any spreadsheet.</p>${admin && s.crm.via === 'settings' ? `<button type="button" class="ui-btn is-quiet is-sm" data-crm-off ${tool('sheets.connect_crm')}>Disconnect</button>` : ''}` : admin ? `<p class="ui-hint">Pull deals, contacts and more into a sheet, kept refreshable. Give the CRM's address and an access token from it (in the CRM, ask for crm.create_access_token, read only is enough).</p><form class="st-crm" ${tool('sheets.connect_crm')}><input class="ui-input" name="url" type="url" placeholder="https://crm.example.com" required><input class="ui-input" name="token" placeholder="Access token (optional for a demo CRM)"><button class="ui-btn is-quiet" type="submit">Connect</button></form>` : '<p class="ui-hint">No CRM connected. An admin can connect one here.</p>'}</section>
    <section class="st-s"><h2>Your data</h2><p class="ui-hint">Everything, any time: every spreadsheet as .xlsx, plus all of it as JSON. Stored in ${h(s.storage === 'postgres' ? 'Postgres' : 'SQLite')}.</p>${admin ? `<button type="button" class="ui-btn is-quiet" data-export ${tool('sheets.export_data')}>${I.download}Export everything</button>` : ''}</section>
    <section class="st-s"><h2>Host it yourself, free</h2><p>Sheets is open source (AGPL-3.0). <code>docker compose up</code> runs it with Postgres on any server; one person can run it with Node and a SQLite file. No licence check, ever. <a href="https://github.com/warOnSaaS/sheets#host-it-yourself">How to host it</a></p></section>
    ${ctx.standalone ? '<section class="st-s"><a class="ui-btn is-quiet" href="/logout">Sign out</a></section>' : ''}`;
  }
  load();
  box.addEventListener('click', async (e) => {
    const t = e.target.closest('[data-theme]');
    try {
      if (t) { await ctx.callTool('sheets.set_preferences', { theme: t.dataset.theme }); document.documentElement.dataset.mode = t.dataset.theme; load(); }
      const r = e.target.closest('[data-remove]');
      if (r) dialog(document.body, { title: 'Remove from the team?', toolName: 'sheets.remove_person', submit: 'Remove', body: '<p>They are signed out everywhere. What they wrote stays.</p>', onSubmit: async () => { await ctx.callTool('sheets.remove_person', { person: r.dataset.remove }); load(); } });
      const d = e.target.closest('[data-decide]');
      if (d) { const out = await ctx.callTool('sheets.decide_approval', { approval: d.dataset.decide, approve: d.dataset.yes === '1' }); toast(out.status === 'done' ? 'Done' : out.status === 'declined' ? 'Declined' : `It did not work: ${out.result?.error ?? ''}`); load(); }
      if (e.target.closest('[data-crm-off]')) { await ctx.callTool('sheets.connect_crm', { disconnect: true }); load(); }
      if (e.target.closest('[data-export]')) { toast('Making the export…'); const out = await ctx.callTool('sheets.export_data', {}); const a = document.createElement('a'); a.href = out.file.url; a.download = out.file.name; document.body.appendChild(a); a.click(); a.remove(); }
    } catch (err) { toast(err.message); }
  });
  box.addEventListener('submit', async (e) => {
    e.preventDefault();
    const f = e.target;
    const v = Object.fromEntries(new FormData(f));
    try {
      if (f.matches('.st-add')) { await ctx.callTool('sheets.add_person', { name: v.name, email: v.email }); toast(`${v.name} can sign in now`); }
      if (f.matches('.st-crm')) { const out = await ctx.callTool('sheets.connect_crm', { url: v.url, ...(v.token ? { token: v.token } : {}) }); toast(out.check ?? 'Connected'); }
      load();
    } catch (err) { showError(f, err.message); }
  });
  return { destroy() {} };
}

// ---------- copy a shared spreadsheet after signing in ----------

function copyPage(el, ctx, shareToken) {
  el.innerHTML = '<div class="ui-page"><p class="ui-empty">Making your copy…</p></div>';
  ctx.callTool('sheets.copy_sheet', { share_token: shareToken }).then((out) => ctx.navigate(`/s/${out.id}`), (e) => { el.innerHTML = `<div class="ui-page"><p class="ui-empty">${h(e.message)}</p><a class="ui-btn is-quiet" href="#/">Your spreadsheets</a></div>`; });
  return { destroy() {} };
}
