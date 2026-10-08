export const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
const json = (o) => JSON.stringify(o).replace(/</g, '\\u003c');

// "Look freely, sign in to use": with the warOnSaaS account, its prompt.js asks for sign-in at the moment of
// an action (any data-tool press) and never walls off a page. Self-hosted installs on GitHub or email skip it.
export const promptTag = (account, signedIn) => (account ? `<script src="${esc((process.env.WOS_ACCOUNT_URL || 'https://account.waronsaas.com').replace(/\/$/, ''))}/prompt.js" defer data-signed-in="${signedIn ? 'true' : 'false'}" data-app="Sheets" data-signin="/auth/waronsaas"></script>` : '');

const head = ({ title, theme = 'auto', v = '' }) => `<!doctype html><html lang="en" data-scheme="ops" data-mode="${esc(theme)}" data-shape="soft" data-type="grotesk" data-surface="bordered"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover,interactive-widget=resizes-content">
<title>${esc(title)}</title><meta name="theme-color" content="#0b0b0b">
<link rel="manifest" href="/manifest.webmanifest"><link rel="icon" href="/icon.svg" type="image/svg+xml"><link rel="apple-touch-icon" href="/icon-192.png">
<link rel="preload" href="/ui/fonts/geist.woff2" as="font" type="font/woff2" crossorigin>
<link rel="stylesheet" href="/ui/src/ui.css${v}"><link rel="stylesheet" href="/ui/src/tokens.css${v}"><link rel="stylesheet" href="/app/sheets.css${v}">`;

// The page every signed-in screen lives in, and the view-only page behind a share link. Screens are drawn
// by the bundled public/app/page.js, which gets all of its data from the tools (/api/tools/*).
export function appShell({ title = 'Sheets', theme = 'auto', version = '', view = null, me = null, account = false }) {
  const v = version ? `?v=${esc(version)}` : '';
  return `${head({ title, theme, v })}<meta name="robots" content="noindex">
</head><body class="sheets-page"><div id="app" class="sheets-root" aria-busy="true"></div>
<script>window.SHEETS=${json({ version, view, me })}</script>
<script type="module" src="/app/page.js${v}"></script>${promptTag(account, !!me)}</body></html>`;
}

// What a visitor sees before signing in: the examples to open, and both ways to run it.
export function landing({ examples = [], version = '', openSignup = false, account = false }) {
  const v = version ? `?v=${esc(version)}` : '';
  const cards = examples.map((e) => `<a class="ex-card" href="/v/${esc(e.token)}"><span class="ex-ic" aria-hidden="true">${GRID}</span><span class="ex-t">${esc(e.title)}</span><span class="ex-a">${esc(e.about)}</span><span class="ex-go">Open, view only</span></a>`).join('');
  return `${head({ title: 'wOS Sheets: spreadsheets your team owns', v })}
<meta name="description" content="Spreadsheets with formulas, tabs, charts, comments and live editing, that your own AI works in over MCP. Self-host free, or host with us.">
</head><body class="land">
<header class="land-top"><a class="land-brand" href="/"><span class="land-mark" aria-hidden="true">${GRID}</span>Sheets</a><nav><a class="ui-btn is-quiet is-sm" href="https://github.com/warOnSaaS/sheets">Source</a><a class="ui-btn is-accent is-sm" href="/login">Sign in</a></nav></header>
<main class="land-main">
  <section class="land-hero">
    <p class="land-kicker">wOS Sheets · open source</p>
    <h1>Spreadsheets your team owns. Your own AI works in them.</h1>
    <p class="land-lede">Formulas, tabs, charts, comments and editing together, like the spreadsheet you know. Connect Claude, ChatGPT or Codex and it reads and writes the sheet for you, with your subscription, not ours. No AI credits, ever.</p>
    <div class="land-cta"><a class="ui-btn is-accent is-lg" href="/login">${openSignup ? 'Sign in to make your own' : 'Sign in'}</a><a class="ui-btn is-quiet is-lg" href="#examples">Look at the examples</a></div>
  </section>
  <section id="examples" class="land-sec"><h2 class="ui-label">Examples, no sign-in needed</h2><div class="ex-grid">${cards || '<p class="ui-hint">No examples on this server.</p>'}</div></section>
  <section class="land-sec"><h2 class="ui-label">Two ways to run it</h2>
    <div class="land-host">
      <div class="host-card"><h3>Host it yourself, free</h3><p>One <code>docker compose up</code> runs Sheets and Postgres on any server. One person can run it with Node and a SQLite file. No licence check, ever. AGPL-3.0.</p><a class="ui-btn is-quiet" href="https://github.com/warOnSaaS/sheets#host-it-yourself">How to host it</a></div>
      <div class="host-card"><h3>Host it with us</h3><p>We run it and charge what it costs us times two, shown openly. Move to your own server any time with one export.</p><a class="ui-btn is-quiet" href="/login">${openSignup ? 'Start free' : 'Sign in'}</a></div>
    </div>
  </section>
  <section class="land-sec land-ai"><h2 class="ui-label">Works with the AI you already pay for</h2><p>Every button in Sheets is also a tool your AI can use over MCP: read a range and get exact values and formulas, write cells, add a chart, build a pivot, pull deals from the CRM. ${'<b>'}Connect it once from the Connect your AI page after you sign in.${'</b>'}</p></section>
</main>
${promptTag(account, false)}<footer class="land-foot"><span>wOS Sheets ${esc(version)}</span><a href="https://github.com/warOnSaaS/sheets">GitHub</a><a href="https://waronsaas.com">warOnSaaS</a></footer>
</body></html>`;
}

export const GRID = '<svg viewBox="0 0 24 24" width="20" height="20" aria-hidden="true"><rect x="3.5" y="3.5" width="17" height="17" rx="3.5" fill="none" stroke="currentColor" stroke-width="1.7"/><path d="M3.5 9.2h17M3.5 14.8h17M9.5 3.5v17" stroke="currentColor" stroke-width="1.7"/></svg>';
