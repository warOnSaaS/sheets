// Small helpers for the screens: escaping, icons, dialogs, menus, popovers and a toast. Every button they
// make names its tool (data-tool), or says data-tool="none" with a reason when it only moves around the page.

export const h = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
export const none = (why) => `data-tool="none" data-why="${h(why)}"`;
export const tool = (name) => `data-tool="${h(name)}"`;

const P = (d, extra = '') => `<svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"${extra}>${d}</svg>`;
export const I = {
  grid: P('<rect x="3.5" y="3.5" width="17" height="17" rx="3.5"/><path d="M3.5 9.2h17M3.5 14.8h17M9.5 3.5v17"/>'),
  plus: P('<path d="M12 5v14M5 12h14"/>'),
  back: P('<path d="M15 18l-6-6 6-6"/>'),
  undo: P('<path d="M9 14L4 9l5-5"/><path d="M4 9h10.5a5.5 5.5 0 010 11H11"/>'),
  redo: P('<path d="M15 14l5-5-5-5"/><path d="M20 9H9.5a5.5 5.5 0 000 11H13"/>'),
  bold: P('<path d="M7 5h6a3.5 3.5 0 010 7H7zM7 12h7a3.5 3.5 0 010 7H7z"/>'),
  italic: P('<path d="M19 4h-9M14 20H5M15 4L9 20"/>'),
  strike: P('<path d="M16 6.5C15 5 13.6 4.5 12 4.5c-2.5 0-4 1.3-4 3.2 0 4.3 8.5 2.6 8.5 7.4 0 2-1.8 3.4-4.5 3.4-2 0-3.6-.8-4.5-2.3M4 12h16"/>'),
  underline: P('<path d="M7 4v6a5 5 0 0010 0V4M5 20h14"/>'),
  color: P('<path d="M5 18L11 4h2l6 14M7.5 13h9"/>'),
  fill: P('<path d="M12 3l8 8-7 7-8-8zM5 21h14"/>'),
  left: P('<path d="M4 6h16M4 10h10M4 14h16M4 18h10"/>'),
  center: P('<path d="M4 6h16M7 10h10M4 14h16M7 18h10"/>'),
  right: P('<path d="M4 6h16M10 10h10M4 14h16M10 18h10"/>'),
  wrap: P('<path d="M4 6h16M4 12h13a3 3 0 010 6h-4M15 16l-2 2 2 2M4 18h5"/>'),
  sortUp: P('<path d="M7 4v16M3 8l4-4 4 4M14 6h7M14 12h5M14 18h3"/>'),
  sortDown: P('<path d="M7 4v16M3 16l4 4 4-4M14 6h3M14 12h5M14 18h7"/>'),
  filter: P('<path d="M4 5h16l-6 7.5V19l-4 1v-7.5z"/>'),
  chart: P('<path d="M4 20V10M10 20V4M16 20v-7M22 20H2"/>'),
  pivot: P('<rect x="3.5" y="3.5" width="17" height="17" rx="2.5"/><path d="M3.5 9h17M9 3.5v17M13 13h4M13 16h4"/>'),
  comment: P('<path d="M20 12a8 8 0 01-11.5 7.2L4 20.5l1.3-4.2A8 8 0 1120 12z"/>'),
  freeze: P('<rect x="3.5" y="3.5" width="17" height="17" rx="2.5"/><path d="M3.5 8.5h17M8.5 3.5v17"/>'),
  crm: P('<circle cx="9" cy="8" r="3.5"/><path d="M2.5 20a6.5 6.5 0 0113 0M16 4.5a3.5 3.5 0 010 7M18.5 14a6 6 0 013 6"/>'),
  refresh: P('<path d="M20 11a8 8 0 10-2.3 5.7M20 4v7h-7"/>'),
  clear: P('<path d="M4 20h16M8 16L17.5 6.5a2.1 2.1 0 013 3L11 19H7z"/>'),
  search: P('<circle cx="11" cy="11" r="6.5"/><path d="M20 20l-4-4"/>'),
  more: P('<circle cx="5" cy="12" r="1.2"/><circle cx="12" cy="12" r="1.2"/><circle cx="19" cy="12" r="1.2"/>'),
  history: P('<path d="M3.5 12a8.5 8.5 0 102.5-6"/><path d="M3 4v4.5h4.5M12 7.5V12l3 2"/>'),
  download: P('<path d="M12 4v11M7 10l5 5 5-5M5 20h14"/>'),
  share: P('<path d="M10 14a4 4 0 005.7 0l3-3a4 4 0 00-5.7-5.7l-1 1M14 10a4 4 0 00-5.7 0l-3 3a4 4 0 005.7 5.7l1-1"/>'),
  x: P('<path d="M6 6l12 12M18 6L6 18"/>'),
  check: P('<path d="M5 12.5l4.5 4.5L19 7.5"/>'),
  upload: P('<path d="M12 20V9M7 14l5-5 5 5M5 4h14"/>'),
  plug: P('<path d="M9 3v5M15 3v5M6 8h12v3a6 6 0 01-12 0zM12 17v4"/>'),
  gear: P('<circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.6 1.6 0 00.3 1.8l.1.1a2 2 0 11-2.8 2.8l-.1-.1a1.6 1.6 0 00-1.8-.3 1.6 1.6 0 00-1 1.5V21a2 2 0 11-4 0v-.1A1.6 1.6 0 009 19.4a1.6 1.6 0 00-1.8.3l-.1.1a2 2 0 11-2.8-2.8l.1-.1a1.6 1.6 0 00.3-1.8 1.6 1.6 0 00-1.5-1H3a2 2 0 110-4h.1A1.6 1.6 0 004.6 9a1.6 1.6 0 00-.3-1.8l-.1-.1a2 2 0 112.8-2.8l.1.1a1.6 1.6 0 001.8.3H9a1.6 1.6 0 001-1.5V3a2 2 0 114 0v.1a1.6 1.6 0 001 1.5 1.6 1.6 0 001.8-.3l.1-.1a2 2 0 112.8 2.8l-.1.1a1.6 1.6 0 00-.3 1.8V9a1.6 1.6 0 001.5 1H21a2 2 0 110 4h-.1a1.6 1.6 0 00-1.5 1z"/>'),
  copy: P('<rect x="8" y="8" width="12" height="12" rx="2"/><path d="M16 8V6a2 2 0 00-2-2H6a2 2 0 00-2 2v8a2 2 0 002 2h2"/>'),
  terminal: P('<path d="M4 17l6-5-6-5M12 19h8"/>'),
  chat: P('<path d="M4 5h16v11H9l-5 4z"/>'),
  caret: P('<path d="M7 10l5 5 5-5"/>', ' width="14" height="14"'),
  dollar: P('<path d="M12 3v18M16.5 7.5c-.8-1.3-2.4-2-4.5-2-2.6 0-4.2 1.3-4.2 3.1 0 4.4 8.9 2.6 8.9 7 0 1.9-1.8 3.3-4.6 3.3-2.3 0-4-.8-4.9-2.4"/>'),
  percent: P('<path d="M19 5L5 19"/><circle cx="7" cy="7" r="2.3"/><circle cx="17" cy="17" r="2.3"/>'),
  decLess: '<span class="tb-txt" aria-hidden="true">.0<sub>←</sub></span>',
  decMore: '<span class="tb-txt" aria-hidden="true">.00<sub>→</sub></span>',
  rows: P('<rect x="3.5" y="4" width="17" height="16" rx="2.5"/><path d="M3.5 12h17"/>'),
};

let toastEl = null, toastTimer = null;
export function toast(msg, root = document.body) {
  if (!toastEl || !toastEl.isConnected) { toastEl = document.createElement('div'); toastEl.className = 'ui-toast'; toastEl.setAttribute('role', 'status'); toastEl.setAttribute('aria-live', 'polite'); root.appendChild(toastEl); }
  toastEl.textContent = msg;
  toastEl.classList.add('is-on');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => toastEl.classList.remove('is-on'), 2600);
}

// A dialog: a title, fields, and a form whose submit runs the tool named. onSubmit gets the form values;
// returning false keeps it open. Escape or Cancel closes it.
export function dialog(root, { title, body, toolName, submit = 'Save', cancel = 'Cancel', onSubmit, wide = false, footer = '', plain = false }) {
  const d = document.createElement('dialog');
  d.className = `ui-dialog sh-dialog${wide ? ' is-wide' : ''}`;
  // A plain dialog holds its own forms (comments, history), so it is not a form itself.
  const wrap = plain ? 'div' : 'form';
  d.innerHTML = `<${wrap} ${plain ? '' : 'method="dialog" '}class="sh-dform" ${plain ? '' : toolName ? tool(toolName) : none('closes the dialog')}><h3>${h(title)}</h3><div class="sh-dbody">${body}</div><div class="ui-dialog-a">${footer}${cancel ? `<button type="button" class="ui-btn is-quiet" data-close ${none('closes the dialog')}>${h(cancel)}</button>` : ''}${submit ? (plain ? `<button type="button" class="ui-btn is-accent" data-close ${none('closes the dialog')}>${h(submit)}</button>` : `<button type="submit" class="ui-btn is-accent">${h(submit)}</button>`) : ''}</div></${wrap}>`;
  root.appendChild(d);
  const form = d.querySelector('.sh-dform');
  d.querySelectorAll('[data-close]').forEach((b) => b.addEventListener('click', () => d.close()));
  form.addEventListener('submit', async (e) => {
    e.preventDefault();
    if (!onSubmit) return d.close();
    const btn = form.querySelector('[type=submit]');
    if (btn) btn.disabled = true;
    try {
      const keep = await onSubmit(Object.fromEntries(new FormData(form)), form, d);
      if (keep !== false) d.close();
    } catch (err) {
      showError(form, err.message);
    } finally { if (btn) btn.disabled = false; }
  });
  d.addEventListener('close', () => d.remove());
  d.showModal();
  const first = d.querySelector('input:not([type=hidden]),select,textarea');
  if (first && matchMedia('(pointer:fine)').matches) first.focus();
  return d;
}

export function showError(form, msg) {
  let p = form.querySelector('.sh-err');
  if (!p) { p = document.createElement('p'); p.className = 'sh-err'; p.setAttribute('role', 'alert'); form.querySelector('.ui-dialog-a, .sh-pop-a')?.before(p) ?? form.append(p); }
  p.textContent = msg;
}

// A menu at an element (or a point). items: { label, tool, run, icon, hint, danger, checked, sep, disabled }.
let openMenu = null;
export function menu(root, at, items, { label = 'Menu' } = {}) {
  closeMenu();
  const m = document.createElement('div');
  m.className = 'sh-menu';
  m.setAttribute('role', 'menu');
  m.setAttribute('aria-label', label);
  m.innerHTML = items.filter(Boolean).map((it, i) => it.sep ? '<hr>' : it.head ? `<p class="sh-menu-h">${h(it.head)}</p>` : `<button type="button" role="menuitem" data-i="${i}" ${it.tool ? tool(it.tool) : none(it.why ?? 'moves around the page')}${it.disabled ? ' disabled' : ''} class="${it.danger ? 'is-danger' : ''}${it.checked ? ' is-on' : ''}">${it.icon ?? '<span class="sh-mi"></span>'}<span>${h(it.label)}</span>${it.hint ? `<kbd class="ui-kbd">${h(it.hint)}</kbd>` : ''}</button>`).join('');
  root.appendChild(m);
  const list = items.filter(Boolean);
  const r = at instanceof Element ? at.getBoundingClientRect() : { left: at.x, right: at.x, top: at.y, bottom: at.y };
  const mw = m.offsetWidth, mh = m.offsetHeight;
  let x = r.left, y = r.bottom + 4;
  if (x + mw > innerWidth - 8) x = Math.max(8, (at instanceof Element ? r.right : r.left) - mw);
  if (y + mh > innerHeight - 8) y = Math.max(8, r.top - mh - 4);
  m.style.left = `${x}px`;
  m.style.top = `${y}px`;
  m.addEventListener('click', (e) => {
    const b = e.target.closest('[data-i]');
    if (!b) return;
    const it = list[Number(b.dataset.i)];
    closeMenu();
    it?.run?.();
  });
  m.addEventListener('keydown', (e) => {
    const bs = [...m.querySelectorAll('button:not([disabled])')];
    const i = bs.indexOf(document.activeElement);
    if (e.key === 'ArrowDown') { e.preventDefault(); bs[(i + 1) % bs.length]?.focus(); }
    if (e.key === 'ArrowUp') { e.preventDefault(); bs[(i - 1 + bs.length) % bs.length]?.focus(); }
    if (e.key === 'Escape') { e.preventDefault(); closeMenu(); }
  });
  setTimeout(() => { document.addEventListener('pointerdown', outside, true); }, 0);
  openMenu = m;
  m.querySelector('button')?.focus({ preventScroll: true });
  return m;
}
function outside(e) { if (openMenu && !openMenu.contains(e.target)) closeMenu(); }
export function closeMenu() { if (openMenu) { openMenu.remove(); openMenu = null; document.removeEventListener('pointerdown', outside, true); } }

// A small panel anchored to a point, for comments, filters and colour picks.
let openPop = null;
export function popover(root, at, html, { className = '' } = {}) {
  closePop();
  const p = document.createElement('div');
  p.className = `sh-pop ${className}`;
  p.innerHTML = html;
  root.appendChild(p);
  const pw = p.offsetWidth, ph = p.offsetHeight;
  let x = at.x, y = at.y;
  if (x + pw > innerWidth - 8) x = Math.max(8, innerWidth - pw - 8);
  if (y + ph > innerHeight - 8) y = Math.max(8, at.y0 !== undefined ? at.y0 - ph - 6 : innerHeight - ph - 8);
  p.style.left = `${x}px`;
  p.style.top = `${y}px`;
  setTimeout(() => document.addEventListener('pointerdown', popOutside, true), 0);
  openPop = p;
  return p;
}
function popOutside(e) { if (openPop && !openPop.contains(e.target) && !e.target.closest?.('.sh-menu')) closePop(); }
export function closePop() { if (openPop) { openPop.remove(); openPop = null; document.removeEventListener('pointerdown', popOutside, true); } }

export const initials = (name) => String(name ?? '?').split(/\s+/).filter(Boolean).slice(0, 2).map((w) => w[0].toUpperCase()).join('') || '?';
export const ago = (iso) => {
  const s = (Date.now() - Date.parse(iso)) / 1000;
  if (!Number.isFinite(s)) return '';
  if (s < 60) return 'just now';
  if (s < 3600) return `${Math.floor(s / 60)} min ago`;
  if (s < 86400) return `${Math.floor(s / 3600)} h ago`;
  if (s < 86400 * 7) return `${Math.floor(s / 86400)} d ago`;
  return new Date(iso).toLocaleDateString(undefined, { month: 'short', day: 'numeric', year: new Date(iso).getFullYear() === new Date().getFullYear() ? undefined : 'numeric' });
};

export function copyText(s) {
  if (navigator.clipboard && window.isSecureContext) return navigator.clipboard.writeText(s);
  const a = document.createElement('textarea');
  a.value = s;
  document.body.appendChild(a);
  a.select();
  try { document.execCommand('copy'); } finally { a.remove(); }
  return Promise.resolve();
}

export const fileToBase64 = (file) => new Promise((ok, bad) => {
  const r = new FileReader();
  r.onload = () => ok(String(r.result).split(',')[1] ?? '');
  r.onerror = () => bad(new Error('That file could not be read.'));
  r.readAsDataURL(file);
});
