// The standalone page: makes the screen context that the suite would otherwise give, and mounts Sheets.
// callTool posts to /api/tools/<name>; live changes come over a WebSocket at /ws, and while it is down a
// "resync" every three seconds makes the open spreadsheet catch up with sheets.sync. The path lives after #.
import { mount } from './sheets.mjs';

const boot = window.SHEETS ?? {};
const view = boot.view ?? null;
const listeners = new Map();
const emit = (name, e) => { for (const fn of listeners.get(name) ?? []) { try { fn(e); } catch (err) { console.error(err); } } };

const ctx = {
  standalone: true,
  view,
  me: boot.me && boot.me.id ? boot.me : null,
  signedIn: !!boot.me,
  live: false,
  path: location.hash.replace(/^#/, '') || '/',
  async callTool(name, input = {}) {
    const r = await fetch(`/api/tools/${name}`, { method: 'POST', headers: { 'content-type': 'application/json', ...(view ? { 'x-share-token': view.token } : {}) }, body: JSON.stringify(input), credentials: 'same-origin' });
    let j = {};
    try { j = await r.json(); } catch {}
    if (r.status === 401 && !view) { location.href = `/login?next=${encodeURIComponent(location.pathname + location.hash)}`; throw new Error('Signed out'); }
    if (r.status === 202 && j.pending) return j;
    if (!r.ok || j.error) {
      if (view && r.status === 403 && window.wosAccount) window.wosAccount.prompt('edit this spreadsheet');
      throw new Error(j.error?.message || 'Something went wrong. Try again.');
    }
    return j.result;
  },
  on(name, fn) {
    if (!listeners.has(name)) listeners.set(name, new Set());
    listeners.get(name).add(fn);
    return () => listeners.get(name)?.delete(fn);
  },
  navigate(path) { location.hash = path; },
  toast() {},
};

const app = mount(document.getElementById('app'), ctx);
document.getElementById('app').removeAttribute('aria-busy');
window.addEventListener('hashchange', () => { ctx.path = location.hash.replace(/^#/, '') || '/'; app.update(ctx.path); });

const live = {
  ws: null, poller: null, backoff: 1000,
  connect() {
    let ws;
    try { ws = new WebSocket(`${location.protocol === 'https:' ? 'wss' : 'ws'}://${location.host}/ws${view ? `?share=${encodeURIComponent(view.token)}` : ''}`); } catch { return this.poll(); }
    this.ws = ws;
    const giveUp = setTimeout(() => { if (ws.readyState !== 1) this.poll(); }, 5000);
    ws.onopen = () => { clearTimeout(giveUp); this.backoff = 1000; ctx.live = true; clearInterval(this.poller); this.poller = null; emit('resync', { type: 'resync' }); };
    ws.onmessage = (m) => { let e; try { e = JSON.parse(m.data); } catch { return; } if (e.type && e.type !== 'hello' && e.type !== 'pong') emit(e.type, e); };
    ws.onclose = () => { clearTimeout(giveUp); ctx.live = false; this.poll(); setTimeout(() => this.connect(), this.backoff); this.backoff = Math.min(this.backoff * 2, 30000); };
  },
  poll() { if (!this.poller) this.poller = setInterval(() => { if (document.visibilityState === 'visible') emit('resync', { type: 'resync' }); }, 3000); },
};
live.connect();
setInterval(() => { if (live.ws?.readyState === 1) live.ws.send('ping'); }, 25000);
