// Records from the wOS CRM in a sheet, kept refreshable. Sheets never reads the CRM's tables: it calls
// the CRM's own tools. Inside the suite that is call.callTool('crm.query'); standalone it is the CRM's
// REST address (POST <crm>/api/tools/crm.query), set by sheets.connect_crm or CRM_URL and CRM_TOKEN.
import { BookError, setRaw, setFmt, tabMap, cellMap, fmtMap, needTab, addTab, setTabProp, checkTabName, usedRange } from './book.mjs';
import { textRaw, parseInput } from './input.mjs';
import { key, unkey } from './a1.mjs';

export const CRM_KINDS = {
  deals: { object: 'Opportunity', fields: ['Name', 'Account.Name', 'StageName', 'Amount', 'CloseDate', 'Owner.Name'] },
  contacts: { object: 'Contact', fields: ['Name', 'Title', 'Email', 'Phone', 'Account.Name', 'Owner.Name'] },
  leads: { object: 'Lead', fields: ['Name', 'Company', 'Email', 'Phone', 'Status', 'Owner.Name'] },
  organizations: { object: 'Account', fields: ['Name', 'Industry', 'Website', 'Phone', 'Owner.Name'] },
  activities: { object: 'Task', fields: ['Subject', 'Type', 'ActivityDate', 'Owner.Name'] },
};

const NICE = { Name: 'Name', 'Account.Name': 'Organization', StageName: 'Stage', Amount: 'Amount', CloseDate: 'Close date', 'Owner.Name': 'Owner', Title: 'Title', Email: 'Email', Phone: 'Phone', Industry: 'Industry', Website: 'Website', Subject: 'Subject', Type: 'Type', ActivityDate: 'Date', Company: 'Company', Status: 'Status', LeadStatus: 'Lead status' };

export function soqlFor({ kind = null, fields = null, where = null, order_by = null, limit = 2000, soql = null }) {
  if (soql) return /\blimit\s+\d+/i.test(soql) ? soql : `${soql} LIMIT ${limit}`;
  const k = CRM_KINDS[kind];
  if (!k) throw new BookError(`kind is one of ${Object.keys(CRM_KINDS).join(', ')}, or give soql.`);
  const f = fields?.length ? fields : k.fields;
  return `SELECT ${f.join(', ')} FROM ${k.object}${where ? ` WHERE ${where}` : ''}${order_by ? ` ORDER BY ${order_by}` : ''} LIMIT ${limit}`;
}

// Calls a CRM tool, in the suite or over REST. Returns the tool's data part.
export async function callCrm(app, me, call, tool, input) {
  if (call?.callTool) {
    let out;
    try { out = await call.callTool(tool, input); } catch (e) {
      if (e.code === 'no_tool' || /no_tool|not on|no tool/i.test(e.message)) throw new BookError('The CRM is off for this team. Turn it on in Settings, Apps, then try again.', 409, 'crm_off');
      throw new BookError(`The CRM said: ${e.message}`, 502, 'crm');
    }
    return out?.data ?? out?.result?.data ?? out;
  }
  const conn = await crmConnection(app, me.team_id);
  if (!conn) throw new BookError('No CRM is connected. An admin connects one with sheets.connect_crm (its address, and an access token from crm.create_access_token), or the server sets CRM_URL.', 409, 'crm_missing');
  let res;
  try {
    res = await fetch(`${conn.url.replace(/\/$/, '')}/api/tools/${encodeURIComponent(tool)}`, {
      method: 'POST',
      headers: { 'content-type': 'application/json', accept: 'application/json', ...(conn.token ? { authorization: `Bearer ${conn.token}` } : {}) },
      body: JSON.stringify(input),
      signal: AbortSignal.timeout(20000),
    });
  } catch (e) { throw new BookError(`The CRM at ${conn.url} could not be reached (${e.message}).`, 502, 'crm'); }
  const body = await res.json().catch(() => ({}));
  if (!res.ok) throw new BookError(`The CRM said: ${typeof body.error === 'string' ? body.error : body.error?.message ?? res.status}`, res.status === 401 || res.status === 403 ? 403 : 502, 'crm');
  return body.data ?? body.result?.data ?? body.result;
}

export async function crmConnection(app, teamId) {
  const saved = await app.team.setting(teamId, 'crm');
  if (saved?.url) return { url: saved.url, token: saved.token ? app.unseal(saved.token) : null, from: 'settings' };
  if (app.env.CRM_URL) return { url: app.env.CRM_URL, token: app.env.CRM_TOKEN || null, from: 'server' };
  return null;
}

// Writes CRM records as a table: a header row, then one row per record, with the record id last so a
// refresh can line them up. Leftover rows from a longer previous pull are cleared.
export function writeTable(doc, tabId, data) {
  const cols = data.columns ?? Object.keys(data.records?.[0] ?? {}).filter((k) => k !== 'id');
  const records = data.records ?? [];
  const header = [...cols.map((c) => NICE[c] ?? c), 'CRM id'];
  const m = cellMap(doc, tabId);
  const width = header.length;
  const prev = usedRange(doc, tabId);
  doc.transact(() => {
    const t = tabMap(doc, tabId);
    if ((t.get('rows') ?? 1000) < records.length + 10) t.set('rows', records.length + 50);
    if ((t.get('cols') ?? 26) < width + 2) t.set('cols', width + 5);
    header.forEach((h, c) => { setRaw(doc, tabId, 0, c, textRaw(h)); setFmt(doc, tabId, 0, c, { b: true }); });
    records.forEach((rec, i) => {
      cols.forEach((c, j) => setRaw(doc, tabId, i + 1, j, cellValue(rec[c])));
      setRaw(doc, tabId, i + 1, cols.length, textRaw(rec.id ?? ''));
    });
    // Clear what the last pull left below and inside the table's columns.
    if (prev) for (const k of [...m.keys()]) { const [r, c] = unkey(k); if (c < width && r > records.length) m.delete(k); }
    // Money and dates read as such.
    cols.forEach((c, j) => {
      const nf = /amount|revenue|value|price/i.test(c) ? '$#,##0' : /date/i.test(c) ? 'yyyy-mm-dd' : null;
      if (nf) for (let i = 1; i <= records.length; i++) setFmt(doc, tabId, i, j, { nf });
    });
    setTabProp(doc, tabId, 'fr', 1);
  });
  return { columns: header, rows: records.length };
}

function cellValue(v) {
  if (v === null || v === undefined) return null;
  if (typeof v === 'number' || typeof v === 'boolean') return v;
  const s = String(v);
  // ISO dates become real dates; everything else stays text exactly.
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return parseInput(s).raw;
  return textRaw(s);
}

// Fetches first (outside any change to the workbook, so a slow CRM never holds it), then writes.
export async function fetchCrm(app, me, call, source) {
  const soql = soqlFor(source);
  const data = await callCrm(app, me, call, 'crm.query', { soql });
  if (!data || !Array.isArray(data.records)) throw new BookError('The CRM did not send records back. Is it a wOS CRM?', 502, 'crm');
  if (data.grouped) throw new BookError('Use a query without GROUP BY or COUNT(): Sheets pulls the records, then a pivot or SUMIFS adds them up and stays live.');
  return { data, soql };
}

export function writeCrmTab(doc, { tab = null, name = null, source, data, soql }) {
  let id;
  if (tab) id = needTab(doc, tab).id;
  else id = addTab(doc, { name: checkTabName(doc, name || `CRM ${source.kind ?? data.kind ?? 'records'}`) });
  const out = writeTable(doc, id, data);
  setTabProp(doc, id, 'link', JSON.stringify({ kind: 'crm', tool: 'crm.query', input: { soql }, source, refreshed_at: new Date().toISOString(), total: data.totalSize ?? out.rows }));
  return { tab: needTab(doc, id).name, ...out, total_in_crm: data.totalSize ?? out.rows, soql };
}
