// Charts as SVG, in kit colours (series take the --sc-1.. shades from sheets.css, so they follow light,
// dark and the colour scheme). Lines also differ by dash, so series read apart in a monochrome scheme.
import { display } from '../../lib/input.mjs';

const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
const DASH = ['', '6 4', '2 3', '10 3 2 3', '1 2', '8 6'];
const color = (i) => `var(--sc-${(i % 6) + 1})`;

function nice(max, min = 0) {
  if (max === min) { max = max === 0 ? 1 : max * 1.2; }
  const span = max - min;
  const step0 = span / 5;
  const mag = 10 ** Math.floor(Math.log10(step0));
  const step = [1, 2, 2.5, 5, 10].map((m) => m * mag).find((s) => span / s <= 6) ?? 10 * mag;
  const lo = Math.floor(min / step) * step, hi = Math.ceil(max / step) * step;
  const ticks = [];
  for (let v = lo; v <= hi + step / 2; v += step) ticks.push(Number(v.toPrecision(12)));
  return { lo, hi, ticks };
}

function short(v, nf) {
  const a = Math.abs(v);
  const cur = nf && nf.includes('$') ? '$' : '';
  if (nf && nf.includes('%')) return display(v, '0%');
  if (a >= 1e9) return `${cur}${(v / 1e9).toFixed(a >= 1e10 ? 0 : 1)}B`;
  if (a >= 1e6) return `${cur}${(v / 1e6).toFixed(a >= 1e7 ? 0 : 1)}M`;
  if (a >= 1e4) return `${cur}${(v / 1e3).toFixed(0)}k`;
  if (a >= 1e3) return `${cur}${(v / 1e3).toFixed(1)}k`;
  return cur ? display(v, '$#,##0') : String(Number(v.toPrecision(6)));
}

export function renderChart(chart, w, h) {
  const { type, data } = chart;
  const labels = data?.labels ?? [];
  const series = (data?.series ?? []).filter((s) => s.values.some((v) => v !== null));
  if (!labels.length || !series.length) return `<svg viewBox="0 0 ${w} ${h}" width="${w}" height="${h}" role="img" aria-label="Empty chart"><text x="${w / 2}" y="${h / 2}" class="ch-empty" text-anchor="middle">No numbers in ${esc(chart.range)} yet</text></svg>`;
  if (type === 'pie') return pie(chart, labels, series[0], w, h);
  const legend = series.length > 1;
  const pad = { l: 52, r: 14, t: legend ? 30 : 12, b: 30 };
  const horiz = type === 'bar';
  if (horiz) pad.l = Math.min(150, 14 + 7 * Math.max(...labels.map((l) => String(l).length)));
  const iw = Math.max(20, w - pad.l - pad.r), ih = Math.max(20, h - pad.t - pad.b);
  const stacked = !!chart.stacked && (type === 'column' || type === 'bar' || type === 'area');
  let vals = series.flatMap((s) => s.values.filter((v) => v !== null));
  if (stacked) vals = labels.map((_, i) => series.reduce((a, s) => a + Math.max(0, s.values[i] ?? 0), 0));
  const scatter = type === 'scatter';
  const { lo, hi, ticks } = nice(Math.max(0, ...vals), Math.min(0, ...vals));
  const fmt = data.format;
  const out = [`<svg viewBox="0 0 ${w} ${h}" width="${w}" height="${h}" role="img" aria-label="${esc(chart.title ?? `${type} chart`)}">`];
  const scale = (v) => (v - lo) / (hi - lo || 1);
  // Grid and value axis.
  for (const t of ticks) {
    if (horiz) { const x = pad.l + scale(t) * iw; out.push(`<line class="ch-grid" x1="${x}" x2="${x}" y1="${pad.t}" y2="${pad.t + ih}"/><text class="ch-tick" x="${x}" y="${pad.t + ih + 18}" text-anchor="middle">${esc(short(t, fmt))}</text>`); }
    else { const y = pad.t + ih - scale(t) * ih; out.push(`<line class="ch-grid" x1="${pad.l}" x2="${pad.l + iw}" y1="${y}" y2="${y}"/><text class="ch-tick" x="${pad.l - 8}" y="${y + 4}" text-anchor="end">${esc(short(t, fmt))}</text>`); }
  }
  const n = labels.length;
  const step = (horiz ? ih : iw) / n;
  const every = Math.max(1, Math.ceil(n / Math.max(1, Math.floor((horiz ? ih / 18 : iw / 64)))));
  labels.forEach((l, i) => {
    if (i % every) return;
    const lab = String(l).length > 14 ? `${String(l).slice(0, 13)}…` : String(l);
    if (horiz) out.push(`<text class="ch-tick" x="${pad.l - 8}" y="${pad.t + step * (i + 0.5) + 4}" text-anchor="end">${esc(lab)}</text>`);
    else if (!scatter) out.push(`<text class="ch-tick" x="${pad.l + step * (i + 0.5)}" y="${pad.t + ih + 18}" text-anchor="middle">${esc(lab)}</text>`);
  });
  const base = scale(Math.max(lo, 0));
  if (type === 'column' || type === 'bar') {
    const groupW = step * 0.72;
    const bw = stacked ? groupW : groupW / series.length;
    const acc = new Array(n).fill(0);
    series.forEach((s, si) => {
      s.values.forEach((v, i) => {
        if (v === null) return;
        const from = stacked ? acc[i] : 0;
        const to = from + v;
        if (stacked) acc[i] = to;
        const a = scale(Math.min(from, to)), b = scale(Math.max(from, to));
        const off = step * i + (step - groupW) / 2 + (stacked ? 0 : bw * si);
        const r = Math.min(3, bw / 3);
        if (horiz) out.push(`<rect class="ch-bar" x="${pad.l + (stacked ? a : Math.min(base, scale(v))) * iw}" y="${pad.t + off}" width="${Math.max(1, (stacked ? b - a : Math.abs(scale(v) - base)) * iw)}" height="${Math.max(1, bw - 2)}" rx="${r}" fill="${color(si)}"><title>${esc(`${s.name}, ${labels[i]}: ${display(v, fmt)}`)}</title></rect>`);
        else out.push(`<rect class="ch-bar" x="${pad.l + off}" y="${pad.t + ih - (stacked ? b : Math.max(base, scale(v))) * ih}" width="${Math.max(1, bw - 2)}" height="${Math.max(1, (stacked ? b - a : Math.abs(scale(v) - base)) * ih)}" rx="${r}" fill="${color(si)}"><title>${esc(`${s.name}, ${labels[i]}: ${display(v, fmt)}`)}</title></rect>`);
      });
    });
  } else if (scatter) {
    const xs = labels.map((l) => Number(String(l).replace(/[^0-9.-]/g, '')));
    const xr = nice(Math.max(...xs.filter(Number.isFinite)), Math.min(0, ...xs.filter(Number.isFinite)));
    for (const t of xr.ticks) { const x = pad.l + ((t - xr.lo) / (xr.hi - xr.lo || 1)) * iw; out.push(`<text class="ch-tick" x="${x}" y="${pad.t + ih + 18}" text-anchor="middle">${esc(short(t))}</text>`); }
    series.forEach((s, si) => s.values.forEach((v, i) => {
      if (v === null || !Number.isFinite(xs[i])) return;
      out.push(`<circle cx="${pad.l + ((xs[i] - xr.lo) / (xr.hi - xr.lo || 1)) * iw}" cy="${pad.t + ih - scale(v) * ih}" r="4" fill="${color(si)}" class="ch-dot"><title>${esc(`${labels[i]}, ${display(v, fmt)}`)}</title></circle>`);
    }));
  } else {
    const acc = new Array(n).fill(0);
    series.forEach((s, si) => {
      const pts = s.values.map((v, i) => {
        if (v === null) return null;
        const y = stacked ? (acc[i] += v) : v;
        return [pad.l + step * (i + 0.5), pad.t + ih - scale(y) * ih, v, i];
      });
      const segs = [];
      let cur = [];
      for (const p of pts) { if (p) cur.push(p); else if (cur.length) { segs.push(cur); cur = []; } }
      if (cur.length) segs.push(cur);
      for (const seg of segs) {
        const d = seg.map((p, i) => `${i ? 'L' : 'M'}${p[0].toFixed(1)},${p[1].toFixed(1)}`).join('');
        if (type === 'area') out.push(`<path class="ch-area" d="${d}L${seg[seg.length - 1][0].toFixed(1)},${(pad.t + ih - base * ih).toFixed(1)}L${seg[0][0].toFixed(1)},${(pad.t + ih - base * ih).toFixed(1)}Z" fill="${color(si)}"/>`);
        out.push(`<path class="ch-line" d="${d}" stroke="${color(si)}" stroke-dasharray="${DASH[si % DASH.length]}"/>`);
      }
      if (n <= 40) for (const p of pts) if (p) out.push(`<circle class="ch-dot" cx="${p[0].toFixed(1)}" cy="${p[1].toFixed(1)}" r="3" fill="${color(si)}"><title>${esc(`${s.name}, ${labels[p[3]]}: ${display(p[2], fmt)}`)}</title></circle>`);
    });
  }
  out.push(`<line class="ch-axis" x1="${pad.l}" x2="${horiz ? pad.l : pad.l + iw}" y1="${horiz ? pad.t : pad.t + ih}" y2="${pad.t + ih}"/>`);
  if (legend) {
    let x = pad.l;
    series.forEach((s, si) => {
      const name = s.name.length > 22 ? `${s.name.slice(0, 21)}…` : s.name;
      out.push(`<rect x="${x}" y="8" width="10" height="10" rx="2" fill="${color(si)}"/><text class="ch-leg" x="${x + 15}" y="17">${esc(name)}</text>`);
      x += 30 + name.length * 6.6;
    });
  }
  out.push('</svg>');
  return out.join('');
}

function pie(chart, labels, s, w, h) {
  const items = labels.map((l, i) => ({ l, v: s.values[i] })).filter((x) => x.v !== null && x.v > 0);
  const total = items.reduce((a, x) => a + x.v, 0) || 1;
  const legendW = w > 360 ? Math.min(180, w * 0.4) : 0;
  const cx = (w - legendW) / 2, cy = h / 2, R = Math.max(10, Math.min(cx, cy) - 14), r = R * 0.58;
  const out = [`<svg viewBox="0 0 ${w} ${h}" width="${w}" height="${h}" role="img" aria-label="${esc(chart.title ?? 'Pie chart')}">`];
  let a = -Math.PI / 2;
  items.forEach((x, i) => {
    const b = a + (x.v / total) * Math.PI * 2;
    const large = b - a > Math.PI ? 1 : 0;
    const p = (ang, rad) => `${(cx + rad * Math.cos(ang)).toFixed(2)},${(cy + rad * Math.sin(ang)).toFixed(2)}`;
    const d = items.length === 1 ? `M${p(0, R)}A${R},${R} 0 1 1 ${p(Math.PI, R)}A${R},${R} 0 1 1 ${p(0, R)}M${p(0, r)}A${r},${r} 0 1 0 ${p(Math.PI, r)}A${r},${r} 0 1 0 ${p(0, r)}Z` : `M${p(a, R)}A${R},${R} 0 ${large} 1 ${p(b, R)}L${p(b, r)}A${r},${r} 0 ${large} 0 ${p(a, r)}Z`;
    out.push(`<path class="ch-slice" d="${d}" fill="${color(i)}" fill-rule="evenodd"><title>${esc(`${x.l}: ${display(x.v, chart.data.format)} (${Math.round((x.v / total) * 100)}%)`)}</title></path>`);
    a = b;
  });
  out.push(`<text class="ch-total" x="${cx}" y="${cy + 5}" text-anchor="middle">${esc(short(total, chart.data.format))}</text>`);
  if (legendW) items.slice(0, 10).forEach((x, i) => out.push(`<rect x="${w - legendW}" y="${16 + i * 20}" width="10" height="10" rx="2" fill="${color(i)}"/><text class="ch-leg" x="${w - legendW + 16}" y="${25 + i * 20}">${esc(String(x.l).slice(0, 20))} · ${Math.round((x.v / total) * 100)}%</text>`));
  out.push('</svg>');
  return out.join('');
}
