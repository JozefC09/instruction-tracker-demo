// "Copy chart" (spec F28): draws a dashboard bar chart as a PNG for slides and email.
// The chart is redrawn from its data on a canvas (not screenshotted), always in the light palette on white,
// because that's what reads well pasted into PowerPoint, Word or Outlook.

const LIGHT = {
  bg: '#ffffff',
  text: '#0b0b0b',
  secondary: '#52514e',
  muted: '#77766f',
  grid: '#e2e1dc',
  'series-1': '#2a78d6',
  'series-2': '#eb6834',
  'series-3': '#1baf7a',
  'series-none': '#a3a29b',
};

const resolve = (color) => {
  const m = String(color).match(/var\(--([\w-]+)\)/);
  return m ? LIGHT[m[1]] ?? LIGHT['series-none'] : color;
};

function roundedBar(ctx, x, y, w, h, r) {
  // Square at the baseline, 4px rounded at the data end.
  r = Math.min(r, w, h / 2);
  ctx.beginPath();
  ctx.moveTo(x, y);
  ctx.lineTo(x + w - r, y);
  ctx.arcTo(x + w, y, x + w, y + r, r);
  ctx.lineTo(x + w, y + h - r);
  ctx.arcTo(x + w, y + h, x + w - r, y + h, r);
  ctx.lineTo(x, y + h);
  ctx.closePath();
  ctx.fill();
}

/**
 * chart = { title, subtitle, rows: [{ label, value, color }], unit, footer }
 * Returns a canvas at 2× resolution (crisp on slides).
 */
export function drawBarChart(chart, { width = 760, scale = 2 } = {}) {
  const font = getComputedStyle(document.body).fontFamily || 'Segoe UI, Arial, sans-serif';
  const pad = 28, rowH = 34, barH = 16;
  const rows = chart.rows.length ? chart.rows : [{ label: 'No sessions match', value: 0, color: LIGHT['series-none'] }];

  const measure = document.createElement('canvas').getContext('2d');
  measure.font = `14px ${font}`;
  const labelW = Math.min(280, Math.max(80, ...rows.map((r) => measure.measureText(r.label).width))) + 16;
  const max = Math.max(1, ...rows.map((r) => r.value));

  const headH = chart.subtitle ? 74 : 54;
  const height = pad + headH + rows.length * rowH + 40;
  const canvas = document.createElement('canvas');
  canvas.width = width * scale;
  canvas.height = height * scale;
  const ctx = canvas.getContext('2d');
  ctx.scale(scale, scale);
  ctx.fillStyle = LIGHT.bg;
  ctx.fillRect(0, 0, width, height);
  ctx.textBaseline = 'middle';

  ctx.fillStyle = LIGHT.text;
  ctx.font = `600 19px ${font}`;
  ctx.fillText(chart.title, pad, pad + 10);
  if (chart.subtitle) {
    ctx.fillStyle = LIGHT.muted;
    ctx.font = `13px ${font}`;
    ctx.fillText(fit(ctx, chart.subtitle, width - pad * 2), pad, pad + 36);
  }

  const x0 = pad + labelW;
  const valueRoom = 56;
  const plotW = width - x0 - pad - valueRoom;
  const top = pad + headH;
  // Recessive baseline.
  ctx.strokeStyle = LIGHT.grid;
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.moveTo(x0 + 0.5, top - 6);
  ctx.lineTo(x0 + 0.5, top + rows.length * rowH - 6);
  ctx.stroke();

  rows.forEach((r, i) => {
    const cy = top + i * rowH + rowH / 2 - 6;
    ctx.font = `14px ${font}`;
    ctx.fillStyle = LIGHT.secondary;
    ctx.textAlign = 'right';
    ctx.fillText(fit(ctx, r.label, labelW - 16), x0 - 12, cy);
    ctx.textAlign = 'left';
    const w = (r.value / max) * plotW;
    if (w > 0) {
      ctx.fillStyle = resolve(r.color);
      roundedBar(ctx, x0 + 1, cy - barH / 2, w, barH, 4);
    }
    ctx.fillStyle = LIGHT.text;
    ctx.font = `600 14px ${font}`;
    ctx.fillText(String(r.value), x0 + 1 + w + 8, cy);
  });

  ctx.fillStyle = LIGHT.muted;
  ctx.font = `11px ${font}`;
  ctx.fillText(fit(ctx, chart.footer || '', width - pad * 2), pad, height - 18);
  return canvas;
}

function fit(ctx, text, maxW) {
  if (ctx.measureText(text).width <= maxW) return text;
  let s = text;
  while (s.length > 1 && ctx.measureText(`${s}…`).width > maxW) s = s.slice(0, -1);
  return `${s}…`;
}

const toBlob = (canvas) => new Promise((res, rej) => canvas.toBlob((b) => (b ? res(b) : rej(new Error('Couldn’t draw the chart.'))), 'image/png'));

/**
 * Copies the chart to the clipboard as a PNG. Where the browser can't (older Firefox, or a page not served
 * over https), it downloads the PNG instead. Returns 'copied' or 'downloaded'.
 */
export async function copyChartPng(chart, filename) {
  const canvas = drawBarChart(chart);
  if (navigator.clipboard?.write && window.ClipboardItem) {
    try {
      // Passing the promise (not the blob) keeps Safari's "user gesture" requirement satisfied.
      await navigator.clipboard.write([new ClipboardItem({ 'image/png': toBlob(canvas) })]);
      return 'copied';
    } catch { /* fall through to download */ }
  }
  const blob = await toBlob(canvas);
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  document.body.append(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
  return 'downloaded';
}
