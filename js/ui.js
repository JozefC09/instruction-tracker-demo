// Small DOM helpers shared by the views.

export const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

export const $ = (sel, root = document) => root.querySelector(sel);
export const $$ = (sel, root = document) => [...root.querySelectorAll(sel)];

export function toast(msg, kind = '') {
  const el = document.createElement('div');
  el.className = `toast ${kind}`;
  el.textContent = msg;
  $('#toasts').append(el);
  setTimeout(() => el.remove(), kind === 'error' ? 8000 : 4000);
}

/** Opens the shared <dialog> with HTML content. Resolves with the value passed to close(). */
export function openDialog(html, onMount) {
  const dlg = $('#dialog');
  dlg.innerHTML = html;
  return new Promise((resolve) => {
    const done = (v) => {
      dlg.removeEventListener('close', onClose);
      if (dlg.open) dlg.close();
      resolve(v);
    };
    const onClose = () => done(null);
    dlg.addEventListener('close', onClose);
    dlg.showModal();
    onMount?.(dlg, done);
  });
}

export async function confirmDialog(title, body, okLabel = 'OK', danger = false) {
  return openDialog(`
    <div class="dlg-head"><h2>${esc(title)}</h2></div>
    <div class="dlg-body">${body}</div>
    <div class="dlg-foot"><button class="btn" data-x="no">Cancel</button><button class="btn ${danger ? 'danger' : 'primary'}" data-x="yes">${esc(okLabel)}</button></div>`,
  (dlg, done) => {
    dlg.querySelector('[data-x=no]').onclick = () => done(false);
    dlg.querySelector('[data-x=yes]').onclick = () => done(true);
  });
}

export async function promptDialog(title, label, placeholder = '') {
  return openDialog(`
    <form method="dialog">
      <div class="dlg-head"><h2>${esc(title)}</h2></div>
      <div class="dlg-body"><label class="field">${esc(label)}<textarea name="v" placeholder="${esc(placeholder)}"></textarea></label></div>
      <div class="dlg-foot"><button class="btn" type="button" data-x="no">Cancel</button><button class="btn primary" type="submit">OK</button></div>
    </form>`,
  (dlg, done) => {
    dlg.querySelector('[data-x=no]').onclick = () => done(null);
    dlg.querySelector('form').onsubmit = (e) => { e.preventDefault(); done(dlg.querySelector('textarea').value); };
  });
}

export function download(filename, content, type) {
  const blob = content instanceof Blob ? content : new Blob([content], { type });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  document.body.append(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
}

/** Writes an .xlsx from { sheetName: array-of-arrays }. */
export function downloadXlsx(filename, sheets) {
  const XLSX = window.XLSX;
  if (!XLSX) return toast('The Excel export library did not load.', 'error');
  const wb = XLSX.utils.book_new();
  for (const [name, aoa] of Object.entries(sheets)) XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), name.slice(0, 31));
  XLSX.writeFile(wb, filename);
}

const tip = () => $('#tooltip');
export function showTip(html, ev) {
  const t = tip();
  t.innerHTML = html;
  t.hidden = false;
  const r = ev.currentTarget?.getBoundingClientRect?.();
  const x = ev.clientX || (r ? r.left + r.width / 2 : 0);
  const y = ev.clientY || (r ? r.top : 0);
  const w = t.offsetWidth, hgt = t.offsetHeight;
  t.style.left = `${Math.min(window.innerWidth - w - 8, Math.max(8, x + 12))}px`;
  t.style.top = `${y + 16 + hgt > window.innerHeight ? y - hgt - 10 : y + 16}px`;
}
export function hideTip() {
  tip().hidden = true;
}

/** Wires hover + keyboard-focus tooltips on elements with data-tip (HTML-escaped already). */
export function bindTips(root) {
  for (const el of $$('[data-tip]', root)) {
    const show = (e) => showTip(el.dataset.tip, e);
    el.addEventListener('mouseenter', show);
    el.addEventListener('mousemove', show);
    el.addEventListener('focus', show);
    el.addEventListener('mouseleave', hideTip);
    el.addEventListener('blur', hideTip);
  }
}

export const fmtDate = (iso, opts = { month: 'short', day: 'numeric', year: 'numeric' }) => {
  if (!iso) return '';
  return new Date(`${iso}T12:00:00`).toLocaleDateString('en-US', opts);
};
export const fmtDateShort = (iso) => fmtDate(iso, { weekday: 'short', month: 'short', day: 'numeric' });

export function timeAgo(ts) {
  const s = Math.round((Date.now() - ts) / 1000);
  if (s < 5) return 'just now';
  if (s < 60) return `${s}s ago`;
  return `${Math.round(s / 60)} min ago`;
}
