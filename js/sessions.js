// Sessions grid with stage views (spec §4.2), the side editor (§4.3), conflict dialog (§5) and the tours grid.
import { app } from './app.js';
import {
  FIELDS, STAGES, parseTime, formatTimeRange, fmtMinutes, parsePrep, formatPrep, normalizeCourse, normalizeLocation,
  suggestFromCourse, validate, needsAttention, addDays, termYearFromSheet, buildIcs, termForDate,
} from './model.js';
import { $, $$, esc, toast, openDialog, confirmDialog, download, downloadXlsx, fmtDateShort } from './ui.js';

// ---------- "Add to my calendar" bookkeeping (per viewer, so it lives in this browser) ----------
const ICS_KEY = 'tracker.ics.v1';
export function icsState() {
  try { return JSON.parse(localStorage.getItem(ICS_KEY) || '{}'); } catch { return {}; }
}
export function markIcsExported(records) {
  const s = icsState();
  for (const r of records) if (r.id) s[r.id] = r.version;
  try { localStorage.setItem(ICS_KEY, JSON.stringify(s)); } catch { /* not remembered */ }
}

// ---------- Grid ----------

const COLS = {
  date: { label: 'Date', get: (r) => (r.fields.date ? fmtDateShort(r.fields.date) : '—') },
  time: { label: 'Time', get: (r) => r.timeDisplay },
  course: { label: 'Course #', get: (r) => r.fields.course },
  faculty: { label: 'Faculty', get: (r) => r.fields.faculty },
  college: { label: 'College/School', get: (r) => r.fields.college },
  studentType: { label: 'Program', get: (r) => r.fields.studentType },
  librarian: { label: 'Librarian', get: (r) => r.fields.librarian },
  sessionType: { label: 'Session type', get: (r) => r.fields.sessionType },
  format: { label: 'Format', get: (r) => r.fields.format },
  location: { label: 'Location', get: (r) => r.fields.location },
  room: { label: 'Room', get: (r) => r.fields.room },
  expected: { label: 'Expected', get: (r) => r.fields.expected ?? '', num: true },
  actual: { label: 'Actual', get: (r) => r.fields.actual ?? '', num: true },
  assessment: { label: 'Assessment', get: (r) => r.fields.assessment },
  prep: { label: 'Prep time', get: (r) => r.fields.prep },
  notes: { label: 'Notes', get: (r) => r.fields.notes },
};
const CONTEXT = ['date', 'time', 'course', 'faculty'];
const STAGE_COLS = {
  all: ['college', 'studentType', 'librarian', 'sessionType', 'format', 'location', 'room', 'expected', 'actual'],
  request: ['college', 'studentType', 'format', 'location', 'expected', 'notes'],
  librarian: ['librarian', 'sessionType', 'room', 'location', 'notes'],
  planning: ['assessment', 'prep', 'notes'],
  after: ['expected', 'actual', 'notes'],
};

/** The built-in "needs attention" filter for each stage view (spec F6). */
function stageAttention(stage, r) {
  const f = r.fields;
  switch (stage) {
    case 'request': return !f.date || !f.course || !f.faculty || !f.college || !f.studentType;
    case 'librarian': return !f.librarian || /pending/i.test(f.room);
    case 'planning': return f.date >= app.today && f.date <= addDays(app.today, 14) && !f.prep;
    case 'after': return f.date && f.date < app.today && (f.actual == null || f.actual === '');
    default: return needsAttention(r, app.today).length > 0;
  }
}

function sstate() {
  app.ui.sessions = { stage: 'all', q: '', librarian: '', attention: false, review: false, table: null, ...(app.ui.sessions || {}) };
  return app.ui.sessions;
}

const sortRecs = (a, b) => (a.fields.date || '9999').localeCompare(b.fields.date || '9999') || (a.start ?? 9999) - (b.start ?? 9999);

function canEdit(data, table) {
  return !data.readOnly && app.canWrite && table && !table.missingSystem.length && !table.closed;
}

const termLabel = (t) => (t.kind === 'tours' ? 'Orientation tours' : t.term?.termYear || termYearFromSheet(t.sheet));

/** Close term (spec F31): list open items, let the coordinator continue anyway, then mark it Closed. */
async function closeTerm(table) {
  const open = table.records.map((r) => ({ r, why: needsAttention(r, app.today) })).filter((x) => x.why.length);
  const counts = {};
  for (const x of open) for (const w of x.why) counts[w] = (counts[w] || 0) + 1;
  const list = open.slice(0, 12).map(({ r, why }) => `<li>${esc(r.fields.date ? fmtDateShort(r.fields.date) : 'no date')} · ${esc(r.fields.course || r.fields.faculty)} · <span class="muted">${esc(why.join(', '))}</span></li>`).join('');
  const ok = await confirmDialog(`Close ${termLabel(table)}?`, `
    <p>Closing makes ${esc(termLabel(table))} read-only in the app on every computer. It stays in the workbook, still counts on the dashboard, and can be reopened later.</p>
    ${open.length ? `<p><span class="badge warn">${open.length} session${open.length === 1 ? '' : 's'} still need attention</span> ${Object.entries(counts).map(([k, v]) => `${esc(k)}: ${v}`).join(' · ')}</p>
      <ul class="small">${list}</ul>${open.length > 12 ? `<p class="small muted">…and ${open.length - 12} more. Use “Needs attention” on this tab to see them all.</p>` : ''}
      <p class="small">You can close it anyway and fix these after reopening, or cancel and fix them first.</p>` : '<p><span class="badge good">No open items</span> Every session has a librarian, a confirmed room and attendance.</p>'}`,
  open.length ? 'Close anyway' : 'Close term');
  if (!ok) return;
  await setClosed(table, true);
}

async function setClosed(table, closed) {
  try {
    const book = app.books.active;
    if (!app.active.hasLists) await app.store.ensureLists(book, app.store.defaultTermsFor(app.active, app.cfg.lists.terms));
    await app.store.setTermClosed(book, table.sheet, closed);
    await app.refresh();
    app.updateBanner();
    toast(closed ? `${termLabel(table)} is closed. It’s read-only now and still counts on the dashboard.` : `${termLabel(table)} is open for editing again.`);
  } catch (e) {
    toast(e.message, 'error');
  }
}

export function renderSessions(main) {
  const s = sstate();
  const data = app.data;
  const terms = data.tables.filter((t) => t.kind === 'term');
  if (!terms.length) {
    main.innerHTML = '<div class="empty">No term sheets (e.g. “FALL 26”) were found in this workbook.</div>';
    return;
  }
  let table = terms.find((t) => t.table === s.table);
  if (!table) {
    const cur = app.lists.terms.find((t) => t.from <= app.today && app.today <= t.to);
    const openTerms = terms.filter((t) => !t.closed);
    table = openTerms.find((t) => cur && t.sheet.toUpperCase() === cur.name.toUpperCase()) ?? openTerms[0] ?? terms[0];
    s.table = table.table;
  }
  renderGrid(main, data, table, terms, s, 'sessions');
}

export function renderTours(main) {
  const data = app.data;
  const table = data.tables.find((t) => t.kind === 'tours');
  if (!table) {
    main.innerHTML = '<div class="page-head"><h1>Orientation tours</h1></div><div class="empty">No tours sheet was found in this workbook.</div>';
    return;
  }
  const s = { stage: 'all', q: app.ui.toursQ || '', librarian: '', attention: false, review: false };
  renderGrid(main, data, table, [], s, 'tours');
}

function renderGrid(main, data, table, terms, s, mode) {
  const editable = canEdit(data, table);
  const q = s.q.trim().toLowerCase();
  const rows = table.records.filter((r) => {
    if (s.librarian && (r.fields.librarian || '(none)') !== s.librarian) return false;
    if (s.attention && !stageAttention(s.stage, r)) return false;
    if (s.review && !r.needsReview) return false;
    if (q && ![r.fields.course, r.fields.faculty, r.fields.notes, r.fields.location].some((v) => String(v).toLowerCase().includes(q))) return false;
    return true;
  }).sort(sortRecs);
  const cols = mode === 'tours' ? ['date', 'time', 'faculty', 'librarian', 'expected'] : [...CONTEXT, ...STAGE_COLS[s.stage]];
  const colLabel = (c) => (mode === 'tours' && c === 'faculty' ? 'Mentors' : COLS[c].label);
  const reviewCount = table.records.filter((r) => r.needsReview).length;
  const libs = [...new Set(table.records.map((r) => r.fields.librarian || '(none)'))].sort();

  main.innerHTML = `
    <div class="page-head">
      <div><h1>${mode === 'tours' ? 'Orientation tours' : 'Sessions'}</h1>
        <div class="muted small">${esc(data.book.name)}${data.readOnly ? ' · past year, read-only' : ''}${!data.readOnly && !editable && !table.closed ? ' · read-only right now' : ''}</div></div>
      <div class="row">
        ${mode === 'sessions' && !data.readOnly && app.canWrite && !table.missingSystem.length ? (table.closed
          ? '<button class="btn" id="btn-reopen">Reopen term</button>'
          : '<button class="btn" id="btn-close-term">Close term…</button>') : ''}
        <button class="btn" id="btn-export-view">Export view</button>
        <a class="btn" href="${esc(data.book.webUrl || '#')}" target="_blank" rel="noopener">Open in Excel</a>
        ${editable ? `<button class="btn primary" id="btn-new">New ${mode === 'tours' ? 'tour' : 'session'}</button>` : ''}
      </div>
    </div>
    ${mode === 'sessions' ? `<div class="tabs" role="tablist">${terms.map((t) => `<button role="tab" aria-selected="${t.table === table.table}" data-tab="${esc(t.table)}" class="${t.closed ? 'closed' : ''}">${t.closed ? '<span aria-label="Closed">🔒</span> ' : ''}${esc(termLabel(t))} <span class="muted small">${t.records.length}</span></button>`).join('')}</div>` : ''}
    ${table.closed ? `<div class="banner" style="border-radius:8px;margin-bottom:12px;border:1px solid var(--border)">🔒 <strong>${esc(termLabel(table))} is closed</strong>${table.term?.closedBy ? ` by ${esc(table.term.closedBy)}` : ''}${table.term?.closedAt ? ` on ${esc(new Date(table.term.closedAt).toLocaleDateString())}` : ''}. It’s read-only here and still counts on the dashboard.${!data.readOnly ? ' Reopen it to make changes.' : ''}</div>` : ''}
    <div class="toolbar">
      ${mode === 'sessions' ? `<label class="small muted" for="sel-stage">Stage view</label>
      <select id="sel-stage">${Object.entries(STAGES).map(([k, v]) => `<option value="${k}" ${s.stage === k ? 'selected' : ''}>${v.label}</option>`).join('')}</select>` : ''}
      <input type="search" id="in-q" placeholder="Search course, faculty, notes…" value="${esc(s.q)}" aria-label="Search">
      ${mode === 'sessions' ? `<select id="sel-lib" aria-label="Librarian"><option value="">All librarians</option>${libs.map((l) => `<option ${s.librarian === l ? 'selected' : ''}>${esc(l)}</option>`).join('')}</select>
      <label class="row small"><input type="checkbox" id="cb-attn" ${s.attention ? 'checked' : ''}> Needs attention</label>` : ''}
      ${reviewCount ? `<label class="row small"><input type="checkbox" id="cb-review" ${s.review ? 'checked' : ''}> Needs review <span class="badge warn">${reviewCount}</span></label>` : ''}
      <span class="spacer"></span><span class="muted small">${rows.length} of ${table.records.length}</span>
    </div>
    <div class="grid-wrap">
      <table class="grid">
        <thead><tr>${cols.map((c) => `<th class="${COLS[c].num ? 'num' : ''}">${esc(colLabel(c))}</th>`).join('')}<th></th></tr></thead>
        <tbody>${rows.map((r) => `<tr data-id="${esc(r.id)}" data-idx="${r.index}" tabindex="0" class="${app.changed.has(r.id) ? 'flash' : ''}">
          ${cols.map((c) => `<td class="${COLS[c].num ? 'num' : ''}" title="${esc(COLS[c].get(r))}">${c === 'librarian' && r.fields.librarian ? `<span class="swatch" style="background:${app.libColor(r.fields.librarian)}"></span>` : ''}${esc(COLS[c].get(r))}</td>`).join('')}
          <td>${flags(r, s.stage)}</td></tr>`).join('') || `<tr><td colspan="${cols.length + 1}" class="empty">No sessions match.</td></tr>`}</tbody>
      </table>
    </div>`;

  $$('[data-tab]', main).forEach((b) => (b.onclick = () => { s.table = b.dataset.tab; renderSessions(main); }));
  $('#btn-close-term', main)?.addEventListener('click', () => closeTerm(table));
  $('#btn-reopen', main)?.addEventListener('click', async () => {
    if (await confirmDialog(`Reopen ${termLabel(table)}?`, '<p>Everyone will be able to edit this term again until it is closed again.</p>', 'Reopen term')) setClosed(table, false);
  });
  $('#sel-stage', main)?.addEventListener('change', (e) => { s.stage = e.target.value; app.render(); });
  $('#sel-lib', main)?.addEventListener('change', (e) => { s.librarian = e.target.value; app.render(); });
  $('#cb-attn', main)?.addEventListener('change', (e) => { s.attention = e.target.checked; app.render(); });
  $('#cb-review', main)?.addEventListener('change', (e) => { s.review = e.target.checked; app.render(); });
  const qIn = $('#in-q', main);
  qIn.oninput = () => {
    s.q = qIn.value;
    if (mode === 'tours') app.ui.toursQ = qIn.value;
    clearTimeout(qIn._t);
    qIn._t = setTimeout(() => { app.render(); const el = $('#in-q'); el.focus(); el.setSelectionRange(el.value.length, el.value.length); }, 250);
  };
  const open = (tr) => {
    const rec = table.records.find((r) => (r.id ? r.id === tr.dataset.id : r.index === Number(tr.dataset.idx)));
    if (rec) openEditor(rec, { readOnly: !editable, stage: s.stage });
  };
  $$('tbody tr[data-idx]', main).forEach((tr) => {
    tr.onclick = () => open(tr);
    tr.onkeydown = (e) => { if (e.key === 'Enter') open(tr); };
  });
  $('#btn-new', main)?.addEventListener('click', () => openEditor(null, { table: table.table, stage: s.stage }));
  $('#btn-export-view', main).onclick = () => {
    const fields = FIELDS.filter((f) => table.map[f.key] !== undefined);
    downloadXlsx(`${table.sheet} – export.xlsx`, {
      [table.sheet]: [fields.map((f) => f.header), ...rows.map((r) => fields.map((f) => (f.key === 'time' ? r.timeDisplay : r.fields[f.key] ?? '')))],
    });
  };
}

function flags(r, stage) {
  const out = [];
  if (r.needsReview) out.push('<span class="badge warn" title="The time or date couldn’t be read reliably">⚠ review</span>');
  for (const why of needsAttention(r, app.today)) if (stage === 'all') out.push(`<span class="badge ${why === 'Attendance missing' ? '' : 'crit'}">${esc(why)}</span>`);
  return out.join(' ');
}

// ---------- Editor ----------

let editorOpen = false;
export const isEditorOpen = () => editorOpen;

export function closeEditor() {
  editorOpen = false;
  $('#drawer').classList.remove('open');
  $('#drawer').setAttribute('aria-hidden', 'true');
  $('#drawer-scrim').hidden = true;
}

function tableOf(rec, tableName) {
  const name = rec?.table ?? tableName;
  return app.data.tables.find((t) => t.table === name) ?? app.active.tables.find((t) => t.table === name);
}

/**
 * Opens the side editor. rec = existing record, or null for a new session.
 * opts: { table, prefill, readOnly, stage, onSaved(id) }
 */
export function openEditor(rec, opts = {}) {
  const table = tableOf(rec, opts.table);
  const bookId = app.data.tables.includes(table) ? app.data.book.id : app.books.active.id;
  const data = app.cache.get(bookId);
  const readOnly = opts.readOnly || !canEdit(data, table);
  const stage = opts.stage || app.ui.sessions?.stage || 'all';
  const isTours = table.kind === 'tours';
  const editableKeys = STAGES[stage].editable;
  const canTouch = (k) => !readOnly && (!editableKeys || editableKeys.includes(k));
  const shown = FIELDS.filter((f) => table.map[f.key] !== undefined);
  const base = rec ? { ...rec.fields } : Object.fromEntries(FIELDS.map((f) => [f.key, f.type === 'int' ? null : '']));
  const init = { ...base, ...(opts.prefill || {}) };
  if (!rec && !init.termYear) init.termYear = table.termYear;
  const t0 = parseTime(init.time);
  const p0 = parsePrep(init.prep);
  const termCfg = app.lists.terms.find((t) => t.name.toUpperCase() === table.sheet.toUpperCase());
  const L = app.lists;
  const opt = (list, v) => {
    const all = v && !list.includes(v) ? [...list, v] : list;
    return `<option value=""></option>${all.map((x) => `<option ${x === v ? 'selected' : ''}>${esc(x)}</option>`).join('')}`;
  };
  const dis = (k) => (canTouch(k) ? '' : 'disabled');
  const has = (k) => shown.some((f) => f.key === k);
  const field = (k, label, inner, full = false) => (has(k) ? `<label class="field ${full ? 'full' : ''}">${esc(label)}${inner}<span class="err" data-err="${k}"></span></label>` : '');
  const sel = (k, label, list) => field(k, label, `<select name="${k}" ${dis(k)}>${opt(list, init[k])}</select>`);
  const ics = icsState();
  const stale = rec && ics[rec.id] !== undefined && ics[rec.id] < rec.version;

  const title = rec ? `${init.course || (isTours ? 'Orientation tour' : 'Session')}${init.date ? ' · ' + fmtDateShort(init.date) : ''}` : opts.prefill ? 'New session from request' : `New ${isTours ? 'tour' : 'session'}`;
  const drawer = $('#drawer');
  drawer.innerHTML = `
    <div class="drawer-head">
      <div><h2>${esc(title)}</h2>
        <div class="small muted">${esc(isTours ? 'Orientation tours' : termYearFromSheet(table.sheet))}${rec?.modifiedBy ? ` · last changed by ${esc(rec.modifiedBy)}${rec.modifiedAt ? ' ' + esc(new Date(rec.modifiedAt).toLocaleString()) : ''}` : ''}</div>
        ${readOnly ? `<div class="small muted">${table.closed ? '🔒 This term is closed, so it’s read-only' : 'Read-only'}</div>` : stage !== 'all' ? `<div class="small muted">${esc(STAGES[stage].label)} stage: other fields are locked. Switch the stage view to All to edit everything.</div>` : ''}
        ${stale ? '<div class="small"><span class="badge warn">Calendar copy may be out of date</span></div>' : ''}
      </div>
      <button class="btn ghost" id="ed-close" aria-label="Close">✕</button>
    </div>
    <form class="drawer-body" id="ed-form" novalidate>
      <div class="section-label">When</div>
      ${field('date', 'Date', `<input type="date" name="date" value="${esc(init.date)}" ${dis('date')}>`)}
      ${has('termYear') ? field('termYear', 'Term Year', `<input type="text" name="termYear" value="${esc(init.termYear)}" ${dis('termYear')}>`) : '<span></span>'}
      ${has('time') ? `<label class="field">Start<input type="time" name="tStart" value="${t0 ? hhmm(t0.start) : ''}" ${dis('time')}></label>
      <label class="field">End<input type="time" name="tEnd" value="${t0?.end != null ? hhmm(t0.end) : ''}" ${dis('time')}></label>
      ${init.time !== '' && (!t0 || t0.guessed) ? `<div class="full hint">Time in the sheet: “${esc(typeof init.time === 'number' ? fmtMinutes(Math.round(init.time * 1440), false) : init.time)}”${t0 ? ' (AM/PM was guessed, please check)' : ' (couldn’t be read; it is kept as-is unless you set a start time)'}</div>` : ''}` : ''}
      <div class="full hint" id="date-warn"></div>

      <div class="section-label">${isTours ? 'Tour' : 'Course'}</div>
      ${isTours ? field('faculty', 'Mentors', `<input type="text" name="faculty" value="${esc(init.faculty)}" ${dis('faculty')}>`, true) : `
      ${field('course', 'Course #', `<input type="text" name="course" value="${esc(init.course)}" ${dis('course')} placeholder="EN 103">`)}
      ${field('faculty', 'Faculty name', `<input type="text" name="faculty" value="${esc(init.faculty)}" ${dis('faculty')}>`)}
      ${field('college', 'College/School', `<select name="college" ${dis('college')}>${opt(L.colleges, init.college)}</select><span class="suggested" data-sug="college"></span>`)}
      ${field('studentType', 'Student type', `<select name="studentType" ${dis('studentType')}>${opt(L.studentTypes, init.studentType)}</select><span class="suggested" data-sug="studentType"></span>`)}`}

      <div class="section-label">Session</div>
      ${sel('librarian', 'Librarian assigned', L.librarians.map((l) => l.name))}
      ${sel('sessionType', 'Session type', L.sessionTypes)}
      ${sel('format', 'Session format', L.sessionFormats)}
      ${field('location', 'Location', `<input type="text" name="location" list="dl-loc" value="${esc(init.location)}" ${dis('location')}><datalist id="dl-loc">${L.locations.map((x) => `<option value="${esc(x)}">`).join('')}</datalist>`)}
      ${sel('room', 'Room reservation', L.roomStatuses)}
      ${field('expected', '# Students (expected)', `<input type="number" min="0" step="1" name="expected" value="${init.expected ?? ''}" ${dis('expected')}>`)}

      ${has('assessment') || has('prep') ? '<div class="section-label">Planning</div>' : ''}
      ${sel('assessment', 'Assessment', L.yesNo)}
      ${has('prep') ? `<label class="field">Prep time<span class="row"><select name="prepKind" ${dis('prep')} aria-label="New or updated">${opt(['New', 'Updated'], p0?.kind || '')}</select><input type="number" min="0" step="0.25" name="prepHours" value="${p0?.hours ?? ''}" ${dis('prep')} aria-label="Hours" style="width:90px"> hours</span>
        ${init.prep && !p0 ? `<span class="hint">In the sheet: “${esc(init.prep)}”</span>` : ''}</label>` : ''}

      ${has('actual') ? '<div class="section-label">After the session</div>' : ''}
      ${field('actual', 'Actual student attendance', `<input type="number" min="0" step="1" name="actual" value="${init.actual ?? ''}" ${dis('actual')}>`)}
      ${field('notes', 'Notes', `<textarea name="notes" ${dis('notes')}>${esc(init.notes)}</textarea>`, true)}
    </form>
    <div class="drawer-foot">
      ${readOnly ? '' : `<button class="btn primary" id="ed-save">${rec ? 'Save' : 'Add session'}</button>`}
      ${rec && !readOnly && !isTours ? '<button class="btn" id="ed-dup">Duplicate</button>' : ''}
      ${rec && !readOnly && !isTours ? `<select id="ed-move" aria-label="Move to term"><option value="">Move to term…</option>${data.tables.filter((t) => t.kind === 'term' && t.table !== table.table && !t.missingSystem.length && !t.closed).map((t) => `<option value="${esc(t.table)}">${esc(termYearFromSheet(t.sheet))}</option>`).join('')}</select>` : ''}
      ${rec && init.date ? `<button class="btn" id="ed-ics">${stale ? 'Update calendar copy' : 'Add to my calendar'}</button>` : ''}
      <span class="spacer"></span>
      ${rec && !readOnly ? '<button class="btn danger" id="ed-del">Delete</button>' : ''}
    </div>`;

  editorOpen = true;
  drawer.classList.add('open');
  drawer.setAttribute('aria-hidden', 'false');
  $('#drawer-scrim').hidden = false;
  $('#drawer-scrim').onclick = closeEditor;
  $('#ed-close').onclick = closeEditor;
  setTimeout(() => drawer.querySelector('input:not([disabled]), select:not([disabled])')?.focus(), 50);

  const form = $('#ed-form');
  const timeTouched = { v: false }, prepTouched = { v: false };
  form.tStart?.addEventListener('input', () => (timeTouched.v = true));
  form.tEnd?.addEventListener('input', () => (timeTouched.v = true));
  form.prepKind?.addEventListener('input', () => (prepTouched.v = true));
  form.prepHours?.addEventListener('input', () => (prepTouched.v = true));

  // Date outside the term's range → warning (not an error).
  const checkDate = () => {
    const d = form.date?.value;
    const w = $('#date-warn');
    if (d && termCfg && (d < termCfg.from || d > termCfg.to)) {
      const other = termForDate(d, app.lists.terms);
      w.innerHTML = `<span class="badge warn">⚠</span> This date is outside ${esc(termYearFromSheet(table.sheet))}${other ? ` (it falls in ${esc(termYearFromSheet(other))}; use “Move to term” after saving)` : ''}.`;
    } else w.textContent = '';
  };
  form.date?.addEventListener('change', checkDate);
  checkDate();

  // Course # → suggest College/School and Student Type when they're blank (spec F8).
  const suggest = () => {
    if (!form.course) return;
    form.course.value = normalizeCourse(form.course.value);
    const sug = suggestFromCourse(form.course.value, L);
    for (const k of ['college', 'studentType']) {
      const elx = form[k];
      if (!elx || elx.disabled || !sug[k]) continue;
      if (!elx.value || elx.dataset.suggested === '1') {
        elx.value = sug[k];
        elx.dataset.suggested = '1';
        $(`[data-sug=${k}]`).textContent = 'Suggested from the course number. Check it.';
      }
    }
  };
  form.course?.addEventListener('change', suggest);
  for (const k of ['college', 'studentType']) form[k]?.addEventListener('change', () => { form[k].dataset.suggested = ''; $(`[data-sug=${k}]`).textContent = ''; });
  if (!rec && opts.prefill?.course) suggest();

  const collect = () => {
    const mine = { ...init };
    for (const f of shown) {
      if (!canTouch(f.key)) continue;
      if (f.key === 'time') {
        if (timeTouched.v || !rec) {
          const a = toMins(form.tStart?.value), b = toMins(form.tEnd?.value);
          mine.time = a == null ? (rec ? '' : init.time) : formatTimeRange(a, b);
        }
        continue;
      }
      if (f.key === 'prep') {
        if (prepTouched.v) mine.prep = form.prepHours.value === '' ? '' : formatPrep(form.prepKind.value, Number(form.prepHours.value));
        continue;
      }
      const el = form[f.key];
      if (!el) continue;
      let v = el.value;
      if (f.type === 'int') v = v === '' ? null : Number(v);
      if (f.key === 'course') v = normalizeCourse(v);
      if (f.key === 'location') v = normalizeLocation(v, L.locations);
      mine[f.key] = typeof v === 'string' ? v.trim() : v;
    }
    return mine;
  };

  const showErrors = (errors) => {
    $$('[data-err]', drawer).forEach((e) => (e.textContent = errors[e.dataset.err] || ''));
    return Object.keys(errors).length === 0;
  };

  const busy = (on) => $$('.drawer-foot button, .drawer-foot select', drawer).forEach((b) => (b.disabled = on));

  const afterWrite = async (msg) => {
    closeEditor();
    toast(msg);
    await app.refresh();
  };

  const handleError = (e) => {
    console.error(e);
    busy(false);
    toast(e.busy ? 'The workbook is busy (someone may have it open in older Excel). Your edits are still here. Try again shortly.' : `Couldn’t save: ${e.message}`, 'error');
  };

  const doSave = async (mine, resolutions = {}) => {
    const r = await app.store.save(bookId, rec, base, mine, resolutions);
    if (r.status === 'saved') return afterWrite(r.mergedOthers ? 'Saved. Someone else’s changes to other fields were kept too.' : 'Saved.');
    if (r.status === 'deleted') {
      busy(false);
      const re = await confirmDialog('This session was deleted', '<p>Someone else deleted this session while you were editing it. Add it back with your changes?</p>', 'Re-add session');
      if (re) {
        await app.store.add(bookId, rec.table, mine);
        return afterWrite('Session re-added.');
      }
      closeEditor();
      return app.refresh();
    }
    // Conflict: same field changed on both sides.
    busy(false);
    const picks = await conflictDialog(r.conflicts, r.theirs);
    if (!picks) return;
    busy(true);
    return doSave(mine, { ...resolutions, ...picks });
  };

  $('#ed-save')?.addEventListener('click', async () => {
    const mine = collect();
    if (!showErrors(isTours ? (mine.date ? {} : { date: 'Required' }) : validate(mine))) return;
    busy(true);
    try {
      if (rec) await doSave(mine);
      else {
        const id = await app.store.add(bookId, table.table, mine);
        await opts.onSaved?.(id, table);
        await afterWrite(isTours ? 'Tour added.' : 'Session added.');
      }
    } catch (e) { handleError(e); }
  });

  $('#ed-dup')?.addEventListener('click', () => {
    const mine = collect();
    closeEditor();
    setTimeout(() => openEditor(null, { table: table.table, stage: 'all', prefill: { ...mine, date: '', actual: null, room: '' } }), 200);
    toast('Copy opened. Pick the new date and save.');
  });

  $('#ed-move')?.addEventListener('change', async (e) => {
    const target = data.tables.find((t) => t.table === e.target.value);
    if (!target) return;
    if (!(await confirmDialog('Move session', `<p>Move this session to <strong>${esc(termYearFromSheet(target.sheet))}</strong>? Unsaved edits in this form are not included.</p>`, 'Move'))) { e.target.value = ''; return; }
    busy(true);
    try {
      const r = await app.store.move(bookId, rec, target.table, target.termYear);
      await afterWrite(r.status === 'deleted' ? 'That session was already deleted by someone else.' : `Moved to ${termYearFromSheet(target.sheet)}.`);
    } catch (err) { handleError(err); }
  });

  $('#ed-del')?.addEventListener('click', async () => {
    if (!(await confirmDialog('Delete session?', `<p>Delete <strong>${esc(init.course || 'this session')}</strong> on ${esc(fmtDateShort(init.date) || 'no date')}? It can be restored from the workbook’s SharePoint version history.</p>`, 'Delete', true))) return;
    busy(true);
    try {
      const r = await app.store.remove(bookId, rec);
      await afterWrite(r.status === 'deleted' ? 'It had already been deleted by someone else.' : 'Session deleted.');
    } catch (e) { handleError(e); }
  });

  $('#ed-ics')?.addEventListener('click', () => {
    download(`${(init.course || 'session').replace(/[\\/:*?"<>|]/g, '-')} ${init.date}.ics`, buildIcs([rec]), 'text/calendar');
    markIcsExported([rec]);
    toast('Saved. Open the file to add it to Outlook.');
  });
}

const hhmm = (m) => `${String(Math.floor(m / 60)).padStart(2, '0')}:${String(m % 60).padStart(2, '0')}`;
const toMins = (v) => (v ? Number(v.slice(0, 2)) * 60 + Number(v.slice(3, 5)) : null);

/** Shows Theirs vs Mine for each conflicting field. Resolves { key: 'mine'|'theirs' } or null. */
function conflictDialog(conflicts, theirs) {
  const who = theirs.modifiedBy || 'someone in Excel';
  const when = theirs.modifiedAt ? new Date(theirs.modifiedAt).toLocaleString() : '';
  const show = (v) => (v === '' || v == null ? '<em class="muted">(blank)</em>' : esc(v));
  return openDialog(`
    <form method="dialog">
      <div class="dlg-head"><h2>Someone else changed the same fields</h2>
        <p class="muted small">${esc(who)}${when ? ` saved changes at ${esc(when)}` : ' changed this row'}. Choose which value to keep for each field. Fields only one of you changed were merged automatically.</p></div>
      <div class="dlg-body">
        ${conflicts.map((c) => `<div class="conflict"><strong>${esc(c.header)}</strong>
          <label><input type="radio" name="${c.key}" value="theirs" checked><span><span class="small muted">Theirs</span><br>${show(c.theirs)}</span></label>
          <label><input type="radio" name="${c.key}" value="mine"><span><span class="small muted">Mine</span><br>${show(c.mine)}</span></label></div>`).join('')}
      </div>
      <div class="dlg-foot"><button class="btn" type="button" data-x="no">Back to editing</button><button class="btn primary" type="submit">Save with these choices</button></div>
    </form>`,
  (dlg, done) => {
    dlg.querySelector('[data-x=no]').onclick = () => done(null);
    dlg.querySelector('form').onsubmit = (e) => {
      e.preventDefault();
      done(Object.fromEntries(conflicts.map((c) => [c.key, dlg.querySelector(`input[name="${c.key}"]:checked`).value])));
    };
  });
}
