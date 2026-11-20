// Librarian dashboard (spec §4.6): period selector, filters, summary tiles, calendar, three bar charts.
import { app } from './app.js';
import { FILTER_DIMS, applyFilters, groupBy, summarize, needsAttention, addDays, termYearFromSheet, buildIcs, DASH_SESSION_TYPE, studentsOf } from './model.js';
import { $, $$, esc, bindTips, hideTip, fmtDate, downloadXlsx, download, openDialog, toast } from './ui.js';
import { openEditor, icsState, markIcsExported } from './sessions.js';
import { copyChartPng } from './chartpng.js';

const PREF_KEY = 'tracker.dashboard.v1';

function state() {
  if (!app.ui.dash) {
    let saved = {};
    try { saved = JSON.parse(localStorage.getItem(PREF_KEY) || '{}'); } catch { /* storage blocked */ }
    const filters = {};
    for (const d of FILTER_DIMS) filters[d.key] = new Set(saved.filters?.[d.key] ?? []);
    app.ui.dash = { period: 'this-term', from: '', to: '', filters, measure: saved.measure || 'sessions', calView: 'month', cal: app.today, asTable: {} };
  }
  return app.ui.dash;
}

function savePrefs(s) {
  try {
    const filters = Object.fromEntries(Object.entries(s.filters).map(([k, v]) => [k, [...v]]));
    localStorage.setItem(PREF_KEY, JSON.stringify({ filters, measure: s.measure }));
  } catch { /* storage blocked: preferences just aren't remembered */ }
}

/** The term table whose configured date range contains today (or the first term). */
function currentTerm(data) {
  const terms = data.tables.filter((t) => t.kind === 'term');
  const hit = app.lists.terms.find((t) => t.from <= app.today && app.today <= t.to && terms.some((x) => x.sheet.toUpperCase() === t.name.toUpperCase()));
  return hit ? terms.find((x) => x.sheet.toUpperCase() === hit.name.toUpperCase()) : terms[0];
}

function periodRecords(data, s) {
  if (s.period === 'year') return { label: `Academic year ${data.book.year ?? ''}`, records: data.records };
  if (s.period === 'range') {
    const recs = applyFilters(data.records, {}, { from: s.from || null, to: s.to || null });
    return { label: `${s.from ? fmtDate(s.from) : '…'} – ${s.to ? fmtDate(s.to) : '…'}`, records: recs };
  }
  const t = s.period.startsWith('term:') ? data.tables.find((x) => x.table === s.period.slice(5)) : currentTerm(data);
  if (!t) return { label: 'No terms', records: [] };
  return { label: t.kind === 'tours' ? 'Orientation tours' : termYearFromSheet(t.sheet), records: t.records, table: t };
}

function filtersExcept(s, key) {
  return Object.fromEntries(Object.entries(s.filters).filter(([k]) => k !== key));
}

export function renderDashboard(main) {
  const s = state();
  const data = app.data;
  // "This term" means nothing for an archived year.
  if (data.readOnly && s.period === 'this-term') s.period = 'year';
  const { label, records: inPeriod } = periodRecords(data, s);
  const filtered = applyFilters(inPeriod, s.filters, null);
  const sum = summarize(filtered, app.today);
  const terms = data.tables;
  const activeFilterCount = Object.values(s.filters).reduce((n, v) => n + v.size, 0);

  main.innerHTML = `
    <div class="page-head">
      <div><h1>Dashboard</h1><div class="muted small">${esc(label)} · ${esc(data.book.name)}${data.readOnly ? ' · read-only' : ''}</div></div>
      <div class="row">
        <button class="btn" id="btn-ics-mine">Add my sessions to calendar</button>
        <button class="btn" id="btn-export">Export report</button>
      </div>
    </div>
    <div class="toolbar" role="group" aria-label="Period">
      <label class="sr-only" for="sel-book">Academic year</label>
      <select id="sel-book" aria-label="Academic year">
        ${[app.books.active, ...app.books.archived].map((b) => `<option value="${esc(b.id)}" ${b.id === data.book.id ? 'selected' : ''}>${esc(b.year || b.name)}${b.id === app.books.active.id ? ' (current)' : ''}</option>`).join('')}
      </select>
      <select id="sel-period" aria-label="Period">
        ${data.readOnly ? '' : `<option value="this-term" ${s.period === 'this-term' ? 'selected' : ''}>This term</option>`}
        ${terms.map((t) => `<option value="term:${esc(t.table)}" ${s.period === `term:${t.table}` ? 'selected' : ''}>${esc(t.kind === 'tours' ? 'Orientation tours' : termYearFromSheet(t.sheet))}</option>`).join('')}
        <option value="year" ${s.period === 'year' ? 'selected' : ''}>Academic year</option>
        <option value="range" ${s.period === 'range' ? 'selected' : ''}>Custom date range…</option>
      </select>
      ${s.period === 'range' ? `<input type="date" id="in-from" value="${esc(s.from)}" aria-label="From"> <span class="muted">to</span> <input type="date" id="in-to" value="${esc(s.to)}" aria-label="To">` : ''}
    </div>
    <div class="filters" id="filters">
      ${FILTER_DIMS.map((d) => filterControl(d, inPeriod, s)).join('')}
      ${Object.entries(s.filters).flatMap(([k, set]) => [...set].map((v) => `<span class="chip">${esc(v)}<button aria-label="Remove filter ${esc(v)}" data-unchip="${esc(k)}" data-v="${esc(v)}">×</button></span>`)).join('')}
      ${activeFilterCount ? '<button class="btn ghost small" id="btn-clear">Clear all</button>' : ''}
    </div>
    <div class="tiles">
      <div class="tile"><div class="tile-label">Sessions</div><div class="tile-value">${sum.sessions}</div><div class="tile-note">One row = one booking</div></div>
      <div class="tile"><div class="tile-label">Students reached</div><div class="tile-value">${sum.students.toLocaleString()}</div><div class="tile-note">${sum.estimates ? `includes ${sum.estimates} estimate${sum.estimates === 1 ? '' : 's'}` : 'actual attendance'}</div></div>
      <div class="tile"><div class="tile-label">Upcoming (next 14 days)</div><div class="tile-value">${sum.upcoming}</div><div class="tile-note">${fmtDate(app.today, { month: 'short', day: 'numeric' })} – ${fmtDate(addDays(app.today, 14), { month: 'short', day: 'numeric' })}</div></div>
      <button class="tile" id="tile-attn"><div class="tile-label">Needs attention</div><div class="tile-value">${sum.attention}</div><div class="tile-note">Pending rooms, no librarian, missing attendance →</div></button>
    </div>
    <div class="dash-grid">
      <div class="dash-col">
        <section class="card" aria-label="Calendar">${calendarHtml(filtered, s)}</section>
      </div>
      <div class="dash-col">
        ${chartCard('librarian', 'Sessions by librarian', librarianData(inPeriod, s), s, true)}
        ${chartCard('studentType', 'Bookings by program', programData(inPeriod, s), s)}
        ${chartCard('college', 'Bookings by college/school', collegeData(inPeriod, s), s)}
        <p class="small muted" title="How numbers are counted">How numbers are counted: one row is one booking (a 10-module series counts as 10). Students reached uses actual attendance, or the expected count where attendance hasn’t been entered. Workshop and Presentation are grouped. Charts ignore their own filter so you can see what you’re choosing between.</p>
      </div>
    </div>`;

  wire(main, s, data, filtered, inPeriod, label);
}

function filterControl(dim, records, s) {
  const values = new Set(records.map(dim.value));
  if (dim.key === 'librarian') app.lists.librarians.forEach((l) => values.add(l.name));
  if (dim.key === 'studentType') app.lists.studentTypes.forEach((v) => values.add(v));
  if (dim.key === 'format') app.lists.sessionFormats.forEach((v) => values.add(v));
  for (const v of s.filters[dim.key]) values.add(v);
  const sorted = [...values].sort((a, b) => (a === 'Unassigned') - (b === 'Unassigned') || a.localeCompare(b));
  const sel = s.filters[dim.key];
  return `<details class="ms" data-active="${sel.size > 0}">
    <summary>${esc(dim.label)}${sel.size ? ` (${sel.size})` : ''}</summary>
    <div class="ms-panel">${sorted.map((v) => `<label><input type="checkbox" data-dim="${dim.key}" value="${esc(v)}" ${sel.has(v) ? 'checked' : ''}>${esc(v)}</label>`).join('')}</div>
  </details>`;
}

// ---------- Chart data ----------

function librarianData(inPeriod, s) {
  const recs = applyFilters(inPeriod, filtersExcept(s, 'librarian'), null);
  const rows = groupBy(recs, (r) => r.fields.librarian || 'Unassigned', s.measure);
  for (const l of app.lists.librarians) if (!rows.some((r) => r.label === l.name)) rows.push({ label: l.name, value: 0 });
  rows.sort((a, b) => (a.label === 'Unassigned') - (b.label === 'Unassigned') || b.value - a.value);
  return rows.map((r) => ({ ...r, color: r.label === 'Unassigned' ? 'var(--series-none)' : app.libColor(r.label) }));
}

function programData(inPeriod, s) {
  const recs = applyFilters(inPeriod, filtersExcept(s, 'studentType'), null);
  const counts = groupBy(recs, (r) => r.fields.studentType || 'Unassigned');
  const order = [...app.lists.studentTypes, 'Unassigned'];
  return order
    .map((label) => ({ label, value: counts.find((c) => c.label === label)?.value ?? 0 }))
    .concat(counts.filter((c) => !order.includes(c.label)))
    .filter((r) => r.label !== 'Unassigned' || r.value > 0)
    .map((r) => ({ ...r, color: r.label === 'Unassigned' ? 'var(--series-none)' : 'var(--series-1)' }));
}

function collegeData(inPeriod, s) {
  const recs = applyFilters(inPeriod, filtersExcept(s, 'college'), null);
  return groupBy(recs, (r) => r.fields.college || 'Unassigned')
    .sort((a, b) => (a.label === 'Unassigned') - (b.label === 'Unassigned') || b.value - a.value)
    .map((r) => ({ ...r, color: r.label === 'Unassigned' ? 'var(--series-none)' : 'var(--series-1)' }));
}

function chartCard(dimKey, title, rows, s, withMeasure = false) {
  const max = Math.max(1, ...rows.map((r) => r.value));
  const total = rows.reduce((n, r) => n + r.value, 0);
  const sel = s.filters[dimKey];
  const unit = withMeasure && s.measure === 'students' ? 'students' : 'sessions';
  const titleText = withMeasure && s.measure === 'students' ? title.replace('Sessions', 'Students reached') : title;
  const body = s.asTable[dimKey]
    ? `<table class="data-table"><thead><tr><th>${esc(dimKey === 'librarian' ? 'Librarian' : dimKey === 'college' ? 'College/School' : 'Program')}</th><th class="num">${esc(unit)}</th><th class="num">Share</th></tr></thead>
       <tbody>${rows.map((r) => `<tr><td>${esc(r.label)}</td><td class="num">${r.value}</td><td class="num">${total ? Math.round((r.value / total) * 100) : 0}%</td></tr>`).join('')}</tbody></table>`
    : rows.length
      ? `<div class="bars">${rows.map((r) => {
        const pressed = sel.has(r.label);
        const tip = `<strong>${esc(r.label)}</strong><br>${r.value} ${unit} · ${total ? Math.round((r.value / total) * 100) : 0}%<br><span style="opacity:.75">Click to ${pressed ? 'remove' : 'add'} filter</span>`;
        return `<button class="bar-row ${sel.size && !pressed ? 'dim' : ''}" aria-pressed="${pressed}" data-bar="${dimKey}" data-v="${esc(r.label)}" data-tip="${esc(tip)}" aria-label="${esc(`${r.label}: ${r.value} ${unit}`)}">
          <span class="bar-label">${dimKey === 'librarian' ? `<span class="swatch" style="background:${r.color}"></span>` : ''}${esc(r.label)}</span>
          <span class="bar-track"><span class="bar-fill" style="display:block;width:${(r.value / max) * 100}%;background:${r.color}"></span></span>
          <span class="bar-value">${r.value}</span></button>`;
      }).join('')}</div>`
      : '<div class="empty">No sessions match.</div>';
  return `<section class="card" aria-label="${esc(titleText)}">
    <div class="card-head"><h2>${esc(titleText)}</h2>
      <div class="row">
        ${withMeasure ? `<div class="seg" role="group" aria-label="Measure"><button data-measure="sessions" aria-pressed="${s.measure === 'sessions'}">Sessions</button><button data-measure="students" aria-pressed="${s.measure === 'students'}">Students</button></div>` : ''}
        <button class="btn ghost small" data-copychart="${dimKey}" title="Copy this chart as a picture, to paste into slides or email">Copy chart</button>
        <button class="btn ghost small" data-astable="${dimKey}">${s.asTable[dimKey] ? 'View as chart' : 'View as table'}</button>
      </div>
    </div>${body}</section>`;
}

// ---------- Calendar (spec F23) ----------

/** 780 → '1p', 570 → '9:30a' (compact, for calendar chips). */
function shortTime(mins) {
  if (mins == null) return '';
  const h24 = Math.floor(mins / 60), m = mins % 60;
  return `${((h24 + 11) % 12) + 1}${m ? ':' + String(m).padStart(2, '0') : ''}${h24 < 12 ? 'a' : 'p'}`;
}

function startOfWeek(iso) {
  const d = new Date(`${iso}T12:00:00`);
  return addDays(iso, -d.getDay());
}

function calendarHtml(records, s) {
  const byDay = new Map();
  for (const r of records) {
    if (!r.fields.date) continue;
    if (!byDay.has(r.fields.date)) byDay.set(r.fields.date, []);
    byDay.get(r.fields.date).push(r);
  }
  for (const list of byDay.values()) list.sort((a, b) => (a.start ?? 9999) - (b.start ?? 9999));

  const cur = new Date(`${s.cal}T12:00:00`);
  let days, title;
  if (s.calView === 'week') {
    const first = startOfWeek(s.cal);
    days = Array.from({ length: 7 }, (_, i) => addDays(first, i));
    title = `Week of ${fmtDate(first, { month: 'short', day: 'numeric', year: 'numeric' })}`;
  } else {
    const firstOfMonth = `${cur.getFullYear()}-${String(cur.getMonth() + 1).padStart(2, '0')}-01`;
    const gridStart = startOfWeek(firstOfMonth);
    days = Array.from({ length: 42 }, (_, i) => addDays(gridStart, i));
    if (days.slice(35).every((d) => d.slice(5, 7) !== firstOfMonth.slice(5, 7))) days = days.slice(0, 35);
    title = cur.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
  }
  const month = s.cal.slice(0, 7);
  const limit = s.calView === 'week' ? 99 : 3;
  const ics = icsState();

  const cells = days.map((d) => {
    const list = byDay.get(d) ?? [];
    const chips = list.slice(0, limit).map((r) => {
      const f = r.fields;
      const stale = ics[r.id] !== undefined && ics[r.id] < r.version;
      const tip = `<strong>${esc(f.course || '(no course)')}</strong> · ${esc(r.timeDisplay || 'time TBA')}<br>${esc(f.faculty)}<br>${esc([f.location, f.format].filter(Boolean).join(' · '))}<br>Librarian: ${esc(f.librarian || 'unassigned')}${r.kind !== 'tours' ? ` · Room: ${esc(f.room || '—')}` : ''}${stale ? '<br>Calendar copy may be out of date' : ''}`;
      return `<button class="cal-chip ${stale ? 'stale' : ''}" style="--lib:${app.libColor(f.librarian)}" data-rec="${esc(r.id)}" data-tip="${esc(tip)}" aria-label="${esc(`${f.course} ${r.timeDisplay} with ${f.librarian || 'no librarian'}`)}"><span class="t">${esc(shortTime(r.start))}</span>${esc(f.course || f.faculty || 'Session')}</button>`;
    }).join('');
    const more = list.length > limit ? `<button class="cal-more" data-week="${d}">+${list.length - limit} more</button>` : '';
    const cls = [s.calView === 'month' && d.slice(0, 7) !== month ? 'other' : '', d === app.today ? 'today' : ''].join(' ');
    return `<div class="cal-day ${cls}"><span class="cal-num">${s.calView === 'week' ? fmtDate(d, { weekday: 'short', day: 'numeric' }) : Number(d.slice(8))}</span>${chips}${more}</div>`;
  }).join('');

  const legend = app.lists.librarians.map((l) => `<span><span class="swatch" style="background:var(--series-${l.color})"></span>${esc(l.name)}</span>`).join('') + '<span><span class="swatch" style="background:var(--series-none)"></span>Unassigned</span>';
  return `<div class="card-head"><h2>Calendar</h2><div class="legend">${legend}</div></div>
    <div class="cal-head">
      <button class="btn small" data-cal="-1" aria-label="Previous">‹</button>
      <button class="btn small" data-cal="today">Today</button>
      <button class="btn small" data-cal="1" aria-label="Next">›</button>
      <span class="cal-title">${esc(title)}</span>
      <span class="spacer"></span>
      <div class="seg" role="group" aria-label="Calendar view"><button data-calview="month" aria-pressed="${s.calView === 'month'}">Month</button><button data-calview="week" aria-pressed="${s.calView === 'week'}">Week</button></div>
    </div>
    <div class="cal ${s.calView}">
      ${['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'].map((d) => `<div class="cal-dow">${d}</div>`).join('')}
      ${cells}
    </div>`;
}

// ---------- Events ----------

function wire(main, s, data, filtered, inPeriod, label) {
  const rerender = () => { savePrefs(s); renderDashboard(main); };
  bindTips(main);

  $('#sel-book', main).onchange = async (e) => {
    app.viewBookId = e.target.value;
    if (!app.cache.has(app.viewBookId)) {
      main.querySelector('.tiles').style.opacity = '.4';
      try { await app.load(app.viewBookId); } catch (err) { toast(err.message, 'error'); app.viewBookId = app.books.active.id; }
    }
    s.period = app.viewBookId === app.books.active.id ? 'this-term' : 'year';
    rerender();
  };
  $('#sel-period', main).onchange = (e) => { s.period = e.target.value; rerender(); };
  $('#in-from', main)?.addEventListener('change', (e) => { s.from = e.target.value; rerender(); });
  $('#in-to', main)?.addEventListener('change', (e) => { s.to = e.target.value; rerender(); });

  // Keep only one filter dropdown open; close on outside click.
  const details = $$('.ms', main);
  details.forEach((d) => d.addEventListener('toggle', () => { if (d.open) details.filter((x) => x !== d).forEach((x) => (x.open = false)); }));
  main.addEventListener('click', (e) => { if (!e.target.closest('.ms')) details.forEach((d) => (d.open = false)); });
  $$('.ms-panel input', main).forEach((cb) => (cb.onchange = () => {
    const set = s.filters[cb.dataset.dim];
    cb.checked ? set.add(cb.value) : set.delete(cb.value);
    const openDim = cb.dataset.dim;
    savePrefs(s);
    renderDashboard(main);
    const reopened = $(`.ms input[data-dim="${openDim}"]`, main)?.closest('details');
    if (reopened) reopened.open = true;
  }));
  $$('[data-unchip]', main).forEach((b) => (b.onclick = () => { s.filters[b.dataset.unchip].delete(b.dataset.v); rerender(); }));
  $('#btn-clear', main)?.addEventListener('click', () => { Object.values(s.filters).forEach((x) => x.clear()); rerender(); });

  // Cross-filtering (spec F27)
  $$('[data-bar]', main).forEach((b) => (b.onclick = () => {
    hideTip();
    const set = s.filters[b.dataset.bar];
    set.has(b.dataset.v) ? set.delete(b.dataset.v) : set.add(b.dataset.v);
    rerender();
  }));
  $$('[data-measure]', main).forEach((b) => (b.onclick = () => { s.measure = b.dataset.measure; rerender(); }));
  $$('[data-astable]', main).forEach((b) => (b.onclick = () => { s.asTable[b.dataset.astable] = !s.asTable[b.dataset.astable]; rerender(); }));

  // Copy chart as PNG (spec F28). Drawn from the same data as the bars, with the period and filters underneath the title.
  $$('[data-copychart]', main).forEach((b) => (b.onclick = async () => {
    const dim = b.dataset.copychart;
    const students = dim === 'librarian' && s.measure === 'students';
    const rows = { librarian: librarianData, studentType: programData, college: collegeData }[dim](inPeriod, s);
    const title = { librarian: students ? 'Students reached by librarian' : 'Sessions by librarian', studentType: 'Bookings by program', college: 'Bookings by college/school' }[dim];
    const filters = Object.entries(s.filters).filter(([k, v]) => k !== dim && v.size)
      .map(([k, v]) => `${FILTER_DIMS.find((d) => d.key === k).label}: ${[...v].join(', ')}`);
    const total = rows.reduce((n, r) => n + r.value, 0);
    try {
      const how = await copyChartPng({
        title,
        subtitle: [label, `${total} ${students ? 'students' : 'sessions'}`, ...filters].join(' · '),
        rows,
        footer: `Trexler Library instruction · ${data.book.name.replace(/\.xlsx$/i, '')} · ${new Date().toLocaleDateString()}`,
      }, `${title} – ${label.replace(/[\\/:*?"<>|]/g, '-')}.png`);
      toast(how === 'copied' ? 'Chart copied. Paste it into your slides or email.' : 'This browser can’t copy pictures, so the chart was downloaded as a PNG instead.');
    } catch (e) {
      toast(e.message, 'error');
    }
  }));

  // Calendar
  $$('[data-cal]', main).forEach((b) => (b.onclick = () => {
    const v = b.dataset.cal;
    if (v === 'today') s.cal = app.today;
    else if (s.calView === 'week') s.cal = addDays(s.cal, 7 * Number(v));
    else {
      const d = new Date(`${s.cal.slice(0, 7)}-01T12:00:00`);
      d.setMonth(d.getMonth() + Number(v));
      s.cal = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-01`;
    }
    rerender();
  }));
  $$('[data-calview]', main).forEach((b) => (b.onclick = () => { s.calView = b.dataset.calview; rerender(); }));
  $$('[data-week]', main).forEach((b) => (b.onclick = () => { s.calView = 'week'; s.cal = b.dataset.week; rerender(); }));
  $$('[data-rec]', main).forEach((b) => (b.onclick = () => {
    hideTip();
    const rec = data.records.find((r) => r.id === b.dataset.rec);
    if (rec) openEditor(rec, { readOnly: data.readOnly });
  }));

  $('#tile-attn', main).onclick = () => {
    app.ui.sessions = { ...(app.ui.sessions || {}), attention: true, stage: 'all' };
    location.hash = '#/sessions';
  };

  $('#btn-export', main).onclick = () => {
    const charts = {
      'By librarian': librarianData(inPeriod, s),
      'By program': programData(inPeriod, s),
      'By college-school': collegeData(inPeriod, s),
    };
    const sheets = {};
    const filterText = Object.entries(s.filters).filter(([, v]) => v.size).map(([k, v]) => `${FILTER_DIMS.find((d) => d.key === k).label}: ${[...v].join(', ')}`).join('; ') || 'none';
    sheets.Summary = [['Period', label], ['Workbook', data.book.name], ['Filters', filterText], ['Generated', new Date().toLocaleString()], [],
      ['Sessions', filtered.length], ['Students reached', summarize(filtered, app.today).students], ['…of which estimates', summarize(filtered, app.today).estimates]];
    for (const [name, rows] of Object.entries(charts)) sheets[name] = [[name.replace('By ', ''), name === 'By librarian' && s.measure === 'students' ? 'Students' : 'Sessions'], ...rows.map((r) => [r.label, r.value])];
    sheets.Sessions = [['Term Year', 'Date', 'Time', 'Course #', 'Faculty Name', 'College/School', 'Student Type', 'Librarian Assigned', 'Session Type', 'Session Format', 'Location', 'Room Reservation Status', '# Students', 'Actual Student Attendance', 'Students counted', 'Needs attention'],
      ...filtered.slice().sort((a, b) => (a.fields.date || '').localeCompare(b.fields.date || '') || (a.start ?? 0) - (b.start ?? 0)).map((r) => {
        const f = r.fields;
        return [f.termYear, f.date, r.timeDisplay, f.course, f.faculty, f.college, f.studentType, f.librarian, DASH_SESSION_TYPE(f.sessionType), f.format, f.location, f.room, f.expected ?? '', f.actual ?? '', studentsOf(r).n, needsAttention(r, app.today).join(', ')];
      })];
    downloadXlsx(`Instruction report – ${label.replace(/[\\/:*?"<>|]/g, '-')}.xlsx`, sheets);
  };

  $('#btn-ics-mine', main).onclick = async () => {
    const first = (app.user.name || '').split(/\s+/)[0].toLowerCase();
    const guess = app.lists.librarians.find((l) => l.name.toLowerCase() === first)?.name ?? app.lists.librarians[0]?.name;
    const who = await openDialog(`
      <form method="dialog"><div class="dlg-head"><h2>Add sessions to my calendar</h2></div>
      <div class="dlg-body"><p class="muted">Saves an .ics file with every upcoming session in <strong>${esc(label)}</strong> assigned to:</p>
        <select name="who">${app.lists.librarians.map((l) => `<option ${l.name === guess ? 'selected' : ''}>${esc(l.name)}</option>`).join('')}</select>
        <p class="small muted">Open the file and Outlook adds the events. Re-exporting later updates the same events instead of duplicating them.</p></div>
      <div class="dlg-foot"><button class="btn" type="button" data-x="no">Cancel</button><button class="btn primary" type="submit">Download .ics</button></div></form>`,
    (dlg, done) => {
      dlg.querySelector('[data-x=no]').onclick = () => done(null);
      dlg.querySelector('form').onsubmit = (e) => { e.preventDefault(); done(dlg.querySelector('select').value); };
    });
    if (!who) return;
    const mine = inPeriod.filter((r) => r.fields.librarian?.toLowerCase() === who.toLowerCase() && r.fields.date >= app.today);
    if (!mine.length) return toast(`No upcoming sessions for ${who} in ${label}.`);
    download(`${who} – instruction sessions.ics`, buildIcs(mine), 'text/calendar');
    markIcsExported(mine);
    toast(`${mine.length} sessions saved. Open the file to add them to Outlook.`);
  };
}
