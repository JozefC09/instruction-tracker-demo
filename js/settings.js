// Settings: connection details, workbooks, "Prepare workbook", and demo-only test helpers.
import { app, syncNow } from './app.js';
import { Store } from './store.js';
import { defaultPlan } from './newyear.js';
import { $, $$, esc, toast, confirmDialog, fmtDate } from './ui.js';

export function renderSettings(main) {
  const a = app.active;
  const issues = a.tables.filter((t) => t.missingSystem.length || t.missingFields.length);
  const sheetList = a.tables.map((t) => `<tr><td>${esc(t.sheet)}</td><td>${esc(t.table)}</td><td class="num">${t.records.length}</td>
    <td>${t.missingSystem.length || t.missingFields.length ? `<span class="badge warn">missing ${esc([...t.missingFields, ...t.missingSystem].map((f) => f.header).join(', '))}</span>` : '<span class="badge good">ready</span>'}</td></tr>`).join('');

  main.innerHTML = `
    <div class="page-head"><h1>Settings</h1></div>
    <div class="dash-col" style="max-width:900px">
      <section class="card">
        <div class="card-head"><h2>Workbooks</h2></div>
        <p class="small">${app.demo ? 'Demo mode (sample data).' : `Signed in as <strong>${esc(app.user.upn)}</strong>.`}</p>
        <table class="data-table">
          <tr><th>Active year</th><td><a href="${esc(app.books.active.webUrl || '#')}" target="_blank" rel="noopener">${esc(app.books.active.name)}</a></td></tr>
          <tr><th>Requests inbox</th><td>${app.books.requests ? `<a href="${esc(app.books.requests.webUrl || '#')}" target="_blank" rel="noopener">${esc(app.books.requests.name)}</a>` : '<span class="muted">not created yet</span>'}</td></tr>
          <tr><th>Past years</th><td>${app.books.archived.map((b) => esc(b.name)).join('<br>') || '<span class="muted">none</span>'}</td></tr>
        </table>
      </section>
      <section class="card">
        <div class="card-head"><h2>Academic year ${esc(app.books.active.year || '')}</h2>
          ${app.canWrite ? '<a class="btn primary" href="#/new-year">Start new academic year…</a>' : ''}</div>
        <table class="data-table"><thead><tr><th>Term</th><th>Dates</th><th>Status</th></tr></thead><tbody>
          ${a.tables.filter((t) => t.kind === 'term').map((t) => `<tr><td><a href="#/sessions" data-term="${esc(t.table)}">${esc(t.term?.termYear || t.sheet)}</a></td>
            <td class="small">${t.term?.from ? `${esc(fmtDate(t.term.from))} – ${esc(fmtDate(t.term.to))}` : '<span class="muted">not set</span>'}</td>
            <td>${t.closed ? `<span class="badge">🔒 Closed</span> <span class="small muted">${esc(t.term.closedBy || '')}</span>` : '<span class="badge good">Open</span>'}</td></tr>`).join('')}
        </tbody></table>
        <p class="small muted">Close a finished term from its tab on the Sessions page. Term dates and Closed flags are stored in the workbook’s hidden <code>Lists</code> sheet, so every computer sees the same thing. Starting a new year copies this workbook (rows emptied), confirms next year’s term dates, carries over sessions already booked for next year, and moves this workbook into <code>Archive/</code>.</p>
      </section>
      <section class="card">
        <div class="card-head"><h2>Sheets in ${esc(app.books.active.name)}</h2></div>
        <table class="data-table"><thead><tr><th>Sheet</th><th>Table</th><th class="num">Rows</th><th>Status</th></tr></thead><tbody>${sheetList || '<tr><td colspan="4" class="muted">No term or tours sheets are Excel tables yet.</td></tr>'}</tbody></table>
        <h2 style="margin-top:16px">Prepare workbook</h2>
        <p class="small muted">One-time setup, for the coordinator. It turns each term sheet into an Excel table and adds any missing columns (Term Year, College/School, Student Type, Actual Student Attendance, plus hidden _RecordId / _RowVersion / _ModifiedAt / _ModifiedBy). It also gives every row an id, fills blank College/Student Type cells with suggestions, and creates Instruction Requests.xlsx. On term sheets, “Entered into LibInsight” is hidden (its values are kept as _LegacyLibInsight), and the unlabeled stray column’s text is appended to Notes before that column becomes Actual Student Attendance. It also updates the Excel dropdowns. Other existing text is never overwritten. New columns are added at the right-hand end of each table, and you can drag them into place in Excel. <strong>Download a backup copy from SharePoint first.</strong></p>
        <button class="btn primary" id="btn-prepare" ${app.offline ? 'disabled' : ''}>${issues.length || !a.tables.length ? 'Prepare workbook…' : 'Run again (check only)'}</button>
        <div class="log" id="prep-log" hidden></div>
      </section>
      ${app.demo ? `<section class="card">
        <div class="card-head"><h2>Demo: try concurrent editing</h2></div>
        <p class="small muted">Open a session in the editor and change its Notes. Then come back here and press the button, which pretends a colleague on another PC changed the Notes of the first FALL 26 session. Save the editor and the conflict dialog appears. Changing a different field from the colleague merges silently.</p>
        <button class="btn" id="btn-sim">Simulate a colleague editing the first Fall session</button>
      </section>
      <section class="card">
        <div class="card-head"><h2>Demo: another computer starts the new year</h2></div>
        <p class="small muted">Pretends Lisa runs “Start new academic year” on her PC, slowed down to about a minute. While it runs, this page shows “Editing is paused” and refuses saves. When it finishes, this page switches to the new year by itself. Press F5 to sync immediately instead of waiting 30 seconds.</p>
        <button class="btn" id="btn-sim-year">Simulate Lisa starting the new year</button>
      </section>` : ''}
      <section class="card">
        <div class="card-head"><h2>Lists</h2></div>
        <p class="small muted">Dropdown values, librarian colors and the course-prefix → college map are in <code>config.js</code> on the web server. Edit that file to change them. Term dates come from the workbook’s hidden Lists sheet; <code>config.js</code> only supplies the defaults used when that sheet is first created.</p>
        <p class="small">Librarians: ${app.lists.librarians.map((l) => `<span class="swatch" style="background:var(--series-${l.color})"></span>${esc(l.name)}`).join(' &nbsp; ')}</p>
        <p class="small muted">After changing a list in <code>config.js</code>, update the dropdowns people see when they edit the workbook directly in Excel. This copies the lists into the Lists sheet and points each column’s dropdown at them. A value that isn’t in the list gets a warning in Excel (Location only suggests). Existing cells aren’t changed.</p>
        ${app.canWrite ? '<button class="btn" id="btn-dropdowns">Update Excel dropdowns</button><div class="log" id="dd-log" hidden></div>' : ''}
      </section>
    </div>`;

  $('#btn-prepare', main).onclick = async () => {
    if (!(await confirmDialog('Prepare workbook', `<p>This changes the structure of <strong>${esc(app.books.active.name)}</strong> on SharePoint (adds tables and columns). Have you saved a backup copy?</p>`, 'Yes, prepare it'))) return;
    const log = $('#prep-log', main);
    log.hidden = false;
    log.textContent = '';
    const write = (m) => { log.textContent += m + '\n'; log.scrollTop = log.scrollHeight; };
    $('#btn-prepare', main).disabled = true;
    try {
      const report = await app.store.prepare(app.lists, write);
      write('\nDone:\n• ' + report.join('\n• '));
      await app.loadRequests();
      await app.load(app.books.active.id);
      app.updateBanner();
      toast('Workbook prepared.');
    } catch (e) {
      write(`\nStopped: ${e.message}`);
      toast(e.message, 'error');
    } finally {
      $('#btn-prepare', main) && ($('#btn-prepare', main).disabled = false);
    }
  };

  $('#btn-dropdowns', main)?.addEventListener('click', async (e) => {
    const log = $('#dd-log', main);
    log.hidden = false;
    log.textContent = '';
    e.target.disabled = true;
    try {
      await app.store.checkWritable(app.books.active.id);
      const lines = await app.store.applyDropdowns(app.books.active, app.lists, (m) => (log.textContent += m + '\n'));
      toast(lines[0] || 'No lists to apply.');
    } catch (err) {
      log.textContent += `Stopped: ${err.message}\n`;
      toast(err.message, 'error');
    } finally {
      e.target.disabled = false;
    }
  });

  $$('[data-term]', main).forEach((a) => (a.onclick = () => { app.ui.sessions = { ...(app.ui.sessions || {}), table: a.dataset.term }; }));

  $('#btn-sim-year', main)?.addEventListener('click', async (e) => {
    e.target.disabled = true;
    const plan = { ...defaultPlan(), lists: app.lists };
    const lisa = new Store(app.memory, { name: 'Lisa', upn: 'lisa@desales.edu' });
    await lisa.init();
    const fast = app.memory.delay;
    app.memory.delay = 1200;
    toast('Lisa started the new academic year on her PC…');
    setTimeout(syncNow, 3000); // pick up the lock without waiting for the 30-second sync
    lisa.startNewYear(plan, (m) => console.info('[Lisa’s PC]', m))
      .then(() => { app.memory.delay = fast; toast('Lisa’s PC finished. This page switches on its next sync (or press F5).'); })
      .catch((err) => { app.memory.delay = fast; toast(`Lisa’s rollover failed: ${err.message}`, 'error'); });
  });

  $('#btn-sim', main)?.addEventListener('click', () => {
    const fall = app.active.tables.find((t) => /fall/i.test(t.sheet));
    const first = fall?.records.slice().sort((x, y) => (x.fields.date || '9').localeCompare(y.fields.date || '9'))[0];
    if (first && app.memory.simulateColleagueEdit(first.id, 'Notes', `Lisa: moved to the PC lab (${new Date().toLocaleTimeString()})`)) {
      toast(`Lisa changed the Notes on ${first.fields.course} (${first.fields.date}). Saving an edit to the same Notes will now show a conflict.`);
    }
  });
}
