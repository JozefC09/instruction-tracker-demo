// Requests inbox (spec §4.4): rows the faculty form adds to tbl_Requests in Instruction Requests.xlsx.
import { app } from './app.js';
import { toIsoDate, toInt, termForDate, parseTime, formatTimeRange } from './model.js';
import { $$, esc, toast, promptDialog, fmtDate } from './ui.js';
import { openEditor } from './sessions.js';

export function renderRequests(main) {
  const r = app.requests;
  app.ui.reqTab = app.ui.reqTab || 'New';
  if (!r) {
    main.innerHTML = `
      <div class="page-head"><h1>Requests</h1></div>
      <div class="card"><h2>No request inbox yet</h2>
        <p class="muted">Faculty requests from the website form are stored in <strong>Instruction Requests.xlsx</strong> in the Instruction folder. It hasn’t been created yet.</p>
        ${app.active.readOnly ? '' : '<button class="btn primary" id="btn-create-req">Create Instruction Requests.xlsx</button>'}
      </div>`;
    main.querySelector('#btn-create-req')?.addEventListener('click', async (e) => {
      e.target.disabled = true;
      try {
        await app.store.createRequestsWorkbook();
        await app.loadRequests();
        toast('Created Instruction Requests.xlsx.');
        app.render();
      } catch (err) {
        e.target.disabled = false;
        toast(err.message, 'error');
      }
    });
    return;
  }
  const tab = app.ui.reqTab;
  const items = r.items.filter((x) => tab === 'All' || x.status === tab)
    .sort((a, b) => String(b.values['Submitted At']).localeCompare(String(a.values['Submitted At'])));
  const count = (s) => r.items.filter((x) => s === 'All' || x.status === s).length;

  main.innerHTML = `
    <div class="page-head"><div><h1>Requests</h1><div class="muted small">From the faculty request form · <a href="${esc(app.books.requests?.webUrl || '#')}" target="_blank" rel="noopener">Open Instruction Requests.xlsx</a></div></div></div>
    <div class="tabs" role="tablist">${['New', 'Accepted', 'Declined', 'All'].map((s) => `<button role="tab" aria-selected="${tab === s}" data-rt="${s}">${s} <span class="muted small">${count(s)}</span></button>`).join('')}</div>
    <div class="req-list">
      ${items.map((x) => card(x)).join('') || '<div class="empty">Nothing here.</div>'}
    </div>`;

  $$('[data-rt]', main).forEach((b) => (b.onclick = () => { app.ui.reqTab = b.dataset.rt; renderRequests(main); }));
  $$('[data-accept]', main).forEach((b) => (b.onclick = () => accept(r.items.find((x) => x.id === b.dataset.accept))));
  $$('[data-decline]', main).forEach((b) => (b.onclick = () => decline(r.items.find((x) => x.id === b.dataset.decline))));
  $$('[data-copy]', main).forEach((b) => (b.onclick = async () => {
    try { await navigator.clipboard.writeText(b.dataset.copy); toast('Email address copied.'); } catch { toast(b.dataset.copy); }
  }));
}

function card(x) {
  const v = x.values;
  const show = (k, label = k) => (v[k] !== '' && v[k] != null ? `<dt>${esc(label)}</dt><dd>${esc(k.includes('Date') ? fmtDate(toIsoDate(v[k])) || v[k] : v[k])}</dd>` : '');
  const badge = x.status === 'New' ? 'warn' : x.status === 'Accepted' ? 'good' : '';
  const editable = app.canWrite;
  return `<article class="card req">
    <div>
      <div class="row"><h2>${esc(v['Course #'] || 'Course not given')} · ${esc(v['Faculty Name'])}</h2><span class="badge ${badge}">${esc(x.status)}</span></div>
      <div class="small muted">${esc(v['Request ID'])}${v['Submitted At'] ? ` · submitted ${esc(new Date(v['Submitted At']).toLocaleString())}` : ''}</div>
      <dl>${show('Preferred Date')}${show('Alternate Date')}${show('Time')}${show('Session Format', 'Format')}${show('Location')}${show('# Students', 'Students')}${show('College/School')}${show('Student Type')}${show('Notes / Goals', 'Goals')}${show('Handled By')}</dl>
    </div>
    <div class="req-actions">
      ${x.status === 'New' && editable ? `<button class="btn primary" data-accept="${esc(x.id)}">Accept…</button><button class="btn" data-decline="${esc(x.id)}">Decline</button>` : ''}
      ${v['Faculty Email'] ? `<button class="btn" data-copy="${esc(v['Faculty Email'])}">Copy email</button>` : ''}
    </div>
  </article>`;
}

/** Accept → New Session prefilled from the request, in the term its preferred date falls in (spec F15). */
function accept(x) {
  const v = x.values;
  const date = toIsoDate(v['Preferred Date']);
  const termName = termForDate(date, app.lists.terms);
  const terms = app.active.tables.filter((t) => t.kind === 'term' && !t.missingSystem.length && !t.closed);
  const table = terms.find((t) => t.sheet.toUpperCase() === termName.toUpperCase()) ?? terms[0];
  if (!table) return toast('No open term sheet to add the session to. Reopen a term, or start the new academic year.', 'error');
  if (termName && table.sheet.toUpperCase() !== termName.toUpperCase()) {
    toast(`The preferred date falls in ${termName}, which is closed or not in this workbook. The session form opens in ${table.sheet}; change the date or term as needed.`);
  }
  const t = parseTime(v.Time);
  app.viewBookId = app.books.active.id;
  openEditor(null, {
    table: table.table,
    stage: 'all',
    prefill: {
      date,
      time: t ? formatTimeRange(t.start, t.end) : String(v.Time || ''),
      course: v['Course #'],
      faculty: v['Faculty Name'],
      college: v['College/School'],
      studentType: v['Student Type'],
      format: v['Session Format'],
      location: v.Location,
      expected: toInt(v['# Students']),
      notes: [v['Notes / Goals'], v['Alternate Date'] ? `Alternate date: ${v['Alternate Date']}` : '', `Request ${v['Request ID']}`].filter(Boolean).join('\n'),
      room: 'Pending',
    },
    onSaved: async (id) => {
      await app.store.updateRequest(x.id, {
        Status: 'Accepted', 'Linked Record ID': id, 'Linked Workbook': app.books.active.name,
        'Handled By': app.user.upn, 'Handled At': new Date().toISOString(),
      });
      await app.loadRequests();
    },
  });
}

async function decline(x) {
  const note = await promptDialog('Decline request', 'Optional note (kept with the request; no email is sent)');
  if (note === null) return;
  try {
    const goals = x.values['Notes / Goals'];
    await app.store.updateRequest(x.id, {
      Status: 'Declined', 'Handled By': app.user.upn, 'Handled At': new Date().toISOString(),
      ...(note ? { 'Notes / Goals': `${goals ? goals + '\n' : ''}Declined: ${note}` } : {}),
    });
    await app.loadRequests();
    toast('Request declined. Use “Copy email” to reply from Outlook.');
    app.render();
  } catch (e) {
    toast(e.message, 'error');
  }
}
