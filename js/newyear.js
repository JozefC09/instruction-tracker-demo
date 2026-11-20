// "Start new academic year" wizard (spec F32, §8): close the year, create next year's workbook from this one,
// confirm term dates, carry over sessions already booked for next year, archive the old workbook.
import { app } from './app.js';
import { needsAttention, nextSheetName, nextYearLabel, shiftYear, termYearFromSheet, isTermSheet } from './model.js';
import { $, $$, esc, toast, confirmDialog, fmtDate, fmtDateShort } from './ui.js';

function plan() {
  const data = app.active;
  if (app.ui.newYear?.forBook === data.book.id) return app.ui.newYear;
  const termTables = data.tables.filter((t) => t.kind === 'term');
  const terms = termTables.map((t) => {
    const meta = t.term ?? data.terms.find((x) => x.sheet.toUpperCase() === t.sheet.toUpperCase());
    const cfgTerm = app.cfg.lists.terms.find((x) => x.name.toUpperCase() === t.sheet.toUpperCase());
    const from = meta?.from || cfgTerm?.from || '';
    const to = meta?.to || cfgTerm?.to || '';
    const sheet = nextSheetName(t.sheet);
    return { fromSheet: t.sheet, oldLabel: meta?.termYear || termYearFromSheet(t.sheet), oldFrom: from, oldTo: to, sheet, termYear: termYearFromSheet(sheet), from: shiftYear(from), to: shiftYear(to) };
  });
  const year = data.book.year || data.book.name.slice(0, 9);
  app.ui.newYear = { forBook: data.book.id, closeOpen: true, terms, newName: `${nextYearLabel(year)} Instruction Classes.xlsx`, unchecked: new Set() };
  return app.ui.newYear;
}

/** Sessions in this year's workbook dated on or after the new year's first term start. */
function carryOver(p) {
  const starts = p.terms.map((t) => t.from).filter(Boolean).sort();
  if (!starts.length) return [];
  return app.active.records.filter((r) => r.id && r.fields.date && r.fields.date >= starts[0])
    .sort((a, b) => a.fields.date.localeCompare(b.fields.date));
}

/** The plan with every default accepted (used by the demo's "another PC" simulation). */
export function defaultPlan() {
  const p = plan();
  return {
    newName: p.newName,
    closeOpen: true,
    terms: p.terms.map(({ fromSheet, sheet, termYear, from, to }) => ({ fromSheet, sheet, termYear, from, to })),
    moveIds: carryOver(p).map((r) => r.id),
  };
}

function problems(p) {
  const out = [];
  const names = p.terms.map((t) => t.sheet.trim().toUpperCase());
  p.terms.forEach((t) => {
    if (!isTermSheet(t.sheet)) out.push(`“${t.sheet}” isn’t a term sheet name like “FALL 27”.`);
    if (!t.from || !t.to) out.push(`${t.sheet}: start and end dates are required.`);
    else if (t.from > t.to) out.push(`${t.sheet}: the start date is after the end date.`);
  });
  if (new Set(names).size !== names.length) out.push('Two terms have the same sheet name.');
  if (app.books.archived.some((b) => b.name === p.newName) || app.books.active.name === p.newName) out.push(`${p.newName} already exists.`);
  return out;
}

export function renderNewYear(main) {
  if (app.data !== app.active) app.viewBookId = app.books.active.id;
  const data = app.active;
  if (!app.canWrite) {
    main.innerHTML = `<div class="page-head"><h1>Start new academic year</h1></div><div class="card"><p>The workbook can’t be changed right now${app.locked ? ` (${esc(app.locked.user)} is ${esc(app.locked.action.toLowerCase())})` : app.offline ? ' (offline)' : ' (it hasn’t been prepared yet; see Settings)'}.</p><a class="btn" href="#/settings">Back to Settings</a></div>`;
    return;
  }
  const p = plan();
  const termTables = data.tables.filter((t) => t.kind === 'term');
  const openTerms = termTables.filter((t) => !t.closed);
  const moves = carryOver(p);
  const chosen = moves.filter((r) => !p.unchecked.has(r.id));
  const issues = problems(p);
  const blocked = issues.length || (openTerms.length && !p.closeOpen);

  main.innerHTML = `
    <div class="page-head">
      <div><h1>Start new academic year</h1><div class="muted small">${esc(data.book.name)} → ${esc(p.newName)}</div></div>
      <a class="btn" href="#/settings">Cancel</a>
    </div>
    <div class="dash-col" style="max-width:920px">
      <section class="card">
        <div class="card-head"><h2>1 · Close ${esc(data.book.year || 'this year')}</h2></div>
        <table class="data-table"><thead><tr><th>Term</th><th>Dates</th><th>Status</th><th class="num">Need attention</th></tr></thead><tbody>
          ${termTables.map((t) => {
            const n = t.records.filter((r) => needsAttention(r, app.today).length).length;
            return `<tr><td>${esc(t.term?.termYear || termYearFromSheet(t.sheet))}</td><td class="small">${t.term?.from ? `${esc(fmtDate(t.term.from))} – ${esc(fmtDate(t.term.to))}` : '—'}</td>
              <td>${t.closed ? '<span class="badge">🔒 Closed</span>' : '<span class="badge warn">Open</span>'}</td><td class="num">${n || ''}</td></tr>`;
          }).join('')}
        </tbody></table>
        ${openTerms.length ? `<label class="row" style="margin-top:10px"><input type="checkbox" id="cb-close" ${p.closeOpen ? 'checked' : ''}> Close the ${openTerms.length} open term${openTerms.length === 1 ? '' : 's'} as part of this (required)</label>` : '<p class="small muted">All terms are closed.</p>'}
      </section>

      <section class="card">
        <div class="card-head"><h2>2 · Confirm next year’s terms</h2></div>
        <p class="small muted">The new workbook is a copy of this one with the rows emptied, so the columns, dropdowns, formatting and hidden Lists sheet carry over. The dates below are this year’s, moved forward one year. Adjust them to the academic calendar.</p>
        <table class="data-table"><thead><tr><th>This year</th><th>New sheet</th><th>Term Year</th><th>Start</th><th>End</th></tr></thead><tbody>
          ${p.terms.map((t, i) => `<tr>
            <td class="small">${esc(t.fromSheet)}</td>
            <td><input type="text" data-i="${i}" data-k="sheet" value="${esc(t.sheet)}" style="width:110px" aria-label="New sheet name for ${esc(t.fromSheet)}"></td>
            <td><input type="text" data-i="${i}" data-k="termYear" value="${esc(t.termYear)}" style="width:120px" aria-label="Term Year"></td>
            <td><input type="date" data-i="${i}" data-k="from" value="${esc(t.from)}" aria-label="Start"></td>
            <td><input type="date" data-i="${i}" data-k="to" value="${esc(t.to)}" aria-label="End"></td></tr>`).join('')}
        </tbody></table>
      </section>

      <section class="card">
        <div class="card-head"><h2>3 · Sessions already booked for next year</h2>
          ${moves.length ? '<div class="row"><button class="btn ghost small" id="mv-all">Select all</button><button class="btn ghost small" id="mv-none">Select none</button></div>' : ''}</div>
        ${moves.length ? `<p class="small muted">These are dated on or after ${esc(fmtDate(p.terms.map((t) => t.from).filter(Boolean).sort()[0]))}. Checked sessions move into the new workbook (same record id) and are removed from this one.</p>
          <div class="grid-wrap" style="max-height:300px"><table class="grid"><tbody>
          ${moves.map((r) => `<tr><td style="width:32px"><input type="checkbox" data-mv="${esc(r.id)}" ${p.unchecked.has(r.id) ? '' : 'checked'} aria-label="Move ${esc(r.fields.course)}"></td>
            <td>${esc(fmtDateShort(r.fields.date))}</td><td>${esc(r.fields.course || r.fields.faculty)}</td><td>${esc(r.fields.faculty)}</td><td class="muted">${esc(r.kind === 'tours' ? 'Tours' : r.sheet)}</td></tr>`).join('')}
          </tbody></table></div>` : '<p class="small muted">None. No sessions in this workbook are dated in the new year.</p>'}
      </section>

      <section class="card">
        <div class="card-head"><h2>4 · Start</h2></div>
        <ol class="small">
          ${openTerms.length && p.closeOpen ? `<li>Close ${openTerms.map((t) => esc(t.term?.termYear || termYearFromSheet(t.sheet))).join(', ')}.</li>` : ''}
          <li>Create <strong>${esc(p.newName)}</strong> in the Instruction folder with sheets ${p.terms.map((t) => esc(t.sheet)).join(', ')}${data.tables.some((t) => t.kind === 'tours') ? ', plus the tours sheet' : ''}.</li>
          ${chosen.length ? `<li>Move ${chosen.length} session${chosen.length === 1 ? '' : 's'} into it.</li>` : ''}
          <li>Move <strong>${esc(data.book.name)}</strong> into <code>Archive/</code>. It stays a normal Excel file and remains on the dashboard as a past year, read-only.</li>
          <li><code>Instruction Requests.xlsx</code> stays where it is, so the request form keeps working.</li>
        </ol>
        <p class="small muted">Other computers pause editing while this runs (usually a minute or two) and switch to the new year on their next sync. Download a backup of the workbook from SharePoint first if you haven’t recently.</p>
        ${issues.length ? `<ul class="err">${issues.map((x) => `<li>${esc(x)}</li>`).join('')}</ul>` : ''}
        <button class="btn primary" id="btn-start" ${blocked ? 'disabled' : ''}>Start ${esc(p.newName.slice(0, 9))}</button>
        <div class="log" id="ny-log" hidden></div>
      </section>
    </div>`;

  const rerender = () => {
    const y = window.scrollY;
    renderNewYear(main);
    window.scrollTo(0, y);
  };
  $('#cb-close', main)?.addEventListener('change', (e) => { p.closeOpen = e.target.checked; rerender(); });
  $$('input[data-i]', main).forEach((inp) => inp.addEventListener('change', () => {
    const t = p.terms[Number(inp.dataset.i)];
    t[inp.dataset.k] = inp.value.trim();
    if (inp.dataset.k === 'sheet') t.termYear = termYearFromSheet(t.sheet) || t.termYear;
    rerender();
  }));
  $$('[data-mv]', main).forEach((cb) => cb.addEventListener('change', () => {
    cb.checked ? p.unchecked.delete(cb.dataset.mv) : p.unchecked.add(cb.dataset.mv);
    rerender();
  }));
  $('#mv-all', main)?.addEventListener('click', () => { p.unchecked.clear(); rerender(); });
  $('#mv-none', main)?.addEventListener('click', () => { moves.forEach((r) => p.unchecked.add(r.id)); rerender(); });

  $('#btn-start', main).onclick = async () => {
    if (!(await confirmDialog(`Start ${p.newName.slice(0, 9)}?`, `<p>This creates <strong>${esc(p.newName)}</strong>${chosen.length ? `, moves ${chosen.length} session${chosen.length === 1 ? '' : 's'}` : ''} and archives <strong>${esc(data.book.name)}</strong>. Everyone will switch to the new year.</p>`, 'Start new year'))) return;
    const log = $('#ny-log', main);
    log.hidden = false;
    const write = (m) => { log.textContent += m + '\n'; log.scrollTop = log.scrollHeight; };
    $$('button, input', main).forEach((el) => (el.disabled = true));
    try {
      if (!data.hasLists) {
        write('Creating the hidden Lists sheet first…');
        await app.store.ensureLists(data.book, app.store.defaultTermsFor(data, app.cfg.lists.terms));
      }
      await app.store.startNewYear({
        newName: p.newName,
        closeOpen: p.closeOpen,
        terms: p.terms.map(({ fromSheet, sheet, termYear, from, to }) => ({ fromSheet, sheet: sheet.trim(), termYear, from, to })),
        moveIds: chosen.map((r) => r.id),
        lists: app.lists,
      }, write);
      app.ui.newYear = null;
      await app.rediscover();
      app.updateBanner();
      write('\nOpening the new year…');
      setTimeout(() => { location.hash = '#/dashboard'; app.render(); }, 1200);
    } catch (e) {
      console.error(e);
      write(`\nStopped: ${e.message}\n${e.archived
        ? 'The new workbook is ready and the old one is marked archived, but it couldn’t be moved into Archive/. Move it there by hand in SharePoint. The app already treats the new year as active.'
        : 'The old workbook is unlocked and still active. Nothing was archived.'}`);
      toast(e.message, 'error');
      $$('button, input', main).forEach((el) => (el.disabled = false));
      try { await app.rediscover(); await app.load(app.books.active.id); app.updateBanner(); } catch { /* keep the log on screen */ }
    }
  };
}
