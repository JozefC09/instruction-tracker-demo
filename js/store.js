// Business logic over a backend (GraphBackend for SharePoint, MemoryBackend for the demo):
// loading workbooks, optimistic-concurrency saves with field-level merge, adds, deletes, moves,
// the requests inbox and the one-time "Prepare workbook" step.
import {
  FIELDS, SYSTEM_FIELDS, mapHeaders, rowsToRecords, recordToRow, mergeFields, newRecordId,
  termYearFromSheet, isTermSheet, isToursSheet, suggestFromCourse, toIsoDate, isoToExcelText, termForDate, tableNameFor,
  LEGACY_LIBINSIGHT, isLibInsightHeader, isStrayHeader, appendToNotes, dropdownLists,
} from './model.js';
import { colLetters } from './graph.js';

/** Dropdown lists live on the Lists sheet from column M (A–G = tbl_Terms, I–K = tbl_AppLock). */
const DROPDOWN_COL = 12;
const DROPDOWN_ROWS = 100;

/** Hidden "Lists" sheet (spec §3.3): term dates + Closed flag, and the advisory lock for transitions (spec §5.6). */
export const TERMS_HEADERS = ['Sheet', 'Term Year', 'Start', 'End', 'Closed', 'Closed By', 'Closed At'];
export const LOCK_HEADERS = ['Action', 'User', 'Since'];
const LOCK_STALE_MS = 10 * 60 * 1000;

/** True when someone else holds the lock (or the workbook has been archived). */
export function lockedByOther(lock, upn) {
  if (!lock?.action) return false;
  if (lock.action === 'Archived') return true;
  if (String(lock.user).toLowerCase() === String(upn).toLowerCase()) return false;
  return Date.now() - Date.parse(lock.since) < LOCK_STALE_MS;
}

export class ReadOnlyError extends Error {}

/** Instruction Requests.xlsx columns: the Phase 2 integration contract (spec §3.5). */
export const REQUEST_HEADERS = [
  'Request ID', 'Submitted At', 'Faculty Name', 'Faculty Email', 'Course #', 'College/School', 'Student Type',
  'Preferred Date', 'Alternate Date', 'Time', 'Session Format', 'Location', '# Students', 'Notes / Goals',
  'Status', 'Linked Record ID', 'Linked Workbook', 'Handled By', 'Handled At',
];

const TOURS_FIELDS = new Set(['date', 'time', 'faculty', 'librarian', 'expected', 'notes', 'actual']);

function classify(sheet) {
  if (isTermSheet(sheet)) return 'term';
  if (isToursSheet(sheet)) return 'tours';
  return null;
}

export class Store {
  constructor(backend, user) {
    this.backend = backend;
    this.user = user; // { name, upn }
    this.books = null;
    this.tables = new Map(); // `${bookId}/${table}` → last read { headers, ... }
    this.meta = new Map(); // bookId → { terms, lock } from the Lists sheet at the last load
  }

  /** Re-reads the lock and term status from the workbook, then applies assertWritable. */
  async checkWritable(bookId, tableName) {
    const meta = this.meta.get(bookId);
    if (meta?.hasLists) {
      const [tt, lt] = await Promise.all([this.backend.readTable(bookId, 'tbl_Terms'), this.backend.readTable(bookId, 'tbl_AppLock')]);
      this.meta.set(bookId, this.parseMeta([tt, lt]));
    }
    this.assertWritable(bookId, tableName);
  }

  /** Refuses writes to past years, closed terms, or while another PC runs a transition. */
  assertWritable(bookId, tableName) {
    if (bookId !== this.books.active.id) throw new ReadOnlyError('Past academic years are read-only.');
    const meta = this.meta.get(bookId);
    if (lockedByOther(meta?.lock, this.user.upn)) {
      throw new ReadOnlyError(meta.lock.action === 'Archived' ? 'This workbook has been archived. Reload to switch to the new year.' : `Editing is paused: ${meta.lock.user} is ${meta.lock.action.toLowerCase()}.`);
    }
    const t = this.tables.get(`${bookId}/${tableName}`);
    const term = t && meta?.terms.find((x) => x.sheet.toUpperCase() === t.sheet.toUpperCase());
    if (term?.closed) throw new ReadOnlyError(`${term.termYear || term.sheet} is closed. Reopen the term to edit it.`);
  }

  async init() {
    this.books = await this.backend.discover();
    return this.books;
  }

  isReadOnly(book) {
    return book.id !== this.books.active.id;
  }

  ctxFor(t) {
    const kind = classify(t.sheet);
    return { table: t.name, sheet: t.sheet, kind, termYear: kind === 'term' ? termYearFromSheet(t.sheet) : '' };
  }

  /** Reads every term and tours table of a workbook (one Graph call per table). */
  async loadBook(book) {
    const raw = await this.backend.readTables(book.id);
    const meta = this.parseMeta(raw);
    this.meta.set(book.id, meta);
    const tables = [];
    for (const t of raw) {
      const ctx = this.ctxFor(t);
      if (!ctx.kind) continue;
      this.tables.set(`${book.id}/${t.name}`, t);
      const map = mapHeaders(t.headers);
      const wanted = ctx.kind === 'tours' ? FIELDS.filter((f) => TOURS_FIELDS.has(f.key)) : FIELDS;
      const missingFields = wanted.filter((f) => map[f.key] === undefined && !(ctx.kind === 'tours' && ['notes', 'actual'].includes(f.key)));
      const missingSystem = SYSTEM_FIELDS.filter((f) => map[f.key] === undefined);
      let records = rowsToRecords(t.headers, t.rows, ctx).filter((r) => !r.blank);
      const term = meta.terms.find((x) => x.sheet.toUpperCase() === t.sheet.toUpperCase());

      // Rows added directly in Excel get an id on the next sync (spec §3.2).
      if (!this.isReadOnly(book) && !missingSystem.length && !term?.closed && !lockedByOther(meta.lock, this.user.upn) && records.some((r) => !r.id)) {
        await this.assignIds(book, t);
        const fresh = await this.backend.readTable(book.id, t.name);
        Object.assign(t, fresh);
        records = rowsToRecords(t.headers, t.rows, ctx).filter((r) => !r.blank);
      }
      tables.push({ ...ctx, headers: t.headers, records, missingFields, missingSystem, map, term, closed: !!term?.closed });
    }
    const order = (x) => (x.kind === 'tours' ? 99 : ['FALL', 'WINTER', 'SPRING', 'SUMMER'].indexOf(x.sheet.toUpperCase().split(/\s/)[0]));
    tables.sort((a, b) => order(a) - order(b));
    return { book, tables, records: tables.flatMap((t) => t.records), readOnly: this.isReadOnly(book), terms: meta.terms, lock: meta.lock, hasLists: meta.hasLists };
  }

  /** Reads tbl_Terms and tbl_AppLock out of an already-loaded set of tables (no extra Graph calls). */
  parseMeta(raw) {
    const find = (n) => raw.find((t) => t.name.toLowerCase() === n.toLowerCase());
    const tt = find('tbl_Terms'), lt = find('tbl_AppLock');
    const col = (t, h) => t.headers.findIndex((x) => String(x).trim().toLowerCase() === h.toLowerCase());
    const terms = tt ? tt.rows.map((r, index) => ({
      index,
      sheet: String(r[col(tt, 'Sheet')] ?? '').trim(),
      termYear: String(r[col(tt, 'Term Year')] ?? '').trim(),
      from: toIsoDate(r[col(tt, 'Start')]),
      to: toIsoDate(r[col(tt, 'End')]),
      closed: /^y/i.test(String(r[col(tt, 'Closed')] ?? '')),
      closedBy: String(r[col(tt, 'Closed By')] ?? ''),
      closedAt: String(r[col(tt, 'Closed At')] ?? ''),
    })).filter((x) => x.sheet) : [];
    const lr = lt?.rows[0];
    const lock = lr && lr[0] ? { action: String(lr[0]), user: String(lr[1] ?? ''), since: String(lr[2] ?? '') } : null;
    return { terms, lock, hasLists: !!(tt && lt) };
  }

  // ---------- Lists sheet, term status, lock ----------

  /** Creates the hidden Lists sheet with tbl_Terms and tbl_AppLock if they don't exist yet. */
  async ensureLists(book, defaultTerms) {
    const raw = await this.backend.readTables(book.id);
    const has = (n) => raw.some((t) => t.name.toLowerCase() === n.toLowerCase());
    if (has('tbl_Terms') && has('tbl_AppLock')) return false;
    const sheets = await this.backend.listSheets(book.id);
    if (!sheets.includes('Lists')) await this.backend.addWorksheet(book.id, 'Lists');
    if (!has('tbl_Terms')) {
      const rows = defaultTerms.map((t) => [t.sheet, t.termYear, isoToExcelText(t.from), isoToExcelText(t.to), 'No', '', '']);
      const body = rows.length ? rows : [TERMS_HEADERS.map(() => '')];
      await this.backend.writeRange(book.id, 'Lists', `A1:G${body.length + 1}`, [TERMS_HEADERS, ...body]);
      await this.backend.createTableAt(book.id, `Lists!A1:G${body.length + 1}`, 'tbl_Terms');
    }
    if (!has('tbl_AppLock')) {
      await this.backend.writeRange(book.id, 'Lists', 'I1:K2', [LOCK_HEADERS, ['', '', '']]);
      await this.backend.createTableAt(book.id, 'Lists!I1:K2', 'tbl_AppLock');
    }
    await this.backend.hideSheet(book.id, 'Lists');
    return true;
  }

  /** Term rows for ensureLists: every term sheet in the workbook, with dates from config where known. */
  defaultTermsFor(data, configTerms) {
    return data.tables.filter((t) => t.kind === 'term').map((t) => {
      const c = configTerms.find((x) => x.name.toUpperCase() === t.sheet.toUpperCase());
      return { sheet: t.sheet, termYear: termYearFromSheet(t.sheet), from: c?.from ?? '', to: c?.to ?? '' };
    });
  }

  async acquireLock(book, action) {
    const t = await this.backend.readTable(book.id, 'tbl_AppLock');
    const r = t.rows[0];
    const lock = r && r[0] ? { action: String(r[0]), user: String(r[1]), since: String(r[2]) } : null;
    if (lockedByOther(lock, this.user.upn)) throw new ReadOnlyError(`${lock.user} is already ${lock.action.toLowerCase()} (since ${new Date(lock.since).toLocaleTimeString()}). Try again when they’re done.`);
    const row = [action, this.user.upn, new Date().toISOString()];
    if (t.rows.length) await this.backend.writeRow(book.id, t, 0, row);
    else await this.backend.addRow(book.id, 'tbl_AppLock', row);
  }

  async releaseLock(book, action = '') {
    const t = await this.backend.readTable(book.id, 'tbl_AppLock');
    const row = action ? [action, this.user.upn, new Date().toISOString()] : ['', '', ''];
    if (t.rows.length) await this.backend.writeRow(book.id, t, 0, row);
  }

  /** Close or reopen a term (spec F31): sets Closed in tbl_Terms. The sheet stays in the workbook. */
  async setTermClosed(book, sheet, closed, { lock = true } = {}) {
    if (book.id !== this.books.active.id) throw new ReadOnlyError('Past academic years are read-only.');
    if (lock) await this.acquireLock(book, closed ? 'Closing a term' : 'Reopening a term');
    try {
      const t = await this.backend.readTable(book.id, 'tbl_Terms');
      const ci = (h) => t.headers.findIndex((x) => String(x).toLowerCase() === h.toLowerCase());
      const idx = t.rows.findIndex((r) => String(r[ci('Sheet')]).trim().toUpperCase() === sheet.toUpperCase());
      const row = idx >= 0 ? t.rows[idx].slice() : TERMS_HEADERS.map(() => '');
      row[ci('Sheet')] = sheet;
      if (!row[ci('Term Year')]) row[ci('Term Year')] = termYearFromSheet(sheet);
      row[ci('Closed')] = closed ? 'Yes' : 'No';
      row[ci('Closed By')] = closed ? this.user.upn : '';
      row[ci('Closed At')] = closed ? new Date().toISOString() : '';
      if (idx >= 0) await this.backend.writeRow(book.id, t, idx, row);
      else {
        const blankIdx = t.rows.findIndex((r) => r.every((v) => v === '' || v == null));
        if (blankIdx >= 0) await this.backend.writeRow(book.id, t, blankIdx, row);
        else await this.backend.addRow(book.id, 'tbl_Terms', row);
      }
    } finally {
      if (lock) await this.releaseLock(book);
    }
  }

  // ---------- Start new academic year (spec F32, §8) ----------

  /**
   * plan = {
   *   newName: '2027-2028 Instruction Classes.xlsx',
   *   closeOpen: true,                       // close terms still open in the old year first
   *   terms: [{ fromSheet: 'FALL 26', sheet: 'FALL 27', termYear: 'Fall 2027', from, to }],
   *   moveIds: ['…']                          // sessions already booked for the new year, moved across
   * }
   * Each step is logged. If it stops part-way, the old workbook is unlocked and stays active.
   */
  async startNewYear(plan, log = () => {}) {
    const old = this.books.active;
    const step = async (msg, fn) => { log(`${msg}…`); return fn(); };
    await step(`Locking ${old.name} so other computers pause`, () => this.acquireLock(old, 'Starting a new academic year'));
    let archived = false;
    try {
      // Check first, so a refused run changes nothing.
      if (await this.backend.findInFolder(plan.newName)) {
        throw new Error(`“${plan.newName}” already exists in the Instruction folder. If it’s left over from an earlier attempt, check it, then delete or rename it and run this again.`);
      }
      if (plan.closeOpen) {
        const meta = this.parseMeta(await this.backend.readTables(old.id));
        for (const t of meta.terms.filter((x) => !x.closed)) await step(`Closing ${t.termYear || t.sheet}`, () => this.setTermClosed(old, t.sheet, true, { lock: false }));
      }
      const nb = await step(`Copying ${old.name} to ${plan.newName} (keeps formatting and dropdowns)`, () => this.backend.copyWorkbook(old.id, plan.newName));

      // Empty and rename the copied sheets.
      const copied = await this.backend.readTables(nb.id);
      const bySheet = new Map(plan.terms.map((t) => [t.fromSheet.toUpperCase(), t]));
      for (const t of copied) {
        const kind = classify(t.sheet);
        const lname = t.name.toLowerCase();
        if (kind === 'term') {
          const target = bySheet.get(t.sheet.toUpperCase());
          await step(`${t.sheet}: clearing rows`, () => this.backend.clearTableRows(nb.id, t));
          if (target && target.sheet !== t.sheet) {
            await step(`Renaming ${t.sheet} → ${target.sheet}`, async () => {
              await this.backend.renameSheet(nb.id, t.sheet, target.sheet);
              await this.backend.renameTable(nb.id, t.name, tableNameFor(target.sheet));
            });
          }
        } else if (kind === 'tours') {
          await step(`${t.sheet}: clearing rows`, () => this.backend.clearTableRows(nb.id, t));
        } else if (lname === 'tbl_terms') {
          await step('Writing the new term dates', async () => {
            await this.backend.clearTableRows(nb.id, t);
            const fresh = await this.backend.readTable(nb.id, t.name);
            const rows = plan.terms.map((x) => [x.sheet, x.termYear, isoToExcelText(x.from), isoToExcelText(x.to), 'No', '', '']);
            if (rows.length) await this.backend.writeRow(nb.id, fresh, 0, rows[0]);
            for (const r of rows.slice(1)) await this.backend.addRow(nb.id, t.name, r);
          });
        } else if (lname === 'tbl_applock') {
          // The copy was taken while we held the lock: clear it in the new workbook.
          await this.backend.writeRow(nb.id, t, 0, ['', '', '']);
        }
      }

      // The Term Year dropdown has to list the new year's terms. Not worth stopping the rollover for.
      if (plan.lists) {
        try {
          await this.applyDropdowns(nb, plan.lists, log, plan.terms.map((x) => x.termYear));
        } catch (e) {
          log(`Couldn’t update the Excel dropdowns (${e.message}). Use Settings → Update Excel dropdowns afterwards.`);
        }
      }

      // Sessions already booked for the new year move across, keeping their record id (spec §8 step 5).
      if (plan.moveIds?.length) {
        const oldData = await this.loadBook(old);
        const newTables = (await this.backend.readTables(nb.id)).map((t) => ({ ...t, kind: classify(t.sheet) }));
        const newTerms = plan.terms.map((x) => ({ name: x.sheet, from: x.from, to: x.to, termYear: x.termYear }));
        let moved = 0;
        for (const id of plan.moveIds) {
          const rec = oldData.records.find((r) => r.id === id);
          if (!rec) continue;
          let target;
          if (rec.kind === 'tours') target = newTables.find((t) => t.kind === 'tours');
          else {
            const name = termForDate(rec.fields.date, newTerms) || (rec.fields.date < newTerms[0]?.from ? newTerms[0]?.name : newTerms.at(-1)?.name);
            target = newTables.find((t) => t.kind === 'term' && t.sheet.toUpperCase() === String(name).toUpperCase());
          }
          if (!target) continue;
          const termYear = rec.kind === 'tours' ? rec.fields.termYear : newTerms.find((x) => x.name.toUpperCase() === target.sheet.toUpperCase())?.termYear;
          const row = recordToRow(target.headers, null, { ...rec.fields, termYear }, { id: rec.id, ...this.sys(rec.version + 1) });
          await this.backend.addRow(nb.id, target.name, row);
          // Delete from the old year only after the copy is safely in the new one.
          const { t, cur } = await this.freshRecord(old.id, rec);
          if (cur) await this.backend.deleteRow(old.id, t, cur.index);
          moved++;
        }
        log(`Moved ${moved} session${moved === 1 ? '' : 's'} into ${plan.newName}.`);
      }

      await step(`Marking ${old.name} as archived`, () => this.releaseLock(old, 'Archived'));
      archived = true;
      const archiveId = await step('Finding the Archive folder', () => this.backend.ensureFolder('Archive'));
      await step(`Moving ${old.name} into Archive`, () => this.backend.moveItem(old.id, archiveId));
      this.books = await this.backend.discover();
      log(`Done. ${this.books.active.name} is now the active year.`);
      return this.books;
    } catch (e) {
      if (!archived) {
        try { await this.releaseLock(old); } catch { /* best effort; a stale lock expires after 10 minutes */ }
      }
      e.archived = archived;
      throw e;
    }
  }

  async assignIds(book, t) {
    const map = mapHeaders(t.headers);
    const ids = t.rows.map((r) => r[map.id] || newRecordId());
    const versions = t.rows.map((r) => r[map.version] || 1);
    await this.backend.writeColumn(book.id, t, map.id, ids);
    await this.backend.writeColumn(book.id, t, map.version, versions);
  }

  async freshRecord(bookId, rec) {
    const t = await this.backend.readTable(bookId, rec.table);
    this.tables.set(`${bookId}/${t.name}`, t);
    const recs = rowsToRecords(t.headers, t.rows, this.ctxFor(t));
    return { t, cur: recs.find((r) => r.id === rec.id) ?? null };
  }

  sys(version) {
    return { version, modifiedAt: new Date().toISOString(), modifiedBy: this.user.upn };
  }

  /**
   * Optimistic save (spec §5.2). base = fields when the editor opened; mine = edited fields.
   * resolutions = { fieldKey: 'mine' | 'theirs' } from the conflict dialog.
   * Returns { status: 'saved' | 'conflict' | 'deleted', ... }.
   */
  async save(bookId, rec, base, mine, resolutions = {}) {
    await this.checkWritable(bookId, rec.table);
    const { t, cur } = await this.freshRecord(bookId, rec);
    if (!cur) return { status: 'deleted' };
    const { merged, conflicts } = mergeFields(base, mine, cur.fields);
    const open = conflicts.filter((c) => !(c.key in resolutions));
    if (open.length) return { status: 'conflict', conflicts: open, theirs: cur };
    for (const c of conflicts) merged[c.key] = resolutions[c.key] === 'mine' ? mine[c.key] : cur.fields[c.key];
    const row = recordToRow(t.headers, cur, merged, { id: cur.id, ...this.sys(cur.version + 1) });
    await this.backend.writeRow(bookId, t, cur.index, row);
    const mergedOthers = cur.version !== rec.version || cur.hash !== rec.hash;
    return { status: 'saved', mergedOthers };
  }

  /** extra = { header: value } for columns the app doesn't edit (e.g. _LegacyLibInsight on a moved row). */
  async add(bookId, tableName, fields, id = newRecordId(), version = 1, extra = {}) {
    await this.checkWritable(bookId, tableName);
    const t = await this.backend.readTable(bookId, tableName);
    const row = recordToRow(t.headers, null, fields, { id, ...this.sys(version) });
    for (const [h, v] of Object.entries(extra)) {
      const i = t.headers.indexOf(h);
      if (i >= 0) row[i] = v;
    }
    await this.backend.addRow(bookId, tableName, row);
    return id;
  }

  async remove(bookId, rec) {
    await this.checkWritable(bookId, rec.table);
    const { t, cur } = await this.freshRecord(bookId, rec);
    if (!cur) return { status: 'deleted' };
    await this.backend.deleteRow(bookId, t, cur.index);
    return { status: 'removed' };
  }

  /** Move to another term sheet, keeping the record id (spec F12). */
  async move(bookId, rec, targetTable, termYear) {
    await this.checkWritable(bookId, rec.table);
    const { t, cur } = await this.freshRecord(bookId, rec);
    if (!cur) return { status: 'deleted' };
    const li = t.headers.indexOf(LEGACY_LIBINSIGHT);
    const extra = li >= 0 && cur.raw[li] !== '' && cur.raw[li] != null ? { [LEGACY_LIBINSIGHT]: cur.raw[li] } : {};
    await this.add(bookId, targetTable, { ...cur.fields, termYear }, cur.id, cur.version + 1, extra);
    const again = await this.freshRecord(bookId, rec); // re-find: the add may have been on the same sheet's neighbour
    if (again.cur) await this.backend.deleteRow(bookId, again.t, again.cur.index);
    return { status: 'moved' };
  }

  // ---------- Requests inbox (spec §4.4) ----------

  async loadRequests() {
    const book = this.books.requests;
    if (!book) return null;
    const tables = await this.backend.readTables(book.id);
    const t = tables.find((x) => /request/i.test(x.name) || /request/i.test(x.sheet)) ?? tables[0];
    if (!t) return null;
    this.tables.set(`${book.id}/${t.name}`, t);
    const items = t.rows.map((row, index) => {
      const v = {};
      t.headers.forEach((h, i) => (v[h] = row[i] ?? ''));
      return { index, values: v, id: String(v['Request ID'] || `row-${index}`), status: String(v.Status || 'New') };
    });
    return { table: t.name, headers: t.headers, items };
  }

  async updateRequest(requestId, patch) {
    const book = this.books.requests;
    const all = await this.backend.readTables(book.id);
    const t = all.find((x) => /request/i.test(x.name) || /request/i.test(x.sheet)) ?? all[0];
    const idCol = t.headers.indexOf('Request ID');
    const index = t.rows.findIndex((r, i) => (idCol >= 0 ? String(r[idCol]) === requestId : `row-${i}` === requestId));
    if (index < 0) throw new Error('That request is no longer in the inbox.');
    const row = t.rows[index].slice();
    for (const [h, v] of Object.entries(patch)) {
      const c = t.headers.indexOf(h);
      if (c >= 0) row[c] = v;
    }
    await this.backend.writeRow(book.id, t, index, row);
  }

  async createRequestsWorkbook() {
    this.books.requests = await this.backend.createRequestsWorkbook(REQUEST_HEADERS);
    return this.books.requests;
  }

  // ---------- One-time migration (spec §7, web version) ----------

  /**
   * Makes the active workbook usable by the app. Idempotent: safe to run again.
   * Never overwrites visible text. Columns are added at the end of each table; the app works by header name,
   * so their position doesn't matter (they can be dragged into place in Excel afterwards).
   */
  async prepare(lists, log = () => {}) {
    const book = this.books.active;
    const report = [];
    const sheets = await this.backend.listSheets(book.id);
    let tables = await this.backend.readTables(book.id);
    for (const sheet of sheets) {
      if (!classify(sheet) || tables.some((t) => t.sheet === sheet)) continue;
      const name = `tbl_${sheet.replace(/[^A-Za-z0-9]/g, '')}`;
      log(`Converting “${sheet}” to an Excel table (${name})…`);
      await this.backend.createTable(book.id, sheet, name);
      report.push(`${sheet}: converted to table ${name}`);
    }
    tables = await this.backend.readTables(book.id);

    for (let t of tables) {
      const ctx = this.ctxFor(t);
      if (!ctx.kind) continue;
      // Before adding columns, so Actual Student Attendance takes the stray column's place instead of going on the end.
      if (ctx.kind === 'term') t = await this.migrateLegacyColumns(book, t, log, report);
      const map = mapHeaders(t.headers);
      const wanted = ctx.kind === 'tours' ? SYSTEM_FIELDS : [...FIELDS, ...SYSTEM_FIELDS];
      const add = wanted.filter((f) => map[f.key] === undefined);
      for (const f of add) {
        log(`${t.sheet}: adding column “${f.header}”…`);
        await this.backend.addColumn(book.id, t, f.header);
        t = await this.backend.readTable(book.id, t.name);
      }
      if (add.length) report.push(`${t.sheet}: added ${add.map((f) => f.header).join(', ')}`);

      const m = mapHeaders(t.headers);
      const missingIds = t.rows.filter((r) => !r[m.id]).length;
      if (missingIds) {
        log(`${t.sheet}: assigning record ids to ${missingIds} rows…`);
        await this.assignIds(book, t);
        report.push(`${t.sheet}: ${missingIds} rows given a record id`);
      }

      if (ctx.kind === 'term' && t.rows.length) {
        // Backfill blanks only: Term Year from the sheet name, College/Student Type suggested from the course.
        let filled = 0;
        const col = (key, fn) => t.rows.map((r) => {
          if (r[m[key]] !== '' && r[m[key]] != null) return r[m[key]];
          const v = fn(r);
          if (v) filled++;
          return v;
        });
        const termYears = col('termYear', () => ctx.termYear);
        const colleges = col('college', (r) => suggestFromCourse(r[m.course], lists).college);
        const types = col('studentType', (r) => suggestFromCourse(r[m.course], lists).studentType);
        if (filled) {
          log(`${t.sheet}: filling ${filled} blank Term Year / College / Student Type cells…`);
          await this.backend.writeColumn(book.id, t, m.termYear, termYears);
          await this.backend.writeColumn(book.id, t, m.college, colleges);
          await this.backend.writeColumn(book.id, t, m.studentType, types);
          report.push(`${t.sheet}: ${filled} blank cells filled with suggestions (check them in the Request stage view)`);
        }
        const undated = t.rows.filter((r) => !toIsoDate(r[m.date])).length;
        if (undated) report.push(`${t.sheet}: ${undated} row(s) have no readable date. Flagged “needs review”.`);
      }
      for (const f of SYSTEM_FIELDS) await this.backend.hideColumns(book.id, t, m[f.key], m[f.key]);
      const legacy = t.headers.findIndex((h) => String(h).trim().toLowerCase() === LEGACY_LIBINSIGHT.toLowerCase());
      if (legacy >= 0) await this.backend.hideColumns(book.id, t, legacy, legacy);
    }

    const termSheets = (await this.backend.readTables(book.id)).filter((t) => classify(t.sheet) === 'term').map((t) => ({ kind: 'term', sheet: t.sheet }));
    log('Checking the hidden Lists sheet (term dates, closed terms)…');
    if (await this.ensureLists(book, this.defaultTermsFor({ tables: termSheets }, lists.terms))) {
      report.push('Created the hidden Lists sheet with tbl_Terms (term dates and Closed flags) and tbl_AppLock');
    }
    report.push(...(await this.applyDropdowns(book, lists, log)));

    if (!this.books.requests) {
      log('Creating Instruction Requests.xlsx…');
      await this.createRequestsWorkbook();
      report.push('Created Instruction Requests.xlsx (table tbl_Requests)');
    }
    return report.length ? report : ['Everything was already prepared. Nothing changed.'];
  }

  /**
   * Spec §7 steps 4–5, on one term table. Idempotent.
   * - "Entered into LibInsight" is renamed _LegacyLibInsight and hidden: the values stay in the workbook, out of view.
   * - The unlabeled stray column: its values are appended to Notes, then it is emptied and becomes
   *   Actual Student Attendance (only if the table doesn't already have that column).
   * Returns the table re-read after any change.
   */
  async migrateLegacyColumns(book, t, log, report) {
    const reread = () => this.backend.readTable(book.id, t.name);
    const lower = (h) => String(h ?? '').trim().toLowerCase();

    const li = t.headers.findIndex(isLibInsightHeader);
    if (li >= 0) {
      if (t.headers.some((h) => lower(h) === LEGACY_LIBINSIGHT.toLowerCase())) {
        report.push(`${t.sheet}: both “${t.headers[li]}” and ${LEGACY_LIBINSIGHT} exist, so “${t.headers[li]}” was left as it is. Check it in Excel.`);
      } else {
        log(`${t.sheet}: moving “${t.headers[li]}” out of view (${LEGACY_LIBINSIGHT}, hidden)…`);
        await this.backend.renameColumn(book.id, t, li, LEGACY_LIBINSIGHT);
        await this.backend.hideColumns(book.id, t, li, li);
        const kept = t.rows.filter((r) => r[li] !== '' && r[li] != null).length;
        report.push(`${t.sheet}: “${t.headers[li]}” cut from view. Its ${kept} value(s) are kept in the hidden ${LEGACY_LIBINSIGHT} column.`);
        t = await reread();
      }
    }

    const strays = t.headers.map((h, i) => (isStrayHeader(h) ? i : -1)).filter((i) => i >= 0);
    if (!strays.length) return t;
    const m = mapHeaders(t.headers);
    const si = strays[0];
    const withValues = t.rows.filter((r) => String(r[si] ?? '').trim() !== '').length;
    if (withValues) {
      if (m.notes === undefined) {
        report.push(`${t.sheet}: the unlabeled column has ${withValues} value(s), but there's no Notes column to move them to. Left as it is.`);
        return t;
      }
      log(`${t.sheet}: moving ${withValues} value(s) from the unlabeled column into Notes…`);
      const notes = t.rows.map((r) => appendToNotes(r[m.notes], r[si]));
      await this.backend.writeColumn(book.id, t, m.notes, notes);
      await this.backend.writeColumn(book.id, t, si, t.rows.map(() => ''));
      report.push(`${t.sheet}: ${withValues} value(s) from the unlabeled column appended to Notes`);
    }
    if (m.actual === undefined) {
      log(`${t.sheet}: turning the unlabeled column into “Actual Student Attendance”…`);
      await this.backend.renameColumn(book.id, t, si, 'Actual Student Attendance');
      report.push(`${t.sheet}: the unlabeled column is now Actual Student Attendance`);
    }
    if (strays.length > 1) report.push(`${t.sheet}: ${strays.length - 1} more unlabeled column(s) left as they are. Check them in Excel.`);
    return reread();
  }

  /**
   * Writes the dropdown lists to the hidden Lists sheet (columns M onwards) and points Excel data validation on
   * every term and tours table at them (spec §3.3), so Excel's dropdowns match the app's.
   * Re-runnable: each run replaces the lists and the validation. Returns report lines.
   */
  async applyDropdowns(book, lists, log = () => {}, termYears = null) {
    const raw = await this.backend.readTables(book.id);
    if (!termYears) {
      const meta = this.parseMeta(raw);
      termYears = meta.terms.map((x) => x.termYear).filter(Boolean);
      if (!termYears.length) termYears = raw.filter((t) => classify(t.sheet) === 'term').map((t) => termYearFromSheet(t.sheet));
    }
    const defs = dropdownLists(lists, termYears);
    if (!defs.length) return [];
    const sheets = await this.backend.listSheets(book.id);
    if (!sheets.includes('Lists')) {
      await this.backend.addWorksheet(book.id, 'Lists');
      await this.backend.hideSheet(book.id, 'Lists');
    }
    log('Writing the dropdown lists to the Lists sheet…');
    const c0 = DROPDOWN_COL;
    const rows = Math.max(DROPDOWN_ROWS, ...defs.map((d) => d.values.length));
    const grid = [defs.map((d) => d.name)];
    for (let r = 0; r < rows; r++) grid.push(defs.map((d) => d.values[r] ?? '')); // blanks clear values from a longer earlier list
    await this.backend.writeRange(book.id, 'Lists', `${colLetters(c0)}1:${colLetters(c0 + defs.length - 1)}${rows + 1}`, grid);

    let applied = 0;
    for (const t of raw) {
      if (!classify(t.sheet)) continue;
      const map = mapHeaders(t.headers);
      for (const [i, d] of defs.entries()) {
        if (map[d.key] === undefined) continue;
        const L = colLetters(c0 + i);
        await this.backend.setListValidation(book.id, t, map[d.key], `=Lists!$${L}$2:$${L}$${d.values.length + 1}`, { strict: d.strict, name: d.name });
        applied++;
      }
    }
    log(`Updated ${applied} Excel dropdowns.`);
    return [`Excel dropdowns updated from the app's lists (${applied} columns)`];
  }
}
