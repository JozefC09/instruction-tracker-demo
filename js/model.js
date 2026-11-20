// Pure data logic: column mapping, parsing, suggestions, filtering and dashboard aggregation.
// No DOM and no network here, so it can be tested on its own (see tests.html).

/** Visible columns of a term sheet, in the order spec §3.2 puts them. The app always works by header name. */
export const FIELDS = [
  { key: 'termYear', header: 'Term Year', stage: 'request' },
  { key: 'date', header: 'Date', stage: 'request', type: 'date', required: true },
  { key: 'time', header: 'Time', stage: 'request', type: 'time' },
  { key: 'course', header: 'Course #', stage: 'request', required: true },
  { key: 'faculty', header: 'Faculty Name', stage: 'request', required: true },
  { key: 'college', header: 'College/School', stage: 'request', list: 'colleges' },
  { key: 'studentType', header: 'Student Type', stage: 'request', list: 'studentTypes' },
  { key: 'librarian', header: 'Librarian Assigned', stage: 'librarian', list: 'librarianNames' },
  { key: 'sessionType', header: 'Session Type', stage: 'librarian', list: 'sessionTypes' },
  { key: 'format', header: 'Session Format', stage: 'request', list: 'sessionFormats' },
  { key: 'location', header: 'Location', stage: 'request', list: 'locations', combo: true },
  { key: 'room', header: 'Room Reservation Status', stage: 'librarian', list: 'roomStatuses' },
  { key: 'expected', header: '# Students', stage: 'request', type: 'int' },
  { key: 'assessment', header: 'Assessment', stage: 'planning', list: 'yesNo' },
  { key: 'notes', header: 'Notes', stage: 'all', type: 'multiline' },
  { key: 'prep', header: 'Prep Time', stage: 'planning' },
  { key: 'actual', header: 'Actual Student Attendance', stage: 'after', type: 'int' },
];

/** Hidden columns the app manages (spec §3.2). */
export const SYSTEM_FIELDS = [
  { key: 'id', header: '_RecordId' },
  { key: 'version', header: '_RowVersion' },
  { key: 'modifiedAt', header: '_ModifiedAt' },
  { key: 'modifiedBy', header: '_ModifiedBy' },
];

/** Other header spellings seen in the workbook (tours sheet, legacy headers). */
const ALIASES = {
  'trexler librarian': 'librarian',
  '# of students': 'expected',
  'students': 'expected',
  'mentors': 'faculty',
  'actual attendance': 'actual',
  'term': 'termYear',
};

export const STAGES = {
  all: { label: 'All', editable: null },
  request: { label: 'Request', editable: ['termYear', 'date', 'time', 'course', 'faculty', 'college', 'studentType', 'format', 'location', 'expected', 'notes'] },
  librarian: { label: 'Librarian', editable: ['librarian', 'sessionType', 'room', 'notes'] },
  planning: { label: 'Planning', editable: ['assessment', 'prep', 'notes'] },
  after: { label: 'After Event', editable: ['actual', 'notes'] },
};

const norm = (s) => String(s ?? '').trim().toLowerCase().replace(/\s+/g, ' ');

/** Maps a header row to { fieldKey: columnIndex }. Unknown headers are kept by name so they round-trip. */
export function mapHeaders(headers) {
  const byHeader = new Map([...FIELDS, ...SYSTEM_FIELDS].map((f) => [norm(f.header), f.key]));
  const map = {};
  headers.forEach((h, i) => {
    const n = norm(h);
    const key = byHeader.get(n) ?? ALIASES[n];
    if (key && map[key] === undefined) map[key] = i;
  });
  return map;
}

// ---------- Excel value conversion ----------

const EXCEL_EPOCH = Date.UTC(1899, 11, 30);

/** Excel serial (or date-like text) → 'YYYY-MM-DD', or '' if it isn't a date. */
export function toIsoDate(v) {
  if (v === null || v === undefined || v === '') return '';
  if (typeof v === 'number') {
    if (v < 1) return ''; // 0 or a bare time is not a date
    const d = new Date(EXCEL_EPOCH + Math.round(v * 86400000));
    return d.toISOString().slice(0, 10);
  }
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (m) {
    const y = m[3].length === 2 ? `20${m[3]}` : m[3];
    return `${y}-${m[1].padStart(2, '0')}-${m[2].padStart(2, '0')}`;
  }
  return '';
}

/** 'YYYY-MM-DD' → 'M/D/YYYY', which Excel (US locale) stores as a real date when written through Graph. */
export function isoToExcelText(iso) {
  if (!iso) return '';
  const [y, m, d] = iso.split('-').map(Number);
  return `${m}/${d}/${y}`;
}

export function toInt(v) {
  if (v === null || v === undefined || v === '') return null;
  const n = typeof v === 'number' ? v : parseFloat(String(v).replace(/[^\d.]/g, ''));
  return Number.isFinite(n) ? Math.round(n) : null;
}

// ---------- Time parsing (spec §12: every legacy shape either parses or is flagged) ----------

function hourGuess(h) {
  // No AM/PM given: class hours run roughly 7am–9pm.
  if (h >= 7 && h <= 11) return h;
  if (h === 12) return 12;
  if (h >= 1 && h <= 6) return h + 12;
  return h; // already 24h (13–23) or 0
}

function readToken(tok) {
  // '1045' → 10:45, '9:30' → 9:30, '1' → 1:00, ':15' → null
  const t = tok.replace(/[^\d:]/g, '');
  if (!t || t.startsWith(':')) return null;
  let h, m;
  if (t.includes(':')) {
    [h, m] = t.split(':').map((x) => parseInt(x || '0', 10));
  } else if (t.length >= 3) {
    h = parseInt(t.slice(0, t.length - 2), 10);
    m = parseInt(t.slice(-2), 10);
  } else {
    h = parseInt(t, 10);
    m = 0;
  }
  if (!(h >= 0 && h <= 23 && m >= 0 && m <= 59)) return null;
  return { h, m };
}

/**
 * Parses a Time cell into minutes after midnight.
 * Returns { start, end|null, guessed } or null when it can't be read (shown as "needs review").
 */
export function parseTime(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') {
    if (v <= 0 || v >= 1) return null;
    let mins = Math.round(v * 1440);
    const guessed = mins < 7 * 60; // fractions without AM/PM, e.g. 0.052 = 1:15 (spec §3.4)
    if (guessed) mins += 12 * 60;
    return { start: mins, end: null, guessed };
  }
  let s = String(v).toLowerCase().replace(/[–—]/g, '-').replace(/\bto\b/g, '-').replace(/noon/g, '12pm').trim();
  if (!/\d/.test(s)) return null;
  // '12:30- 1-:15' — a stray '-' before the minutes; '1-:15' → '1:15'
  s = s.replace(/(\d)\s*-\s*:/g, '$1:');
  const suffixOf = (x) => (/(p\.?m?\.?)\s*$/.test(x) ? 'pm' : /(a\.?m?\.?)\s*$/.test(x) ? 'am' : null);
  const parts = s.split('-').map((p) => p.trim()).filter(Boolean);
  if (parts.length === 0 || parts.length > 2) return null;
  const toks = parts.map((p) => ({ t: readToken(p), suf: suffixOf(p) }));
  if (toks.some((x) => !x.t)) return null;

  const apply = (t, suf) => {
    let h = t.h;
    if (suf === 'pm' && h < 12) h += 12;
    if (suf === 'am' && h === 12) h = 0;
    return h * 60 + t.m;
  };

  const [a, b] = toks;
  let guessed = false;
  if (!b) {
    const start = a.suf ? apply(a.t, a.suf) : (guessed = !a.suf && a.t.h <= 12, hourGuess(a.t.h) * 60 + a.t.m);
    return { start, end: null, guessed };
  }
  let end, start;
  if (b.suf) {
    end = apply(b.t, b.suf);
    if (a.suf) start = apply(a.t, a.suf);
    else {
      // '1-2pm' → 13:00–14:00; '11-12 pm' → 11:00–12:00; '10-1pm' → 10:00–13:00
      start = apply(a.t, b.suf);
      if (start > end) start = apply(a.t, b.suf === 'pm' ? 'am' : 'pm');
    }
  } else {
    guessed = !a.suf;
    start = a.suf ? apply(a.t, a.suf) : hourGuess(a.t.h) * 60 + a.t.m;
    end = hourGuess(b.t.h) * 60 + b.t.m;
    if (end <= start && end + 720 <= 24 * 60) end += 720;
  }
  if (end <= start) return null;
  return { start, end, guessed };
}

export function fmtMinutes(mins, withSuffix = true) {
  const h24 = Math.floor(mins / 60), m = mins % 60;
  const h = ((h24 + 11) % 12) + 1;
  return `${h}:${String(m).padStart(2, '0')}${withSuffix ? (h24 < 12 ? ' AM' : ' PM') : ''}`;
}

/** Start/end minutes → '1:00–2:00 PM' (the display text the app writes into Time). */
export function formatTimeRange(start, end) {
  if (start == null) return '';
  if (end == null) return fmtMinutes(start);
  const samePeriod = start < 720 === end < 720;
  return `${fmtMinutes(start, !samePeriod)}–${fmtMinutes(end)}`;
}

// ---------- Normalization and suggestions ----------

/** 'en103' → 'EN 103', 'pa-517' → 'PA 517'. Anything else is left alone. */
export function normalizeCourse(s) {
  const m = String(s ?? '').trim().match(/^([A-Za-z]{2,4})\s*-?\s*(\d{3,4}[A-Za-z]?)(.*)$/);
  return m ? `${m[1].toUpperCase()} ${m[2].toUpperCase()}${m[3]}` : String(s ?? '').trim();
}

/** 'GMB 210' → 'Gmb 210', matching the list's casing when the room is already known. */
export function normalizeLocation(s, known = []) {
  const v = String(s ?? '').trim().replace(/\s+/g, ' ');
  if (!v) return '';
  const hit = known.find((k) => k.toLowerCase() === v.toLowerCase());
  if (hit) return hit;
  return v.replace(/^([A-Za-z]{3,})(\s+\d)/, (_, w, rest) => (w.length <= 4 && w === w.toUpperCase() && !['TTC'].includes(w) ? w[0] + w.slice(1).toLowerCase() : w) + rest);
}

/** Suggests College/School and Student Type from the course number (spec F8). */
export function suggestFromCourse(course, lists) {
  const m = normalizeCourse(course).match(/^([A-Z]{2,4}) (\d{3,4})/);
  if (!m) return { college: '', studentType: '' };
  const college = lists.prefixToCollege?.[m[1]] ?? '';
  const studentType = parseInt(m[2], 10) >= (lists.graduateThreshold ?? 500) ? 'Graduate' : 'Undergraduate';
  return { college, studentType };
}

/** 'FALL 26' → 'Fall 2026'. */
export function termYearFromSheet(sheet) {
  const m = String(sheet).trim().match(/^(fall|winter|spring|summer)\s*'?(\d{2}|\d{4})$/i);
  if (!m) return '';
  const y = m[2].length === 2 ? `20${m[2]}` : m[2];
  return `${m[1][0].toUpperCase()}${m[1].slice(1).toLowerCase()} ${y}`;
}

/** 'FALL 26' → 'FALL 27' (same season, next year). Non-term names are returned unchanged. */
export function nextSheetName(sheet) {
  const m = String(sheet).trim().match(/^(fall|winter|spring|summer)(\s*'?)(\d{2}|\d{4})$/i);
  if (!m) return sheet;
  const n = Number(m[3]) + 1;
  return `${m[1]}${m[2]}${m[3].length === 2 ? String(n % 100).padStart(2, '0') : n}`;
}

/** 'tbl_FALL26' style name for a sheet. */
export function tableNameFor(sheet) {
  return `tbl_${String(sheet).replace(/[^A-Za-z0-9]/g, '')}`;
}

/** '2026-2027' → '2027-2028' */
export function nextYearLabel(year) {
  const [a, b] = String(year).split('-').map(Number);
  return `${a + 1}-${b + 1}`;
}

/** Same month/day one year later (Feb 29 → Feb 28). */
export function shiftYear(iso, years = 1) {
  if (!iso) return '';
  const [y, m, d] = iso.split('-').map(Number);
  const last = new Date(Date.UTC(y + years, m, 0)).getUTCDate();
  return `${y + years}-${String(m).padStart(2, '0')}-${String(Math.min(d, last)).padStart(2, '0')}`;
}

export function isTermSheet(sheet) {
  return !!termYearFromSheet(sheet);
}

export function isToursSheet(sheet) {
  return /tour/i.test(sheet);
}

// ---------- Legacy columns (spec §3.2, §7 steps 4–5) ----------

/** Hidden column that keeps the values of the cut "Entered into LibInsight" column. */
export const LEGACY_LIBINSIGHT = '_LegacyLibInsight';

export function isLibInsightHeader(h) {
  return /libinsight/i.test(String(h ?? '')) && String(h).trim().toLowerCase() !== LEGACY_LIBINSIGHT.toLowerCase();
}

/** The unlabeled stray column: blank in the sheet, or "Column15"-style once Excel has made the sheet a table. */
export function isStrayHeader(h) {
  const s = String(h ?? '').trim();
  return s === '' || /^column\s*\d+$/i.test(s);
}

/** Notes with a stray value appended (spec Q8). Unchanged if it's already there, so re-running is safe. */
export function appendToNotes(notes, extra) {
  const n = String(notes ?? '').trim();
  const x = String(extra ?? '').trim();
  if (!x || n.includes(x)) return n;
  return n ? `${n}\n${x}` : x;
}

// ---------- Excel dropdowns (spec §3.3: data validation rewritten from Lists) ----------

/**
 * The dropdown lists written to the Lists sheet and applied as Excel data validation.
 * strict = Excel warns on a value that isn't in the list; Location only suggests (pick or type).
 */
export function dropdownLists(lists, termYears = []) {
  const uniq = (a) => [...new Set(a.map((x) => String(x ?? '').trim()).filter(Boolean))];
  return [
    { key: 'termYear', name: 'Term Year', values: uniq(termYears), strict: true },
    { key: 'college', name: 'College/School', values: uniq(lists.colleges || []), strict: true },
    { key: 'studentType', name: 'Student Type', values: uniq(lists.studentTypes || []), strict: true },
    { key: 'librarian', name: 'Librarian', values: uniq((lists.librarians || []).map((l) => l.name)), strict: true },
    { key: 'sessionType', name: 'Session Type', values: uniq(lists.sessionTypes || []), strict: true },
    { key: 'format', name: 'Session Format', values: uniq(lists.sessionFormats || []), strict: true },
    { key: 'location', name: 'Location', values: uniq(lists.locations || []), strict: false },
    { key: 'room', name: 'Room Reservation Status', values: uniq(lists.roomStatuses || []), strict: true },
    { key: 'assessment', name: 'Assessment', values: uniq(lists.yesNo || []), strict: true },
  ].filter((d) => d.values.length);
}

/** 'Updated - 1 hour' / 'New – 8 hours' / '2' → { kind, hours }. */
export function parsePrep(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') return { kind: '', hours: v };
  const s = String(v).toLowerCase();
  const n = s.match(/(\d+(?:\.\d+)?)\s*(h|hr|hrs|hour|hours|m|min|mins|minutes)?/);
  if (!n) return null;
  let hours = parseFloat(n[1]);
  if (n[2] && n[2].startsWith('m')) hours = hours / 60;
  const kind = /new/.test(s) ? 'New' : /updat/.test(s) ? 'Updated' : '';
  return { kind, hours };
}

export function formatPrep(kind, hours) {
  if (hours == null || hours === '') return '';
  const h = Number(hours);
  const label = `${h} ${h === 1 ? 'hour' : 'hours'}`;
  return kind ? `${kind} – ${label}` : label;
}

// ---------- Records ----------

/**
 * Turns one table's header + rows into records.
 * Each record keeps the raw row so columns the app doesn't know about round-trip untouched.
 */
export function rowsToRecords(headers, rows, ctx) {
  const map = mapHeaders(headers);
  return rows.map((row, index) => {
    const get = (k) => (map[k] === undefined ? '' : row[map[k]]);
    // An empty row (e.g. the one Excel keeps in a cleared table) isn't a session.
    const blank = FIELDS.every((fd) => map[fd.key] === undefined || get(fd.key) === '' || get(fd.key) == null);
    const f = {};
    for (const fd of FIELDS) {
      const v = get(fd.key);
      if (fd.type === 'date') f[fd.key] = toIsoDate(v);
      else if (fd.type === 'int') f[fd.key] = toInt(v);
      else if (fd.type === 'time') f[fd.key] = v === null || v === undefined ? '' : v;
      else f[fd.key] = v === null || v === undefined ? '' : String(v).trim();
    }
    if (!f.termYear && ctx.termYear) f.termYear = ctx.termYear;
    if (ctx.kind === 'tours') {
      f.sessionType = f.sessionType || 'Tour';
      f.studentType = f.studentType || 'Undergraduate';
      f.college = f.college || 'All / N/A';
    }
    const t = parseTime(f.time);
    const rec = {
      table: ctx.table,
      sheet: ctx.sheet,
      kind: ctx.kind,
      termYear: ctx.termYear,
      index,
      blank,
      id: String(get('id') || ''),
      version: toInt(get('version')) ?? 0,
      modifiedAt: String(get('modifiedAt') || ''),
      modifiedBy: String(get('modifiedBy') || ''),
      fields: f,
      start: t?.start ?? null,
      end: t?.end ?? null,
      timeDisplay: typeof f.time === 'number' ? (t ? formatTimeRange(t.start, t.end) : '') : String(f.time),
      // "Online"/"TBA" are fine; text with digits that can't be read, Excel fractions without AM/PM, or no date → review.
      needsReview: (/\d/.test(String(f.time)) && !t) || (typeof f.time === 'number' && (t?.guessed ?? false)) || !f.date,
      raw: row.slice(),
    };
    rec.hash = visibleHash(rec.fields);
    return rec;
  });
}

/** Hash of the visible values, used to detect edits made directly in Excel (spec §5 step 2.4). */
export function visibleHash(fields) {
  const s = FIELDS.map((fd) => String(fields[fd.key] ?? '')).join('\u241F');
  let h = 2166136261;
  for (let i = 0; i < s.length; i++) {
    h ^= s.charCodeAt(i);
    h = Math.imul(h, 16777619);
  }
  return (h >>> 0).toString(16);
}

/** Builds a full row for writing: starts from the raw row, overlays known fields. */
export function recordToRow(headers, rec, fields, sys) {
  const map = mapHeaders(headers);
  const row = rec?.raw ? rec.raw.slice() : headers.map(() => '');
  while (row.length < headers.length) row.push('');
  for (const fd of FIELDS) {
    if (map[fd.key] === undefined || !(fd.key in fields)) continue;
    let v = fields[fd.key];
    if (fd.type === 'date') v = isoToExcelText(v);
    if (fd.type === 'int') v = v === null || v === undefined || v === '' ? '' : Number(v);
    row[map[fd.key]] = v ?? '';
  }
  for (const [k, v] of Object.entries(sys)) if (map[k] !== undefined) row[map[k]] = v;
  return row;
}

/**
 * Field-level three-way merge (spec §5 step 2.3).
 * base = values when the user opened the editor, mine = their edits, theirs = what's in the workbook now.
 */
export function mergeFields(base, mine, theirs) {
  const merged = { ...theirs };
  const conflicts = [];
  for (const fd of FIELDS) {
    const k = fd.key;
    const b = String(base[k] ?? ''), m = String(mine[k] ?? ''), t = String(theirs[k] ?? '');
    const iChanged = m !== b, theyChanged = t !== b;
    if (iChanged && !theyChanged) merged[k] = mine[k];
    else if (iChanged && theyChanged && m !== t) conflicts.push({ key: k, header: fd.header, base: base[k], mine: mine[k], theirs: theirs[k] });
  }
  return { merged, conflicts };
}

export function newRecordId() {
  return crypto.randomUUID ? crypto.randomUUID() : 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c) => {
    const r = (Math.random() * 16) | 0;
    return (c === 'x' ? r : (r & 0x3) | 0x8).toString(16);
  });
}

// ---------- Validation ----------

export function validate(fields) {
  const errors = {};
  for (const fd of FIELDS) {
    const v = fields[fd.key];
    if (fd.required && (v === '' || v == null)) errors[fd.key] = 'Required';
    if (fd.type === 'int' && v !== '' && v != null && (!Number.isInteger(Number(v)) || Number(v) < 0)) errors[fd.key] = 'Whole number, 0 or more';
  }
  return errors;
}

// ---------- Dashboard: filtering and aggregation (spec §4.6) ----------

export const DASH_SESSION_TYPE = (t) => (/^(workshop|presentation)$/i.test(t) ? 'Workshop/Presentation' : t || 'Unassigned');

export const FILTER_DIMS = [
  { key: 'studentType', label: 'Student type', value: (r) => r.fields.studentType || 'Unassigned' },
  { key: 'librarian', label: 'Librarian', value: (r) => r.fields.librarian || 'Unassigned' },
  { key: 'college', label: 'College/School', value: (r) => r.fields.college || 'Unassigned' },
  { key: 'format', label: 'Session format', value: (r) => r.fields.format || 'Unassigned' },
  { key: 'sessionType', label: 'Session type', value: (r) => DASH_SESSION_TYPE(r.fields.sessionType) },
];

/** filters = { dimKey: Set of values } (empty/missing = no filter). period = { from, to } ISO dates or null. */
export function applyFilters(records, filters, period) {
  return records.filter((r) => {
    if (period?.from && (!r.fields.date || r.fields.date < period.from)) return false;
    if (period?.to && (!r.fields.date || r.fields.date > period.to)) return false;
    for (const dim of FILTER_DIMS) {
      const sel = filters[dim.key];
      if (sel && sel.size && !sel.has(dim.value(r))) return false;
    }
    return true;
  });
}

export function studentsOf(r) {
  if (r.fields.actual != null && r.fields.actual !== '') return { n: Number(r.fields.actual), estimate: false };
  if (r.fields.expected != null && r.fields.expected !== '') return { n: Number(r.fields.expected), estimate: true };
  return { n: 0, estimate: false };
}

export function groupBy(records, valueFn, measure = 'sessions') {
  const m = new Map();
  for (const r of records) {
    const k = valueFn(r);
    m.set(k, (m.get(k) ?? 0) + (measure === 'students' ? studentsOf(r).n : 1));
  }
  return [...m.entries()].map(([label, value]) => ({ label, value })).sort((a, b) => b.value - a.value || a.label.localeCompare(b.label));
}

export function needsAttention(r, todayIso) {
  const reasons = [];
  if (r.kind !== 'tours') {
    if (/pending/i.test(r.fields.room)) reasons.push('Room pending');
    if (!r.fields.librarian) reasons.push('No librarian');
  }
  if (r.fields.date && r.fields.date < todayIso && (r.fields.actual == null || r.fields.actual === '')) reasons.push('Attendance missing');
  return reasons;
}

export function summarize(records, todayIso) {
  const in14 = addDays(todayIso, 14);
  let students = 0, estimates = 0, upcoming = 0, attention = 0;
  for (const r of records) {
    const s = studentsOf(r);
    students += s.n;
    if (s.estimate) estimates++;
    if (r.fields.date >= todayIso && r.fields.date <= in14) upcoming++;
    if (needsAttention(r, todayIso).length) attention++;
  }
  return { sessions: records.length, students, estimates, upcoming, attention };
}

export function addDays(iso, n) {
  const d = new Date(`${iso}T12:00:00Z`);
  d.setUTCDate(d.getUTCDate() + n);
  return d.toISOString().slice(0, 10);
}

export function todayIso() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

/** Term the date falls in, from config terms [{ name, from, to }]. */
export function termForDate(iso, terms) {
  return terms.find((t) => t.from && t.to && iso >= t.from && iso <= t.to)?.name ?? '';
}

// ---------- .ics (spec §4.7) ----------

function icsEscape(s) {
  return String(s ?? '').replace(/\\/g, '\\\\').replace(/;/g, '\\;').replace(/,/g, '\\,').replace(/\r?\n/g, '\\n');
}

function icsLocal(iso, mins) {
  const h = Math.floor(mins / 60), m = mins % 60;
  return `${iso.replace(/-/g, '')}T${String(h).padStart(2, '0')}${String(m).padStart(2, '0')}00`;
}

/** One VEVENT per record. The UID is stable (record id) so re-exporting updates the same Outlook event. */
export function buildIcs(records, { domain = 'instruction-tracker', tz = 'America/New_York' } = {}) {
  const stamp = new Date().toISOString().replace(/[-:]/g, '').replace(/\.\d+/, '');
  const lines = [
    'BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Trexler Library//Instruction Tracker//EN', 'CALSCALE:GREGORIAN', 'METHOD:PUBLISH',
    'BEGIN:VTIMEZONE', `TZID:${tz}`,
    'BEGIN:DAYLIGHT', 'TZOFFSETFROM:-0500', 'TZOFFSETTO:-0400', 'TZNAME:EDT', 'DTSTART:19700308T020000', 'RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=2SU', 'END:DAYLIGHT',
    'BEGIN:STANDARD', 'TZOFFSETFROM:-0400', 'TZOFFSETTO:-0500', 'TZNAME:EST', 'DTSTART:19701101T020000', 'RRULE:FREQ=YEARLY;BYMONTH=11;BYDAY=1SU', 'END:STANDARD',
    'END:VTIMEZONE',
  ];
  for (const r of records) {
    if (!r.fields.date) continue;
    const f = r.fields;
    lines.push('BEGIN:VEVENT', `UID:${r.id || r.table + '-' + r.index}@${domain}`, `DTSTAMP:${stamp}`, `SEQUENCE:${r.version || 0}`);
    if (r.start != null) {
      lines.push(`DTSTART;TZID=${tz}:${icsLocal(f.date, r.start)}`, `DTEND;TZID=${tz}:${icsLocal(f.date, r.end ?? r.start + 60)}`);
    } else {
      lines.push(`DTSTART;VALUE=DATE:${f.date.replace(/-/g, '')}`);
    }
    lines.push(
      `SUMMARY:${icsEscape(`${f.sessionType || 'Instruction'}: ${f.course}${f.faculty ? ' (' + f.faculty + ')' : ''}`)}`,
      `LOCATION:${icsEscape(f.location)}`,
      `DESCRIPTION:${icsEscape([`Format: ${f.format}`, `Room: ${f.room}`, `Expected students: ${f.expected ?? ''}`, f.notes].filter(Boolean).join('\n'))}`,
      'END:VEVENT',
    );
  }
  lines.push('END:VCALENDAR');
  return lines.join('\r\n');
}
