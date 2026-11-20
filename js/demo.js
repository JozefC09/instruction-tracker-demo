// In-memory stand-in for the SharePoint workbooks, used when config.js has no clientId (or with ?demo=1).
// It implements the same methods as GraphBackend, so the demo exercises the real save/merge/conflict code.
import { FIELDS, SYSTEM_FIELDS, newRecordId, isoToExcelText } from './model.js';
import { REQUEST_HEADERS, TERMS_HEADERS, LOCK_HEADERS } from './store.js';
import { parseAddress } from './graph.js';

const listsTables = (terms) => ({
  tbl_Terms: { sheet: 'Lists', headers: TERMS_HEADERS, rows: terms.map(([sheet, termYear, from, to, closed]) => [sheet, termYear, isoToExcelText(from), isoToExcelText(to), closed ? 'Yes' : 'No', closed ? 'kathleen@desales.edu' : '', closed ? `${to}T17:00:00Z` : '']) },
  tbl_AppLock: { sheet: 'Lists', headers: LOCK_HEADERS, rows: [['', '', '']] },
});

const clone = (x) => JSON.parse(JSON.stringify(x));

function rng(seed) {
  let s = seed;
  return () => ((s = (s * 1664525 + 1013904223) % 4294967296) / 4294967296);
}

const COURSES = [
  ['EN 103', 'Dr. Alvarez'], ['EN 103', 'Prof. Brennan'], ['NU 800', 'Dr. Okafor'], ['PA 517', 'Dr. Feld'],
  ['HI 210', 'Dr. Marsh'], ['BU 330', 'Prof. Chen'], ['MBA 610', 'Dr. Ruiz'], ['DPT 702', 'Dr. Patel'],
  ['CJ 250', 'Prof. Gallo'], ['TH 109', 'Fr. Dolan'], ['BI 151', 'Dr. Weiss'], ['ED 520', 'Dr. Novak'],
  ['CSD 540', 'Dr. Hart'], ['PS 101', 'Prof. Kim'], ['AC 210', 'Prof. Shaw'], ['CO 380', 'Dr. Lane'],
];
const LIBS = ['Kathleen', 'Lisa', 'Michele'];
const TYPES = ['One shot', 'One shot', 'One shot', 'Embedded', 'Workshop', 'Presentation', 'Orientation'];
const FORMATS = ['Face to face', 'Face to face', 'Synchronous', 'Asynchronous'];
const ROOMS = ['Gmb 155', 'Gmb 210', 'Dlg 105', 'TTC', 'PC Lab 2', 'Library', 'Wills 223', 'TEAMS', 'Online'];
const TIMES = [['9:30 -1045', 0], ['1-2pm', 0], ['10:00–10:50 AM', 0], ['11-12 pm', 0], ['12:30- 1-:15', 0], ['2:00–3:15 PM', 0], ['6:00–8:30 PM', 0], ['Online', 0]];
const PREFIX = { EN: 'College of Arts & Sciences', HI: 'College of Arts & Sciences', TH: 'College of Arts & Sciences', BI: 'College of Arts & Sciences', CJ: 'College of Arts & Sciences', ED: 'College of Arts & Sciences', PS: 'College of Arts & Sciences', CO: 'College of Arts & Sciences', BU: 'College of Business', MBA: 'College of Business', AC: 'College of Business', PA: 'College of Health Professions', DPT: 'College of Health Professions', CSD: 'College of Health Professions', NU: 'School of Nursing' };

const HEADERS = [...FIELDS.map((f) => f.header), ...SYSTEM_FIELDS.map((f) => f.header)];

function makeTerm(rand, termYear, from, to, n, today) {
  const rows = [];
  const span = (new Date(to) - new Date(from)) / 86400000;
  for (let i = 0; i < n; i++) {
    const d = new Date(new Date(from).getTime() + Math.floor(rand() * span) * 86400000);
    if (d.getUTCDay() === 0) d.setUTCDate(d.getUTCDate() + 1);
    if (d.getUTCDay() === 6) d.setUTCDate(d.getUTCDate() + 2);
    const iso = d.toISOString().slice(0, 10);
    const [course, faculty] = COURSES[Math.floor(rand() * COURSES.length)];
    const prefix = course.split(' ')[0];
    const num = parseInt(course.split(' ')[1], 10);
    const past = iso < today;
    const expected = 8 + Math.floor(rand() * 25);
    const adult = prefix === 'BU' && rand() < 0.5;
    const lib = rand() < 0.08 && !past ? '' : LIBS[Math.floor(rand() * (rand() < 0.2 ? 3 : 2.6))];
    const fmt = FORMATS[Math.floor(rand() * FORMATS.length)];
    const time = TIMES[Math.floor(rand() * TIMES.length)][0];
    const prep = rand() < 0.5 ? `${rand() < 0.4 ? 'New' : 'Updated'} – ${1 + Math.floor(rand() * 4)} hours` : '';
    const v = {
      'Term Year': termYear,
      Date: isoToExcelText(iso),
      Time: fmt === 'Asynchronous' ? 'Online' : time,
      'Course #': course,
      'Faculty Name': faculty,
      'College/School': rand() < 0.07 ? '' : PREFIX[prefix],
      'Student Type': adult ? 'Adult Studies' : num >= 500 ? 'Graduate' : 'Undergraduate',
      'Librarian Assigned': lib,
      'Session Type': TYPES[Math.floor(rand() * TYPES.length)],
      'Session Format': fmt,
      Location: fmt === 'Face to face' ? ROOMS[Math.floor(rand() * 7)] : fmt === 'Synchronous' ? 'TEAMS' : 'Online',
      'Room Reservation Status': fmt === 'Face to face' ? (rand() < (past ? 0.02 : 0.3) ? 'Pending' : 'Confirmed') : '',
      '# Students': expected,
      Assessment: rand() < 0.3 ? 'Yes' : 'No',
      Notes: rand() < 0.2 ? 'Focus on database searching for the annotated bibliography.' : '',
      'Prep Time': prep,
      'Actual Student Attendance': past && rand() < 0.85 ? Math.max(0, expected - Math.floor(rand() * 5)) : '',
      _RecordId: newRecordId(),
      _RowVersion: 1,
      _ModifiedAt: new Date(Date.now() - rand() * 9e9).toISOString(),
      _ModifiedBy: `${lib || 'kathleen'}@desales.edu`.toLowerCase(),
    };
    rows.push(HEADERS.map((h) => v[h] ?? ''));
  }
  // Workbook order is by date, like the real sheets.
  return rows.sort((a, b) => new Date(a[1]) - new Date(b[1]));
}

function makeTours(rand, from, n) {
  const headers = ['Date', 'Time', 'Mentors', 'Trexler Librarian', '# of students', '_RecordId', '_RowVersion', '_ModifiedAt', '_ModifiedBy'];
  const rows = [];
  for (let i = 0; i < n; i++) {
    const d = new Date(new Date(from).getTime() + Math.floor(rand() * 12) * 86400000).toISOString().slice(0, 10);
    rows.push([isoToExcelText(d), 0.052 + Math.floor(rand() * 3) * 0.042, ['A. Rossi', 'J. Park', 'M. Diaz'][i % 3], LIBS[i % 3], 15 + Math.floor(rand() * 20), newRecordId(), 1, '', '']);
  }
  return { headers, rows };
}

export function buildDemoWorkbooks(today) {
  const rand = rng(42);
  const tours = makeTours(rand, '2026-08-20', 6);
  const active = {
    name: '2026-2027 Instruction Classes.xlsx',
    tables: {
      tbl_oriettours: { sheet: 'oriet tours', ...tours },
      tbl_FALL26: { sheet: 'FALL 26', headers: HEADERS, rows: makeTerm(rand, 'Fall 2026', '2026-08-24', '2026-12-18', 64, today) },
      tbl_WINTER26: { sheet: 'WINTER 26', headers: HEADERS, rows: makeTerm(rand, 'Winter 2026', '2027-01-04', '2027-01-15', 4, today) },
      tbl_SPRING27: { sheet: 'SPRING 27', headers: HEADERS, rows: makeTerm(rand, 'Spring 2027', '2027-01-19', '2027-05-07', 18, today) },
      tbl_SUMMER27: { sheet: 'SUMMER 27', headers: HEADERS, rows: makeTerm(rand, 'Summer 2027', '2027-05-17', '2027-08-13', 3, today) },
      ...listsTables([
        ['FALL 26', 'Fall 2026', '2026-08-24', '2026-12-18'],
        ['WINTER 26', 'Winter 2026', '2026-12-19', '2027-01-17'],
        ['SPRING 27', 'Spring 2027', '2027-01-18', '2027-05-14'],
        ['SUMMER 27', 'Summer 2027', '2027-05-15', '2027-08-23'],
      ]),
    },
  };
  // A few legacy rows the way the real workbook has them: unparseable time, missing date.
  const fall = active.tables.tbl_FALL26.rows;
  fall[3][2] = 'TBD w/ prof';
  fall.push(HEADERS.map((h) => ({ 'Term Year': 'Fall 2026', 'Course #': 'EN 104', 'Faculty Name': 'Prof. Brennan', _RecordId: newRecordId(), _RowVersion: 1 }[h] ?? '')));

  const archived = {
    name: '2025-2026 Instruction Classes.xlsx',
    tables: {
      tbl_FALL25: { sheet: 'FALL 25', headers: HEADERS, rows: makeTerm(rng(7), 'Fall 2025', '2025-08-25', '2025-12-19', 52, today) },
      tbl_SPRING26: { sheet: 'SPRING 26', headers: HEADERS, rows: makeTerm(rng(9), 'Spring 2026', '2026-01-20', '2026-05-08', 41, today) },
      ...listsTables([
        ['FALL 25', 'Fall 2025', '2025-08-25', '2025-12-19', true],
        ['SPRING 26', 'Spring 2026', '2026-01-20', '2026-05-15', true],
      ]),
    },
  };
  archived.tables.tbl_AppLock.rows = [['Archived', 'kathleen@desales.edu', '2026-08-01T12:00:00Z']];
  const reqRow = (o) => REQUEST_HEADERS.map((h) => o[h] ?? '');
  const requests = {
    name: 'Instruction Requests.xlsx',
    tables: {
      tbl_Requests: {
        sheet: 'REQUESTS',
        headers: REQUEST_HEADERS,
        rows: [
          reqRow({ 'Request ID': 'R-1041', 'Submitted At': '2026-11-16T14:02:00Z', 'Faculty Name': 'Dr. Okafor', 'Faculty Email': 'faculty.one@example.edu', 'Course #': 'NU 805', 'College/School': 'School of Nursing', 'Student Type': 'Graduate', 'Preferred Date': '12/2/2026', 'Alternate Date': '12/3/2026', Time: '6–8:30 PM', 'Session Format': 'Synchronous', Location: 'TEAMS', '# Students': 22, 'Notes / Goals': 'Evidence-based practice databases (CINAHL) for the capstone lit review.', Status: 'New' }),
          reqRow({ 'Request ID': 'R-1042', 'Submitted At': '2026-11-17T09:40:00Z', 'Faculty Name': 'Prof. Gallo', 'Faculty Email': 'faculty.two@example.edu', 'Course #': 'CJ 250', 'College/School': 'College of Arts & Sciences', 'Student Type': 'Undergraduate', 'Preferred Date': '1/26/2027', Time: 'TR 11–12:15', 'Session Format': 'Face to face', Location: 'Dlg 105', '# Students': 28, 'Notes / Goals': 'Case law research basics.', Status: 'New' }),
          reqRow({ 'Request ID': 'R-1037', 'Submitted At': '2026-11-02T16:12:00Z', 'Faculty Name': 'Dr. Weiss', 'Course #': 'BI 151', 'Preferred Date': '11/10/2026', 'Session Format': 'Face to face', Status: 'Accepted', 'Handled By': 'lisa@desales.edu' }),
        ],
      },
    },
  };
  return { active, archived, requests };
}

export class MemoryBackend {
  constructor(today) {
    // Cloned so each table has its own header array (the builder shares one), like separate sheets in Excel.
    const { active, archived, requests } = clone(buildDemoWorkbooks(today));
    // Item ids are opaque, like SharePoint's: "active" keeps its id after it is archived.
    this.books = { active: { ...active, folder: 'root' }, archived: { ...archived, folder: 'Archive' }, requests: { ...requests, folder: 'root' } };
    this.delay = 120;
    this.nextId = 1;
  }

  async wait() {
    await new Promise((r) => setTimeout(r, this.delay));
  }

  /** Same rules as GraphBackend.discover: newest year workbook in the main folder is active. */
  async discover() {
    await this.wait();
    const yearRe = /^(\d{4})-(\d{4}) Instruction Classes\.xlsx$/i;
    const pick = (id) => ({ id, name: this.books[id].name, year: this.books[id].name.match(yearRe).slice(1, 3).join('-'), webUrl: '#' });
    const years = Object.keys(this.books).filter((id) => yearRe.test(this.books[id].name)).sort((a, b) => this.books[b].name.localeCompare(this.books[a].name));
    const inRoot = years.filter((id) => this.books[id].folder === 'root');
    const reqId = Object.keys(this.books).find((id) => /^instruction requests\.xlsx$/i.test(this.books[id].name));
    return {
      active: pick(inRoot[0]),
      archived: years.filter((id) => id !== inRoot[0]).map(pick),
      requests: reqId ? { id: reqId, name: this.books[reqId].name, webUrl: '#' } : null,
    };
  }

  activeId() {
    const yearRe = /^\d{4}-\d{4} Instruction Classes\.xlsx$/i;
    return Object.keys(this.books).filter((id) => this.books[id].folder === 'root' && yearRe.test(this.books[id].name)).sort((a, b) => this.books[b].name.localeCompare(this.books[a].name))[0];
  }

  tableInfo(itemId, name) {
    const t = this.books[itemId].tables[name];
    return { name, sheet: t.sheet, address: { sheet: t.sheet, c0: 0, r0: 1, c1: t.headers.length - 1, r1: t.rows.length + 1 }, headers: clone(t.headers), rows: clone(t.rows) };
  }

  async readTables(itemId) {
    await this.wait();
    return Object.keys(this.books[itemId].tables).map((n) => this.tableInfo(itemId, n));
  }

  async readTable(itemId, name) {
    await this.wait();
    return this.tableInfo(itemId, name);
  }

  async writeRow(itemId, t, index, row) {
    await this.wait();
    this.books[itemId].tables[t.name].rows[index] = clone(row);
  }

  async addRow(itemId, tableName, row) {
    await this.wait();
    this.books[itemId].tables[tableName].rows.push(clone(row));
  }

  async deleteRow(itemId, t, index) {
    await this.wait();
    this.books[itemId].tables[t.name].rows.splice(index, 1);
  }

  async listSheets(itemId) {
    const b = this.books[itemId];
    return [...new Set([...Object.values(b.tables).map((t) => t.sheet), ...(b.extraSheets || [])])];
  }

  async addWorksheet(itemId, name) {
    (this.books[itemId].extraSheets ||= []).push(name);
  }

  async writeRange(itemId, sheet, address, values) {
    const b = this.books[itemId];
    const a = parseAddress(`'${sheet}'!${address}`);
    const grid = ((b.grids ||= {})[sheet] ||= {});
    values.forEach((row, i) => row.forEach((v, j) => (grid[`${a.r0 + i},${a.c0 + j}`] = v)));
  }

  async createTableAt(itemId, address, name) {
    const b = this.books[itemId];
    const a = parseAddress(address);
    const grid = b.grids?.[a.sheet] || {};
    const read = (r) => Array.from({ length: a.c1 - a.c0 + 1 }, (_, j) => grid[`${r},${a.c0 + j}`] ?? '');
    const rows = [];
    for (let r = a.r0 + 1; r <= a.r1; r++) rows.push(read(r));
    b.tables[name] = { sheet: a.sheet, headers: read(a.r0).map(String), rows: rows.length ? rows : [read(a.r0).map(() => '')] };
  }

  async hideSheet() {}

  async renameSheet(itemId, from, to) {
    for (const t of Object.values(this.books[itemId].tables)) if (t.sheet === from) t.sheet = to;
  }

  async renameTable(itemId, from, to) {
    const tables = this.books[itemId].tables;
    tables[to] = tables[from];
    delete tables[from];
  }

  /** Like Excel: an emptied table keeps one blank row. */
  async clearTableRows(itemId, t) {
    await this.wait();
    const tb = this.books[itemId].tables[t.name];
    tb.rows = [tb.headers.map(() => '')];
  }

  async findInFolder(name) {
    const id = Object.keys(this.books).find((k) => this.books[k].name === name && this.books[k].folder === 'root');
    return id ? { id, name } : null;
  }

  async copyWorkbook(itemId, name) {
    await this.wait();
    const id = `wb${this.nextId++}`;
    this.books[id] = { ...clone(this.books[itemId]), name, folder: 'root' };
    return { id, name, webUrl: '#' };
  }

  async ensureFolder(name) {
    return name;
  }

  async moveItem(itemId, folderId) {
    const b = this.books[itemId];
    if (Object.values(this.books).some((x) => x !== b && x.folder === folderId && x.name === b.name)) throw new Error(`“${b.name}” already exists in ${folderId}.`);
    b.folder = folderId;
  }
  async createTable() {}
  async addColumn(itemId, t, header) {
    const tb = this.books[itemId].tables[t.name];
    tb.headers.push(header);
    tb.rows.forEach((r) => r.push(''));
  }
  async renameColumn(itemId, t, colIdx, header) {
    await this.wait();
    this.books[itemId].tables[t.name].headers[colIdx] = header;
  }
  /** Kept per table + header so tests can check what Excel would show. */
  async setListValidation(itemId, t, colIdx, source, { strict = true } = {}) {
    const b = this.books[itemId];
    (b.validation ||= {})[`${t.name}:${b.tables[t.name].headers[colIdx]}`] = { source, strict };
  }
  /** Reads a block of the Lists sheet grid (demo/test helper). */
  gridColumn(itemId, sheet, col, fromRow, toRow) {
    const grid = this.books[itemId].grids?.[sheet] || {};
    const out = [];
    for (let r = fromRow; r <= toRow; r++) out.push(grid[`${r},${col}`] ?? '');
    return out;
  }
  async writeColumn(itemId, t, colIdx, values) {
    values.forEach((v, i) => (this.books[itemId].tables[t.name].rows[i][colIdx] = v));
  }
  async hideColumns() {}
  async createRequestsWorkbook(headers) {
    this.books.requests = { name: 'Instruction Requests.xlsx', folder: 'root', tables: { tbl_Requests: { sheet: 'REQUESTS', headers, rows: [] } } };
    return { id: 'requests', name: 'Instruction Requests.xlsx', webUrl: '#' };
  }

  /** Demo helper: pretend a colleague on another PC edited a row (to try the merge and conflict dialog). */
  simulateColleagueEdit(recordId, header, value) {
    for (const t of Object.values(this.books[this.activeId()].tables)) {
      const idCol = t.headers.indexOf('_RecordId');
      const row = t.rows.find((r) => r[idCol] === recordId);
      if (!row) continue;
      row[t.headers.indexOf(header)] = value;
      row[t.headers.indexOf('_RowVersion')] = Number(row[t.headers.indexOf('_RowVersion')] || 0) + 1;
      row[t.headers.indexOf('_ModifiedBy')] = 'lisa@desales.edu';
      row[t.headers.indexOf('_ModifiedAt')] = new Date().toISOString();
      return true;
    }
    return false;
  }
}
