// Low-level SharePoint workbook access through the Microsoft Graph Excel API.
// Everything here addresses rows by table + index right after a fresh read; the store (store.js)
// re-finds rows by _RecordId before every write, so indexes are never cached between operations.
import { getToken } from './auth.js';

const GRAPH = 'https://graph.microsoft.com/v1.0';
const RETRY_STATUSES = new Set([409, 423, 429, 500, 502, 503, 504]);

export class GraphError extends Error {
  constructor(status, body) {
    super(body?.error?.message || `Microsoft Graph returned ${status}`);
    this.status = status;
    this.code = body?.error?.code;
  }
}

async function call(method, path, body, { raw = false, headers = {} } = {}) {
  for (let attempt = 0; ; attempt++) {
    const token = await getToken();
    const res = await fetch(path.startsWith('http') ? path : GRAPH + path, {
      method,
      headers: {
        Authorization: `Bearer ${token}`,
        ...(body && !raw ? { 'Content-Type': 'application/json' } : {}),
        ...headers,
      },
      body: body ? (raw ? body : JSON.stringify(body)) : undefined,
    });
    if (res.ok) {
      const text = await res.text(); // 202/204 (e.g. copy, delete) have no body
      return text ? JSON.parse(text) : null;
    }
    // Locked by a non-co-authoring Excel session, throttled, or a transient server error: back off and retry (spec §5.5).
    if (RETRY_STATUSES.has(res.status) && attempt < 3) {
      const wait = Number(res.headers.get('Retry-After')) * 1000 || 800 * 2 ** attempt;
      await new Promise((r) => setTimeout(r, wait));
      continue;
    }
    let err = null;
    try { err = await res.json(); } catch { /* no body */ }
    const e = new GraphError(res.status, err);
    if (RETRY_STATUSES.has(res.status)) e.busy = true;
    throw e;
  }
}

/** 'A' → 0, 'AB' → 27 */
export function colIndex(letters) {
  return [...letters.toUpperCase()].reduce((n, c) => n * 26 + (c.charCodeAt(0) - 64), 0) - 1;
}
export function colLetters(i) {
  let s = '';
  for (i += 1; i > 0; i = Math.floor((i - 1) / 26)) s = String.fromCharCode(65 + ((i - 1) % 26)) + s;
  return s;
}

/** "'FALL 26'!A1:R40" → { sheet: 'FALL 26', c0: 0, r0: 1, c1: 17, r1: 40 } */
export function parseAddress(address) {
  const bang = address.lastIndexOf('!');
  let sheet = address.slice(0, bang);
  if (sheet.startsWith("'")) sheet = sheet.slice(1, -1).replace(/''/g, "'");
  const [a, b = a] = address.slice(bang + 1).split(':');
  const p = (x) => {
    const m = x.match(/^\$?([A-Z]+)\$?(\d+)$/i);
    return { c: colIndex(m[1]), r: Number(m[2]) };
  };
  const s = p(a), e = p(b);
  return { sheet, c0: s.c, r0: s.r, c1: e.c, r1: e.r };
}

const sheetPath = (name) => `worksheets/${encodeURIComponent(name)}`;

/** Text that Excel would treat as a formula is stored as text instead. */
function safeCell(v) {
  return typeof v === 'string' && /^[=]/.test(v) ? `'${v}` : v;
}

const shareId = (url) => 'u!' + btoa(unescape(encodeURIComponent(url))).replace(/=+$/, '').replace(/\//g, '_').replace(/\+/g, '-');

export class GraphBackend {
  constructor(cfg) {
    this.cfg = cfg;
    this.driveId = null;
    this.folderId = null;
  }

  wb(itemId) {
    return `/drives/${this.driveId}/items/${itemId}/workbook`;
  }

  /** Resolves the Instruction folder link and finds the active year, the archive and the requests workbook (spec F2). */
  async discover() {
    if (!this.cfg.instructionFolderUrl) throw new Error('config.js has no instructionFolderUrl. Paste the SharePoint link to the Instruction folder there.');
    let item = await call('GET', `/shares/${shareId(this.cfg.instructionFolderUrl)}/driveItem?$select=id,name,folder,file,parentReference`);
    this.driveId = item.parentReference.driveId;
    if (!item.folder) item = await call('GET', `/drives/${this.driveId}/items/${item.parentReference.id}?$select=id,name,folder`);
    this.folderId = item.id;

    const list = async (id) => (await call('GET', `/drives/${this.driveId}/items/${id}/children?$select=id,name,folder,file,webUrl&$top=500`)).value;
    const children = await list(this.folderId);
    const yearRe = /^(\d{4})-(\d{4}) Instruction Classes\.xlsx$/i;
    const years = children.filter((c) => c.file && yearRe.test(c.name)).sort((a, b) => b.name.localeCompare(a.name));
    const archiveFolder = children.find((c) => c.folder && /^archive$/i.test(c.name));
    const archived = archiveFolder ? (await list(archiveFolder.id)).filter((c) => c.file && yearRe.test(c.name)).sort((a, b) => b.name.localeCompare(a.name)) : [];
    const requests = children.find((c) => c.file && /^instruction requests\.xlsx$/i.test(c.name)) ?? null;
    if (!years.length) throw new Error('No "YYYY-YYYY Instruction Classes.xlsx" workbook was found in the Instruction folder.');
    const pick = (c) => ({ id: c.id, name: c.name, webUrl: c.webUrl, year: c.name.match(yearRe)?.slice(1, 3).join('-') });
    return {
      active: pick(years[0]),
      // Older workbooks still sitting in the main folder are treated like archived ones.
      archived: [...years.slice(1), ...archived].map(pick),
      requests: requests ? pick(requests) : null,
    };
  }

  async readTables(itemId) {
    const tables = (await call('GET', `${this.wb(itemId)}/tables?$select=id,name`)).value;
    return Promise.all(tables.map((t) => this.readTable(itemId, t.name)));
  }

  async readTable(itemId, name) {
    const r = await call('GET', `${this.wb(itemId)}/tables/${encodeURIComponent(name)}/range?$select=address,values`);
    const a = parseAddress(r.address);
    return { name, sheet: a.sheet, address: a, headers: r.values[0].map(String), rows: r.values.slice(1) };
  }

  rowAddress(t, index) {
    const row = t.address.r0 + 1 + index;
    return `${colLetters(t.address.c0)}${row}:${colLetters(t.address.c1)}${row}`;
  }

  async writeRow(itemId, t, index, row) {
    await call('PATCH', `${this.wb(itemId)}/${sheetPath(t.sheet)}/range(address='${this.rowAddress(t, index)}')`, { values: [row.map(safeCell)] });
  }

  async addRow(itemId, tableName, row) {
    await call('POST', `${this.wb(itemId)}/tables/${encodeURIComponent(tableName)}/rows/add`, { index: null, values: [row.map(safeCell)] });
  }

  async deleteRow(itemId, t, index) {
    await call('POST', `${this.wb(itemId)}/${sheetPath(t.sheet)}/range(address='${this.rowAddress(t, index)}')/delete`, { shift: 'Up' });
  }

  // ---- Used by "Prepare workbook" ----

  async listSheets(itemId) {
    return (await call('GET', `${this.wb(itemId)}/worksheets?$select=name`)).value.map((w) => w.name);
  }

  async createTable(itemId, sheet, tableName) {
    const used = await call('GET', `${this.wb(itemId)}/${sheetPath(sheet)}/usedRange(valuesOnly=true)?$select=address`);
    const t = await call('POST', `${this.wb(itemId)}/tables/add`, { address: used.address, hasHeaders: true });
    await call('PATCH', `${this.wb(itemId)}/tables/${t.id}`, { name: tableName, showBandedRows: true });
  }

  async addColumn(itemId, t, header) {
    const values = [[header], ...t.rows.map(() => [''])];
    await call('POST', `${this.wb(itemId)}/tables/${encodeURIComponent(t.name)}/columns/add`, { index: null, values });
  }

  /** Writes one whole column of data rows (used to assign _RecordId / _RowVersion). */
  async writeColumn(itemId, t, colIdx, values) {
    if (!values.length) return;
    const L = colLetters(t.address.c0 + colIdx);
    const addr = `${L}${t.address.r0 + 1}:${L}${t.address.r0 + values.length}`;
    await call('PATCH', `${this.wb(itemId)}/${sheetPath(t.sheet)}/range(address='${addr}')`, { values: values.map((v) => [v]) });
  }

  async hideColumns(itemId, t, fromIdx, toIdx) {
    const addr = `${colLetters(t.address.c0 + fromIdx)}:${colLetters(t.address.c0 + toIdx)}`;
    await call('PATCH', `${this.wb(itemId)}/${sheetPath(t.sheet)}/range(address='${addr}')/format`, { columnHidden: true });
  }

  /** Renames a table column by writing its header cell (Excel renames the table column to match). */
  async renameColumn(itemId, t, colIdx, header) {
    const addr = `${colLetters(t.address.c0 + colIdx)}${t.address.r0}`;
    await call('PATCH', `${this.wb(itemId)}/${sheetPath(t.sheet)}/range(address='${addr}')`, { values: [[safeCell(header)]] });
  }

  /**
   * Replaces the data validation on one table column's data rows with an in-cell dropdown.
   * source is a range on the Lists sheet, e.g. "=Lists!$M$2:$M$5". New table rows inherit it.
   * strict: Excel shows a warning (with "continue anyway") for a value that isn't in the list.
   */
  async setListValidation(itemId, t, colIdx, source, { strict = true, name = '' } = {}) {
    const L = colLetters(t.address.c0 + colIdx);
    const addr = `${L}${t.address.r0 + 1}:${L}${Math.max(t.address.r1, t.address.r0 + 1)}`;
    const base = `${this.wb(itemId)}/${sheetPath(t.sheet)}/range(address='${addr}')/dataValidation`;
    await call('POST', `${base}/clear`);
    await call('PATCH', base, {
      rule: { list: { inCellDropDown: true, source } },
      ignoreBlanks: true,
      errorAlert: { showAlert: strict, style: 'Warning', title: 'Not in the list', message: `That ${name || 'value'} isn't in the Lists sheet. The tracker app uses the lists in its config.js.` },
    });
  }

  // ---- Used by the Lists sheet, Close term and Start new academic year ----

  async writeRange(itemId, sheet, address, values) {
    await call('PATCH', `${this.wb(itemId)}/${sheetPath(sheet)}/range(address='${address}')`, { values: values.map((r) => r.map(safeCell)) });
  }

  /** address like "Lists!A1:G5" (header row first). */
  async createTableAt(itemId, address, name) {
    const t = await call('POST', `${this.wb(itemId)}/tables/add`, { address, hasHeaders: true });
    await call('PATCH', `${this.wb(itemId)}/tables/${t.id}`, { name });
  }

  async addWorksheet(itemId, name) {
    await call('POST', `${this.wb(itemId)}/worksheets/add`, { name });
  }

  async hideSheet(itemId, name) {
    await call('PATCH', `${this.wb(itemId)}/${sheetPath(name)}`, { visibility: 'Hidden' });
  }

  async renameSheet(itemId, from, to) {
    await call('PATCH', `${this.wb(itemId)}/${sheetPath(from)}`, { name: to });
  }

  async renameTable(itemId, from, to) {
    await call('PATCH', `${this.wb(itemId)}/tables/${encodeURIComponent(from)}`, { name: to });
  }

  /** Empties a table. Excel keeps one (blank) data row in an emptied table; the app ignores blank rows. */
  async clearTableRows(itemId, t) {
    const n = t.rows.length;
    if (!n) return;
    const L0 = colLetters(t.address.c0), L1 = colLetters(t.address.c1);
    const first = t.address.r0 + 1;
    if (n > 1) {
      await call('POST', `${this.wb(itemId)}/${sheetPath(t.sheet)}/range(address='${L0}${first + 1}:${L1}${first + n - 1}')/delete`, { shift: 'Up' });
    }
    await call('POST', `${this.wb(itemId)}/${sheetPath(t.sheet)}/range(address='${L0}${first}:${L1}${first}')/clear`, { applyTo: 'Contents' });
  }

  async findInFolder(name, folderId = this.folderId) {
    try {
      return await call('GET', `/drives/${this.driveId}/items/${folderId}:/${encodeURIComponent(name)}?$select=id,name,webUrl,size`);
    } catch (e) {
      if (e.status === 404) return null;
      throw e;
    }
  }

  /** Server-side copy into the Instruction folder. Keeps formatting, validation and table styles. */
  async copyWorkbook(itemId, name) {
    await call('POST', `/drives/${this.driveId}/items/${itemId}/copy?@microsoft.graph.conflictBehavior=fail`, {
      parentReference: { driveId: this.driveId, id: this.folderId }, name,
    });
    // The copy runs in the background; wait for the file, then for the workbook API to accept it.
    for (let i = 0; i < 60; i++) {
      const item = await this.findInFolder(name);
      if (item?.size) {
        for (let j = 0; ; j++) {
          try {
            await call('GET', `${this.wb(item.id)}/tables?$select=name`);
            return { id: item.id, name: item.name, webUrl: item.webUrl };
          } catch (e) {
            if (j >= 10) throw e;
            await new Promise((r) => setTimeout(r, 3000));
          }
        }
      }
      await new Promise((r) => setTimeout(r, 2000));
    }
    throw new Error(`SharePoint didn’t finish copying “${name}” within 2 minutes. Check the Instruction folder, then try again.`);
  }

  async ensureFolder(name) {
    const existing = await this.findInFolder(name);
    if (existing) return existing.id;
    const f = await call('POST', `/drives/${this.driveId}/items/${this.folderId}/children`, { name, folder: {}, '@microsoft.graph.conflictBehavior': 'fail' });
    return f.id;
  }

  async moveItem(itemId, folderId) {
    await call('PATCH', `/drives/${this.driveId}/items/${itemId}`, { parentReference: { id: folderId } });
  }

  /** Creates Instruction Requests.xlsx with the §3.5 header row as tbl_Requests. */
  async createRequestsWorkbook(headers) {
    const XLSX = window.XLSX;
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([headers]), 'REQUESTS');
    const bytes = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
    const item = await call('PUT', `/drives/${this.driveId}/items/${this.folderId}:/${encodeURIComponent('Instruction Requests.xlsx')}:/content?@microsoft.graph.conflictBehavior=fail`,
      new Blob([bytes]), { raw: true, headers: { 'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' } });
    // The workbook API can lag a few seconds behind a fresh upload.
    for (let i = 0; ; i++) {
      try {
        const t = await call('POST', `${this.wb(item.id)}/tables/add`, { address: `REQUESTS!A1:${colLetters(headers.length - 1)}1`, hasHeaders: true });
        await call('PATCH', `${this.wb(item.id)}/tables/${t.id}`, { name: 'tbl_Requests' });
        break;
      } catch (e) {
        if (i >= 4) throw e;
        await new Promise((r) => setTimeout(r, 2000));
      }
    }
    return { id: item.id, name: item.name, webUrl: item.webUrl };
  }
}
