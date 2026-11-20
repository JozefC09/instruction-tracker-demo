// App shell: sign-in, workbook discovery, routing, and the sync loop (spec §5.1).
import { initAuth, signIn, signOut, currentUser } from './auth.js';
import { GraphBackend } from './graph.js';
import { MemoryBackend } from './demo.js';
import { Store, lockedByOther } from './store.js';
import { todayIso } from './model.js';
import { $, $$, esc, timeAgo, hideTip, toast } from './ui.js';
import { renderDashboard } from './dashboard.js';
import { renderSessions, renderTours, closeEditor, isEditorOpen } from './sessions.js';
import { renderRequests } from './requests.js';
import { renderSettings } from './settings.js';
import { renderNewYear } from './newyear.js';

const cfg = window.TRACKER_CONFIG || {};
const params = new URLSearchParams(location.search);

export const app = {
  cfg,
  lists: cfg.lists,
  demo: !cfg.clientId || params.has('demo'),
  today: todayIso(),
  store: null,
  books: null,
  viewBookId: null, // workbook shown on Dashboard / Sessions (active year or an archived one)
  cache: new Map(), // bookId → loaded data
  requests: null,
  lastSync: 0,
  offline: false,
  changed: new Set(), // record ids changed by someone else on the last sync (highlighted)
  ui: {}, // per-view state (filters, tabs…) — kept across re-renders

  get data() { return this.cache.get(this.viewBookId); },
  get active() { return this.cache.get(this.books.active.id); },
  bookById(id) { return [this.books.active, ...this.books.archived].find((b) => b.id === id); },

  libColor(name) {
    const lib = this.lists.librarians.find((l) => l.name.toLowerCase() === String(name || '').toLowerCase());
    return lib ? `var(--series-${lib.color})` : 'var(--series-none)';
  },

  async load(bookId, { quiet = false } = {}) {
    const prev = this.cache.get(bookId);
    const data = await this.store.loadBook(this.bookById(bookId));
    if (prev && quiet) {
      const old = new Map(prev.records.map((r) => [r.id, r.hash]));
      for (const r of data.records) if (old.has(r.id) && old.get(r.id) !== r.hash) this.changed.add(r.id);
      // Rows added or deleted elsewhere also need a repaint.
      if (data.records.length !== prev.records.length || data.records.some((r) => !old.has(r.id))) this.changed.add('*');
    }
    this.cache.set(bookId, data);
    if (bookId === this.books.active.id) {
      // Term dates live in the workbook's Lists sheet once it exists; config.js only supplies the defaults.
      const terms = data.terms.filter((t) => t.from && t.to).map((t) => ({ name: t.sheet, from: t.from, to: t.to }));
      this.lists = { ...cfg.lists, terms: terms.length ? terms : cfg.lists.terms };
      this.locked = lockedByOther(data.lock, this.user.upn) ? data.lock : null;
    }
    return data;
  },

  /** Writes are allowed in the active year when online, prepared, and nobody else is mid-transition. */
  get canWrite() {
    return !this.offline && !this.locked && !this.active.tables.some((t) => t.missingSystem.length);
  },

  updateBanner() {
    const unprepared = this.active?.tables.some((t) => t.missingSystem.length) || (this.active && !this.active.hasLists);
    if (this.offline) return this.banner('<strong>Offline.</strong> Showing the last data loaded. Changes can’t be saved until the connection is back.', 'error');
    if (this.locked?.action === 'Archived') return this.banner('<strong>This year has been archived.</strong> Switching to the new academic year…', 'warn');
    if (this.locked) return this.banner(`<strong>Editing is paused.</strong> ${esc(this.locked.user)} is ${esc(this.locked.action.toLowerCase())}. This page updates by itself when they’re done.`, 'warn');
    if (unprepared) return this.banner(`This workbook hasn’t been fully prepared for the app yet${this.active.tables.some((t) => t.missingSystem.length) ? ', so it is <strong>read-only</strong> here' : ''}. The coordinator can run <a href="#/settings">Prepare workbook</a> once.`, 'warn');
    if (this.demo) return this.banner(this.demoBannerHtml, 'warn');
    this.banner('');
  },

  /** Re-checks the Instruction folder; if a new academic year has been started elsewhere, switch to it (spec F32). */
  async rediscover() {
    const prevActive = this.books.active.id;
    const books = await this.store.init();
    this.books = books;
    lastDiscover = Date.now();
    if (books.active.id === prevActive) return false;
    const wasViewingActive = this.viewBookId === prevActive;
    this.cache.clear();
    this.ui = {};
    this.viewBookId = wasViewingActive || !this.bookById(this.viewBookId) ? books.active.id : this.viewBookId;
    await Promise.all([this.load(books.active.id), this.loadRequests()]);
    if (this.viewBookId !== books.active.id) await this.load(this.viewBookId);
    toast(`Now using ${books.active.name}. Last year is under Archive (read-only).`);
    return true;
  },

  async loadRequests() {
    try {
      this.requests = await this.store.loadRequests();
    } catch (e) {
      console.warn('Requests workbook could not be read', e);
      this.requests = null;
    }
    const n = this.requests?.items.filter((r) => r.status === 'New').length ?? 0;
    $('#req-count').textContent = n;
    $('#req-count').hidden = !n;
  },

  /** Re-read after a write, then re-render the current route. */
  async refresh() {
    await Promise.all([this.load(this.books.active.id), this.viewBookId !== this.books.active.id ? this.load(this.viewBookId) : null]);
    this.lastSync = Date.now();
    this.render();
  },

  render() {
    const route = (location.hash.replace(/^#\/?/, '') || 'dashboard').split('/');
    const navRoute = route[0] === 'new-year' ? 'settings' : route[0];
    $$('#sidenav a').forEach((a) => a.classList.toggle('active', a.dataset.route === navRoute));
    const main = $('#main');
    hideTip();
    switch (route[0]) {
      case 'sessions': return renderSessions(main, route.slice(1));
      case 'tours': return renderTours(main);
      case 'requests': return renderRequests(main);
      case 'settings': return renderSettings(main);
      case 'new-year': return renderNewYear(main);
      default: return renderDashboard(main);
    }
  },

  banner(html, kind = '') {
    const b = $('#banner');
    b.className = `banner ${kind}`;
    b.innerHTML = html;
    b.hidden = !html;
  },
};
window.trackerApp = app; // handy for debugging from the console

function setSync(text, cls = '') {
  const s = $('#sync-status');
  s.className = `sync ${cls}`;
  s.textContent = text;
}

// ---------- Sync loop: 30 s while visible, 2 min in the background ----------
let syncing = false;
async function syncTick() {
  if (syncing || !app.store) return;
  // Don't repaint the page under someone mid-edit; conflicts are caught on save anyway.
  if (isEditorOpen() || document.querySelector('#dialog[open]')) return;
  syncing = true;
  try {
    app.changed.clear();
    const before = JSON.stringify([app.locked, app.active?.terms, app.offline]);
    // Look for a new academic year every 5 minutes, and right away once a rollover lock clears or the year is archived.
    const lockWasHeld = !!app.locked;
    let switched = false;
    if (Date.now() - lastDiscover > 5 * 60 * 1000 || app.locked?.action === 'Archived') switched = await app.rediscover();
    if (!switched) {
      await app.load(app.books.active.id, { quiet: true });
      if (lockWasHeld && !app.locked) switched = await app.rediscover();
    }
    if (app.viewBookId !== app.books.active.id && !app.cache.has(app.viewBookId)) await app.load(app.viewBookId);
    await app.loadRequests();
    app.lastSync = Date.now();
    app.offline = false;
    if (switched || app.changed.size || JSON.stringify([app.locked, app.active?.terms, app.offline]) !== before) {
      app.updateBanner();
      app.render();
    }
  } catch (e) {
    console.warn('Sync failed', e);
    if (!app.offline) {
      app.offline = true;
      app.updateBanner();
    }
  } finally {
    syncing = false;
  }
}

let lastTick = 0;
let lastDiscover = Date.now();
export function syncNow() {
  lastTick = Date.now();
  return syncTick();
}
setInterval(() => {
  if (!app.store) return;
  setSync(app.offline ? 'Offline' : app.lastSync ? `Synced ${timeAgo(app.lastSync)}` : 'Loading…', app.offline ? 'offline' : '');
  const every = (document.hidden ? cfg.syncSecondsHidden || 120 : cfg.syncSecondsVisible || 30) * 1000;
  if (Date.now() - lastTick >= every) {
    lastTick = Date.now();
    syncTick();
  }
}, 1000);

document.addEventListener('keydown', (e) => {
  if (e.key === 'F5' && app.store && !e.ctrlKey) {
    e.preventDefault();
    lastTick = Date.now();
    syncTick().then(() => app.render());
  }
  if (e.key === 'Escape' && isEditorOpen() && !document.querySelector('#dialog[open]')) closeEditor();
});

window.addEventListener('hashchange', () => app.store && app.render());

// ---------- Start-up ----------
function showSignIn(message = '') {
  $('#main').innerHTML = `
    <div class="signin card">
      <h1>Library Instruction Tracker</h1>
      <p>Sign in with your DeSales Microsoft 365 account to view and edit instruction sessions.</p>
      ${message ? `<p class="err">${esc(message)}</p>` : ''}
      <button class="btn primary" id="btn-signin">Sign in with Microsoft</button>
      <p class="small muted">Faculty requesting a session: use the <a href="request-form.html">request form</a>.</p>
    </div>`;
  $('#btn-signin').onclick = async () => {
    try {
      if (await signIn()) start();
    } catch (e) {
      showSignIn(e.message);
    }
  };
}

async function start() {
  let backend, user;
  if (app.demo) {
    backend = new MemoryBackend(app.today);
    app.memory = backend;
    user = { name: 'Demo Librarian', upn: 'demo.librarian@desales.edu' };
    app.demoBannerHtml = '<strong>Demo mode.</strong> Sample data only, kept in this browser tab. Nothing is read from or written to SharePoint. Set <code>clientId</code> in config.js to connect.';
  } else {
    user = currentUser();
    backend = new GraphBackend(cfg);
    $('#btn-signout').hidden = false;
  }
  $('#user-name').textContent = user.name;
  app.store = new Store(backend, user);
  app.user = user;
  $('#main').innerHTML = '<div class="empty">Loading the instruction workbook…</div>';
  setSync('Loading…', 'busy');
  try {
    app.books = await app.store.init();
    app.viewBookId = app.books.active.id;
    await Promise.all([app.load(app.books.active.id), app.loadRequests()]);
    app.lastSync = lastTick = Date.now();
    $('#sidenav').hidden = false;
    lastDiscover = Date.now();
    app.updateBanner();
    app.render();
  } catch (e) {
    console.error(e);
    setSync('Error', 'offline');
    $('#main').innerHTML = `<div class="card"><h1>Couldn’t open the workbook</h1><p class="err">${esc(e.message)}</p>
      <p class="muted">Check the Instruction folder link in config.js and that your account can open it in SharePoint.</p>
      <button class="btn" onclick="location.reload()">Try again</button></div>`;
  }
}

async function boot() {
  $('#btn-signout').onclick = () => signOut();
  if (app.demo) return start();
  try {
    const account = await initAuth(cfg);
    if (account) start();
    else showSignIn();
  } catch (e) {
    console.error(e);
    showSignIn(e.message);
  }
}

// MSAL and SheetJS are loaded with `defer`; module scripts run after them, so they're ready here.
boot();
