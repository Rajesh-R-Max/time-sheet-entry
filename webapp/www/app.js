/* Timesheet entry - talks to server.ps1, which writes straight into
   Timesheet.xlsx via Excel COM. No build step, no dependencies. */

const $ = id => document.getElementById(id);

let authUser = null;
let users = [];

const FIELDS = ['date','project','task','hours','category','budget','billable',
                'notes','ticket','incidentType','nonBillableReason'];
const DATALISTS = ['project','task','category','budget','billable',
                   'incidentType','nonBillableReason'];
// Everything the form owns, used when loading an entry in for edit/copy.
const CONTEXT = ['project','task','category','budget','billable','nonBillableReason'];

const DAYS = ['Mon','Tue','Wed','Thu','Fri','Sat','Sun'];

let state = { options: {}, projectTasks: {}, projects: [], entries: [], totals: {} };
let editingRow = null;        // null = add mode, otherwise the worksheet row
let reportWeek = null;        // Monday of the week on screen, yyyy-mm-dd
let reportData = null;
let reportCollapsed = new Set();   // collapsed pivot paths, see PSEP below
let dayDate = null;           // day on screen in the Day view, yyyy-mm-dd
let dayWeek = null;           // cached /api/week payload that covers dayDate

/* Bootstrap only sends the most recent rows, so a Day or Report view showing an
   older date has to fall back to whatever that view loaded. */
const findEntry = row =>
  (state.entries || []).find(x => x.row === row) ||
  (dayWeek?.entries || []).find(x => x.row === row) ||
  (reportData?.entries || []).find(x => x.row === row);

/* ---------------------------------------------------------------- utils */

// Mirrors ConvertTo-Hours in server.ps1 so the hint matches what gets saved.
function parseHours(raw) {
  const s = String(raw || '').replace(/\s/g, '').toLowerCase();
  if (!s) return null;
  let m;
  if ((m = s.match(/^(\d+):([0-5]?\d)$/)))         return +m[1] + +m[2] / 60;
  if ((m = s.match(/^(\d+(?:\.\d+)?)h(\d+)m?$/)))  return +m[1] + +m[2] / 60;
  if ((m = s.match(/^(\d+(?:\.\d+)?)h$/)))         return +m[1];
  if ((m = s.match(/^(\d+(?:\.\d+)?)m$/)))         return +m[1] / 60;
  if (/^\.?\d+(\.\d+)?$/.test(s) || /^\d*\.\d+$/.test(s)) return +s;
  return null;
}

const fmtHM = h => {
  const w = Math.floor(h), mi = Math.round((h - w) * 60);
  return `${w}h ${String(mi).padStart(2, '0')}m`;
};

// 5.5 -> "5:30". Rounded to whole minutes first so 1.4999 cannot land on :89.
const fmtClock = h => {
  const mins = Math.round((h || 0) * 60);
  return `${Math.floor(mins / 60)}:${String(mins % 60).padStart(2, '0')}`;
};

const esc = s => String(s).replace(/[&<>"']/g, c =>
  ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

const iso = d => `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
const parseIso = s => { const [y,m,d] = s.split('-').map(Number); return new Date(y, m-1, d); };
const mondayOf = d => { const x = new Date(d); x.setDate(x.getDate() - ((x.getDay() + 6) % 7)); return x; };

// ISO-8601 week number, so it lines up with the Week Number column in Excel.
function isoWeek(d) {
  const t = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  t.setUTCDate(t.getUTCDate() + 4 - (t.getUTCDay() || 7));
  const y0 = new Date(Date.UTC(t.getUTCFullYear(), 0, 1));
  return { year: t.getUTCFullYear(), week: Math.ceil(((t - y0) / 86400000 + 1) / 7) };
}

const prettyDate = d => d.toLocaleDateString(undefined, { day: 'numeric', month: 'short' });

function setMsg(text, kind) {
  const el = $('msg');
  el.textContent = text;
  el.className = 'msg' + (kind ? ' ' + kind : '');
  if (kind === 'ok') setTimeout(() => { if (el.textContent === text) el.textContent = ''; }, 4000);
}

const fillList = (id, values) => {
  $('dl-' + id).innerHTML = (values || []).map(v => `<option value="${esc(v)}">`).join('');
};

function flashForm() {
  const fc = document.querySelector('.form-card');
  fc.classList.remove('form-flash');
  void fc.offsetWidth;                       // restart the animation
  fc.classList.add('form-flash');
  fc.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

/* ---------------------------------------------------------------- render */

function render() {
  $('totalToday').textContent = fmtClock(state.totals.today);
  $('totalWeek').textContent  = fmtClock(state.totals.week);
  $('totalCount').textContent = state.totals.count  ?? 0;

  DATALISTS.forEach(f => fillList(f, state.options[f]));
  refreshTaskList();
  renderProjectAdmin();

  const box = $('entries');
  const rows = [...(state.entries || [])].reverse();   // newest first
  if (!rows.length) { box.innerHTML = '<p class="empty">No entries yet.</p>'; return; }

  // group consecutive rows by date
  const groups = [];
  for (const e of rows) {
    const last = groups[groups.length - 1];
    if (last && last.date === e.date) last.items.push(e);
    else groups.push({ date: e.date, dateText: e.dateText, day: e.day, items: [e] });
  }

  box.innerHTML = groups.map(g => {
    const total = g.items.reduce((s, e) => s + (e.hours || 0), 0);
    return `<div class="daygroup">
      <div class="dayhead">
        <span>${esc(g.day)} ${esc(g.dateText)}</span>
        <span class="dh-total">${total.toFixed(2)} h</span>
      </div>
      ${g.items.map(entryHtml).join('')}
    </div>`;
  }).join('');

  box.querySelectorAll('.del').forEach(b =>
    b.addEventListener('click', () => del(+b.dataset.row)));
  box.querySelectorAll('.copy-btn').forEach(b =>
    b.addEventListener('click', () => copyEntry(+b.dataset.row)));
  box.querySelectorAll('.edit-btn').forEach(b =>
    b.addEventListener('click', () => startEdit(+b.dataset.row)));
}

function entryHtml(e) {
  const tagDefs = [];
  if (e.category)                       tagDefs.push({ text: e.category,          cls: 'tag-category' });
  if (e.budget === 'Fixed Fee')         tagDefs.push({ text: e.budget,            cls: 'tag-fixed' });
  else if (e.budget === 'Non Billable') tagDefs.push({ text: e.budget,            cls: 'tag-nb-budget' });
  else if (e.budget)                    tagDefs.push({ text: e.budget,            cls: '' });
  if (e.billable === 'Yes')             tagDefs.push({ text: 'Billable',          cls: 'tag-billable' });
  else if (e.billable === 'No')         tagDefs.push({ text: 'Non-billable',      cls: 'tag-nonbill' });
  if (e.nonBillableReason)              tagDefs.push({ text: e.nonBillableReason, cls: 'tag-reason' });
  if (e.ticket)                         tagDefs.push({ text: e.ticket,            cls: 'tag-ticket' });
  if (e.incidentType)                   tagDefs.push({ text: e.incidentType,      cls: 'tag-ticket' });

  const tags   = tagDefs.map(t => `<span class="tag ${t.cls}">${esc(t.text)}</span>`).join('');
  const border = e.billable === 'Yes' ? 'is-billable' : e.billable === 'No' ? 'is-nonbillable' : '';
  const active = editingRow === e.row ? ' is-editing' : '';

  return `<div class="entry ${border}${active}">
    <div class="e-main">
      <div class="e-proj">${esc(e.project || '(no project)')}</div>
      ${e.task  ? `<div class="e-task">${esc(e.task)}</div>`   : ''}
      ${e.notes ? `<div class="e-notes">${esc(e.notes)}</div>` : ''}
      ${tags    ? `<div class="e-tags">${tags}</div>`          : ''}
    </div>
    <div class="e-hours">${(e.hours || 0).toFixed(2)}h</div>
    <button class="edit-btn" data-row="${e.row}" title="Edit this entry">&#9998;</button>
    <button class="copy-btn" data-row="${e.row}" title="Copy to form">&#x29C9;</button>
    <button class="del"      data-row="${e.row}" title="Delete">&times;</button>
  </div>`;
}

/* Tasks belong to a project on the Projects tab, so the list is exactly that
   project's tasks - never a fallback to every task ever used. */
function refreshTaskList() {
  const proj = $('project').value.trim();
  const scoped = state.projectTasks[proj];
  const hint = $('taskHint');
  if (scoped && scoped.length) {
    fillList('task', scoped);
    hint.textContent = `- ${scoped.length} on this project`;
  } else {
    fillList('task', []);
    hint.textContent = proj ? '- no tasks set up for this project' : '';
  }
}

/* ---------------------------------------------------------------- api */

async function call(url, opts) {
  const res  = await fetch(url, opts);
  const data = await res.json().catch(() => ({ error: 'Bad response from server' }));
  if (!res.ok) throw new Error(data.error || `HTTP ${res.status}`);
  return data;
}

function showApp() {
  $('appShell').hidden = false;
  $('authView').hidden = true;
  $('currentUser').textContent = `Signed in as ${authUser?.name || authUser?.username || ''}`;
  $('currentUser').hidden = false;
  $('usersTab').hidden = authUser?.username?.toLowerCase() !== 'admin';
  if (authUser?.mustChangePassword) showPasswordView(true);
}

function showLogin() {
  $('authView').hidden = false;
  $('appShell').hidden = true;
  $('currentUser').hidden = true;
  $('usersTab').hidden = true;
  $('passwordView').hidden = true;
}

async function login(username, password) {
  const data = await call('/api/login', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ username, password })
  });
  authUser = data.user;
  showApp();
  load();
  loadYesterdayUpdate();
  loadTasks();
  loadHolidays();
  return data;
}

async function logout() {
  try { await call('/api/logout', { method: 'POST' }); } catch (_) {}
  authUser = null;
  showLogin();
}

async function initAuth() {
  try {
    const data = await call('/api/session');
    authUser = data.user;
    showApp();
    load();
    loadYesterdayUpdate();
    loadTasks();
  } catch {
    authUser = null;
    showLogin();
  }
}

function setUserMsg(text, kind) {
  const el = $('userMsg');
  el.textContent = text || '';
  el.className = 'msg' + (kind ? ' ' + kind : '');
  if (kind === 'ok') setTimeout(() => { if (el.textContent === text) el.textContent = ''; }, 4000);
}

function setPasswordMsg(text, kind) {
  const el = $('passwordMsg');
  el.textContent = text || '';
  el.className = 'msg' + (kind ? ' ' + kind : '');
}

function showPasswordView(required) {
  $('passwordView').hidden = false;
  $('passwordTitle').textContent = required ? 'Set your password' : 'Change password';
  $('passwordIntro').textContent = required
    ? 'Your administrator provided a temporary password. Choose a new password to continue.'
    : 'Choose a new password for your account.';
  $('currentPasswordField').hidden = required;
  $('currentPassword').required = !required;
  $('cancelPassword').hidden = required;
  $('passwordForm').dataset.required = required ? 'true' : 'false';
  setPasswordMsg('');
  $('currentPassword').value = '';
  $('newPassword').value = '';
  $('confirmPassword').value = '';
  (required ? $('newPassword') : $('currentPassword')).focus();
}

function hidePasswordView() {
  if (authUser?.mustChangePassword) return;
  $('passwordView').hidden = true;
  $('passwordForm').reset();
}

async function changePassword() {
  const required = $('passwordForm').dataset.required === 'true';
  const currentPassword = $('currentPassword').value;
  const newPassword = $('newPassword').value;
  const confirmPassword = $('confirmPassword').value;
  if (newPassword.length < 8) return setPasswordMsg('Password must be at least 8 characters.', 'err');
  if (newPassword !== confirmPassword) return setPasswordMsg('New passwords do not match.', 'err');
  try {
    setPasswordMsg('Saving password...');
    await call('/api/password/change', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ currentPassword, newPassword })
    });
    authUser.mustChangePassword = false;
    hidePasswordView();
    setPasswordMsg('Password saved.', 'ok');
  } catch (err) {
    setPasswordMsg(err.message, 'err');
  }
}

async function loadUsers() {
  try {
    const data = await call('/api/users');
    users = data.users || [];
    $('usersBody').innerHTML = users.map(user =>
      `<div class="user-row"><strong>${esc(user.username)}</strong><span>${esc(user.name)}</span></div>`
    ).join('') || '<p class="empty">No users found.</p>';
    $('resetUsername').innerHTML = users.map(user => `<option value="${esc(user.username)}">${esc(user.name)} (${esc(user.username)})</option>`).join('');
  } catch (err) {
    $('usersBody').innerHTML = `<p class="empty">${esc(err.message)}</p>`;
  }
}

async function createUser(payload) {
  try {
    setUserMsg('Creating user...');
    const data = await call('/api/users', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
    users = data.users || [];
    $('usersBody').innerHTML = users.map(user =>
      `<div class="user-row"><strong>${esc(user.username)}</strong><span>${esc(user.name)}</span></div>`
    ).join('');
    $('resetUsername').innerHTML = users.map(user => `<option value="${esc(user.username)}">${esc(user.name)} (${esc(user.username)})</option>`).join('');
    $('userForm').reset();
    setUserMsg('User created.', 'ok');
  } catch (err) {
    setUserMsg(err.message, 'err');
  }
}

async function resetUserPassword(username, password) {
  try {
    setUserMsg('Resetting password...');
    await call('/api/users/reset-password', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ username, password })
    });
    $('resetUserForm').reset();
    setUserMsg('Password reset. The user must choose a new password at next login.', 'ok');
  } catch (err) {
    setUserMsg(err.message, 'err');
  }
}

function apply(data) {
  state = { ...state, ...data };
  allEntries = null;          // entries changed, so the Projects view must refetch
  render();
}

async function load() {
  try { apply(await call('/api/bootstrap')); }
  catch (err) { setMsg(err.message, 'err'); $('entries').innerHTML = '<p class="empty">Could not reach the server.</p>'; }
}

async function loadYesterdayUpdate() {
  const d = new Date();
  d.setDate(d.getDate() - 1);
  const date = iso(d);
  $('updateDate').textContent = longDate(d);
  $('updateStatus').textContent = 'Loading yesterday\'s entries...';
  try {
    if (!dayWeek || date < dayWeek.weekStart || date > dayWeek.weekEnd) {
      dayWeek = await call(`/api/week?start=${date}`);
    }
    const entries = (dayWeek.entries || []).filter(e => e.date === date)
      .sort((a, b) => String(a.project || '').localeCompare(String(b.project || '')) ||
                      String(a.task || '').localeCompare(String(b.task || '')));
    if (entries.length) {
      const projects = new Map();
      for (const entry of entries) {
        const project = entry.project || 'General work';
        const notes = String(entry.notes || '').trim();
        if (!projects.has(project)) projects.set(project, []);
        if (notes && !projects.get(project).includes(notes)) projects.get(project).push(notes);
      }
      $('yesterdayUpdate').innerHTML = [...projects].map(([project, notes]) =>
        `<div class="project"><strong>${esc(project)}</strong><br>&bull; ${esc(notes.join('; ') || 'No notes entered')}</div>`
      ).join('');
      $('updateStatus').textContent = 'Generated from yesterday\'s entries. You can edit it before sending.';
    } else {
      $('yesterdayUpdate').innerHTML = '';
      $('updateStatus').textContent = 'No entries were found for yesterday.';
    }
  } catch (err) {
    $('updateStatus').textContent = err.message;
    $('yesterdayUpdate').textContent = '';
  }
}

async function copyYesterdayUpdate() {
  const text = $('yesterdayUpdate').innerText.trim();
  if (!text) return;
  try {
    await navigator.clipboard.writeText(text);
    $('btnCopyUpdate').textContent = 'Copied';
    setTimeout(() => $('btnCopyUpdate').textContent = 'Copy update', 1600);
  } catch {
    window.prompt('Copy the update below:', text);
  }
}

/* ---------------------------------------------------------------- edit */

function setEditMode(on, row) {
  editingRow = on ? row : null;
  $('editBanner').hidden = !on;
  if (on) $('editRow').textContent = row;
  $('btnAdd').textContent   = on ? 'Save changes' : 'Add entry';
  $('btnKeep').hidden       = on;
  document.querySelector('.form-card').classList.toggle('is-editing', on);
}

function startEdit(row) {
  const e = findEntry(row);
  if (!e) return;
  showView('entry');            // the form lives on the entry tab
  $('date').value  = e.date  || '';
  $('hours').value = (e.hours || 0).toFixed(2).replace(/\.00$/, '');
  $('notes').value = e.notes || '';
  $('ticket').value       = e.ticket       || '';
  $('incidentType').value = e.incidentType || '';
  CONTEXT.forEach(f => $(f).value = e[f] || '');
  $('hours').dispatchEvent(new Event('input'));
  refreshTaskList();
  setEditMode(true, row);
  render();
  flashForm();
  setMsg(`Editing row ${row}. Save to overwrite it, or Cancel to leave it alone.`);
  $('hours').focus();
}

function cancelEdit() {
  setEditMode(false);
  clearForm(false);
  render();
  setMsg('Edit cancelled - nothing was changed.');
}

function clearForm(keepContext) {
  $('hours').value  = '';
  $('notes').value  = '';
  $('ticket').value = '';
  $('hoursHint').textContent = '';
  if (!keepContext) {
    [...CONTEXT, 'incidentType'].forEach(f => $(f).value = '');
    refreshTaskList();
  }
}

function copyEntry(row) {
  const e = findEntry(row);
  if (!e) return;
  showView('entry');
  if (editingRow !== null) setEditMode(false);
  CONTEXT.forEach(f => $(f).value = e[f] || '');
  clearForm(true);
  $('notes').value = e.notes || '';    // set after clearForm, which blanks it
  $('incidentType').value = '';
  refreshTaskList();
  render();
  flashForm();
  $('hours').focus();
  setMsg('Copied - fill in the hours.', 'ok');
}

/* ---------------------------------------------------------------- submit */

async function submit(keepValues) {
  const hours = parseHours($('hours').value);
  if (hours === null)           return setMsg('Enter hours as 1.5, 1:30, 1h 30m or 90m.', 'err');
  if (hours <= 0 || hours > 24) return setMsg('Hours must be between 0 and 24.', 'err');
  if (!$('project').value.trim()) return setMsg('Project is required.', 'err');
  if (!$('budget').value.trim())  return setMsg('Budget is required.', 'err');

  const payload = {};
  FIELDS.forEach(f => payload[f] = $(f).value.trim());

  const editing = editingRow !== null;
  if (editing) payload.row = editingRow;

  const buttons = [$('btnAdd'), $('btnKeep')];
  buttons.forEach(b => b.disabled = true);
  setMsg(editing ? 'Updating Excel...' : 'Saving to Excel...');

  try {
    const data = await call(editing ? '/api/update' : '/api/entry', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });

    if (editing) setEditMode(false);
    apply(data);

    setMsg(editing
      ? `Row ${data.result.row} updated to ${fmtHM(data.result.hours)}.`
      : `Added ${fmtHM(data.result.hours)} to row ${data.result.row}.`, 'ok');

    clearForm(keepValues && !editing);
    if (editing) { clearForm(false); render(); }
    $((keepValues && !editing) ? 'hours' : 'project').focus();
    refreshOpenViews();
  } catch (err) {
    setMsg(err.message, 'err');
  } finally {
    buttons.forEach(b => b.disabled = false);
  }
}

async function del(row) {
  const e = findEntry(row);
  const what = e ? `${e.project} - ${(e.hours || 0).toFixed(2)}h on ${e.dateText}` : `row ${row}`;
  if (!confirm(`Delete this entry?\n\n${what}`)) return;
  try {
    if (editingRow === row) setEditMode(false);
    apply(await call('/api/delete', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ row })
    }));
    setMsg('Entry deleted.', 'ok');
    refreshOpenViews();
  } catch (err) { setMsg(err.message, 'err'); }
}

/* Deleting shifts every later worksheet row up by one, so any cached view has
   stale row numbers until it refetches. */
function refreshOpenViews() {
  if (reportData) loadReport(reportWeek);
  if (dayWeek)    loadDay(dayDate, true);
  loadYesterdayUpdate();
}

/* ---------------------------------------------------------------- report */

async function loadReport(start) {
  try {
    const q = start ? `?start=${start}` : '';
    reportData = await call('/api/week' + q);
    reportWeek = reportData.weekStart;
    renderReport();
  } catch (err) {
    $('reportBody').innerHTML = `<p class="empty">${esc(err.message)}</p>`;
  }
}

/* Sum `hours` into a Map keyed by whatever `key` returns. */
function tally(entries, key) {
  const m = new Map();
  for (const e of entries) {
    const k = key(e);
    if (k === null || k === undefined) continue;
    m.set(k, (m.get(k) || 0) + (e.hours || 0));
  }
  return [...m.entries()].sort((a, b) => b[1] - a[1]);
}

function barRows(pairs, total, cls) {
  if (!pairs.length) return '<p class="empty small">Nothing recorded.</p>';
  return pairs.map(([label, h]) => {
    const pct = total > 0 ? (h / total) * 100 : 0;
    return `<div class="brow">
      <div class="brow-label" title="${esc(label)}">${esc(label)}</div>
      <div class="brow-track"><div class="brow-fill ${cls}" style="width:${pct.toFixed(1)}%"></div></div>
      <div class="brow-val">${h.toFixed(2)}</div>
      <div class="brow-pct">${pct.toFixed(0)}%</div>
    </div>`;
  }).join('');
}

/* Per-project totals with the billable split shown inside each bar. */
function projectRows(entries, total) {
  const m = new Map();
  for (const e of entries) {
    const k = e.project || '(no project)';
    const p = m.get(k) || { hours: 0, billable: 0, tasks: new Set(), count: 0 };
    p.hours += e.hours || 0;
    if (e.billable === 'Yes') p.billable += e.hours || 0;
    if (e.task) p.tasks.add(e.task);
    p.count++;
    m.set(k, p);
  }
  const rows = [...m.entries()].sort((a, b) => b[1].hours - a[1].hours);
  if (!rows.length) return '<p class="empty small">Nothing recorded.</p>';

  return rows.map(([label, p]) => {
    const pct  = total > 0 ? (p.hours / total) * 100 : 0;
    const bPct = p.hours > 0 ? (p.billable / p.hours) * 100 : 0;
    return `<div class="brow">
      <div class="brow-label" title="${esc(label)}">${esc(label)}
        <span class="brow-sub">${p.count} entr${p.count === 1 ? 'y' : 'ies'}${p.tasks.size ? ` &middot; ${p.tasks.size} task${p.tasks.size === 1 ? '' : 's'}` : ''}</span>
      </div>
      <div class="brow-track" title="Billable ${p.billable.toFixed(2)}h of ${p.hours.toFixed(2)}h">
        <div class="brow-fill f-proj" style="width:${pct.toFixed(1)}%">
          <div class="brow-bill" style="width:${bPct.toFixed(1)}%"></div>
        </div>
      </div>
      <div class="brow-val">${p.hours.toFixed(2)}</div>
      <div class="brow-pct">${pct.toFixed(0)}%</div>
    </div>`;
  }).join('');
}

/* ------------------------------------------- pivot: week/project/task/date */

let pivotRows = [];   // flattened rows behind the table, indexed by data-row

// Joins the ancestor labels into a row path. A unit separator can't turn up in
// a project name, so "is a descendant of" stays a plain startsWith test. The
// path itself never reaches the DOM - rows are keyed by their pivotRows index.
const PSEP = String.fromCharCode(31);   // unit separator

/* Week -> Project -> Task -> entry rows, mirroring TimesheetPivot in the
   workbook. Excel already writes "2026-W34" into the Week Number column;
   fall back to computing it when that column comes back empty. */
function buildPivot(entries) {
  const weeks = new Map();
  for (const e of entries) {
    let wk = (e.weekNumber || '').trim();
    if (!wk && e.date) {
      const { year, week } = isoWeek(parseIso(e.date));
      wk = `${year}-W${String(week).padStart(2, '0')}`;
    }
    if (!wk) wk = '(no week)';

    const pn = e.project || '(no project)';
    const tn = e.task    || '(no task)';
    const h  = e.hours   || 0;

    if (!weeks.has(wk)) weeks.set(wk, { hours: 0, projects: new Map() });
    const w = weeks.get(wk);
    w.hours += h;

    if (!w.projects.has(pn)) w.projects.set(pn, { hours: 0, tasks: new Map() });
    const p = w.projects.get(pn);
    p.hours += h;

    if (!p.tasks.has(tn)) p.tasks.set(tn, { hours: 0, rows: [] });
    const t = p.tasks.get(tn);
    t.hours += h;
    t.rows.push(e);
  }
  return weeks;
}

const byName = (a, b) => String(a[0]).localeCompare(String(b[0]));

/* Depth-first walk into the flat row list the table renders from. Groups carry
   their own rolled-up hours, so a collapsed group still shows its total. */
function flattenPivot(tree) {
  const out = [];
  for (const [wk, w] of [...tree].sort(byName)) {
    out.push({ kind: 'group', level: 0, path: wk, label: wk, notes: '', hours: w.hours });

    for (const [pn, p] of [...w.projects].sort(byName)) {
      const pp = wk + PSEP + pn;
      out.push({ kind: 'group', level: 1, path: pp, label: pn, notes: '', hours: p.hours });

      for (const [tn, t] of [...p.tasks].sort(byName)) {
        const tp = pp + PSEP + tn;
        out.push({ kind: 'group', level: 2, path: tp, label: tn, notes: '', hours: t.hours });

        // Several entries on the same task and day become one line: hours added
        // up and their notes clubbed together, so a day is never repeated.
        const days = new Map();
        for (const e of t.rows) {
          const key = e.date || e.dateText || '';
          if (!days.has(key)) days.set(key, { label: e.dateText || e.date || '', notes: [], hours: 0 });
          const d = days.get(key);
          d.hours += e.hours || 0;
          const n = (e.notes || '').trim();
          if (n && !d.notes.includes(n)) d.notes.push(n);
        }

        [...days]
          .sort((a, b) => String(a[0]).localeCompare(String(b[0])))
          .forEach(([, d], i) => out.push({
            kind: 'leaf', level: 3, path: tp + PSEP + i,
            label: d.label, notes: d.notes.join('\n'), hours: d.hours
          }));
      }
    }
  }
  return out;
}

function pivotHtml(entries) {
  pivotRows = flattenPivot(buildPivot(entries));
  if (!pivotRows.length) return '<p class="empty small">Nothing recorded.</p>';
  const total = entries.reduce((s, e) => s + (e.hours || 0), 0);

  const body = pivotRows.map((r, i) => {
    const group = r.kind === 'group';
    const label = group ? `<span class="pvt-twisty">&#9662;</span>${esc(r.label)}` : esc(r.label);
    return `<tr class="pvt-l${r.level} ${group ? 'pvt-group' : 'pvt-leaf'}"
                data-row="${i}"${group ? ' data-group="1"' : ''}>
      <td class="pvt-lbl">${label}</td>
      <td class="pvt-notes">${esc(r.notes).replace(/\n/g, '<br>')}</td>
      <td class="pvt-hrs">${r.hours.toFixed(2)}</td>
    </tr>`;
  }).join('');

  return `<table class="pvt">
    <thead><tr>
      <th>Week / Project / Task / Date</th><th>Notes</th><th class="pvt-hrs">Hours</th>
    </tr></thead>
    <tbody>${body}</tbody>
    <tfoot><tr class="pvt-grand">
      <td>Total</td><td></td><td class="pvt-hrs">${total.toFixed(2)}</td>
    </tr></tfoot>
  </table>`;
}

const pathOf = tr => (pivotRows[+tr.dataset.row] || {}).path;

// A row is hidden when any strict ancestor is collapsed, so nesting works
// without having to track which level did the collapsing.
function applyPivotCollapse() {
  const collapsed = [...reportCollapsed];
  document.querySelectorAll('#reportBody tr[data-row]').forEach(tr => {
    const path = pathOf(tr);
    tr.hidden = collapsed.some(c => path.startsWith(c + PSEP));
    if (tr.dataset.group === '1') tr.classList.toggle('is-collapsed', reportCollapsed.has(path));
  });
}

function togglePivot(path) {
  if (reportCollapsed.has(path)) reportCollapsed.delete(path);
  else                           reportCollapsed.add(path);
  applyPivotCollapse();
}

/* ---------------------------------------------------------------- day view */

/* The week endpoint snaps whatever date it is given back to that week's Monday,
   so a single fetch covers seven days of Prev/Next without touching Excel
   again. `force` is for after a write, when the cache is stale. */
async function loadDay(dateStr, force) {
  dayDate = dateStr || iso(new Date());
  const covered = dayWeek && !force &&
                  dayDate >= dayWeek.weekStart && dayDate <= dayWeek.weekEnd;
  try {
    if (!covered) dayWeek = await call(`/api/week?start=${dayDate}`);
    renderDay();
  } catch (err) {
    $('dayBody').innerHTML = `<p class="empty">${esc(err.message)}</p>`;
  }
}

const longDate = d =>
  d.toLocaleDateString(undefined, { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });

/* "Today" / "Yesterday" / "Tomorrow" when it applies, always the ISO week so
   the label ties back to the weekly report. */
function dayCaption(d) {
  const shift = n => { const x = new Date(); x.setDate(x.getDate() + n); return iso(x); };
  const rel = { [shift(0)]: 'Today', [shift(-1)]: 'Yesterday', [shift(1)]: 'Tomorrow' }[iso(d)];
  const { year, week } = isoWeek(d);
  return `${rel ? rel + ' - ' : ''}${year}-W${String(week).padStart(2, '0')}`;
}

const entriesOn = date => (dayWeek?.entries || [])
  .filter(e => e.date === date)
  .sort((a, b) => String(a.project || '').localeCompare(String(b.project || '')) ||
                  String(a.task    || '').localeCompare(String(b.task    || '')));

function renderDay() {
  const d       = parseIso(dayDate);
  const entries = entriesOn(dayDate);

  $('dyName').textContent = longDate(d);
  $('dyRel').textContent  = dayCaption(d);
  $('dayPick').value      = dayDate;
  $('printDayTitle').textContent = `Daily timesheet - ${longDate(d)}`;

  if (!entries.length) {
    $('dayBody').innerHTML = '<p class="empty">No entries logged on this day.</p>';
    return;
  }

  const total    = entries.reduce((s, e) => s + (e.hours || 0), 0);
  const billable = entries.filter(e => e.billable === 'Yes').reduce((s, e) => s + (e.hours || 0), 0);

  /* Same project can appear on several rows; roll them up so the day reads as
     "where did the time go" rather than a flat list. */
  const projects = tally(entries, e => e.project || '(no project)');

  $('dayBody').innerHTML = `
    <div class="sumstrip sumstrip-4">
      <div class="sumcard s-total"><div class="sum-val">${total.toFixed(2)}</div><div class="sum-lbl">Total hours</div></div>
      <div class="sumcard s-bill"><div class="sum-val">${billable.toFixed(2)}</div><div class="sum-lbl">Billable</div></div>
      <div class="sumcard s-nonbill"><div class="sum-val">${(total-billable).toFixed(2)}</div><div class="sum-lbl">Non-billable</div></div>
      <div class="sumcard s-count"><div class="sum-val">${entries.length}</div><div class="sum-lbl">Entries</div></div>
    </div>

    <div class="rsection">
      <h3>Where the time went</h3>
      ${barRows(projects, total, 'f-day')}
    </div>

    <div class="rsection">
      <h3>Entries</h3>
      <div class="dayentries">${entries.map(entryHtml).join('')}</div>
    </div>
  `;

  /* Rebuilt on every render, so the listeners go back on each time. */
  const box = $('dayBody');
  box.querySelectorAll('.del').forEach(b =>
    b.addEventListener('click', () => del(+b.dataset.row)));
  box.querySelectorAll('.copy-btn').forEach(b =>
    b.addEventListener('click', () => copyEntry(+b.dataset.row)));
  box.querySelectorAll('.edit-btn').forEach(b =>
    b.addEventListener('click', () => startEdit(+b.dataset.row)));
}

/* Plain text for a standup or a daily update. */
function dayAsText() {
  const d = parseIso(dayDate);
  const entries = entriesOn(dayDate);
  const total    = entries.reduce((s, e) => s + (e.hours || 0), 0);
  const billable = entries.filter(e => e.billable === 'Yes').reduce((s, e) => s + (e.hours || 0), 0);

  const L = [longDate(d), '='.repeat(60)];
  if (!entries.length) { L.push('No entries logged on this day.'); return L.join('\n'); }

  L.push(`Total ${total.toFixed(2)} h   Billable ${billable.toFixed(2)} h   Non-billable ${(total-billable).toFixed(2)} h`);
  L.push('');
  for (const e of entries) {
    L.push(`  ${e.project || '(no project)'}`.padEnd(58) + (e.hours || 0).toFixed(2).padStart(8));
    if (e.task)  L.push(`    ${e.task}`);
    if (e.notes) L.push(`    ${e.notes}`);
  }
  return L.join('\n');
}

function shiftDay(days) {
  const d = parseIso(dayDate);
  d.setDate(d.getDate() + days);
  loadDay(iso(d));
}

function renderReport() {
  const entries = reportData.entries || [];
  const mon     = parseIso(reportData.weekStart);
  const sun     = parseIso(reportData.weekEnd);
  const { year, week } = isoWeek(mon);

  $('wkIso').textContent   = `${year}-W${String(week).padStart(2, '0')}`;
  $('wkRange').textContent = `${prettyDate(mon)} - ${prettyDate(sun)} ${sun.getFullYear()}`;
  $('printTitle').textContent =
    `Weekly timesheet - ${year}-W${String(week).padStart(2,'0')} (${prettyDate(mon)} - ${prettyDate(sun)} ${sun.getFullYear()})`;

  if (!entries.length) {
    $('reportBody').innerHTML = '<p class="empty">No entries logged in this week.</p>';
    return;
  }

  const total    = entries.reduce((s, e) => s + (e.hours || 0), 0);
  const billable = entries.filter(e => e.billable === 'Yes').reduce((s, e) => s + (e.hours || 0), 0);
  const nonBill  = total - billable;
  const daysLogged = new Set(entries.map(e => e.date)).size;

  /* ---- day by day, always all 7 days so gaps are visible ---- */
  const byDate = new Map();
  for (const e of entries) byDate.set(e.date, (byDate.get(e.date) || 0) + (e.hours || 0));
  const dayMax = Math.max(...DAYS.map((_, i) => {
    const d = new Date(mon); d.setDate(mon.getDate() + i);
    return byDate.get(iso(d)) || 0;
  }), 1);

  const dayHtml = DAYS.map((name, i) => {
    const d = new Date(mon); d.setDate(mon.getDate() + i);
    const h = byDate.get(iso(d)) || 0;
    const weekend = i >= 5;
    return `<div class="dayrow ${h === 0 ? 'is-zero' : ''} ${weekend ? 'is-weekend' : ''}">
      <div class="dayrow-name">${name} <span class="dayrow-date">${prettyDate(d)}</span></div>
      <div class="brow-track"><div class="brow-fill f-day" style="width:${((h/dayMax)*100).toFixed(1)}%"></div></div>
      <div class="brow-val">${h ? h.toFixed(2) : '-'}</div>
    </div>`;
  }).join('');

  /* ---- non-billable, grouped by reason ---- */
  const reasons = tally(entries.filter(e => e.billable !== 'Yes'),
                        e => e.nonBillableReason || '(no reason given)');

  const pctB = total > 0 ? (billable / total) * 100 : 0;

  $('reportBody').innerHTML = `
    <div class="sumstrip">
      <div class="sumcard s-total"><div class="sum-val">${total.toFixed(2)}</div><div class="sum-lbl">Total hours</div></div>
      <div class="sumcard s-bill"><div class="sum-val">${billable.toFixed(2)}</div><div class="sum-lbl">Billable</div></div>
      <div class="sumcard s-nonbill"><div class="sum-val">${nonBill.toFixed(2)}</div><div class="sum-lbl">Non-billable</div></div>
      <div class="sumcard s-days"><div class="sum-val">${daysLogged}</div><div class="sum-lbl">Days logged</div></div>
      <div class="sumcard s-count"><div class="sum-val">${entries.length}</div><div class="sum-lbl">Entries</div></div>
    </div>

    <div class="rsection">
      <h3>Billable split</h3>
      <div class="splitbar">
        <div class="split-b"  style="width:${pctB.toFixed(1)}%"  title="Billable ${billable.toFixed(2)}h"></div>
        <div class="split-nb" style="width:${(100-pctB).toFixed(1)}%" title="Non-billable ${nonBill.toFixed(2)}h"></div>
      </div>
      <div class="splitlegend">
        <span><i class="sw sw-b"></i> Billable ${billable.toFixed(2)} h (${pctB.toFixed(0)}%)</span>
        <span><i class="sw sw-nb"></i> Non-billable ${nonBill.toFixed(2)} h (${(100-pctB).toFixed(0)}%)</span>
      </div>
      ${reasons.length ? `<h4>Non-billable by reason</h4>${barRows(reasons, nonBill, 'f-reason')}` : ''}
    </div>

    <div class="rsection">
      <h3>Day by day</h3>
      <div class="daytable">${dayHtml}</div>
    </div>

    <div class="rsection">
      <h3>Hours by week / project / task / date
        <span class="pvt-tools no-print">
          <button type="button" id="pvtExpand"   class="btn btn-tiny">Expand all</button>
          <button type="button" id="pvtCollapse" class="btn btn-tiny">Collapse all</button>
        </span>
      </h3>
      ${pivotHtml(entries)}
    </div>

    <div class="rsection">
      <h3>By project</h3>
      ${projectRows(entries, total)}
    </div>

    <div class="rsection">
      <h3>By category</h3>
      ${barRows(tally(entries, e => e.category || '(no category)'), total, 'f-cat')}
    </div>
  `;

  /* Rebuilt on every render, so the listeners have to go back on each time. */
  const body = $('reportBody');
  body.querySelectorAll('tr[data-group]').forEach(tr =>
    tr.addEventListener('click', () => togglePivot(pathOf(tr))));
  $('pvtExpand').addEventListener('click', () => {
    reportCollapsed.clear();
    applyPivotCollapse();
  });
  $('pvtCollapse').addEventListener('click', () => {
    body.querySelectorAll('tr[data-group]').forEach(tr => reportCollapsed.add(pathOf(tr)));
    applyPivotCollapse();
  });
  applyPivotCollapse();
}

/* Plain-text version for pasting into email or a status update. */
function reportAsText() {
  const entries = reportData?.entries || [];
  const mon = parseIso(reportData.weekStart), sun = parseIso(reportData.weekEnd);
  const { year, week } = isoWeek(mon);
  const total    = entries.reduce((s, e) => s + (e.hours || 0), 0);
  const billable = entries.filter(e => e.billable === 'Yes').reduce((s, e) => s + (e.hours || 0), 0);

  const L = [];
  L.push(`Weekly timesheet - ${year}-W${String(week).padStart(2,'0')} (${prettyDate(mon)} - ${prettyDate(sun)} ${sun.getFullYear()})`);
  L.push('='.repeat(60));
  L.push(`Total ${total.toFixed(2)} h   Billable ${billable.toFixed(2)} h   Non-billable ${(total-billable).toFixed(2)} h`);
  L.push('');

  // Same tree the table renders from, so the two can't drift apart.
  L.push('BY WEEK / PROJECT / TASK / DATE');
  for (const r of flattenPivot(buildPivot(entries))) {
    const pad   = '  '.repeat(r.level + 1);
    const lines = r.notes ? r.notes.split('\n') : [];
    const first = lines.length ? `${r.label}  ${lines[0]}` : r.label;
    L.push((pad + first).padEnd(70) + r.hours.toFixed(2).padStart(8));
    // Extra notes for the same day sit under the date, aligned with the first.
    for (const n of lines.slice(1)) L.push(pad + ' '.repeat(r.label.length + 2) + n);
  }
  L.push('');

  L.push('BY DAY');
  const byDate = new Map();
  for (const e of entries) byDate.set(e.date, (byDate.get(e.date) || 0) + (e.hours || 0));
  DAYS.forEach((name, i) => {
    const d = new Date(mon); d.setDate(mon.getDate() + i);
    const h = byDate.get(iso(d)) || 0;
    L.push(`  ${name} ${prettyDate(d)}   ${h ? h.toFixed(2) : '-'}`);
  });
  L.push('');

  L.push('BY PROJECT');
  for (const [p, h] of tally(entries, e => e.project || '(no project)')) L.push(`  ${p}  ${h.toFixed(2)}`);
  L.push('');

  L.push('BY CATEGORY');
  for (const [c, h] of tally(entries, e => e.category || '(no category)')) L.push(`  ${c}  ${h.toFixed(2)}`);

  const reasons = tally(entries.filter(e => e.billable !== 'Yes'), e => e.nonBillableReason || '(no reason given)');
  if (reasons.length) {
    L.push('');
    L.push('NON-BILLABLE BY REASON');
    for (const [r, h] of reasons) L.push(`  ${r}  ${h.toFixed(2)}`);
  }
  return L.join('\n');
}

function shiftWeek(days) {
  const d = parseIso(reportWeek);
  d.setDate(d.getDate() + days);
  loadReport(iso(d));
}

/* ------------------------------------------------------- projects tab */

/* The Projects sheet is the master list: what exists, what state it is in, and
   which tasks belong to it. Entry reads its dropdowns straight off this. */
const PROJ_STATUSES = ['New', 'In-Progress', 'Hold', 'Complete'];

function setPadmMsg(text, kind) {
  const el = $('padmMsg');
  el.textContent = text || '';
  el.className = 'msg' + (kind ? ' ' + kind : '');
  if (kind === 'ok') setTimeout(() => { if (el.textContent === text) el.textContent = ''; }, 3000);
}

async function projectAction(payload, okMsg) {
  try {
    setPadmMsg('Saving to Excel...');
    apply(await call('/api/project', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    }));
    setPadmMsg(okMsg, 'ok');
  } catch (err) {
    setPadmMsg(err.message, 'err');
  }
}

function renderProjectAdmin() {
  const list = state.projects || [];
  const active = list.filter(p => p.status === 'In-Progress').length;
  $('padmStats').textContent = list.length
    ? `${list.length} project${list.length === 1 ? '' : 's'} - ${active} in progress`
    : '';

  if (!list.length) {
    $('padmBody').innerHTML = '<p class="empty">No projects yet. Add one above.</p>';
    return;
  }

  $('padmBody').innerHTML = list.map(p => `
    <div class="padm-row is-${statusKey(p.status)}">
      <div class="padm-head">
        <span class="padm-name">${esc(p.name)}</span>
        <span class="padm-count">${p.tasks.length} task${p.tasks.length === 1 ? '' : 's'}</span>
        <select class="padm-pick" data-proj="${esc(p.name)}" aria-label="Status for ${esc(p.name)}">
          ${PROJ_STATUSES.map(s => `<option value="${s}"${s === p.status ? ' selected' : ''}>${s}</option>`).join('')}
        </select>
        <button type="button" class="padm-del" data-proj="${esc(p.name)}" title="Delete project">&times;</button>
      </div>
      <div class="padm-tasks">
        ${p.tasks.length
          ? p.tasks.map(t => `<span class="tag padm-task">${esc(t)}<button type="button" class="padm-task-del" data-proj="${esc(p.name)}" data-task="${esc(t)}" title="Remove task">&times;</button></span>`).join('')
          : '<span class="empty small">No tasks yet.</span>'}
      </div>
      <form class="padm-taskform" data-proj="${esc(p.name)}">
        <input type="text" class="padm-taskinput" placeholder="Add a task to this project..." maxlength="120">
        <button type="submit" class="btn btn-tiny">Add task</button>
      </form>
    </div>`).join('');

  const box = $('padmBody');
  box.querySelectorAll('.padm-pick').forEach(sel =>
    sel.addEventListener('change', () =>
      projectAction({ action: 'setStatus', name: sel.dataset.proj, status: sel.value },
                    `${sel.dataset.proj} is now ${sel.value}.`)));

  box.querySelectorAll('.padm-del').forEach(btn =>
    btn.addEventListener('click', () => {
      const name = btn.dataset.proj;
      if (!confirm(`Delete "${name}" and its tasks?\n\nTimesheet entries already booked against it are not touched.`)) return;
      projectAction({ action: 'removeProject', name }, `${name} deleted.`);
    }));

  box.querySelectorAll('.padm-task-del').forEach(btn =>
    btn.addEventListener('click', () =>
      projectAction({ action: 'removeTask', name: btn.dataset.proj, task: btn.dataset.task },
                    `Removed "${btn.dataset.task}".`)));

  box.querySelectorAll('.padm-taskform').forEach(form =>
    form.addEventListener('submit', e => {
      e.preventDefault();
      const input = form.querySelector('.padm-taskinput');
      const task = input.value.trim();
      if (!task) return;
      input.value = '';
      projectAction({ action: 'addTask', name: form.dataset.proj, task }, `Added "${task}".`);
    }));
}

/* ------------------------------------------------------- team activities */

/* WFH is booked in whole Mon-Fri weeks, so a week is the unit throughout: one
   row per person per week, counted straight against the yearly allowance. */
let team = { allowance: 4, activities: [], people: [], entries: [] };
let teamYear = new Date().getFullYear();

/* "2026-W11" -> the Monday of that ISO week, which is what the server stores. */
function mondayOfWeekValue(value) {
  const m = /^(\d{4})-W(\d{1,2})$/.exec(String(value || '').trim());
  if (!m) return null;
  const monday = mondayOf(new Date(+m[1], 0, 4));     // Jan 4 is always in week 1
  monday.setDate(monday.getDate() + (+m[2] - 1) * 7);
  return iso(monday);
}

function setTeamMsg(text, kind) {
  const el = $('teamMsg');
  el.textContent = text || '';
  el.className = 'msg' + (kind ? ' ' + kind : '');
  if (kind === 'ok') setTimeout(() => { if (el.textContent === text) el.textContent = ''; }, 4000);
}

async function loadTeam() {
  try {
    team = await call('/api/team');
    renderTeam();
  } catch (err) {
    $('teamBody').innerHTML = `<p class="empty">${esc(err.message)}</p>`;
  }
}

async function teamAction(payload, okMsg) {
  try {
    setTeamMsg('Saving to Excel...');
    const data = await call('/api/team', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
    team = data;
    renderTeam();
    const warn = data.result && data.result.warning;
    setTeamMsg(warn || okMsg, warn ? 'warn' : 'ok');
  } catch (err) {
    setTeamMsg(err.message, 'err');
  }
}

function renderTeam() {
  const entries = team.entries || [];
  const allowance = team.allowance || 4;

  // Year list comes from what is actually logged, plus this year so it is never empty.
  const years = [...new Set([...entries.map(e => +e.week.slice(0, 4)), new Date().getFullYear()])]
                  .sort((a, b) => b - a);
  if (!years.includes(teamYear)) teamYear = years[0];
  $('teamYear').innerHTML = years.map(y => `<option value="${y}"${y === teamYear ? ' selected' : ''}>${y}</option>`).join('');

  const scoped = entries.filter(e => +e.week.slice(0, 4) === teamYear);
  const used = {};
  for (const e of scoped) used[e.person] = (used[e.person] || 0) + 1;

  const active = (team.people || []).filter(p => p.status === 'Active');
  $('teamPerson').innerHTML = '<option value="">Select a person...</option>' +
    active.map(p => `<option value="${esc(p.name)}">${esc(p.name)}</option>`).join('');
  $('teamActivity').innerHTML = (team.activities || []).map(a => `<option value="${esc(a)}">${esc(a)}</option>`).join('');

  $('teamStats').textContent = `${scoped.length} week${scoped.length === 1 ? '' : 's'} logged in ${teamYear}`;

  // Anyone with history this year stays on the grid even once they go Inactive.
  const names = [...new Set([...active.map(p => p.name), ...scoped.map(e => e.person)])];
  if (!names.length) {
    $('teamBody').innerHTML = '<p class="empty">No one on the team yet. Add someone under Manage people.</p>';
    renderTeamPeople();
    return;
  }

  $('teamBody').innerHTML = `<div class="tm-grid">${names.map(name => {
    const n = used[name] || 0;
    const over = n > allowance;
    const pct = Math.min(n / allowance, 1) * 100;
    const weeks = scoped.filter(e => e.person === name).sort((a, b) => a.weekStart.localeCompare(b.weekStart));
    return `<div class="tm-row${over ? ' is-over' : ''}">
      <div class="tm-name">${esc(name)}</div>
      <div class="tm-track"><div class="tm-fill" style="width:${pct.toFixed(0)}%"></div></div>
      <div class="tm-count">${n} of ${allowance}</div>
      <div class="tm-left">${over ? `${n - allowance} over` : `${allowance - n} left`}</div>
      <div class="tm-weeks">${weeks.length
        ? weeks.map(w => `<span class="tag tm-week" title="${esc(w.weekStart)} to ${esc(w.weekEnd)}${w.notes ? ' - ' + esc(w.notes) : ''}">${esc(w.week)}<button type="button" class="tm-del" data-row="${w.row}" title="Remove">&times;</button></span>`).join('')
        : '<span class="empty small">none booked</span>'}</div>
    </div>`;
  }).join('')}</div>`;

  $('teamBody').querySelectorAll('.tm-del').forEach(btn =>
    btn.addEventListener('click', () => {
      const e = (team.entries || []).find(x => x.row === +btn.dataset.row);
      if (!e) return;
      if (!confirm(`Remove ${e.activity} ${e.week} for ${e.person}?`)) return;
      teamAction({ action: 'removeWeek', row: e.row }, `Removed ${e.week} for ${e.person}.`);
    }));

  renderTeamPeople();
}

function renderTeamPeople() {
  const people = team.people || [];
  $('teamPeopleBody').innerHTML = people.length
    ? people.map(p => `<div class="tm-person">
        <span class="tm-person-name">${esc(p.name)}</span>
        <select class="padm-pick" data-person="${esc(p.name)}" aria-label="Status for ${esc(p.name)}">
          <option value="Active"${p.status === 'Active' ? ' selected' : ''}>Active</option>
          <option value="Inactive"${p.status === 'Inactive' ? ' selected' : ''}>Inactive</option>
        </select>
        <button type="button" class="padm-del" data-person="${esc(p.name)}" title="Remove">&times;</button>
      </div>`).join('')
    : '<p class="empty small">Nobody yet.</p>';

  $('teamPeopleBody').querySelectorAll('.padm-pick').forEach(sel =>
    sel.addEventListener('change', () =>
      teamAction({ action: 'setPersonStatus', person: sel.dataset.person, status: sel.value },
                 `${sel.dataset.person} is now ${sel.value}.`)));

  $('teamPeopleBody').querySelectorAll('.padm-del').forEach(btn =>
    btn.addEventListener('click', () => {
      if (!confirm(`Remove ${btn.dataset.person} from the team?`)) return;
      teamAction({ action: 'removePerson', person: btn.dataset.person }, `${btn.dataset.person} removed.`);
    }));
}

/* ------------------------------------------------- effort by project */

let allEntries = null;        // every row, fetched once per visit
let projOpen = new Set();     // expanded project names

/* Returns [fromIso, toIso]; either end may be null for open-ended. */
function projRange() {
  const sel = $('pjRangeSel').value;
  const now = new Date();
  const y = now.getFullYear(), m = now.getMonth();
  if (sel === 'all')     return [null, null];
  if (sel === 'ytd')     return [iso(new Date(y, 0, 1)), null];
  if (sel === 'quarter') return [iso(new Date(y, Math.floor(m / 3) * 3, 1)), null];
  if (sel === 'month')   return [iso(new Date(y, m, 1)), null];
  if (sel === '30' || sel === '90') {
    const d = new Date(now); d.setDate(d.getDate() - (+sel - 1));
    return [iso(d), null];
  }
  return [$('pjFrom').value || null, $('pjTo').value || null];
}

async function loadProjects() {
  try {
    if (!allEntries) {
      $('projBody').innerHTML = '<p class="empty">Loading...</p>';
      const data = await call('/api/all');
      allEntries = data.entries || [];
    }
    renderProjects();
  } catch (err) {
    $('projBody').innerHTML = `<p class="empty">${esc(err.message)}</p>`;
  }
}

function projScoped() {
  const [from, to] = projRange();
  return (allEntries || []).filter(e =>
    e.date && (!from || e.date >= from) && (!to || e.date <= to));
}

/* project -> { hours, billable, entries, tasks: Map(task -> {hours,count}) } */
function groupByProject(entries) {
  const m = new Map();
  for (const e of entries) {
    const key = e.project || '(no project)';
    let p = m.get(key);
    if (!p) { p = { hours: 0, billable: 0, count: 0, first: e.date, last: e.date, tasks: new Map() }; m.set(key, p); }
    const h = e.hours || 0;
    p.hours += h;
    if (e.billable === 'Yes') p.billable += h;
    p.count++;
    if (e.date < p.first) p.first = e.date;
    if (e.date > p.last)  p.last  = e.date;

    const tk = e.task || '(no task)';
    const t = p.tasks.get(tk) || { hours: 0, count: 0 };
    t.hours += h; t.count++;
    p.tasks.set(tk, t);
  }
  return [...m.entries()].sort((a, b) => b[1].hours - a[1].hours);
}

function renderProjects() {
  const entries = projScoped();
  const [from, to] = projRange();
  const rows = groupByProject(entries);
  const total = entries.reduce((s, e) => s + (e.hours || 0), 0);
  const billable = entries.filter(e => e.billable === 'Yes').reduce((s, e) => s + (e.hours || 0), 0);
  const days = new Set(entries.map(e => e.date)).size;

  const dates = entries.map(e => e.date).sort();
  $('pjRange').textContent = dates.length
    ? `${prettyDate(parseIso(dates[0]))} ${parseIso(dates[0]).getFullYear()} - ${prettyDate(parseIso(dates[dates.length-1]))} ${parseIso(dates[dates.length-1]).getFullYear()}`
    : (from || to ? 'No entries in this range' : 'No entries');

  if (!rows.length) {
    $('projBody').innerHTML = '<p class="empty">Nothing recorded in this range.</p>';
    return;
  }

  const max = rows[0][1].hours || 1;

  $('projBody').innerHTML = `
    <div class="sumstrip">
      <div class="sumcard s-total"><div class="sum-val">${total.toFixed(2)}</div><div class="sum-lbl">Total hours</div></div>
      <div class="sumcard s-bill"><div class="sum-val">${billable.toFixed(2)}</div><div class="sum-lbl">Billable</div></div>
      <div class="sumcard s-nonbill"><div class="sum-val">${(total-billable).toFixed(2)}</div><div class="sum-lbl">Non-billable</div></div>
      <div class="sumcard s-days"><div class="sum-val">${rows.length}</div><div class="sum-lbl">Projects</div></div>
      <div class="sumcard s-count"><div class="sum-val">${days}</div><div class="sum-lbl">Days logged</div></div>
    </div>

    <div class="projlist">
      ${rows.map(([name, p]) => projCardHtml(name, p, total, max)).join('')}
    </div>`;

  $('projBody').querySelectorAll('.pj-head').forEach(h =>
    h.addEventListener('click', () => {
      const n = h.dataset.proj;
      projOpen.has(n) ? projOpen.delete(n) : projOpen.add(n);
      renderProjects();
    }));
}

function projCardHtml(name, p, total, max) {
  const open = projOpen.has(name);
  const share = total > 0 ? (p.hours / total) * 100 : 0;
  const bPct  = p.hours > 0 ? (p.billable / p.hours) * 100 : 0;
  const tasks = [...p.tasks.entries()].sort((a, b) => b[1].hours - a[1].hours);

  return `<div class="pjrow${open ? ' is-open' : ''}">
    <div class="pj-head" data-proj="${esc(name)}" title="Click to show tasks">
      <span class="pj-twisty">${open ? '&#9662;' : '&#9656;'}</span>
      <span class="pj-label">${esc(name)}</span>
      <span class="pj-meta">${p.count} entr${p.count === 1 ? 'y' : 'ies'} &middot; ${p.tasks.size} task${p.tasks.size === 1 ? '' : 's'} &middot; ${prettyDate(parseIso(p.first))} - ${prettyDate(parseIso(p.last))}</span>
      <div class="pj-track" title="Billable ${p.billable.toFixed(2)}h of ${p.hours.toFixed(2)}h">
        <div class="pj-fill" style="width:${((p.hours/max)*100).toFixed(1)}%">
          <div class="pj-bill" style="width:${bPct.toFixed(1)}%"></div>
        </div>
      </div>
      <span class="pj-hrs">${p.hours.toFixed(2)}</span>
      <span class="pj-pct">${share.toFixed(0)}%</span>
    </div>
    ${open ? `<div class="pj-tasks">
      ${tasks.map(([t, v]) => `<div class="pj-task">
        <span class="pjt-label">${esc(t)}</span>
        <span class="pjt-meta">${v.count}&times;</span>
        <div class="pjt-track"><div class="pjt-fill" style="width:${((v.hours/p.hours)*100).toFixed(1)}%"></div></div>
        <span class="pjt-hrs">${v.hours.toFixed(2)}</span>
      </div>`).join('')}
    </div>` : ''}
  </div>`;
}

function projectsAsText() {
  const entries = projScoped();
  const rows = groupByProject(entries);
  const total = entries.reduce((s, e) => s + (e.hours || 0), 0);
  const L = [];
  L.push(`Effort by project - ${$('pjRange').textContent}`);
  L.push('='.repeat(60));
  L.push(`Total ${total.toFixed(2)} h across ${rows.length} project(s)`);
  L.push('');
  for (const [name, p] of rows) {
    const share = total > 0 ? (p.hours / total) * 100 : 0;
    L.push(`${name.padEnd(46)}${p.hours.toFixed(2).padStart(9)}  ${share.toFixed(0).padStart(3)}%`);
    for (const [t, v] of [...p.tasks.entries()].sort((a, b) => b[1].hours - a[1].hours)) {
      L.push(`  ${t.padEnd(44)}${v.hours.toFixed(2).padStart(9)}`);
    }
    L.push('');
  }
  return L.join('\n');
}

/* ------------------------------------------------------- task management */

let tasks = {};  // { taskId: { id, title, status, priority, date, createdAt } }
let editingTaskId = null;   // card whose title is open for editing
let showDone = false;       // also show tasks completed on an earlier day

const TASK_STATUSES = ['New', 'In-Progress', 'Hold', 'Complete'];
const statusKey = s => s.toLowerCase().replace('-', '');

/* The day a card belongs to: the day it was finished if it is done, otherwise
   the day it was raised. Keeps work that spanned days on the day it closed. */
function taskDay(t) {
  if (t.completedAt) {
    const d = new Date(t.completedAt);
    if (!isNaN(d)) return iso(d);
  }
  return t.date;
}

async function loadTasks() {
  try {
    const data = await call('/api/tasks');
    tasks = data.tasks || {};
    renderTasks();
  } catch (err) {
    setTaskMsg('Could not load tasks: ' + err.message, 'err');
  }
}

/* Highest priority (1) first, then the pick order within that priority, then
   oldest first. Unsequenced cards (0) sink below sequenced ones. */
const seqRank = t => (t.seq || 0) || Infinity;
const byPriority = (a, b) =>
  (a.priority || 3) - (b.priority || 3) ||
  seqRank(a) - seqRank(b) ||
  new Date(a.createdAt) - new Date(b.createdAt);

/* Next free slot in the pick order for a priority, counting open work only. */
function nextSeq(priority) {
  const used = Object.values(tasks)
    .filter(t => (t.priority || 3) === priority && t.status !== 'Complete')
    .map(t => t.seq || 0);
  return Math.max(0, ...used) + 1;
}

function renderTasks() {
  const today = iso(new Date());
  // Anything still open stays on the board rather than disappearing at midnight;
  // finished work drops off with the day it was completed unless it is asked for.
  const all = Object.values(tasks);
  const earlierDone = all.filter(t => t.status === 'Complete' && taskDay(t) !== today);
  const todays = all.filter(t => t.status !== 'Complete' || taskDay(t) === today || showDone);

  $('taskDate').textContent = prettyDate(new Date());

  const toggle = $('btnShowDone');
  toggle.hidden = earlierDone.length === 0;
  toggle.textContent = showDone ? 'Hide earlier completed'
                                : `Show earlier completed (${earlierDone.length})`;
  toggle.classList.toggle('is-on', showDone);

  const done = todays.filter(t => t.status === 'Complete').length;
  $('boardStats').textContent = todays.length
    ? `${done} of ${todays.length} done`
    : '';

  $('tasksBoard').innerHTML = TASK_STATUSES.map(status => {
    const items = todays.filter(t => t.status === status).sort(byPriority);
    return `<section class="swimlane lane-${statusKey(status)}" data-status="${status}">
      <div class="lane-head">
        <span class="lane-dot"></span>
        <h3>${status}</h3>
        <span class="lane-count">${items.length}</span>
      </div>
      <div class="lane-body" data-status="${status}">
        ${items.map(taskHtml).join('')}
        <div class="lane-empty${items.length ? ' is-hidden' : ''}">Drop a card here</div>
      </div>
    </section>`;
  }).join('');

  const board = $('tasksBoard');
  board.querySelectorAll('.prio-pick').forEach(sel =>
    sel.addEventListener('change', () => updateTaskPriority(sel.dataset.taskId, +sel.value)));
  board.querySelectorAll('.seq-pick').forEach(inp => {
    inp.addEventListener('change', () => updateTaskSeq(inp.dataset.taskId, +inp.value));
    inp.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); inp.blur(); } });
  });
  board.querySelectorAll('.del-task').forEach(btn =>
    btn.addEventListener('click', () => deleteTask(btn.dataset.taskId)));
  board.querySelectorAll('.move-btn').forEach(btn =>
    btn.addEventListener('click', () => updateTaskStatus(btn.dataset.taskId, btn.dataset.status)));
  board.querySelectorAll('.edit-task').forEach(btn =>
    btn.addEventListener('click', () => { editingTaskId = btn.dataset.taskId; renderTasks(); }));
  board.querySelectorAll('.save-edit').forEach(btn =>
    btn.addEventListener('click', () => saveTaskTitle(btn.dataset.taskId)));
  board.querySelectorAll('.cancel-edit').forEach(btn =>
    btn.addEventListener('click', cancelTaskEdit));

  // Only ever one open editor, so it can take focus as soon as it is drawn.
  const editor = board.querySelector('.task-edit-input');
  if (editor) {
    editor.addEventListener('keydown', e => {
      e.stopPropagation();                       // keep Esc away from the entry form
      if (e.key === 'Enter')  { e.preventDefault(); saveTaskTitle(editor.dataset.taskId); }
      if (e.key === 'Escape') { e.preventDefault(); cancelTaskEdit(); }
    });
    editor.focus();
    editor.select();
  }

  wireDragDrop(board);
}

/* Stamped on every card, and doubles as the "still open from an earlier day"
   marker once the task stops being today's. */
function createdStamp(task) {
  const d = task.createdAt ? new Date(task.createdAt)
          : task.date      ? parseIso(task.date)
          : null;
  if (!d || isNaN(d)) return null;
  return {
    text: task.createdAt
      ? d.toLocaleString(undefined, { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' })
      : d.toLocaleDateString(undefined, { day: 'numeric', month: 'short' }),
    title: `Created ${d.toLocaleString()}`
  };
}

/* Shown on finished cards so the board says when the work actually closed. */
function completedStamp(task) {
  if (task.status !== 'Complete' || !task.completedAt) return null;
  const d = new Date(task.completedAt);
  if (isNaN(d)) return null;
  const sameDay = iso(d) === iso(new Date());
  return {
    text: sameDay
      ? d.toLocaleTimeString(undefined, { hour: '2-digit', minute: '2-digit' })
      : d.toLocaleString(undefined, { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' }),
    title: `Completed ${d.toLocaleString()}`
  };
}

function taskHtml(task) {
  const prio = task.priority || 3;
  const idx = TASK_STATUSES.indexOf(task.status);
  const prev = TASK_STATUSES[idx - 1];
  const next = TASK_STATUSES[idx + 1];
  // Only unfinished work from an earlier day is worth flagging; finished work
  // just carries its date.
  const stale = task.date && task.date !== iso(new Date()) && task.status !== 'Complete';
  const editing = task.id === editingTaskId;
  const made = createdStamp(task);
  const finished = completedStamp(task);
  return `<article class="card-task p${prio}-edge is-${statusKey(task.status)}${editing ? ' is-editing' : ''}"
                   draggable="${editing ? 'false' : 'true'}" data-task-id="${esc(task.id)}">
    <div class="ct-top">
      <span class="prio-badge p${prio}" title="Priority ${prio}">P${prio}</span>
      <select class="prio-pick" data-task-id="${esc(task.id)}" aria-label="Priority">
        ${[1,2,3,4,5].map(p => `<option value="${p}"${p === prio ? ' selected' : ''}>P${p}</option>`).join('')}
      </select>
      <input type="number" class="seq-pick" data-task-id="${esc(task.id)}" min="0" max="99" step="1"
             value="${task.seq || ''}" placeholder="#" aria-label="Pick order"
             title="Pick order within P${prio} - lower goes first">
      ${made ? `<span class="ct-created${stale ? ' is-carried' : ''}" title="${esc(made.title)}${stale ? ' - still open' : ''}">${esc(made.text)}</span>` : ''}
      ${finished ? `<span class="ct-created is-done" title="${esc(finished.title)}">&#10003; ${esc(finished.text)}</span>` : ''}
      <span class="ct-spacer"></span>
      ${editing ? '' : `<button type="button" class="edit-task" data-task-id="${esc(task.id)}" title="Edit title">&#9998;</button>`}
      <button type="button" class="del-task" data-task-id="${esc(task.id)}" title="Delete task">&times;</button>
    </div>
    ${editing ? `<div class="ct-edit">
      <input type="text" class="task-edit-input" data-task-id="${esc(task.id)}" maxlength="200" value="${esc(task.title)}">
      <div class="ct-edit-actions">
        <button type="button" class="btn btn-tiny cancel-edit">Cancel</button>
        <button type="button" class="btn btn-tiny btn-primary save-edit" data-task-id="${esc(task.id)}">Save</button>
      </div>
    </div>` : `<div class="ct-title">${esc(task.title)}</div>`}
    <div class="ct-move">
      ${prev ? `<button type="button" class="move-btn" data-task-id="${esc(task.id)}" data-status="${prev}" title="Move to ${prev}">&#8592; ${prev}</button>` : '<span></span>'}
      ${next ? `<button type="button" class="move-btn" data-task-id="${esc(task.id)}" data-status="${next}" title="Move to ${next}">${next} &#8594;</button>` : '<span></span>'}
    </div>
  </article>`;
}

/* Cards move between lanes by drag, with the arrow buttons as the fallback. */
function wireDragDrop(board) {
  let draggedId = null;

  board.querySelectorAll('.card-task').forEach(card => {
    card.addEventListener('dragstart', e => {
      draggedId = card.dataset.taskId;
      card.classList.add('is-dragging');
      e.dataTransfer.effectAllowed = 'move';
      e.dataTransfer.setData('text/plain', draggedId);
    });
    card.addEventListener('dragend', () => {
      draggedId = null;
      card.classList.remove('is-dragging');
      board.querySelectorAll('.lane-body').forEach(l => l.classList.remove('is-over'));
    });
  });

  board.querySelectorAll('.lane-body').forEach(lane => {
    lane.addEventListener('dragover', e => {
      e.preventDefault();
      e.dataTransfer.dropEffect = 'move';
      lane.classList.add('is-over');
    });
    lane.addEventListener('dragleave', () => lane.classList.remove('is-over'));
    lane.addEventListener('drop', e => {
      e.preventDefault();
      lane.classList.remove('is-over');
      const id = draggedId || e.dataTransfer.getData('text/plain');
      if (id) updateTaskStatus(id, lane.dataset.status);
    });
  });
}

async function logTaskEvent(payload) {
  return call('/api/events', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ ...payload, timestamp: new Date().toISOString() })
  });
}

async function addTask(title, priority) {
  const taskId = 'task-' + Date.now();
  const date = iso(new Date());
  const seq = nextSeq(priority);
  try {
    await logTaskEvent({ type: 'task:created', taskId, title, status: 'New', priority, seq, date });
    tasks[taskId] = { id: taskId, title, status: 'New', priority, seq, date, createdAt: new Date().toISOString(), completedAt: null };
    $('taskTitle').value = '';
    renderTasks();
    setTaskMsg('Task added.', 'ok');
  } catch (err) {
    setTaskMsg('Failed to add task: ' + err.message, 'err');
  }
}

async function updateTaskStatus(taskId, newStatus) {
  const task = tasks[taskId];
  if (!task || task.status === newStatus) return;
  const oldStatus = task.status;
  try {
    await logTaskEvent({ type: 'task:status_changed', taskId, oldStatus, newStatus });
    task.status = newStatus;
    task.completedAt = newStatus === 'Complete' ? new Date().toISOString() : null;
    renderTasks();
    setTaskMsg(`Moved to ${newStatus}.`, 'ok');
  } catch (err) {
    setTaskMsg('Failed to update task: ' + err.message, 'err');
  }
}

async function updateTaskPriority(taskId, newPriority) {
  const task = tasks[taskId];
  if (!task || task.priority === newPriority) return;
  const oldPriority = task.priority;
  try {
    await logTaskEvent({ type: 'task:priority_changed', taskId, oldPriority, newPriority });
    task.priority = newPriority;
    renderTasks();
    setTaskMsg(`Priority set to P${newPriority}.`, 'ok');
  } catch (err) {
    setTaskMsg('Failed to update priority: ' + err.message, 'err');
  }
}

async function updateTaskSeq(taskId, newSeq) {
  const task = tasks[taskId];
  if (!task) return;
  newSeq = Number.isFinite(newSeq) && newSeq > 0 ? Math.min(99, Math.round(newSeq)) : 0;
  if ((task.seq || 0) === newSeq) return renderTasks();
  const oldSeq = task.seq || 0;
  try {
    await logTaskEvent({ type: 'task:seq_changed', taskId, oldSeq, newSeq });
    task.seq = newSeq;
    renderTasks();
    setTaskMsg(newSeq ? `Pick order set to #${newSeq}.` : 'Pick order cleared.', 'ok');
  } catch (err) {
    setTaskMsg('Failed to update order: ' + err.message, 'err');
  }
}

async function deleteTask(taskId) {
  if (!confirm('Delete this task?')) return;
  try {
    await logTaskEvent({ type: 'task:deleted', taskId });
    delete tasks[taskId];
    if (editingTaskId === taskId) editingTaskId = null;
    renderTasks();
    setTaskMsg('Task deleted.', 'ok');
  } catch (err) {
    setTaskMsg('Failed to delete task: ' + err.message, 'err');
  }
}

function cancelTaskEdit() {
  editingTaskId = null;
  renderTasks();
}

async function saveTaskTitle(taskId) {
  const task = tasks[taskId];
  const input = $('tasksBoard').querySelector('.task-edit-input');
  if (!task || !input) return;

  const newTitle = input.value.trim();
  if (!newTitle)               return setTaskMsg('A task needs a title.', 'err');
  if (newTitle === task.title) return cancelTaskEdit();

  try {
    await logTaskEvent({ type: 'task:title_changed', taskId, oldTitle: task.title, newTitle });
    task.title = newTitle;
    editingTaskId = null;
    renderTasks();
    setTaskMsg('Task updated.', 'ok');
  } catch (err) {
    setTaskMsg('Failed to rename task: ' + err.message, 'err');
  }
}

function setTaskMsg(text, kind) {
  const el = $('taskMsg');
  el.textContent = text;
  el.className = 'msg' + (kind ? ' ' + kind : '');
  if (kind === 'ok') setTimeout(() => { if (el.textContent === text) el.textContent = ''; }, 3000);
}

/* ---------------------------------------------------------------- holidays */

/* Comes from the Holidays sheet in the workbook - Excel stays the only place
   this is edited. */
let holidays = [];

async function loadHolidays() {
  try {
    const data = await call('/api/holidays');
    holidays = (data.holidays || []).slice().sort((a, b) => a.date.localeCompare(b.date));
  } catch {
    holidays = [];
  }
  renderHolidaySide();
  if (!$('viewHolidays').hidden) renderHolidays();
}

const daysUntil = d => Math.round((parseIso(d) - parseIso(iso(new Date()))) / 86400000);

const relDays = n => n === 0 ? 'today' : n === 1 ? 'tomorrow' : `in ${n} days`;

const upcomingHolidays = () => holidays.filter(h => daysUntil(h.date) >= 0);

const weekdayOf = h => h.day || parseIso(h.date).toLocaleDateString(undefined, { weekday: 'long' });

const holidayLong = h => `${weekdayOf(h)}, ${parseIso(h.date).toLocaleDateString(undefined, { day: 'numeric', month: 'long' })}`;

/* Sits in the sidebar, so every tab shows what is coming up. */
function renderHolidaySide() {
  const next = upcomingHolidays().slice(0, 3);
  $('holidaySide').innerHTML = next.length
    ? next.map((h, i) => {
        const n = daysUntil(h.date);
        return `<div class="hs-row${i === 0 ? ' is-next' : ''}">
          <span class="hs-name" title="${esc(holidayLong(h))}">${esc(h.name)}</span>
          <span class="hs-when">${esc(weekdayOf(h).slice(0, 3))} ${esc(prettyDate(parseIso(h.date)))}</span>
          <span class="hs-in">${esc(relDays(n))}</span>
        </div>`;
      }).join('')
    : '<div class="hs-none">Nothing coming up.</div>';
}

function renderHolidays() {
  const upcoming = upcomingHolidays();
  const years = [...new Set(holidays.map(h => h.date.slice(0, 4)))];

  $('holYear').textContent  = years.join(', ');
  $('holStats').textContent = `${upcoming.length} of ${holidays.length} left`;
  $('holNext').textContent  = upcoming.length
    ? `Next: ${upcoming[0].name} - ${relDays(daysUntil(upcoming[0].date))}`
    : 'Nothing coming up';

  if (!holidays.length) {
    $('holBody').innerHTML = '<p class="empty">No dates on the Holidays sheet in the workbook.</p>';
    return;
  }

  $('holBody').innerHTML = `<div class="hol-grid">` + holidays.map(h => {
    const n   = daysUntil(h.date);
    const cls = n < 0 ? 'is-past' : n === 0 ? 'is-today' : '';
    return `<div class="hol-row ${cls}">
      <span class="hol-date">${esc(prettyDate(parseIso(h.date)))}</span>
      <span class="hol-day">${esc(weekdayOf(h))}</span>
      <span class="hol-name">${esc(h.name)}</span>
      <span class="hol-in">${n < 0 ? 'past' : relDays(n)}</span>
      <button type="button" class="hol-del" data-date="${esc(h.date)}" title="Remove">&times;</button>
    </div>`;
  }).join('') + `</div>`;

  $('holBody').querySelectorAll('.hol-del').forEach(b =>
    b.addEventListener('click', () => removeHoliday(b.dataset.date)));
}

function setHolMsg(text, kind) {
  const el = $('holMsg');
  el.textContent = text;
  el.className = 'msg' + (kind ? ' ' + kind : '');
  if (kind === 'ok') setTimeout(() => { if (el.textContent === text) el.textContent = ''; }, 3000);
}

async function holidayAction(payload, okText) {
  try {
    const data = await call('/api/holiday', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
    holidays = (data.holidays || []).slice().sort((a, b) => a.date.localeCompare(b.date));
    renderHolidaySide();
    renderHolidays();
    setHolMsg(okText, 'ok');  } catch (err) {
    setHolMsg(err.message, 'err');
  }
}

function removeHoliday(date) {
  const h = holidays.find(x => x.date === date);
  if (!h || !confirm(`Remove ${h.name} on ${h.date}?`)) return;
  holidayAction({ action: 'removeHoliday', date }, `Removed ${h.name}.`);
}

/* ---------------------------------------------------------------- views */

const VIEWS = {
  entry: 'viewEntry', day: 'viewDay', report: 'viewReport', tasks: 'viewTasks',
  projects: 'viewProjectAdmin', team: 'viewTeam', effort: 'viewEffort',
  holidays: 'viewHolidays', users: 'viewUsers'
};

function showView(name) {
  for (const [v, id] of Object.entries(VIEWS)) $(id).hidden = v !== name;
  document.querySelectorAll('.vtab').forEach(t => t.classList.toggle('is-on', t.dataset.view === name));
  if (name === 'report' && !reportData) loadReport(null);
  if (name === 'day') {
    if (!dayDate) dayDate = iso(new Date());
    if (!dayWeek || dayDate < dayWeek.weekStart || dayDate > dayWeek.weekEnd) {
      loadDay(dayDate);
    } else {
      renderDay();
    }
  }
  if (name === 'tasks') loadTasks();
  if (name === 'team')  loadTeam();
  if (name === 'effort') loadProjects();
  if (name === 'holidays') renderHolidays();
  if (name === 'users') loadUsers();
}

/* ---------------------------------------------------------------- wire up */

$('entryForm').addEventListener('submit', e => { e.preventDefault(); submit(false); });
$('btnKeep').addEventListener('click', () => submit(true));
$('btnReload').addEventListener('click', load);
$('btnCancelEdit').addEventListener('click', cancelEdit);
$('project').addEventListener('input', refreshTaskList);
$('btnLogout').addEventListener('click', () => logout());
$('btnChangePassword').addEventListener('click', () => showPasswordView(false));
$('cancelPassword').addEventListener('click', hidePasswordView);
$('passwordForm').addEventListener('submit', e => { e.preventDefault(); changePassword(); });
$('forgotPassword').addEventListener('click', () => { $('forgotMsg').hidden = !$('forgotMsg').hidden; });

$('loginForm').addEventListener('submit', async e => {
  e.preventDefault();
  const username = $('loginUser').value.trim();
  const password = $('loginPass').value;
  const msg = $('loginMsg');
  msg.textContent = 'Signing in...';
  msg.className = 'msg';
  try {
    await login(username, password);
    msg.textContent = '';
    $('loginForm').reset();
  } catch (err) {
    msg.textContent = err.message;
    msg.className = 'msg err';
  }
});

$('userForm').addEventListener('submit', e => {
  e.preventDefault();
  createUser({
    username: $('newUsername').value.trim(),
    name: $('newUserName').value.trim(),
    password: $('newUserPassword').value
  });
});

$('resetUserForm').addEventListener('submit', e => {
  e.preventDefault();
  resetUserPassword($('resetUsername').value, $('resetUserPassword').value);
});

$('taskForm').addEventListener('submit', e => {
  e.preventDefault();
  const title = $('taskTitle').value.trim();
  if (title) addTask(title, +$('taskPriority').value);
});

$('btnShowDone').addEventListener('click', () => { showDone = !showDone; renderTasks(); });

$('padmForm').addEventListener('submit', e => {
  e.preventDefault();
  const name = $('padmName').value.trim();
  if (!name) return;
  $('padmName').value = '';
  projectAction({ action: 'addProject', name, status: $('padmStatus').value }, `Added "${name}".`);
});

$('teamForm').addEventListener('submit', e => {
  e.preventDefault();
  const weekStart = mondayOfWeekValue($('teamWeek').value);
  if (!weekStart) return setTeamMsg('Pick a week (format 2026-W11).', 'err');
  const person = $('teamPerson').value;
  if (!person) return setTeamMsg('Pick a person.', 'err');
  const notes = $('teamNotes').value.trim();
  $('teamNotes').value = '';
  teamAction({ action: 'addWeek', person, activity: $('teamActivity').value, weekStart, notes },
             `Logged ${$('teamActivity').value} for ${person}.`);
});

$('teamYear').addEventListener('change', () => { teamYear = +$('teamYear').value; renderTeam(); });

$('holForm').addEventListener('submit', e => {
  e.preventDefault();
  const date = $('holDate').value;
  const name = $('holName').value.trim();
  if (!date) return setHolMsg('Pick a date.', 'err');
  if (!name) return setHolMsg('Give the holiday a name.', 'err');
  $('holDate').value = '';
  $('holName').value = '';
  holidayAction({ action: 'addHoliday', date, name }, `Added ${name}.`);
});

$('teamPersonForm').addEventListener('submit', e => {
  e.preventDefault();
  const name = $('teamNewPerson').value.trim();
  if (!name) return;
  $('teamNewPerson').value = '';
  teamAction({ action: 'addPerson', person: name }, `Added ${name}.`);
});

document.querySelectorAll('.vtab').forEach(t =>
  t.addEventListener('click', () => showView(t.dataset.view)));

$('wkPrev').addEventListener('click', () => shiftWeek(-7));
$('wkNext').addEventListener('click', () => shiftWeek(7));
$('wkThis').addEventListener('click', () => loadReport(iso(mondayOf(new Date()))));
$('btnPrint').addEventListener('click', () => window.print());

$('dayPrev').addEventListener('click', () => shiftDay(-1));
$('dayNext').addEventListener('click', () => shiftDay(1));
$('dayToday').addEventListener('click', () => loadDay(iso(new Date())));
$('dayPick').addEventListener('change', () => { if ($('dayPick').value) loadDay($('dayPick').value); });
$('btnPrintDay').addEventListener('click', () => window.print());

$('pjRangeSel').addEventListener('change', () => {
  const custom = $('pjRangeSel').value === 'custom';
  $('pjFrom').hidden = !custom;
  $('pjTo').hidden   = !custom;
  renderProjects();
});
$('pjFrom').addEventListener('change', renderProjects);
$('pjTo').addEventListener('change', renderProjects);
$('btnPrintProjects').addEventListener('click', () => window.print());

/* Clipboard needs a user gesture and can still be blocked (file://, no HTTPS),
   so fall back to a prompt the text can be copied out of by hand. */
function wireCopy(btnId, ready, getText) {
  const btn = $(btnId);
  btn.addEventListener('click', async () => {
    if (!ready()) return;
    const txt = getText();
    try {
      await navigator.clipboard.writeText(txt);
      btn.textContent = 'Copied';
      setTimeout(() => btn.textContent = 'Copy as text', 1600);
    } catch {
      window.prompt('Copy the text below:', txt);
    }
  });
}

wireCopy('btnCopyReport', () => !!reportData, reportAsText);
wireCopy('btnCopyDay',    () => !!dayWeek,    dayAsText);
wireCopy('btnCopyProjects', () => !!allEntries, projectsAsText);

$('hours').addEventListener('input', () => {
  const h = parseHours($('hours').value);
  $('hoursHint').textContent = h === null ? '' : `= ${fmtHM(h)} (${h.toFixed(2)})`;
});

document.querySelectorAll('.chip').forEach(c => c.addEventListener('click', () => {
  $('hours').value = c.dataset.h;
  $('hours').dispatchEvent(new Event('input'));
}));

// Ctrl+Enter submits from anywhere in the form; Esc leaves edit mode.
document.addEventListener('keydown', e => {
  if (e.ctrlKey && e.key === 'Enter') { e.preventDefault(); submit(false); }
  if (e.key === 'Escape' && editingRow !== null) { e.preventDefault(); cancelEdit(); }
});

$('date').value = iso(new Date());
loadHolidays();
$('btnCopyUpdate').addEventListener('click', copyYesterdayUpdate);
$('btnOpenYesterday').addEventListener('click', () => {
  showView('day');
  loadDay(iso(new Date(Date.now() - 86400000)));
});
initAuth();
