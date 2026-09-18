const $ = id => document.getElementById(id);
let entries = [];
let currentUser = null;
let projects = [];
let editingEntryId = null;
let reportWeekStart = null;
const isoDate = d => `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
let dayDate = isoDate(new Date());
let optionCatalogLoaded = false;
let tasks = {};
let editingTaskId = null;
let showDone = false;
let reminders = [];
let reminderFilter = 'all';
const TASK_STATUSES = ['New', 'In-Progress', 'Hold', 'Complete'];
const statusKey = s => s.toLowerCase().replace(/[^a-z0-9]+/g, '');
const esc = value => String(value ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

function isoWeek(date) {
  const value = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  value.setUTCDate(value.getUTCDate() + 4 - (value.getUTCDay() || 7));
  const yearStart = new Date(Date.UTC(value.getUTCFullYear(), 0, 1));
  return { year: value.getUTCFullYear(), week: Math.ceil((((value - yearStart) / 86400000) + 1) / 7) };
}

function formatDayMonth(d) {
  const day = String(d.getDate()).padStart(2, '0');
  const month = d.toLocaleDateString('en-GB', { month: 'short' });
  return `${day} ${month}`;
}

function getWeekRange(dateOrIso) {
  const d = typeof dateOrIso === 'string'
    ? (/^\d{4}-\d{2}-\d{2}$/.test(dateOrIso)
        ? (([y, m, day]) => new Date(+y, +m - 1, +day))(dateOrIso.split('-'))
        : new Date(dateOrIso))
    : new Date(dateOrIso);
  const monday = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  monday.setDate(monday.getDate() - ((monday.getDay() + 6) % 7));
  const friday = new Date(monday.getFullYear(), monday.getMonth(), monday.getDate() + 4);
  const w = isoWeek(monday);
  const weekCode = `${w.year}-W${String(w.week).padStart(2, '0')}`;
  const mondayStr = formatDayMonth(monday);
  const fridayStr = `${formatDayMonth(friday)} ${friday.getFullYear()}`;
  return { weekCode, rangeText: `${mondayStr} – ${fridayStr}`, monday, friday, mondayIso: isoDate(monday) };
}

function weekToDateRange(weekCode) {
  const m = /^(\d{4})-W(\d{1,2})$/.exec(String(weekCode || '').trim());
  if (!m) return weekCode;
  const year = +m[1];
  const weekNum = +m[2];
  const jan4 = new Date(Date.UTC(year, 0, 4));
  const monday = new Date(jan4);
  monday.setHours(0, 0, 0, 0);
  monday.setDate(monday.getDate() - ((monday.getDay() + 6) % 7));
  monday.setUTCDate(monday.getUTCDate() + (weekNum - 1) * 7);
  const friday = new Date(monday);
  friday.setUTCDate(friday.getUTCDate() + 4);
  return `${formatDayMonth(monday)} – ${formatDayMonth(friday)} ${friday.getFullYear()}`;
}

function formatUserAgent(ua) {
  if (!ua || ua === 'Unknown') return 'Unknown device';
  let browser = 'Browser';
  if (/Edg\//i.test(ua)) browser = 'Edge';
  else if (/Chrome\//i.test(ua)) browser = 'Chrome';
  else if (/Firefox\//i.test(ua)) browser = 'Firefox';
  else if (/Safari\//i.test(ua)) browser = 'Safari';

  let os = 'Device';
  if (/Windows/i.test(ua)) os = 'Windows';
  else if (/Macintosh|Mac OS/i.test(ua)) os = 'macOS';
  else if (/Android/i.test(ua)) os = 'Android';
  else if (/iPhone|iPad|iOS/i.test(ua)) os = 'iOS';
  else if (/Linux/i.test(ua)) os = 'Linux';

  return `${browser} on ${os}`;
}

async function api(path, options = {}) {
  const response = await fetch(path, { credentials: 'same-origin', ...options });
  const data = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(data.error || `Request failed (${response.status})`);
  return data;
}

function showMessage(id, text, error = true) {
  const element = $(id);
  element.textContent = text || '';
  element.style.color = error ? 'var(--danger)' : 'var(--accent)';
}

async function showApp(user) {
  currentUser = { ...user, isAdmin: Boolean(user.isAdmin) || String(user.username || '').toLowerCase() === 'admin' };
  $('loginPanel').hidden = true;
  $('appPanel').hidden = false;
  $('passwordPanel').hidden = true;
  const displayName = user.name || user.username;
  if ($('currentUser')) $('currentUser').textContent = `Signed in as ${displayName}`;
  const wsUser = $('workspaceUser');
  if (wsUser) wsUser.textContent = `👤 ${displayName}`;
  if ($('usersTab')) $('usersTab').hidden = !currentUser.isAdmin;
  if ($('sessionsTab')) $('sessionsTab').hidden = !currentUser.isAdmin;
  initScratchpad();
  await Promise.all([loadEntries(), loadProjectOptions(), loadOptionCatalog(), loadSidebarHolidays(), loadTasks(), loadReminders()]);
  if (!reportWeekStart) reportWeekStart = getMonday(new Date());
  document.querySelectorAll('.tab').forEach(item => item.classList.toggle('active', item.dataset.view === 'tasks'));
  $('entryWorkspace').hidden = true;
  $('tabWorkspace').hidden = false;
  $('viewTitle').textContent = "Today's Tasks";
  $('viewMessage').textContent = '';
  renderTab('tasks').catch(error => { $('tabWorkspace').innerHTML = `<div class="tab-card"><p class="message">${esc(error.message)}</p></div>`; });
}

function showLogin() {
  currentUser = null;
  $('loginPanel').hidden = false;
  $('appPanel').hidden = true;
  $('passwordPanel').hidden = true;
  if ($('usersTab')) $('usersTab').hidden = true;
  if ($('sessionsTab')) $('sessionsTab').hidden = true;
  if ($('currentUser')) $('currentUser').textContent = '';
  if ($('workspaceUser')) $('workspaceUser').textContent = '';
}

async function loadEntries() {
  try {
    const data = await api('/api/entries');
    entries = data.entries || [];
    $('entryTotal').textContent = entries.length;
    $('todayTotal').textContent = `${entries.filter(entry => entry.date === new Date().toISOString().slice(0, 10)).reduce((total, entry) => total + Number(entry.hours || 0), 0).toFixed(2)}h`;
    const monday = getMonday(new Date()).toISOString().slice(0, 10);
    const weekEnd = shiftDate(monday, 6);
    $('weekTotal').textContent = `${entries.filter(entry => entry.date >= monday && entry.date <= weekEnd).reduce((total, entry) => total + Number(entry.hours || 0), 0).toFixed(2)}h`;
    const groups = new Map();
    entries.forEach(entry => { if (!groups.has(entry.date)) groups.set(entry.date, []); groups.get(entry.date).push(entry); });
    $('entries').innerHTML = [...groups.entries()].map(([date, items]) => {
      const dayTotal = items.reduce((sum, entry) => sum + Number(entry.hours || 0), 0);
      const dayLabel = new Date(`${date}T00:00:00`).toLocaleDateString('en-GB', { weekday: 'short', day: '2-digit', month: 'short', year: 'numeric' });
      return `<section class="recent-day"><header><strong>${esc(dayLabel)}</strong><b>${dayTotal.toFixed(2)} h</b></header>${items.map(entry => `<article class="recent-entry"><div class="recent-main"><strong>${esc(entry.project)}</strong><span>${esc(entry.task || '')}</span><p>${esc(entry.notes || 'No notes')}</p><div class="recent-tags"><em>${esc(entry.category || 'Other')}</em><em>${esc(entry.billable === 'Yes' ? 'Billable' : 'Non Billable')}</em><em>${esc(entry.nonBillableReason || '')}</em><em>${esc(entry.budget || '')}</em></div></div><strong class="recent-hours">${Number(entry.hours || 0).toFixed(2)}h</strong><div class="entry-row-actions"><button type="button" class="entry-action" data-copy="${entry.id}">Copy</button><button type="button" class="entry-action" data-edit="${entry.id}">Edit</button><button type="button" class="entry-action delete-action" data-delete="${entry.id}">Delete</button></div></article>`).join('')}</section>`;
    }).join('') || '<p class="empty">No entries yet.</p>';
    $('entries').querySelectorAll('[data-copy]').forEach(button => button.addEventListener('click', () => loadEntryIntoForm(Number(button.dataset.copy), false)));
    $('entries').querySelectorAll('[data-edit]').forEach(button => button.addEventListener('click', () => loadEntryIntoForm(Number(button.dataset.edit), true)));
    $('entries').querySelectorAll('[data-delete]').forEach(button => button.addEventListener('click', () => deleteEntry(Number(button.dataset.delete))));
    if (!optionCatalogLoaded) updateEntryOptions();
  } catch (error) { showMessage('entryMessage', error.message); }
}

async function loadSidebarHolidays() {
  try {
    const data = await api('/api/holidays');
    const today = new Date(); today.setHours(0, 0, 0, 0);
    const upcoming = (data.holidays || []).filter(item => new Date(`${item.date}T00:00:00`) >= today).slice(0, 3);
    $('sidebarHolidayList').innerHTML = upcoming.map(item => { const date = new Date(`${item.date}T00:00:00`); const days = Math.ceil((date - today) / 86400000); return `<div class="sidebar-holiday"><strong>${esc(item.name)}</strong><span>${esc(date.toLocaleDateString('en-GB', { weekday: 'short', day: 'numeric', month: 'short' }))}</span><small>${days === 0 ? 'today' : `in ${days} days`}</small></div>`; }).join('') || '<p class="sidebar-empty">No upcoming holidays.</p>';
  } catch (_) { $('sidebarHolidayList').innerHTML = ''; }
}

async function loadProjectOptions() {
  try {
    const data = await api('/api/projects');
    projects = data.projects || [];
    const activeProjects = projects.filter(p => p.status === 'In-Progress');
    $('project').innerHTML = '<option value="">Select project</option>' + activeProjects
      .map(project => `<option value="${esc(project.name)}">${esc(project.name)}</option>`).join('');
    updateTaskOptions();
  } catch (_) {
    $('project').innerHTML = '<option value="">Select project</option>';
  }
}

function updateTaskOptions() {
  const project = projects.find(item => item.name === $('project').value);
  $('task').innerHTML = '<option value="">Select task</option>' + (project?.tasks || [])
    .map(task => `<option value="${esc(task)}">${esc(task)}</option>`).join('');
}

function updateEntryOptions() {
  [['category', 'category'], ['budget', 'budget'], ['billable', 'billable'], ['nonBillableReason', 'nonBillableReason']].forEach(([selectId, field]) => {
    const values = [...new Set(entries.map(entry => entry[field]).filter(Boolean))].sort();
    const placeholder = selectId === 'budget' ? 'Select budget' : selectId === 'billable' ? 'Select billable status' : selectId === 'nonBillableReason' ? 'Select reason' : 'Select category';
    $(selectId).innerHTML = `<option value="">${placeholder}</option>` + values.map(value => `<option value="${esc(value)}">${esc(value)}</option>`).join('');
  });
}

async function loadOptionCatalog() {
  try {
    const data = await api('/api/options');
    optionCatalogLoaded = true;
    [['category', 'category'], ['budget', 'budget'], ['billable', 'billable'], ['nonBillableReason', 'nonBillableReason']].forEach(([selectId, key]) => {
      const values = data.options?.[key] || [];
      const placeholder = selectId === 'budget' ? 'Select budget' : selectId === 'billable' ? 'Select billable status' : selectId === 'nonBillableReason' ? 'Select reason' : 'Select category';
      $(selectId).innerHTML = `<option value="">${placeholder}</option>` + values.map(value => `<option value="${esc(value)}">${esc(value)}</option>`).join('');
    });
  } catch (_) { updateEntryOptions(); }
}

async function renderTab(view) {
  const tabWorkspace = $('tabWorkspace');
  const totalHours = entries.reduce((total, entry) => total + Number(entry.hours || 0), 0).toFixed(2);
  const projectNames = [...new Set(entries.map(entry => entry.project).filter(Boolean))].sort();
  const rows = entries.slice(0, 12).map(entry => `<div class="data-row"><strong>${esc(entry.date)}</strong><span>${esc(entry.project)}</span><span>${esc(entry.hours)}h</span></div>`).join('');
  const reportHtml = view === 'report' ? buildWeeklyReport() : '';
  const content = {
    day: buildDayView(),
    tasks: `<div class="taskboard-shell tab-card tasks-card"><div class="tasks-head"><h2>Today's Tasks</h2><span class="board-stats" id="boardStats">0 of 0 done</span><button type="button" id="btnShowDone" class="btn-tiny" hidden>Show earlier completed</button><span class="task-date" id="taskDate">${new Date().toLocaleDateString('en-GB', { day: 'numeric', month: 'short' })}</span></div><form id="taskForm" autocomplete="off" class="task-input-form"><div class="task-input-group"><input type="text" id="taskTitle" placeholder="Add a new task..." maxlength="200" required><select id="taskPriority" class="prio-select" aria-label="Priority"><option value="1">P1 - Highest</option><option value="2">P2 - High</option><option value="3" selected>P3 - Medium</option><option value="4">P4 - Low</option><option value="5">P5 - Lowest</option></select><button type="submit" class="btn btn-primary btn-sm">+ Add Task</button></div><div id="taskMsg" class="msg"></div></form><div id="tasksBoard" class="tasks-board"></div></div>`,
    reminders: `<div class="tab-card reminder-card"><div class="reminder-head"><div><p class="eyebrow">REMINDERS &amp; SCHEDULE</p><h2>Reminders <span id="reminderStats" class="reminder-stats-badge"></span></h2></div><div class="reminder-head-actions"><button type="button" id="btnQuickAddReminder" class="btn btn-primary btn-sm">+ New Reminder</button></div></div><form id="reminderForm" autocomplete="off" class="reminder-input-form"><div class="reminder-form-grid"><input type="text" id="reminderTitle" placeholder="What do you need to remember?" maxlength="200" required><input type="date" id="reminderDueDate" class="date-input" required aria-label="Due date"><input type="time" id="reminderDueTime" class="time-input" aria-label="Due time (optional)"><select id="reminderPriority" class="prio-select" aria-label="Priority"><option value="1">P1 - Highest</option><option value="2">P2 - High</option><option value="3" selected>P3 - Medium</option><option value="4">P4 - Low</option><option value="5">P5 - Lowest</option></select><button type="submit" class="btn btn-primary btn-sm">+ Add</button></div><div class="reminder-form-extra"><input type="text" id="reminderNotes" placeholder="Additional details or notes (optional)..." maxlength="300"></div><div id="reminderMsg" class="msg"></div></form><div class="reminder-filters"><button type="button" class="reminder-filter-btn is-active" data-filter="all">All Active</button><button type="button" class="reminder-filter-btn" data-filter="today">Due Today</button><button type="button" class="reminder-filter-btn" data-filter="upcoming">Upcoming</button><button type="button" class="reminder-filter-btn" data-filter="completed">Completed</button></div><div id="reminderList" class="reminder-list"></div></div>`,
    report: `<div class="report-card"><div class="report-toolbar"><div><button type="button" id="reportPrev" class="report-nav">&#8592; Previous</button><button type="button" id="reportThis" class="report-nav">This week</button><button type="button" id="reportNext" class="report-nav">Next &#8594;</button><p class="eyebrow">HOURS BY WEEK / PROJECT / TASK / DATE</p></div><div class="report-toolbar-actions"><button type="button" id="expandReport">Expand all</button><button type="button" id="collapseReport">Collapse all</button><button type="button" id="exportReportPdf" class="btn-pdf-export" title="Export as PDF / Print">📄 Export as PDF</button></div></div>${reportHtml}</div>`,
    projects: `<div class="tab-card"><p class="eyebrow">WORK CATALOG</p><h2>Projects</h2><form id="projectForm" class="management-form"><input id="newProject" placeholder="Project name" required><select id="projectStatus"><option>New</option><option>In-Progress</option><option>Hold</option><option>Complete</option></select><button type="submit">Create project</button></form><div id="projectList" class="project-list"></div></div>`,
    team: `<div class="team-card"><div class="team-head"><h2>Team activities <span id="teamStats"></span></h2><label>YEAR<select id="teamYear"></select></label></div><form id="teamForm" class="team-form"><select id="teamPerson" required aria-label="Team member"><option value="">Select a person...</option></select><select id="teamActivity" aria-label="Activity"><option value="WFH">WFH</option></select><div class="team-date-field"><div class="team-input-calendar-wrap"><input id="teamDate" class="team-date-input" type="text" readonly required placeholder="Select week..." aria-label="Select week date" title="Click to open calendar and select a week"><button type="button" id="teamCalendarBtn" class="team-calendar-btn" title="Open Calendar">📅</button></div><input id="teamWeek" type="hidden" required><div id="teamWeekPreview" class="team-week-preview">Pick date to select Mon–Fri week</div><div id="teamCalendarPopup" class="team-calendar-popup" hidden></div></div><input id="teamNotes" placeholder="Notes (optional)" aria-label="Notes"><button type="submit">+ Log week</button></form><div id="teamWeekStrip" class="team-week-strip"><button type="button" id="twsPrev" class="tws-btn" title="Previous Week">&larr; Prev Week</button><div id="twsDays" class="tws-days"></div><button type="button" id="twsNext" class="tws-btn" title="Next Week">Next Week &rarr;</button></div><div id="teamMessage" class="msg"></div><div id="teamList" class="team-list"></div><details class="manage-people"><summary>Manage people</summary><form id="newTeamForm"><input id="newTeamMember" placeholder="Add someone to the team..." required><button type="submit">Add person</button></form></details></div>`,
    holidays: `<div class="holiday-card"><div class="holiday-head"><h2>Holidays <span id="holidayYears"></span></h2><span id="holidayNext"></span></div><form id="holidayForm" class="holiday-form"><input id="holidayDate" class="date-input" type="date" placeholder="dd-mm-yyyy" required><input id="holidayName" placeholder="Holiday name..." required><button type="submit">+ Add Holiday</button></form><p class="holiday-note">Holidays live on the Cloudflare D1 calendar. Adding one here makes it available to all users.</p><div id="holidayList" class="holiday-list"></div></div>`,
    sessions: `<div class="tab-card admin-panel-card sessions-panel-card">
      <div class="admin-panel-head">
        <div class="section-title-bar">
          <div>
            <p class="eyebrow">SECURITY &amp; MONITORING</p>
            <h2>Active Logged In Sessions <span id="activeSessionsCount" class="sessions-count-badge"></span></h2>
            <p class="muted small">Real-time authenticated active sessions across all devices.</p>
          </div>
          <button type="button" id="reloadSessionsBtn" class="btn btn-tiny" title="Refresh active sessions">🔄 Refresh</button>
        </div>
      </div>
      <div id="sessionsMessage" class="msg"></div>
      <div id="activeSessionsList" class="sessions-list"></div>
    </div>`,
    users: `<div class="tab-card admin-panel-card users-panel-card">
      <div class="admin-panel-head">
        <p class="eyebrow">ADMINISTRATION</p>
        <h2>User Management</h2>
      </div>
      <div class="user-mgmt-layout">
        <div class="user-forms-column">
          <h3>Create User</h3>
          <form id="userForm" class="management-form"><input id="newUsername" placeholder="Username" required><input id="newUserName" placeholder="Display name" required><input id="newUserEmail" type="email" placeholder="Email address" required><input id="newUserPassword" type="password" minlength="8" placeholder="Temporary password" required><button type="submit">Create user</button></form>

          <h3>Update Email</h3>
          <form id="emailForm" class="management-form reset-form"><select id="emailUsername" required aria-label="User to update"></select><input id="editUserEmail" type="email" placeholder="Email address" required><button type="submit">Save email</button></form>

          <h3>Reset Password</h3>
          <form id="resetUserForm" class="management-form reset-form"><select id="resetUsername" required aria-label="User to reset"></select><input id="resetPassword" type="password" minlength="8" placeholder="New temporary password" required><button type="submit">Reset password</button></form>

          <h3>One-Time Reset Code</h3>
          <form id="resetCodeForm" class="management-form reset-form"><select id="codeUsername" required aria-label="User for reset code"></select><button type="submit">Generate code</button></form>

          <div id="userMessage" class="message"></div>
        </div>

        <div class="user-list-column">
          <h3>Registered Users</h3>
          <div id="userList" class="data-list"></div>
        </div>
      </div>
    </div>`,
    options: `<div class="tab-card"><p class="eyebrow">ENTRY CATALOG</p><h2>Dropdown options</h2><p class="muted">Add values that will appear in the Entry form dropdowns.</p><div id="optionForms" class="option-forms"></div><div id="optionMessage" class="message"></div></div>`
  };
  if ((view === 'users' || view === 'sessions') && !currentUser?.isAdmin) return;
  tabWorkspace.innerHTML = content[view] || content.day;
  tabWorkspace.hidden = false;
  $('entryWorkspace').hidden = true;
  await loadManagementData(view);
  if (view === 'report') {
    $('reportPrev').addEventListener('click', () => { reportWeekStart = shiftWeek(reportWeekStart, -7); renderTab('report'); });
    $('reportThis').addEventListener('click', () => { reportWeekStart = getMonday(new Date()); renderTab('report'); });
    $('reportNext').addEventListener('click', () => { reportWeekStart = shiftWeek(reportWeekStart, 7); renderTab('report'); });
    $('expandReport').addEventListener('click', () => document.querySelectorAll('.report-week, .report-project, .report-task').forEach(item => item.open = true));
    $('collapseReport').addEventListener('click', () => document.querySelectorAll('.report-week, .report-project, .report-task').forEach(item => item.open = false));
    $('exportReportPdf')?.addEventListener('click', () => {
      document.querySelectorAll('.report-week, .report-project, .report-task').forEach(item => item.open = true);
      window.print();
    });
  }
  if (view === 'day') {
    $('dayPrev').addEventListener('click', () => { dayDate = shiftDate(dayDate, -1); renderTab('day'); });
    $('dayNext').addEventListener('click', () => { dayDate = shiftDate(dayDate, 1); renderTab('day'); });
    $('dayToday').addEventListener('click', () => { dayDate = new Date().toISOString().slice(0, 10); renderTab('day'); });
    $('dayPicker').addEventListener('change', event => { if (event.target.value) { dayDate = event.target.value; renderTab('day'); } });
    $('dayPrintBtn')?.addEventListener('click', () => window.print());
    $('tabWorkspace').querySelectorAll('[data-copy]').forEach(button => button.addEventListener('click', () => loadEntryIntoForm(Number(button.dataset.copy), false)));
    $('tabWorkspace').querySelectorAll('[data-edit]').forEach(button => button.addEventListener('click', () => loadEntryIntoForm(Number(button.dataset.edit), true)));
    $('tabWorkspace').querySelectorAll('[data-delete]').forEach(button => button.addEventListener('click', () => deleteEntry(Number(button.dataset.delete))));
  }
}

function shiftDate(value, days) {
  const date = new Date(`${value}T00:00:00`);
  date.setDate(date.getDate() + days);
  return date.toISOString().slice(0, 10);
}

function buildDayView() {
  const dayEntries = entries.filter(entry => entry.date === dayDate);
  const total = dayEntries.reduce((sum, entry) => sum + Number(entry.hours || 0), 0);
  const billable = dayEntries.filter(entry => String(entry.billable).toLowerCase() === 'yes').reduce((sum, entry) => sum + Number(entry.hours || 0), 0);
  const nonBillable = total - billable;
  const projectTotals = [...new Set(dayEntries.map(entry => entry.project).filter(Boolean))].map(name => ({ name, hours: dayEntries.filter(entry => entry.project === name).reduce((sum, entry) => sum + Number(entry.hours || 0), 0) })).sort((a, b) => b.hours - a.hours);
  const maxProjectHours = Math.max(...projectTotals.map(item => item.hours), 1);
  const projectBars = projectTotals.map(item => `<div class="day-project-row"><span>${esc(item.name)}</span><div class="day-bar-track"><i style="width:${item.hours / maxProjectHours * 100}%"></i></div><strong>${item.hours.toFixed(2)}</strong><small>${total ? Math.round(item.hours / total * 100) : 0}%</small></div>`).join('') || '<p class="empty">No entries logged on this day.</p>';
  const entryRows = dayEntries.map(entry => `<div class="day-entry-row"><div class="day-entry-main"><strong>${esc(entry.project)}</strong><span>${esc(entry.task || '')}</span><p>${esc(entry.notes || 'No notes')}</p><div class="day-tags"><em>${esc(entry.category || 'Other')}</em><em>${esc(entry.billable === 'Yes' ? 'Billable' : 'Non Billable')}</em><em>${esc(entry.nonBillableReason || '')}</em></div></div><strong class="day-entry-hours">${Number(entry.hours || 0).toFixed(2)}h</strong><div class="entry-row-actions"><button type="button" class="entry-action" data-copy="${entry.id}">Copy</button><button type="button" class="entry-action" data-edit="${entry.id}">Edit</button><button type="button" class="entry-action delete-action" data-delete="${entry.id}">Delete</button></div></div>`).join('') || '<p class="empty">No entries logged on this day.</p>';
  const dateLabel = new Date(`${dayDate}T00:00:00`).toLocaleDateString('en-GB', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });
  const printHead = `
    <div class="print-only print-header">
      <h2>Daily Timesheet — ${esc(dateLabel)}</h2>
      <p class="print-meta">${currentUser ? `${esc(currentUser.name || currentUser.username)} · ` : ''}${total.toFixed(2)} hours logged</p>
    </div>
  `;
  return `<div class="day-card">${printHead}<div class="day-toolbar"><button type="button" id="dayPrev">&#8592; Previous</button><strong>${esc(dateLabel)}</strong><button type="button" id="dayNext">Next &#8594;</button><button type="button" id="dayToday">Today</button><input id="dayPicker" class="date-input" type="date" value="${dayDate}"><span></span><button type="button" id="dayPrintBtn" class="btn-pdf-export" title="Export as PDF / Print">📄 Export as PDF</button></div><div class="day-kpis"><div><strong>${total.toFixed(2)}</strong><span>TOTAL HOURS</span></div><div><strong>${billable.toFixed(2)}</strong><span>BILLABLE</span></div><div><strong>${nonBillable.toFixed(2)}</strong><span>NON-BILLABLE</span></div><div><strong>${dayEntries.length}</strong><span>ENTRIES</span></div></div><h3 class="report-section-title">WHERE THE TIME WENT</h3><div class="day-projects">${projectBars}</div><h3 class="report-section-title">ENTRIES</h3><div class="day-entries">${entryRows}</div></div>`;
}

function getMonday(date) {
  const monday = new Date(date);
  monday.setHours(0, 0, 0, 0);
  monday.setDate(monday.getDate() - ((monday.getDay() + 6) % 7));
  return monday;
}

function shiftWeek(date, days) {
  const next = new Date(date);
  next.setDate(next.getDate() + days);
  return next;
}

function buildWeeklyReport() {
  if (!entries.length) return '<p class="empty">No entries logged yet.</p>';
  const weekStartDate = reportWeekStart || getMonday(new Date());
  const weekStart = isoDate(weekStartDate);
  const weekEndDate = shiftWeek(weekStartDate, 6);
  const weekEnd = isoDate(weekEndDate);
  const weekEntries = entries.filter(entry => entry.date >= weekStart && entry.date <= weekEnd);
  const weekNum = isoWeek(weekStartDate);
  const weekCode = `${weekNum.year}-W${String(weekNum.week).padStart(2, '0')}`;
  const rangeDisplay = `${formatDayMonth(weekStartDate)} – ${formatDayMonth(weekEndDate)} ${weekEndDate.getFullYear()}`;
  const printHead = `
    <div class="print-only print-header">
      <h2>Weekly Timesheet Report — ${esc(weekCode)}</h2>
      <p class="print-meta">${esc(rangeDisplay)}${currentUser ? ` · ${esc(currentUser.name || currentUser.username)}` : ''}</p>
    </div>
  `;
  if (!weekEntries.length) return `<div class="report-empty">${printHead}No entries for ${esc(rangeDisplay)}.</div>`;
  const total = weekEntries.reduce((sum, entry) => sum + Number(entry.hours || 0), 0);
  const billableHours = weekEntries.filter(entry => String(entry.billable).toLowerCase() === 'yes').reduce((sum, entry) => sum + Number(entry.hours || 0), 0);
  const nonBillableHours = total - billableHours;
  const dayTotals = new Map();
  weekEntries.forEach(entry => dayTotals.set(entry.date, (dayTotals.get(entry.date) || 0) + Number(entry.hours || 0)));
  const dayBars = [...Array(7)].map((_, index) => {
    const date = shiftWeek(weekStartDate, index);
    const key = date.toISOString().slice(0, 10);
    const hours = dayTotals.get(key) || 0;
    const label = date.toLocaleDateString('en-GB', { weekday: 'short', day: 'numeric', month: 'short' });
    return `<div class="day-bar-row"><span>${esc(label)}</span><div class="day-bar-track"><i style="width:${total ? hours / Math.max(...dayTotals.values(), 1) * 100 : 0}%"></i></div><strong>${hours.toFixed(2)}</strong></div>`;
  }).join('');
  const reasonTotals = [...new Set(weekEntries.map(entry => entry.nonBillableReason).filter(Boolean))].map(name => ({ name, hours: weekEntries.filter(entry => entry.nonBillableReason === name).reduce((sum, entry) => sum + Number(entry.hours || 0), 0) }));
  const reasonBars = reasonTotals.map(item => `<div class="reason-row"><span>${esc(item.name)}</span><div class="reason-track"><i style="width:${nonBillableHours ? item.hours / nonBillableHours * 100 : 0}%"></i></div><strong>${item.hours.toFixed(2)}</strong></div>`).join('') || '<div class="muted">No non-billable reasons.</div>';
  const weeks = new Map();
  weekEntries.forEach(entry => {
    const date = new Date(`${entry.date}T00:00:00`);
    const monday = new Date(date);
    monday.setDate(date.getDate() - ((date.getDay() + 6) % 7));
    const weekStart = monday.toISOString().slice(0, 10);
    if (!weeks.has(weekStart)) weeks.set(weekStart, new Map());
    const projects = weeks.get(weekStart);
    const projectName = entry.project || '(No project)';
    if (!projects.has(projectName)) projects.set(projectName, new Map());
    const tasks = projects.get(projectName);
    const taskName = entry.task || '(No task)';
    if (!tasks.has(taskName)) tasks.set(taskName, new Map());
    const dates = tasks.get(taskName);
    if (!dates.has(entry.date)) dates.set(entry.date, []);
    dates.get(entry.date).push(entry);
  });

  const reportRows = [...weeks.entries()].sort(([a], [b]) => b.localeCompare(a)).map(([weekStart, projects]) => {
    const weekEnd = new Date(`${weekStart}T00:00:00`);
    weekEnd.setDate(weekEnd.getDate() + 6);
    const weekEntries = entries.filter(entry => entry.date >= weekStart && entry.date <= weekEnd.toISOString().slice(0, 10));
    const weekHours = weekEntries.reduce((total, entry) => total + Number(entry.hours || 0), 0);
    const projectHtml = [...projects.entries()].map(([project, tasks]) => {
      const projectHours = [...tasks.values()].flatMap(dates => [...dates.values()].flat()).reduce((total, entry) => total + Number(entry.hours || 0), 0);
      const taskHtml = [...tasks.entries()].map(([task, dates]) => {
        const taskHours = [...dates.values()].flat().reduce((total, entry) => total + Number(entry.hours || 0), 0);
        const dateHtml = [...dates.entries()].sort(([a], [b]) => a.localeCompare(b)).map(([date, dayEntries]) => {
          const notes = dayEntries.filter(entry => entry.notes).map(entry => `<div class="report-note">${esc(entry.notes)}</div>`).join('');
          const hours = dayEntries.reduce((total, entry) => total + Number(entry.hours || 0), 0);
          return `<div class="report-row report-date"><span>${esc(new Date(`${date}T00:00:00`).toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' }))}</span><span>${notes || '<span class="muted">No notes</span>'}</span><strong>${hours.toFixed(2)}</strong></div>`;
        }).join('');
        return `<details class="report-task" open><summary><span>${esc(task)}</span><strong>${taskHours.toFixed(2)}</strong></summary>${dateHtml}</details>`;
      }).join('');
      return `<details class="report-project" open><summary><span>${esc(project)}</span><strong>${projectHours.toFixed(2)}</strong></summary>${taskHtml}</details>`;
    }).join('');
    const weekNumber = isoWeek(new Date(`${weekStart}T00:00:00`));
    return `<details class="report-week" open><summary><span>${weekNumber.year}-W${String(weekNumber.week).padStart(2, '0')}</span><strong>${weekHours.toFixed(2)}</strong></summary>${projectHtml}</details>`;
  }).join('');
  const byProject = [...new Set(weekEntries.map(entry => entry.project).filter(Boolean))].map(name => ({ name, hours: weekEntries.filter(entry => entry.project === name).reduce((sum, entry) => sum + Number(entry.hours || 0), 0), count: weekEntries.filter(entry => entry.project === name).length }));
  const byCategory = [...new Set(weekEntries.map(entry => entry.category).filter(Boolean))].map(name => ({ name, hours: weekEntries.filter(entry => entry.category === name).reduce((sum, entry) => sum + Number(entry.hours || 0), 0) }));
  const bars = (items, color) => items.sort((a, b) => b.hours - a.hours).map(item => `<div class="bar-row"><span>${esc(item.name)}</span><div class="bar-track"><i style="width:${total ? item.hours / total * 100 : 0}%;background:${color}"></i></div><strong>${item.hours.toFixed(2)}</strong><small>${total ? Math.round(item.hours / total * 100) : 0}%</small></div>`).join('');
  const weekLabel = `${isoWeek(weekStartDate).year}-W${String(isoWeek(weekStartDate).week).padStart(2, '0')}`;
  return `${printHead}<div class="report-kpis"><div><strong>${total.toFixed(2)}</strong><span>TOTAL HOURS</span></div><div><strong>${billableHours.toFixed(2)}</strong><span>BILLABLE</span></div><div><strong>${nonBillableHours.toFixed(2)}</strong><span>NON-BILLABLE</span></div><div><strong>${new Set(weekEntries.map(entry => entry.date)).size}</strong><span>DAYS LOGGED</span></div><div><strong>${weekEntries.length}</strong><span>ENTRIES</span></div></div><h3 class="report-section-title">BILLABLE SPLIT</h3><div class="split-bar"><i style="width:${total ? billableHours / total * 100 : 0}%"></i><b style="width:${total ? nonBillableHours / total * 100 : 0}%"></b></div><div class="split-labels"><span>Billable ${billableHours.toFixed(2)}h</span><span>Non-billable ${nonBillableHours.toFixed(2)}h</span></div><h3 class="report-section-title">NON-BILLABLE REASON</h3>${reasonBars}<h3 class="report-section-title">DAY BY DAY</h3><div class="day-bars">${dayBars}</div>${reportRows}<div class="report-total"><strong>Total</strong><strong>${total.toFixed(2)}</strong></div><h3 class="report-section-title">BY PROJECT</h3>${bars(byProject, '#8067c8')}<h3 class="report-section-title">BY CATEGORY</h3>${bars(byCategory, '#2862ae')}`;
}

async function loadManagementData(view) {
  if (view === 'tasks') {
    await loadTasks();
    const taskForm = $('taskForm');
    if (taskForm) {
      taskForm.addEventListener('submit', async e => {
        e.preventDefault();
        const input = $('taskTitle');
        const prioSel = $('taskPriority');
        const title = (input?.value || '').trim();
        const prio = +(prioSel?.value || 3);
        if (title) await addTask(title, prio);
      });
    }
    const btnShowDone = $('btnShowDone');
    if (btnShowDone) {
      btnShowDone.addEventListener('click', () => {
        showDone = !showDone;
        renderTasks();
      });
    }
  }
  if (view === 'reminders') {
    await loadReminders();
    const reminderForm = $('reminderForm');
    if (reminderForm) {
      const dueDateInput = $('reminderDueDate');
      if (dueDateInput && !dueDateInput.value) {
        dueDateInput.value = isoDate(new Date());
      }
      reminderForm.addEventListener('submit', async e => {
        e.preventDefault();
        const title = ($('reminderTitle')?.value || '').trim();
        const dueDate = $('reminderDueDate')?.value;
        const dueTime = $('reminderDueTime')?.value || null;
        const priority = +($('reminderPriority')?.value || 3);
        const notes = ($('reminderNotes')?.value || '').trim() || null;
        if (title && dueDate) {
          await addReminder(title, dueDate, dueTime, priority, notes);
          reminderForm.reset();
          if ($('reminderDueDate')) $('reminderDueDate').value = isoDate(new Date());
        }
      });
    }
    document.querySelectorAll('.reminder-filter-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        reminderFilter = btn.dataset.filter;
        document.querySelectorAll('.reminder-filter-btn').forEach(b => b.classList.toggle('is-active', b === btn));
        renderReminders();
      });
    });
    $('btnQuickAddReminder')?.addEventListener('click', () => {
      $('reminderTitle')?.focus();
    });
  }
  if (view === 'projects') {
    const data = await api('/api/projects');
    const projectStatuses = ['New', 'In-Progress', 'Hold', 'Complete'];
    $('projectList').innerHTML = (data.projects || []).map(project => `
      <div class="project-row">
        <strong>${esc(project.name)}</strong>
        <select class="project-status-select" data-project-id="${project.id}" aria-label="Status for ${esc(project.name)}">
          ${projectStatuses.map(status => `<option value="${status}"${status === project.status ? ' selected' : ''}>${status}</option>`).join('')}
        </select>
        <div class="task-list">${(project.tasks || []).map(task => `<span>${esc(task)}</span>`).join('') || '<em>No tasks</em>'}</div>
        <form class="task-add-form" data-project-id="${project.id}">
          <input name="task" placeholder="Add task" required>
          <button type="submit">Add task</button>
        </form>
      </div>
    `).join('') || '<p class="empty">No projects yet.</p>';
    $('projectForm').addEventListener('submit', async event => {
      event.preventDefault();
      await api('/api/projects', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name: $('newProject').value, status: $('projectStatus').value })
      });
      await loadProjectOptions();
      renderTab('projects');
    });
    document.querySelectorAll('.project-status-select').forEach(select => {
      select.addEventListener('change', async event => {
        const projectId = Number(event.target.dataset.projectId);
        const status = event.target.value;
        try {
          await api('/api/projects/status', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ projectId, status })
          });
          await loadProjectOptions();
        } catch (error) {
          alert(error.message || 'Failed to update project status.');
          renderTab('projects');
        }
      });
    });
    document.querySelectorAll('.task-add-form').forEach(form => form.addEventListener('submit', async event => {
      event.preventDefault();
      await api('/api/project-tasks', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ projectId: Number(form.dataset.projectId), name: form.elements.task.value })
      });
      await loadProjectOptions();
      renderTab('projects');
    }));
  }
  if (view === 'team') {
    const data = await api('/api/team');
    const teamPerson = document.querySelector('#tabWorkspace #teamPerson');
    const teamStats = document.querySelector('#tabWorkspace #teamStats');
    const teamList = document.querySelector('#tabWorkspace #teamList');
    const teamForm = document.querySelector('#tabWorkspace #teamForm');
    const newTeamForm = document.querySelector('#tabWorkspace #newTeamForm');
    const teamDate = document.querySelector('#tabWorkspace #teamDate');
    const teamWeek = document.querySelector('#tabWorkspace #teamWeek');
    const teamWeekPreview = document.querySelector('#tabWorkspace #teamWeekPreview');
    const teamCalendarBtn = document.querySelector('#tabWorkspace #teamCalendarBtn');
    const teamCalendarPopup = document.querySelector('#tabWorkspace #teamCalendarPopup');
    const teamWeekStrip = document.querySelector('#tabWorkspace #teamWeekStrip');
    const twsDays = document.querySelector('#tabWorkspace #twsDays');
    const twsPrev = document.querySelector('#tabWorkspace #twsPrev');
    const twsNext = document.querySelector('#tabWorkspace #twsNext');
    const teamYearSel = document.querySelector('#tabWorkspace #teamYear');
    const teamMessage = document.querySelector('#tabWorkspace #teamMessage');

    if (!teamPerson || !teamStats || !teamList || !teamForm || !newTeamForm) {
      $('tabWorkspace').innerHTML = `<div class="team-card"><p class="message">Team view could not be initialized. Please refresh the page.</p></div>`;
      return;
    }

    const members = data.members || [];
    teamPerson.innerHTML = '<option value="">Select a person...</option>' + members.map(member => `<option value="${member.id}">${esc(member.name)}</option>`).join('');
    const activity = data.activity || [];

    const currentYear = new Date().getFullYear();
    const recordedYears = activity.map(item => {
      const y = parseInt(item.week?.slice(0, 4), 10);
      return isNaN(y) ? currentYear : y;
    });
    const years = [...new Set([currentYear, ...recordedYears])].sort((a, b) => b - a);

    if (teamYearSel) {
      teamYearSel.innerHTML = years.map(y => `<option value="${y}"${y === currentYear ? ' selected' : ''}>${y}</option>`).join('');
      teamYearSel.addEventListener('change', () => {
        renderTeamActivityList();
      });
    }

    let calViewDate = new Date();

    function renderDaysStrip(details) {
      if (!twsDays || !details) return;
      const days = [];
      const monday = new Date(details.monday);
      const dayNames = ['MON', 'TUE', 'WED', 'THU', 'FRI'];
      for (let i = 0; i < 5; i++) {
        const d = new Date(monday.getFullYear(), monday.getMonth(), monday.getDate() + i);
        days.push({
          name: dayNames[i],
          dateStr: formatDayMonth(d),
          iso: isoDate(d)
        });
      }
      twsDays.innerHTML = days.map(d => `
        <div class="tws-day is-active" title="${d.name} ${d.dateStr} (${details.weekCode})">
          <span class="tws-dname">${d.name}</span>
          <span class="tws-dnum">${d.dateStr}</span>
        </div>
      `).join('');
    }

    function syncDateToWeek(dateVal) {
      if (!dateVal) {
        if (teamDate) teamDate.value = '';
        if (teamWeek) teamWeek.value = '';
        if (teamWeekPreview) teamWeekPreview.textContent = 'Pick date to select Mon–Fri week';
        if (twsDays) twsDays.innerHTML = '<span class="tws-empty">No week selected</span>';
        return;
      }
      const details = getWeekRange(dateVal);
      if (teamDate) teamDate.value = `${details.rangeText} (${details.weekCode})`;
      if (teamWeek) teamWeek.value = details.weekCode;
      if (teamWeekPreview) {
        teamWeekPreview.innerHTML = `<strong>${details.weekCode}</strong> · ${esc(details.rangeText)} <span class="team-days-hint">(Mon–Fri selected)</span>`;
      }
      renderDaysStrip(details);
    }

    let activeSelectedDate = isoDate(new Date());

    function renderCalendarPopup() {
      if (!teamCalendarPopup) return;
      const selectedDate = activeSelectedDate
        ? (([y, m, d]) => new Date(+y, +m - 1, +d))(activeSelectedDate.split('-'))
        : new Date();
      const selWeek = getWeekRange(selectedDate);
      const year = calViewDate.getFullYear();
      const month = calViewDate.getMonth();
      const monthName = calViewDate.toLocaleDateString('en-GB', { month: 'long', year: 'numeric' });

      const firstDay = new Date(year, month, 1);
      const lastDay = new Date(year, month + 1, 0);

      const startDate = new Date(firstDay);
      startDate.setDate(startDate.getDate() - ((startDate.getDay() + 6) % 7));

      const endDate = new Date(lastDay);
      const endDayOfWeek = (endDate.getDay() + 6) % 7;
      if (endDayOfWeek < 6) {
        endDate.setDate(endDate.getDate() + (6 - endDayOfWeek));
      }

      const todayIso = isoDate(new Date());

      let html = `
        <div class="cal-pop-head">
          <button type="button" class="cal-nav-btn" data-cal-nav="-1" aria-label="Previous month">&#8249;</button>
          <strong class="cal-month-title">${esc(monthName)}</strong>
          <button type="button" class="cal-nav-btn" data-cal-nav="1" aria-label="Next month">&#8250;</button>
          <button type="button" class="cal-close-btn" data-cal-close="1" title="Close">&times;</button>
        </div>
        <table class="cal-table">
          <thead>
            <tr>
              <th class="cal-th-wk" title="Week Number">Wk</th>
              <th>Mo</th><th>Tu</th><th>We</th><th>Th</th><th>Fr</th><th>Sa</th><th>Su</th>
            </tr>
          </thead>
          <tbody>
      `;

      let curr = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
      while (curr <= endDate) {
        const rowWeek = getWeekRange(curr);
        const isSelWeek = rowWeek.weekCode === selWeek.weekCode;
        const monIso = isoDate(curr);
        const wkNum = rowWeek.weekCode.split('-W')[1] || '';

        html += `<tr class="cal-week-row${isSelWeek ? ' is-selected-week' : ''}" data-week-code="${rowWeek.weekCode}" data-mon-date="${monIso}" title="Select Week ${rowWeek.weekCode} (${rowWeek.rangeText})">`;
        html += `<td class="cal-td-wk" title="Week ${rowWeek.weekCode}"><span class="cal-wk-pill">W${wkNum}</span></td>`;

        for (let day = 0; day < 7; day++) {
          const isoD = isoDate(curr);
          const isCurMonth = curr.getMonth() === month;
          const isToday = isoD === todayIso;
          const isPickedDate = isoD === activeSelectedDate;
          const isWorkday = day < 5;
          const isFirstWorkDay = day === 0;
          const isLastWorkDay = day === 4;
          const dayNum = curr.getDate();

          let cellClass = 'cal-day-cell';
          if (!isCurMonth) cellClass += ' is-other-month';
          if (isToday) cellClass += ' is-today';
          if (isWorkday) cellClass += ' is-workday';
          if (isFirstWorkDay) cellClass += ' is-week-start';
          if (isLastWorkDay) cellClass += ' is-week-end';
          if (isPickedDate && isWorkday) cellClass += ' is-picked-date';

          html += `<td class="${cellClass}" data-date="${isoD}">
            <span class="cal-day-num">${dayNum}</span>
          </td>`;
          curr.setDate(curr.getDate() + 1);
        }
        html += `</tr>`;
      }

      html += `</tbody></table><div class="cal-pop-footer"><button type="button" class="cal-today-link" data-cal-today="1">Today / Current Week</button></div>`;
      teamCalendarPopup.innerHTML = html;

      teamCalendarPopup.querySelectorAll('[data-cal-nav]').forEach(b => {
        b.addEventListener('click', e => {
          e.stopPropagation();
          calViewDate.setMonth(calViewDate.getMonth() + Number(b.dataset.calNav));
          renderCalendarPopup();
        });
      });

      teamCalendarPopup.querySelectorAll('[data-cal-today]').forEach(b => {
        b.addEventListener('click', e => {
          e.stopPropagation();
          const now = new Date();
          calViewDate = new Date(now);
          const nowIso = isoDate(now);
          activeSelectedDate = nowIso;
          syncDateToWeek(nowIso);
          renderCalendarPopup();
        });
      });

      teamCalendarPopup.querySelectorAll('[data-cal-close]').forEach(b => {
        b.addEventListener('click', e => {
          e.stopPropagation();
          teamCalendarPopup.hidden = true;
        });
      });

      teamCalendarPopup.querySelectorAll('.cal-day-cell').forEach(cell => {
        cell.addEventListener('click', e => {
          e.stopPropagation();
          const targetDate = cell.dataset.date;
          if (targetDate) {
            activeSelectedDate = targetDate;
            syncDateToWeek(targetDate);
            renderCalendarPopup();
            setTimeout(() => { if (teamCalendarPopup) teamCalendarPopup.hidden = true; }, 160);
          }
        });
      });

      teamCalendarPopup.querySelectorAll('.cal-week-row').forEach(row => {
        row.addEventListener('click', e => {
          e.stopPropagation();
          const monDate = row.dataset.monDate;
          if (monDate) {
            activeSelectedDate = monDate;
            syncDateToWeek(monDate);
            renderCalendarPopup();
            setTimeout(() => { if (teamCalendarPopup) teamCalendarPopup.hidden = true; }, 160);
          }
        });
      });
    }

    activeSelectedDate = isoDate(new Date());
    syncDateToWeek(activeSelectedDate);

    if (teamDate) {
      teamDate.addEventListener('click', e => {
        e.stopPropagation();
        teamCalendarPopup.hidden = false;
        calViewDate = activeSelectedDate
          ? (([y, m, d]) => new Date(+y, +m - 1, +d))(activeSelectedDate.split('-'))
          : new Date();
        renderCalendarPopup();
      });
    }

    if (teamCalendarBtn && teamCalendarPopup) {
      teamCalendarBtn.addEventListener('click', e => {
        e.stopPropagation();
        teamCalendarPopup.hidden = !teamCalendarPopup.hidden;
        if (!teamCalendarPopup.hidden) {
          calViewDate = activeSelectedDate
            ? (([y, m, d]) => new Date(+y, +m - 1, +d))(activeSelectedDate.split('-'))
            : new Date();
          renderCalendarPopup();
        }
      });
    }

    if (twsPrev) {
      twsPrev.addEventListener('click', () => {
        const curDate = activeSelectedDate
          ? (([y, m, d]) => new Date(+y, +m - 1, +d))(activeSelectedDate.split('-'))
          : new Date();
        curDate.setDate(curDate.getDate() - 7);
        const isoD = isoDate(curDate);
        activeSelectedDate = isoD;
        calViewDate = new Date(curDate);
        syncDateToWeek(isoD);
        if (!teamCalendarPopup.hidden) renderCalendarPopup();
      });
    }

    if (twsNext) {
      twsNext.addEventListener('click', () => {
        const curDate = activeSelectedDate
          ? (([y, m, d]) => new Date(+y, +m - 1, +d))(activeSelectedDate.split('-'))
          : new Date();
        curDate.setDate(curDate.getDate() + 7);
        const isoD = isoDate(curDate);
        activeSelectedDate = isoD;
        calViewDate = new Date(curDate);
        syncDateToWeek(isoD);
        if (!teamCalendarPopup.hidden) renderCalendarPopup();
      });
    }

    if (!window._teamCalDocClickBound) {
      window._teamCalDocClickBound = true;
      document.addEventListener('click', e => {
        const popup = document.querySelector('#tabWorkspace #teamCalendarPopup');
        const btn = document.querySelector('#tabWorkspace #teamCalendarBtn');
        const inp = document.querySelector('#tabWorkspace #teamDate');
        if (popup && !popup.hidden && !popup.contains(e.target) && e.target !== btn && e.target !== inp) {
          popup.hidden = true;
        }
      });
    }

    function renderTeamActivityList() {
      const year = parseInt(teamYearSel?.value, 10) || currentYear;
      const scopedActivity = activity.filter(item => (item.week || '').startsWith(String(year)));
      teamStats.textContent = `${scopedActivity.length} week${scopedActivity.length === 1 ? '' : 's'} logged in ${year}`;

      teamList.innerHTML = members.map(member => {
        const weeks = scopedActivity.filter(item => item.memberId === member.id);
        const count = weeks.length;
        const allowance = 4;
        const over = count > allowance;
        const pct = Math.min(count / allowance, 1) * 100;
        return `<div class="team-row">
          <div class="team-member-info">
            <strong>${esc(member.name)}</strong>
          </div>
          <div class="team-track"><i style="width:${pct}%"></i></div>
          <b>${count} of ${allowance}</b>
          <small>${over ? `${count - allowance} over` : `${Math.max(allowance - count, 0)} left`}</small>
          <div class="team-weeks">${weeks.map(item => {
            const range = weekToDateRange(item.week);
            return `<span class="team-week-chip" title="${esc(item.week)}: ${esc(range)}${item.notes ? ' · ' + esc(item.notes) : ''}">
              <span class="twc-code">${esc(item.week)}</span>
              <span class="twc-range">${esc(range)}</span>
              <button type="button" class="team-week-del" data-activity-id="${item.id}" title="Remove this week">&times;</button>
            </span>`;
          }).join('') || '<em>none booked</em>'}</div>
          <button type="button" class="team-delete" data-team-delete="${member.id}" title="Remove team member">Delete</button>
        </div>`;
      }).join('') || '<p class="empty">No team members yet. Add someone below.</p>';

      teamList.querySelectorAll('.team-week-del').forEach(btn => {
        btn.addEventListener('click', async e => {
          e.stopPropagation();
          const activityId = Number(btn.dataset.activityId);
          if (!confirm('Remove this booked week?')) return;
          try {
            await api('/api/team/activity/delete', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ activityId }) });
            renderTab('team');
          } catch (error) {
            if (teamMessage) {
              teamMessage.textContent = error.message;
              teamMessage.className = 'msg err';
            }
          }
        });
      });

      teamList.querySelectorAll('[data-team-delete]').forEach(button => {
        button.addEventListener('click', async () => {
          const member = members.find(item => item.id === Number(button.dataset.teamDelete));
          if (!member || !confirm(`Delete ${member.name} and all their booked activities?`)) return;
          try {
            await api('/api/team/delete', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ memberId: member.id }) });
            renderTab('team');
          } catch (error) {
            if (teamMessage) {
              teamMessage.textContent = error.message;
              teamMessage.className = 'msg err';
            }
          }
        });
      });
    }

    renderTeamActivityList();

    teamForm.addEventListener('submit', async event => {
      event.preventDefault();
      try {
        await api('/api/team/activity', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            memberId: Number(teamPerson.value),
            activity: $('teamActivity').value,
            week: teamWeek.value,
            notes: $('teamNotes').value
          })
        });
        renderTab('team');
      } catch (error) {
        if (teamMessage) {
          teamMessage.textContent = error.message;
          teamMessage.className = 'msg err';
        }
      }
    });

    newTeamForm.addEventListener('submit', async event => {
      event.preventDefault();
      try {
        await api('/api/team', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ name: $('newTeamMember').value })
        });
        renderTab('team');
      } catch (error) {
        if (teamMessage) {
          teamMessage.textContent = error.message;
          teamMessage.className = 'msg err';
        }
      }
    });
  }
  if (view === 'holidays') {
    const data = await api('/api/holidays');
    const today = new Date(); today.setHours(0, 0, 0, 0);
    const holidays = data.holidays || [];
    const years = [...new Set(holidays.map(item => item.date.slice(0, 4)))].join(', ');
    const upcoming = holidays.find(item => new Date(`${item.date}T00:00:00`) >= today);
    $('holidayYears').textContent = `${years} · ${holidays.filter(item => new Date(`${item.date}T00:00:00`) >= today).length} of ${holidays.length} left`;
    $('holidayNext').textContent = upcoming ? `Next: ${upcoming.name} · in ${Math.max(0, Math.ceil((new Date(`${upcoming.date}T00:00:00`) - today) / 86400000))} days` : '';
    $('holidayList').innerHTML = holidays.map(holiday => { const date = new Date(`${holiday.date}T00:00:00`); const past = date < today; const days = Math.ceil((date - today) / 86400000); return `<div class="holiday-row ${past ? 'is-past' : ''}"><strong>${esc(date.toLocaleDateString('en-GB', { day: 'numeric', month: 'short' }))}</strong><span>${esc(date.toLocaleDateString('en-GB', { weekday: 'long' }))}</span><b>${esc(holiday.name)}</b><small>${past ? 'past' : days === 0 ? 'today' : `in ${days} days`}</small><button type="button" title="Remove" disabled>×</button></div>`; }).join('') || '<p class="empty">No holidays yet.</p>';
    $('holidayForm').addEventListener('submit', async event => { event.preventDefault(); await api('/api/holidays', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ date: $('holidayDate').value, name: $('holidayName').value }) }); await loadSidebarHolidays(); renderTab('holidays'); });
  }
  if (view === 'sessions') {
    async function loadSessionsView() {
      const container = $('activeSessionsList');
      const countBadge = $('activeSessionsCount');
      if (!container) return;
      try {
        const res = await api('/api/sessions');
        const list = res.sessions || [];
        if (countBadge) countBadge.textContent = `${list.length} active`;
        container.innerHTML = list.length
          ? list.map(s => {
              const device = formatUserAgent(s.userAgent);
              const loginDate = s.createdAt ? new Date(s.createdAt).toLocaleString('en-GB', { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' }) : '—';
              const lastActive = s.lastActiveAt ? new Date(s.lastActiveAt).toLocaleString('en-GB', { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' }) : '—';
              return `
                <div class="session-card ${s.isCurrent ? 'is-current-session' : ''}">
                  <div class="session-card-header">
                    <div class="session-user-info">
                      <strong>${esc(s.name || s.username)} <span class="session-username">(${esc(s.username)})</span></strong>
                      ${s.isAdmin ? '<span class="session-admin-pill">Admin</span>' : ''}
                      ${s.isCurrent ? '<span class="session-current-pill">This Device</span>' : ''}
                    </div>
                    ${s.isCurrent
                      ? '<span class="session-status-badge is-active">Active Now</span>'
                      : `<button type="button" class="btn-revoke-session" data-revoke-session="${esc(s.id)}" data-session-user="${esc(s.username)}" title="Terminate session">End Session</button>`}
                  </div>
                  <div class="session-meta-grid">
                    <div class="session-meta-item">
                      <span class="sm-label">Device</span>
                      <span class="sm-val">${esc(device)}</span>
                    </div>
                    <div class="session-meta-item">
                      <span class="sm-label">IP Address</span>
                      <span class="sm-val">${esc(s.ipAddress)}</span>
                    </div>
                    <div class="session-meta-item">
                      <span class="sm-label">Signed In</span>
                      <span class="sm-val">${esc(loginDate)}</span>
                    </div>
                    <div class="session-meta-item">
                      <span class="sm-label">Last Active</span>
                      <span class="sm-val">${esc(lastActive)}</span>
                    </div>
                  </div>
                </div>
              `;
            }).join('')
          : '<p class="empty">No active sessions found.</p>';

        container.querySelectorAll('[data-revoke-session]').forEach(btn => {
          btn.addEventListener('click', async () => {
            const sid = btn.dataset.revokeSession;
            const u = btn.dataset.sessionUser;
            if (!confirm(`Revoke and terminate the active session for ${u}?`)) return;
            try {
              await api('/api/sessions/revoke', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ sessionId: sid })
              });
              await loadSessionsView();
            } catch (err) {
              const msg = $('sessionsMessage');
              if (msg) {
                msg.textContent = err.message;
                msg.className = 'msg err';
              }
            }
          });
        });
      } catch (err) {
        container.innerHTML = `<p class="empty">${esc(err.message)}</p>`;
      }
    }

    $('reloadSessionsBtn')?.addEventListener('click', loadSessionsView);
    await loadSessionsView();
  }
  if (view === 'users') {
    const data = await api('/api/users');
    $('resetUsername').innerHTML = data.users.map(user => `<option value="${esc(user.username)}">${esc(user.name)} (${esc(user.username)})</option>`).join('');
    $('codeUsername').innerHTML = $('resetUsername').innerHTML;
    $('emailUsername').innerHTML = data.users.map(user => `<option value="${esc(user.username)}" data-email="${esc(user.email || '')}">${esc(user.name)} (${esc(user.username)})</option>`).join('');
    const selectedEmailUser = data.users[0];
    $('editUserEmail').value = selectedEmailUser?.email || '';
    $('emailUsername').addEventListener('change', () => { const user = data.users.find(item => item.username === $('emailUsername').value); $('editUserEmail').value = user?.email || ''; });
    $('userList').innerHTML = data.users.map(user => `<div class="data-row"><strong>${esc(user.username)}</strong><span>${esc(user.name)} · ${esc(user.email || 'Email not set')}</span></div>`).join('') || '<p class="empty">No users yet.</p>';
    $('userForm').addEventListener('submit', async event => { event.preventDefault(); try { await api('/api/users', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ username: $('newUsername').value, name: $('newUserName').value, email: $('newUserEmail').value, password: $('newUserPassword').value }) }); $('userMessage').textContent = 'User created.'; renderTab('users'); } catch (error) { $('userMessage').textContent = error.message; } });
    $('emailForm').addEventListener('submit', async event => { event.preventDefault(); try { await api('/api/users/email', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ username: $('emailUsername').value, email: $('editUserEmail').value }) }); $('userMessage').textContent = 'Email address saved.'; renderTab('users'); } catch (error) { $('userMessage').textContent = error.message; } });
    $('resetUserForm').addEventListener('submit', async event => { event.preventDefault(); try { await api('/api/users/reset-password', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ username: $('resetUsername').value, password: $('resetPassword').value }) }); $('userMessage').textContent = 'Password reset. The user must choose a new password at next login.'; $('resetPassword').value = ''; } catch (error) { $('userMessage').textContent = error.message; } });
    $('resetCodeForm').addEventListener('submit', async event => { event.preventDefault(); try { const data = await api('/api/users/reset-code', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ username: $('codeUsername').value }) }); $('userMessage').textContent = `One-time reset code: ${data.code} (expires in 15 minutes).`; } catch (error) { $('userMessage').textContent = error.message; } });
  }
  if (view === 'options') {
    const data = await api('/api/options');
    const definitions = [['category', 'Category'], ['budget', 'Budget'], ['billable', 'Billable'], ['nonBillableReason', 'Non-billable Reason']];
    $('optionForms').innerHTML = definitions.map(([key, label]) => `<form class="option-form" data-list="${key}"><label>${label}<input name="value" required maxlength="120"></label><button type="submit">Add ${label}</button><div class="option-values">${(data.options?.[key] || []).map(value => `<span>${esc(value)}<button type="button" class="option-value-delete" data-list="${key}" data-value="${esc(value)}" title="Delete">×</button></span>`).join('') || '<em>No values yet.</em>'}</div></form>`).join('');
    document.querySelectorAll('.option-form').forEach(form => form.addEventListener('submit', async event => { event.preventDefault(); try { await api('/api/options', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ listName: form.dataset.list, value: form.elements.value.value }) }); $('optionMessage').textContent = 'Option saved.'; renderTab('options'); loadOptionCatalog(); } catch (error) { $('optionMessage').textContent = error.message; } }));
    document.querySelectorAll('.option-value-delete').forEach(button => button.addEventListener('click', async () => { const listName = button.dataset.list; const value = button.dataset.value; if (!confirm(`Delete ${value} from ${listName}? Existing entries will not change.`)) return; try { await api('/api/options/delete', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ listName, value }) }); $('optionMessage').textContent = 'Option deleted.'; renderTab('options'); loadOptionCatalog(); } catch (error) { $('optionMessage').textContent = error.message; } }));
  }
}

document.querySelectorAll('.tab').forEach(tab => tab.addEventListener('click', () => {
  document.querySelectorAll('.tab').forEach(item => item.classList.toggle('active', item === tab));
  const view = tab.dataset.view;
  const viewTitles = {
    tasks: "Today's Tasks",
    reminders: "Reminders",
    entry: "Timesheet entry",
    day: "Daily view",
    report: "Weekly report",
    projects: "Projects",
    team: "Team activities",
    holidays: "Holidays",
    sessions: "Active Sessions",
    users: "User Management",
    options: "Dropdown options"
  };
  $('viewTitle').textContent = viewTitles[view] || tab.textContent.trim();
  $('viewMessage').textContent = view === 'entry' ? 'The Cloudflare workspace is using D1 for its data.' : '';
  if (view === 'entry') {
    $('entryWorkspace').hidden = false;
    $('tabWorkspace').hidden = true;
  } else {
    $('entryWorkspace').hidden = true;
    $('tabWorkspace').hidden = false;
    renderTab(view).catch(error => { $('tabWorkspace').innerHTML = `<div class="tab-card"><p class="message">${esc(error.message)}</p></div>`; });
  }
}));

$('loginForm').addEventListener('submit', async event => {
  event.preventDefault();
  showMessage('loginMessage', 'Signing in...', false);
  try {
    const data = await api('/api/login', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ username: $('username').value, password: $('password').value }) });
    showApp(data.user);
  } catch (error) { showMessage('loginMessage', error.message); }
});

$('forgotPassword').addEventListener('click', () => {
  $('loginForm').hidden = true;
  $('registerForm').hidden = true;
  $('forgotForm').hidden = false;
  $('forgotUsername').value = $('username').value;
  $('forgotUsername').focus();
});
$('registerLink').addEventListener('click', () => { $('loginForm').hidden = true; $('forgotForm').hidden = true; $('registerForm').hidden = false; $('registerUsername').focus(); });
$('cancelForgot').addEventListener('click', () => {
  $('forgotForm').hidden = true;
  $('loginForm').hidden = false;
  $('forgotMessage').textContent = '';
});
$('cancelRegister').addEventListener('click', () => { $('registerForm').hidden = true; $('loginForm').hidden = false; $('registerMessage').textContent = ''; });
$('registerForm').addEventListener('submit', event => {
  event.preventDefault();
  api('/api/register', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ username: $('registerUsername').value, name: $('registerName').value, email: $('registerEmail').value, password: $('registerPassword').value, recoveryPassword: $('registerRecovery').value }) }).then(data => { $('registerMessage').textContent = data.message; $('registerForm').reset(); }).catch(error => { $('registerMessage').textContent = error.message; });
});
$('forgotForm').addEventListener('submit', event => {
  event.preventDefault();
  if ($('forgotNewPassword').value !== $('forgotConfirmPassword').value) { $('forgotMessage').textContent = 'New passwords do not match.'; return; }
  api('/api/password/recover', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ username: $('forgotUsername').value, email: $('forgotEmail').value, recoveryPassword: $('forgotRecovery').value, newPassword: $('forgotNewPassword').value }) }).then(data => { $('forgotMessage').textContent = data.message; $('forgotForm').reset(); }).catch(error => { $('forgotMessage').textContent = error.message; });
});

$('logout').addEventListener('click', async () => {
  await api('/api/logout', { method: 'POST' }).catch(() => {});
  showLogin();
});

$('changePassword').addEventListener('click', () => {
  $('passwordPanel').hidden = false;
  $('passwordMessage').textContent = '';
  $('passwordForm').reset();
  $('currentPassword').focus();
});
$('cancelPassword').addEventListener('click', () => { $('passwordPanel').hidden = true; });
$('passwordForm').addEventListener('submit', async event => {
  event.preventDefault();
  const newPassword = $('newPassword').value;
  if (newPassword !== $('confirmPassword').value) { $('passwordMessage').textContent = 'New passwords do not match.'; return; }
  try {
    await api('/api/password/change', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ currentPassword: $('currentPassword').value, newPassword }) });
    $('passwordMessage').textContent = 'Password changed successfully.';
    $('passwordForm').reset();
  } catch (error) { $('passwordMessage').textContent = error.message; }
});

$('entryForm').addEventListener('submit', async event => {
  event.preventDefault();
  await saveEntry(false);
});
$('reloadEntries').addEventListener('click', loadEntries);
$('project').addEventListener('input', updateTaskOptions);

async function saveEntry(keepValues) {
  showMessage('entryMessage', 'Saving...', false);
  const body = { date: $('date').value, project: $('project').value, task: $('task').value, hours: Number($('hours').value), category: $('category').value, budget: $('budget').value, billable: $('billable').value, nonBillableReason: $('nonBillableReason').value, notes: $('notes').value, ticket: $('ticket').value, incidentType: $('incidentType').value };
  try {
    const endpoint = editingEntryId ? '/api/entries/update' : '/api/entries';
    const payload = editingEntryId ? { ...body, id: editingEntryId } : body;
    await api(endpoint, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
    if (editingEntryId) cancelEntryEdit();
    else if (keepValues) { $('hours').value = ''; $('notes').value = ''; $('ticket').value = ''; }
    else $('entryForm').reset();
    showMessage('entryMessage', 'Entry saved.', false);
    loadEntries();
  } catch (error) { showMessage('entryMessage', error.message); }
}

$('addKeep').addEventListener('click', () => saveEntry(true));
$('cancelEdit').addEventListener('click', cancelEntryEdit);
document.querySelectorAll('[data-hours]').forEach(button => button.addEventListener('click', () => {
  $('hours').value = button.dataset.hours;
}));

/* Rough Work / Scratchpad Logic */
function initScratchpad() {
  const pad = $('scratchpadText');
  const status = $('scratchpadStatus');
  const copyBtn = $('copyToNotesBtn');
  const clearBtn = $('clearScratchpadBtn');
  if (!pad) return;

  const storageKey = `timesheet_scratchpad_${currentUser?.username || 'user'}`;
  try {
    const saved = localStorage.getItem(storageKey);
    if (saved) pad.value = saved;
  } catch (_) {}

  let timeoutId = null;
  pad.addEventListener('input', () => {
    if (status) status.textContent = 'Saving...';
    clearTimeout(timeoutId);
    timeoutId = setTimeout(() => {
      try {
        localStorage.setItem(storageKey, pad.value);
        if (status) status.textContent = 'Saved';
      } catch (_) {}
    }, 300);
  });

  if (copyBtn) {
    copyBtn.addEventListener('click', () => {
      const text = pad.value.trim();
      if (!text) return;
      const notesField = $('notes');
      if (notesField) {
        notesField.value = notesField.value ? `${notesField.value}\n${text}` : text;
        notesField.focus();
        notesField.scrollIntoView({ behavior: 'smooth', block: 'center' });
      }
    });
  }

  if (clearBtn) {
    clearBtn.addEventListener('click', () => {
      if (!pad.value.trim() || confirm('Clear rough work text?')) {
        pad.value = '';
        try { localStorage.removeItem(storageKey); } catch (_) {}
        if (status) status.textContent = 'Cleared';
      }
    });
  }
}

function loadEntryIntoForm(id, edit) {
  const entry = entries.find(item => item.id === id);
  if (!entry) return;
  const activeTab = document.querySelector('.tab.active');
  if (activeTab?.dataset.view !== 'entry') document.querySelector('.tab[data-view="entry"]')?.click();
  $('date').value = entry.date || '';
  if (entry.project && !Array.from($('project').options).some(opt => opt.value === entry.project)) {
    const opt = document.createElement('option');
    opt.value = entry.project;
    opt.textContent = entry.project;
    $('project').appendChild(opt);
  }
  $('project').value = entry.project || '';
  updateTaskOptions();
  $('task').value = entry.task || '';
  $('category').value = entry.category || '';
  $('budget').value = entry.budget || '';
  $('billable').value = entry.billable || '';
  $('nonBillableReason').value = entry.nonBillableReason || '';
  $('hours').value = entry.hours || '';
  $('notes').value = entry.notes || '';
  $('ticket').value = entry.ticket || '';
  $('incidentType').value = entry.incidentType || '';
  editingEntryId = edit ? id : null;
  $('saveEntry').textContent = edit ? 'Save changes' : 'Add entry';
  $('cancelEdit').hidden = !edit;
  $('entryForm').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function cancelEntryEdit() {
  editingEntryId = null;
  $('entryForm').reset();
  $('saveEntry').textContent = 'Add entry';
  $('cancelEdit').hidden = true;
}

async function deleteEntry(id) {
  const entry = entries.find(item => item.id === id);
  const label = entry ? `${entry.project} on ${entry.date}` : `entry ${id}`;
  if (!confirm(`Delete ${label}? This cannot be undone.`)) return;
  try {
    await api('/api/entries/delete', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ id }) });
    if (editingEntryId === id) cancelEntryEdit();
    await loadEntries();
    if (document.querySelector('.tab.active')?.dataset.view === 'day') renderTab('day');
  } catch (error) { showMessage('entryMessage', error.message); }
}

/* ------------------------------------------------------- reminder management */

function getReminderDueStatus(rem) {
  if (rem.isCompleted) return { label: 'Done', cls: 'is-done', diffDays: 0 };
  const today = isoDate(new Date());
  if (rem.dueDate === today) return { label: 'Today', cls: 'is-today', diffDays: 0 };
  const todayDate = new Date(`${today}T00:00:00`);
  const dueDate = new Date(`${rem.dueDate}T00:00:00`);
  const diffDays = Math.round((dueDate - todayDate) / 86400000);
  if (diffDays < 0) return { label: `${Math.abs(diffDays)}d overdue`, cls: 'is-overdue', diffDays };
  if (diffDays === 1) return { label: 'Tomorrow', cls: 'is-soon', diffDays };
  return { label: `in ${diffDays}d`, cls: 'is-upcoming', diffDays };
}

async function loadReminders() {
  try {
    const data = await api('/api/reminders');
    reminders = data.reminders || [];
    renderReminders();
    renderSidebarReminders();
  } catch (err) {
    console.error('Could not load reminders:', err);
  }
}

function renderSidebarReminders() {
  const container = $('sidebarReminderList');
  if (!container) return;
  const activeReminders = reminders.filter(r => !r.isCompleted);
  const countBadge = $('sidebarReminderCount');
  if (countBadge) countBadge.textContent = activeReminders.length ? String(activeReminders.length) : '';

  const addBtn = $('sidebarAddReminderBtn');
  if (addBtn && !addBtn._bound) {
    addBtn._bound = true;
    addBtn.addEventListener('click', e => {
      e.stopPropagation();
      document.querySelector('.tab[data-view="reminders"]')?.click();
      setTimeout(() => $('reminderTitle')?.focus(), 100);
    });
  }

  container.innerHTML = activeReminders.length
    ? activeReminders.slice(0, 5).map(rem => {
        const status = getReminderDueStatus(rem);
        const timeStr = rem.dueTime ? ` @ ${rem.dueTime}` : '';
        return `
          <div class="sidebar-reminder-item p${rem.priority || 3}-edge ${status.cls}" data-reminder-id="${rem.id}" title="${esc(rem.title)} · Due ${esc(rem.dueDate)}${timeStr}">
            <button type="button" class="sbr-check" data-toggle-reminder="${rem.id}" title="Mark as completed">&#10003;</button>
            <div class="sbr-body">
              <strong>${esc(rem.title)}</strong>
              <span class="sbr-due ${status.cls}">${esc(status.label)}${timeStr}</span>
            </div>
            <span class="prio-badge p${rem.priority || 3}">P${rem.priority || 3}</span>
          </div>
        `;
      }).join('')
    : '<p class="sidebar-empty">No active reminders. <button type="button" class="sbr-add-link" id="sbrAddLink">+ Add one</button></p>';

  container.querySelectorAll('[data-toggle-reminder]').forEach(btn => {
    btn.addEventListener('click', async e => {
      e.stopPropagation();
      await toggleReminder(Number(btn.dataset.toggleReminder));
    });
  });

  container.querySelectorAll('.sidebar-reminder-item').forEach(item => {
    item.addEventListener('click', () => {
      document.querySelector('.tab[data-view="reminders"]')?.click();
    });
  });

  container.querySelector('#sbrAddLink')?.addEventListener('click', e => {
    e.preventDefault();
    document.querySelector('.tab[data-view="reminders"]')?.click();
    setTimeout(() => $('reminderTitle')?.focus(), 100);
  });
}

function renderReminders() {
  const list = $('reminderList');
  if (!list) return;

  const today = isoDate(new Date());
  const activeCount = reminders.filter(r => !r.isCompleted).length;
  const todayCount = reminders.filter(r => !r.isCompleted && r.dueDate === today).length;
  const overdueCount = reminders.filter(r => !r.isCompleted && r.dueDate < today).length;
  const statsEl = $('reminderStats');
  if (statsEl) {
    statsEl.textContent = `${activeCount} active${todayCount ? ` · ${todayCount} due today` : ''}${overdueCount ? ` · ${overdueCount} overdue` : ''}`;
  }

  let filtered = [...reminders];
  if (reminderFilter === 'today') {
    filtered = filtered.filter(r => !r.isCompleted && r.dueDate === today);
  } else if (reminderFilter === 'upcoming') {
    filtered = filtered.filter(r => !r.isCompleted && r.dueDate > today);
  } else if (reminderFilter === 'completed') {
    filtered = filtered.filter(r => r.isCompleted);
  } else {
    filtered = filtered.filter(r => !r.isCompleted);
  }

  list.innerHTML = filtered.length
    ? filtered.map(rem => {
        const status = getReminderDueStatus(rem);
        const timeStr = rem.dueTime ? ` @ ${rem.dueTime}` : '';
        const dateFormatted = new Date(`${rem.dueDate}T00:00:00`).toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' });
        return `
          <div class="reminder-item p${rem.priority || 3}-edge ${rem.isCompleted ? 'is-completed' : ''}" data-reminder-id="${rem.id}">
            <button type="button" class="reminder-checkbox ${rem.isCompleted ? 'checked' : ''}" data-toggle="${rem.id}" title="${rem.isCompleted ? 'Mark incomplete' : 'Mark complete'}">
              ${rem.isCompleted ? '&#10003;' : ''}
            </button>
            <div class="reminder-main">
              <div class="reminder-title-row">
                <strong class="reminder-title">${esc(rem.title)}</strong>
                <span class="prio-badge p${rem.priority || 3}">P${rem.priority || 3}</span>
                <span class="reminder-due-tag ${status.cls}">${esc(status.label)} (${esc(dateFormatted)}${timeStr})</span>
              </div>
              ${rem.notes ? `<p class="reminder-notes">${esc(rem.notes)}</p>` : ''}
            </div>
            <div class="reminder-actions">
              <button type="button" class="reminder-del-btn" data-delete-reminder="${rem.id}" title="Delete reminder">&times;</button>
            </div>
          </div>
        `;
      }).join('')
    : `<p class="empty">${reminderFilter === 'completed' ? 'No completed reminders.' : 'No reminders found in this view.'}</p>`;

  list.querySelectorAll('[data-toggle]').forEach(btn => {
    btn.addEventListener('click', async () => {
      await toggleReminder(Number(btn.dataset.toggle));
    });
  });

  list.querySelectorAll('[data-delete-reminder]').forEach(btn => {
    btn.addEventListener('click', async () => {
      const id = Number(btn.dataset.deleteReminder);
      if (!confirm('Delete this reminder?')) return;
      await deleteReminder(id);
    });
  });
}

async function addReminder(title, dueDate, dueTime, priority, notes) {
  try {
    await api('/api/reminders', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ title, dueDate, dueTime, priority, notes })
    });
    showMessage('reminderMsg', 'Reminder added.', false);
    await loadReminders();
  } catch (err) {
    showMessage('reminderMsg', err.message, true);
  }
}

async function toggleReminder(id) {
  try {
    await api('/api/reminders/toggle', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ id })
    });
    await loadReminders();
  } catch (err) {
    console.error('Failed to toggle reminder:', err);
  }
}

async function deleteReminder(id) {
  try {
    await api('/api/reminders/delete', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ id })
    });
    await loadReminders();
  } catch (err) {
    console.error('Failed to delete reminder:', err);
  }
}

/* ------------------------------------------------------- task management */

function taskDay(t) {
  if (t.completedAt) {
    const d = new Date(t.completedAt);
    if (!isNaN(d)) return d.toISOString().slice(0, 10);
  }
  return t.date;
}

function renderSidebarTasks() {
  const container = $('sidebarTaskList');
  if (!container) return;
  const inProgress = Object.values(tasks)
    .filter(t => t.status === 'In-Progress')
    .sort(byPriority);

  const countBadge = $('sidebarTaskCount');
  if (countBadge) countBadge.textContent = inProgress.length ? String(inProgress.length) : '';

  container.innerHTML = inProgress.length
    ? inProgress.map(task => `
        <div class="sidebar-task p${task.priority || 3}-edge" data-task-id="${esc(task.id)}" title="Priority P${task.priority || 3} · Click to open Today's Tasks">
          <span class="prio-badge p${task.priority || 3}">P${task.priority || 3}</span>
          <strong>${esc(task.title)}</strong>
        </div>
      `).join('')
    : '<p class="sidebar-empty">No active tasks in progress.</p>';

  container.querySelectorAll('.sidebar-task').forEach(card => {
    card.addEventListener('click', () => {
      document.querySelector('.tab[data-view="tasks"]')?.click();
    });
  });
}

async function loadTasks() {
  try {
    const data = await api('/api/tasks');
    tasks = data.tasks || {};
    renderTasks();
    renderSidebarTasks();
  } catch (err) {
    setTaskMsg('Could not load tasks: ' + err.message, 'err');
  }
}

const seqRank = t => (t.seq || 0) || Infinity;
const byPriority = (a, b) =>
  (a.priority || 3) - (b.priority || 3) ||
  seqRank(a) - seqRank(b) ||
  new Date(a.createdAt || 0) - new Date(b.createdAt || 0);

function nextSeq(priority) {
  const used = Object.values(tasks)
    .filter(t => (t.priority || 3) === priority && t.status !== 'Complete')
    .map(t => t.seq || 0);
  return Math.max(0, ...used) + 1;
}

function renderTasks() {
  renderSidebarTasks();
  const board = $('tasksBoard');
  if (!board) return;
  const today = new Date().toISOString().slice(0, 10);
  const all = Object.values(tasks);
  const earlierDone = all.filter(t => t.status === 'Complete' && taskDay(t) !== today);
  const todays = all.filter(t => t.status !== 'Complete' || taskDay(t) === today || showDone);

  const dateEl = $('taskDate');
  if (dateEl) dateEl.textContent = new Date().toLocaleDateString('en-GB', { day: 'numeric', month: 'short' });

  const toggle = $('btnShowDone');
  if (toggle) {
    toggle.hidden = earlierDone.length === 0;
    toggle.textContent = showDone ? 'Hide earlier completed'
                                  : `Show earlier completed (${earlierDone.length})`;
    toggle.classList.toggle('is-on', showDone);
  }

  const done = todays.filter(t => t.status === 'Complete').length;
  const stats = $('boardStats');
  if (stats) {
    stats.textContent = todays.length ? `${done} of ${todays.length} done` : '0 of 0 done';
  }

  board.innerHTML = TASK_STATUSES.map(status => {
    const items = todays.filter(t => t.status === status).sort(byPriority);
    return `<section class="swimlane lane-${statusKey(status)}" data-status="${status}">
      <div class="lane-head">
        <span class="lane-dot"></span>
        <h3>${status.toUpperCase()}</h3>
        <span class="lane-count">${items.length}</span>
      </div>
      <div class="lane-body" data-status="${status}">
        ${items.map(taskHtml).join('')}
        <div class="lane-empty${items.length ? ' is-hidden' : ''}">Drop a card here</div>
      </div>
    </section>`;
  }).join('');

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

  const editor = board.querySelector('.task-edit-input');
  if (editor) {
    editor.addEventListener('keydown', e => {
      e.stopPropagation();
      if (e.key === 'Enter') { e.preventDefault(); saveTaskTitle(editor.dataset.taskId); }
      if (e.key === 'Escape') { e.preventDefault(); cancelTaskEdit(); }
    });
    editor.focus();
    editor.select();
  }

  wireDragDrop(board);
}

function createdStamp(task) {
  const d = task.createdAt ? new Date(task.createdAt)
          : task.date ? new Date(`${task.date}T00:00:00`)
          : null;
  if (!d || isNaN(d)) return null;
  return {
    text: task.createdAt
      ? d.toLocaleString('en-GB', { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' })
      : d.toLocaleDateString('en-GB', { day: 'numeric', month: 'short' }),
    title: `Created ${d.toLocaleString()}`
  };
}

function completedStamp(task) {
  if (task.status !== 'Complete' || !task.completedAt) return null;
  const d = new Date(task.completedAt);
  if (isNaN(d)) return null;
  const today = new Date().toISOString().slice(0, 10);
  const sameDay = d.toISOString().slice(0, 10) === today;
  return {
    text: sameDay
      ? d.toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' })
      : d.toLocaleString('en-GB', { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' }),
    title: `Completed ${d.toLocaleString()}`
  };
}

function taskHtml(task) {
  const prio = task.priority || 3;
  const idx = TASK_STATUSES.indexOf(task.status);
  const prev = TASK_STATUSES[idx - 1];
  const next = TASK_STATUSES[idx + 1];
  const today = new Date().toISOString().slice(0, 10);
  const stale = task.date && task.date !== today && task.status !== 'Complete';
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
  return api('/api/events', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ ...payload, timestamp: new Date().toISOString() })
  });
}

async function addTask(title, priority) {
  const taskId = 'task-' + Date.now();
  const date = new Date().toISOString().slice(0, 10);
  const seq = nextSeq(priority);
  try {
    await logTaskEvent({ type: 'task:created', taskId, title, status: 'New', priority, seq, date });
    tasks[taskId] = { id: taskId, title, status: 'New', priority, seq, date, createdAt: new Date().toISOString(), completedAt: null };
    const taskTitleInput = $('taskTitle');
    if (taskTitleInput) taskTitleInput.value = '';
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
  const input = $('tasksBoard')?.querySelector('.task-edit-input');
  if (!task || !input) return;

  const newTitle = input.value.trim();
  if (!newTitle) return setTaskMsg('A task needs a title.', 'err');
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
  if (!el) return;
  el.textContent = text;
  el.className = 'msg' + (kind ? ' ' + kind : '');
  if (kind === 'ok') setTimeout(() => { if (el.textContent === text) el.textContent = ''; }, 3000);
}

api('/api/session').then(data => showApp(data.user)).catch(showLogin);
