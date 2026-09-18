const SESSION_DAYS = 7;
const PASSWORD_ITERATIONS = 100000;

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    const cors = { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' };

    try {
      if (request.method === 'OPTIONS') return new Response(null, { headers: cors });
      if (url.pathname === '/api/login' && request.method === 'POST') return await login(request, env, cors);
      if (url.pathname === '/api/register' && request.method === 'POST') return await register(request, env, cors);
      if (url.pathname === '/api/logout' && request.method === 'POST') return await logout(request, env, cors);
      if (url.pathname === '/api/session' && request.method === 'GET') return await session(request, env, cors);
      if (url.pathname === '/api/password/change' && request.method === 'POST') return await changePassword(request, env, cors);
      if (url.pathname === '/api/password/forgot' && request.method === 'POST') return await forgotPassword(request, env, cors);
      if (url.pathname === '/api/password/recover' && request.method === 'POST') return await recoverWithSecondaryPassword(request, env, cors);
      if (url.pathname === '/api/users' && request.method === 'GET') return await listUsers(request, env, cors);
      if (url.pathname === '/api/users' && request.method === 'POST') return await createUser(request, env, cors);
      if (url.pathname === '/api/users/email' && request.method === 'POST') return await updateUserEmail(request, env, cors);
      if (url.pathname === '/api/users/reset-password' && request.method === 'POST') return await resetPassword(request, env, cors);
      if (url.pathname === '/api/users/reset-code' && request.method === 'POST') return await createResetCode(request, env, cors);
      if (url.pathname === '/api/sessions' && request.method === 'GET') return await listActiveSessions(request, env, cors);
      if (url.pathname === '/api/sessions/revoke' && request.method === 'POST') return await revokeSession(request, env, cors);
      if (url.pathname === '/api/projects' && request.method === 'GET') return await listProjects(request, env, cors);
      if (url.pathname === '/api/projects' && request.method === 'POST') return await addProject(request, env, cors);
      if (url.pathname === '/api/projects/status' && request.method === 'POST') return await updateProjectStatus(request, env, cors);
      if (url.pathname === '/api/project-tasks' && request.method === 'POST') return await addProjectTask(request, env, cors);
      if (url.pathname === '/api/team' && request.method === 'GET') return await listTeam(request, env, cors);
      if (url.pathname === '/api/team' && request.method === 'POST') return await addTeamMember(request, env, cors);
      if (url.pathname === '/api/team/activity' && request.method === 'POST') return await addTeamActivity(request, env, cors);
      if (url.pathname === '/api/team/activity/delete' && request.method === 'POST') return await deleteTeamActivity(request, env, cors);
      if (url.pathname === '/api/team/delete' && request.method === 'POST') return await deleteTeamMember(request, env, cors);
      if (url.pathname === '/api/holidays' && request.method === 'GET') return await listHolidays(request, env, cors);
      if (url.pathname === '/api/holidays' && request.method === 'POST') return await addHoliday(request, env, cors);
      if (url.pathname === '/api/options' && request.method === 'GET') return await listOptions(request, env, cors);
      if (url.pathname === '/api/options' && request.method === 'POST') return await addOption(request, env, cors);
      if (url.pathname === '/api/options/delete' && request.method === 'POST') return await deleteOption(request, env, cors);
      if (url.pathname === '/api/tasks' && request.method === 'GET') return await listDailyTasks(request, env, cors);
      if (url.pathname === '/api/events' && request.method === 'POST') return await handleTaskEvent(request, env, cors);
      if (url.pathname === '/api/reminders' && request.method === 'GET') return await listReminders(request, env, cors);
      if (url.pathname === '/api/reminders' && request.method === 'POST') return await addReminder(request, env, cors);
      if (url.pathname === '/api/reminders/toggle' && request.method === 'POST') return await toggleReminder(request, env, cors);
      if (url.pathname === '/api/reminders/delete' && request.method === 'POST') return await deleteReminder(request, env, cors);
      if (url.pathname === '/api/entries' && request.method === 'GET') return await listEntries(request, env, cors);
      if (url.pathname === '/api/entries' && request.method === 'POST') return await addEntry(request, env, cors);
      if (url.pathname === '/api/entries/update' && request.method === 'POST') return await updateEntry(request, env, cors);
      if (url.pathname === '/api/entries/delete' && request.method === 'POST') return await deleteEntry(request, env, cors);
      if (env.ASSETS) return env.ASSETS.fetch(request);
      return json({ error: 'Not found' }, 404, cors);
    } catch (error) {
      const message = error.message || 'Request failed';
      const status = /authentication required|session expired|invalid session/i.test(message) ? 401 : 400;
      return json({ error: message }, status, cors);
    }
  }
};

async function login(request, env, headers) {
  const body = await request.json();
  const username = String(body.username || '').trim();
  const password = String(body.password || '');
  const user = await env.DB.prepare('SELECT * FROM users WHERE username = ? COLLATE NOCASE').bind(username).first();
  if (!user || !(await verifyPassword(password, user.password_hash, user.password_salt, user.password_iterations))) {
    return json({ error: 'Invalid username or password.' }, 401, headers);
  }

  const token = randomToken(32);
  const expires = new Date(Date.now() + SESSION_DAYS * 86400000).toISOString();
  const ip = request.headers.get('CF-Connecting-IP') || request.headers.get('x-forwarded-for') || 'Unknown';
  const userAgent = request.headers.get('user-agent') || 'Unknown';

  await env.DB.prepare(`
    INSERT INTO sessions (id, user_id, ip_address, user_agent, last_active_at, expires_at)
    VALUES (?, ?, ?, ?, strftime('%Y-%m-%dT%H:%M:%fZ', 'now'), ?)
  `).bind(token, user.id, ip, userAgent, expires).run();

  return new Response(JSON.stringify({ ok: true, user: publicUser(user) }), {
    status: 200,
    headers: { ...headers, 'Set-Cookie': cookie(token, SESSION_DAYS * 86400) }
  });
}

async function register(request, env, headers) {
  const body = await request.json();
  const username = String(body.username || '').trim().toUpperCase();
  const name = String(body.name || username).trim();
  const email = String(body.email || '').trim().toLowerCase();
  const password = String(body.password || '');
  const recoveryPassword = String(body.recoveryPassword || '');
  if (!/^[A-Z0-9._-]{2,40}$/.test(username)) throw new Error('Username must be 2-40 letters, numbers, dots, hyphens, or underscores.');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('A valid email address is required.');
  if (password.length < 8 || recoveryPassword.length < 8) throw new Error('Both passwords must be at least 8 characters.');
  if (password === recoveryPassword) throw new Error('Use different primary and secondary passwords.');
  if (name.toLowerCase() === password.toLowerCase() || name.toLowerCase() === recoveryPassword.toLowerCase()) throw new Error('Display name cannot be a password.');
  const primary = await hashPassword(password);
  const recovery = await hashPassword(recoveryPassword);
  await env.DB.prepare(`INSERT INTO users (username, display_name, email, password_hash, password_salt, password_iterations, recovery_hash, recovery_salt, recovery_iterations, must_change_password, is_admin)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 0, 0)`).bind(username, name, email, primary.hash, primary.salt, primary.iterations, recovery.hash, recovery.salt, recovery.iterations).run();
  return json({ ok: true, message: 'Registration successful. You can now sign in.' }, 201, headers);
}

async function logout(request, env, headers) {
  const token = readCookie(request, 'timesheet_session');
  if (token) await env.DB.prepare('DELETE FROM sessions WHERE id = ?').bind(token).run();
  return new Response(JSON.stringify({ ok: true }), {
    headers: { ...headers, 'Set-Cookie': cookie('', 0) }
  });
}

async function session(request, env, headers) {
  const user = await authenticatedUser(request, env);
  return json({ ok: true, user: publicUser(user) }, 200, headers);
}

async function changePassword(request, env, headers) {
  const user = await authenticatedUser(request, env);
  const body = await request.json();
  const currentPassword = String(body.currentPassword || '');
  const newPassword = String(body.newPassword || '');
  if (!(await verifyPassword(currentPassword, user.password_hash, user.password_salt, user.password_iterations))) {
    throw new Error('Current password is incorrect.');
  }
  if (newPassword.length < 8) throw new Error('Password must be at least 8 characters.');
  const record = await hashPassword(newPassword);
  await env.DB.prepare(`UPDATE users SET password_hash = ?, password_salt = ?, password_iterations = ?, must_change_password = 0 WHERE id = ?`)
    .bind(record.hash, record.salt, record.iterations, user.id).run();
  return json({ ok: true }, 200, headers);
}

async function forgotPassword(request, env, headers) {
  const body = await request.json();
  const username = String(body.username || '').trim();
  const user = await env.DB.prepare('SELECT * FROM users WHERE username = ? COLLATE NOCASE').bind(username).first();
  if (user?.email) {
    const code = String(Math.floor(100000 + Math.random() * 900000));
    const tokenHash = await hashToken(code);
    const expires = new Date(Date.now() + 15 * 60 * 1000).toISOString();
    await env.DB.prepare('DELETE FROM password_reset_tokens WHERE user_id = ?').bind(user.id).run();
    await env.DB.prepare('INSERT INTO password_reset_tokens (user_id, token_hash, expires_at) VALUES (?, ?, ?)').bind(user.id, tokenHash, expires).run();
    await sendResetEmail(env, user, code);
  }
  return json({ ok: true, message: 'If the account exists, a reset email has been sent.' }, 200, headers);
}

async function recoverWithSecondaryPassword(request, env, headers) {
  const body = await request.json();
  const username = String(body.username || '').trim();
  const email = String(body.email || '').trim().toLowerCase();
  const recoveryPassword = String(body.recoveryPassword || '');
  const newPassword = String(body.newPassword || '');
  const user = await env.DB.prepare('SELECT * FROM users WHERE username = ? COLLATE NOCASE AND email = ? COLLATE NOCASE').bind(username, email).first();
  if (!user?.recovery_hash || !(await verifyPassword(recoveryPassword, user.recovery_hash, user.recovery_salt, user.recovery_iterations))) throw new Error('Recovery details are invalid.');
  if (newPassword.length < 8) throw new Error('New password must be at least 8 characters.');
  const record = await hashPassword(newPassword);
  await env.DB.prepare('UPDATE users SET password_hash = ?, password_salt = ?, password_iterations = ?, must_change_password = 0 WHERE id = ?').bind(record.hash, record.salt, record.iterations, user.id).run();
  await env.DB.prepare('DELETE FROM sessions WHERE user_id = ?').bind(user.id).run();
  return json({ ok: true, message: 'Password reset successfully. You can now sign in.' }, 200, headers);
}

async function sendResetEmail(env, user, code) {
  if (!env.RESEND_API_KEY || !env.RESEND_FROM_EMAIL) throw new Error('Password recovery email is not configured.');
  const response = await fetch('https://api.resend.com/emails', {
    method: 'POST',
    headers: { Authorization: `Bearer ${env.RESEND_API_KEY}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({
      from: env.RESEND_FROM_EMAIL,
      to: [user.email],
      subject: 'Timesheet password reset code',
      text: `Your Timesheet password reset code is ${code}. It expires in 15 minutes. If you did not request this, ignore this email.`
    })
  });
  if (!response.ok) throw new Error('Password recovery email could not be sent.');
}

async function listUsers(request, env, headers) {
  const admin = await adminUser(request, env);
  const result = await env.DB.prepare('SELECT username, display_name AS name, email, is_admin AS isAdmin FROM users ORDER BY username').all();
  return json({ users: result.results }, 200, headers);
}

async function createUser(request, env, headers) {
  await adminUser(request, env);
  const body = await request.json();
  const username = String(body.username || '').trim().toUpperCase();
  const name = String(body.name || username).trim();
  const email = String(body.email || '').trim().toLowerCase();
  const password = String(body.password || '');
  if (!/^[A-Z0-9._-]{2,40}$/.test(username)) throw new Error('Username must be 2-40 letters, numbers, dots, hyphens, or underscores.');
  if (password.length < 8) throw new Error('Password must be at least 8 characters.');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('A valid email address is required.');
  const record = await hashPassword(password);
  await env.DB.prepare(`INSERT INTO users (username, display_name, email, password_hash, password_salt, password_iterations, must_change_password)
    VALUES (?, ?, ?, ?, ?, ?, 1)`).bind(username, name, email, record.hash, record.salt, record.iterations).run();
  return json({ ok: true }, 201, headers);
}

async function updateUserEmail(request, env, headers) {
  await adminUser(request, env);
  const body = await request.json();
  const username = String(body.username || '').trim();
  const email = String(body.email || '').trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('A valid email address is required.');
  const result = await env.DB.prepare('UPDATE users SET email = ? WHERE username = ? COLLATE NOCASE').bind(email, username).run();
  if (!result.meta.changes) throw new Error('User not found.');
  return json({ ok: true }, 200, headers);
}

async function resetPassword(request, env, headers) {
  await adminUser(request, env);
  const body = await request.json();
  const password = String(body.password || '');
  if (password.length < 8) throw new Error('Password must be at least 8 characters.');
  const record = await hashPassword(password);
  const result = await env.DB.prepare(`UPDATE users SET password_hash = ?, password_salt = ?, password_iterations = ?, must_change_password = 1 WHERE username = ? COLLATE NOCASE`)
    .bind(record.hash, record.salt, record.iterations, String(body.username || '')).run();
  if (!result.meta.changes) throw new Error('User not found.');
  return json({ ok: true }, 200, headers);
}

async function createResetCode(request, env, headers) {
  await adminUser(request, env);
  const body = await request.json();
  const username = String(body.username || '').trim();
  const user = await env.DB.prepare('SELECT id FROM users WHERE username = ? COLLATE NOCASE').bind(username).first();
  if (!user) throw new Error('User not found.');
  const code = String(Math.floor(100000 + Math.random() * 900000));
  const tokenHash = await hashToken(code);
  const expires = new Date(Date.now() + 15 * 60 * 1000).toISOString();
  await env.DB.prepare('DELETE FROM password_reset_tokens WHERE user_id = ?').bind(user.id).run();
  await env.DB.prepare('INSERT INTO password_reset_tokens (user_id, token_hash, expires_at) VALUES (?, ?, ?)').bind(user.id, tokenHash, expires).run();
  return json({ ok: true, code, expiresAt: expires }, 200, headers);
}

async function listActiveSessions(request, env, headers) {
  const admin = await adminUser(request, env);
  const currentToken = readCookie(request, 'timesheet_session');
  const result = await env.DB.prepare(`
    SELECT s.id, s.ip_address AS ipAddress, s.user_agent AS userAgent,
           s.created_at AS createdAt, s.last_active_at AS lastActiveAt, s.expires_at AS expiresAt,
           u.id AS userId, u.username, u.display_name AS name, u.email, u.is_admin AS isAdmin
    FROM sessions s
    JOIN users u ON u.id = s.user_id
    WHERE s.expires_at > strftime('%Y-%m-%dT%H:%M:%fZ', 'now')
    ORDER BY s.last_active_at DESC, s.created_at DESC
  `).all();

  const sessions = (result.results || []).map(row => ({
    id: row.id,
    userId: row.userId,
    username: row.username,
    name: row.name,
    email: row.email,
    isAdmin: Boolean(row.isAdmin),
    ipAddress: row.ipAddress || 'Unknown',
    userAgent: row.userAgent || 'Unknown',
    createdAt: row.createdAt,
    lastActiveAt: row.lastActiveAt || row.createdAt,
    expiresAt: row.expiresAt,
    isCurrent: row.id === currentToken
  }));

  return json({ sessions }, 200, headers);
}

async function revokeSession(request, env, headers) {
  await adminUser(request, env);
  const body = await request.json();
  const sessionId = String(body.sessionId || '').trim();
  if (!sessionId) throw new Error('Session ID is required.');
  const result = await env.DB.prepare('DELETE FROM sessions WHERE id = ?').bind(sessionId).run();
  if (!result.meta.changes) throw new Error('Session not found or already ended.');
  return json({ ok: true, sessionId }, 200, headers);
}

async function listProjects(request, env, headers) {
  const user = await authenticatedUser(request, env);
  const result = user.is_admin
    ? await env.DB.prepare('SELECT id, name, status FROM projects ORDER BY name').all()
    : await env.DB.prepare('SELECT id, name, status FROM projects WHERE user_id = ? OR user_id IS NULL ORDER BY name').bind(user.id).all();
  const tasks = await env.DB.prepare('SELECT project_id, name FROM project_tasks ORDER BY name').all();
  const projects = result.results.map(project => ({ ...project, tasks: tasks.results.filter(task => task.project_id === project.id).map(task => task.name) }));
  return json({ projects }, 200, headers);
}

async function addProject(request, env, headers) {
  const user = await authenticatedUser(request, env);
  const body = await request.json();
  const name = String(body.name || '').trim();
  const status = String(body.status || 'New').trim();
  if (!name) throw new Error('Project name is required.');
  if (!['New', 'In-Progress', 'Hold', 'Complete'].includes(status)) throw new Error('Invalid project status.');
  await env.DB.prepare('INSERT INTO projects (user_id, name, status) VALUES (?, ?, ?)').bind(user.id, name, status).run();
  return json({ ok: true }, 201, headers);
}

async function updateProjectStatus(request, env, headers) {
  const user = await authenticatedUser(request, env);
  const body = await request.json();
  const projectId = Number(body.projectId || body.id);
  const status = String(body.status || '').trim();
  if (!Number.isInteger(projectId)) throw new Error('Project ID is required.');
  if (!['New', 'In-Progress', 'Hold', 'Complete'].includes(status)) throw new Error('Invalid project status.');
  const project = user.is_admin
    ? await env.DB.prepare('SELECT id FROM projects WHERE id = ?').bind(projectId).first()
    : await env.DB.prepare('SELECT id FROM projects WHERE id = ? AND (user_id = ? OR user_id IS NULL)').bind(projectId, user.id).first();
  if (!project) throw new Error('Project not found or you do not have permission to modify it.');
  await env.DB.prepare('UPDATE projects SET status = ? WHERE id = ?').bind(status, projectId).run();
  return json({ ok: true, projectId, status }, 200, headers);
}

async function addProjectTask(request, env, headers) {
  const user = await authenticatedUser(request, env);
  const body = await request.json();
  const projectId = Number(body.projectId);
  const name = String(body.name || '').trim();
  if (!Number.isInteger(projectId) || !name) throw new Error('Project and task name are required.');
  const project = user.is_admin
    ? await env.DB.prepare('SELECT id FROM projects WHERE id = ?').bind(projectId).first()
    : await env.DB.prepare('SELECT id FROM projects WHERE id = ? AND (user_id = ? OR user_id IS NULL)').bind(projectId, user.id).first();
  if (!project) throw new Error('Project not found.');
  await env.DB.prepare('INSERT INTO project_tasks (project_id, name) VALUES (?, ?)').bind(projectId, name).run();
  return json({ ok: true }, 201, headers);
}

async function listTeam(request, env, headers) {
  const user = await authenticatedUser(request, env);
  const result = user.is_admin
    ? await env.DB.prepare('SELECT id, name, status FROM team_members ORDER BY name').all()
    : await env.DB.prepare('SELECT id, name, status FROM team_members WHERE user_id = ? ORDER BY name').bind(user.id).all();
  const activity = await env.DB.prepare('SELECT id, member_id AS memberId, week_start AS week, activity, notes FROM team_activity ORDER BY week_start').all();
  return json({ members: result.results, activity: activity.results }, 200, headers);
}

async function addTeamMember(request, env, headers) {
  const user = await authenticatedUser(request, env);
  const body = await request.json();
  const name = String(body.name || '').trim();
  if (!name) throw new Error('Team member name is required.');
  await env.DB.prepare('INSERT INTO team_members (user_id, name, status) VALUES (?, ?, ?)').bind(user.id, name, 'Active').run();
  return json({ ok: true }, 201, headers);
}

async function addTeamActivity(request, env, headers) {
  const user = await authenticatedUser(request, env);
  const body = await request.json();
  const memberId = Number(body.memberId);
  let week = String(body.week || '').trim();
  const activity = String(body.activity || 'WFH').trim();

  if (/^\d{4}-\d{2}-\d{2}$/.test(week)) {
    const d = new Date(`${week}T00:00:00`);
    const w = isoWeek(d);
    week = `${w.year}-W${String(w.week).padStart(2, '0')}`;
  }

  if (!Number.isInteger(memberId) || !/^\d{4}-W\d{2}$/.test(week)) throw new Error('Select a person and a valid week.');
  const member = user.is_admin
    ? await env.DB.prepare('SELECT id FROM team_members WHERE id = ?').bind(memberId).first()
    : await env.DB.prepare('SELECT id FROM team_members WHERE id = ? AND user_id = ?').bind(memberId, user.id).first();
  if (!member) throw new Error('Team member not found.');
  await env.DB.prepare('INSERT INTO team_activity (member_id, week_start, activity, notes) VALUES (?, ?, ?, ?) ON CONFLICT(member_id, week_start, activity) DO UPDATE SET notes=excluded.notes').bind(memberId, week, activity, body.notes || null).run();
  return json({ ok: true }, 201, headers);
}

async function deleteTeamActivity(request, env, headers) {
  const user = await authenticatedUser(request, env);
  const body = await request.json();
  const activityId = Number(body.activityId);
  if (!Number.isInteger(activityId)) throw new Error('Activity ID is required.');
  const result = user.is_admin
    ? await env.DB.prepare('DELETE FROM team_activity WHERE id = ?').bind(activityId).run()
    : await env.DB.prepare(`DELETE FROM team_activity WHERE id = ? AND member_id IN (SELECT id FROM team_members WHERE user_id = ?)`).bind(activityId, user.id).run();
  if (!result.meta.changes) throw new Error('Activity record not found.');
  return json({ ok: true, activityId }, 200, headers);
}

async function deleteTeamMember(request, env, headers) {
  const user = await authenticatedUser(request, env);
  const body = await request.json();
  const memberId = Number(body.memberId);
  if (!Number.isInteger(memberId)) throw new Error('Team member id is required.');
  const result = user.is_admin
    ? await env.DB.prepare('DELETE FROM team_members WHERE id = ?').bind(memberId).run()
    : await env.DB.prepare('DELETE FROM team_members WHERE id = ? AND user_id = ?').bind(memberId, user.id).run();
  if (!result.meta.changes) throw new Error('Team member not found.');
  return json({ ok: true, memberId }, 200, headers);
}

async function listHolidays(request, env, headers) {
  await authenticatedUser(request, env);
  const result = await env.DB.prepare('SELECT id, holiday_date AS date, name FROM holidays ORDER BY holiday_date').all();
  return json({ holidays: result.results }, 200, headers);
}

async function addHoliday(request, env, headers) {
  await authenticatedUser(request, env);
  const body = await request.json();
  const date = String(body.date || '').trim();
  const name = String(body.name || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !name) throw new Error('Holiday date and name are required.');
  await env.DB.prepare('INSERT INTO holidays (holiday_date, name) VALUES (?, ?)').bind(date, name).run();
  return json({ ok: true }, 201, headers);
}

async function listOptions(request, env, headers) {
  await authenticatedUser(request, env);
  const result = await env.DB.prepare('SELECT list_name AS listName, value FROM entry_options ORDER BY list_name, value').all();
  const options = {};
  for (const row of result.results) (options[row.listName] ||= []).push(row.value);
  return json({ options }, 200, headers);
}

async function addOption(request, env, headers) {
  await authenticatedUser(request, env);
  const body = await request.json();
  const listName = String(body.listName || '').trim();
  const value = String(body.value || '').trim();
  const allowed = ['category', 'budget', 'billable', 'nonBillableReason'];
  if (!allowed.includes(listName)) throw new Error('Invalid option list.');
  if (!value) throw new Error('Option value is required.');
  if (value.length > 120) throw new Error('Option value is too long.');
  await env.DB.prepare('INSERT OR IGNORE INTO entry_options (list_name, value) VALUES (?, ?)').bind(listName, value).run();
  return json({ ok: true }, 201, headers);
}

async function deleteOption(request, env, headers) {
  await authenticatedUser(request, env);
  const body = await request.json();
  const listName = String(body.listName || '').trim();
  const value = String(body.value || '').trim();
  const allowed = ['category', 'budget', 'billable', 'nonBillableReason'];
  if (!allowed.includes(listName) || !value) throw new Error('Option list and value are required.');
  const result = await env.DB.prepare('DELETE FROM entry_options WHERE list_name = ? AND value = ?').bind(listName, value).run();
  if (!result.meta.changes) throw new Error('Option not found.');
  return json({ ok: true }, 200, headers);
}

async function listEntries(request, env, headers) {
  const user = await authenticatedUser(request, env);
  const result = await env.DB.prepare(`SELECT id, entry_date AS date, project, task, hours, billable, category, budget,
    notes, ticket, incident_type AS incidentType, non_billable_reason AS nonBillableReason
    FROM timesheet_entries WHERE user_id = ? ORDER BY entry_date DESC, id DESC`).bind(user.id).all();
  return json({ entries: result.results }, 200, headers);
}

async function addEntry(request, env, headers) {
  const user = await authenticatedUser(request, env);
  const body = await request.json();
  const date = String(body.date || '');
  const project = String(body.project || '').trim();
  const budget = String(body.budget || '').trim();
  const hours = Number(body.hours);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !project || !budget || !Number.isFinite(hours) || hours <= 0 || hours > 24) {
    throw new Error('Date, project, budget, and valid hours are required.');
  }
  const result = await env.DB.prepare(`INSERT INTO timesheet_entries
    (user_id, entry_date, project, task, hours, billable, category, budget, notes, ticket, incident_type, non_billable_reason)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`)
    .bind(user.id, date, project, body.task || null, hours, body.billable || null, body.category || null, budget,
      body.notes || null, body.ticket || null, body.incidentType || null, body.nonBillableReason || null).run();
  return json({ ok: true, id: result.meta.last_row_id }, 201, headers);
}

async function updateEntry(request, env, headers) {
  const user = await authenticatedUser(request, env);
  const body = await request.json();
  const id = Number(body.id);
  const date = String(body.date || '');
  const project = String(body.project || '').trim();
  const budget = String(body.budget || '').trim();
  const hours = Number(body.hours);
  if (!Number.isInteger(id) || !/^\d{4}-\d{2}-\d{2}$/.test(date) || !project || !budget || !Number.isFinite(hours) || hours <= 0 || hours > 24) {
    throw new Error('Entry, date, project, budget, and valid hours are required.');
  }
  const result = await env.DB.prepare(`UPDATE timesheet_entries SET entry_date = ?, project = ?, task = ?, hours = ?, billable = ?, category = ?, budget = ?, notes = ?, ticket = ?, incident_type = ?, non_billable_reason = ?, updated_at = strftime('%Y-%m-%dT%H:%M:%fZ', 'now') WHERE id = ? AND user_id = ?`)
    .bind(date, project, body.task || null, hours, body.billable || null, body.category || null, budget, body.notes || null, body.ticket || null, body.incidentType || null, body.nonBillableReason || null, id, user.id).run();
  if (!result.meta.changes) throw new Error('Entry not found.');
  return json({ ok: true, id }, 200, headers);
}

async function deleteEntry(request, env, headers) {
  const user = await authenticatedUser(request, env);
  const body = await request.json();
  const id = Number(body.id);
  if (!Number.isInteger(id)) throw new Error('Entry id is required.');
  const result = await env.DB.prepare('DELETE FROM timesheet_entries WHERE id = ? AND user_id = ?').bind(id, user.id).run();
  if (!result.meta.changes) throw new Error('Entry not found.');
  return json({ ok: true, id }, 200, headers);
}

async function listDailyTasks(request, env, headers) {
  const user = await authenticatedUser(request, env);
  const result = await env.DB.prepare(`
    SELECT id, title, status, priority, seq, task_date AS date, created_at AS createdAt, completed_at AS completedAt
    FROM daily_tasks
    WHERE user_id = ?
    ORDER BY priority ASC, seq ASC, created_at ASC
  `).bind(user.id).all();

  const tasks = {};
  for (const t of (result.results || [])) {
    tasks[t.id] = {
      id: t.id,
      title: t.title,
      status: t.status,
      priority: Number(t.priority || 3),
      seq: Number(t.seq || 0),
      date: t.date,
      createdAt: t.createdAt,
      completedAt: t.completedAt || null
    };
  }
  const today = new Date().toISOString().slice(0, 10);
  return json({ tasks, today }, 200, headers);
}

async function handleTaskEvent(request, env, headers) {
  const user = await authenticatedUser(request, env);
  const body = await request.json();
  const event = body || {};
  const type = String(event.type || '');
  const taskId = String(event.taskId || '').trim();

  if (!taskId) throw new Error('Task ID is required.');

  if (type === 'task:created') {
    const title = String(event.title || '').trim();
    if (!title) throw new Error('Task title is required.');
    const status = String(event.status || 'New');
    const priority = Number(event.priority || 3);
    const seq = Number(event.seq || 0);
    const date = String(event.date || new Date().toISOString().slice(0, 10));
    const createdAt = String(event.timestamp || new Date().toISOString());
    await env.DB.prepare(`
      INSERT INTO daily_tasks (id, user_id, title, status, priority, seq, task_date, created_at, completed_at)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, NULL)
      ON CONFLICT(id) DO UPDATE SET title=excluded.title, status=excluded.status, priority=excluded.priority, seq=excluded.seq, task_date=excluded.task_date
    `).bind(taskId, user.id, title, status, priority, seq, date, createdAt).run();
    return json({ ok: true, event: type }, 201, headers);
  }

  if (type === 'task:status_changed') {
    const newStatus = String(event.newStatus || 'New');
    const completedAt = newStatus === 'Complete' ? String(event.timestamp || new Date().toISOString()) : null;
    await env.DB.prepare(`
      UPDATE daily_tasks SET status = ?, completed_at = ? WHERE id = ? AND user_id = ?
    `).bind(newStatus, completedAt, taskId, user.id).run();
    return json({ ok: true, event: type }, 200, headers);
  }

  if (type === 'task:priority_changed') {
    const newPriority = Number(event.newPriority || 3);
    await env.DB.prepare(`
      UPDATE daily_tasks SET priority = ? WHERE id = ? AND user_id = ?
    `).bind(newPriority, taskId, user.id).run();
    return json({ ok: true, event: type }, 200, headers);
  }

  if (type === 'task:seq_changed') {
    const newSeq = Number(event.newSeq || 0);
    await env.DB.prepare(`
      UPDATE daily_tasks SET seq = ? WHERE id = ? AND user_id = ?
    `).bind(newSeq, taskId, user.id).run();
    return json({ ok: true, event: type }, 200, headers);
  }

  if (type === 'task:title_changed') {
    const newTitle = String(event.newTitle || '').trim();
    if (!newTitle) throw new Error('Task title cannot be empty.');
    await env.DB.prepare(`
      UPDATE daily_tasks SET title = ? WHERE id = ? AND user_id = ?
    `).bind(newTitle, taskId, user.id).run();
    return json({ ok: true, event: type }, 200, headers);
  }

  if (type === 'task:deleted') {
    await env.DB.prepare(`
      DELETE FROM daily_tasks WHERE id = ? AND user_id = ?
    `).bind(taskId, user.id).run();
    return json({ ok: true, event: type }, 200, headers);
  }

  throw new Error(`Unknown event type: ${type}`);
}

async function listReminders(request, env, headers) {
  const user = await authenticatedUser(request, env);
  const result = await env.DB.prepare(`
    SELECT id, title, due_date AS dueDate, due_time AS dueTime, priority, notes,
           is_completed AS isCompleted, completed_at AS completedAt, created_at AS createdAt
    FROM reminders
    WHERE user_id = ?
    ORDER BY is_completed ASC, due_date ASC, due_time ASC, priority ASC, id DESC
  `).bind(user.id).all();
  return json({ reminders: result.results || [] }, 200, headers);
}

async function addReminder(request, env, headers) {
  const user = await authenticatedUser(request, env);
  const body = await request.json();
  const title = String(body.title || '').trim();
  const dueDate = String(body.dueDate || '').trim();
  const dueTime = body.dueTime ? String(body.dueTime).trim() : null;
  const priority = Number(body.priority || 3);
  const notes = body.notes ? String(body.notes).trim() : null;

  if (!title) throw new Error('Reminder title is required.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dueDate)) throw new Error('A valid due date (YYYY-MM-DD) is required.');

  const result = await env.DB.prepare(`
    INSERT INTO reminders (user_id, title, due_date, due_time, priority, notes, is_completed, created_at)
    VALUES (?, ?, ?, ?, ?, ?, 0, strftime('%Y-%m-%dT%H:%M:%fZ', 'now'))
  `).bind(user.id, title, dueDate, dueTime, priority, notes).run();

  return json({ ok: true, id: result.meta.last_row_id }, 201, headers);
}

async function toggleReminder(request, env, headers) {
  const user = await authenticatedUser(request, env);
  const body = await request.json();
  const id = Number(body.id);
  if (!Number.isInteger(id)) throw new Error('Reminder ID is required.');

  const current = await env.DB.prepare('SELECT id, is_completed FROM reminders WHERE id = ? AND user_id = ?').bind(id, user.id).first();
  if (!current) throw new Error('Reminder not found.');

  const nextCompleted = current.is_completed ? 0 : 1;
  const completedAt = nextCompleted ? new Date().toISOString() : null;

  await env.DB.prepare(`
    UPDATE reminders SET is_completed = ?, completed_at = ? WHERE id = ? AND user_id = ?
  `).bind(nextCompleted, completedAt, id, user.id).run();

  return json({ ok: true, id, isCompleted: nextCompleted }, 200, headers);
}

async function deleteReminder(request, env, headers) {
  const user = await authenticatedUser(request, env);
  const body = await request.json();
  const id = Number(body.id);
  if (!Number.isInteger(id)) throw new Error('Reminder ID is required.');

  const result = await env.DB.prepare('DELETE FROM reminders WHERE id = ? AND user_id = ?').bind(id, user.id).run();
  if (!result.meta.changes) throw new Error('Reminder not found.');

  return json({ ok: true, id }, 200, headers);
}

async function authenticatedUser(request, env) {
  const token = readCookie(request, 'timesheet_session');
  if (!token) throw new Error('Authentication required.');
  const row = await env.DB.prepare(`SELECT u.* FROM sessions s JOIN users u ON u.id = s.user_id
    WHERE s.id = ? AND s.expires_at > strftime('%Y-%m-%dT%H:%M:%fZ', 'now')`).bind(token).first();
  if (!row) throw new Error('Session expired or invalid.');

  const ip = request.headers.get('CF-Connecting-IP') || request.headers.get('x-forwarded-for') || null;
  const ua = request.headers.get('user-agent') || null;
  env.DB.prepare(`
    UPDATE sessions
    SET last_active_at = strftime('%Y-%m-%dT%H:%M:%fZ', 'now'),
        ip_address = COALESCE(ip_address, ?),
        user_agent = COALESCE(user_agent, ?)
    WHERE id = ?
  `).bind(ip, ua, token).run().catch(() => {});

  return row;
}

async function adminUser(request, env) {
  const user = await authenticatedUser(request, env);
  if (!user.is_admin) throw new Error('Administrator access required.');
  return user;
}

function publicUser(user) {
  return { username: user.username, name: user.display_name, mustChangePassword: Boolean(user.must_change_password), isAdmin: Boolean(user.is_admin) };
}

async function hashPassword(password) {
  const salt = crypto.getRandomValues(new Uint8Array(16));
  const hash = await derive(password, salt, PASSWORD_ITERATIONS);
  return { hash: toBase64(hash), salt: toBase64(salt), iterations: PASSWORD_ITERATIONS };
}

async function verifyPassword(password, encodedHash, encodedSalt, iterations) {
  const hash = await derive(password, fromBase64(encodedSalt), Number(iterations));
  return timingSafeEqual(hash, fromBase64(encodedHash));
}

async function derive(password, salt, iterations) {
  const key = await crypto.subtle.importKey('raw', new TextEncoder().encode(password), 'PBKDF2', false, ['deriveBits']);
  const bits = await crypto.subtle.deriveBits({ name: 'PBKDF2', salt, iterations, hash: 'SHA-256' }, key, 256);
  return new Uint8Array(bits);
}

async function hashToken(value) {
  const bytes = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(value));
  return toBase64(new Uint8Array(bytes));
}

function json(value, status, headers) {
  return new Response(JSON.stringify(value), { status, headers });
}

function cookie(value, maxAge) {
  return `timesheet_session=${value}; Path=/; HttpOnly; Secure; SameSite=Lax; Max-Age=${maxAge}`;
}

function readCookie(request, name) {
  const header = request.headers.get('Cookie') || '';
  const item = header.split(';').map(x => x.trim()).find(x => x.startsWith(`${name}=`));
  return item ? item.slice(name.length + 1) : null;
}

function randomToken(bytes) {
  return toBase64(crypto.getRandomValues(new Uint8Array(bytes))).replace(/[+/=]/g, '').slice(0, bytes * 2);
}

function toBase64(bytes) { return btoa(String.fromCharCode(...bytes)); }
function fromBase64(value) { return Uint8Array.from(atob(value), char => char.charCodeAt(0)); }
function timingSafeEqual(a, b) {
  if (a.length !== b.length) return false;
  let result = 0;
  for (let i = 0; i < a.length; i++) result |= a[i] ^ b[i];
  return result === 0;
}
