import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const file = process.argv[2] || path.join(path.dirname(fileURLToPath(import.meta.url)), 'excel-export.json');
const environment = process.env.CF_ENV || 'dev';
const password = process.env.MIGRATION_USER_PASSWORD;
const preserveExistingUser = process.env.PRESERVE_EXISTING_USER === '1';
const username = (process.env.MIGRATION_USER || 'RAJESH').trim().toUpperCase();
if (!preserveExistingUser && (!password || password.length < 8)) throw new Error('Set MIGRATION_USER_PASSWORD to a password of at least 8 characters.');
if (!fs.existsSync(file)) throw new Error(`Export file not found: ${file}`);
if (!/^[A-Z0-9._-]{2,40}$/.test(username)) throw new Error('MIGRATION_USER is invalid.');

const exportData = JSON.parse(fs.readFileSync(file, 'utf8'));
const entries = Array.isArray(exportData.entries) ? exportData.entries : [];
const iterations = 100000;
const sql = [];
const q = value => `'${String(value ?? '').replaceAll("'", "''")}'`;
if (!preserveExistingUser) {
  const salt = crypto.randomBytes(16);
  const hash = crypto.pbkdf2Sync(password, salt, iterations, 32, 'sha256');
  sql.push(`INSERT INTO users (username, display_name, password_hash, password_salt, password_iterations, must_change_password, is_admin) VALUES (${q(username)}, ${q(username)}, ${q(hash.toString('base64'))}, ${q(salt.toString('base64'))}, ${iterations}, 0, 0) ON CONFLICT(username) DO UPDATE SET password_hash=excluded.password_hash, password_salt=excluded.password_salt, password_iterations=excluded.password_iterations, must_change_password=0;`);
} else {
  sql.push(`SELECT 1 AS existing_user WHERE EXISTS (SELECT 1 FROM users WHERE username = ${q(username)} COLLATE NOCASE);`);
}
sql.push(`DELETE FROM timesheet_entries WHERE user_id = (SELECT id FROM users WHERE username = ${q(username)} COLLATE NOCASE);`);
for (const entry of entries) {
  sql.push(`INSERT INTO timesheet_entries (user_id, entry_date, project, task, hours, billable, category, budget, notes, ticket, incident_type, non_billable_reason) SELECT id, ${q(entry.date)}, ${q(entry.project)}, ${q(entry.task)}, ${Number(entry.hours)}, ${q(entry.billable)}, ${q(entry.category)}, ${q(entry.budget)}, ${q(entry.notes)}, ${q(entry.ticket)}, ${q(entry.incidentType)}, ${q(entry.nonBillableReason)} FROM users WHERE username = ${q(username)} COLLATE NOCASE;`);
}
const database = environment === 'prod' ? 'timesheet-db' : `timesheet-db-${environment}`;
const sqlFile = path.join(path.dirname(fileURLToPath(import.meta.url)), `.import-${process.pid}.sql`);
fs.writeFileSync(sqlFile, sql.join('\n'), 'utf8');
const npx = process.platform === 'win32' ? 'npx.cmd' : 'npx';
try {
  execFileSync(npx, ['--yes', 'wrangler', 'd1', 'execute', database, '--env', environment, '--remote', '--file', sqlFile], { stdio: 'inherit', shell: true });
} finally {
  fs.rmSync(sqlFile, { force: true });
}
console.log(`Imported ${entries.length} entries for ${username} into ${environment}.`);
