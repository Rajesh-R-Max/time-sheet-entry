import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const file = process.argv[2] || path.join(path.dirname(fileURLToPath(import.meta.url)), 'projects-export.json');
const environment = process.env.CF_ENV || 'dev';
const replace = process.env.REPLACE_PROJECTS === '1';
if (!fs.existsSync(file)) throw new Error(`Project export not found: ${file}`);
const data = JSON.parse(fs.readFileSync(file, 'utf8'));
const projects = Array.isArray(data.projects) ? data.projects : [];
const q = value => `'${String(value ?? '').replaceAll("'", "''")}'`;
const sql = [];
if (replace) { sql.push('DELETE FROM project_tasks;'); sql.push('DELETE FROM projects;'); }
for (const project of projects) {
  sql.push(`INSERT INTO projects (name, status) VALUES (${q(project.name)}, ${q(project.status)}) ON CONFLICT(name) DO UPDATE SET status=excluded.status;`);
  for (const task of (project.tasks || [])) {
    sql.push(`INSERT OR IGNORE INTO project_tasks (project_id, name) SELECT id, ${q(task)} FROM projects WHERE name = ${q(project.name)};`);
  }
}
const database = environment === 'prod' ? 'timesheet-db' : `timesheet-db-${environment}`;
const sqlFile = path.join(path.dirname(fileURLToPath(import.meta.url)), `.projects-${process.pid}.sql`);
fs.writeFileSync(sqlFile, sql.join('\n'), 'utf8');
const npx = process.platform === 'win32' ? 'npx.cmd' : 'npx';
try {
  execFileSync(npx, ['--yes', 'wrangler', 'd1', 'execute', database, '--env', environment, '--remote', '--file', sqlFile], { stdio: 'inherit', shell: true });
} finally {
  fs.rmSync(sqlFile, { force: true });
}
const taskCount = projects.reduce((total, project) => total + (project.tasks || []).length, 0);
console.log(`Imported ${projects.length} projects and ${taskCount} tasks into ${environment}.`);
