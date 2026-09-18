import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const file = process.argv[2] || path.join(path.dirname(fileURLToPath(import.meta.url)), 'options-export.json');
const environment = process.env.CF_ENV || 'dev';
const replace = process.env.REPLACE_OPTIONS === '1';
if (!fs.existsSync(file)) throw new Error(`Options export not found: ${file}`);
const data = JSON.parse(fs.readFileSync(file, 'utf8'));
const options = Array.isArray(data.options) ? data.options : (data.options ? Object.values(data.options) : []);
const q = value => `'${String(value ?? '').replaceAll("'", "''")}'`;
const sql = [];
if (replace) sql.push('DELETE FROM entry_options;');
for (const option of options) sql.push(`INSERT OR IGNORE INTO entry_options (list_name, value) VALUES (${q(option.listName)}, ${q(option.value)});`);
const database = environment === 'prod' ? 'timesheet-db' : `timesheet-db-${environment}`;
const sqlFile = path.join(path.dirname(fileURLToPath(import.meta.url)), `.options-${process.pid}.sql`);
fs.writeFileSync(sqlFile, sql.join('\n'), 'utf8');
const npx = process.platform === 'win32' ? 'npx.cmd' : 'npx';
try { execFileSync(npx, ['--yes', 'wrangler', 'd1', 'execute', database, '--env', environment, '--remote', '--file', sqlFile], { stdio: 'inherit', shell: true }); }
finally { fs.rmSync(sqlFile, { force: true }); }
console.log(`Imported ${options.length} option values into ${environment}.`);
