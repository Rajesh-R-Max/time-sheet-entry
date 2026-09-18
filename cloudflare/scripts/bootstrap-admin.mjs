import crypto from 'node:crypto';
import { execFileSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const password = process.env.ADMIN_PASSWORD;
const environment = process.env.CF_ENV || 'prod';
const username = (process.env.ADMIN_USERNAME || 'admin').trim().toUpperCase();
const displayName = (process.env.ADMIN_NAME || 'Administrator').trim();
if (!password || password.length < 8) throw new Error('Set ADMIN_PASSWORD to a password of at least 8 characters.');
if (!/^[A-Z0-9._-]{2,40}$/.test(username)) throw new Error('ADMIN_USERNAME is invalid.');

const iterations = 100000;
const salt = crypto.randomBytes(16);
const hash = crypto.pbkdf2Sync(password, salt, iterations, 32, 'sha256');
const sql = `INSERT INTO users (username, display_name, password_hash, password_salt, password_iterations, must_change_password, is_admin)
VALUES ('${sqlString(username)}', '${sqlString(displayName)}', '${hash.toString('base64')}', '${salt.toString('base64')}', ${iterations}, 0, 1)
ON CONFLICT(username) DO UPDATE SET display_name=excluded.display_name, password_hash=excluded.password_hash,
password_salt=excluded.password_salt, password_iterations=excluded.password_iterations, must_change_password=0, is_admin=1;`;

const npx = process.platform === 'win32' ? 'npx.cmd' : 'npx';
const database = environment === 'prod' ? 'timesheet-db' : `timesheet-db-${environment}`;
const sqlFile = path.join(path.dirname(fileURLToPath(import.meta.url)), `.bootstrap-${process.pid}.sql`);
fs.writeFileSync(sqlFile, sql, { encoding: 'utf8' });
try {
	execFileSync(npx, ['--yes', 'wrangler', 'd1', 'execute', database, '--env', environment, '--remote', '--file', sqlFile], {
		stdio: 'inherit',
		windowsHide: false,
		shell: true
	});
} catch (error) {
	throw new Error(`Wrangler D1 bootstrap failed: ${error.message}`);
} finally {
	fs.rmSync(sqlFile, { force: true });
}
console.log(`Admin account '${username}' bootstrapped in Cloudflare D1.`);

function sqlString(value) { return value.replaceAll("'", "''"); }
