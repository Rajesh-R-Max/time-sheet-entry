# Timesheet Cloudflare Package

This is a fully Cloudflare-hosted version, isolated from the existing `webapp/` Excel/PowerShell server. It does not modify or replace the local server.

## Architecture

- Cloudflare Worker API in `src/index.js`
- Cloudflare D1 database schema in `schema.sql`
- Cloudflare-hosted static assets in `public/`
- No runtime dependency on PowerShell, Excel, localhost, `users.json`, or the local server

## Environments

The package has isolated `dev`, `test`, and `prod` Wrangler environments. Always pass `--env` or use `deploy.ps1` with an explicit environment.

## Promote Dev data to Test

Test uses its own D1 database. To refresh Test with Dev application data, the promotion script exports users, projects, tasks, team members, team activity, holidays, options, and timesheet entries from Dev, clears the corresponding Test data, and imports the Dev data. Sessions and password reset tokens are not copied.

This is destructive to existing Test application data and requires an explicit flag:

```powershell
.\scripts\promote-dev-to-test.ps1 -ReplaceExisting
```

The script does not modify Dev, Prod, Excel, or the local server. Review the Test data backup/export before running the refresh.

## Prerequisites

1. Install Node.js.
2. From this directory, run `npm install`.
3. Authenticate Wrangler with `npx wrangler login`.
4. Confirm the account with `npx wrangler whoami`.

## D1 databases

Create Dev and Test databases if they do not already exist:

```powershell
npx wrangler d1 create timesheet-db-dev
npx wrangler d1 create timesheet-db-test
```

Copy the returned IDs into `wrangler.toml`. Production already points to the existing `timesheet-db` database.

Apply the schema per environment:

```powershell
npx wrangler d1 execute timesheet-db-dev --env dev --remote --file .\schema.sql
npx wrangler d1 execute timesheet-db-test --env test --remote --file .\schema.sql
npx wrangler d1 execute timesheet-db --env prod --remote --file .\schema.sql
```

## Deploy

```powershell
.\deploy.ps1 -Environment dev
.\deploy.ps1 -Environment test
.\deploy.ps1 -Environment prod
```

Use `-ApplySchema` when an environment needs the schema applied before deployment.

## Bootstrap an administrator

Set a password without committing it to a file:

```powershell
$env:CF_ENV = 'dev'
$env:ADMIN_NAME = 'Administrator'
$env:ADMIN_PASSWORD = 'choose-a-strong-password'
npm run bootstrap-admin
Remove-Item Env:ADMIN_PASSWORD
Remove-Item Env:CF_ENV
Remove-Item Env:ADMIN_NAME
```

## Import Excel data for RAJESH

The migration reads `tblTimesheet` from the workbook, skips invalid rows, and imports valid rows into the selected D1 environment under `RAJESH`. D1 entries are isolated by `user_id`, so each user sees only their own rows. The migration replaces existing D1 rows for that username when `-ReplaceExisting` is supplied, which makes retries safe and prevents duplicates.

Set a temporary Cloudflare password for RAJESH, then run:

```powershell
$env:MIGRATION_USER_PASSWORD = 'choose-a-strong-rajesh-password'
.\scripts\migrate-excel.ps1 -Environment dev -UserName RAJESH -ReplaceExisting
Remove-Item Env:MIGRATION_USER_PASSWORD
```

The utility only reads Excel and writes hashed user credentials and entries to the selected D1 database. It does not edit the workbook or stop the local server. Review the exported row count before confirming a production migration.

## Import projects and tasks

The Projects sheet stores one project/task pair per row. Import it into D1 with:

```powershell
.\scripts\migrate-projects.ps1 -Environment dev -ReplaceExisting
```

The replace flag clears only Cloudflare `projects` and `project_tasks` rows before importing the Excel Projects sheet. It does not touch timesheet entries or the Excel workbook.

## Import Entry dropdown options

Import Category, Budget, Billable, and Non-billable Reason values from the Excel Lists sheet:

```powershell
.\scripts\migrate-options.ps1 -Environment dev -ReplaceExisting
```
## Current Cloudflare scope

- Users, sessions, password hashes, and timesheet entries are stored in D1.
- Login, logout, session validation, password change, admin user creation, admin password reset, list entries, and add entry are implemented.
- Projects, team members, and holidays have D1 tables, API routes, and UI forms.
- The local Excel server is not contacted by the deployed Worker.
