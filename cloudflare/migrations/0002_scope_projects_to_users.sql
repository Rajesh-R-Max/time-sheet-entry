PRAGMA foreign_keys = OFF;

CREATE TABLE projects_scoped (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  user_id INTEGER NOT NULL REFERENCES users(id),
  name TEXT NOT NULL,
  status TEXT NOT NULL DEFAULT 'New',
  created_at TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ', 'now'))
);

INSERT INTO projects_scoped (id, user_id, name, status, created_at)
SELECT p.id, (SELECT id FROM users WHERE username = 'RAJESH' COLLATE NOCASE LIMIT 1), p.name, p.status, p.created_at
FROM projects p;

CREATE TABLE project_tasks_scoped (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  project_id INTEGER NOT NULL REFERENCES projects_scoped(id) ON DELETE CASCADE,
  name TEXT NOT NULL,
  created_at TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ', 'now')),
  UNIQUE(project_id, name)
);

INSERT INTO project_tasks_scoped (id, project_id, name, created_at)
SELECT t.id, t.project_id, t.name, t.created_at FROM project_tasks t;

DROP TABLE project_tasks;
DROP TABLE projects;
ALTER TABLE projects_scoped RENAME TO projects;
ALTER TABLE project_tasks_scoped RENAME TO project_tasks;
CREATE UNIQUE INDEX idx_projects_user_name ON projects(user_id, name);
CREATE INDEX idx_project_tasks_project ON project_tasks(project_id);

PRAGMA foreign_keys = ON;
