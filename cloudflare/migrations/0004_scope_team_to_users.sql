PRAGMA foreign_keys = OFF;

CREATE TABLE team_members_scoped (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  user_id INTEGER NOT NULL REFERENCES users(id),
  name TEXT NOT NULL,
  status TEXT NOT NULL DEFAULT 'Active',
  created_at TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ', 'now'))
);

INSERT INTO team_members_scoped (id, user_id, name, status, created_at)
SELECT m.id, (SELECT id FROM users WHERE username = 'RAJESH' COLLATE NOCASE LIMIT 1), m.name, m.status, m.created_at
FROM team_members m;

CREATE TABLE team_activity_scoped (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  member_id INTEGER NOT NULL REFERENCES team_members_scoped(id) ON DELETE CASCADE,
  week_start TEXT NOT NULL,
  activity TEXT NOT NULL DEFAULT 'WFH',
  notes TEXT,
  created_at TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ', 'now')),
  UNIQUE(member_id, week_start, activity)
);

INSERT INTO team_activity_scoped (id, member_id, week_start, activity, notes, created_at)
SELECT a.id, a.member_id, a.week_start, a.activity, a.notes, a.created_at FROM team_activity a;

DROP TABLE team_activity;
DROP TABLE team_members;
ALTER TABLE team_members_scoped RENAME TO team_members;
ALTER TABLE team_activity_scoped RENAME TO team_activity;
CREATE UNIQUE INDEX idx_team_members_user_name ON team_members(user_id, name);
CREATE INDEX idx_team_activity_member ON team_activity(member_id, week_start);

PRAGMA foreign_keys = ON;
