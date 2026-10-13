CREATE TABLE teachers (
  id TEXT PRIMARY KEY,
  token_hash TEXT NOT NULL UNIQUE,
  label TEXT NOT NULL,
  expires_at INTEGER NOT NULL,
  revoked_at INTEGER
);
CREATE TABLE rounds (
  id TEXT PRIMARY KEY,
  teacher_id TEXT NOT NULL REFERENCES teachers(id),
  title TEXT NOT NULL,
  criteria TEXT NOT NULL,
  deadline INTEGER NOT NULL,
  expires_at INTEGER NOT NULL,
  created_at INTEGER NOT NULL,
  closed_at INTEGER,
  max_responses INTEGER NOT NULL CHECK(max_responses BETWEEN 1 AND 200)
);
CREATE INDEX rounds_expiry ON rounds(expires_at);
CREATE INDEX rounds_owner ON rounds(teacher_id);
CREATE TABLE responses (
  seq INTEGER PRIMARY KEY AUTOINCREMENT,
  id TEXT NOT NULL UNIQUE,
  round_id TEXT NOT NULL REFERENCES rounds(id) ON DELETE CASCADE,
  name TEXT NOT NULL,
  name_key TEXT NOT NULL,
  email TEXT NOT NULL,
  browser_hash TEXT NOT NULL,
  answers TEXT NOT NULL,
  additional_quality TEXT NOT NULL,
  comment TEXT NOT NULL,
  submitted_at INTEGER NOT NULL,
  UNIQUE(round_id, name_key),
  UNIQUE(round_id, browser_hash)
);
