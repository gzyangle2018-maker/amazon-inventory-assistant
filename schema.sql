-- D1 Database Schema for Amazon Inventory Assistant

CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT UNIQUE NOT NULL,
    password_hash TEXT NOT NULL,
    role TEXT NOT NULL DEFAULT 'user',
    created_at TEXT,
    is_active INTEGER DEFAULT 1,
    machine_limit INTEGER DEFAULT 1,
    department TEXT DEFAULT '',
    created_by TEXT DEFAULT ''
);

CREATE TABLE IF NOT EXISTS login_logs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT,
    machine_id TEXT,
    ip TEXT,
    login_time TEXT,
    success INTEGER
);

CREATE TABLE IF NOT EXISTS user_machines (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT NOT NULL,
    machine_id TEXT NOT NULL,
    os_type TEXT DEFAULT 'Windows',
    machine_name TEXT,
    bound_at TEXT,
    is_active INTEGER DEFAULT 1,
    UNIQUE(username, machine_id)
);

CREATE TABLE IF NOT EXISTS versions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    version TEXT UNIQUE NOT NULL,
    description TEXT,
    is_active INTEGER DEFAULT 1,
    created_at TEXT
);

CREATE TABLE IF NOT EXISTS upload_history (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT NOT NULL,
    filename TEXT NOT NULL,
    original_path TEXT,
    stored_path TEXT,
    uploaded_at TEXT,
    row_count INTEGER DEFAULT 0
);

CREATE TABLE IF NOT EXISTS seckill_reports (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT NOT NULL,
    created_at TEXT NOT NULL,
    items TEXT,
    ziniao_info TEXT,
    file_path TEXT
);

CREATE TABLE IF NOT EXISTS llm_config (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    api_key TEXT,
    base_url TEXT,
    model_name TEXT,
    updated_at TEXT
);

CREATE TABLE IF NOT EXISTS column_mappings (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT NOT NULL DEFAULT 'global',
    map_type TEXT NOT NULL,
    columns_json TEXT,
    updated_by TEXT DEFAULT '',
    updated_at TEXT,
    UNIQUE(username, map_type)
);

CREATE TABLE IF NOT EXISTS upload_file_data (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    history_id INTEGER NOT NULL UNIQUE,
    headers_json TEXT NOT NULL,
    data_json TEXT NOT NULL,
    green_rows_json TEXT DEFAULT '[]',
    created_at TEXT,
    FOREIGN KEY (history_id) REFERENCES upload_history(id)
);

-- Insert default admin
INSERT OR IGNORE INTO users (username, password_hash, role, created_at, is_active)
VALUES ('yangle', 'a8f5f167f44f4964e6c998dee827110c9a0c5e1e7a5b6e5f9d8c7b6a5f4e3d2c', 'admin', datetime('now'), 1);
