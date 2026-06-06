CREATE TABLE IF NOT EXISTS users (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  username TEXT UNIQUE NOT NULL,
  password TEXT NOT NULL,
  name TEXT NOT NULL,
  created_at DATETIME DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS records (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  user_id INTEGER NOT NULL,
  part_number TEXT NOT NULL,
  product_name TEXT NOT NULL,
  bom_cost REAL NOT NULL,
  margin REAL NOT NULL,
  exchange_rate REAL NOT NULL,
  qty TEXT,
  date TEXT,
  to_customer TEXT,
  attention TEXT,
  term TEXT,
  tel TEXT,
  remark TEXT,
  notes TEXT,
  file_name TEXT,
  created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
  FOREIGN KEY (user_id) REFERENCES users(id)
);

INSERT OR IGNORE INTO users (username, password, name) VALUES ('admin', '1234', 'Admin');
INSERT OR IGNORE INTO users (username, password, name) VALUES ('user1', '1234', 'User 1');