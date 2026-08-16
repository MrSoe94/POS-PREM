const fs = require('fs');
const path = require('path');

const SINGLETON_FILES = new Set([
  'settings.json',
  'sync_config.json',
  'banners.json',
  'qris.json',
  'deletions.json',
  'lastSync.json',
  'cart.json',
  'trial-info.json',
]);

const CACHE_TTL_MS = 5000;

let db = null;
let dbPath = null;
let dataDir = null;
let initialized = false;
let SQL = null;
let dirty = false;
const cache = new Map();

function isSingletonFile(filename) {
  return SINGLETON_FILES.has(filename);
}

function cacheKey(filename) {
  return path.join(dataDir || '', filename);
}

function getCache(filename) {
  const k = cacheKey(filename);
  const v = cache.get(k);
  if (!v) return null;
  if (Date.now() - v.t > CACHE_TTL_MS) {
    cache.delete(k);
    return null;
  }
  return v.data;
}

function setCache(filename, data) {
  cache.set(cacheKey(filename), { t: Date.now(), data });
}

function invalidateCache(filename) {
  try {
    cache.delete(cacheKey(filename));
  } catch {}
}

function getDocId(doc) {
  if (!doc || typeof doc !== 'object') return null;
  const id = doc._id ?? doc.id;
  if (id != null && String(id).trim()) return String(id);
  if (doc.updatedAt) return `_sync_${doc.updatedAt}_${Math.random().toString(36).slice(2,8)}`;
  return `_sync_${Date.now()}_${Math.random().toString(36).slice(2,8)}`;
}

function getDbPath(dir) {
  return path.join(dir, 'pos.db');
}

async function loadSqlJs() {
  if (SQL) return SQL;
  const initSqlJs = require('sql.js');
  SQL = await initSqlJs();
  return SQL;
}

function markDirty() {
  dirty = true;
}

function saveDb(retries = 3) {
  if (!db || !dbPath || !dirty) return;
  const data = db._raw.export();
  const tmpPath = dbPath + '.tmp';
  try {
    fs.writeFileSync(tmpPath, Buffer.from(data));
    fs.renameSync(tmpPath, dbPath);
    dirty = false;
  } catch (e) {
    try { fs.unlinkSync(tmpPath); } catch {}
    if ((e.code === 'EBUSY' || e.code === 'EPERM') && retries > 0) {
      const start = Date.now();
      while (Date.now() - start < 30) {} // brief spin-wait
      saveDb(retries - 1);
    } else {
      console.error('[saveDb] Gagal menulis database:', e && e.message ? e.message : e);
      throw e;
    }
  }
}

function makePrepare(database) {
  return function prepare(sql) {
    return {
      get(...params) {
        const stmt = database.prepare(sql);
        try {
          if (params.length) stmt.bind(params);
          if (stmt.step()) return stmt.getAsObject();
          return undefined;
        } finally {
          stmt.free();
        }
      },
      all(...params) {
        const stmt = database.prepare(sql);
        const rows = [];
        try {
          if (params.length) stmt.bind(params);
          while (stmt.step()) rows.push(stmt.getAsObject());
          return rows;
        } finally {
          stmt.free();
        }
      },
      run(...params) {
        database.run(sql, params);
        markDirty();
      },
    };
  };
}

function wrapDatabase(database) {
  return {
    _raw: database,
    prepare: makePrepare(database),
    exec(sql) {
      database.exec(sql);
      markDirty();
    },
    transaction(fn) {
      database.run('BEGIN');
      try {
        fn();
        database.run('COMMIT');
      } catch (err) {
        try {
          database.run('ROLLBACK');
        } catch {}
        throw err;
      }
      saveDb();
    },
    close() {
      saveDb();
      try {
        database.close();
      } catch {}
    },
  };
}

function createSchema(database) {
  database.exec(`
    CREATE TABLE IF NOT EXISTS meta (
      key TEXT PRIMARY KEY,
      value TEXT NOT NULL
    );
    CREATE TABLE IF NOT EXISTS documents (
      collection TEXT NOT NULL,
      doc_id TEXT NOT NULL,
      data TEXT NOT NULL,
      updated_at INTEGER DEFAULT 0,
      PRIMARY KEY (collection, doc_id)
    );
    CREATE INDEX IF NOT EXISTS idx_documents_collection ON documents(collection);
    CREATE INDEX IF NOT EXISTS idx_documents_updated ON documents(collection, updated_at);
    CREATE TABLE IF NOT EXISTS singletons (
      name TEXT PRIMARY KEY,
      data TEXT NOT NULL
    );
  `);
}

function getMeta(database, key, fallback = null) {
  const row = database.prepare('SELECT value FROM meta WHERE key = ?').get(key);
  return row ? row.value : fallback;
}

function setMeta(database, key, value) {
  database.prepare('INSERT OR REPLACE INTO meta (key, value) VALUES (?, ?)').run(key, String(value));
}

function readSingleton(database, filename, fallback) {
  const row = database.prepare('SELECT data FROM singletons WHERE name = ?').get(filename);
  if (!row) return fallback;
  try {
    return JSON.parse(row.data);
  } catch {
    return fallback;
  }
}

function readCollection(database, filename, fallback) {
  const rows = database
    .prepare('SELECT data FROM documents WHERE collection = ? ORDER BY rowid')
    .all(filename);
  if (!rows.length) return fallback;
  const out = [];
  for (const row of rows) {
    try {
      out.push(JSON.parse(row.data));
    } catch {}
  }
  return out;
}

function writeSingleton(database, filename, data) {
  database.prepare('INSERT OR REPLACE INTO singletons (name, data) VALUES (?, ?)').run(
    filename,
    JSON.stringify(data)
  );
  saveDb();
}

function writeCollection(database, filename, data) {
  const arr = Array.isArray(data) ? data : [];
  const del = database.prepare('DELETE FROM documents WHERE collection = ?');
  const ins = database.prepare(
    'INSERT OR REPLACE INTO documents (collection, doc_id, data, updated_at) VALUES (?, ?, ?, ?)'
  );
  database.transaction(() => {
    del.run(filename);
    for (const doc of arr) {
      const docId = getDocId(doc);
      if (!docId) continue;
      ins.run(filename, docId, JSON.stringify(doc), Number(doc && doc.updatedAt) || 0);
    }
  });
}

function listCollections(database) {
  const docCols = database
    .prepare('SELECT DISTINCT collection AS name FROM documents ORDER BY collection')
    .all()
    .map((r) => r.name);
  const singletons = database
    .prepare('SELECT name FROM singletons ORDER BY name')
    .all()
    .map((r) => r.name);
  return [...new Set([...docCols, ...singletons])].sort();
}

async function migrateFromJson(dir) {
  const stats = { files: 0, documents: 0, errors: [] };
  if (!db) throw new Error('SQLite storage not initialized');
  const targetDir = dir || dataDir;
  if (!targetDir) return stats;

  let entries = [];
  try {
    entries = await fs.promises.readdir(targetDir);
  } catch (e) {
    stats.errors.push({ file: targetDir, error: e.message || String(e) });
    return stats;
  }

  const jsonFiles = entries.filter((name) => String(name).toLowerCase().endsWith('.json'));
  for (const filename of jsonFiles) {
    try {
      const filePath = path.join(targetDir, filename);
      const raw = await fs.promises.readFile(filePath, 'utf8');
      const parsed = JSON.parse(raw);
      if (isSingletonFile(filename) || (parsed && typeof parsed === 'object' && !Array.isArray(parsed))) {
        writeSingleton(db, filename, parsed);
        stats.files += 1;
        stats.documents += 1;
      } else if (Array.isArray(parsed)) {
        writeCollection(db, filename, parsed);
        stats.files += 1;
        stats.documents += parsed.length;
      } else {
        continue;
      }
      invalidateCache(filename);
    } catch (e) {
      stats.errors.push({ file: filename, error: e.message || String(e) });
    }
  }

  setMeta(db, 'json_migrated', '1');
  saveDb();
  return stats;
}

async function exportAllToJson() {
  if (!db) throw new Error('SQLite storage not initialized');
  const out = {};
  const collections = listCollections(db);
  for (const filename of collections) {
    invalidateCache(filename);
    out[filename] = await readData(filename);
  }
  return out;
}

async function init(dir, options = {}) {
  if (initialized && dataDir === dir && db) return { ok: true, already: true };

  dataDir = dir;
  await fs.promises.mkdir(dir, { recursive: true });

  const SqlJs = await loadSqlJs();
  dbPath = getDbPath(dir);
  const isNew = !fs.existsSync(dbPath);

  const rawDb = isNew
    ? new SqlJs.Database()
    : new SqlJs.Database(fs.readFileSync(dbPath));

  db = wrapDatabase(rawDb);
  createSchema(db);

  if (!getMeta(db, 'schema_version')) {
    setMeta(db, 'schema_version', '1');
    saveDb();
  }

  let hasJsonFiles = false;
  try {
    const entries = await fs.promises.readdir(dir);
    hasJsonFiles = entries.some((name) => String(name).toLowerCase().endsWith('.json'));
  } catch {}

  const shouldMigrate = hasJsonFiles && (isNew || !getMeta(db, 'json_migrated'));
  if (shouldMigrate && typeof options.onMigrated === 'function') {
    const stats = await migrateFromJson(dir);
    try {
      options.onMigrated(stats);
    } catch {}
  }

  initialized = true;
  return { ok: true, mode: 'sqlite', dbPath, isNew };
}

async function readData(filename) {
  if (!initialized) throw new Error('SQLite storage not initialized');
  const cached = getCache(filename);
  if (cached !== null) return cached;
  const fallback = filename.includes('.json') ? [] : {};
  let result;
  if (isSingletonFile(filename)) {
    result = readSingleton(db, filename, fallback);
  } else {
    const singleton = readSingleton(db, filename, null);
    result = singleton !== null ? singleton : readCollection(db, filename, fallback);
  }
  setCache(filename, result);
  return result;
}

async function writeData(filename, data) {
  if (!initialized) throw new Error('SQLite storage not initialized');
  if (isSingletonFile(filename) || (data && typeof data === 'object' && !Array.isArray(data))) {
    writeSingleton(db, filename, data);
  } else {
    writeCollection(db, filename, data);
  }
  setCache(filename, data);
}

function close() {
  if (db) {
    try {
      db.close();
    } catch {}
    db = null;
  }
  dbPath = null;
  initialized = false;
  dirty = false;
  cache.clear();
}

function isSqliteMode() {
  return true;
}

module.exports = {
  init,
  readData,
  writeData,
  isSqliteMode,
  setCache,
  invalidateCache,
  listCollections: () => (db ? listCollections(db) : []),
  getDbPath,
  close,
  isSingletonFile,
  SINGLETON_FILES,
  migrateFromJson,
  exportAllToJson,
};
