/**
 * Redis cache wrapper — lazy connect, auto-reconnect, graceful degradation.
 */
let client = null;
let enabled = false;

function getClient() {
  if (client) return client;
  const url = String(process.env.REDIS_URL || '').trim();
  if (!url) return null;
  let Redis;
  try {
    Redis = require('ioredis');
  } catch {
    return null;
  }
  client = new Redis(url, {
    maxRetriesPerRequest: 3,
    retryStrategy(times) {
      if (times > 5) return null;
      return Math.min(times * 200, 2000);
    },
    lazyConnect: true,
  });
  client._ttl = Math.max(1, Number(process.env.REDIS_CACHE_TTL) || 300);
  client.on('connect', () => { enabled = true; });
  client.on('close', () => { enabled = false; });
  client.on('error', (err) => {
    if (process.env.NODE_ENV !== 'production') {
      try { console.error('[redis]', err.message); } catch {}
    }
  });
  return client;
}

async function ensureConnected() {
  const c = getClient();
  if (!c) return null;
  if (c.status === 'ready' || c.status === 'connecting') return c;
  try {
    await c.connect();
    enabled = true;
  } catch {
    enabled = false;
    return null;
  }
  return c;
}

function cacheKey(prefix, filename) {
  return `pos:${prefix}:${filename}`;
}

async function get(prefix, filename) {
  const c = await ensureConnected();
  if (!c) return null;
  try {
    const raw = await c.get(cacheKey(prefix, filename));
    if (raw == null) return null;
    return JSON.parse(raw);
  } catch {
    return null;
  }
}

async function set(prefix, filename, data) {
  const c = await ensureConnected();
  if (!c) return;
  try {
    const ttl = c._ttl || 300;
    await c.set(cacheKey(prefix, filename), JSON.stringify(data), 'EX', ttl);
  } catch {}
}

async function del(prefix, filename) {
  const c = await ensureConnected();
  if (!c) return;
  try {
    await c.del(cacheKey(prefix, filename));
  } catch {}
}

async function flushAll() {
  const c = await ensureConnected();
  if (!c) return;
  try {
    let cursor = '0';
    do {
      const [next, keys] = await c.scan(cursor, { match: 'pos:*', count: 100 });
      if (keys.length) await c.del(keys);
      cursor = next;
    } while (cursor !== '0');
  } catch {}
}

async function close() {
  if (!client) return;
  try {
    enabled = false;
    await client.quit();
  } catch {}
  client = null;
}

function isEnabled() {
  return enabled;
}

module.exports = { get, set, del, flushAll, close, isEnabled, getClient, ensureConnected };
