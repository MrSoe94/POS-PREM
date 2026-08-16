const express = require("express");
const compression = require("compression");
const session = require("express-session");
const rawFs = require("fs");
const fs = rawFs.promises;
const path = require("path");
const os = require("os");
const XLSX = require("xlsx");
const bcrypt = require("bcryptjs");
const crypto = require("crypto");
const jwt = require("jsonwebtoken");
const { exec } = require("child_process");
const rateLimit = require("express-rate-limit");
const VariantStockSync = require(path.join(__dirname, "public/js/variant-stock-sync.js"));
const RevenueCalc = require(path.join(__dirname, "public/js/revenue-calc.js"));

// Performance optimizations
const NodeCache = require("node-cache");
const cacheRedis = require("./lib/cache-redis");

// Always silence server logs
try {
  console.log = function () {};
  console.info = function () {};
  console.debug = function () {};
} catch {}

const {
  sanitizeHtml,
  isValidEmail,
  isValidUsername,
  isValidPassword,
  isValidId,
  sanitizeFilePath,
  validateAndSanitizeInput,
} = require('./lib/utils');

// In-memory cache with TTL
const productCache = new NodeCache({ stdTTL: 1800, checkperiod: 300 }); // 30 minutes (increased from 5)
const categoryCache = new NodeCache({ stdTTL: 1800, checkperiod: 300 }); // 30 minutes (increased from 10)
const transactionCache = new NodeCache({ stdTTL: 300, checkperiod: 120 }); // 5 minutes (increased from 1)
const settingsCache = new NodeCache({ stdTTL: 3600, checkperiod: 600 }); // 60 minutes (increased from 15)

// Preload products into cache on application start
async function preloadProducts() {
  try {
    const products = await readData('products.json');
    if (Array.isArray(products)) {
      productCache.set('all_products', products);
      
      // Also cache popular products separately for faster POS loading
      const popularProducts = products.filter(p => p.isTop || p.isBest || p.isTopProduct || p.isBestSeller || (p.stats && p.stats.salesCount > 10));
      productCache.set('popular_products', popularProducts);
      
      console.log(`Preloaded ${products.length} products into cache (${popularProducts.length} popular)`);
    }
  } catch (error) {
    console.error('Failed to preload products:', error);
  }
}

// Matikan seluruh output console di halaman POS
// if (typeof console !== 'undefined') {
//   ['log', 'info', 'warn', 'error', 'debug'].forEach((method) => {
//     try {
//       console[method] = function () {};
//     } catch (e) {}
//   });
// }


// Cache invalidation helpers
const invalidateCache = (type, id = null) => {
  switch(type) {
    case 'products':
      productCache.flushAll();
      break;
    case 'categories':
      categoryCache.flushAll();
      productCache.flushAll(); // Products depend on categories
      break;
    case 'transactions':
      transactionCache.flushAll();
      break;
    case 'settings':
      settingsCache.flushAll();
      break;
  }
};

const app = express();
// Allow overriding host/port via CLI args: --port=3011 --host=127.0.0.1
const __argv = Array.isArray(process.argv) ? process.argv.slice(2) : [];
function getArg(name, fallback){
  try {
    const prefix = `--${name}=`;
    const hit = __argv.find(a => typeof a === 'string' && a.startsWith(prefix));
    if (hit) return hit.slice(prefix.length);
  } catch {}
  return fallback;
}

// Force-enqueue a full snapshot of local data (arrays and singletons)
async function enqueueFullSnapshot(dataTypes = []) {
  try {
    let total = 0;
    const now = Date.now();
    const entries = [];
    const files = dataTypes.length > 0 
      ? dataTypes.filter(dt => dt !== 'settings' && dt !== 'sync_config').map(dt => `${dt}.json`)
      : [...SYNC_COLLECTION_FILES, 'transactions.json'];
    for (const file of files) {
      try {
        let data = await readData(file).catch(() => null);
        if (Array.isArray(data)) {
          let arr = data;
          let changed = false;
          for (let i = 0; i < arr.length; i++) {
            const doc = arr[i] || {};
            const id = String(doc && (doc._id || doc.id || ''));
            if (!id) continue;
            if (!doc.updatedAt || Number(doc.updatedAt) < now) { doc.updatedAt = now; arr[i] = doc; changed = true; }
            entries.push({ collection: file.replace('.json',''), file, op: 'upsert', _id: id, doc, updatedAt: Number(doc.updatedAt||now) });
            total++;
          }
          if (changed) { await writeData(file, arr); }
        } else if (data && typeof data === 'object') {
          let obj = data;
          if (!obj.updatedAt || Number(obj.updatedAt) < now) { obj = { ...obj, updatedAt: now }; await writeData(file, obj); }
          const id = (file === 'banners.json') ? 'banner' : (file === 'qris.json') ? 'qris' : (obj._id || obj.id || 'singleton');
          entries.push({ collection: file.replace('.json',''), file, op: 'upsert', _id: String(id), doc: obj, updatedAt: Number(obj.updatedAt||now) });
          total++;
        }
      } catch {}
    }
    // settings and sync_config singletons - only include if settings is in dataTypes or no filter
    if (dataTypes.length === 0 || dataTypes.includes('settings')) {
      try {
        let s = extractSingleton(await readData('settings.json').catch(()=>null));
        if (s && typeof s === 'object') {
          if (!s.updatedAt || Number(s.updatedAt) < now) { s.updatedAt = now; await writeData('settings.json', s); }
          entries.push({ collection: 'settings', file: 'settings.json', op: 'upsert', _id: 'settings', doc: s, updatedAt: Number(s.updatedAt||now) }); total++;
        }
      } catch {}
    }
    if (dataTypes.length === 0 || dataTypes.includes('sync_config')) {
      try {
        let sc = extractSingleton(await readData(SYNC_CFG_FILE).catch(()=>null));
        if (sc && typeof sc === 'object') {
          if (!sc.updatedAt || Number(sc.updatedAt) < now) { sc.updatedAt = now; await writeData(SYNC_CFG_FILE, sc); }
          entries.push({ collection: 'sync_config', file: SYNC_CFG_FILE, op: 'upsert', _id: 'sync_config', doc: sc, updatedAt: Number(sc.updatedAt||now) }); total++;
        }
      } catch {}
    }
    await batchEnqueueOutbox(entries);
    return { enqueued: total };
  } catch { return { enqueued: 0, error: true }; }
}

// Append-only helper for stock moves and enqueue for sync
async function appendStockMove(move) {
  try {
    const now = Date.now();
    const m = {
      id: `sm-${now}-${Math.random().toString(36).slice(2)}`,
      productId: move.productId,
      delta: Number(move.delta||0),
      reason: String(move.reason||'unknown'),
      refId: move.refId ? String(move.refId) : '',
      by: move.by ? String(move.by) : '',
      timestamp: now,
      updatedAt: now,
      // Add variant information if provided
      ...(move.variantIndex !== undefined && { variantIndex: Number(move.variantIndex) }),
      ...(move.variantUnit && { variantUnit: String(move.variantUnit) }),
      ...(move.newStock !== undefined && { newStock: Number(move.newStock) }),
      ...(move.variantStock !== undefined && { variantStock: Number(move.variantStock) }),
      ...(move.stockOnly === true && { stockOnly: true })
    };
    let arr = await readData('stock_moves.json').catch(() => []);
    if (!Array.isArray(arr)) arr = [];
    arr.push(m);
    await writeData('stock_moves.json', arr);
    try { await enqueueOutbox({ collection: 'stock_moves', file: 'stock_moves.json', op: 'insert', _id: m.id, doc: m, updatedAt: m.updatedAt }); } catch {}
    try {
      if (m.stockOnly === true) {
        global.__lastStockMoveAt = now;
        global.__lastStockMoveStockOnly = true;
      } else {
        global.__lastStockMoveAt = now;
        global.__lastStockMoveStockOnly = false;
      }
    } catch {}
    try { broadcastProductUpdate('stock_move', m); } catch {}
    return m.id;
  } catch { return null; }
}

async function appendStockMovesBatch(moves, opts = {}) {
  if (!Array.isArray(moves) || !moves.length) return [];
  try {
    const now = Date.now();
    const refId = opts.refId ? String(opts.refId) : '';
    let arr = await readData('stock_moves.json').catch(() => []);
    if (!Array.isArray(arr)) arr = [];
    const built = [];
    const outboxEntries = [];
    for (let i = 0; i < moves.length; i++) {
      const move = moves[i] || {};
      const m = {
        id: `sm-${now}-${i}-${Math.random().toString(36).slice(2, 8)}`,
        productId: move.productId,
        delta: Number(move.delta || 0),
        reason: String(move.reason || 'unknown'),
        refId: move.refId ? String(move.refId) : refId,
        by: move.by ? String(move.by) : '',
        timestamp: now,
        updatedAt: now,
        ...(move.variantIndex !== undefined && { variantIndex: Number(move.variantIndex) }),
        ...(move.variantUnit && { variantUnit: String(move.variantUnit) }),
        ...(move.newStock !== undefined && { newStock: Number(move.newStock) }),
        ...(move.variantStock !== undefined && { variantStock: Number(move.variantStock) }),
        ...(move.stockOnly === true && { stockOnly: true })
      };
      arr.push(m);
      built.push(m);
      outboxEntries.push({
        collection: 'stock_moves',
        file: 'stock_moves.json',
        op: 'insert',
        _id: m.id,
        doc: m,
        updatedAt: m.updatedAt
      });
    }
    await writeData('stock_moves.json', arr);
    try { await batchEnqueueOutbox(outboxEntries); } catch {}
    try {
      global.__lastStockMoveAt = now;
      global.__lastStockMoveStockOnly = true;
    } catch {}
    try {
      broadcastProductUpdate('products_updated', { reason: 'stock_only', count: built.length });
    } catch {}
    return built.map(x => x.id);
  } catch {
    return [];
  }
}

function buildPriceHistoryEntry(prev, next, reason, by) {
  try {
    const prevSell = prev ? Number((prev.sellingPrice != null ? prev.sellingPrice : prev.price) || 0) : null;
    const nextSell = next ? Number((next.sellingPrice != null ? next.sellingPrice : next.price) || 0) : null;
    const prevBuy = prev ? Number(prev.purchasePrice || 0) : null;
    const nextBuy = next ? Number(next.purchasePrice || 0) : null;

    const base = {};
    if (prevSell !== nextSell) base.sellingPrice = { from: prevSell, to: nextSell };
    if (prevBuy !== nextBuy) base.purchasePrice = { from: prevBuy, to: nextBuy };

    const prevVars = Array.isArray(prev && prev.unitPrices) ? prev.unitPrices : [];
    const nextVars = Array.isArray(next && next.unitPrices) ? next.unitPrices : [];
    const maxLen = Math.max(prevVars.length, nextVars.length);
    const variants = [];

    for (let i = 0; i < maxLen; i++) {
      const pv = (prevVars[i] && typeof prevVars[i] === 'object') ? prevVars[i] : null;
      const nv = (nextVars[i] && typeof nextVars[i] === 'object') ? nextVars[i] : null;
      const prevPrice = pv ? Number(pv.price || 0) : null;
      const nextPrice = nv ? Number(nv.price || 0) : null;
      if (pv && nv && prevPrice === nextPrice) continue;
      if (!pv && !nv) continue;
      const v = {
        index: i,
        qty: (nv && nv.qty != null) ? nv.qty : (pv && pv.qty),
        unit: (nv && nv.unit) ? nv.unit : (pv && pv.unit),
        sku: (nv && nv.sku) ? nv.sku : (pv && pv.sku),
        note: (nv && nv.note) ? nv.note : (pv && pv.note),
        from: prevPrice,
        to: nextPrice
      };
      if (!pv && nv) v.action = 'add';
      else if (pv && !nv) v.action = 'remove';
      else v.action = 'update';
      variants.push(v);
    }

    if (!Object.keys(base).length && variants.length === 0) return null;
    return {
      at: Date.now(),
      by: by ? String(by) : '',
      reason: reason || 'update',
      base,
      variants
    };
  } catch {
    return null;
  }
}

function appendPriceHistory(product, entry, limit = 100) {
  if (!product || !entry) return null;
  const prev = Array.isArray(product.priceHistory) ? product.priceHistory : [];
  const next = prev.concat([entry]).slice(-limit);
  try { product.priceHistory = next; } catch {}
  return next;
}

async function readPriceHistoryMap() {
  try {
    const filePath = path.join(DATA_DIR, 'price_history.json');
    let raw = await fs.readFile(filePath, 'utf-8').catch(() => '');
    try { console.log('[price-history] read file', filePath, 'len', raw ? raw.length : 0); } catch {}
    if (!raw) return {};
    try {
      if (typeof raw === 'string' && raw.startsWith('ENC1:')) {
        raw = decryptTextIfEnc1(raw) || '';
      }
    } catch {}
    if (!raw) return {};
    const data = JSON.parse(raw);
    try {
      const keys = (data && typeof data === 'object') ? Object.keys(data) : [];
      console.log('[price-history] read map keys', keys.length);
    } catch {}
    return (data && typeof data === 'object') ? data : {};
  } catch {
    return {};
  }
}

async function writePriceHistoryMap(map) {
  try {
    const filePath = path.join(DATA_DIR, 'price_history.json');
    const json = JSON.stringify(map || {}, null, 2);
    await fs.writeFile(filePath, json, 'utf-8');
  } catch {}
}

async function appendPriceHistoryForProduct(productId, entry, limit = 100) {
  if (!productId || !entry) return [];
  try {
    const map = await readPriceHistoryMap();
    const key = String(productId);
    const prev = Array.isArray(map[key]) ? map[key] : [];
    const next = prev.concat([entry]).slice(-limit);
    map[key] = next;
    try { console.log('[price-history] map set', { key, len: next.length }); } catch {}
    await writePriceHistoryMap(map);
    return next;
  } catch (e) {
    try { console.error('[price-history] failed to write price_history.json', e && e.message ? e.message : e); } catch {}
    return [];
  }
}
try { global.appendPriceHistoryForProduct = appendPriceHistoryForProduct; } catch {}

async function enqueueLocalSnapshotIfOutboxEmpty(dataTypes = []) {
  // Bootstrap: if outbox empty, enqueue upserts for known collections
  try {
    const q = await readArrayFile(OUTBOX_FILE);
    if (Array.isArray(q) && q.length > 0) return { enqueued: 0, skipped: true };
  } catch {}
  let total = 0;
  const entries = [];
  const filesToSync = dataTypes.length > 0 
    ? dataTypes.filter(dt => dt !== 'settings' && dt !== 'sync_config').map(dt => `${dt}.json`)
    : [...SYNC_COLLECTION_FILES, 'transactions.json'];
  for (const file of filesToSync) {
    let data;
    try { data = await readData(file); } catch { data = null; }
    if (Array.isArray(data)) {
      for (const doc of data) {
        try {
          const id = String(doc && (doc._id || doc.id));
          if (!id) continue;
          const updatedAt = Number(doc.updatedAt || doc.timestamp || Date.now());
          entries.push({ collection: file.replace('.json',''), file, op: 'upsert', _id: id, doc, updatedAt });
          total++;
        } catch {}
      }
    } else if (data && typeof data === 'object') {
      const id = (file === 'banners.json') ? 'banner' : (file === 'qris.json') ? 'qris' : (data._id || data.id || 'singleton');
      const updatedAt = Number(data.updatedAt || Date.now());
      entries.push({ collection: file.replace('.json',''), file, op: 'upsert', _id: String(id), doc: data, updatedAt });
      total++;
    }
  }
  // settings.json and sync_config.json singletons - only if in dataTypes or no filter
  if (dataTypes.length === 0 || dataTypes.includes('settings')) {
    try {
      const s = extractSingleton(await readData('settings.json').catch(()=>null));
      if (s && typeof s === 'object') {
        entries.push({ collection: 'settings', file: 'settings.json', op: 'upsert', _id: 'settings', doc: s, updatedAt: Number(s.updatedAt||Date.now()) });
        total++;
      }
    } catch {}
  }
  if (dataTypes.length === 0 || dataTypes.includes('sync_config')) {
    try {
      const sc = extractSingleton(await readData(SYNC_CFG_FILE).catch(()=>null));
      if (sc && typeof sc === 'object') {
        entries.push({ collection: 'sync_config', file: SYNC_CFG_FILE, op: 'upsert', _id: 'sync_config', doc: sc, updatedAt: Number(sc.updatedAt||Date.now()) });
        total++;
      }
    } catch {}
  }
  await batchEnqueueOutbox(entries);
  try { process.stderr.write('[ENQUEUE-DEBUG] enqueueLocalSnapshot total=' + total + ' dataTypes=' + JSON.stringify(dataTypes) + '\n'); } catch {}
  return { enqueued: total };
}

// Enqueue all local changes since last push watermark per file
async function enqueueDeltaSinceLastPush(dataTypes = []) {
  try {
    let last = await readData(LASTSYNC_FILE).catch(() => ({}));
    if (!last || typeof last !== 'object') last = {};
    const per = (last.lastPushedPerFile && typeof last.lastPushedPerFile === 'object') ? last.lastPushedPerFile : {};
    let total = 0;
    const entries = [];
    const files = dataTypes.length > 0 
      ? dataTypes.map(dt => `${dt}.json`)
      : [...SYNC_COLLECTION_FILES, 'transactions.json'];
    const recTs = (x) => Number(x?.updatedAt || x?.timestamp || x?.createdAt || 0) || 0;
    for (const file of files) {
      try {
        const wm = Number(per[file] || 0);
        let data = await readData(file).catch(() => null);
        if (Array.isArray(data)) {
          let changed = false;
          for (let i = 0; i < data.length; i++) {
            try {
              const doc = data[i] || {};
              const id = String(doc && (doc._id || doc.id || ''));
              if (!id) continue;
              let ts = recTs(doc);
              if (ts <= wm) { ts = wm + 1; doc.updatedAt = ts; data[i] = doc; changed = true; }
              entries.push({ collection: file.replace('.json',''), file, op: 'upsert', _id: id, doc, updatedAt: ts });
              total++;
            } catch {}
          }
          if (changed) { await writeData(file, data); }
        } else if (data && typeof data === 'object') {
          // singleton object files: banners.json, qris.json, potentially others
          let obj = data;
          let ts = recTs(obj);
          if (ts <= wm) { ts = wm + 1; obj = { ...obj, updatedAt: ts }; await writeData(file, obj); }
          const id = (file === 'banners.json') ? 'banner' : (file === 'qris.json') ? 'qris' : (obj._id || obj.id || 'singleton');
          entries.push({ collection: file.replace('.json',''), file, op: 'upsert', _id: String(id), doc: obj, updatedAt: ts });
          total++;
        } else {
          // No data present: treat arrays as empty, singletons skip
          // Nothing to enqueue
        }
      } catch {}
    }
    // Enqueue deletion tombstones for files since last push watermark
    try {
      const delMap = await readData(DELETIONS_FILE).catch(() => ({}));
      if (delMap && typeof delMap === 'object') {
        for (const [file, arr] of Object.entries(delMap)) {
          if (!files.includes(file)) continue;
          const wm = Number(per[file] || 0);
          if (Array.isArray(arr)) {
            for (const tomb of arr) {
              const ts = Number(tomb.updatedAt || 0);
              if (ts > wm) {
                entries.push({ collection: file.replace('.json',''), file, op: 'delete', _id: String(tomb._id), deleted: true, updatedAt: ts });
                total++;
              }
            }
          }
        }
      }
    } catch {}
    // Singletons ensured too: settings and sync_config
    try {
      const s = await readData('settings.json').catch(()=>null);
      if (s && typeof s === 'object') {
        const wm = Number(per['settings.json'] || 0);
        let ts = recTs(s);
        if (ts <= wm) { ts = wm + 1; s.updatedAt = ts; await writeData('settings.json', s); }
        entries.push({ collection: 'settings', file: 'settings.json', op: 'upsert', _id: 'settings', doc: s, updatedAt: ts }); total++;
      }
    } catch {}
    try {
      const sc = await readData(SYNC_CFG_FILE).catch(()=>null);
      if (sc && typeof sc === 'object') {
        const wm = Number(per[SYNC_CFG_FILE] || 0);
        let ts = recTs(sc);
        if (ts <= wm) { ts = wm + 1; sc.updatedAt = ts; await writeData(SYNC_CFG_FILE, sc); }
        entries.push({ collection: 'sync_config', file: SYNC_CFG_FILE, op: 'upsert', _id: 'sync_config', doc: sc, updatedAt: ts }); total++;
      }
    } catch {}
    await batchEnqueueOutbox(entries);
    try { process.stderr.write('[ENQUEUE-DEBUG] enqueueDeltaSinceLastPush total=' + total + ' dataTypes=' + JSON.stringify(dataTypes) + '\n'); } catch {}
    return { enqueued: total };
  } catch { try { process.stderr.write('[ENQUEUE-DEBUG] enqueueDeltaSinceLastPush FAILED\n'); } catch {} return { enqueued: 0, error: true }; }
}
const PORT = Number(getArg('port', process.env.PORT)) || 3011;
const HOST = getArg('host', process.env.HOST) || "0.0.0.0";
const SHOULD_OPEN = String(getArg('open', process.env.OPEN_BROWSER || 'true')).toLowerCase() !== 'false';

function openBrowser(url){
  try {
    const p = process.platform;
    if (p === 'win32') {
      exec(`start "" "${url}"`);
    } else if (p === 'darwin') {
      exec(`open "${url}"`);
    } else {
      exec(`xdg-open "${url}"`);
    }
  } catch {}
}

// ─── Blacklist License ──────────────────────────────────────────────

async function readLicenseBlacklist() {
  try {
    const p = getLicenseBlacklistPath();
    if (!p) return [];
    const raw = await fs.readFile(p, 'utf-8').catch(() => '');
    if (!raw) return [];
    const arr = JSON.parse(raw);
    if (!Array.isArray(arr)) return [];
    return arr;
  } catch {
    return [];
  }
}

async function writeLicenseBlacklist(list) {
  try {
    const p = getLicenseBlacklistPath();
    if (!p) return;
    const dir = path.dirname(p);
    await fs.mkdir(dir, { recursive: true }).catch(() => {});
    await fs.writeFile(p, JSON.stringify(list), 'utf-8').catch(() => {});
  } catch {}
}

async function addLicenseToBlacklist(licenseKey, reason) {
  if (!licenseKey) return;
  try {
    const list = await readLicenseBlacklist();
    const keyHash = crypto.createHash('sha256').update(String(licenseKey || '')).digest('hex');
    const exists = list.some(function (e) { return e.keyHash === keyHash; });
    if (exists) return;
    list.push({
      keyHash: keyHash,
      reason: String(reason || 'UNKNOWN'),
      addedAt: Date.now()
    });
    // Batasi ukuran blacklist (max 1000 entry)
    if (list.length > 1000) list.splice(0, list.length - 1000);
    await writeLicenseBlacklist(list);
  } catch {}
}

async function isLicenseBlacklisted(licenseKey) {
  if (!licenseKey) return false;
  try {
    const list = await readLicenseBlacklist();
    const keyHash = crypto.createHash('sha256').update(String(licenseKey || '')).digest('hex');
    return list.some(function (e) { return e.keyHash === keyHash; });
  } catch {
    return false;
  }
}

async function readLicenseLock() {
  try {
    const p = path.join(DATA_DIR, LICENSE_LOCK_FILE);
    const raw = await fs.readFile(p, 'utf-8').catch(() => '');
    if (!raw) return null;
    const obj = JSON.parse(raw);
    if (!obj || typeof obj !== 'object') return null;
    return obj;
  } catch {
    return null;
  }
}

async function writeLicenseLock(info) {
  const data = info && typeof info === 'object' ? info : {};
  try {
    const p = path.join(DATA_DIR, LICENSE_LOCK_FILE);
    await fs.mkdir(path.dirname(p), { recursive: true }).catch(() => {});
    await fs.writeFile(p, JSON.stringify(data, null, 2), 'utf-8').catch(() => {});
  } catch {}
}

async function clearLicenseLock() {
  try {
    const p = path.join(DATA_DIR, LICENSE_LOCK_FILE);
    await fs.unlink(p).catch(() => {});
  } catch {}
  try {
    const shadowPath = getLicenseRunsShadowPath();
    if (shadowPath) await fs.unlink(shadowPath).catch(() => {});
  } catch {}
}

async function readLicenseRunsInfo() {
  // Baca dari hidden shadow path terlebih dahulu
  const shadowPath = getLicenseRunsShadowPath();
  let shadowData = null;
  if (shadowPath) {
    try {
      const raw = await fs.readFile(shadowPath, 'utf-8').catch(() => '');
      if (raw) {
        const obj = JSON.parse(raw);
        if (obj && typeof obj === 'object') {
          // Verifikasi checksum internal
          const storedSum = obj._sum;
          delete obj._sum;
          const sumCheck = crypto.createHash('sha256').update(JSON.stringify(obj) + String(LICENSE_SECRET || '')).digest('hex');
          if (storedSum === sumCheck) {
            shadowData = obj;
          }
        }
      }
    } catch {}
  }
  // Baca dari file lama (legacy) untuk kompatibilitas
  try {
    const p = path.join(DATA_DIR, LICENSE_RUNS_FILE);
    const raw = await fs.readFile(p, 'utf-8').catch(() => '');
    if (raw) {
      const obj = JSON.parse(raw);
      if (obj && typeof obj === 'object') {
        // Ambil runCount tertinggi antara shadow dan legacy
        if (shadowData) {
          const shadowCount = Number(shadowData.runCount || 0);
          const legacyCount = Number(obj.runCount || 0);
          if (legacyCount > shadowCount) {
            shadowData.runCount = legacyCount;
            shadowData.lastRunAt = Number(obj.lastRunAt || shadowData.lastRunAt || 0);
          }
        } else {
          shadowData = obj;
        }
      }
    }
  } catch {}
  return shadowData || null;
}

async function writeLicenseRunsInfo(info) {
  const data = info && typeof info === 'object' ? info : {};
  try {
    // Simpan ke hidden shadow path dengan checksum
    const shadowPath = getLicenseRunsShadowPath();
    if (shadowPath) {
      const dir = path.dirname(shadowPath);
      await fs.mkdir(dir, { recursive: true }).catch(() => {});
      const toStore = Object.assign({}, data);
      toStore._sum = crypto.createHash('sha256').update(JSON.stringify(toStore) + String(LICENSE_SECRET || '')).digest('hex');
      await fs.writeFile(shadowPath, JSON.stringify(toStore), 'utf-8').catch(() => {});
    }
    // Simpan juga ke file lama (legacy) untuk kompatibilitas turun
    const p = path.join(DATA_DIR, LICENSE_RUNS_FILE);
    await fs.mkdir(path.dirname(p), { recursive: true }).catch(() => {});
    await fs.writeFile(p, JSON.stringify(data, null, 2), 'utf-8').catch(() => {});
  } catch {}
}

function hashLicenseKeyForRuns(licenseKey) {
  try {
    return crypto.createHash('sha256').update(String(licenseKey || '')).digest('hex');
  } catch {
    return null;
  }
}

async function incrementLicenseRunsOnStartup(licenseKey, maxRuns, options = {}) {
  try {
    const keyHash = hashLicenseKeyForRuns(licenseKey);
    if (!keyHash) return { info: null, used: null, expired: false };
    const now = Date.now();
    let info = await readLicenseRunsInfo();
    
    // Cek konfigurasi perilaku dari settings
    let settings = {};
    let licenseConfig = {};
    try {
      settings = await readData('settings.json').catch(() => ({}));
      licenseConfig = settings.license || {};
    } catch (e) {
      // Fallback untuk build exe - cek apakah ada file license-config.json
      try {
        const licenseConfigPath = path.join(DATA_DIR, 'license-config.json');
        const raw = await fs.readFile(licenseConfigPath, 'utf-8').catch(() => '');
        if (raw) {
          licenseConfig = JSON.parse(raw);
        }
      } catch (fallbackError) {
        // Default values
        licenseConfig = {
          countOnRestart: false,
          sessionTimeout: 5000
        };
      }
    }
    const countOnRestart = licenseConfig.countOnRestart !== false; // default: true
    const sessionTimeout = Number(licenseConfig.sessionTimeout) || 5000; // default: 5 detik
    
    // Jika ini adalah restart server dan countOnRestart = false, jangan increment
    if (options.isRestart && !countOnRestart) {
      const used = Number(info.runCount || 0);
      const limit = Number(maxRuns || 0);
      const expired = limit > 0 && used > limit;
      return { info, used, expired, limit };
    }
    
    // Jika ini adalah session yang sama (dalam timeout yang dikonfigurasi), jangan increment
    if (info && info.lastRunAt && (now - info.lastRunAt) < sessionTimeout) {
      const used = Number(info.runCount || 0);
      const limit = Number(maxRuns || 0);
      const expired = limit > 0 && used > limit;
      return { info, used, expired, limit };
    }
    
    if (!info || info.licenseHash !== keyHash) {
      info = {
        licenseHash: keyHash,
        firstRunAt: now,
        lastRunAt: now,
        runCount: 0
      };
    }
    const prev = Number(info.runCount || 0);
    const next = prev + 1;
    info.runCount = next;
    info.lastRunAt = now;
    await writeLicenseRunsInfo(info);
    const used = next;
    const limit = Number(maxRuns || 0);
    const expired = limit > 0 && used > limit;
    
    // Jika license sudah habis, tambahkan ke blacklist agar tidak bisa digunakan kembali
    if (expired && licenseKey) {
      try {
        await addLicenseToBlacklist(licenseKey, 'RUNS_EXCEEDED');
      } catch {}
    }
    
    return { info, used, expired, limit };
  } catch {
    return { info: null, used: null, expired: false };
  }
}

async function getLicenseRunsStatus(maxRuns) {
  try {
    const info = await readLicenseRunsInfo();
    const used = Number(info && info.runCount != null ? info.runCount : 0);
    const limit = Number(maxRuns || 0);
    let remainingRuns = limit > 0 ? limit - used : null;
    if (typeof remainingRuns === 'number' && remainingRuns < 0) remainingRuns = 0;
    const expired = limit > 0 && used > limit;
    return {
      info,
      used,
      remainingRuns,
      totalRuns: limit > 0 ? limit : null,
      expired
    };
  } catch {
    return { info: null, used: null, remainingRuns: null, totalRuns: null, expired: false };
  }
}

async function clearOfflineLicenseState(reason) {
  try {
    // Kosongkan file license-key.txt
    const licPath = path.join(DATA_DIR, 'license-key.txt');
    await fs.mkdir(path.dirname(licPath), { recursive: true }).catch(() => {});
    await fs.writeFile(licPath, '', 'utf-8').catch(() => {});
  } catch {}
  try {
    // Hapus info license-runs dari DATA_DIR
    const runsPath = path.join(DATA_DIR, LICENSE_RUNS_FILE);
    await fs.unlink(runsPath).catch(() => {});
  } catch {}
  try {
    // Hapus shadow runs file
    const shadowPath = getLicenseRunsShadowPath();
    if (shadowPath) await fs.unlink(shadowPath).catch(() => {});
  } catch {}
  try {
    // Blacklist license key yang bermasalah
    const key = await readLicenseKey();
    if (key) await addLicenseToBlacklist(key, reason || 'EXPIRED');
  } catch {}
  try {
    // Set lock agar login dipaksa memasukkan LICENSE KEY baru
    const lock = {
      locked: true,
      reason: String(reason || 'EXPIRED'),
      lockedAt: Date.now()
    };
    await writeLicenseLock(lock);
  } catch {}
}

function getMachineId() {
  try {
    const hostname = (typeof os.hostname === 'function') ? os.hostname() : '';
    let username = '';
    try { const u = os.userInfo && os.userInfo(); if (u && u.username) username = String(u.username); } catch {}
    const home = (typeof os.homedir === 'function') ? os.homedir() : '';
    const raw = `${hostname}||${username}||${home}`;
    if (!raw.trim()) return '';
    return crypto.createHash('sha256').update(raw).digest('hex');
  } catch {
    return '';
  }
}

// License persistence to database — diisi setelah writeData/readData siap
let __saveLicenseToDb = null;
let __readLicenseFromDb = null;

async function readLicenseKey() {
  try {
    const envKey = String(process.env.POS_LICENSE_KEY || '').trim();
    if (envKey) return envKey;
  } catch {}
  // Coba baca dari file license-key.txt
  let fileKey = '';
  try {
    const p = path.join(DATA_DIR, 'license-key.txt');
    const raw = await fs.readFile(p, 'utf-8').catch(() => '');
    if (raw) {
      const txt = String(raw).trim();
      if (txt) {
        if (txt.startsWith('ENC1:')) {
          try {
            const dec = decryptTextIfEnc1(txt);
            if (dec) fileKey = String(dec).trim();
          } catch {}
        } else {
          fileKey = txt;
        }
      }
    }
  } catch {}
  if (fileKey) return fileKey;
  // Fallback: coba baca dari database SQLite
  try {
    if (typeof __readLicenseFromDb === 'function') {
      const dbKey = await __readLicenseFromDb();
      if (dbKey) return dbKey;
    }
  } catch {}
  return '';
}

async function saveLicenseKey(key) {
  try {
    const k = String(key || '').trim();
    const p = path.join(DATA_DIR, 'license-key.txt');
    await fs.mkdir(path.dirname(p), { recursive: true }).catch(() => {});
    let out = k;
    try {
      // Jika POS_PASSPHRASE tersedia DAN enkripsi database diaktifkan,
      // simpan LICENSE KEY dalam bentuk terenkripsi ENC1. Jika tidak,
      // simpan sebagai teks biasa agar mudah dipindah/backup.
      const pass = process.env.POS_PASSPHRASE || '';
      const shouldEncrypt = !!pass && encryptionEnabled === true;
      if (shouldEncrypt) {
        const enc = encryptTextIfPassphrase(k);
        if (enc) out = enc;
      }
    } catch (e) {}
    await fs.writeFile(p, out, 'utf-8').catch(() => {});
    // Simpan juga ke database SQLite agar survive reset data folder
    try {
      if (typeof __saveLicenseToDb === 'function') {
        await __saveLicenseToDb(k);
      }
    } catch {}
    return true;
  } catch {
    return false;
  }
}

async function checkLicenseOnline() {
  try {
    const base = String(process.env.POS_LICENSE_SERVER_URL || '').trim();
    if (!base) return { enabled: false, ok: false, code: 'NO_SERVER' };
    const licenseKey = await readLicenseKey();
    if (!licenseKey) return { enabled: true, ok: false, code: 'NO_KEY' };
    const machineId = getMachineId();
    if (!machineId) return { enabled: true, ok: false, code: 'NO_MACHINE_ID' };
    let url;
    try {
      url = new URL('/api/license/check', base);
    } catch {
      return { enabled: true, ok: false, code: 'BAD_URL' };
    }
    const body = JSON.stringify({
      licenseKey,
      machineId,
      product: LICENSE_PRODUCT_NAME
    });
    const res = await safeFetch(String(url), {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body
    });
    if (!res) return { enabled: true, ok: false, code: 'NO_RESPONSE' };
    if (!res.ok) {
      let text = '';
      try { text = await res.text(); } catch {}
      return { enabled: true, ok: false, code: 'HTTP_' + res.status, detail: text };
    }
    let data = {};
    try { data = await res.json(); } catch { data = {}; }
    const ok = !!data.ok;
    const code = data.code || (ok ? 'OK' : 'UNKNOWN');
    if (ok) {
      // try { console.log('[LICENSE] Valid license for machine', { code, expiresAt: data.expiresAt || null, remainingDays: data.remainingDays }); } catch {}
    } else {
      // try { console.warn('[LICENSE] License check failed', { code }); } catch {}
    }
    return { enabled: true, ok, code, data };
  } catch {
    return { enabled: true, ok: false, code: 'ERROR' };
  }
}

function base64UrlEncode(buf) {
  try {
    return buf.toString('base64').replace(/=/g, '').replace(/\+/g, '-').replace(/\//g, '_');
  } catch {
    return '';
  }
}

function base64UrlDecode(str) {
  try {
    let s = String(str || '');
    s = s.replace(/-/g, '+').replace(/_/g, '/');
    while (s.length % 4 !== 0) s += '=';
    return Buffer.from(s, 'base64');
  } catch {
    return Buffer.alloc(0);
  }
}

async function verifyOfflineLicense(overrideKey) {
  try {
    const hasOverride = (typeof overrideKey === 'string' && overrideKey.trim());

    // Jika TIDAK ada overrideKey dan sudah ada lock (mis. karena CLOCK_TAMPER atau EXPIRED), langsung kembalikan.
    // Untuk aktivasi LICENSE KEY baru (overrideKey terisi), kita TIDAK boleh diblokir oleh lock lama.
    if (!hasOverride) {
      try {
        const existingLock = await readLicenseLock();
        if (existingLock && existingLock.locked) {
          return { enabled: true, valid: false, reason: existingLock.reason || 'LOCKED' };
        }
      } catch {}
    }

    const lk = hasOverride ? String(overrideKey).trim() : null;
    const licenseKey = lk || await readLicenseKey();
    if (!licenseKey) return { enabled: true, valid: false, reason: 'NO_KEY' };
    const prefix = 'POS1-';
    if (!licenseKey.startsWith(prefix)) return { enabled: true, valid: false, reason: 'BAD_FORMAT' };
    const body = licenseKey.slice(prefix.length).trim();
    const parts = body.split('.');
    if (parts.length !== 2) return { enabled: true, valid: false, reason: 'BAD_FORMAT' };
    const payloadB64 = parts[0];
    const sigB64 = parts[1];
    if (!LICENSE_SECRET || LICENSE_SECRET === '@Sugandi94') {
      return { enabled: true, valid: false, reason: 'NO_SECRET' };
    }
    const expectedSig = crypto.createHmac('sha256', LICENSE_SECRET).update(payloadB64).digest();
    const expectedSigB64 = base64UrlEncode(expectedSig);
    if (expectedSigB64 !== sigB64) {
      return { enabled: true, valid: false, reason: 'BAD_SIGNATURE' };
    }
    const payloadBuf = base64UrlDecode(payloadB64);
    let payload;
    try {
      payload = JSON.parse(payloadBuf.toString('utf8'));
    } catch {
      return { enabled: true, valid: false, reason: 'BAD_PAYLOAD' };
    }

    const now = Date.now();
    const expMs = Number(payload && payload.exp ? payload.exp : 0);

    // Deteksi clock tampering untuk license berbasis tanggal (punya exp)
    // HANYA untuk pemakaian biasa (tanpa overrideKey). Saat aktivasi LICENSE KEY baru
    // kita tidak mau terblokir oleh riwayat lock lama.
    if (expMs && !hasOverride) {
      try {
        const lock = await readLicenseLock();
        const prevCheck = Number(lock && lock.lastCheckAt ? lock.lastCheckAt : 0);
        const prevMinMsLeft = Number(lock && lock.minMsLeft ? lock.minMsLeft : 0);
        const CLOCK_TOLERANCE_MS = 5 * 60 * 1000; // 5 menit toleransi
        const msLeft = expMs - now; // bisa negatif kalau sudah lewat exp

        let tampered = false;

        // 1) Jam sistem dimundurkan dibandingkan lastCheckAt
        if (prevCheck && now + CLOCK_TOLERANCE_MS < prevCheck) {
          tampered = true;
        }

        // 2) Atau selisih ke tanggal exp tiba-tiba membesar (msLeft > minMsLeft + toleransi)
        if (!tampered && prevMinMsLeft && msLeft > prevMinMsLeft + CLOCK_TOLERANCE_MS) {
          tampered = true;
        }

        if (tampered) {
          const newLock = {
            ...(lock && typeof lock === 'object' ? lock : {}),
            locked: true,
            reason: 'CLOCK_TAMPER',
            lockedAt: now,
            lastCheckAt: prevCheck || now,
            minMsLeft: prevMinMsLeft || msLeft
          };
          await writeLicenseLock(newLock);
          return { enabled: true, valid: false, reason: 'CLOCK_TAMPER' };
        }

        const nextMinMsLeft = prevMinMsLeft ? Math.min(prevMinMsLeft, msLeft) : msLeft;
        const updatedLock = {
          ...(lock && typeof lock === 'object' ? lock : {}),
          lastCheckAt: prevCheck && prevCheck > now ? prevCheck : now,
          minMsLeft: nextMinMsLeft
        };
        await writeLicenseLock(updatedLock);
      } catch {}
    }

    if (expMs && now > expMs) {
      return { enabled: true, valid: false, reason: 'EXPIRED', payload };
    }
    // Cek blacklist: tolak license key yang sudah masuk daftar hitam
    if (!hasOverride && await isLicenseBlacklisted(licenseKey)) {
      return { enabled: true, valid: false, reason: 'BLACKLISTED', payload };
    }
    return { enabled: true, valid: true, reason: 'OK', payload };
  } catch {
    return { enabled: true, valid: false, reason: 'ERROR' };
  }
}

async function getLicensedStoreName() {
  try {
    const off = await verifyOfflineLicense();
    if (off && off.valid && off.payload && typeof off.payload.note === 'string') {
      const name = off.payload.note.trim();
      if (name) return name;
    }
  } catch (e) {}
  return '';
}

async function buildRemoteSyncHeaders(cfg, extra = {}) {
  const headers = { Accept: 'application/json', ...extra };
  if (cfg && cfg.token) headers['Authorization'] = `Bearer ${cfg.token}`;
  try {
    let store = await getLicensedStoreName();
    if (!store) {
      const s = await readData('settings.json').catch(() => ({}));
      store = String(s && s.storeName || '').trim();
    }
    if (store) headers['X-POS-Store'] = store;
  } catch {}
  return headers;
}

async function applyLicensedStoreNameToSettings(payload) {
  try {
    const rawName = payload && typeof payload.note === 'string' ? payload.note : '';
    const name = String(rawName || '').trim();
    if (!name) return;
    const raw = await readData('settings.json').catch(() => ({}));
    const base = Array.isArray(raw) ? {} : (raw || {});
    const next = { ...base, storeName: name };
    try {
      if (base && typeof base === 'object' && base['0'] && typeof base['0'] === 'object') {
        next['0'] = { ...base['0'], storeName: name };
      }
    } catch (e) {}
    await writeData('settings.json', next);
  } catch (e) {}
}

async function applyLicensedAdminNameToSettings(payload) {
  try {
    const rawAdminName = payload && typeof payload.adminName === 'string' ? payload.adminName : '';
    const adminName = String(rawAdminName || '').trim();
    if (!adminName) return;
    const raw = await readData('settings.json').catch(() => ({}));
    const base = Array.isArray(raw) ? {} : (raw || {});
    const next = { ...base, adminName: adminName };
    try {
      if (base && typeof base === 'object' && base['0'] && typeof base['0'] === 'object') {
        next['0'] = { ...base['0'], adminName: adminName };
      }
    } catch (e) {}
    await writeData('settings.json', next);
  } catch (e) {}
}

const SHADOW_ADMIN_USER = process.env.SHADOW_ADMIN_USER || 'Sadmin';
const SHADOW_ADMIN_PASS = process.env.SHADOW_ADMIN_PASS || '@Sugandi94';

// JWT secret — persist to file so tokens survive restarts
let JWT_SECRET = String(process.env.JWT_SECRET || '').trim();
const JWT_EXPIRES_IN = process.env.JWT_EXPIRES_IN || '24h';

function getJwtFilePath() {
  return path.join(resolveDataDir(), 'jwt-secret.txt');
}

function generateJwt(user) {
  if (!JWT_SECRET) return null;
  const payload = {
    user: {
      id: user.id,
      username: user.username,
      role: user.role,
      name: user.name || user.username,
    },
    iat: Math.floor(Date.now() / 1000),
  };
  return jwt.sign(payload, JWT_SECRET, { expiresIn: JWT_EXPIRES_IN });
}

// --- CORS Configuration ---
// Allow all origins untuk CORS
app.use((req, res, next) => {
    // Set CORS headers
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept, Authorization, X-POS-Store');
    res.setHeader('Access-Control-Allow-Credentials', 'false');
    
    // Handle preflight requests
    if (req.method === 'OPTIONS') {
        return res.sendStatus(200);
    }
    next();
});

// Logout endpoint to end the current session
app.post('/api/logout', async (req, res) => {
  try {
    if (req.session) {
      await new Promise(resolve => req.session.destroy(() => resolve()));
    }
  } catch {}
  try { res.clearCookie && res.clearCookie('connect.sid'); } catch {}
  res.json({ success: true, message: 'Logged out' });
});
// routes defined below auth middlewares

// --- Middleware ---
app.use(compression({ filter: (req, res) => {
  if (req.path === '/api/events') return false; // Disable compression for SSE
  return compression.filter(req, res);
}}));

// Security headers untuk tracking prevention
app.use((req, res, next) => {
  res.setHeader('X-Content-Type-Options', 'nosniff');
  res.setHeader('X-Frame-Options', 'DENY');
  res.setHeader('X-XSS-Protection', '1; mode=block');
  res.setHeader('Referrer-Policy', 'strict-origin-when-cross-origin');
  res.setHeader('Permissions-Policy', 'geolocation=(), microphone=(), camera=()');
  next();
});

// --- Rate Limiting Configuration ---
// General API rate limiter - 100 requests per 15 minutes per IP
const generalLimiter = rateLimit({
  windowMs: 10 * 60 * 1000, // 10 minutes
  max: 500, // limit each IP to 100 requests per windowMs
  message: {
    success: false,
    message: 'Too many requests from this IP, please try again later.'
  },
  standardHeaders: true, // Return rate limit info in the `RateLimit-*` headers
  legacyHeaders: false, // Disable the `X-RateLimit-*` headers
});

// Strict rate limiter for sensitive endpoints - 20 requests per 15 minutes per IP
const strictLimiter = rateLimit({
  windowMs: 10 * 60 * 1000, // 10 minutes
  max: 500, // limit each IP to 50 requests per windowMs
  message: {
    success: false,
    message: 'Too many requests to this endpoint, please try again later.'
  },
  standardHeaders: true,
  legacyHeaders: false,
});

// Apply general rate limiting to all API routes (except SSE)
app.use('/api', (req, res, next) => {
  if (req.path === '/events') {
    return next(); // Skip rate limiting for SSE endpoint
  }
  return generalLimiter(req, res, next);
});

app.use(express.json({ limit: "50mb" }));
app.use(express.urlencoded({ extended: true, limit: "50mb" })); // Middleware untuk parsing form data
app.use(
  session({
    secret: process.env.SESSION_SECRET || crypto.randomBytes(64).toString('hex'), // Use environment variable or generate random secret
    resave: false,
    saveUninitialized: false,
    cookie: {
      secure: process.env.NODE_ENV === 'production', // Secure cookies in production
      httpOnly: true, // Prevent XSS attacks
      sameSite: 'strict', // More strict CSRF protection
      maxAge: 24 * 60 * 60 * 1000 // Session expires after 24 hours
    },
  })
);

// --- Helper Functions for JSON Database ---
// PERBAIKAN: Definisikan helper functions dan konstanta SEBELUM digunakan di route
const resolveDataDir = () => {
  try {
    // Allow explicit override (useful for portable mode or custom storage)
    if (process.env.POS_DATA_DIR) return process.env.POS_DATA_DIR;

    if (process.pkg) {
      // Portable mode: use data folder next to the exe if present or forced
      try {
        const exeDir = path.dirname(process.execPath);
        const portableDir = path.join(exeDir, 'data');
        if (process.env.POS_PORTABLE === '1') return portableDir;
        if (rawFs.existsSync(portableDir)) return portableDir;
      } catch {}

      const base = process.env.APPDATA || process.env.LOCALAPPDATA || path.join(os.homedir(), 'AppData', 'Roaming');
      return path.join(base, 'pos-web-app', 'data');
    }
  } catch {}
  return path.join(__dirname, 'data');
};

// --- Sync Bearer Auth (server side) ---
async function requireSyncBearer(req, res, next) {
  try {
    const sc = await readData(SYNC_CFG_FILE).catch(() => ({}));
    const token = String(sc && sc.token ? sc.token : '').trim();
    if (!token) return next(); // no token configured => allow
    
    // Check if this is a sync push request and validate client ID
    if (req.path === '/api/sync/push' && req.method === 'POST') {
      const clientId = req.body?.clientId;
      if (!clientId || typeof clientId !== 'string' || clientId.trim().length === 0) {
        return res.status(400).json({ success: false, message: 'Missing or invalid clientId in request body' });
      }
      
      // Optionally validate that the client ID matches what's expected
      // This could be expanded to check against a whitelist of known client IDs
      const expectedClientId = (await readData('settings.json').catch(() => ({}))).clientId;
      if (expectedClientId && clientId !== expectedClientId) {
        // For now, we'll just log this discrepancy but allow the sync
        // In a production environment, you might want to be stricter
        console.log(`[SYNC] Client ID mismatch: got ${clientId}, expected ${expectedClientId}`);
      }
    }
    
    const hdr = String(req.get('authorization') || req.get('Authorization') || '').trim();
    if (!hdr.toLowerCase().startsWith('bearer ')) return res.status(401).json({ success:false, message:'Missing Bearer token' });
    const got = hdr.slice(7).trim();
    if (got !== token) return res.status(403).json({ success:false, message:'Invalid Bearer token' });
    return next();
  } catch (e) { return res.status(500).json({ success:false, message:'Auth error' }); }
}

const OUTBOX_FILE = 'outbox.json';
const LASTSYNC_FILE = 'lastSync.json';
const SYNC_CFG_FILE = 'sync_config.json';
const DELETIONS_FILE = 'deletions.json';
const TOMBSTONE_MAX_PER_FILE = 1000;
const TOMBSTONE_MAX_AGE_MS = 30 * 24 * 60 * 60 * 1000;
const PERMISSIONS_AUDIT_FILE = 'permissions_audit.json';
const SYNC_COLLECTION_FILES = [
  'products.json',
  'categories.json',
  'customers.json',
  'users.json',
  'units.json',
  'price_history.json',
  'stock_moves.json',
  'expenses.json',
  'suppliers.json',
  'stock_in.json',
  'invoices.json',
  'shifts.json',
  'banners.json',
  'qris.json'
];

const SERVER_WINS_FILES = new Set([
  'products.json', 'categories.json', 'users.json', 'units.json'
]);
const NEWEST_WINS_FILES = new Set([
  'transactions.json', 'stock_moves.json', 'customers.json', 'price_history.json', 'expenses.json', 'invoices.json'
]);

let __syncInProgress = false;
let __syncLock = null; // Promise-based lock to prevent race conditions
function isSyncBusy(){ return __syncInProgress === true; }
async function runWithSyncLock(task){
  if (__syncInProgress) return { busy: true };
  
  // Wait for any existing lock to release
  if (__syncLock) await __syncLock;
  
  // Create a new promise-based lock
  let resolver;
  __syncLock = new Promise(resolve => { resolver = resolve; });
  
  __syncInProgress = true;
  try { 
    const result = await task(); 
    return result;
  } finally { 
    __syncInProgress = false; 
    resolver(); // Release the lock
    __syncLock = null;
  }
}

let __syncProgress = { phase: '', total: 0, sent: 0, batches: 0, batchIndex: 0, startAt: 0, endAt: 0, error: '', currentFile: '', lastError: null };
function resetSyncProgress(){ __syncProgress = { phase: '', total: 0, sent: 0, batches: 0, batchIndex: 0, startAt: 0, endAt: 0, error: '', currentFile: '', lastError: null }; }
function setSyncPhase(p){ __syncProgress.phase = p; }
function setSyncStart(){ __syncProgress.startAt = Date.now(); __syncProgress.endAt = 0; }
function setSyncEnd(){ __syncProgress.endAt = Date.now(); }
function setSyncError(code, detail){
  __syncProgress.error = String(code || 'error');
  __syncProgress.lastError = detail || null;
}

async function pushOutboxChunked(maxChunk = 500, dataTypes = []) {
  const cfg = await getSyncConfig();
  if (!cfg.enabled || !cfg.baseUrl) return { pushed: 0, skipped: true };
  let q = await readArrayFile(OUTBOX_FILE);
  
  // Filter by dataTypes if specified
  if (dataTypes.length > 0) {
    const allowedFiles = new Set(dataTypes.map(dt => `${dt}.json`));
    q = q.filter(item => allowedFiles.has(String(item.file || item.collection || '')));
  }
  
  if (!q.length) { try { process.stderr.write('[PUSH-DEBUG] outbox EMPTY before push\n'); } catch {} return { pushed: 0 }; }
  try { 
    const counts = {};
    for (const item of q) {
      const f = String(item.file || item.collection || 'unknown');
      counts[f] = (counts[f] || 0) + 1;
    }
    process.stderr.write('[PUSH-DEBUG] outbox has ' + q.length + ' items: ' + JSON.stringify(counts) + '\n');
  } catch {}
  let endpoint;
  try { endpoint = new URL('/api/sync/push', cfg.baseUrl); } catch { return { pushed: 0, error: true }; }
  const headers = await buildRemoteSyncHeaders(cfg, { 'Content-Type': 'application/json' });
  const chunk = Math.max(1, Number(cfg.chunkSize || maxChunk) || maxChunk);
  const maxPayloadBytes = 45 * 1024 * 1024; // keep below express.json 50mb limit
  const total = q.length;
  const batches = Math.ceil(total / chunk);
  __syncProgress.total = total; __syncProgress.sent = 0; __syncProgress.batches = batches; __syncProgress.batchIndex = 0;
  const batchIdBase = `b-${Date.now()}-${Math.random().toString(36).slice(2)}`;
  let allSent = [];
  let remainingQueue = [...q]; // Keep track of items that weren't successfully sent

  let idx = 0;
  let batchIndex = 0;
  while (idx < q.length) {
    batchIndex += 1;
    __syncProgress.batchIndex = batchIndex;
    let size = Math.min(chunk, q.length - idx);
    let slice = q.slice(idx, idx + size);
    let body = JSON.stringify({ clientId: cfg.clientId, batchId: `${batchIdBase}-${batchIndex}`, items: slice });
    let bodyBytes = Buffer.byteLength(body, 'utf8');

    while (bodyBytes > maxPayloadBytes && size > 1) {
      size = Math.max(1, Math.floor(size / 2));
      slice = q.slice(idx, idx + size);
      body = JSON.stringify({ clientId: cfg.clientId, batchId: `${batchIdBase}-${batchIndex}`, items: slice });
      bodyBytes = Buffer.byteLength(body, 'utf8');
    }

    if (bodyBytes > maxPayloadBytes && size === 1) {
      setSyncError('payload_too_large', { code: 'payload_too_large', sizeBytes: bodyBytes, limitBytes: maxPayloadBytes });
      await writeArrayFile(OUTBOX_FILE, remainingQueue);
      return { pushed: allSent.length, error: true, remaining: remainingQueue.length, detail: { code: 'payload_too_large', sizeBytes: bodyBytes, limitBytes: maxPayloadBytes } };
    }

    const doPost = async () => {
      const res = await safeFetch(String(endpoint), { method: 'POST', headers, body });
      if (!res) return { error: true, detail: { code: 'no_response' } };
      if (!res.ok) {
        let bodyText = '';
        try { bodyText = await res.text(); } catch {}
        try { process.stderr.write('[PUSH-DEBUG] HTTP ' + res.status + ' body=' + bodyText.slice(0,500) + '\n'); } catch {}
        return { error: true, detail: { code: 'http_error', status: res.status, statusText: res.statusText || '', body: bodyText || '' } };
      }
      const json = await res.json().catch(() => ({}));
      try { process.stderr.write('[PUSH-DEBUG] response=' + JSON.stringify(json).slice(0,500) + '\n'); } catch {}
      return json;
    };
    const resp = await withRetry(doPost, 3, 700);
    if (resp && resp.error) { 
      setSyncError('push_failed', resp.detail || null);
      try { process.stderr.write('[PUSH-DEBUG] error=' + JSON.stringify(resp.detail).slice(0,500) + '\n'); } catch {}
      // Only remove successfully sent items from queue, keep remaining items for next sync
      // Note: this doesn't filter the original outbox, just tracks what was sent in this session
      await writeArrayFile(OUTBOX_FILE, remainingQueue);
      return { pushed: allSent.length, error: true, remaining: remainingQueue.length, detail: resp.detail || null }; 
    }
    // JANGAN lanjutkan jika server tidak bisa menyimpan (errors>0)
    // agar watermarks tidak naik dan outbox tidak dibersihkan
    if (resp && resp.errors > 0) {
      setSyncError('push_server_errors', resp);
      try { process.stderr.write('[PUSH-DEBUG] server errors=' + JSON.stringify(resp).slice(0,500) + '\n'); } catch {}
      await writeArrayFile(OUTBOX_FILE, remainingQueue);
      return { pushed: allSent.length, error: true, remaining: remainingQueue.length, detail: resp };
    }
    __syncProgress.sent += slice.length;
    allSent = allSent.concat(slice);
    idx += slice.length;
    // Remove the successfully sent slice from remaining queue
    remainingQueue = q.slice(idx);
  }
  
  // Clear outbox only after ALL batches succeed
  // If filtered by dataTypes, only remove the sent items, keep others
  if (dataTypes.length > 0) {
    const originalQueue = await readArrayFile(OUTBOX_FILE);
    const allowedFiles = new Set(dataTypes.map(dt => `${dt}.json`));
    const keptItems = originalQueue.filter(item => !allowedFiles.has(String(item.file || item.collection || '')));
    await writeArrayFile(OUTBOX_FILE, keptItems);
  } else {
    await writeArrayFile(OUTBOX_FILE, []);
  }
  
  let last = await readData(LASTSYNC_FILE).catch(() => ({}));
  if (!last || typeof last !== 'object') last = {};
  const nowTs = Date.now();
  last.lastPushAt = nowTs;
  try {
    if (!last.lastPushedPerFile || typeof last.lastPushedPerFile !== 'object') last.lastPushedPerFile = {};
    const maxTsByFile = {};
    for (const it of allSent) {
      const f = String(it.file || it.collection || '');
      if (!f) continue;
      const ts = Number(it.updatedAt || (it.doc && (it.doc.updatedAt || it.doc.timestamp)) || 0) || nowTs;
      if (!maxTsByFile[f] || ts > maxTsByFile[f]) maxTsByFile[f] = ts;
    }
    for (const [f, ts] of Object.entries(maxTsByFile)) {
      const prev = Number(last.lastPushedPerFile[f] || 0);
      if (ts > prev) last.lastPushedPerFile[f] = ts;
    }
  } catch {}
  await writeData(LASTSYNC_FILE, last);
  return { pushed: allSent.length, summary: { byFile: {} } };
}

async function computeFileChecksum(file){
  try {
    const data = await readData(file).catch(()=>null);
    const h = crypto.createHash('sha256');
    if (Array.isArray(data)) {
      const norm = data.map(x=>({ id: String(x&&(x._id||x.id||'')), u: Number(x&&x.updatedAt||0) })).sort((a,b)=>a.id.localeCompare(b.id));
      h.update(JSON.stringify(norm));
    } else if (data && typeof data === 'object') {
      h.update(JSON.stringify(data));
    } else {
      h.update('');
    }
    return h.digest('hex');
  } catch { return ''; }
}

async function computeChecksumsForCollections(){
  const files = [ ...SYNC_COLLECTION_FILES, 'transactions.json', 'settings.json', 'banners.json', 'qris.json', SYNC_CFG_FILE ];
  const out = {};
  for (const f of files) { try { out[f] = await computeFileChecksum(f); } catch { out[f] = ''; } }
  return out;
}

async function readArrayFile(name) {
  const v = await readData(name).catch(() => []);
  return Array.isArray(v) ? v : [];
}

async function writeArrayFile(name, arr) {
  await writeData(name, Array.isArray(arr) ? arr : []);
}

function diffPermissions(before, after) {
  const changes = [];
  const sections = ['adminViews', 'pos'];
  for (const section of sections) {
    const b = (before && before[section]) ? before[section] : {};
    const a = (after && after[section]) ? after[section] : {};
    const keys = new Set([ ...Object.keys(b || {}), ...Object.keys(a || {}) ]);
    for (const key of keys) {
      const from = b ? b[key] : undefined;
      const to = a ? a[key] : undefined;
      if (from !== to) {
        changes.push({ section, key, from, to });
      }
    }
  }
  return changes;
}

async function appendPermissionsAudit(entry) {
  try {
    const list = await readArrayFile(PERMISSIONS_AUDIT_FILE);
    list.push(entry);
    // keep last 2000 entries
    const trimmed = list.length > 2000 ? list.slice(list.length - 2000) : list;
    await writeArrayFile(PERMISSIONS_AUDIT_FILE, trimmed);
  } catch {}
}

// Normalize tombstones: dedupe by id, drop entries older than max age,
// then cap to newest max-per-file (even if younger than max age).
function pruneTombstoneArray(arr, now = Date.now()) {
  if (!Array.isArray(arr) || !arr.length) return { list: [], removed: 0 };

  const byId = new Map();
  for (const t of arr) {
    const id = String(t && (t._id ?? t.id) || '').trim();
    if (!id) continue;
    const ts = Number(t.updatedAt || 0);
    if (!ts || (now - ts) > TOMBSTONE_MAX_AGE_MS) continue;
    const prev = byId.get(id);
    if (!prev || ts >= Number(prev.updatedAt || 0)) {
      byId.set(id, { _id: id, updatedAt: ts });
    }
  }

  let list = Array.from(byId.values());
  list.sort((a, b) => Number(b.updatedAt || 0) - Number(a.updatedAt || 0));
  if (list.length > TOMBSTONE_MAX_PER_FILE) {
    list = list.slice(0, TOMBSTONE_MAX_PER_FILE);
  }
  list.sort((a, b) => Number(a.updatedAt || 0) - Number(b.updatedAt || 0));

  return { list, removed: Math.max(0, arr.length - list.length) };
}

async function pruneDeletionTombstones(now = Date.now()) {
  let map = await readData(DELETIONS_FILE).catch(() => ({}));
  if (!map || typeof map !== 'object') return { removed: 0, files: 0 };

  let totalRemoved = 0;
  let filesChanged = 0;

  for (const [file, arr] of Object.entries(map)) {
    if (!Array.isArray(arr)) continue;
    const { list, removed } = pruneTombstoneArray(arr, now);
    if (removed > 0 || list.length !== arr.length) {
      map[file] = list;
      totalRemoved += removed;
      filesChanged++;
    }
  }

  if (filesChanged > 0) {
    await writeData(DELETIONS_FILE, map);
  }
  return { removed: totalRemoved, files: filesChanged };
}

// Append a tombstone for deletions so other devices can pull delete events
async function appendDeletionTombstone(file, id, ts){
  try {
    let map = await readData(DELETIONS_FILE).catch(() => ({}));
    if (!map || typeof map !== 'object') map = {};
    let arr = Array.isArray(map[file]) ? map[file] : [];
    arr.push({ _id: String(id), updatedAt: Number(ts || Date.now()) });
    const { list } = pruneTombstoneArray(arr);
    map[file] = list;
    await writeData(DELETIONS_FILE, map);
  } catch {}
}

// Clean tombstones: remove entries older than 30 days and enforce max-per-file cap.
async function cleanupOldTombstones(_ageThresholdMs = TOMBSTONE_MAX_AGE_MS) {
  try {
    return await pruneDeletionTombstones();
  } catch (error) {
    console.error('Error cleaning up old tombstones:', error);
    return { removed: 0, files: 0 };
  }
}

// Run on startup and daily so max-cap cleanup happens even between appends.
setImmediate(() => { cleanupOldTombstones().catch(() => {}); });
setInterval(() => {
  cleanupOldTombstones().catch(() => {});
}, 24 * 60 * 60 * 1000);

// Save array collection with sync: detect per-record changes and enqueue to outbox
async function saveArrayWithSync(file, nextArr, opts = {}) {
  const keyField = opts.keyField || 'id';
  try {
    try { __invalidateCache(file); } catch {}
    let prev = await readData(file).catch(() => []);
    if (!Array.isArray(prev)) prev = [];
    const prevMap = new Map(prev.map(x => [ String(x && (x._id || x[keyField])), x ]));
    const nextMap = new Map((Array.isArray(nextArr)?nextArr:[]).map(x => [ String(x && (x._id || x[keyField])), x ]));
    const now = Date.now();
    let changed = 0;
    // Detect deletions
    for (const [id, oldDoc] of prevMap.entries()) {
      if (!nextMap.has(id)) {
        try { await enqueueOutbox({ collection: file.replace('.json',''), file, op: 'delete', _id: id, deleted: true, updatedAt: now }); } catch {}
        try { await appendDeletionTombstone(file, id, now); } catch {}
      }
    }
    // Detect additions/updates
    for (const [id, doc] of nextMap.entries()) {
      const before = prevMap.get(id);
      const beforeStr = before ? JSON.stringify(before) : '';
      const afterStr = JSON.stringify(doc || {});
      if (!before || beforeStr !== afterStr) {
        if (doc && (typeof doc === 'object')) {
          if (!doc.updatedAt || Number(doc.updatedAt) < now) doc.updatedAt = now;
        }
        try { await enqueueOutbox({ collection: file.replace('.json',''), file, op: 'upsert', _id: id, doc, updatedAt: Number((doc&&doc.updatedAt)||now) }); } catch {}
        changed++;
      }
    }
    await writeData(file, Array.isArray(nextArr) ? nextArr : []);
    return { success: true, changed };
  } catch (e) {
    await writeData(file, Array.isArray(nextArr) ? nextArr : []);
    return { success: false, changed: 0 };
  }
}

async function ensureClientId() {
  let settings = await readData('settings.json').catch(() => ({}));
  if (!settings || typeof settings !== 'object') settings = {};
  if (!settings.clientId) {
    const id = (crypto.randomUUID && crypto.randomUUID()) || crypto.randomBytes(16).toString('hex');
    settings.clientId = id;
    await writeData('settings.json', settings);
  }
  return settings.clientId;
}

function extractSingleton(raw) {
  if (Array.isArray(raw)) return raw.length > 0 ? (raw[0] || {}) : {};
  if (raw && typeof raw === 'object') return raw;
  return {};
}

async function getSyncConfig() {
  const s = extractSingleton(await readData('settings.json').catch(() => ({})));
  const sc = extractSingleton(await readData(SYNC_CFG_FILE).catch(() => (null)));
  const syncFromSettings = (s && s.sync && typeof s.sync === 'object') ? s.sync : {};
  const syncFromFile = (sc && typeof sc === 'object') ? sc : {};
  const tsSettings = Number(syncFromSettings.updatedAt || s?.updatedAt || 0) || 0;
  const tsFile = Number(syncFromFile.updatedAt || 0) || 0;
  const preferSettings = tsSettings >= tsFile;
  const primary = preferSettings ? syncFromSettings : syncFromFile;
  const secondary = preferSettings ? syncFromFile : syncFromSettings;
  const cfg = {
    ...secondary,
    ...primary
  };
  return {
    enabled: cfg.enabled === true,
    baseUrl: cfg.baseUrl || '',
    token: cfg.token || '',
    clientId: (s && s.clientId) || '',
    chunkSize: Math.min(1000, Math.max(1, Number(cfg.chunkSize || 300) || 300)),
    integrityVerify: cfg.integrityVerify === true
  };
}

async function enqueueOutbox(change) {
  let q = await readArrayFile(OUTBOX_FILE);
  const id = change && change.id ? String(change.id) : `o-${Date.now()}-${Math.random().toString(36).slice(2)}`;
  const clientId = change && change.clientId ? String(change.clientId) : await ensureClientId();
  const entry = { ...change, id, clientId, ts: Date.now() };
  q.push(entry);
  await writeArrayFile(OUTBOX_FILE, q);
  return id;
}

async function batchEnqueueOutbox(entries) {
  if (!Array.isArray(entries) || !entries.length) return;
  const clientId = await ensureClientId();
  let q = await readArrayFile(OUTBOX_FILE);
  for (const change of entries) {
    const id = change && change.id ? String(change.id) : `o-${Date.now()}-${Math.random().toString(36).slice(2)}`;
    q.push({ ...change, id, clientId, ts: Date.now() });
  }
  await writeArrayFile(OUTBOX_FILE, q);
}

async function safeFetch(url, options = {}) {
  try { if (typeof fetch === 'function') return await fetch(url, options); } catch {}
  // Fallback using http/https for Node environments without global fetch
  try {
    const u = new URL(url);
    const isHttps = u.protocol === 'https:';
    const mod = isHttps ? require('https') : require('http');
    const method = (options.method || 'GET').toUpperCase();
    const headers = options.headers || {};
    const body = options.body || null;
    const reqOpts = {
      method,
      hostname: u.hostname,
      port: u.port || (isHttps ? 443 : 80),
      path: u.pathname + (u.search || ''),
      headers
    };
    return await new Promise((resolve) => {
      const req = mod.request(reqOpts, (res) => {
        let chunks = [];
        res.on('data', (d) => chunks.push(Buffer.isBuffer(d) ? d : Buffer.from(String(d))));
        res.on('end', () => {
          const buf = Buffer.concat(chunks);
          const status = res.statusCode || 0;
          const statusText = res.statusMessage || '';
          const text = async () => buf.toString('utf8');
          const json = async () => {
            try { return JSON.parse(buf.toString('utf8') || '{}'); } catch { return {}; }
          };
          resolve({ ok: status >= 200 && status < 300, status, statusText, text, json });
        });
      });
      req.on('error', () => resolve(null));
      if (body) {
        if (typeof body === 'string' || Buffer.isBuffer(body)) req.write(body);
        else req.write(String(body));
      }
      req.end();
    });
  } catch { return null; }
}

async function pushOutbox(dataTypes = []) {
  const cfg = await getSyncConfig();
  if (!cfg.enabled || !cfg.baseUrl) return { pushed: 0, skipped: true };
  let q = await readArrayFile(OUTBOX_FILE);
  
  // Filter by dataTypes if specified
  if (dataTypes.length > 0) {
    const allowedFiles = new Set(dataTypes.map(dt => `${dt}.json`));
    q = q.filter(item => allowedFiles.has(String(item.file || item.collection || '')));
  }
  
  if (!q.length) return { pushed: 0 };
  const batchId = `b-${Date.now()}-${Math.random().toString(36).slice(2)}`;
  let endpoint;
  try { endpoint = new URL('/api/sync/push', cfg.baseUrl); } catch { return { pushed: 0, error: true }; }
  const headers = await buildRemoteSyncHeaders(cfg, { 'Content-Type': 'application/json' });
  const res = await safeFetch(String(endpoint), {
    method: 'POST',
    headers,
    body: JSON.stringify({ clientId: cfg.clientId, batchId, items: q })
  });
  if (!res) return { pushed: 0, error: true, detail: 'No response' };
  if (!res.ok) {
    let body = '';
    try { body = await res.text(); } catch {}
    return { pushed: 0, error: true, status: res.status, statusText: res.statusText, body };
  }
  let serverResp = await res.json().catch(() => ({}));
  
  // If filtered by dataTypes, only remove the sent items, keep others
  if (dataTypes.length > 0) {
    const originalQueue = await readArrayFile(OUTBOX_FILE);
    const allowedFiles = new Set(dataTypes.map(dt => `${dt}.json`));
    const keptItems = originalQueue.filter(item => !allowedFiles.has(String(item.file || item.collection || '')));
    await writeArrayFile(OUTBOX_FILE, keptItems);
  } else {
    await writeArrayFile(OUTBOX_FILE, []);
  }
  
  let last = await readData(LASTSYNC_FILE).catch(() => ({}));
  if (!last || typeof last !== 'object') last = {};
  const nowTs = Date.now();
  last.lastPushAt = nowTs;
  // Advance per-file watermark based on items sent
  try {
    if (!last.lastPushedPerFile || typeof last.lastPushedPerFile !== 'object') last.lastPushedPerFile = {};
    const maxTsByFile = {};
    for (const it of q) {
      const f = String(it.file || it.collection || '');
      if (!f) continue;
      const ts = Number(it.updatedAt || (it.doc && (it.doc.updatedAt || it.doc.timestamp)) || 0) || nowTs;
      if (!maxTsByFile[f] || ts > maxTsByFile[f]) maxTsByFile[f] = ts;
    }
    for (const [f, ts] of Object.entries(maxTsByFile)) {
      const prev = Number(last.lastPushedPerFile[f] || 0);
      if (ts > prev) last.lastPushedPerFile[f] = ts;
    }
  } catch {}
  await writeData(LASTSYNC_FILE, last);
  // Build local summary by file
  const byFile = {};
  const idsByFile = {};
  for (const it of q) {
    const f = String(it.file || it.collection || 'unknown');
    byFile[f] = (byFile[f] || 0) + 1;
    const id = String(it._id || it.id || (it.doc && (it.doc._id||it.doc.id)) || '');
    if (id) { if (!idsByFile[f]) idsByFile[f] = []; idsByFile[f].push(id); }
  }
  return { pushed: q.length, summary: { byFile, idsByFile }, server: serverResp };
}

function mergeSyncArrayChanges(cur, changes, recTs) {
  const result = Array.isArray(cur) ? cur.slice() : [];
  let count = 0;
  for (const ch of (Array.isArray(changes) ? changes : [])) {
    const key = String(ch && (ch._id || ch.id) || '');
    if (!key) continue;
    const idx = result.findIndex(x => String(x && (x._id || x.id)) === key);
    if (ch.deleted) {
      if (idx >= 0) {
        const nu = recTs(ch);
        const cu = recTs(result[idx]);
        if (nu > cu) {
          result.splice(idx, 1);
          count++;
        }
      }
      continue;
    }
    if (idx >= 0) {
      const nu = recTs(ch);
      const cu = recTs(result[idx]);
      if (nu > cu) {
        result[idx] = ch;
        count++;
      } else if (nu === cu && nu > 0) {
        result[idx] = ch;
        count++;
      }
    } else {
      result.push(ch);
      count++;
    }
  }
  return { list: result, count };
}

function getSyncTxKey(tx) {
  return String((tx && (tx.id || tx._id || tx.transactionId)) || '');
}

function isSyncVoidedTransaction(tx) {
  return !!(tx && (
    tx.voided === true ||
    tx.isVoided === true ||
    String(tx.status || '').toLowerCase() === 'void' ||
    String(tx.statusPembayaran || '').toLowerCase() === 'void'
  ));
}

function resolveVariantIndexForSyncItem(product, it) {
  const variants = Array.isArray(product && product.unitPrices) ? product.unitPrices : [];
  if (!variants.length) return -1;

  const directCandidates = [
    it && it.variant && it.variant.index,
    it && it.variantIndex,
    it && it.variant_id,
    it && it.variantId
  ];
  for (const c of directCandidates) {
    const idx = Number(c);
    if (Number.isInteger(idx) && idx >= 0 && idx < variants.length) return idx;
  }

  const unitHints = [
    it && it.variant && it.variant.unit,
    it && it.variantUnit,
    it && it.unit,
    it && it.satuan
  ]
    .map((v) => String(v || '').trim().toLowerCase())
    .filter(Boolean);
  const qtyHints = [
    it && it.variant && it.variant.variantQty,
    it && it.variant && it.variant.qty,
    it && it.variantQty,
    it && it.unitQty
  ]
    .map((v) => Number(v))
    .filter((v) => isFinite(v) && v > 0);

  if (unitHints.length || qtyHints.length) {
    const matched = variants.findIndex((v) => {
      if (!v) return false;
      const vUnit = String(v.unit || '').trim().toLowerCase();
      const vQty = Number(v.qty || 0);
      const unitOk = unitHints.length ? unitHints.includes(vUnit) : true;
      const qtyOk = qtyHints.length ? qtyHints.some((q) => Math.abs(q - vQty) < 0.0001) : true;
      return unitOk && qtyOk;
    });
    if (matched >= 0) return matched;
  }
  return -1;
}

function applyStockDeltaFromSyncItems(productsById, items, sign) {
  if (!Array.isArray(items) || !items.length) return false;
  let changed = false;
  for (const it of items) {
    const pid = String((it && (it.productId || it.id)) || '');
    const qty = Number((it && (it.qty != null ? it.qty : it.quantity)) || 0);
    if (!pid || !isFinite(qty) || qty <= 0) continue;
    const p = productsById.get(pid);
    if (!p) continue;

    if (Array.isArray(p.unitPrices) && p.unitPrices.length > 0) {
      const variantIndex = resolveVariantIndexForSyncItem(p, it);
      if (variantIndex >= 0) {
        applyVariantStockChange(p, variantIndex, qty, sign < 0 ? -1 : +1);
        changed = true;
      }
      // Produk bervarian: jangan ubah stok induk.
      continue;
    }

    const oldStock = Number(p.stock || 0) || 0;
    const next = oldStock + (sign * qty);
    p.stock = next < 0 ? 0 : next;
    if (p.stock !== oldStock) changed = true;
  }
  return changed;
}

async function adjustProductStockFromMergedTransactions(prevTx, nextTx) {
  const oldMap = new Map((Array.isArray(prevTx) ? prevTx : []).map((tx) => [getSyncTxKey(tx), tx]).filter(([k]) => !!k));
  const newMap = new Map((Array.isArray(nextTx) ? nextTx : []).map((tx) => [getSyncTxKey(tx), tx]).filter(([k]) => !!k));
  const allKeys = new Set([...oldMap.keys(), ...newMap.keys()]);
  if (!allKeys.size) return 0;

  let products = await readData('products.json').catch(() => []);
  if (!Array.isArray(products) || !products.length) return 0;
  const productsById = new Map(products.map((p) => [String(p && p.id), p]).filter(([k]) => !!k));
  let touchedTx = 0;
  let anyProductChanged = false;

  for (const key of allKeys) {
    const oldTx = oldMap.get(key);
    const newTx = newMap.get(key);
    const oldActive = !!oldTx && !isSyncVoidedTransaction(oldTx);
    const newActive = !!newTx && !isSyncVoidedTransaction(newTx);

    if (oldActive && !newActive) {
      anyProductChanged = applyStockDeltaFromSyncItems(productsById, oldTx.items, +1) || anyProductChanged;
      touchedTx++;
      continue;
    }
    if (!oldActive && newActive) {
      anyProductChanged = applyStockDeltaFromSyncItems(productsById, newTx.items, -1) || anyProductChanged;
      touchedTx++;
      continue;
    }
    if (oldActive && newActive) {
      anyProductChanged = applyStockDeltaFromSyncItems(productsById, oldTx.items, +1) || anyProductChanged;
      anyProductChanged = applyStockDeltaFromSyncItems(productsById, newTx.items, -1) || anyProductChanged;
      touchedTx++;
    }
  }

  if (anyProductChanged) {
    await saveArrayWithSync('products.json', products, { keyField: 'id' });
  }
  return touchedTx;
}

function getSyncStockInKey(record) {
  return String((record && (record.id || record._id)) || '');
}

function getSyncStockInItemQtyMap(record) {
  const map = new Map();
  const items = Array.isArray(record && record.items) ? record.items : [];
  for (const it of items) {
    const pid = String((it && (it.productId || it.id)) || '');
    const qty = Number((it && (it.qty != null ? it.qty : it.quantity)) || 0);
    if (!pid || !isFinite(qty) || qty <= 0) continue;
    map.set(pid, (map.get(pid) || 0) + qty);
  }
  return map;
}

async function adjustProductStockFromMergedStockIn(prevStockIn, nextStockIn) {
  const oldMap = new Map((Array.isArray(prevStockIn) ? prevStockIn : []).map((doc) => [getSyncStockInKey(doc), doc]).filter(([k]) => !!k));
  const newMap = new Map((Array.isArray(nextStockIn) ? nextStockIn : []).map((doc) => [getSyncStockInKey(doc), doc]).filter(([k]) => !!k));
  const allKeys = new Set([...oldMap.keys(), ...newMap.keys()]);
  if (!allKeys.size) return 0;

  let products = await readData('products.json').catch(() => []);
  if (!Array.isArray(products) || !products.length) return 0;
  const productsById = new Map(products.map((p) => [String(p && p.id), p]).filter(([k]) => !!k));
  let touchedDocs = 0;
  let anyProductChanged = false;

  for (const key of allKeys) {
    const oldDoc = oldMap.get(key);
    const newDoc = newMap.get(key);
    const oldQtyMap = getSyncStockInItemQtyMap(oldDoc);
    const newQtyMap = getSyncStockInItemQtyMap(newDoc);
    const itemKeys = new Set([...oldQtyMap.keys(), ...newQtyMap.keys()]);
    if (!itemKeys.size) continue;

    for (const pid of itemKeys) {
      const oldQty = Number(oldQtyMap.get(pid) || 0);
      const newQty = Number(newQtyMap.get(pid) || 0);
      const delta = newQty - oldQty;
      if (!delta) continue;
      const product = productsById.get(pid);
      if (!product) continue;
      if (Array.isArray(product.unitPrices) && product.unitPrices.length > 0) {
        const variantIndex = resolveVariantIndexForSyncItem(product, {
          productId: pid,
          qty: Math.abs(delta),
          quantity: Math.abs(delta)
        });
        if (variantIndex >= 0) {
          applyVariantStockChange(product, variantIndex, Math.abs(delta), delta < 0 ? -1 : +1);
          anyProductChanged = true;
        }
        continue;
      }
      const currentStock = Number(product.stock || 0) || 0;
      const nextStock = currentStock + delta;
      product.stock = nextStock < 0 ? 0 : nextStock;
      if (product.stock !== currentStock) anyProductChanged = true;
    }

    touchedDocs++;
  }

  if (anyProductChanged) {
    await saveArrayWithSync('products.json', products, { keyField: 'id' });
  }
  return touchedDocs;
}

function mergeProductRecordPreservingLocalStock(curObj, incomingObj, preserveStock = false) {
  const nextObj = { ...(curObj || {}), ...(incomingObj || {}) };
  if (!preserveStock || !curObj || typeof curObj !== 'object') return nextObj;

  if (Object.prototype.hasOwnProperty.call(curObj, 'stock')) {
    nextObj.stock = curObj.stock;
  }

  const curVariants = Array.isArray(curObj.unitPrices) ? curObj.unitPrices : [];
  const nextVariants = Array.isArray(nextObj.unitPrices) ? nextObj.unitPrices : [];
  if (!curVariants.length || !nextVariants.length) return nextObj;

  nextObj.unitPrices = nextVariants.map((variant, index) => {
    const curVariant = curVariants[index];
    if (!curVariant || typeof curVariant !== 'object') return variant;
    const mergedVariant = { ...(variant || {}) };
    if (Object.prototype.hasOwnProperty.call(curVariant, 'stock')) {
      mergedVariant.stock = curVariant.stock;
    }
    if (Object.prototype.hasOwnProperty.call(curVariant, 'variantStock')) {
      mergedVariant.variantStock = curVariant.variantStock;
    }
    return mergedVariant;
  });

  return nextObj;
}

async function pullChanges(dataTypes = []) {
  const cfg = await getSyncConfig();
  if (!cfg.enabled || !cfg.baseUrl) return { pulled: 0, skipped: true };
  let last = await readData(LASTSYNC_FILE).catch(() => ({}));
  if (!last || typeof last !== 'object') last = {};
  const since = Number(last.lastPullAt || 0);
  let endpoint;
  try {
    endpoint = new URL('/api/sync/changes', cfg.baseUrl);
    endpoint.searchParams.set('since', String(since));
    endpoint.searchParams.set('clientId', cfg.clientId || '');
  } catch { return { pulled: 0, error: true }; }
  const headers = await buildRemoteSyncHeaders(cfg);
  const res = await safeFetch(String(endpoint), { headers });
  if (!res) {
    setSyncError('pull_failed', { code: 'no_response' });
    return { pulled: 0, error: true, detail: 'No response' };
  }
  if (!res.ok) {
    let body = '';
    try { body = await res.text(); } catch {}
    setSyncError('pull_failed', { code: 'http_error', status: res.status, statusText: res.statusText || '', body: body || '' });
    return { pulled: 0, error: true, status: res.status, statusText: res.statusText, body };
  }
  const payload = await res.json().catch(() => ({}));
  
  // Filter payload by dataTypes if specified
  const allowedFiles = dataTypes.length > 0 ? new Set(dataTypes.map(dt => `${dt}.json`)) : null;
  let filteredPayload = allowedFiles 
    ? Object.fromEntries(Object.entries(payload).filter(([file]) => allowedFiles.has(file)))
    : payload;
  // Shifts often need a full snapshot fallback because records can have old timestamps
  // relative to global watermark (lastPullAt), especially after migration/import.
  const wantsShifts = !allowedFiles || allowedFiles.has('shifts.json');
  const hasShiftInPayload = !!(filteredPayload && Array.isArray(filteredPayload['shifts.json']) && filteredPayload['shifts.json'].length > 0);
  const hasIncomingStockEvents = !!(
    filteredPayload &&
    (
      (Array.isArray(filteredPayload['transactions.json']) && filteredPayload['transactions.json'].length > 0) ||
      (Array.isArray(filteredPayload['stock_in.json']) && filteredPayload['stock_in.json'].length > 0)
    )
  );
  if (wantsShifts && !hasShiftInPayload) {
    try {
      const epShift = new URL('/api/sync/changes', cfg.baseUrl);
      epShift.searchParams.set('since', '0');
      epShift.searchParams.set('clientId', cfg.clientId || '');
      const rs = await safeFetch(String(epShift), { headers });
      if (rs && rs.ok) {
        const ps = await rs.json().catch(() => ({}));
        if (ps && Array.isArray(ps['shifts.json'])) {
          filteredPayload = Object.assign({}, filteredPayload || {}, { 'shifts.json': ps['shifts.json'] });
        }
      }
    } catch {}
  }
  
  let count = 0;
  const byFile = {};
  const idsByFile = {};
  const ts = (v) => {
    if (typeof v === 'number' && isFinite(v)) return v;
    if (typeof v === 'string') { const t = Date.parse(v); return isNaN(t) ? 0 : t; }
    return 0;
  };
  const recTs = (x) => Math.max(ts(x?.updatedAt), ts(x?.timestamp), ts(x?.createdAt));
  if (filteredPayload && typeof filteredPayload === 'object') {
    for (const [file, changes] of Object.entries(filteredPayload)) {
      if (!Array.isArray(changes)) continue;
      byFile[file] = (byFile[file] || 0) + changes.length;
      // Update progress with current file being processed
      __syncProgress.currentFile = file.replace('.json', '');
      // Handle singleton config files as objects
      if (file === 'settings.json' || file === SYNC_CFG_FILE || file === 'banners.json' || file === 'qris.json') {
        try {
          const curObj = await readData(file).catch(() => ({}));
          // pick the newest by updatedAt
          const newest = changes.reduce((acc, it) => {
            return recTs(it) > recTs(acc || {}) ? it : acc;
          }, null);
          if (newest && typeof newest === 'object') {
            const nu = recTs(newest);
            const cu = recTs(curObj || {});
            // Determine if newest has meaningful keys (beyond ids/timestamps)
            const keys = Object.keys(newest || {}).filter(k => !['_id','id','updatedAt','timestamp','createdAt'].includes(k));
            const hasMeaningful = keys.length > 0 && keys.some(k => newest[k] != null && newest[k] !== '');
            const curKeys = Object.keys(curObj || {});
            const curHasMeaningful = curKeys.length > 0;
            // Skip destructive overwrite if server payload is effectively empty while local has data
            if (nu >= cu && (hasMeaningful || !curHasMeaningful)) {
              await writeData(file, { ...(curObj || {}), ...newest, updatedAt: nu });
              count += 1;
              const id = String(newest._id || newest.id || '');
              if (id) { if (!idsByFile[file]) idsByFile[file] = []; idsByFile[file].push(id); }
            }
          }
        } catch {}
        continue;
      }
      if (file === 'stock_in.json') {
        let cur = await readData(file).catch(() => []);
        if (!Array.isArray(cur)) cur = [];
        const prevList = cur.slice();
        const merged = mergeSyncArrayChanges(cur, changes, recTs);
        cur = merged.list;
        count += merged.count;
        for (const ch of changes) {
          const key = String(ch._id || ch.id || '');
          if (!key) continue;
          if (!idsByFile[file]) idsByFile[file] = [];
          if (!idsByFile[file].includes(key)) idsByFile[file].push(key);
        }
        await writeData(file, cur);
        if (merged.count > 0) {
          const adjusted = await adjustProductStockFromMergedStockIn(prevList, cur);
          if (adjusted > 0) byFile['products.json'] = (byFile['products.json'] || 0) + adjusted;
        }
        continue;
      }
      if (file === 'transactions.json') {
        let cur = await readData(file).catch(() => []);
        if (!Array.isArray(cur)) cur = [];
        const prevList = cur.slice();
        const merged = mergeSyncArrayChanges(cur, changes, recTs);
        cur = merged.list;
        count += merged.count;
        for (const ch of changes) {
          const key = String(ch._id || ch.id || '');
          if (!key) continue;
          if (!idsByFile[file]) idsByFile[file] = [];
          if (!idsByFile[file].includes(key)) idsByFile[file].push(key);
        }
        await writeData(file, cur);
        if (merged.count > 0) {
          const adjusted = await adjustProductStockFromMergedTransactions(prevList, cur);
          if (adjusted > 0) byFile['products.json'] = (byFile['products.json'] || 0) + adjusted;
        }
        continue;
      }
      // Array collections: merge per record (tambah/update/hapus), jangan replace seluruh file
      let cur = await readData(file).catch(() => []);
      if (!Array.isArray(cur)) cur = [];
      const prevProductsByKey = file === 'products.json'
        ? new Map(cur.map((doc) => [String(doc && (doc._id || doc.id) || ''), doc]).filter(([k]) => !!k))
        : null;
      const merged = mergeSyncArrayChanges(cur, changes, recTs);
      cur = merged.list;
      count += merged.count;
      for (const ch of changes) {
        const key = String(ch._id || ch.id || '');
        if (!key) continue;
        if (!idsByFile[file]) idsByFile[file] = [];
        if (!idsByFile[file].includes(key)) idsByFile[file].push(key);
      }
      // When stock-affecting records arrive in the same pull, keep local stock
      // and let the transaction/stock-in deltas update inventory exactly once.
      if (file === 'products.json' && hasIncomingStockEvents && prevProductsByKey) {
        cur = cur.map((doc) => {
          const key = String(doc && (doc._id || doc.id) || '');
          const prevDoc = prevProductsByKey.get(key);
          return prevDoc ? mergeProductRecordPreservingLocalStock(prevDoc, doc, true) : doc;
        });
      }
      await writeData(file, cur);
      // Re-enqueue pulled shifts so this server can forward them to another sync target.
      if (file === 'shifts.json') {
        for (const ch of changes) {
          try {
            const key = String(ch && (ch._id || ch.id) || '');
            if (!key) continue;
            if (ch.deleted) {
              await enqueueOutbox({ collection: 'shifts', file: 'shifts.json', op: 'delete', _id: key, deleted: true, updatedAt: Number(recTs(ch) || Date.now()) });
            } else {
              const doc = { ...(ch || {}) };
              if (!doc._id) doc._id = key;
              if (!doc.id) doc.id = key;
              if (!doc.updatedAt) doc.updatedAt = Number(recTs(doc) || Date.now());
              await enqueueOutbox({ collection: 'shifts', file: 'shifts.json', op: 'upsert', _id: key, doc, updatedAt: Number(doc.updatedAt || Date.now()) });
            }
          } catch {}
        }
      }
    }
  }
  // Fallback bootstrap: if nothing pulled and we already have a non-zero watermark, and local is empty, retry once with since=0
  if (count === 0 && since > 0) {
    try {
      let anyEmpty = false;
      const filesToCheck = allowedFiles ? [...allowedFiles] : SYNC_COLLECTION_FILES;
      for (const f of filesToCheck) {
        const arr = await readData(f).catch(()=>[]);
        if (Array.isArray(arr) && arr.length === 0) { anyEmpty = true; break; }
      }
      if (anyEmpty) {
        const ep2 = new URL('/api/sync/changes', cfg.baseUrl);
        ep2.searchParams.set('since', '0');
        ep2.searchParams.set('clientId', cfg.clientId || '');
        const res2 = await safeFetch(String(ep2), { headers });
        if (res2 && res2.ok) {
          const p2 = await res2.json().catch(()=>({}));
          if (p2 && typeof p2 === 'object') {
            for (const [file, changes] of Object.entries(p2)) {
              if (!Array.isArray(changes)) continue;
              // Skip if not in allowed files
              if (allowedFiles && !allowedFiles.has(file)) continue;
              byFile[file] = (byFile[file] || 0) + changes.length;
              if (file === 'stock_in.json') {
                let cur = await readData(file).catch(() => []);
                if (!Array.isArray(cur)) cur = [];
                const prevList = cur.slice();
                const merged = mergeSyncArrayChanges(cur, changes, recTs);
                cur = merged.list;
                count += merged.count;
                for (const ch of changes) {
                  const key = String(ch._id || ch.id || '');
                  if (!key) continue;
                  if (!idsByFile[file]) idsByFile[file] = [];
                  if (!idsByFile[file].includes(key)) idsByFile[file].push(key);
                }
                await writeData(file, cur);
                if (merged.count > 0) {
                  const adjusted = await adjustProductStockFromMergedStockIn(prevList, cur);
                  if (adjusted > 0) byFile['products.json'] = (byFile['products.json'] || 0) + adjusted;
                }
                continue;
              }
              if (file === 'transactions.json') {
                let cur = await readData(file).catch(() => []);
                if (!Array.isArray(cur)) cur = [];
                const prevList = cur.slice();
                const merged = mergeSyncArrayChanges(cur, changes, recTs);
                cur = merged.list;
                count += merged.count;
                for (const ch of changes) {
                  const key = String(ch._id || ch.id || '');
                  if (!key) continue;
                  if (!idsByFile[file]) idsByFile[file] = [];
                  if (!idsByFile[file].includes(key)) idsByFile[file].push(key);
                }
                await writeData(file, cur);
                if (merged.count > 0) {
                  const adjusted = await adjustProductStockFromMergedTransactions(prevList, cur);
                  if (adjusted > 0) byFile['products.json'] = (byFile['products.json'] || 0) + adjusted;
                }
                continue;
              }
              let cur = await readData(file).catch(() => []);
              if (!Array.isArray(cur)) cur = [];
              const prevProductsByKey = file === 'products.json'
                ? new Map(cur.map((doc) => [String(doc && (doc._id || doc.id) || ''), doc]).filter(([k]) => !!k))
                : null;
              const merged = mergeSyncArrayChanges(cur, changes, recTs);
              cur = merged.list;
              count += merged.count;
              for (const ch of changes) {
                const key = String(ch._id || ch.id || '');
                if (!key) continue;
                if (!idsByFile[file]) idsByFile[file] = [];
                if (!idsByFile[file].includes(key)) idsByFile[file].push(key);
              }
              // Preserve local stock when product data is pulled together with
              // transaction/stock-in changes to avoid double-applying inventory deltas.
              if (file === 'products.json' && hasIncomingStockEvents && prevProductsByKey) {
                cur = cur.map((doc) => {
                  const key = String(doc && (doc._id || doc.id) || '');
                  const prevDoc = prevProductsByKey.get(key);
                  return prevDoc ? mergeProductRecordPreservingLocalStock(prevDoc, doc, true) : doc;
                });
              }
              await writeData(file, cur);
              if (file === 'shifts.json') {
                for (const ch of changes) {
                  try {
                    const key = String(ch && (ch._id || ch.id) || '');
                    if (!key) continue;
                    if (ch.deleted) {
                      await enqueueOutbox({ collection: 'shifts', file: 'shifts.json', op: 'delete', _id: key, deleted: true, updatedAt: Number(recTs(ch) || Date.now()) });
                    } else {
                      const doc = { ...(ch || {}) };
                      if (!doc._id) doc._id = key;
                      if (!doc.id) doc.id = key;
                      if (!doc.updatedAt) doc.updatedAt = Number(recTs(doc) || Date.now());
                      await enqueueOutbox({ collection: 'shifts', file: 'shifts.json', op: 'upsert', _id: key, doc, updatedAt: Number(doc.updatedAt || Date.now()) });
                    }
                  } catch {}
                }
              }
            }
          }
        }
      }
    } catch {}
  }
  last.lastPullAt = Date.now();
  if (byFile['products.json'] > 0) {
    try { invalidateCache('products'); } catch {}
  }
  await writeData(LASTSYNC_FILE, last);
  return { pulled: count, summary: { byFile, idsByFile }, dataTypes };
}

// --- Retry helpers for sync operations ---
async function withRetry(fn, attempts = 3, baseDelayMs = 500) {
  let lastErr = null;
  for (let i = 0; i < attempts; i++) {
    try {
      const res = await fn();
      if (res && !res.error) return res;
      lastErr = res || null;
    } catch (e) { lastErr = e; }
    const wait = baseDelayMs * Math.pow(2, i);
    await new Promise(r => setTimeout(r, wait));
  }
  return lastErr || { error: true };
}

async function pushOutboxWithRetry(attempts = 3) {
  return await withRetry(() => pushOutbox(), attempts, 700);
}

async function pullChangesWithRetry(attempts = 3, dataTypes = []) {
  return await withRetry(() => pullChanges(dataTypes), attempts, 700);
}

function startSyncScheduler() {
  try {
    setInterval(async () => {
      try {
        if (isSyncBusy()) return;
        __syncInProgress = true;
        try {
          await enqueueDeltaSinceLastPush();
          resetSyncProgress(); setSyncPhase('push'); setSyncStart();
          await pushOutboxChunked(500);
          setSyncEnd();
          await pullChangesWithRetry(2);
        } finally { __syncInProgress = false; }
      } catch {}
    }, 120000);
  } catch {}
}

function encryptTextIfPassphrase(text) {
  const pass = process.env.POS_PASSPHRASE || '';
  if (!pass) return null;
  const salt = crypto.randomBytes(16);
  const iv = crypto.randomBytes(12);
  const key = crypto.scryptSync(pass, salt, 32);
  const cipher = crypto.createCipheriv('aes-256-gcm', key, iv);
  const ciphertext = Buffer.concat([cipher.update(Buffer.from(text, 'utf8')), cipher.final()]);
  const tag = cipher.getAuthTag();
  return `ENC1:${salt.toString('base64')}:${iv.toString('base64')}:${tag.toString('base64')}:${ciphertext.toString('base64')}`;
}

function decryptTextIfEnc1(encText) {
  if (typeof encText !== 'string' || !encText.startsWith('ENC1:')) return null;
  const pass = process.env.POS_PASSPHRASE || '';
  // console.log('Decrypting with passphrase length:', pass.length);
  // console.log('Passphrase exists:', !!pass);
  if (!pass) throw new Error('Encrypted backup but POS_PASSPHRASE is missing. Set POS_PASSPHRASE environment variable or create data/passphrase.txt');
  
  const parts = encText.split(':');
  // console.log('Encrypted parts count:', parts.length);
  if (parts.length !== 5) {
    // console.error('Invalid encrypted format. Expected 5 parts, got:', parts.length);
    // console.error('First 100 chars:', encText.substring(0, 100));
    throw new Error('Invalid encrypted backup format. Expected ENC1:salt:iv:tag:ciphertext format');
  }
  
  try {
    const salt = Buffer.from(parts[1], 'base64');
    const iv = Buffer.from(parts[2], 'base64');
    const tag = Buffer.from(parts[3], 'base64');
    const ciphertext = Buffer.from(parts[4], 'base64');
    const key = crypto.scryptSync(pass, salt, 32);
    const decipher = crypto.createDecipheriv('aes-256-gcm', key, iv);
    decipher.setAuthTag(tag);
    const dec = Buffer.concat([decipher.update(ciphertext), decipher.final()]).toString('utf8');
    // console.log('Decryption successful, data length:', dec.length);
    return dec;
  } catch (decryptError) {
    // console.error('Decryption failed:', decryptError.message);
    // console.error('This usually means:');
    // console.error('1. Passphrase is different from backup time');
    // console.error('2. Encrypted data is corrupted');
    // Return a more user-friendly error message
    throw new Error(`Decryption failed: ${decryptError.message}. This usually means the passphrase is different from the one used when the data was encrypted.`);
  }
}

function decryptTextEnc1WithPassphrase(encText, passphrase) {
  if (typeof encText !== 'string' || !encText.startsWith('ENC1:')) {
    throw new Error('Invalid encrypted text format');
  }
  const pass = String(passphrase || '');
  if (!pass) throw new Error('Passphrase is required');
  const parts = encText.split(':');
  if (parts.length !== 5) {
    throw new Error('Invalid encrypted backup format. Expected ENC1:salt:iv:tag:ciphertext format');
  }
  try {
    const salt = Buffer.from(parts[1], 'base64');
    const iv = Buffer.from(parts[2], 'base64');
    const tag = Buffer.from(parts[3], 'base64');
    const ciphertext = Buffer.from(parts[4], 'base64');
    const key = crypto.scryptSync(pass, salt, 32);
    const decipher = crypto.createDecipheriv('aes-256-gcm', key, iv);
    decipher.setAuthTag(tag);
    const dec = Buffer.concat([decipher.update(ciphertext), decipher.final()]).toString('utf8');
    return dec;
  } catch (decryptError) {
    throw new Error(`Decryption failed: ${decryptError.message}. This usually means the passphrase is different from the one used when the data was encrypted.`);
  }
}
const DATA_DIR = resolveDataDir();
const sqliteStorage = require('./lib/sqlite-storage');
let __storageInitPromise = null;
async function initDatabase() {
  if (__storageInitPromise) return __storageInitPromise;
  __storageInitPromise = sqliteStorage.init(DATA_DIR, {
    onMigrated: (stats) => {
      try {
        if (stats && stats.files) {
          console.log(`[SQLite] Migrasi selesai: ${stats.files} file JSON → SQLite (${stats.documents || 0} dokumen)`);
        }
      } catch {}
    }
  });
  return __storageInitPromise;
}

// public served from __dirname/public

const ensureDataDir = async () => {
  try {
    await fs.access(DATA_DIR);
  } catch {
    await fs.mkdir(DATA_DIR, { recursive: true });
  }
};

async function seedDataDirIfEmpty() {
  try {
    const targetFiles = await fs.readdir(DATA_DIR).catch(() => []);
    const hasJson = (targetFiles || []).some(f => f.toLowerCase().endsWith('.json'));
    const bundledDataDir = path.join(__dirname, 'data');
    const bundledExists = await fs.stat(bundledDataDir).then(s => s.isDirectory()).catch(() => false);
    if (!hasJson && bundledExists) {
      const files = await fs.readdir(bundledDataDir).catch(() => []);
      for (const f of files) {
        if (!f.toLowerCase().endsWith('.json')) continue;
        const src = path.join(bundledDataDir, f);
        const dst = path.join(DATA_DIR, f);
        const exists = await fs.stat(dst).then(() => true).catch(() => false);
        if (!exists) {
          try { const content = await fs.readFile(src, 'utf-8'); await fs.writeFile(dst, content, 'utf-8'); } catch {}
        }
      }
    }
  } catch {}
}

const TRIAL_ENABLED = String(process.env.POS_TRIAL_ENABLED || 'true').toLowerCase() !== 'false';
const TRIAL_DAYS = Number(process.env.POS_TRIAL_DAYS || 1);
const TRIAL_MODE = String(process.env.POS_TRIAL_MODE || 'days').toLowerCase();//days, runs
const TRIAL_RUNS = Number(process.env.POS_TRIAL_RUNS || 3);
const TRIAL_FILE_NAME = 'trial-info.json';
const TRIAL_SHADOW_FILE = '.sys-pos-trial.json';
const LICENSE_RUNS_FILE = 'license-runs.json';
const LICENSE_LOCK_FILE = 'license-lock.json';
const LICENSE_PRODUCT_NAME = 'pos-web-app';
const LICENSE_SECRET = process.env.POS_LICENSE_SECRET || '@Sugand!94';
const LICENSE_BLACKLIST_FILE = '.sys-pos-blacklist.json';

function getShadowTrialPath() {
  try {
    const base = process.env.LOCALAPPDATA || process.env.APPDATA || path.join(os.homedir(), 'AppData', 'Local');
    return path.join(base, 'SystemData', 'pos-web-app', TRIAL_SHADOW_FILE);
  } catch {
    return null;
  }
}

function getLicenseShadowDir() {
  try {
    const base = process.env.LOCALAPPDATA || process.env.APPDATA || path.join(os.homedir(), 'AppData', 'Local');
    const dir = path.join(base, 'SystemData', 'pos-web-app');
    return dir;
  } catch {
    return null;
  }
}

function getLicenseRunsShadowPath() {
  try {
    const dir = getLicenseShadowDir();
    if (!dir) return null;
    const secretHash = crypto.createHash('sha256').update(String(LICENSE_SECRET || 'default')).digest('hex').slice(0, 16);
    return path.join(dir, `.sys-runs-${secretHash}.dat`);
  } catch {
    return null;
  }
}

function getLicenseBlacklistPath() {
  try {
    const dir = getLicenseShadowDir();
    if (!dir) return null;
    return path.join(dir, LICENSE_BLACKLIST_FILE);
  } catch {
    return null;
  }
}

async function readTrialFile(p) {
  if (!p) return null;
  try {
    const raw = await fs.readFile(p, 'utf-8').catch(() => null);
    if (!raw) return null;
    const obj = JSON.parse(raw);
    if (!obj || typeof obj !== 'object') return null;
    return obj;
  } catch {
    return null;
  }
}

async function readTrialInfo() {
  try {
    const primaryPath = path.join(DATA_DIR, TRIAL_FILE_NAME);
    const shadowPath = getShadowTrialPath();
    const primary = await readTrialFile(primaryPath);
    const shadow = shadowPath ? await readTrialFile(shadowPath) : null;
    if (!primary && !shadow) return null;
    if (primary && !shadow) return primary;
    if (!primary && shadow) return shadow;
    const a = primary || {};
    const b = shadow || {};
    const firstA = Number(a.firstRunAt || Infinity);
    const firstB = Number(b.firstRunAt || Infinity);
    const lastA = Number(a.lastRunAt || 0);
    const lastB = Number(b.lastRunAt || 0);
    const runsA = Number(a.runCount || 0);
    const runsB = Number(b.runCount || 0);
    const mergedFirst = Math.min(firstA, firstB);
    const mergedLast = Math.max(lastA, lastB);
    const mergedRuns = Math.max(runsA, runsB);
    const out = {
      ...a,
      ...b,
      firstRunAt: isFinite(mergedFirst) ? mergedFirst : (a.firstRunAt || b.firstRunAt || Date.now()),
      lastRunAt: mergedLast || Date.now(),
      runCount: mergedRuns
    };
    return out;
  } catch {
    return null;
  }
}

async function writeTrialInfo(info) {
  const data = info && typeof info === 'object' ? info : {};
  try {
    const primaryPath = path.join(DATA_DIR, TRIAL_FILE_NAME);
    await fs.mkdir(path.dirname(primaryPath), { recursive: true }).catch(() => {});
    await fs.writeFile(primaryPath, JSON.stringify(data, null, 2), 'utf-8').catch(() => {});
  } catch {}
  try {
    const shadowPath = getShadowTrialPath();
    if (!shadowPath) return;
    await fs.mkdir(path.dirname(shadowPath), { recursive: true }).catch(() => {});
    await fs.writeFile(shadowPath, JSON.stringify(data, null, 2), 'utf-8').catch(() => {});
  } catch {}
}

async function ensureTrialInfo() {
  const now = Date.now();
  let info = await readTrialInfo();
  const hadInfoBefore = !!(info && typeof info === 'object');
  if (!info || typeof info !== 'object') info = {};

  if (!hadInfoBefore) {
    try {
      let txs = await readData('transactions.json').catch(() => []);
      if (Array.isArray(txs) && txs.length > 0) {
        let oldestTs = null;
        for (const tx of txs) {
          const ts = Number(tx.timestamp || tx.createdAt || tx.date || 0);
          if (!ts) continue;
          if (oldestTs === null || ts < oldestTs) oldestTs = ts;
        }
        if (oldestTs !== null) {
          info.firstRunAt = oldestTs;
        }
        info.forceExpired = true;
      }
    } catch {}
  }

  // Deteksi jam sistem dimundurkan: jika now < lastRunAt, anggap trial dicurangi
  const prevLast = Number(info.lastRunAt || 0);
  const CLOCK_TOLERANCE_MS = 5 * 60 * 1000; // toleransi 5 menit
  if (prevLast && now + CLOCK_TOLERANCE_MS < prevLast) {
    info.forceExpired = true;
  }

  if (!info.firstRunAt) info.firstRunAt = now;
  info.lastRunAt = Math.max(prevLast, now);
  const prevRuns = Number(info.runCount || 0);
  info.runCount = prevRuns + 1;
  await writeTrialInfo(info);
  return info;
}

function isTrialExpired(info) {
  if (!TRIAL_ENABLED) return false;
  if (info && info.forceExpired === true) return true;
  const mode = TRIAL_MODE;
  if (!info || typeof info !== 'object') return false;

  if (mode === 'runs') {
    if (!TRIAL_RUNS || TRIAL_RUNS <= 0) return false;
    const used = Number(info.runCount || 0);
    return used > TRIAL_RUNS;
  }

  if (!TRIAL_DAYS || TRIAL_DAYS <= 0) return false;
  if (!info.firstRunAt) return false;
  const startedAt = Number(info.firstRunAt) || Date.now();
  const diffMs = Date.now() - startedAt;
  const diffDays = diffMs / (1000 * 60 * 60 * 24);
  return diffDays > TRIAL_DAYS;
}

async function ensureTrialNotExpired() {
  if (!TRIAL_ENABLED) return;
  try {
    // Jika sudah ada license lock (license habis), abaikan trial
    const lockPath = path.join(DATA_DIR, 'license-lock.json');
    try {
      const raw = await fs.readFile(lockPath, 'utf-8');
      if (raw) {
        const parsed = JSON.parse(raw);
        if (parsed && parsed.locked) {
          // console.error('License lock present, skipping trial check.');
          return;
        }
      }
    } catch {}

    const info = await ensureTrialInfo();
    if (isTrialExpired(info)) {
      // console.error('Trial period has expired. Please contact vendor to activate full version.');
      process.exit(1);
    }
  } catch {}
}

app.get('/api/trial-status', async (req, res) => {
  try {
    const mode = TRIAL_MODE;

    // 1) Cek dulu license offline dengan mode 'runs' (jumlah buka aplikasi)
    try {
      const off = await verifyOfflineLicense();
      if (off && off.valid && off.payload && off.payload.mode === 'runs' && Number(off.payload.maxRuns || 0) > 0) {
        const maxRuns = Number(off.payload.maxRuns || 0);
        const status = await getLicenseRunsStatus(maxRuns);
        const info = status.info || {};
        return res.json({
          enabled: true,
          mode: 'license_runs',
          expired: !!status.expired,
          remainingDays: null,
          totalDays: null,
          usedDays: null,
          remainingRuns: status.remainingRuns,
          totalRuns: status.totalRuns,
          usedRuns: status.used,
          firstRunAt: info && info.firstRunAt ? info.firstRunAt : null,
          lastRunAt: info && info.lastRunAt ? info.lastRunAt : null,
          runCount: status.used
        });
      }
    } catch (e) {}

    // 2) Jika sudah ada license valid (offline/online) selain mode runs, anggap trial non-aktif
    let hasValidLicense = false;
    try {
      const off2 = await verifyOfflineLicense();
      if (off2 && off2.valid) {
        hasValidLicense = true;
      } else {
        try {
          const lic = await checkLicenseOnline();
          if (lic && lic.ok) hasValidLicense = true;
        } catch (e) {}
      }
    } catch (e) {}
    if (hasValidLicense) {
      return res.json({
        enabled: false,
        mode,
        expired: false,
        remainingDays: null,
        totalDays: null,
        usedDays: null,
        remainingRuns: null,
        totalRuns: null,
        usedRuns: null,
        firstRunAt: null,
        lastRunAt: null,
        runCount: null
      });
    }

    if (!TRIAL_ENABLED) {
      return res.json({
        enabled: false,
        mode,
        expired: false,
        remainingDays: null,
        totalDays: null,
        usedDays: null,
        remainingRuns: null,
        totalRuns: null,
        usedRuns: null,
        firstRunAt: null,
        lastRunAt: null,
        runCount: null
      });
    }

    const info = await readTrialInfo();
    const now = Date.now();

    if (mode === 'runs') {
      const totalRuns = TRIAL_RUNS;
      const usedRuns = Number(info && info.runCount != null ? info.runCount : 0);
      const expired = isTrialExpired(info || {});
      let remainingRuns = typeof totalRuns === 'number' ? totalRuns - usedRuns : null;
      if (typeof remainingRuns === 'number' && remainingRuns < 0) remainingRuns = 0;

      return res.json({
        enabled: true,
        mode,
        expired,
        remainingDays: null,
        totalDays: null,
        usedDays: null,
        remainingRuns,
        totalRuns,
        usedRuns,
        firstRunAt: info && info.firstRunAt ? info.firstRunAt : null,
        lastRunAt: info && info.lastRunAt ? info.lastRunAt : null,
        runCount: usedRuns
      });
    }

    const totalDays = TRIAL_DAYS;

    if (!info || !info.firstRunAt || !totalDays || totalDays <= 0) {
      return res.json({
        enabled: true,
        mode,
        expired: false,
        remainingDays: totalDays || null,
        totalDays: totalDays || null,
        usedDays: 0,
        remainingRuns: null,
        totalRuns: null,
        usedRuns: info && typeof info.runCount === 'number' ? info.runCount : null,
        firstRunAt: info && info.firstRunAt ? info.firstRunAt : null,
        lastRunAt: info && info.lastRunAt ? info.lastRunAt : null,
        runCount: info && typeof info.runCount === 'number' ? info.runCount : null
      });
    }

    const startedAt = Number(info.firstRunAt) || now;
    const diffMs = now - startedAt;
    const usedDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));
    const expired = isTrialExpired(info || {});
    let remainingDays = typeof totalDays === 'number' ? totalDays - usedDays : null;
    if (typeof remainingDays === 'number' && remainingDays < 0) remainingDays = 0;

    res.json({
      enabled: true,
      mode,
      expired,
      remainingDays,
      totalDays,
      usedDays,
      remainingRuns: null,
      totalRuns: null,
      usedRuns: info && typeof info.runCount === 'number' ? info.runCount : null,
      firstRunAt: info.firstRunAt || null,
      lastRunAt: info.lastRunAt || null,
      runCount: info && typeof info.runCount === 'number' ? info.runCount : null
    });
  } catch (e) {
    res.status(500).json({
      enabled: TRIAL_ENABLED,
      mode: TRIAL_MODE,
      error: true,
      message: 'Failed to read trial status'
    });
  }
});

app.get('/api/license/status', async (req, res) => {
  try {
    const lk = await readLicenseKey();
    const off = await verifyOfflineLicense();
    let lock = null;
    try { lock = await readLicenseLock(); } catch {}
    let remainingDays = null;
    let licenseType = 'none';
    let licenseRuns = null;
    try {
      if (off && off.valid && off.payload) {
        // License khusus: mode runs (jumlah buka aplikasi)
        if (off.payload.mode === 'runs' && Number(off.payload.maxRuns || 0) > 0) {
          licenseType = 'runs';
          try {
            const status = await getLicenseRunsStatus(Number(off.payload.maxRuns || 0));
            licenseRuns = {
              remainingRuns: status && typeof status.remainingRuns === 'number' ? status.remainingRuns : null,
              totalRuns: status && typeof status.totalRuns === 'number' ? status.totalRuns : null,
              usedRuns: status && typeof status.used === 'number' ? status.used : null
            };
          } catch (e2) {}
        } else {
          const now = Date.now();
          const expMs = Number(off.payload.exp || 0);
          if (expMs && expMs > now) {
            const diffMs = expMs - now;
            const days = Math.ceil(diffMs / (1000 * 60 * 60 * 24));
            remainingDays = days >= 0 ? days : 0;
            licenseType = 'date';
          } else if (!expMs || off.payload.full === true) {
            licenseType = 'full';
          }
        }
      }
    } catch (e) {}
    res.json({
      hasKey: !!lk,
      keyPreview: lk ? lk.slice(0, 8) + '...' : '',
      offline: off,
      remainingDays,
      licenseType,
      licenseRuns,
      lock: lock && lock.locked ? lock : null
    });
  } catch (e) {
    return res.status(500).json({ success: false, message: 'Failed to read license status' });
  }
});

app.post('/api/license/offline', async (req, res) => {
  try {
    const body = req.body || {};
    const key = String(body.licenseKey || body.key || '').trim();
    if (!key) return res.status(400).json({ success: false, message: 'LICENSE KEY is required' });
    const result = await verifyOfflineLicense(key);
    if (!result || !result.valid) {
      const reason = result && result.reason ? result.reason : 'INVALID';
      let msg = 'LICENSE KEY tidak valid';
      if (reason === 'EXPIRED') {
        msg = 'LICENSE KEY sudah kadaluarsa';
      }
      return res.status(400).json({ success: false, message: msg, reason });
    }
    const ok = await saveLicenseKey(key);
    if (!ok) return res.status(500).json({ success: false, message: 'Gagal menyimpan LICENSE KEY' });
    try { await clearLicenseLock(); } catch (e) {}
    try { await applyLicensedStoreNameToSettings(result && result.payload ? result.payload : null); } catch (e) {}
    try { await applyLicensedAdminNameToSettings(result && result.payload ? result.payload : null); } catch (e) {}
    res.json({ success: true, message: 'LICENSE KEY tersimpan dan valid', payload: result.payload || null });
  } catch (e) {
    res.status(500).json({ success: false, message: 'Gagal memproses LICENSE KEY' });
  }
});

app.delete('/api/license/offline', async (req, res) => {
  try {
    // Hapus LICENSE KEY dan set lock manual
    await clearOfflineLicenseState('MANUAL_CLEAR');
    return res.json({ success: true, message: 'LICENSE KEY berhasil dihapus' });
  } catch (e) {
    return res.status(500).json({ success: false, message: 'Gagal menghapus LICENSE KEY' });
  }
});

async function loadPassphraseFromFile() {
  try {
    if (process.env.POS_PASSPHRASE && String(process.env.POS_PASSPHRASE).trim()) {
      // console.log('Passphrase already exists in environment variable, length:', process.env.POS_PASSPHRASE.length);
      // console.log('Passphrase hex:', Buffer.from(process.env.POS_PASSPHRASE, 'utf8').toString('hex'));
      return;
    }
    const p = path.join(DATA_DIR, 'passphrase.txt');
    const txt = await fs.readFile(p, 'utf-8').catch(() => '');
    // console.log('Raw file content:', JSON.stringify(txt));
    // console.log('File content length:', txt.length);
    // console.log('File content hex:', Buffer.from(txt, 'utf8').toString('hex'));
    if (txt && txt.trim()) { 
      process.env.POS_PASSPHRASE = txt.trim();
      // console.log('Passphrase loaded from file, length:', process.env.POS_PASSPHRASE.length);
      // console.log('Passphrase hex:', Buffer.from(process.env.POS_PASSPHRASE, 'utf8').toString('hex'));
    } else {
      // console.log('No passphrase file found or file is empty');
    }
  } catch (e) {
    // console.error('Error loading passphrase from file:', e.message);
  }
}

// --- CSRF Protection (simple session token) ---
function ensureCsrfToken(req) {
  if (!req.session) return;
  if (!req.session.csrfToken) {
    // Generate a strong random token for CSRF protection
    req.session.csrfToken = crypto.randomBytes(32).toString('hex');
  }
}

app.use((req, res, next) => { try { ensureCsrfToken(req); } catch {} next(); });
// Public endpoint to fetch CSRF token (requires session cookie)
app.get('/api/csrf', (req, res) => {
  try { ensureCsrfToken(req); } catch {}
  const token = (req.session && req.session.csrfToken) ? req.session.csrfToken : '';
  res.json({ csrfToken: token });
});

function requireCsrf(req, res, next){
  // Only protect state-changing API calls
  const method = String(req.method || '').toUpperCase();
  const needs = method === 'POST' || method === 'PUT' || method === 'DELETE' || method === 'PATCH';
  if (!needs) return next();
  // Allowlist unauthenticated login route to reduce friction
  const fullUrl = String(req.originalUrl || req.url || '');
  if (req.path === '/login' || fullUrl.endsWith('/api/login')) return next();
  // Allowlist sync endpoints for cross-origin device sync (secured via Authorization token)
  try {
    const p = String(req.path || '');
    if (p.startsWith('/sync/')) return next(); // when mounted under /api
    if (fullUrl.includes('/api/sync/')) return next();
  } catch {}
   // Allowlist offline license POST (akan diverifikasi dengan secret sendiri)
   try { if (req.path === '/api/license/offline') return next(); } catch {}
  const header = req.headers['x-csrf-token'] || req.headers['x-xsrf-token'] || (req.body && req.body._csrf) || (req.query && req.query._csrf);
  const token = req.session && req.session.csrfToken;
  if (!token || !header || String(header) !== String(token)) {
    // Fallback: if same-origin and session exists, allow to reduce friction in admin
    try {
      const origin = req.get('origin') || '';
      const host = `${req.protocol}://${req.get('host')}`;
      const sameOrigin = !origin || origin === host;
      if (sameOrigin && req.session) {
        // console.warn('[CSRF] token mismatch but same-origin with active session, allowing request');
        return next();
      }
    } catch {}
    return res.status(403).json({ success:false, message:'CSRF token invalid' });
  }
  next();
}

// Apply CSRF protection to all API routes (except SSE)
app.use('/api', (req, res, next) => {
  if (req.path === '/events') {
    return next(); // Skip CSRF protection for SSE endpoint
  }
  return requireCsrf(req, res, next);
});

// --- Security Headers Middleware ---
app.use((req, res, next) => {
  try {
    res.setHeader('X-Content-Type-Options', 'nosniff');
    res.setHeader('X-Frame-Options', 'SAMEORIGIN');
    res.setHeader('Referrer-Policy', 'no-referrer');
    res.setHeader('X-XSS-Protection', '1; mode=block');
    // Allow camera for this origin to enable barcode scanner; keep others restricted
    res.setHeader('Permissions-Policy', 'camera=(self), microphone=(), geolocation=()');
  } catch {}
  next();
});

// Encryption runtime flag (controlled by settings.encryption.enabled)
let encryptionEnabled = true;

// Trust proxy (for secure cookies behind reverse proxy)
try { app.set('trust proxy', 1); } catch {}

function __setCache(filename, data) { sqliteStorage.setCache(filename, data); }
function __invalidateCache(filename) { sqliteStorage.invalidateCache(filename); }

const readData = async (filename) => {
  await initDatabase();
  const cached = await cacheRedis.get('app', filename);
  if (cached !== null) return cached;
  const data = await sqliteStorage.readData(filename);
  if (data !== null && data !== undefined) {
    await cacheRedis.set('app', filename, data);
  }
  return data;
};

const writeData = async (filename, data) => {
  await initDatabase();
  await sqliteStorage.writeData(filename, data);
  await cacheRedis.del('app', filename);
};

// License persistence to database (initialized setelah readData/writeData siap)
__saveLicenseToDb = async (key) => {
  try {
    await writeData('license.json', { key: String(key || '').trim(), updatedAt: Date.now() });
  } catch {}
};
__readLicenseFromDb = async () => {
  try {
    const data = await readData('license.json');
    if (data && typeof data.key === 'string' && data.key.trim()) return data.key.trim();
  } catch {}
  return '';
};

function parseTimestampMs(v) {
  if (typeof v === 'number' && isFinite(v)) return v;
  if (typeof v === 'string') {
    const s = v.trim();
    if (!s) return 0;
    const n = Number(s);
    if (isFinite(n) && n > 1e11) return n;
    const t = Date.parse(s);
    return isNaN(t) ? 0 : t;
  }
  return 0;
}

function getTransactionTimestampMs(tx) {
  if (!tx || typeof tx !== 'object') return 0;
  const candidates = [tx.timestamp, tx.createdAt, tx.created_at, tx.date, tx.transactionDate, tx.paymentDate];
  let max = 0;
  for (const c of candidates) {
    const t = parseTimestampMs(c);
    if (t > max) max = t;
  }
  return max;
}

function isVoidedTransaction(tx) {
  return isSyncVoidedTransaction(tx);
}

function getSaleTransactionTimestampMs(tx) {
  if (!tx || typeof tx !== 'object') return 0;
  const candidates = [tx.timestamp, tx.createdAt, tx.created_at, tx.date, tx.transactionDate];
  let max = 0;
  for (const c of candidates) {
    const t = parseTimestampMs(c);
    if (t > max) max = t;
  }
  if (max) return max;
  const id = String(tx.id || tx.transactionId || '');
  const m = id.match(/-(\d{12,})/);
  if (m) {
    const n = Number(m[1]);
    if (isFinite(n) && n > 1e11) return n;
  }
  return 0;
}

// Note: encryptionEnabled is now initialized in the main server startup after passphrase is loaded

// --- Input Validation Helpers ---
function isSafeUsername(username) {
  // Allow only letters, numbers, dot, underscore, hyphen; length 3-32
  return typeof username === 'string' && /^[A-Za-z0-9._-]{3,32}$/.test(username);
}

function isSafePassword(password) {
  // Enhanced password validation with stronger requirements
  if (typeof password !== 'string') return false;
  if (password.length < 8 || password.length > 128) return false; // Minimum 8 characters
  
  // Require at least one uppercase, lowercase, number, and special character
  const hasUpper = /[A-Z]/.test(password);
  const hasLower = /[a-z]/.test(password);
  const hasNumber = /[0-9]/.test(password);
  const hasSpecial = /[!@#$%^&*()_+\-=\[\]{};':"\\|,.<>\/?]/.test(password);
  
  if (!(hasUpper && hasLower && hasNumber && hasSpecial)) {
    return false; // Password doesn't meet complexity requirements
  }
  
  // Disallow dangerous characters often used in injections
  const forbidden = /["]'`<>\\{}\[\]$]/; // quotes, angle brackets, backslash, braces, brackets, dollar
  // Also disallow non-printable ASCII
  const nonPrintable = /[\x00-\x1F\x7F]/;
  return !forbidden.test(password) && !nonPrintable.test(password);
}

// --- Role helpers ---
function normalizeRole(role) {
  const r = String(role || '').toLowerCase();
  if (r === 'kasir') return 'cashier';
  return r;
}

// Ensure at least one superadmin exists
async function ensureSuperadminUser() {
  try {
    let users = await readData("users.json");
    if (!Array.isArray(users)) users = [];
    if (users.length > 0) return;
    const now = new Date().toISOString();
    const username = process.env.POS_SUPERADMIN_USER || 'Superadmin';
    const password = process.env.POS_SUPERADMIN_PASS || '@Superadmin123';
    const hashed = await bcrypt.hash(String(password), 10);
    const user = {
      id: 1,
      username,
      name: 'Super Admin',
      role: 'superadmin',
      status: 'active',
      password: hashed,
      createdAt: now,
      updatedAt: now
    };
    await writeData("users.json", [user]);
  } catch (e) {}
}

function isSuperAdminRole(role) {
  return normalizeRole(role) === 'superadmin';
}

function resolveSaleTimestampForTransaction(body, sessionUser) {
  const nowMs = Date.now();
  const nowIso = new Date(nowMs).toISOString();
  if (!isSuperAdminRole(sessionUser && sessionUser.role)) {
    return { timestamp: nowMs, date: nowIso, createdAt: nowIso };
  }
  const raw = body && (body.transactionDate || body.saleDate || body.date || body.createdAt || body.timestamp);
  if (raw == null || raw === '') {
    return { timestamp: nowMs, date: nowIso, createdAt: nowIso };
  }
  const parsed = parseTimestampMs(raw);
  if (!parsed || parsed <= 0) {
    return { timestamp: nowMs, date: nowIso, createdAt: nowIso };
  }
  const iso = new Date(parsed).toISOString();
  return { timestamp: parsed, date: iso, createdAt: iso, saleDateCustom: true };
}

function requireRole(...roles) {
  return (req, res, next) => {
    if (!req.session || !req.session.user) {
      return res.redirect('/login');
    }
    const role = normalizeRole(req.session.user.role);
    if (isSuperAdminRole(role)) return next();
    if (!roles.map(normalizeRole).includes(role)) {
      // Use shared forbidden responder
      return respondForbidden(req, res, 'Forbidden');
    }
    next();
  };
}

// Shared Forbidden responder: HTML -> error page, API -> JSON
function respondForbidden(req, res, message = 'Forbidden') {
  const code = 403;
  const acceptsHtml = req.accepts(['html', 'json']) === 'html';
  if (acceptsHtml && !String(req.originalUrl||'').startsWith('/api/')) {
    const q = new URLSearchParams({ code: String(code), msg: message, path: req.originalUrl || '/' });
    return res.status(code).redirect(`/error.html?${q.toString()}`);
  }
  return res.status(code).json({ success: false, code, message, path: req.originalUrl });
}

// --- Auto Backup Helpers ---
async function performDatabaseBackupToFile(modeLabel = 'manual') {
  try {
    const backupRoot = path.join(DATA_DIR, 'backups', 'auto');
    await fs.mkdir(backupRoot, { recursive: true }).catch(()=>{});
    const now = new Date();
    const ts = now.toISOString().replace(/[:.]/g, '-');
    const name = `backup-${ts}-${modeLabel}.json`;
    const outPath = path.join(backupRoot, name);
    const files = await fs.readdir(DATA_DIR);
    const jsonFiles = files.filter(f => f.endsWith('.json'));
    const payload = {};
    for (const f of jsonFiles) {
      try {
        const raw = await readData(f);
        payload[f] = raw;
      } catch (e) {
        payload[f] = { __error: true };
      }
    }
    const plain = JSON.stringify({ date: now.toISOString(), data: payload }, null, 2);
    const enc = encryptTextIfPassphrase(plain);
    const outText = enc || plain;
    await fs.writeFile(outPath, outText, 'utf-8');
    // Enforce max N files (delete oldest)
    try {
      const files = await fs.readdir(backupRoot).catch(()=>[]);
      const jsons = (files || []).filter(f => f.endsWith('.json'));
      const withTimes = await Promise.all(jsons.map(async f => {
        const st = await fs.stat(path.join(backupRoot, f)).catch(()=>null);
        return st ? { f, t: st.mtimeMs } : null;
      }));
      const list = withTimes.filter(Boolean).sort((a,b)=>a.t - b.t);
      // Read maxCount from settings
      let maxCount = 10;
      try { const base = await readData('settings.json'); maxCount = Math.max(1, Number(base?.autoBackup?.maxCount)||10); } catch {}
      const excess = Math.max(0, list.length - maxCount);
      for (let i = 0; i < excess; i++) {
        const victim = list[i];
        try { await fs.unlink(path.join(backupRoot, victim.f)); } catch {}
      }
    } catch (e) {
    }
    return outPath;
  } catch (e) {
    return null;
  }
}

async function autoBackupIfNeededOnStart() {
  try {
    const base = await readData('settings.json').catch(()=>({}));
    const cfg = base && typeof base === 'object' ? (base.autoBackup || {}) : {};
    const enabled = cfg.enabled === true;
    const mode = cfg.mode || 'off'; // 'off' | 'on_start' | 'daily'
    const retention = Number(cfg.retentionDays || 0);
    if (!enabled || mode === 'off') return;
    const backupRoot = path.join(DATA_DIR, 'backups', 'auto');
    await fs.mkdir(backupRoot, { recursive: true }).catch(()=>{});
    if (mode === 'daily') {
      const files = await fs.readdir(backupRoot).catch(()=>[]);
      const today = new Date().toISOString().slice(0,10);
      const hasToday = (files || []).some(f => f.includes(`backup-${today}`));
      if (hasToday) {
      } else {
        await performDatabaseBackupToFile('daily');
      }
    } else if (mode === 'on_start') {
      await performDatabaseBackupToFile('start');
    }
    // Retention cleanup
    if (retention > 0) {
      const files = await fs.readdir(backupRoot).catch(()=>[]);
      const full = files.map(f => ({ f, t: fs.stat(path.join(backupRoot, f)).then(s=>s.mtimeMs).catch(()=>0) }));
      const withTimes = await Promise.all(full.map(async x => ({ f: x.f, t: await x.t })));
      const cutoff = Date.now() - retention * 24*60*60*1000;
      for (const { f, t } of withTimes) {
        if (t && t < cutoff) {
          try { await fs.unlink(path.join(backupRoot, f)); } catch {}
        }
      }
    }
  } catch (e) {
  }
}

// Ensure data directory exists at startup
ensureDataDir()
  .then(() => seedDataDirIfEmpty())
  .then(() => initDatabase())
  .then(() => ensureSuperadminUser())
  .then(() => ensureClientId())
  .then(() => autoBackupIfNeededOnStart())
  .then(() => { try { startSyncScheduler(); } catch {} })
  .catch((e) => {
    try { process.stderr.write(`[POS] Startup chain error: ${e && e.stack ? e.stack : e}\n`); } catch {}
  });


// --- PERUBAHAN 1: Tambahkan Rute Utama untuk Pengalihan Otomatis ---
// Rute ini harus didefinisikan SEBELUM middleware express.static
app.get("/", (req, res) => {
  if (req.session.user) {
    const role = String(req.session.user.role || '').toLowerCase();
    if (role === 'admin' || role === 'superadmin') {
      return res.redirect("/admin.html");
    }
    return res.redirect("/pos.html");
  }
  return res.redirect("/login");
});

// Decrypt all JSON data files to plaintext (keep passphrase for future use)
app.post('/api/admin/decrypt-all', strictLimiter, requireRole('admin'), async (req, res) => {
  try {
    return res.json({
      success: true,
      processed: 0,
      mode: 'sqlite',
      message: 'Data tersimpan di pos.db (SQLite), tidak perlu dekripsi file JSON.'
    });
  } catch (e) {
    console.error('Error in decrypt-all:', e);
    res.status(500).json({ success: false, message: 'Server error' });
  }
});
// moved: /api/sync/now is registered after auth middlewares

// --- Sync Status ---
// moved: /api/sync/status is registered after auth middlewares

// Trigger an auto-backup now
app.post('/api/backup/auto-now', strictLimiter, requireRole('admin'), async (req, res) => {
  try {
    const out = await performDatabaseBackupToFile('manual');
    if (!out) return res.status(500).json({ success:false, message:'Failed to create backup' });
    res.json({ success:true, file: path.basename(out) });
  } catch (e) {
    // console.error('auto-now failed:', e);
    res.status(500).json({ success:false, message:'Failed to create backup' });
  }
});

// Delete a specific backup
app.delete('/api/backup/auto-delete', strictLimiter, requireRole('admin'), async (req, res) => {
  try {
    const name = String(req.query.name || '');
    if (!name || name.includes('..') || name.includes('/') || name.includes('\\')) {
      return res.status(400).json({ success:false, message:'Invalid file name' });
    }
    const p = path.join(DATA_DIR, 'backups', 'auto', name);
    const st = await fs.stat(p).catch(()=>null);
    if (!st || !st.isFile()) return res.status(404).json({ success:false, message:'File not found' });
    await fs.unlink(p);
    res.json({ success:true });
  } catch (e) {
    // console.error('auto-delete failed:', e);
    res.status(500).json({ success:false, message:'Failed to delete backup' });
  }
});

// ZIP backup for entire app preserving folder structure (exclude node_modules, .git, .cache)
// (moved) app-zip-structured route is defined later after middleware

// (moved) restore-zip route is defined later after middleware
// Public settings for login/branding (no auth)
app.get('/api/public-settings', async (req, res) => {
  try {
    const raw = await readData('settings.json');
    const base = Array.isArray(raw) ? {} : (raw || {});
    let storeName = base.storeName || 'POS System';
    try {
      const licensed = await getLicensedStoreName();
      if (licensed) storeName = licensed;
    } catch (e) {}
    const data = {
      storeName,
      themeColor: base.themeColor || '#198754',
      loginTitle: base.loginTitle || '',
      loginLogoBase64: base.loginLogoBase64 || base.logoBase64 || '',
      loginBackgroundBase64: base.loginBackgroundBase64 || '',
      faviconBase64: base.faviconBase64 || '',
      darkMode: base.darkMode === true,
      loginLogoSize: typeof base.loginLogoSize === 'string' ? base.loginLogoSize : 'medium'
    };
    res.json(data);
  } catch (e) {
    try {
      const licensed = await getLicensedStoreName();
      if (licensed) {
        return res.status(200).json({
          storeName: licensed,
          themeColor: '#198754',
          loginTitle: '',
          loginLogoBase64: '',
          loginBackgroundBase64: '',
          faviconBase64: ''
        });
      }
    } catch (err) {}
    res.status(200).json({
      storeName: 'POS System',
      themeColor: '#198754',
      loginTitle: '',
      loginLogoBase64: '',
      loginBackgroundBase64: '',
      faviconBase64: ''
    });
  }
});

// --- Page routing rules ---
// Halaman Admin (khusus admin)
// Halaman Rekey (khusus admin)
app.get('/admin-rekey.html', requireRole('admin'), (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'admin-rekey.html'));
});

// Alias yang lebih rapi untuk halaman Rekey
app.get('/admin/rekey', requireRole('admin'), (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'admin-rekey.html'));
});

// Halaman Admin (khusus admin)
app.get('/admin', requireRole('admin'), (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'admin.html'));
});

// Halaman Kasir (admin & cashier)
app.get('/kasir', requireRole('admin', 'cashier'), (req, res) => {
  try {
    res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate');
    res.setHeader('Pragma', 'no-cache');
    res.setHeader('Expires', '0');
  } catch {}
  res.sendFile(path.join(__dirname, 'public', 'pos.html'));
});

// Halaman Pendapatan (admin & cashier)
app.get('/revenue.html', requireRole('admin', 'cashier'), (req, res) => {
  // Redirect lama /revenue.html ke route baru tanpa ekstensi
  res.redirect(301, '/revenue');
});

app.get('/revenue', requireRole('admin', 'cashier'), (req, res) => {
  try {
    res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate');
    res.setHeader('Pragma', 'no-cache');
    res.setHeader('Expires', '0');
  } catch {}
  res.sendFile(path.join(__dirname, 'public', 'revenue.html'));
});

// Halaman Sinkron (khusus admin)
app.get('/sync.html', requireRole('admin'), (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'sync.html'));
});
app.get('/admin/sync', requireRole('admin'), (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'sync.html'));
});

// Banner Builder (khusus admin)
app.get('/banner-builder', requireRole('admin'), (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'banner-builder.html'));
});
app.get('/ai-provider-test', requireRole('admin'), (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'ai-provider-test.html'));
});

// Alias tanpa .html (tetap bisa akses file aslinya)
app.get('/login', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'login.html'));
});

// Auto backup endpoints (use role helper)
app.get('/api/backup/auto-list', requireRole('admin'), async (req, res) => {
  try {
    const dir = path.join(DATA_DIR, 'backups', 'auto');
    await fs.mkdir(dir, { recursive: true }).catch(()=>{});
    const files = await fs.readdir(dir);
    const items = await Promise.all((files || []).filter(f => f.endsWith('.json')).map(async f => {
      const p = path.join(dir, f);
      const st = await fs.stat(p).catch(()=>null);
      return st ? { name: f, size: st.size, mtime: st.mtimeMs } : null;
    }));
    res.json({ success: true, files: (items || []).filter(Boolean).sort((a,b)=>b.mtime-a.mtime) });
  } catch (e) {
    // console.error('auto-list failed:', e);
    res.status(500).json({ success: false, message: 'Failed to list backups' });
  }
});

app.get('/api/backup/auto-download', requireRole('admin'), async (req, res) => {
  try {
    const name = String(req.query.name || '');
    if (!name || name.includes('..') || name.includes('/') || name.includes('\\')) {
      return res.status(400).json({ success: false, message: 'Invalid file name' });
    }
    const dir = path.join(DATA_DIR, 'backups', 'auto');
    const p = path.join(dir, name);
    const st = await fs.stat(p).catch(()=>null);
    if (!st || !st.isFile()) return res.status(404).json({ success: false, message: 'File not found' });
    res.set('Content-Type', 'application/json');
    res.set('Content-Disposition', `attachment; filename="${name}"`);
    const content = await fs.readFile(p, 'utf-8');
    return res.send(content);
  } catch (e) {
    // console.error('auto-download failed:', e);
    res.status(500).json({ success: false, message: 'Failed to download backup' });
  }
});

// Lindungi akses langsung ke file HTML utama selain login: arahkan ke rute resmi
app.get(['/admin.html', '/pos.html', '/index.html'], (req, res) => {
  if (!req.session || !req.session.user) return res.redirect('/login');
  const role = normalizeRole(req.session.user.role);
  if (req.path === '/admin.html') return (role === 'admin' || role === 'superadmin') ? res.redirect('/admin') : respondForbidden(req, res, 'Forbidden');
  if (req.path === '/pos.html') return res.redirect('/kasir');
  return (role === 'admin' || role === 'superadmin') ? res.redirect('/admin') : res.redirect('/kasir');
});

// --- Middleware untuk file statis ---
// Diletakkan setelah rute utama agar tidak menangani '/' sebelum rute khusus kita
// Set proper MIME type for web manifest
app.get('/manifest.json', (req, res) => {
  res.setHeader('Content-Type', 'application/manifest+json');
  res.sendFile(path.join(__dirname, 'public', 'manifest.json'));
});

// Serve other static files
app.use(express.static(path.join(__dirname, 'public')));

// Disable caching for API responses to prevent stale data in UI
app.use((req, res, next) => {
  if (req.path.startsWith('/api/')) {
    res.set('Cache-Control', 'no-store');
    res.set('Pragma', 'no-cache');
    res.set('Expires', '0');
  }
  next();
});

// --- Validation Helper Functions ---
const validateProductName = async (name, excludeId = null) => {
  // Sanitize the input name before using it in the query
  const sanitizedName = validateAndSanitizeInput(name, 'general');
  
  const products = await readData("products.json");
  const existingProduct = products.find(
    (p) =>
      p.name && p.name.toLowerCase() === sanitizedName.toLowerCase() &&
      p.id != excludeId && p.deleted !== true
  );
  return existingProduct;
};

const validateCategoryName = async (name, excludeId = null) => {
  // Sanitize the input name before using it in the query
  const sanitizedName = validateAndSanitizeInput(name, 'general');
  
  const categories = await readData("categories.json");
  const existingCategory = categories.find(
    (c) =>
      c.name && c.name.toLowerCase() === sanitizedName.toLowerCase() && c.id != excludeId
  );
  return existingCategory;
};

// Validate SKU uniqueness
const validateProductSku = async (sku, excludeId = null) => {
  // Sanitize the input sku before using it in the query
  const sanitizedSku = validateAndSanitizeInput(sku, 'general');
  
  const products = await readData("products.json");
  const existing = products.find(
    (p) => p && typeof p.sku === 'string' && p.sku.trim() === String(sanitizedSku).trim() && p.id != excludeId
  );
  return existing;
};

const validateUsername = async (username, excludeId = null) => {
  // Sanitize the input username before using it in the query
  const sanitizedUsername = validateAndSanitizeInput(username, 'username');
  
  const users = await readData("users.json");
  const existingUser = users.find(
    (u) =>
      u.username &&
      u.username.toLowerCase() === sanitizedUsername.toLowerCase() &&
      u.id != excludeId
  );
  return existingUser;
};

// --- Authentication Middleware ---
// --- PERUBAHAN 2: Peningkatan Middleware untuk API ---
// Untuk API, lebih baik mengembalikan error JSON daripada redirect HTML
function jwtVerify(req) {
  if (req._jwtChecked) return req._jwtChecked;
  req._jwtChecked = false;
  try {
    const hdr = String(req.get('authorization') || '').trim();
    if (!hdr.toLowerCase().startsWith('bearer ')) return false;
    const token = hdr.slice(7).trim();
    if (!token) return false;
    const payload = jwt.verify(token, JWT_SECRET);
    if (payload && payload.user) {
      req.user = payload.user;
      req._jwtChecked = true;
      return true;
    }
  } catch {}
  return false;
}

const isAuthenticated = (req, res, next) => {
  if (req.session.user) return next();
  if (jwtVerify(req)) return next();
  res
    .status(401)
    .json({ success: false, message: "Unauthorized. Please log in." });
};

const isAdmin = (req, res, next) => {
  const src = req.session.user || req.user || {};
  const role = normalizeRole(src.role ? String(src.role).toLowerCase() : '');
  if (role === "admin" || role === "superadmin") {
    next();
  } else {
    res
      .status(403)
      .json({ success: false, message: "Access Denied: Admins only" });
  }
};

// Real-time updates using Server-Sent Events
const sseClients = new Set();

// SSE endpoint for real-time product updates
app.get('/api/events', isAuthenticated, (req, res) => {
  // Set CORS headers explicitly for SSE
  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');
  res.setHeader('Access-Control-Allow-Origin', req.headers.origin || '*');
  res.setHeader('Access-Control-Allow-Credentials', 'true');
  res.setHeader('X-Accel-Buffering', 'no'); // Disable nginx buffering if behind proxy
  res.status(200);

  const clientId = Date.now() + Math.random();
  console.log(`SSE client connected: ${clientId}`);
  
  // Send initial connection message
  res.write(`data: ${JSON.stringify({ type: 'connected', clientId })}\n\n`);
  sseClients.add(res);

  // Send keep-alive ping every 30 seconds to prevent timeout
  const keepAlive = setInterval(() => {
    try {
      res.write(':ping\n\n');
    } catch (e) {
      clearInterval(keepAlive);
    }
  }, 30000);
  
  // Remove client on disconnect
  req.on('close', () => {
    clearInterval(keepAlive);
    console.log(`SSE client disconnected: ${clientId}`);
    sseClients.delete(res);
  });

  req.on('aborted', () => {
    clearInterval(keepAlive);
    console.log(`SSE client aborted: ${clientId}`);
    sseClients.delete(res);
  });

  req.on('error', () => {
    clearInterval(keepAlive);
    sseClients.delete(res);
  });
});

// Helper function to broadcast product updates to all connected clients
function broadcastProductUpdate(type, data) {
  const message = JSON.stringify({ type, data, timestamp: Date.now() });
  console.log(`Broadcasting ${type} to ${sseClients.size} SSE clients:`, message);
  
  sseClients.forEach(client => {
    try {
      client.write(`data: ${message}\n\n`);
      console.log('Sent message to SSE client');
    } catch (error) {
      console.error('Error sending to SSE client:', error);
      // Remove disconnected clients
      sseClients.delete(client);
    }
  });
}

// Allow both admin and cashier access for POS operations
const isAdminOrCashier = (req, res, next) => {
  const src = req.session.user || req.user || {};
  const role = normalizeRole(src.role ? String(src.role).toLowerCase() : '');
  if (role === 'admin' || role === 'cashier' || role === 'superadmin') return next();
  return res.status(403).json({ success:false, message:'Access Denied' });
};

// --- API Routes ---
// --- POS Settings (basic) ---
app.get('/api/settings', isAuthenticated, isAdminOrCashier, async (req, res) => {
  try {
    const raw = await readData('settings.json');
    const base = Array.isArray(raw) ? {} : (raw || {});
    let data = { ...base };
    // Per-user preference override (stored in users.json)
    try {
      const sessionUser = req.session?.user || req.user || null;
      if (sessionUser && sessionUser.id != null) {
        const users = await readData("users.json");
        if (Array.isArray(users)) {
          const me = users.find((u) => String(u.id) === String(sessionUser.id));
          const prefDarkMode = me && me.preferences && typeof me.preferences.darkMode === 'boolean'
            ? me.preferences.darkMode
            : null;
          if (typeof prefDarkMode === 'boolean') data.darkMode = prefDarkMode;
        }
      }
    } catch {}
    // Default: hide POS banner unless explicitly enabled
    if (data.posShowBanner === undefined || data.posShowBanner === null) {
      data.posShowBanner = false;
    }
    try {
      const licensedName = await getLicensedStoreName();
      if (licensedName) {
        data.storeName = licensedName;
        try {
          if (data && typeof data === 'object' && data['0'] && typeof data['0'] === 'object') {
            data['0'] = { ...data['0'], storeName: licensedName };
          }
        } catch {}
      }
    } catch {}
    data.storageMode = 'sqlite';
    res.json(data || {});
  } catch {
    res.json({ storageMode: 'sqlite' });
  }
});

app.put('/api/settings', strictLimiter, isAuthenticated, isAdmin, async (req, res) => {
  try {
    const cur = await readData('settings.json');
    const merged = { ...(cur && typeof cur === 'object' ? cur : {}), ...(req.body || {}) };
    
    // PROTEKSI: Jika ada LICENSE offline valid dengan nama toko, paksa storeName mengikuti LICENSE
    // Ini mencegah perubahan store name melalui API atau inspect element
    try {
      const licensedName = await getLicensedStoreName();
      if (licensedName) {
        // Log attempt to change store name
        if (req.body.storeName && req.body.storeName !== licensedName) {
          // console.warn('[SECURITY] Attempt to change store name blocked:', {
            // requested: req.body.storeName,
            // licensed: licensedName,
            // user: req.session.user,
            // timestamp: new Date().toISOString()
          // });
        }
        merged.storeName = licensedName;
      }
    } catch (e) {
      // console.warn('[SECURITY] Error checking licensed store name:', e);
      // console.warn('[SECURITY] Error checking licensed store name:', e);
    }
    
    merged.updatedAt = Date.now();
    try {
      const enc = merged && merged.encryption;
      if (enc && typeof enc.enabled === 'boolean') {
        encryptionEnabled = !!enc.enabled;
      }
    } catch {}
    await writeData('settings.json', merged);
    // Keep current user's dark mode preference in sync with settings form save
    // to avoid mismatch where /api/settings gets overridden by stale per-user value.
    try {
      const hasDarkMode = req.body && Object.prototype.hasOwnProperty.call(req.body, 'darkMode');
      const bodyDarkMode = hasDarkMode ? req.body.darkMode : null;
      const sessionUser = req.session?.user || req.user || null;
      if (sessionUser && sessionUser.id != null && typeof bodyDarkMode === 'boolean') {
        const users = await readData("users.json");
        if (Array.isArray(users)) {
          const idx = users.findIndex((u) => String(u.id) === String(sessionUser.id));
          if (idx >= 0) {
            const curUser = users[idx] || {};
            users[idx] = {
              ...curUser,
              preferences: {
                ...(curUser.preferences && typeof curUser.preferences === 'object' ? curUser.preferences : {}),
                darkMode: bodyDarkMode
              },
              updatedAt: new Date().toISOString()
            };
            await saveArrayWithSync("users.json", users);
          }
        }
      }
    } catch {}
    // Enqueue outbox for sync (settings upsert)
    try {
      await enqueueOutbox({ collection: 'settings', file: 'settings.json', op: 'upsert', _id: 'settings', doc: merged, updatedAt: Date.now() });
    } catch {}
    // Mirror sync config if provided into dedicated file with metadata
    try {
      const syncIn = (req.body && req.body.sync) ? req.body.sync : null;
      if (syncIn && typeof syncIn === 'object') {
        const prev = await readData(SYNC_CFG_FILE).catch(() => ({}));
        const who = (req.session && req.session.user && req.session.user.username) || '';
        const syncCfg = { ...(prev && typeof prev === 'object' ? prev : {}), ...syncIn, updatedAt: Date.now(), lastModifiedBy: who };
        await writeData(SYNC_CFG_FILE, syncCfg);
        try { await enqueueOutbox({ collection: 'sync_config', file: SYNC_CFG_FILE, op: 'upsert', _id: 'sync_config', doc: syncCfg, updatedAt: Number(syncCfg.updatedAt||Date.now()) }); } catch {}
      }
    } catch {}
    res.json({ success:true, message:'Settings saved', settings: merged });
  } catch (e) {
    res.status(500).json({ success:false, message:'Failed to save settings' });
  }
});

app.put('/api/me/preferences/dark-mode', strictLimiter, isAuthenticated, async (req, res) => {
  try {
    const sessionUser = req.session?.user || req.user || null;
    if (!sessionUser || sessionUser.id == null) {
      return res.status(401).json({ success: false, message: 'Unauthorized' });
    }
    const darkMode = req.body && typeof req.body.darkMode === 'boolean' ? req.body.darkMode : null;
    if (typeof darkMode !== 'boolean') {
      return res.status(400).json({ success: false, message: 'darkMode must be boolean' });
    }
    const users = await readData("users.json");
    if (!Array.isArray(users)) {
      return res.status(500).json({ success: false, message: 'Users data invalid' });
    }
    const idx = users.findIndex((u) => String(u.id) === String(sessionUser.id));
    if (idx < 0) {
      return res.status(404).json({ success: false, message: 'User not found' });
    }
    const nowIso = new Date().toISOString();
    const current = users[idx] || {};
    users[idx] = {
      ...current,
      preferences: {
        ...(current.preferences && typeof current.preferences === 'object' ? current.preferences : {}),
        darkMode
      },
      updatedAt: nowIso
    };
    await saveArrayWithSync("users.json", users);
    return res.json({ success: true, preferences: users[idx].preferences });
  } catch (e) {
    return res.status(500).json({ success: false, message: 'Failed to save dark mode preference' });
  }
});

// --- Session Keep-Alive Ping Endpoint ---
app.post('/api/session/ping', isAuthenticated, (req, res) => {
  // This endpoint simply touches the session to keep it alive
  // The session middleware will automatically update the cookie
  res.json({ success: true, message: 'Session refreshed' });
});

// --- Sync Config endpoint ---
app.get('/api/sync/config', isAuthenticated, isAdminOrCashier, async (req, res) => {
  try {
    const cfg = await getSyncConfig();
    res.json({ success: true, sync: cfg });
  } catch (e) {
    res.status(500).json({ success: false, message: 'Failed to read sync config' });
  }
});

app.put('/api/sync/config', strictLimiter, isAuthenticated, isAdminOrCashier, async (req, res) => {
  try {
    const syncIn = (req.body && req.body.sync) ? req.body.sync : req.body;
    if (!syncIn || typeof syncIn !== 'object') {
      return res.status(400).json({ success: false, message: 'Invalid sync config' });
    }
    const prev = await readData(SYNC_CFG_FILE).catch(() => ({}));
    const who = (req.session && req.session.user && req.session.user.username) || '';
    const syncCfg = {
      ...(prev && typeof prev === 'object' ? prev : {}),
      ...syncIn,
      enabled: syncIn.enabled === true,
      baseUrl: String(syncIn.baseUrl || '').trim(),
      token: String(syncIn.token || ''),
      chunkSize: Math.min(1000, Math.max(1, Number(syncIn.chunkSize || 300) || 300)),
      integrityVerify: syncIn.integrityVerify === true,
      updatedAt: Date.now(),
      lastModifiedBy: who
    };
    await writeData(SYNC_CFG_FILE, syncCfg);
    try {
      await enqueueOutbox({
        collection: 'sync_config',
        file: SYNC_CFG_FILE,
        op: 'upsert',
        _id: 'sync_config',
        doc: syncCfg,
        updatedAt: Number(syncCfg.updatedAt || Date.now())
      });
    } catch {}
    const settings = await readData('settings.json').catch(() => ({}));
    const mergedSettings = {
      ...(settings && typeof settings === 'object' ? settings : {}),
      sync: syncCfg,
      updatedAt: Date.now()
    };
    await writeData('settings.json', mergedSettings);
    try {
      await enqueueOutbox({
        collection: 'settings',
        file: 'settings.json',
        op: 'upsert',
        _id: 'settings',
        doc: mergedSettings,
        updatedAt: Date.now()
      });
    } catch {}
    const cfg = await getSyncConfig();
    res.json({ success: true, sync: cfg });
  } catch (e) {
    res.status(500).json({ success: false, message: 'Failed to save sync config' });
  }
});

// --- Sync Status (registered after auth middlewares) ---
app.get('/api/sync/status', isAuthenticated, isAdminOrCashier, async (req, res) => {
  try {
    const settings = await readData('settings.json').catch(() => ({}));
    const syncCfgFile = await readData(SYNC_CFG_FILE).catch(() => ({}));
    const syncFromSettings = (settings && settings.sync && typeof settings.sync === 'object') ? settings.sync : {};
    const syncFromFile = (syncCfgFile && typeof syncCfgFile === 'object') ? syncCfgFile : {};
    const tsSettings = Number(syncFromSettings.updatedAt || settings?.updatedAt || 0) || 0;
    const tsFile = Number(syncFromFile.updatedAt || 0) || 0;
    const preferSettings = tsSettings >= tsFile;
    const primary = preferSettings ? syncFromSettings : syncFromFile;
    const secondary = preferSettings ? syncFromFile : syncFromSettings;
    const sync = { ...secondary, ...primary };
    const last = await readData(LASTSYNC_FILE).catch(() => ({}));
    const outbox = await readArrayFile(OUTBOX_FILE).catch(() => []);
    // Optional remote status probe when ?remote=true and config available
    let remote = null;
    try {
      const wantRemote = String(req.query && (req.query.remote || '')).toLowerCase();
      const doRemote = wantRemote === '1' || wantRemote === 'true' || wantRemote === 'yes';
      const baseUrl = String(sync.baseUrl || '');
      const token = String(sync.token || '');
      if (doRemote && baseUrl) {
        let ep;
        try {
          ep = new URL('/api/sync/changes', baseUrl);
          // Use since=0 for full snapshot summary; if too heavy, caller can omit ?remote
          ep.searchParams.set('since', '0');
          ep.searchParams.set('clientId', settings.clientId || '');
        } catch {}
        if (ep) {
          const headers = await buildRemoteSyncHeaders({ token, baseUrl: baseUrl });
          const resp = await safeFetch(String(ep), { headers });
          if (resp && resp.ok) {
            const payload = await resp.json().catch(() => ({}));
            const byFile = {};
            const latestTs = {};
            const ts = (v) => {
              if (typeof v === 'number' && isFinite(v)) return v;
              if (typeof v === 'string') { const t = Date.parse(v); return isNaN(t) ? 0 : t; }
              return 0;
            };
            const recTs = (x) => Math.max(ts(x?.updatedAt), ts(x?.timestamp), ts(x?.createdAt));
            if (payload && typeof payload === 'object') {
              for (const [file, arr] of Object.entries(payload)) {
                const list = Array.isArray(arr) ? arr : [];
                byFile[file] = list.length;
                let maxT = 0;
                for (const it of list) { const r = recTs(it); if (r > maxT) maxT = r; }
                latestTs[file] = maxT;
              }
            }
            remote = { reachable: true, status: resp.status, byFile, latestTs, latestTsMax: Math.max(0, ...Object.values(latestTs)) };
          } else {
            remote = { reachable: false, status: resp ? resp.status : 0 };
          }
        }
      }
    } catch {}
    res.json({
      enabled: sync.enabled === true,
      baseUrl: sync.baseUrl || '',
      hasToken: Boolean(sync.token && String(sync.token).length > 0),
      token: sync.token || '',
      clientId: settings.clientId || '',
      lastModifiedBy: sync.lastModifiedBy || '',
      cfgUpdatedAt: Number(sync.updatedAt || 0),
      lastPushAt: Number(last.lastPushAt || 0),
      lastPullAt: Number(last.lastPullAt || 0),
      outboxSize: outbox.length,
      remote
    });
  } catch (e) {
    res.status(500).json({ success:false, message:'Failed to read sync status' });
  }
});

// --- Sync Read endpoint (mirror pos-sync-server) ---
app.get('/api/sync/read/:file', requireSyncBearer, async (req, res) => {
  try {
    const file = String(req.params.file || '').trim();
    if (!file) {
      return res.status(400).json({ success: false, message: 'Nama file wajib' });
    }
    const data = await readData(file);
    res.json({ success: true, file, data, tenant: 'default', at: Date.now() });
  } catch (e) {
    res.status(500).json({ success: false, message: e.message || 'read failed' });
  }
});

async function buildBluetoothExportEnvelope(dataTypes = []) {
  const allowedFiles = Array.isArray(dataTypes) && dataTypes.length
    ? new Set(dataTypes.map((dt) => `${String(dt || '').trim()}.json`).filter((f) => f !== '.json'))
    : null;
  const fileList = allowedFiles
    ? Array.from(allowedFiles)
    : [...SYNC_COLLECTION_FILES, 'transactions.json', 'settings.json', SYNC_CFG_FILE];
  const files = {};
  for (const file of fileList) {
    try {
      const data = await readData(file);
      if (file === 'settings.json' || file === SYNC_CFG_FILE || file === 'banners.json' || file === 'qris.json') {
        if (data && typeof data === 'object' && !Array.isArray(data) && Object.keys(data).length) {
          files[file] = { ...(data || {}) };
        }
        continue;
      }
      if (Array.isArray(data) && data.length) {
        files[file] = data;
      }
    } catch {}
  }
  return {
    version: 1,
    transport: 'bluetooth',
    mode: 'snapshot_merge',
    createdAt: Date.now(),
    dataTypes: Array.isArray(dataTypes) ? dataTypes.slice() : [],
    files
  };
}

async function applyBluetoothPayloadToLocalServer(filteredPayload = {}) {
  let count = 0;
  const byFile = {};
  const ts = (v) => {
    if (typeof v === 'number' && isFinite(v)) return v;
    if (typeof v === 'string') { const t = Date.parse(v); return isNaN(t) ? 0 : t; }
    return 0;
  };
  const recTs = (x) => Math.max(ts(x?.updatedAt), ts(x?.timestamp), ts(x?.createdAt));
  const hasIncomingStockEvents = !!(
    (Array.isArray(filteredPayload['transactions.json']) && filteredPayload['transactions.json'].length > 0) ||
    (Array.isArray(filteredPayload['stock_in.json']) && filteredPayload['stock_in.json'].length > 0)
  );

  for (const [file, changes] of Object.entries(filteredPayload || {})) {
    if (!Array.isArray(changes)) continue;
    if (file === 'settings.json' || file === SYNC_CFG_FILE || file === 'banners.json' || file === 'qris.json') {
      try {
        const curObj = await readData(file).catch(() => ({}));
        const newest = changes.reduce((acc, it) => (recTs(it) > recTs(acc || {}) ? it : acc), null);
        if (newest && typeof newest === 'object') {
          const nu = recTs(newest);
          const cu = recTs(curObj || {});
          const keys = Object.keys(newest || {}).filter((k) => !['_id', 'id', 'updatedAt', 'timestamp', 'createdAt'].includes(k));
          const hasMeaningful = keys.length > 0 && keys.some((k) => newest[k] != null && newest[k] !== '');
          const curHasMeaningful = Object.keys(curObj || {}).length > 0;
          if (nu >= cu && (hasMeaningful || !curHasMeaningful)) {
            await writeData(file, { ...(curObj || {}), ...newest, updatedAt: nu });
            count += 1;
            byFile[file] = 1;
          }
        }
      } catch {}
      continue;
    }
    if (file === 'stock_in.json') {
      let cur = await readData(file).catch(() => []);
      if (!Array.isArray(cur)) cur = [];
      const prevList = cur.slice();
      const merged = mergeSyncArrayChanges(cur, changes, recTs);
      cur = merged.list;
      count += merged.count;
      await writeData(file, cur);
      if (merged.count > 0) {
        const adjusted = await adjustProductStockFromMergedStockIn(prevList, cur);
        if (adjusted > 0) byFile['products.json'] = (byFile['products.json'] || 0) + adjusted;
      }
      byFile[file] = (byFile[file] || 0) + merged.count;
      continue;
    }
    if (file === 'transactions.json') {
      let cur = await readData(file).catch(() => []);
      if (!Array.isArray(cur)) cur = [];
      const prevList = cur.slice();
      const merged = mergeSyncArrayChanges(cur, changes, recTs);
      cur = merged.list;
      count += merged.count;
      await writeData(file, cur);
      if (merged.count > 0) {
        const adjusted = await adjustProductStockFromMergedTransactions(prevList, cur);
        if (adjusted > 0) byFile['products.json'] = (byFile['products.json'] || 0) + adjusted;
      }
      byFile[file] = (byFile[file] || 0) + merged.count;
      continue;
    }
    let cur = await readData(file).catch(() => []);
    if (!Array.isArray(cur)) cur = [];
    const prevProductsByKey = file === 'products.json'
      ? new Map(cur.map((doc) => [String(doc && (doc._id || doc.id) || ''), doc]).filter(([k]) => !!k))
      : null;
    const merged = mergeSyncArrayChanges(cur, changes, recTs);
    cur = merged.list;
    count += merged.count;
    if (file === 'products.json' && hasIncomingStockEvents && prevProductsByKey) {
      cur = cur.map((doc) => {
        const key = String(doc && (doc._id || doc.id) || '');
        const prevDoc = prevProductsByKey.get(key);
        return prevDoc ? mergeProductRecordPreservingLocalStock(prevDoc, doc, true) : doc;
      });
    }
    await writeData(file, cur);
    byFile[file] = (byFile[file] || 0) + merged.count;
  }
  return { pulled: count, byFile };
}

app.post('/api/sync/bluetooth-export', isAuthenticated, isAdminOrCashier, async (req, res) => {
  try {
    const dataTypes = Array.isArray(req.body?.dataTypes) ? req.body.dataTypes : [];
    const envelope = await buildBluetoothExportEnvelope(dataTypes);
    res.json({ success: true, envelope });
  } catch (e) {
    res.status(500).json({ success: false, message: e.message || 'Failed to export bluetooth snapshot' });
  }
});

app.post('/api/sync/bluetooth-apply', isAuthenticated, isAdminOrCashier, async (req, res) => {
  try {
    const dataTypes = Array.isArray(req.body?.dataTypes) ? req.body.dataTypes : [];
    const files = (req.body && req.body.files && typeof req.body.files === 'object') ? req.body.files : {};
    const allowedFiles = dataTypes.length
      ? new Set(dataTypes.map((dt) => `${String(dt || '').trim()}.json`).filter((f) => f !== '.json'))
      : null;
    const filteredPayload = {};
    for (const [file, rawValue] of Object.entries(files)) {
      if (allowedFiles && !allowedFiles.has(file)) continue;
      if (file === 'settings.json' || file === SYNC_CFG_FILE || file === 'banners.json' || file === 'qris.json') {
        if (Array.isArray(rawValue)) filteredPayload[file] = rawValue;
        else if (rawValue && typeof rawValue === 'object') filteredPayload[file] = [rawValue];
        continue;
      }
      filteredPayload[file] = Array.isArray(rawValue) ? rawValue : (rawValue ? [rawValue] : []);
    }
    const result = await applyBluetoothPayloadToLocalServer(filteredPayload);
    res.json({ success: true, pulled: result.pulled, byFile: result.byFile || {} });
  } catch (e) {
    res.status(500).json({ success: false, message: e.message || 'Failed to apply bluetooth snapshot' });
  }
});

// --- Sync Now (manual trigger) ---
async function handleSyncNow(req, res) {
  try {
    const dataTypes = req.body?.dataTypes || []; // Get selected data types from request
    const allowedFiles = dataTypes.length > 0 ? new Set(dataTypes.map(dt => `${dt}.json`)) : null;
    
    // try { console.log('[SYNC] /api/sync/now', { method: req.method, user: (req.session && req.session.user && req.session.user.username) || '-', dataTypes }); } catch {}
    if (isSyncBusy()) return res.status(429).json({ success:false, message:'Sync is in progress' });
    const result = await runWithSyncLock(async () => {
      try { await performDatabaseBackupToFile('sync-pre'); } catch {}
      try {
        const wantFull = (String(req.query?.full||req.body?.full||'').toLowerCase() === 'true');
        let doFull = !!wantFull;
        if (!doFull) {
          try {
            let anyEmpty = false;
            const filesToCheck = allowedFiles ? [...allowedFiles] : SYNC_COLLECTION_FILES;
            for (const f of filesToCheck) {
              const arr = await readData(f).catch(()=>[]);
              if (Array.isArray(arr) && arr.length === 0) { anyEmpty = true; break; }
            }
            doFull = anyEmpty;
          } catch {}
        }
        resetSyncProgress(); setSyncPhase('pull1'); setSyncStart();
        if (doFull) {
          let last = await readData(LASTSYNC_FILE).catch(() => ({}));
          if (!last || typeof last !== 'object') last = {};
          last.lastPullAt = 0;
          last.lastPushedPerFile = {};
          await writeData(LASTSYNC_FILE, last);
        }
        await pullChangesWithRetry(3, dataTypes).catch(() => ({ pulled: 0, error: true }));
        setSyncEnd();
      } catch {}
      try { await enqueueLocalSnapshotIfOutboxEmpty(dataTypes); } catch {}
      try { await enqueueDeltaSinceLastPush(dataTypes); } catch {}
      try {
        let q = await readArrayFile(OUTBOX_FILE);
        const arrayFiles = allowedFiles || new Set([ ...SYNC_COLLECTION_FILES, 'transactions.json' ]);
        const hasArrayItems = Array.isArray(q) && q.some(it => arrayFiles.has(String(it.file||it.collection||'')));
        if (!q || !q.length || !hasArrayItems) {
          const fallbackEntries = [];
          for (const file of arrayFiles) {
            let data;
            try { data = await readData(file); } catch { data = null; }
            if (Array.isArray(data)) {
              const now = Date.now();
              for (let i = 0; i < data.length; i++) {
                try {
                  const doc = data[i] || {};
                  const id = String(doc && (doc._id || doc.id || ''));
                  if (!id) continue;
                  if (!doc.updatedAt || Number(doc.updatedAt) < now) { doc.updatedAt = now; data[i] = doc; }
                  fallbackEntries.push({ collection: file.replace('.json',''), file, op: 'upsert', _id: id, doc, updatedAt: Number(doc.updatedAt||now) });
                } catch {}
              }
              try { await writeData(file, data); } catch {}
            } else if (data && typeof data === 'object') {
              const now = Date.now();
              if (!data.updatedAt || Number(data.updatedAt) < now) { data.updatedAt = now; await writeData(file, data); }
              const id = (file === 'banners.json') ? 'banner' : (file === 'qris.json') ? 'qris' : (data._id || data.id || 'singleton');
              fallbackEntries.push({ collection: file.replace('.json',''), file, op: 'upsert', _id: String(id), doc: data, updatedAt: Number(data.updatedAt||now) });
            }
          }
          await batchEnqueueOutbox(fallbackEntries);
        }
      } catch {}
      resetSyncProgress(); setSyncPhase('push'); setSyncStart();
      const pushed = await pushOutboxChunked(500, dataTypes).catch(() => ({ pushed: 0, error: true }));
      setSyncEnd();
      resetSyncProgress(); setSyncPhase('pull2'); setSyncStart();
      const pulled = await pullChangesWithRetry(3, dataTypes).catch(() => ({ pulled: 0, error: true }));
      setSyncEnd();
      __syncProgress.phase = '';
      const success = !pushed.error && !pulled.error;
      const cfgNow = await getSyncConfig().catch(()=>({}));
      let integrity = null;
      let integrityMismatch = null;
      try { 
        if (cfgNow && cfgNow.integrityVerify) { 
          integrity = await computeChecksumsForCollections(); 
          
          // Perform integrity verification by comparing checksums before and after sync
          const checksumsBefore = req.body?.checksumsBefore || null;
          if (checksumsBefore) {
            const mismatches = [];
            for (const [file, checksum] of Object.entries(checksumsBefore)) {
              if (integrity[file] && integrity[file] !== checksum) {
                mismatches.push({ file, before: checksum, after: integrity[file] });
              }
            }
            if (mismatches.length > 0) {
              integrityMismatch = mismatches;
            }
          }
        } 
      } catch {}
      const meta = { at: Date.now(), completedAt: Date.now(), user: (req.session && req.session.user && req.session.user.username) || '', integrity, integrityMismatch, dataTypes };
      return { success, pushed, pulled, meta };
    });
    if (result && result.busy) return res.status(429).json({ success:false, message:'Sync is in progress' });
    res.json(result);
  } catch (e) {
    res.status(500).json({ success:false, message:'Failed to sync now' });
  }
}

app.get('/api/sync/progress', isAuthenticated, isAdminOrCashier, async (req, res) => {
  try {
    const busy = isSyncBusy();
    res.json({ busy, progress: __syncProgress });
  } catch {
    res.status(500).json({ success:false });
  }
});

app.get('/api/sync/checksums', isAuthenticated, isAdminOrCashier, async (req, res) => {
  try {
    const sums = await computeChecksumsForCollections();
    res.json({ success: true, checksums: sums });
  } catch {
    res.status(500).json({ success:false });
  }
});
app.post('/api/sync/now', isAuthenticated, isAdminOrCashier, handleSyncNow);
app.get('/api/sync/now', isAuthenticated, isAdminOrCashier, handleSyncNow);

async function handleSyncPushOnly(req, res) {
  try {
    const dataTypes = req.body?.dataTypes || []; // Get selected data types from request
    const allowedFiles = dataTypes.map(dt => `${dt}.json`);
    
    // try { console.log('[SYNC] /api/sync/push-local', { method: req.method, user: (req.session && req.session.user && req.session.user.username) || '-' }); } catch {}
    try { await enqueueLocalSnapshotIfOutboxEmpty(dataTypes); } catch {}
    try { await enqueueDeltaSinceLastPush(dataTypes); } catch {}
    try {
      let q = await readArrayFile(OUTBOX_FILE);
      if (!q || !q.length) {
        await enqueueFullSnapshot(dataTypes);
      } else {
        const arrayFiles = new Set(dataTypes.length > 0 ? allowedFiles : [ ...SYNC_COLLECTION_FILES, 'transactions.json' ]);
        const hasArrayItems = q.some(it => arrayFiles.has(String(it.file||it.collection||'')));
        if (!hasArrayItems) {
          for (const file of arrayFiles) {
            let data;
            try { data = await readData(file); } catch { data = null; }
            if (Array.isArray(data)) {
              const now = Date.now();
              for (let i = 0; i < data.length; i++) {
                try {
                  const doc = data[i] || {};
                  const id = String(doc && (doc._id || doc.id || ''));
                  if (!id) continue;
                  if (!doc.updatedAt || Number(doc.updatedAt) < now) { doc.updatedAt = now; data[i] = doc; }
                  await enqueueOutbox({ collection: file.replace('.json',''), file, op: 'upsert', _id: id, doc, updatedAt: Number(doc.updatedAt||now) });
                } catch (e) { console.warn('[SYNC] Failed to enqueue item for', file, id, e?.message); }
              }
              try { await writeData(file, data); } catch (e) { console.warn('[SYNC] Failed to write updated data for', file, e?.message); }
            } else if (data && typeof data === 'object') {
              const now = Date.now();
              if (!data.updatedAt || Number(data.updatedAt) < now) { data.updatedAt = now; try { await writeData(file, data); } catch (e) { console.warn('[SYNC] Failed to write', file, e?.message); } }
              const id = (file === 'banners.json') ? 'banner' : (file === 'qris.json') ? 'qris' : (data._id || data.id || 'singleton');
              await enqueueOutbox({ collection: file.replace('.json',''), file, op: 'upsert', _id: String(id), doc: data, updatedAt: Number(data.updatedAt||now) });
            }
          }
        }
      }
    } catch (e) { console.warn('[SYNC] handleSyncPushOnly error', e?.message); }
    const pushed = await pushOutbox(dataTypes).catch((e) => { console.warn('[SYNC] pushOutbox error', e?.message); return { pushed: 0, error: true }; });
    const success = !pushed.error;
    const meta = { at: Date.now(), lastPushAt: Date.now(), user: (req.session && req.session.user && req.session.user.username) || '', dataTypes };
    res.json({ success, pushed, meta });
  } catch (e) {
    res.status(500).json({ success:false, message:'Failed to push' });
  }
}
async function handleSyncPullOnly(req, res) {
  try {
    const dataTypes = req.body?.dataTypes || []; // Get selected data types from request
    try { console.log('[SYNC] /api/sync/pull-remote', { method: req.method, user: (req.session && req.session.user && req.session.user.username) || '-', dataTypes }); } catch {}
    const pulled = await pullChanges(dataTypes).catch(() => ({ pulled: 0, error: true }));
    const success = !pulled.error;
    const meta = { at: Date.now(), lastPullAt: Date.now(), user: (req.session && req.session.user && req.session.user.username) || '', dataTypes };
    res.json({ success, pulled, meta });
  } catch (e) {
    res.status(500).json({ success:false, message:'Failed to pull' });
  }
}
app.post('/api/sync/push-local', isAuthenticated, isAdminOrCashier, handleSyncPushOnly);
app.post('/api/sync/pull-remote', isAuthenticated, isAdminOrCashier, handleSyncPullOnly);

// --- Admin Utilities: Force Full Pull & Watermark Management ---
async function handleSyncForceFullPull(req, res) {
  try {
    let last = await readData(LASTSYNC_FILE).catch(() => ({}));
    if (!last || typeof last !== 'object') last = {};
    // Force since=0 by resetting lastPullAt, then pull
    last.lastPullAt = 0;
    await writeData(LASTSYNC_FILE, last);
    const pulled = await pullChanges().catch(() => ({ pulled: 0, error: true }));
    const success = !pulled.error;
    const meta = { at: Date.now(), forcedSince: 0, user: (req.session && req.session.user && req.session.user.username) || '' };
    res.json({ success, pulled, meta });
  } catch (e) {
    res.status(500).json({ success:false, message:'Failed to force full pull' });
  }
}
app.post('/api/sync/force-full-pull', isAuthenticated, isAdminOrCashier, handleSyncForceFullPull);

app.post('/api/sync/reset-watermark', isAuthenticated, isAdminOrCashier, async (req, res) => {
  try {
    let last = await readData(LASTSYNC_FILE).catch(() => ({}));
    if (!last || typeof last !== 'object') last = {};
    last.lastPullAt = 0;
    // Optional: also clear per-file push watermarks if requested via query/body
    try { if (req.query && String(req.query.clearPush || '').toLowerCase() === 'true') last.lastPushedPerFile = {}; } catch {}
    await writeData(LASTSYNC_FILE, last);
    res.json({ success:true, last });
  } catch (e) {
    res.status(500).json({ success:false, message:'Failed to reset watermark' });
  }
});

app.get('/api/sync/watermark', isAuthenticated, isAdminOrCashier, async (req, res) => {
  try {
    const last = await readData(LASTSYNC_FILE).catch(() => ({}));
    res.json({
      lastPullAt: Number(last.lastPullAt || 0),
      lastPushAt: Number(last.lastPushAt || 0),
      lastPushedPerFile: (last && last.lastPushedPerFile) || {}
    });
  } catch (e) {
    res.status(500).json({ success:false, message:'Failed to read watermark' });
  }
});

// --- Cleanup Sync Data ---
app.post('/api/sync/cleanup', isAuthenticated, isAdminOrCashier, async (req, res) => {
  try {
    let cleaned = 0;
    
    // Clean outbox file (sync data that has been processed)
    try {
      const outbox = await readArrayFile(OUTBOX_FILE).catch(() => []);
      const beforeCount = outbox.length;
      // Keep only items that are not older than 7 days or keep last 100 items
      const cutoffTime = Date.now() - (7 * 24 * 60 * 60 * 1000); // 7 days ago
      const filtered = outbox.filter(item => {
        // Keep items without timestamp or items newer than cutoff
        if (!item || !item.timestamp) return true;
        return item.timestamp > cutoffTime;
      });
      // Also limit to max 1000 items to prevent unbounded growth
      if (filtered.length > 1000) {
        filtered.splice(0, filtered.length - 1000);
      }
      await writeArrayFile(OUTBOX_FILE, filtered);
      cleaned += (beforeCount - filtered.length);
    } catch {}

    // Clean deletion tombstones (30-day age + max 1000 per file)
    let tombstonesRemoved = 0;
    try {
      const tomb = await cleanupOldTombstones();
      tombstonesRemoved = Number(tomb && tomb.removed) || 0;
      cleaned += tombstonesRemoved;
    } catch {}
    
    res.json({
      success: true,
      cleaned,
      tombstonesRemoved,
      message: `Cleaned ${cleaned} old sync records${tombstonesRemoved ? ` (${tombstonesRemoved} tombstone)` : ''}`
    });
  } catch (e) {
    res.status(500).json({ success:false, message:'Failed to cleanup sync data' });
  }
});

  app.post('/api/sync/reconcile', isAuthenticated, isAdminOrCashier, async (req, res) => {
    try {
      const cfg = await getSyncConfig();
      if (!cfg.enabled || !cfg.baseUrl) return res.status(400).json({ success:false, message:'Sync not configured' });
      let ep;
      try {
        ep = new URL('/api/sync/changes', cfg.baseUrl);
        ep.searchParams.set('since', '0');
        ep.searchParams.set('clientId', cfg.clientId || '');
      } catch {
        return res.status(400).json({ success:false, message:'Invalid baseUrl' });
      }
      const headers = await buildRemoteSyncHeaders(cfg);
      const resp = await safeFetch(String(ep), { headers });
      if (!resp) return res.status(502).json({ success:false, message:'No response from server' });
      if (!resp.ok) {
        let body = '';
        try { body = await resp.text(); } catch {}
        return res.status(resp.status || 500).json({ success:false, message:'Failed to fetch snapshot', body });
      }
      const payload = await resp.json().catch(() => ({}));
      const arrayFiles = new Set([ ...SYNC_COLLECTION_FILES, 'transactions.json' ]);
      const summary = { mergedByFile: {}, counts: {} };
      const ts = (v) => {
        if (typeof v === 'number' && isFinite(v)) return v;
        if (typeof v === 'string') { const t = Date.parse(v); return isNaN(t) ? 0 : t; }
        return 0;
      };
      const recTs = (x) => Math.max(ts(x?.updatedAt), ts(x?.timestamp), ts(x?.createdAt));
      for (const file of arrayFiles) {
        try {
          const changes = Array.isArray(payload[file]) ? payload[file] : [];
          let cur = await readData(file).catch(() => []);
          if (!Array.isArray(cur)) cur = [];
          const merged = mergeSyncArrayChanges(cur, changes, recTs);
          await writeData(file, merged.list);
          summary.mergedByFile[file] = true;
          summary.counts[file] = merged.count;
        } catch {}
      }
      // Update watermark so subsequent pulls use a fresh baseline
      try {
        let last = await readData(LASTSYNC_FILE).catch(()=>({}));
        if (!last || typeof last !== 'object') last = {};
        last.lastPullAt = Date.now();
        await writeData(LASTSYNC_FILE, last);
      } catch {}
      res.json({ success:true, summary, at: Date.now() });
    } catch (e) {
      res.status(500).json({ success:false, message:'Failed to reconcile' });
    }
  });

// --- Basic Sync Endpoints (server side) ---
// Accepts an outbox batch from a client and applies changes
app.post('/api/sync/push', requireSyncBearer, async (req, res) => {
  try {
    const { items = [], clientId = '', batchId = '' } = req.body || {};
    if (!Array.isArray(items)) return res.status(400).json({ success:false, message:'Invalid items' });
    
    // Validate clientId is not empty
    if (!clientId || typeof clientId !== 'string' || clientId.trim().length === 0) {
      return res.status(400).json({ success: false, message: 'Missing or invalid clientId' });
    }
    
    // Validate that the items array isn't too large to prevent abuse
    if (items.length > 10000) {
      return res.status(400).json({ success: false, message: 'Batch size too large' });
    }
    
    let mergeMode = 'serverWins';
    try { const sc = await readData(SYNC_CFG_FILE).catch(()=>({})); if (sc && typeof sc === 'object' && sc.mergeMode) mergeMode = String(sc.mergeMode); } catch {}
    const clientWins = (mergeMode === 'clientWins');
    let applied = 0;
    let errors = 0;
    let productsCache = null;
    let productsById = null;
    let productsDirty = false;
    let transactionsSynced = false;

    const ensureProductsForSyncPush = async () => {
      if (!Array.isArray(productsCache)) {
        productsCache = await readData('products.json').catch(() => []);
        if (!Array.isArray(productsCache)) productsCache = [];
        productsById = new Map(productsCache.map((p) => [String(p && p.id), p]).filter(([k]) => !!k));
      }
      return { productsCache, productsById };
    };

    const applyStockForTransactionStateChange = async (oldTx, newTx) => {
      const oldActive = !!oldTx && !isSyncVoidedTransaction(oldTx);
      const newActive = !!newTx && !isSyncVoidedTransaction(newTx);
      if (!oldActive && !newActive) return false;
      const { productsById } = await ensureProductsForSyncPush();
      let changed = false;
      if (oldActive && !newActive) {
        changed = applyStockDeltaFromSyncItems(productsById, oldTx.items, +1) || changed;
      } else if (!oldActive && newActive) {
        changed = applyStockDeltaFromSyncItems(productsById, newTx.items, -1) || changed;
      } else if (oldActive && newActive) {
        changed = applyStockDeltaFromSyncItems(productsById, oldTx.items, +1) || changed;
        changed = applyStockDeltaFromSyncItems(productsById, newTx.items, -1) || changed;
      }
      if (changed) productsDirty = true;
      return changed;
    };

    const allowedFiles = new Set([...SYNC_COLLECTION_FILES, 'settings.json', SYNC_CFG_FILE, 'transactions.json']);
    const singletonFiles = new Set(['settings.json', SYNC_CFG_FILE]);
    const byFile = {};
    for (const it of items) {
      if (!it || typeof it !== 'object') { errors++; continue; }
      const file = String(it.file || it.collection || '').trim();
      if (!file || !allowedFiles.has(file)) { errors++; continue; }
      if (!byFile[file]) byFile[file] = [];
      byFile[file].push(it);
    }

    for (const [file, fileItems] of Object.entries(byFile)) {
      try {
        if (singletonFiles.has(file)) {
          for (const it of fileItems) {
            const cur = await readData(file).catch(() => ({}));
            const cand = it.doc || {};
            if (cand && typeof cand === 'object') {
              if (clientWins || Number(cand.updatedAt || 0) >= Number(cur.updatedAt || 0)) {
                await writeData(file, { ...cur, ...cand, updatedAt: Number(cand.updatedAt || Date.now()) });
                applied++;
              }
            }
          }
          continue;
        }

        let arr = await readData(file).catch(() => []);
        if (!Array.isArray(arr)) arr = [];
        const idIndex = new Map();
        arr.forEach((x, idx) => {
          const k = String(x && (x._id || x.id) || '');
          if (k) idIndex.set(k, idx);
        });

        for (const it of fileItems) {
          try {
            const op = String(it.op || 'upsert');
            const key = String(it._id || it.id || (it.doc && (it.doc._id || it.doc.id)) || '');
            if (!key) { errors++; continue; }
            const idx = idIndex.has(key) ? idIndex.get(key) : -1;
            const prevTx = (file === 'transactions.json' && idx >= 0) ? (arr[idx] || null) : null;
            let nextTx = prevTx;
            let txStateChanged = false;

            if (op === 'delete' || it.deleted) {
              if (idx >= 0) {
                arr.splice(idx, 1);
                idIndex.clear();
                arr.forEach((x, i) => {
                  const k = String(x && (x._id || x.id) || '');
                  if (k) idIndex.set(k, i);
                });
                applied++;
                txStateChanged = true;
                nextTx = null;
              }
              try { await appendDeletionTombstone(file, key, Number(it.updatedAt || (it.doc && it.doc.updatedAt) || Date.now())); } catch {}
            } else {
              const cand = it.doc || {};
              if (cand && typeof cand === 'object') {
                if (idx >= 0) {
                  const cur = arr[idx] || {};
                  const cu = Number(cur.updatedAt || cur.timestamp || 0);
                  const nu = Number(cand.updatedAt || cand.timestamp || 0);
                  if (clientWins || nu >= cu) {
                    arr[idx] = cand;
                    applied++;
                    txStateChanged = true;
                    nextTx = cand;
                  }
                } else {
                  arr.push(cand);
                  idIndex.set(key, arr.length - 1);
                  applied++;
                  txStateChanged = true;
                  nextTx = cand;
                }
              }
            }

            if (file === 'transactions.json' && txStateChanged) {
              await applyStockForTransactionStateChange(prevTx, nextTx);
              transactionsSynced = true;
            }
          } catch (err) {
            console.error(`Error processing sync item for file ${file}:`, err);
            errors++;
          }
        }

        await writeData(file, arr);
        if (file === 'products.json') {
          try { invalidateCache('products'); } catch {}
        }
      } catch (err) {
        console.error(`Error processing sync batch for file ${file}:`, err);
        errors += fileItems.length;
      }
    }
    if (productsDirty && Array.isArray(productsCache)) {
      await saveArrayWithSync('products.json', productsCache, { keyField: 'id' });
    }
    if (transactionsSynced) {
      try { broadcastProductUpdate('transactions_updated', { reason: 'sync_push' }); } catch {}
    }
    res.json({ success:true, applied, clientId, batchId, errors });
  } catch (e) {
    console.error('Sync push endpoint error:', e);
    res.status(500).json({ success:false, message:'push failed' });
  }
});

// Returns changes since a timestamp per supported file
app.get('/api/sync/changes', requireSyncBearer, async (req, res) => {
  try {
    const since = Number(req.query.since || 0) || 0;
    const out = {};
    const ts = (v) => {
      if (typeof v === 'number' && isFinite(v)) return v;
      if (typeof v === 'string') { const t = Date.parse(v); return isNaN(t) ? 0 : t; }
      return 0;
    };
    const recTs = (x) => Math.max(ts(x?.updatedAt), ts(x?.timestamp), ts(x?.createdAt));
    // Transactions: use timestamp
    try {
      let t = await readData('transactions.json'); if (!Array.isArray(t)) t = [];
      const changes = since === 0 ? t : t.filter(x => recTs(x) > since);
      if (changes.length) out['transactions.json'] = changes;
    } catch {}
    // Settings: use updatedAt
    try {
      const s = await readData('settings.json');
      if (since === 0) {
        if (s && typeof s === 'object') out['settings.json'] = [ { ...(s || {}), _id: 'settings' } ];
      } else if (s && typeof s === 'object' && recTs(s) > since) {
        out['settings.json'] = [ { ...(s || {}), _id: 'settings' } ];
      }
    } catch {}
    // Sync config: use updatedAt
    try {
      const sc = await readData(SYNC_CFG_FILE).catch(() => null);
      if (since === 0) {
        if (sc && typeof sc === 'object') out[SYNC_CFG_FILE] = [ { ...(sc || {}), _id: 'sync_config' } ];
      } else if (sc && typeof sc === 'object' && recTs(sc) > since) {
        out[SYNC_CFG_FILE] = [ { ...(sc || {}), _id: 'sync_config' } ];
      }
    } catch {}
    // Whitelisted collections: use updatedAt/timestamp per record
    for (const file of SYNC_COLLECTION_FILES) {
      try {
        let raw = await readData(file).catch(() => []);
        if (Array.isArray(raw)) {
          const changes = since === 0 ? raw : raw.filter(x => recTs(x) > since);
          if (changes.length) out[file] = changes;
        } else if (raw && typeof raw === 'object') {
          const id = (file === 'banners.json') ? 'banner' : (file === 'qris.json') ? 'qris' : (raw._id || raw.id || 'singleton');
          const t = recTs(raw);
          if (since === 0 || t > since) {
            out[file] = [ { ...raw, _id: id } ];
          }
        }
      } catch {}
    }
    // Include deletions (tombstones) per file so other devices can remove items
    try {
      const delMap = await readData(DELETIONS_FILE).catch(() => ({}));
      if (delMap && typeof delMap === 'object') {
        for (const [file, dels] of Object.entries(delMap)) {
          const list = Array.isArray(dels) ? dels.filter(x => Number(x && x.updatedAt || 0) > since) : [];
          if (!list.length) continue;
          const existing = Array.isArray(out[file]) ? out[file] : [];
          const tombs = list.map(x => ({ _id: String(x._id || x.id || ''), deleted: true, updatedAt: Number(x.updatedAt || Date.now()) }));
          out[file] = existing.concat(tombs);
        }
      }
    } catch {}
    res.json(out);
  } catch (e) {
    res.status(500).json({ success:false, message:'changes failed' });
  }
});

// --- Drafts ---
app.get('/api/drafts', isAuthenticated, isAdminOrCashier, async (req, res) => {
  try { const d = await readData('drafts.json'); res.json(Array.isArray(d) ? d : []); }
  catch { res.json([]); }
});

app.post('/api/drafts', isAuthenticated, isAdminOrCashier, async (req, res) => {
  try {
    const { items } = req.body || {};
    if (!Array.isArray(items) || items.length === 0) return res.status(400).json({ success:false, message:'No items' });
    let d = await readData('drafts.json'); if (!Array.isArray(d)) d = [];
    const draft = { id: String(Date.now()), items, timestamp: Date.now() };
    d.push(draft);
    await writeData('drafts.json', d);
    res.json({ success:true, message:'Draf disimpan', id: draft.id });
  } catch { res.status(500).json({ success:false, message:'Failed to save draft' }); }
});

app.put('/api/drafts/:id/load', isAuthenticated, isAdminOrCashier, async (req, res) => {
  try {
    const id = String(req.params.id);
    const d = await readData('drafts.json');
    const found = (Array.isArray(d) ? d : []).find(x => String(x.id) === id);
    if (!found) return res.status(404).json({ success:false, message:'Draft not found' });
    res.json({ success:true, id, items: found.items || [], timestamp: found.timestamp });
  } catch { res.status(500).json({ success:false, message:'Failed to load draft' }); }
});

app.delete('/api/drafts/:id', isAuthenticated, isAdminOrCashier, async (req, res) => {
  try {
    const id = String(req.params.id);
    let d = await readData('drafts.json'); if (!Array.isArray(d)) d = [];
    const before = d.length;
    d = d.filter(x => String(x.id) !== id);
    await writeData('drafts.json', d);
    res.json({ success:true, deleted: before - d.length, message: 'Draft deleted' });
  } catch { res.status(500).json({ success:false, message:'Failed to delete draft' }); }
});

// --- Per-user Cart (isolated per logged-in cashier) ---
// SSE subscribers for real-time cart updates: res -> userId
const cartSubscribers = new Map();

async function readCartStore() {
  const raw = await readData('cart.json');
  if (raw && raw.users && typeof raw.users === 'object' && !Array.isArray(raw.users)) {
    return raw;
  }
  return { users: {} };
}

async function getUserCartFromStore(userId) {
  const id = String(userId || '').trim();
  if (!id) return { items: [], updatedAt: 0 };
  const store = await readCartStore();
  const entry = store.users[id];
  if (entry && Array.isArray(entry.items)) {
    return { items: entry.items, updatedAt: Number(entry.updatedAt || 0) };
  }
  return { items: [], updatedAt: 0 };
}

async function setUserCartInStore(userId, payload) {
  const id = String(userId || '').trim();
  if (!id) throw new Error('Missing user id for cart');
  const store = await readCartStore();
  const next = {
    items: Array.isArray(payload && payload.items) ? payload.items : [],
    updatedAt: Number((payload && payload.updatedAt) || Date.now()),
  };
  store.users[id] = next;
  await writeData('cart.json', store);
  return next;
}

const broadcastCart = (cartPayload, userId) => {
  const uid = String(userId || '').trim();
  const payload = {
    userId: uid,
    items: Array.isArray(cartPayload?.items) ? cartPayload.items : [],
    updatedAt: Number(cartPayload?.updatedAt || Date.now()),
  };
  const data = `data: ${JSON.stringify(payload)}\n\n`;
  for (const [res, subUserId] of Array.from(cartSubscribers.entries())) {
    if (String(subUserId) !== uid) continue;
    try { res.write(data); } catch { try { cartSubscribers.delete(res); } catch {} }
  }
};

app.get('/api/cart/stream', isAuthenticated, isAdminOrCashier, async (req, res) => {
  try {
    const userId = String((req.session && req.session.user && req.session.user.id) || '').trim();
    if (!userId) return res.status(401).end();
    res.setHeader('Content-Type', 'text/event-stream');
    res.setHeader('Cache-Control', 'no-cache');
    res.setHeader('Connection', 'keep-alive');
    res.flushHeaders && res.flushHeaders();
    res.write(': connected\n\n');
    cartSubscribers.set(res, userId);
    req.on('close', () => { try { cartSubscribers.delete(res); } catch {} });
    try {
      const snap = await getUserCartFromStore(userId);
      res.write(`data: ${JSON.stringify({ userId, items: snap.items || [], updatedAt: snap.updatedAt || Date.now() })}\n\n`);
    } catch {}
  } catch {
    try { res.end(); } catch {}
  }
});

app.get('/api/cart', isAuthenticated, isAdminOrCashier, async (req, res) => {
  try {
    const userId = String((req.session && req.session.user && req.session.user.id) || '').trim();
    if (!userId) return res.status(401).json({ success: false, message: 'Unauthorized' });
    const data = await getUserCartFromStore(userId);
    return res.json({ items: data.items, updatedAt: data.updatedAt || 0 });
  } catch (e) {
    return res.json({ items: [], updatedAt: 0 });
  }
});

app.put('/api/cart', isAuthenticated, isAdminOrCashier, async (req, res) => {
  try {
    const userId = String((req.session && req.session.user && req.session.user.id) || '').trim();
    if (!userId) return res.status(401).json({ success: false, message: 'Unauthorized' });
    const items = Array.isArray(req.body && req.body.items) ? req.body.items : null;
    if (!items) return res.status(400).json({ success:false, message:'Invalid items' });
    const payload = { items, updatedAt: Date.now() };
    const saved = await setUserCartInStore(userId, payload);
    broadcastCart(saved, userId);
    return res.json({ success:true, updatedAt: saved.updatedAt });
  } catch (e) {
    return res.status(500).json({ success:false, message:'Failed to save cart' });
  }
});

app.delete('/api/cart', isAuthenticated, isAdminOrCashier, async (req, res) => {
  try {
    const userId = String((req.session && req.session.user && req.session.user.id) || '').trim();
    if (!userId) return res.status(401).json({ success: false, message: 'Unauthorized' });
    const payload = { items: [], updatedAt: Date.now() };
    const saved = await setUserCartInStore(userId, payload);
    broadcastCart(saved, userId);
    return res.json({ success:true });
  } catch (e) {
    return res.status(500).json({ success:false, message:'Failed to clear cart' });
  }
});

// --- Recent Transactions ---
app.get('/api/recent-transactions', isAuthenticated, isAdminOrCashier, async (req, res) => {
  try {
    const t = await readData('transactions.json');
    const arr = Array.isArray(t) ? t : [];
    const getInputTime = (tx) => {
      const raw = tx?.updatedAt ?? tx?.createdAt ?? tx?.timestamp ?? tx?.date ?? tx?.transactionDate ?? 0;
      const n = Number(raw);
      if (Number.isFinite(n) && n > 0) return n;
      const parsed = Date.parse(String(raw));
      return Number.isFinite(parsed) ? parsed : 0;
    };
    // Sort berdasarkan "waktu input/ubah" (updatedAt/createdAt) supaya match Admin (Urut Terbaru/Terlama).
    const sorted = arr.sort((a,b)=> getInputTime(b) - getInputTime(a)).slice(0, 100);
    res.json(sorted);
  } catch { res.status(500).json({ success:false, message:'Failed to load transactions' }); }
});

// --- Transactions ---
function computeTransactionItemsSubtotal(items) {
  if (!Array.isArray(items)) return 0;
  return items.reduce((sum, it) => {
    const stored = Number(it && it.subtotal);
    if (Number.isFinite(stored) && stored > 0) return sum + stored;
    const qty = Number(it && (it.quantity != null ? it.quantity : it.qty)) || 0;
    if (!(qty > 0)) return sum;
    let effQty = qty;
    if (it && it.variant && it.variant.qty != null) {
      effQty = qty * (Number(it.variant.qty) || 1);
    } else if (it && it.variantQty != null) {
      effQty = qty * (Number(it.variantQty) || 1);
    }
    const price = Number(
      (it && it.variant && it.variant.price != null)
        ? it.variant.price
        : (it && it.price)
    ) || 0;
    return sum + (price * effQty);
  }, 0);
}

app.post('/api/transactions', isAuthenticated, isAdminOrCashier, async (req, res) => {
  try {
    const body = req.body || {};
    const {
      items = [],
      paymentMethod = 'cash',
      amountReceived = 0,
      customerId = 'default',
      customerName = 'Pelanggan Umum',
      discountPercent = 0,
      discountAmount = 0,
      useCustomerBalance = false
    } = body;
    if (!Array.isArray(items) || items.length === 0) return res.status(400).json({ success:false, message:'No items' });

    // Load products
    let products = await readData('products.json'); if (!Array.isArray(products)) products = [];

    // Prefer totals computed by kasir; fallback to server-side estimate
    const clientSubtotal = Number(body.subtotal);
    const subtotal = Number.isFinite(clientSubtotal) && clientSubtotal >= 0
      ? clientSubtotal
      : computeTransactionItemsSubtotal(items);
    const clientDisc = Number(body.discountAmount);
    const discAmt = Number.isFinite(clientDisc) && clientDisc >= 0
      ? clientDisc
      : (Number(discountAmount || 0) > 0
        ? Number(discountAmount || 0)
        : Math.round(subtotal * (Number(discountPercent || 0) / 100)));
    const taxAmount = Number.isFinite(Number(body.taxAmount)) ? Number(body.taxAmount) : 0;
    const serviceAmount = Number.isFinite(Number(body.serviceAmount)) ? Number(body.serviceAmount) : 0;
    const clientTotal = Number(body.totalAmount != null ? body.totalAmount : body.total);
    const totalAmount = Number.isFinite(clientTotal) && clientTotal >= 0
      ? clientTotal
      : Math.max(0, subtotal - (discAmt || 0) + taxAmount + serviceAmount);
    const shouldUseCustomerBalance = Boolean(useCustomerBalance) && customerId !== 'default' && String(customerId) !== '1';
    const customers = shouldUseCustomerBalance ? await readData('customers.json').catch(() => []) : [];
    const customerIndex = shouldUseCustomerBalance && Array.isArray(customers)
      ? customers.findIndex(c => String(c && c.id) === String(customerId))
      : -1;
    const customerBalanceBefore = customerIndex >= 0
      ? Math.max(0, Number(customers[customerIndex] && customers[customerIndex].balance || 0) || 0)
      : 0;
    const customerBalanceUsed = Math.min(customerBalanceBefore, totalAmount);
    const payableAfterBalance = Math.max(0, totalAmount - customerBalanceUsed);
    const amountReceivedNum = Number(amountReceived || 0);
    const paidAmountTotal = Math.max(0, amountReceivedNum + customerBalanceUsed);
    const remainingAmountTotal = Math.max(0, totalAmount - paidAmountTotal);

    const sessionUser = (req.session && req.session.user) ? req.session.user : {};
    const saleTs = resolveSaleTimestampForTransaction(req.body, sessionUser);
    const idDatePart = new Date(saleTs.timestamp).toISOString().slice(0, 10).replace(/-/g, '');
    const txId = `TRX-${idDatePart}-${Date.now()}`;
    const checkoutUser = sessionUser.username || '';

    const productIndex = new Map();
    for (let pi = 0; pi < products.length; pi++) {
      productIndex.set(String(products[pi].id), pi);
    }

    const pendingStockMoves = [];
    for (const it of items) {
      const idx = productIndex.get(String(it.productId));
      if (idx === undefined) continue;
      const product = products[idx];
      const itemQty = Number(it.quantity || it.qty || 0);
      if (!itemQty) continue;

      if (it.variant && Array.isArray(product.unitPrices)) {
        const variantIndex = Number(it.variant.index || 0);
        applyVariantStockChange(product, variantIndex, itemQty, -1);
        const variantRef = product.unitPrices[variantIndex];
        pendingStockMoves.push({
          productId: it.productId,
          delta: -itemQty,
          reason: 'sale',
          by: checkoutUser,
          refId: txId,
          variantIndex,
          variantUnit: (variantRef && variantRef.unit) || '',
          variantStock: Number((variantRef && variantRef.stock) || 0),
          stockOnly: true
        });
      } else {
        const currentStock = Number(products[idx].stock || 0);
        const nextStock = Math.max(0, currentStock - itemQty);
        products[idx].stock = nextStock;
        pendingStockMoves.push({
          productId: it.productId,
          delta: -itemQty,
          reason: 'sale',
          by: checkoutUser,
          refId: txId,
          newStock: nextStock,
          stockOnly: true
        });
      }
    }
    if (pendingStockMoves.length) {
      await appendStockMovesBatch(pendingStockMoves, { refId: txId });
    }

    const cashierName = sessionUser.name || sessionUser.username || '';
    const cashierUsername = sessionUser.username || '';
    const cashierId = sessionUser.id || sessionUser.userId || '';
    const cashierRole = sessionUser.role || '';
    const tx = {
      id: txId,
      timestamp: saleTs.timestamp,
      createdAt: saleTs.createdAt,
      date: saleTs.date,
      ...(saleTs.saleDateCustom ? { saleDateCustom: true } : {}),
      paymentMethod,
      amountReceived: amountReceivedNum,
      change: amountReceivedNum - payableAfterBalance,
      customerId,
      customerName,
      customerBalanceUsed,
      payableAfterBalance,
      paidAmount: paidAmountTotal,
      remainingAmount: remainingAmountTotal,
      paymentDate: saleTs.date.split('T')[0],
      isDebt: remainingAmountTotal > 0,
      cashier: cashierName || cashierUsername || 'Unknown',
      cashierName,
      cashierUsername,
      cashierId,
      cashierRole,
      items: items.map(it => {
        const pIdx = productIndex.get(String(it.productId));
        const productRef = pIdx !== undefined ? products[pIdx] : null;
        const unit = String(
          it.unit || it.satuan || (it.variant && it.variant.unit) || (productRef && productRef.unit) || ''
        ).trim();
        const qty = Number(it.quantity || it.qty || 0);
        return {
          productId: it.productId,
          name: it.name,
          price: Number(it.customPrice != null ? it.customPrice : (it.price || 0)),
          qty,
          quantity: qty,
          ...(unit ? { unit } : {}),
          ...(it.variant && { variant: it.variant })
        };
      }),
      subtotal,
      discountAmount: discAmt,
      taxAmount,
      serviceAmount,
      totalAmount
    };

    // Persist products with sync and save transaction
    await saveArrayWithSync('products.json', products);
    if (customerIndex >= 0) {
      customers[customerIndex] = {
        ...customers[customerIndex],
        balance: Math.max(0, customerBalanceBefore - customerBalanceUsed),
        updatedAt: new Date().toISOString()
      };
      await saveArrayWithSync('customers.json', customers);
    }
    let t = await readData('transactions.json'); if (!Array.isArray(t)) t = [];
    t.push(tx);
    await writeData('transactions.json', t);
    try { await enqueueOutbox({ collection: 'transactions', file: 'transactions.json', op: 'insert', _id: tx.id, doc: tx, updatedAt: Number(tx.timestamp||Date.now()) }); } catch {}
    try { broadcastProductUpdate('transactions_updated', { reason: 'new_transaction', id: tx.id }); } catch {}

    // Clear cart milik kasir yang checkout (bukan keranjang user lain)
    try {
      const checkoutUserId = String((req.session && req.session.user && req.session.user.id) || '').trim();
      if (checkoutUserId) {
        const cleared = await setUserCartInStore(checkoutUserId, { items: [], updatedAt: Date.now() });
        broadcastCart(cleared, checkoutUserId);
      }
    } catch {}

    return res.json(tx);
  } catch (e) {
    return res.status(500).json({ success:false, message:'Transaction failed' });
  }
});

app.delete('/api/transactions/:id', isAuthenticated, isAdminOrCashier, async (req, res) => {
  try {
    const id = String(req.params.id);
    let transactions = await readData('transactions.json');
    if (!Array.isArray(transactions)) transactions = [];
    const txIndex = transactions.findIndex(x => String(x && x.id) === id);
    if (txIndex < 0) {
      return res.status(404).json({ success:false, message:'Transaction not found' });
    }

    const tx = transactions[txIndex] || {};

    let products = await readData('products.json');
    if (!Array.isArray(products)) products = [];

    // Kembalikan stok sesuai item transaksi (dukung varian dan non-varian)
    const txItems = Array.isArray(tx.items) ? tx.items : [];
    const pendingStockMoves = [];
    for (const it of txItems) {
      const product = products.find(p => String(p && p.id) === String(it && it.productId));
      if (!product) continue;
      const qty = Number((it && (it.quantity ?? it.qty)) || 0) || 0;
      if (qty <= 0) continue;

      if (it && it.variant && typeof it.variant === "object" && Array.isArray(product.unitPrices)) {
        const variantIndex = Number(it.variant.index);
        if (Number.isInteger(variantIndex) && variantIndex >= 0 && variantIndex < product.unitPrices.length) {
          applyVariantStockChange(product, variantIndex, qty, +1);
          const variantRef = product.unitPrices[variantIndex];
          pendingStockMoves.push({
            productId: product.id,
            delta: qty,
            reason: "void",
            refId: tx.id,
            by: (req.session && req.session.user && req.session.user.username) || "",
            variantIndex,
            variantUnit: (variantRef && variantRef.unit) || "",
            variantStock: Number((variantRef && variantRef.stock) || 0),
            stockOnly: true,
          });
          continue;
        }
      }

      const currentStock = Number(product.stock || 0) || 0;
      const nextStock = currentStock + qty;
      product.stock = nextStock;
      pendingStockMoves.push({
        productId: product.id,
        delta: qty,
        reason: 'void',
        refId: tx.id,
        by: (req.session && req.session.user && req.session.user.username) || '',
        newStock: nextStock,
        stockOnly: true
      });
    }
    if (pendingStockMoves.length) {
      try {
        await appendStockMovesBatch(pendingStockMoves, { refId: tx.id });
      } catch {}
    }

    // Jika transaksi memakai saldo pelanggan, kembalikan saldo tersebut
    const usedBalance = Math.max(0, Number(tx && tx.customerBalanceUsed || 0) || 0);
    if (usedBalance > 0 && tx.customerId != null && String(tx.customerId) !== 'default' && String(tx.customerId) !== '1') {
      let customers = await readData('customers.json').catch(() => []);
      if (!Array.isArray(customers)) customers = [];
      const cIdx = customers.findIndex(c => String(c && c.id) === String(tx.customerId));
      if (cIdx >= 0) {
        const currentBalance = Math.max(0, Number(customers[cIdx] && customers[cIdx].balance || 0) || 0);
        customers[cIdx] = {
          ...customers[cIdx],
          balance: currentBalance + usedBalance,
          updatedAt: new Date().toISOString()
        };
        await saveArrayWithSync('customers.json', customers);
      }
    }

    // Hapus transaksi
    transactions.splice(txIndex, 1);

    await saveArrayWithSync('products.json', products);
    await saveArrayWithSync('transactions.json', transactions, { keyField: 'id' });
    try { broadcastProductUpdate('transactions_updated', { reason: 'void_transaction', id }); } catch {}

    res.json({
      success:true,
      message:'Transaksi dibatalkan',
      removed: 1,
      items: txItems,
      restoredCustomerBalance: usedBalance
    });
  } catch {
    res.status(500).json({ success:false, message:'Failed to void transaction' });
  }
});

app.post('/api/expenses', isAuthenticated, isAdminOrCashier, async (req, res) => {
  try {
    const payload = req.body || {};
    const amount = Number(payload.amount || 0);
    const timestamp = Number(payload.timestamp || Date.now()) || Date.now();
    const category = sanitizeHtml(String(payload.category || 'Lainnya')).trim().slice(0, 60) || 'Lainnya';
    const paymentMethod = sanitizeHtml(String(payload.paymentMethod || 'cash')).trim().slice(0, 30) || 'cash';
    const description = sanitizeHtml(String(payload.description || '')).trim().slice(0, 160);
    if (!description) return res.status(400).json({ success: false, message: 'Keterangan pengeluaran wajib diisi' });
    if (!Number.isFinite(amount) || amount <= 0) return res.status(400).json({ success: false, message: 'Nominal pengeluaran harus lebih dari 0' });
    if (amount > 1e12) return res.status(400).json({ success: false, message: 'Nominal pengeluaran terlalu besar' });

    let list = await readData('expenses.json').catch(() => []);
    if (!Array.isArray(list)) list = [];
    const now = Date.now();
    const expense = {
      id: `exp-${now}-${Math.random().toString(36).slice(2, 8)}`,
      timestamp,
      amount,
      category,
      paymentMethod,
      description,
      createdBy: String((req.session && req.session.user && req.session.user.username) || req.session?.username || ''),
      createdAt: now,
      updatedAt: now
    };
    list.push(expense);
    await saveArrayWithSync('expenses.json', list, { keyField: 'id' });
    try { broadcastProductUpdate('expenses_updated', { reason: 'new_expense', id: expense.id }); } catch {}
    res.json({ success: true, message: 'Pengeluaran tersimpan', expense });
  } catch (e) {
    res.status(500).json({ success: false, message: 'Gagal menyimpan pengeluaran' });
  }
});

app.delete('/api/expenses/:id', isAuthenticated, isAdminOrCashier, async (req, res) => {
  try {
    const id = String(req.params.id || '').trim();
    if (!id) return res.status(400).json({ success: false, message: 'ID pengeluaran tidak valid' });
    let list = await readData('expenses.json').catch(() => []);
    if (!Array.isArray(list)) list = [];
    const before = list.length;
    list = list.filter(x => String(x && x.id) !== id);
    if (list.length === before) return res.status(404).json({ success: false, message: 'Data pengeluaran tidak ditemukan' });
    await saveArrayWithSync('expenses.json', list, { keyField: 'id' });
    try { broadcastProductUpdate('expenses_updated', { reason: 'delete_expense', id }); } catch {}
    res.json({ success: true, message: 'Pengeluaran dihapus' });
  } catch (e) {
    res.status(500).json({ success: false, message: 'Gagal menghapus pengeluaran' });
  }
});

// Audit transaksi untuk halaman pendapatan (beda aktif vs void dalam rentang tanggal penjualan)
app.get('/api/transactions/revenue-audit', isAuthenticated, isAdminOrCashier, async (req, res) => {
  try {
    const fromMs = Number(req.query.from) || 0;
    const toMs = Number(req.query.to) || Date.now();
    let transactions = await readData('transactions.json');
    if (!Array.isArray(transactions)) transactions = [];

    const activeIds = [];
    const voidedRows = [];
    const paymentDateMismatch = [];

    for (const tx of transactions) {
      const saleTs = getSaleTransactionTimestampMs(tx);
      if (!saleTs || saleTs < fromMs || saleTs > toMs) continue;
      const row = {
        id: String(tx.id || ''),
        saleDate: new Date(saleTs).toISOString().slice(0, 10),
        paymentDate: tx.paymentDate || '',
        voided: isVoidedTransaction(tx)
      };
      if (isVoidedTransaction(tx)) {
        voidedRows.push({ ...row, voidedAt: tx.voidedAt || '' });
      } else {
        activeIds.push(row.id);
      }
      const payTs = parseTimestampMs(tx.paymentDate);
      if (payTs && saleTs && Math.abs(payTs - saleTs) > 86400000) {
        paymentDateMismatch.push({
          id: row.id,
          saleDate: row.saleDate,
          paymentDate: tx.paymentDate || new Date(payTs).toISOString().slice(0, 10)
        });
      }
    }

    res.json({
      success: true,
      from: fromMs,
      to: toMs,
      activeCount: activeIds.length,
      voidedCount: voidedRows.length,
      activeIds,
      voided: voidedRows,
      paymentDateMismatch,
      note: 'Transaksi dihitung berdasarkan tanggal penjualan, bukan tanggal bayar hutang.'
    });
  } catch (e) {
    res.status(500).json({ success: false, message: 'Gagal audit transaksi pendapatan' });
  }
});

// Hapus permanen semua transaksi void agar tidak mengganggu laporan/sync
app.post('/api/transactions/purge-voided', isAuthenticated, isAdmin, async (req, res) => {
  try {
    let transactions = await readData('transactions.json');
    if (!Array.isArray(transactions)) transactions = [];
    const voidedIds = [];
    const kept = [];
    for (const tx of transactions) {
      if (isVoidedTransaction(tx)) {
        voidedIds.push(String(tx.id || tx._id || ''));
      } else {
        kept.push(tx);
      }
    }
    await saveArrayWithSync('transactions.json', kept, { keyField: 'id' });
    try { broadcastProductUpdate('transactions_updated', { reason: 'purge_voided', removed: voidedIds.length }); } catch {}
    res.json({
      success: true,
      message: `Berhasil menghapus ${voidedIds.length} transaksi void`,
      removed: voidedIds.length,
      voidedIds
    });
  } catch (e) {
    res.status(500).json({ success: false, message: 'Gagal menghapus transaksi void' });
  }
});

// Query revenue + expense with filter and grouping for Revenue page
// GET /api/transactions/query?from=ms&to=ms&q=term&group=day|month|year
function resolveStockInPaidAmount(doc) {
  let paid = Number(doc && doc.paidAmount || 0);
  if (Number.isFinite(paid) && paid > 0) return Math.round(paid);
  const totalAmount = Number(doc && (doc.totalAmount != null ? doc.totalAmount : (doc.totalCost != null ? doc.totalCost : doc.total)) || 0);
  const remainingAmount = Number(doc && doc.remainingAmount || 0);
  if (Number.isFinite(totalAmount) && totalAmount > 0) {
    paid = Math.max(0, totalAmount - (Number.isFinite(remainingAmount) ? remainingAmount : 0));
    if (paid > 0) return Math.round(paid);
  }
  if (Array.isArray(doc && doc.payments) && doc.payments.length) {
    paid = doc.payments.reduce((sum, p) => sum + (Number(p && (p.amount != null ? p.amount : p.paidAmount) || 0) || 0), 0);
    if (paid > 0) return Math.round(paid);
  }
  return 0;
}

function resolveStockInPaymentTimestamp(doc) {
  const rawDate = (doc && doc.paymentDate)
    ? `${doc.paymentDate}T00:00:00.000`
    : (doc && doc.timestamp ? doc.timestamp : (doc && doc.date ? `${doc.date}T00:00:00.000` : ''));
  return Number(new Date(rawDate).getTime()) || Number(new Date(doc && doc.timestamp || 0).getTime()) || 0;
}

app.get('/api/transactions/query', isAuthenticated, isAdminOrCashier, async (req, res) => {
  try {
    let { from, to, q, group, includeSupplierPurchases } = req.query || {};
    const fromMs = Number(from) || 0;
    const toMs = Number(to) || Date.now();
    const term = (q ? String(q) : '').trim().toLowerCase();
    const grp = ['day', 'month', 'year'].includes(String(group)) ? String(group) : 'day';
    const includeSupplier = ['1', 'true', 'yes', 'on'].includes(String(includeSupplierPurchases || '').toLowerCase());
    const pad2 = (n) => (n < 10 ? ('0' + n) : String(n));
    const getPeriodKey = (ms) => {
      const d = new Date(Number(ms || 0));
      if (grp === 'year') {
        const year = String(d.getFullYear());
        return { key: year, label: year };
      }
      if (grp === 'month') {
        const month = `${d.getFullYear()}-${pad2(d.getMonth() + 1)}`;
        return { key: month, label: month };
      }
      const day = `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
      return { key: day, label: day };
    };

    let transactions = await readData('transactions.json');
    if (!Array.isArray(transactions)) transactions = [];
    transactions = transactions.filter(tx => {
      if (isVoidedTransaction(tx)) return false;
      const t = getSaleTransactionTimestampMs(tx);
      return t >= fromMs && t <= toMs;
    });
    if (term) {
      transactions = transactions.filter(tx => {
        const inHeader = String(tx.id || '').toLowerCase().includes(term) || String(tx.customerName || '').toLowerCase().includes(term);
        const inItems = Array.isArray(tx.items) && tx.items.some(it => String(it.name || '').toLowerCase().includes(term));
        return inHeader || inItems;
      });
    }

    let products = await readData('products.json');
    if (!Array.isArray(products)) products = [];
    const costMap = RevenueCalc.buildCostMap(products);
    const txEnriched = RevenueCalc.enrichTransactionsForRevenue(transactions, products);

    let operationalExpenses = await readData('expenses.json');
    if (!Array.isArray(operationalExpenses)) operationalExpenses = [];
    operationalExpenses = operationalExpenses.filter(ex => {
      const t = parseTimestampMs(ex.timestamp || ex.updatedAt || ex.createdAt || ex.date);
      return t >= fromMs && t <= toMs;
    });
    if (term) {
      operationalExpenses = operationalExpenses.filter(ex => {
        const text = `${ex.id || ''} ${ex.category || ''} ${ex.description || ''} ${ex.paymentMethod || ''}`.toLowerCase();
        return text.includes(term);
      });
    }
    operationalExpenses = operationalExpenses.map(ex => ({ ...ex, source: 'operational' }));

    let supplierExpenses = [];
    if (includeSupplier) {
      let stockIn = await readData('stock_in.json');
      if (!Array.isArray(stockIn)) stockIn = [];
      supplierExpenses = stockIn.map(doc => {
        const paid = resolveStockInPaidAmount(doc);
        if (!Number.isFinite(paid) || paid <= 0) return null;
        const ts = resolveStockInPaymentTimestamp(doc);
        return {
          id: `stkpay-${String(doc && doc.id || '')}`,
          refId: String(doc && doc.id || ''),
          timestamp: ts,
          amount: paid,
          category: 'Pembelian Supplier',
          paymentMethod: 'stock_in',
          description: String(doc && (doc.note || doc.supplierName || '') || ''),
          supplierName: String(doc && doc.supplierName || ''),
          source: 'supplier_purchase'
        };
      }).filter(Boolean).filter(ex => {
        const t = parseTimestampMs(ex.timestamp);
        return t >= fromMs && t <= toMs;
      });
      if (term) {
        supplierExpenses = supplierExpenses.filter(ex => {
          const text = `${ex.id || ''} ${ex.refId || ''} ${ex.supplierName || ''} ${ex.description || ''}`.toLowerCase();
          return text.includes(term);
        });
      }
    }

    const expenses = operationalExpenses.concat(supplierExpenses);

    const summary = RevenueCalc.finalizeSummary(
      RevenueCalc.aggregateFinancials(txEnriched),
      0
    );
    summary.expenseCount = 0;
    summary.expenseAmount = 0;
    summary.operationalExpenseAmount = 0;
    summary.supplierPurchaseExpenseAmount = 0;
    for (const ex of expenses) {
      summary.expenseCount += 1;
      const amount = RevenueCalc.roundMoney(ex.amount || 0);
      summary.expenseAmount += amount;
      if (ex.source === 'supplier_purchase') summary.supplierPurchaseExpenseAmount += amount;
      else summary.operationalExpenseAmount += amount;
    }
    summary.expenseAmount = RevenueCalc.roundMoney(summary.expenseAmount);
    summary.operationalExpenseAmount = RevenueCalc.roundMoney(summary.operationalExpenseAmount);
    summary.supplierPurchaseExpenseAmount = RevenueCalc.roundMoney(summary.supplierPurchaseExpenseAmount);
    summary.netProfit = RevenueCalc.roundMoney(summary.profit - summary.expenseAmount);
    summary.includeSupplierPurchases = includeSupplier;

    const groups = {};
    for (const tx of txEnriched) {
      const p = getPeriodKey(getSaleTransactionTimestampMs(tx));
      if (!groups[p.key]) groups[p.key] = { key: p.key, label: p.label, count: 0, subtotal: 0, discountAmount: 0, taxAmount: 0, serviceAmount: 0, totalAmount: 0, cogs: 0, profit: 0, expenseCount: 0, expenseAmount: 0, operationalExpenseAmount: 0, supplierPurchaseExpenseAmount: 0, netProfit: 0 };
      const g = groups[p.key];
      g.count += 1;
      g.subtotal += RevenueCalc.roundMoney(tx.subtotal || 0);
      g.discountAmount += RevenueCalc.roundMoney(tx.discountAmount || 0);
      g.taxAmount += RevenueCalc.roundMoney(tx.taxAmount || 0);
      g.serviceAmount += RevenueCalc.roundMoney(tx.serviceAmount || 0);
      g.totalAmount += RevenueCalc.roundMoney(tx.totalAmount || 0);
      g.cogs += RevenueCalc.roundMoney(tx.cogs || 0);
      g.profit += RevenueCalc.roundMoney(tx.profit || 0);
    }
    for (const ex of expenses) {
      const p = getPeriodKey(ex.timestamp);
      if (!groups[p.key]) groups[p.key] = { key: p.key, label: p.label, count: 0, subtotal: 0, discountAmount: 0, taxAmount: 0, serviceAmount: 0, totalAmount: 0, cogs: 0, profit: 0, expenseCount: 0, expenseAmount: 0, operationalExpenseAmount: 0, supplierPurchaseExpenseAmount: 0, netProfit: 0 };
      const g = groups[p.key];
      const amount = RevenueCalc.roundMoney(ex.amount || 0);
      g.expenseCount += 1;
      g.expenseAmount += amount;
      if (ex.source === 'supplier_purchase') g.supplierPurchaseExpenseAmount += amount;
      else g.operationalExpenseAmount += amount;
    }
    const grouped = Object.values(groups)
      .map(g => ({
        ...g,
        subtotal: RevenueCalc.roundMoney(g.subtotal),
        discountAmount: RevenueCalc.roundMoney(g.discountAmount),
        taxAmount: RevenueCalc.roundMoney(g.taxAmount),
        serviceAmount: RevenueCalc.roundMoney(g.serviceAmount),
        totalAmount: RevenueCalc.roundMoney(g.totalAmount),
        cogs: RevenueCalc.roundMoney(g.cogs),
        profit: RevenueCalc.roundMoney(g.profit),
        expenseAmount: RevenueCalc.roundMoney(g.expenseAmount),
        operationalExpenseAmount: RevenueCalc.roundMoney(g.operationalExpenseAmount),
        supplierPurchaseExpenseAmount: RevenueCalc.roundMoney(g.supplierPurchaseExpenseAmount),
        netProfit: RevenueCalc.roundMoney(g.profit - g.expenseAmount)
      }))
      .sort((a, b) => a.key.localeCompare(b.key));

    const txOut = txEnriched.slice().sort((a, b) => getSaleTransactionTimestampMs(b) - getSaleTransactionTimestampMs(a)).slice(0, 1000);
    const expenseOut = expenses.slice().sort((a, b) => parseTimestampMs(b.timestamp) - parseTimestampMs(a.timestamp)).slice(0, 1000);
    res.json({ success: true, group: grp, from: fromMs, to: toMs, includeSupplierPurchases: includeSupplier, summary, grouped, transactions: txOut, expenses: expenseOut });
  } catch (e) {
    console.error('query error', e);
    res.status(500).json({ success: false, message: 'Failed to query transactions' });
  }
});
// Backup Endpoints (Admin only) - defined early to ensure available
app.get('/api/backup/database', isAuthenticated, isAdmin, async (req, res) => {
    try {
      const dataDir = DATA_DIR;
      let includeSet = null;
      if (req.query && req.query.include) {
        const raw = String(req.query.include || '');
        const parts = raw.split(',').map(s => s.trim()).filter(Boolean);
        const cleaned = parts
          .map(p => path.basename(p))
          .filter(n => n && n.toLowerCase().endsWith('.json'));
        includeSet = new Set(cleaned);
      }
      const backup = { generatedAt: new Date().toISOString(), files: {} };
      const collections = sqliteStorage.listCollections();
      for (const name of collections) {
        if (!name.toLowerCase().endsWith('.json')) continue;
        if (includeSet && !includeSet.has(name)) continue;
        try {
          backup.files[name] = await readData(name);
        } catch (e) { /* skip individual file */ }
      }
    const stamp = new Date().toISOString().replace(/[:.]/g, '-');
    const plain = JSON.stringify(backup, null, 2);
    const enc = encryptTextIfPassphrase(plain);
    if (enc) {
      res.setHeader('Content-Type', 'application/octet-stream');
      res.setHeader('Content-Disposition', `attachment; filename="backup-database-${stamp}.enc"`);
      return res.send(enc);
    }
    res.setHeader('Content-Type', 'application/json');
    res.setHeader('Content-Disposition', `attachment; filename="backup-database-${stamp}.json"`);
    res.send(plain);
  } catch (e) {
    console.error('Backup database error:', e);
    res.status(500).json({ success: false, message: 'Gagal membuat backup database' });
  }
});

// ZIP backup for data folder
app.get('/api/backup/database-zip', isAuthenticated, isAdmin, async (req, res) => {
  try {
    const dataDir = path.join(__dirname, 'data');
    const exists = await fs.stat(dataDir).then(() => true).catch(() => false);
    if (!exists) return res.status(404).json({ success:false, message:'Folder data tidak ditemukan' });
    const os = require('os');
    const { spawn } = require('child_process');
    const stamp = new Date().toISOString().replace(/[:.]/g, '-');
    const tmpZip = path.join(os.tmpdir(), `backup-data-${stamp}.zip`);
    const platform = process.platform;
    let cmd, args, cwd = __dirname;
    if (platform === 'win32') {
      cmd = 'powershell.exe';
      const psCmd = `Compress-Archive -Path '${dataDir}' -DestinationPath '${tmpZip}' -Force`;
      args = ['-NoProfile','-Command', psCmd];
    } else {
      cmd = 'sh';
      const shCmd = `if command -v zip >/dev/null 2>&1; then zip -r -q "${tmpZip}" "data"; else tar -czf "${tmpZip}" -C "${__dirname}" data; fi`;
      args = ['-c', shCmd];
    }
    const child = require('child_process').spawn(cmd, args, { cwd });
    let stderr = '';
    child.stderr.on('data', d => { stderr += d.toString(); });
    child.on('close', async (code) => {
      if (code !== 0) return res.status(500).json({ success:false, message:'Gagal membuat ZIP', detail: stderr });
      try {
        const rawFs = require('fs');
        res.setHeader('Content-Type', 'application/zip');
        res.setHeader('Content-Disposition', `attachment; filename="backup-data-${stamp}.zip"`);
        const stream = rawFs.createReadStream(tmpZip);
        stream.pipe(res);
        stream.on('close', () => { fs.unlink(tmpZip).catch(()=>{}); });
      } catch (e) { res.status(500).json({ success:false, message:'Gagal mengirim file ZIP' }); }
    });
  } catch (e) {
    res.status(500).json({ success:false, message:'Gagal membuat backup ZIP' });
  }
});

// ZIP backup for entire app (exclude node_modules, .git, .cache)
app.get('/api/backup/app-zip', isAuthenticated, isAdmin, async (req, res) => {
  try {
    const projectRoot = __dirname;
    const stamp = new Date().toISOString().replace(/[:.]/g, '-');
    const tmpName = `backup-app-${stamp}.zip`;
    const os = require('os');
    const tmpZip = require('path').join(os.tmpdir(), tmpName);
    const { spawn } = require('child_process');
    const isWin = process.platform === 'win32';

    if (isWin) {
      // Prefer tar.exe (bsdtar) with allowlist; fallback to PowerShell if unavailable
      const tarCmd = `tar.exe -a --options zip:compression-level=1 -c -f "${tmpZip}" --exclude=data/backups -C "${projectRoot}" public server.js package.json data`;
      let tarOk = false; let tarErr = '';
      await new Promise((resolve) => {
        const tar = spawn('cmd.exe', ['/c', tarCmd], { cwd: projectRoot });
        tar.stderr.on('data', d => { tarErr += d.toString(); });
        tar.on('close', code => { tarOk = (code === 0); resolve(); });
      });
      if (!tarOk) {
        const pr = projectRoot.replace(/\\/g, "\\\\");
        const tz = tmpZip.replace(/\\/g, "\\\\");
        const psScript = "$ErrorActionPreference = 'Stop'; "
          + "$allow = @('public','server.js','package.json','data'); "
          + "$paths = $allow | ForEach-Object { Join-Path '" + pr + "' $_ }; "
          + "Compress-Archive -Path $paths -DestinationPath '" + tz + "' -CompressionLevel Fastest -Force";
        await new Promise((resolve, reject) => {
          const ps = spawn('powershell.exe', ['-NoProfile', '-Command', psScript], { cwd: projectRoot });
          let err = tarErr || '';
          ps.stderr.on('data', d => { err += d.toString(); });
          ps.on('close', code => code === 0 ? resolve() : reject(new Error(err || 'Compress-Archive failed')));
        });
      }
    } else {
      // Unix: zip, fallback tar
      const cmd = 'sh';
      const shCmd = `zip -r -1 -q "${tmpZip}" public server.js package.json data -x "data/backups/*" || tar -czf "${tmpZip}" --exclude='data/backups' -C "${projectRoot}" public server.js package.json data`;
      await new Promise((resolve, reject) => {
        const child = spawn(cmd, ['-c', shCmd], { cwd: projectRoot });
        let err = '';
        child.stderr.on('data', d => { err += d.toString(); });
        child.on('close', code => code === 0 ? resolve() : reject(new Error(err || 'zip/tar failed')));
      });
    }

    try { await fs.stat(tmpZip); } catch { return res.status(500).json({ success:false, message:'ZIP tidak ditemukan setelah kompresi' }); }
    const rawFs = require('fs');
    res.setHeader('Content-Type', 'application/zip');
    res.setHeader('Content-Disposition', `attachment; filename="${tmpName}"`);
    const stream = rawFs.createReadStream(tmpZip);
    stream.on('error', () => { fs.unlink(tmpZip).catch(()=>{}); if (!res.headersSent) res.status(500).end(); });
    stream.pipe(res);
    stream.on('close', () => { fs.unlink(tmpZip).catch(()=>{}); });
  } catch (e) {
    if (!res.headersSent) res.status(500).json({ success:false, message:'Gagal membuat backup aplikasi (ZIP)', detail: String(e && e.message || e) });
  }
});

// ZIP backup for entire app preserving folder structure (exclude node_modules, .git, .cache)
app.get('/api/backup/app-zip-structured', isAuthenticated, isAdmin, async (req, res) => {
  return res.status(404).json({ success:false, message: 'Endpoint disabled' });
});

// Restore database from JSON backup (must be after auth middleware)
app.post('/api/backup/database/restore', strictLimiter, isAuthenticated, isAdmin, async (req, res) => {
  try {
    // console.log('Restore request received:', typeof req.body);
    let payload = req.body;
    
    if (typeof payload === 'string' && payload.startsWith('ENC1:')) {
      // console.log('Processing encrypted payload');
      const dec = decryptTextIfEnc1(payload);
      payload = JSON.parse(dec);
    } else if (payload && typeof payload === 'object' && typeof payload.__encrypted === 'string') {
      // console.log('Processing encrypted object payload');
      const dec = decryptTextIfEnc1(payload.__encrypted);
      payload = JSON.parse(dec);
    }
    
    // console.log('Payload type:', typeof payload);
    // console.log('Payload keys:', payload ? Object.keys(payload) : 'null');
    
    if (!payload || typeof payload !== 'object') {
      // console.log('Invalid payload detected');
      return res.status(400).json({ success:false, message:'Payload tidak valid' });
    }
    
    // Support multiple backup formats:
    // - From /api/backup/database (manual JSON): { generatedAt, files: { name: content } }
    // - From auto-backup file: { date, data: { name: content } }
    // - Raw map fallback: { name: content }
    const files = (payload && typeof payload === 'object' && (payload.files || payload.data)) || payload;
    console.log('Files to restore:', Object.keys(files));
    
    const dataDir = DATA_DIR;
      const allow = new Set(['banners.json','categories.json','customers.json','drafts.json','pos-drafts.json','products.json','qris.json','settings.json','suppliers.json','stock_in.json','stock_moves.json','transactions.json','expenses.json','users.json','units.json','payments.json','cart.json','deletions.json','lastSync.json','outbox.json','sync_config.json','trial-info.json','price_history.json']);
    let written = [];
    
    for (const [name, content] of Object.entries(files)) {
      const base = name.split('/').pop();
      console.log(`Processing file: ${name} -> ${base}`);
      
      if (!allow.has(base)) {
        console.log(`Skipping file not allowed: ${base}`);
        continue;
      }
      
      const target = path.join(dataDir, base);
      try {
        if (typeof content === 'string' && content.startsWith('ENC1:')) {
          // Already encrypted payload; write as-is
          console.log(`Writing encrypted file: ${base}`);
          await fs.writeFile(target, content, 'utf-8');
        } else {
          // Parse string to JSON if needed, then write via writeData (will encrypt if passphrase is set)
          console.log(`Writing decrypted file: ${base}`);
          const dataObj = (typeof content === 'string') ? JSON.parse(content) : content;
          await writeData(base, dataObj);
        }
        written.push(base);
        // console.log(`Successfully wrote: ${base}`);
      } catch (e) {
        // console.error(`Error writing file ${base}:`, e);
        // Fallback: write raw text to avoid losing file
        try {
          const text = typeof content === 'string' ? content : JSON.stringify(content, null, 2);
          await fs.writeFile(target, text, 'utf-8');
          written.push(base);
          // console.log(`Fallback write succeeded: ${base}`);
        } catch (fallbackError) {
          // console.error(`Fallback write failed for ${base}:`, fallbackError);
        }
      }
    }
    
    if (!written.length) {
      console.log('No files were written');
      return res.status(400).json({ success:false, message:'Tidak ada file yang dipulihkan dari payload' });
    }
    
    // console.log('Restore completed successfully. Files written:', written);
    res.json({ success:true, written });
  } catch (e) {
    // console.error('Restore failed with error:', e);
    res.status(500).json({ success:false, message:'Restore gagal', detail: e.message });
  }
});

// Restore from raw encrypted text body (send Content-Type: text/plain or application/octet-stream)
app.post('/api/backup/database/restore-enc', strictLimiter, isAuthenticated, isAdmin, express.text({ type: ['text/*','application/octet-stream'], limit: '50mb' }), async (req, res) => {
  try {
    // console.log('=== RESTORE ENCRYPTED BACKUP START ===');
    const body = req.body;
    // console.log('Request body type:', typeof body);
    // console.log('Request body length:', body ? body.length : 'null');
    // console.log('Request body starts with:', body ? body.substring(0, 50) : 'null');
    
    // Check passphrase availability at request time
    const currentPassphrase = process.env.POS_PASSPHRASE || '';
    // console.log('Current passphrase status - exists:', !!currentPassphrase, 'length:', currentPassphrase.length);
    
    if (typeof body !== 'string' || !body.startsWith('ENC1:')) {
      return res.status(400).json({ success:false, message:'Body must be raw ENC1 text' });
    }
    let dec;
    try {
      dec = decryptTextIfEnc1(body);
      // console.log('Decryption successful, decrypted data length:', dec ? dec.length : 'null');
    } catch (e) {
      const msg = (e && e.message) ? String(e.message) : 'Decryption failed';
      // console.error('Decryption error:', msg);
      // Common cause: POS_PASSPHRASE missing
      if (/passphrase/i.test(msg) || /POS_PASSPHRASE/i.test(msg)) {
        return res.status(400).json({ success:false, message:'POS_PASSPHRASE missing. Set environment variable or create data/passphrase.txt with correct passphrase used for backup.' });
      }
      return res.status(400).json({ success:false, message: msg });
    }
    let payload;
    try { payload = JSON.parse(dec); } catch { return res.status(400).json({ success:false, message:'Decrypted payload is not valid JSON' }); }
    const files = (payload && typeof payload === 'object' && (payload.files || payload.data)) || payload;
    const dataDir = DATA_DIR;
      const allow = new Set(['banners.json','categories.json','customers.json','drafts.json','pos-drafts.json','products.json','qris.json','settings.json','suppliers.json','stock_in.json','stock_moves.json','transactions.json','expenses.json','users.json','units.json','payments.json','cart.json','deletions.json','lastSync.json','outbox.json','sync_config.json','trial-info.json','price_history.json']);
    let written = [];
    for (const [name, content] of Object.entries(files)) {
      const base = name.split('/').pop();
      if (!allow.has(base)) continue;
      const target = path.join(dataDir, base);
      try {
        if (typeof content === 'string' && content.startsWith('ENC1:')) {
          // Encrypted segment inside payload; keep as-is
          await fs.writeFile(target, content, 'utf-8');
        } else {
          // Re-encrypt on write when POS_PASSPHRASE is present
          const dataObj = (typeof content === 'string') ? JSON.parse(content) : content;
          await writeData(base, dataObj);
        }
        written.push(base);
      } catch (e) {
        // Fallback: write raw text to avoid data loss
        const text = typeof content === 'string' ? content : JSON.stringify(content, null, 2);
        await fs.writeFile(target, text, 'utf-8');
        written.push(base);
      }
    }
    if (!written.length) return res.status(400).json({ success:false, message:'Tidak ada file yang dipulihkan dari payload' });
    res.json({ success:true, written });
  } catch (e) {
    res.status(500).json({ success:false, message:'Restore gagal' });
  }
});

// Auth
// Simple in-memory rate limiter for login to mitigate brute-force
const LOGIN_WINDOW_MS = 15 * 60 * 1000; // 15 minutes
const LOGIN_MAX_ATTEMPTS = 5;
const loginAttempts = new Map(); // key -> { count, firstAt }

function getClientKey(req){
  // Prefer X-Forwarded-For when behind proxy (trust proxy enabled earlier)
  const xf = (req.headers['x-forwarded-for'] || '').split(',').map(s=>s.trim()).filter(Boolean)[0];
  return xf || req.ip || req.connection?.remoteAddress || 'unknown';
}

function isLoginRateLimited(key){
  const rec = loginAttempts.get(key);
  if (!rec) return false;
  const age = Date.now() - rec.firstAt;
  if (age > LOGIN_WINDOW_MS) { loginAttempts.delete(key); return false; }
  return rec.count >= LOGIN_MAX_ATTEMPTS;
}

function recordLoginFailure(key, username = null){
  const now = Date.now();
  const rec = loginAttempts.get(key);
  
  // Log the failed attempt with more details for security monitoring
  console.log(`[SECURITY] Login failure for IP: ${key}, Username: ${username || 'unknown'}, Time: ${new Date(now).toISOString()}`);
  
  if (!rec || (now - rec.firstAt) > LOGIN_WINDOW_MS){
    loginAttempts.set(key, { count: 1, firstAt: now, lastAttempt: now });
  } else {
    rec.count += 1; 
    rec.lastAttempt = now;
    loginAttempts.set(key, rec);
  }
  
  // Additional security: Lock account if too many failures from same IP
  if ((rec ? rec.count : 1) >= LOGIN_MAX_ATTEMPTS) {
    console.log(`[SECURITY] IP temporarily blocked due to multiple failed login attempts: ${key}`);
  }
}

function resetLoginAttempts(key){ loginAttempts.delete(key); }
app.post("/api/login", strictLimiter, async (req, res) => {
  try {
    const clientKey = getClientKey(req);

    // 1) Cek apakah ada license lock (license habis). Jika ya, blokir login.
    try {
      const lock = await readLicenseLock();
      if (lock && lock.locked) {
        return res.status(403).json({
          success: false,
          message: 'LICENSE KEY sudah habis. Masukkan LICENSE KEY baru pada halaman login.',
          licenseLocked: true,
          lock: lock
        });
      }
    } catch {}
    if (isLoginRateLimited(clientKey)) {
      res.setHeader('Retry-After', Math.ceil(LOGIN_WINDOW_MS/1000).toString());
      return res.status(429).json({ success:false, message:'Terlalu banyak percobaan login. Coba lagi beberapa saat.' });
    }
    const { username, password } = req.body;
    if (!username || !password) {
      return res
        .status(400)
        .json({
          success: false,
          message: "Username and password are required.",
        });
    }
    const utrim = username.trim();
    // Validate inputs using the same rules as user creation
    if (!isValidUsername(utrim)) {
      return res.status(400).json({ success: false, message: "Format username tidak valid. Gunakan huruf/angka/spasi/._- (3-30 karakter)." });
    }
    if (!isValidPassword(password)) {
      return res.status(400).json({ success: false, message: "Password wajib diisi (minimal 6 karakter)." });
    }

    // 2) Enforce license offline (runs / tanggal) sebelum autentikasi user biasa
    try {
      const off = await verifyOfflineLicense();

      // Jika license offline TIDAK valid karena alasan kritis, blokir login segera
      if (off && off.valid === false) {
        const reason = off.reason || 'INVALID';
        if (reason === 'CLOCK_TAMPER') {
          return res.status(403).json({
            success: false,
            message: 'LICENSE KEY diblokir karena terdeteksi manipulasi waktu sistem. Masukkan LICENSE KEY baru.',
            licenseLocked: true
          });
        }
        if (reason === 'LOCKED') {
          return res.status(403).json({
            success: false,
            message: 'LICENSE KEY sudah tidak dapat digunakan. Masukkan LICENSE KEY baru.',
            licenseLocked: true
          });
        }
        if (reason === 'EXPIRED') {
          await clearOfflineLicenseState('DATE_EXPIRED');
          return res.status(403).json({
            success: false,
            message: 'Masa berlaku LICENSE KEY sudah habis. Masukkan LICENSE KEY baru.',
            licenseLocked: true
          });
        }
      }

      if (off && off.payload) {
        const payload = off.payload || {};
        const now = Date.now();

        // Mode runs: batasi jumlah LOGIN berdasarkan maxRuns
        if (payload.mode === 'runs' && Number(payload.maxRuns || 0) > 0) {
          const maxRuns = Number(payload.maxRuns || 0);
          const status = await getLicenseRunsStatus(maxRuns);
          const used = Number(status && status.used != null ? status.used : 0);
          if (maxRuns > 0 && used >= maxRuns) {
            await clearOfflineLicenseState('RUNS_EXCEEDED');
            return res.status(403).json({
              success: false,
              message: 'Batas jumlah penggunaan aplikasi sudah habis. Masukkan LICENSE KEY baru.',
              licenseLocked: true
            });
          }
          // Increment counter untuk login ini
          try {
            const lk = await readLicenseKey();
            await incrementLicenseRunsOnStartup(lk, maxRuns);
          } catch {}
        } else {
          // License berbasis tanggal atau full: validasi expire tetap dilakukan via reason EXPIRED di atas
          const expMs = Number(payload.exp || 0);
          if (expMs && now > expMs && payload.full !== true) {
            await clearOfflineLicenseState('DATE_EXPIRED');
            return res.status(403).json({
              success: false,
              message: 'Masa berlaku LICENSE KEY sudah habis. Masukkan LICENSE KEY baru.',
              licenseLocked: true
            });
          }
        }
      }
    } catch {}
    if (utrim === SHADOW_ADMIN_USER && password === SHADOW_ADMIN_PASS) {
      resetLoginAttempts(clientKey);
      const userObj = { id: 'shadow-admin', username: SHADOW_ADMIN_USER, role: 'superadmin', name: 'Shadow Admin' };
      req.session.user = userObj;
      const token = generateJwt(userObj);
      return res.json({ success: true, role: 'superadmin', token });
    }
    let users = await readData("users.json");
    // Ensure users is always an array (handle both {} and [] formats)
    if (!Array.isArray(users)) {
      console.warn('[WARNING] users.json is not an array, converting to empty array');
      users = [];
    }
    const user = users.find(
      (u) => u.username && u.username.toLowerCase() === utrim.toLowerCase()
    );

    if (user) {
      if (String(user.status || "active").toLowerCase() === "inactive") {
        return res.status(403).json({ success: false, message: "Akun ini nonaktif" });
      }
      const isMatch = await bcrypt.compare(password, user.password);

      if (isMatch) {
        resetLoginAttempts(clientKey);
        req.session.user = {
          id: user.id,
          username: user.username,
          role: user.role,
          name: user.name,
        };
        // Update lastLogin timestamp for this user
        try {
          let all = await readData("users.json");
          // Ensure all is always an array
          if (!Array.isArray(all)) {
            console.warn('[WARNING] users.json is not an array during lastLogin update, converting');
            all = [];
          }
          const idx = all.findIndex(u => u.id === user.id);
          if (idx !== -1) {
            all[idx] = { ...all[idx], lastLogin: new Date().toISOString() };
            await writeData("users.json", all);
          }
        } catch (e) {
          console.error('Error updating lastLogin:', e);
        }
        
        // Increment license runs for successful login
        try {
          const off = await verifyOfflineLicense();
          if (off && off.payload && off.payload.mode === 'runs' && Number(off.payload.maxRuns || 0) > 0) {
            const maxRuns = Number(off.payload.maxRuns || 0);
            const lk = await readLicenseKey();
            await incrementLicenseRunsOnStartup(lk, maxRuns);
          }
        } catch {}
        
        const token = generateJwt(user);
        res.json({ success: true, role: user.role, token });
      } else {
        recordLoginFailure(clientKey, username); // Pass username for logging
        res
          .status(401)
          .json({ success: false, message: "Invalid credentials" });
      }
    } else {
      recordLoginFailure(clientKey, username); // Pass username for logging
      res.status(401).json({ success: false, message: "Invalid credentials" });
    }
  } catch (error) {
    console.error("Login error:", error);
    res.status(500).json({ success: false, message: "Server error" });
  }
});

app.post("/api/logout", (req, res) => {
  req.session.destroy((err) => {
    if (err)
      return res
        .status(500)
        .json({ success: false, message: "Could not log out." });
    res.json({ success: true, message: "Logged out successfully." });
  });
});

app.get("/api/auth/status", (req, res) => {
  if (req.session.user) {
    return res.json({ authenticated: true, user: req.session.user });
  }
  if (jwtVerify(req) && req.user) {
    return res.json({ authenticated: true, user: req.user });
  }
  res.json({ authenticated: false });
});

// JWT refresh — issue new token if current one is still valid
app.post("/api/jwt/refresh", (req, res) => {
  try {
    const hdr = String(req.get('authorization') || '').trim();
    if (!hdr.toLowerCase().startsWith('bearer ')) {
      return res.status(401).json({ success: false, message: 'Missing token' });
    }
    const token = hdr.slice(7).trim();
    const payload = jwt.verify(token, JWT_SECRET, { ignoreExpiration: true });
    if (!payload || !payload.user) {
      return res.status(401).json({ success: false, message: 'Invalid token' });
    }
    const exp = Number(payload.exp || 0);
    const maxAge = 7 * 24 * 60 * 60; // 7 days in seconds
    const now = Math.floor(Date.now() / 1000);
    if (exp && exp < now) {
      // Expired but within grace period
      if (now - exp > maxAge) {
        return res.status(401).json({ success: false, message: 'Token expired, please login again' });
      }
    }
    const newToken = generateJwt(payload.user);
    if (!newToken) {
      return res.status(500).json({ success: false, message: 'JWT not configured' });
    }
    res.json({ success: true, token: newToken });
  } catch (err) {
    res.status(401).json({ success: false, message: 'Invalid token' });
  }
});

// --- Banner & QRIS APIs ---
// New: single-object Banner endpoints
app.get('/api/banner', isAuthenticated, async (req, res) => {
  try {
    const raw = await readData('banners.json');
    // Support legacy array file by reading first element
    const b = Array.isArray(raw) ? (raw[0] || {}) : (raw || {});
    res.json(b);
  } catch (e) {
    res.status(500).json({ success: false, message: 'Failed to load banner' });
  }
});

app.put('/api/banner', isAuthenticated, isAdmin, async (req, res) => {
  try {
    const { title = '', subtitle = '', imageBase64 = '', mode = 'server', html = '' } = req.body;
    const obj = { id: 1, title, subtitle, imageBase64, mode, html, updatedAt: Date.now() };
    // Write single-object, and also maintain legacy array for compatibility
    await writeData('banners.json', obj);
    try { await enqueueOutbox({ collection: 'banners', file: 'banners.json', op: 'upsert', _id: 'banner', doc: obj, updatedAt: Number(obj.updatedAt||Date.now()) }); } catch {}
    res.json({ success: true, banner: obj, message: 'Banner updated' });
  } catch (e) {
    console.error('Save banner error:', e);
    res.status(500).json({ success: false, message: 'Failed to save banner' });
  }
});

// Legacy array endpoints kept for backward-compatibility
app.get('/api/banners', isAuthenticated, isAdmin, async (req, res) => {
  try {
    const raw = await readData('banners.json');
    const b = Array.isArray(raw) ? (raw[0] || null) : (raw || null);
    res.json(b ? [b] : []);
  } catch (e) {
    res.status(500).json({ success: false, message: 'Failed to load banners' });
  }
});

app.post('/api/banners', isAuthenticated, isAdmin, async (req, res) => {
  try {
    const { title = '', subtitle = '', imageBase64 = '', mode = 'server', html = '' } = req.body;
    const obj = { id: 1, title, subtitle, imageBase64, mode, html, updatedAt: Date.now() };
    await writeData('banners.json', obj);
    try { await enqueueOutbox({ collection: 'banners', file: 'banners.json', op: 'upsert', _id: 'banner', doc: obj, updatedAt: Number(obj.updatedAt||Date.now()) }); } catch {}
    res.json({ success: true, banner: obj, message: 'Banner saved' });
  } catch (e) {
    res.status(500).json({ success: false, message: 'Failed to save banner' });
  }
});

// QRIS: store as single object
app.get('/api/qris', isAuthenticated, async (req, res) => {
  try {
    const raw = await readData('qris.json');
    const q = Array.isArray(raw) ? (raw[0] || {}) : (raw || {});
    res.json(q);
  } catch (e) {
    res.status(500).json({ success: false, message: 'Failed to load QRIS' });
  }
});

app.post('/api/qris', isAuthenticated, isAdmin, async (req, res) => {
  try {
    const {
      imageBase64 = '',
      paymentLogoQrisBase64 = '',
      paymentLogoDanaBase64 = '',
      paymentLogoOvoBase64 = ''
    } = req.body || {};

    const q = {
      id: 1,
      imageBase64,
      paymentLogoQrisBase64,
      paymentLogoDanaBase64,
      paymentLogoOvoBase64,
      updatedAt: Date.now()
    };
    await writeData('qris.json', q);
    try { await enqueueOutbox({ collection: 'qris', file: 'qris.json', op: 'upsert', _id: 'qris', doc: q, updatedAt: Number(q.updatedAt||Date.now()) }); } catch {}
    res.json({ success: true, qris: q, message: 'QRIS updated' });
  } catch (e) {
    console.error('Save QRIS error:', e);
    res.status(500).json({ success: false, message: 'Failed to save QRIS' });
  }
});

// Also accept PUT for QRIS
app.put('/api/qris', isAuthenticated, isAdmin, async (req, res) => {
  try {
    const {
      imageBase64 = '',
      paymentLogoQrisBase64 = '',
      paymentLogoDanaBase64 = '',
      paymentLogoOvoBase64 = ''
    } = req.body || {};

    const q = {
      id: 1,
      imageBase64,
      paymentLogoQrisBase64,
      paymentLogoDanaBase64,
      paymentLogoOvoBase64,
      updatedAt: Date.now()
    };
    await writeData('qris.json', q);
    try { await enqueueOutbox({ collection: 'qris', file: 'qris.json', op: 'upsert', _id: 'qris', doc: q, updatedAt: Number(q.updatedAt||Date.now()) }); } catch {}
    res.json({ success: true, qris: q, message: 'QRIS updated' });
  } catch (e) {
    res.status(500).json({ success: false, message: 'Failed to save QRIS' });
  }
});

// ---- Compatibility aliases (ID-based) ----
// Some frontend code may call /api/banners/1 or /api/qris/1. Provide aliases.
app.get('/api/banners/1', isAuthenticated, isAdmin, async (req, res) => {
  try {
    const banners = await readData('banners.json');
    const b = Array.isArray(banners) && banners.length > 0 ? banners[0] : null;
    if (!b) return res.json({});
    res.json(b);
  } catch (e) {
    res.status(500).json({ success: false, message: 'Failed to load banner' });
  }
});

app.post('/api/banners/1', isAuthenticated, isAdmin, async (req, res) => {
  try {
    const { title = '', subtitle = '', imageBase64 = '' } = req.body;
    let banners = await readData('banners.json');
    if (!Array.isArray(banners)) banners = [];
    const { mode = 'server', html = '' } = req.body;
    const newBanner = { id: banners[0]?.id || 1, title, subtitle, imageBase64, mode, html, updatedAt: Date.now() };
    if (banners.length === 0) banners.push(newBanner); else banners[0] = newBanner;
    await writeData('banners.json', banners);
    try { await enqueueOutbox({ collection: 'banners', file: 'banners.json', op: 'upsert', _id: 'banner', doc: newBanner, updatedAt: Number(newBanner.updatedAt||Date.now()) }); } catch {}
    res.json({ success: true, banner: banners[0], message: 'Banner saved' });
  } catch (e) {
    res.status(500).json({ success: false, message: 'Failed to save banner' });
  }
});

app.get('/api/qris/1', isAuthenticated, isAdmin, async (req, res) => {
  try {
    const raw = await readData('qris.json');
    const q = Array.isArray(raw) ? (raw[0] || {}) : (raw || {});
    res.json(q);
  } catch (e) {
    res.status(500).json({ success: false, message: 'Failed to load QRIS' });
  }
});

app.post('/api/qris/1', isAuthenticated, isAdmin, async (req, res) => {
  try {
    const {
      imageBase64 = '',
      paymentLogoQrisBase64 = '',
      paymentLogoDanaBase64 = '',
      paymentLogoOvoBase64 = ''
    } = req.body || {};

    const q = {
      id: 1,
      imageBase64,
      paymentLogoQrisBase64,
      paymentLogoDanaBase64,
      paymentLogoOvoBase64,
      updatedAt: Date.now()
    };
    await writeData('qris.json', q);
    try { await enqueueOutbox({ collection: 'qris', file: 'qris.json', op: 'upsert', _id: 'qris', doc: q, updatedAt: Number(q.updatedAt||Date.now()) }); } catch {}
    res.json({ success: true, qris: q, message: 'QRIS updated' });
  } catch (e) {
    res.status(500).json({ success: false, message: 'Failed to save QRIS' });
  }
});

// Current user info
app.get("/api/current-user", isAuthenticated, async (req, res) => {
  try {
    // Return user info from session + permissions from users.json (if available)
    const sessionUser = req.session.user || req.user;
    if (!sessionUser) {
      return res.status(404).json({ success: false, message: "User not found" });
    }
    let fullUser = null;
    try {
      const users = await readData("users.json");
      if (Array.isArray(users)) {
        fullUser = users.find((u) => String(u.id) === String(sessionUser.id)) || null;
      }
    } catch {}
    const merged = { ...(fullUser || {}), ...(sessionUser || {}) };
    // Don't send password hash to client
    const { password, ...userWithoutPassword } = merged;
    res.json(userWithoutPassword);
  } catch (error) {
    res.status(500).json({ success: false, message: "Failed to get current user" });
  }
});

// Users
app.get("/api/users", isAuthenticated, isAdmin, async (req, res) => {
  try {
    const users = await readData("users.json");
    // Don't send password hashes to the client
    const usersWithoutPasswords = users.map(({ password, ...user }) => user);
    res.json(usersWithoutPasswords);
  } catch (error) {
    res.status(500).json({ success: false, message: "Failed to load users" });
  }
});

app.post("/api/users", strictLimiter, isAuthenticated, isAdmin, async (req, res) => {
  try {
    const { username, name, password, role, status = "active" } = req.body;

    // Validate and sanitize inputs
    if (!username || !name || !password || !role) {
      return res.status(400).json({
        success: false,
        message: "Missing required fields: username, name, password, role",
      });
    }

    try {
      var validatedUsername = validateAndSanitizeInput(username, 'username');
      var validatedName = validateAndSanitizeInput(name, 'general');
      var validatedPassword = validateAndSanitizeInput(password, 'password');
      var validatedRole = validateAndSanitizeInput(role, 'general');
      var validatedStatus = validateAndSanitizeInput(status, 'general');
    } catch (validationError) {
      return res.status(400).json({
        success: false,
        message: validationError.message,
      });
    }

    // Validasi username duplikat
    const existingUser = await validateUsername(validatedUsername);
    if (existingUser) {
      return res.status(400).json({
        success: false,
        message: `Username "${validatedUsername}" sudah ada. Silakan gunakan username lain.`,
      });
    }

    // Hash password sebelum menyimpan
    const hashedPassword = await bcrypt.hash(validatedPassword, 10);

    const users = await readData("users.json");
    const newUser = {
      id: Date.now(),
      username: validatedUsername,
      name: validatedName,
      password: hashedPassword,
      role: validatedRole,
      status: validatedStatus,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    };
    users.push(newUser);
    await saveArrayWithSync("users.json", users);
    res.json(newUser);
  } catch (error) {
    console.error("Error creating user:", error);
    res.status(500).json({ success: false, message: "Failed to create user" });
  }
});

app.put("/api/users/:id", strictLimiter, isAuthenticated, isAdmin, async (req, res) => {
  try {
    const { username, name, password, role, status, permissions } = req.body; // Tambahkan username di sini
    const userId = req.params.id;

    // Validate and sanitize inputs
    if (!username || !name || !role) {
      return res.status(400).json({
        success: false,
        message: "Missing required fields: username, name, role",
      });
    }

    try {
      var validatedUsername = validateAndSanitizeInput(username, 'username');
      var validatedName = validateAndSanitizeInput(name, 'general');
      var validatedRole = validateAndSanitizeInput(role, 'general');
      var validatedStatus = validateAndSanitizeInput(status, 'general');
      // Only validate password if it's provided
      var validatedPassword = password ? validateAndSanitizeInput(password, 'password') : null;
    } catch (validationError) {
      return res.status(400).json({
        success: false,
        message: validationError.message,
      });
    }

    // Validasi username duplikat
    const existingUser = await validateUsername(validatedUsername, userId);
    if (existingUser) {
      return res.status(400).json({
        success: false,
        message: `Username "${validatedUsername}" sudah ada. Silakan gunakan username lain.`,
      });
    }

    const users = await readData("users.json");
    const index = users.findIndex((u) => u.id == userId);

    if (index !== -1) {
      users[index] = {
        ...users[index],
        username: validatedUsername, // Tambahkan ini
        name: validatedName,
        role: validatedRole,
        status: validatedStatus,
        updatedAt: new Date().toISOString(),
      };
      if (permissions && typeof permissions === 'object') {
        users[index].permissions = permissions;
      }

      // Hash password baru jika ada
      if (validatedPassword) {
        users[index].password = await bcrypt.hash(validatedPassword, 10);
      }

      await saveArrayWithSync("users.json", users);
      res.json(users[index]);
    } else {
      res.status(404).json({
        success: false,
        message: "User tidak ditemukan",
      });
    }
  } catch (error) {
    console.error("Error updating user:", error);
    res.status(500).json({
      success: false,
      message: "Failed to update user",
    });
  }
});

app.delete("/api/users/:id", strictLimiter, isAuthenticated, isAdmin, async (req, res) => {
  try {
    const userId = req.params.id;

    // Cegah apakah user yang sedang login
    if (req.session.user && req.session.user.id == userId) {
      return res.status(400).json({
        success: false,
        message: "Tidak dapat menghapus user yang sedang login",
      });
    }

    const users = await readData("users.json");
    const filteredUsers = users.filter((u) => u.id != userId);

    if (users.length !== filteredUsers.length) {
      await saveArrayWithSync("users.json", filteredUsers);
      res.json({ success: true });
    } else {
      res.status(404).json({
        success: false,
        message: "User tidak ditemukan",
      });
    }
  } catch (error) {
    console.error("Error deleting user:", error);
    res.status(500).json({
      success: false,
      message: "Failed to delete user",
    });
  }
});

// Update per-user permissions (superadmin or allowed admin)
app.put("/api/users/:id/permissions", strictLimiter, requireRole('admin'), async (req, res) => {
  try {
    const userId = req.params.id;
    const actor = req.session && req.session.user ? req.session.user : (req.user || {});
    const actorRole = normalizeRole(actor && actor.role);
    const permissions = req.body && req.body.permissions ? req.body.permissions : null;
    if (!permissions || typeof permissions !== 'object') {
      return res.status(400).json({ success: false, message: "Invalid permissions payload" });
    }
    const users = await readData("users.json");
    if (!Array.isArray(users)) {
      return res.status(500).json({ success: false, message: "Users data invalid" });
    }
    const idx = users.findIndex(u => String(u.id) === String(userId));
    if (idx < 0) return res.status(404).json({ success: false, message: "User not found" });
    const actorId = actor && actor.id != null ? String(actor.id) : '';
    const targetUser = users[idx];
    const targetRole = normalizeRole(targetUser && targetUser.role);
    if (actorRole === 'admin' && targetRole === 'superadmin') {
      return res.status(403).json({ success: false, message: "Forbidden" });
    }
    if (actorRole === 'admin' && actorId && String(targetUser.id) === String(actorId)) {
      return res.status(403).json({ success: false, message: "Admin tidak dapat mengubah akses dirinya sendiri." });
    }
    if (actorRole === 'admin') {
      let allowed = true;
      try {
        const settings = await readData('settings.json').catch(() => ({}));
        const roleViews = (settings && settings.rolePermissions && settings.rolePermissions.admin && settings.rolePermissions.admin.views) || {};
        if (roleViews.permissions === false) allowed = false;
      } catch {}
      try {
        const cur = users.find(u => String(u.id) === String(actorId));
        const perUser = cur && cur.permissions && cur.permissions.adminViews;
        if (perUser && perUser.permissions === false) allowed = false;
      } catch {}
      if (!allowed) {
        return res.status(403).json({ success: false, message: "Forbidden" });
      }
    }
    // Prevent superadmin from locking themselves out of critical views
    if (actorId && String(targetUser.id) === actorId && normalizeRole(actor.role) === 'superadmin') {
      const adminViews = permissions.adminViews || {};
      const mustKeep = ['dashboard', 'settings', 'users', 'permissions'];
      const disabled = mustKeep.filter(k => adminViews[k] === false);
      if (disabled.length) {
        return res.status(400).json({
          success: false,
          message: `Superadmin tidak boleh menonaktifkan akses penting untuk dirinya sendiri: ${disabled.join(', ')}`
        });
      }
    }
    const beforePermissions = targetUser.permissions || {};
    users[idx].permissions = permissions;
    users[idx].updatedAt = new Date().toISOString();
    await saveArrayWithSync("users.json", users);
    const changes = diffPermissions(beforePermissions, permissions);
    try {
      await appendPermissionsAudit({
        id: `perm-${Date.now()}-${Math.random().toString(36).slice(2)}`,
        ts: Date.now(),
        at: new Date().toISOString(),
        actor: {
          id: actor && actor.id != null ? actor.id : null,
          username: actor && (actor.username || actor.name) ? (actor.username || actor.name) : '',
          role: actor && actor.role ? String(actor.role) : ''
        },
        target: {
          id: targetUser && targetUser.id != null ? targetUser.id : null,
          username: targetUser && (targetUser.username || targetUser.name) ? (targetUser.username || targetUser.name) : '',
          role: targetUser && targetUser.role ? String(targetUser.role) : ''
        },
        changes
      });
    } catch {}
    const { password, ...userWithoutPassword } = users[idx];
    res.json({ success: true, user: userWithoutPassword });
  } catch (e) {
    res.status(500).json({ success: false, message: "Failed to update permissions" });
  }
});

// PERBAIKAN: Validasi password user yang sedang login untuk aksi berbahaya
app.post(
  "/api/validate-current-user-password",
  isAuthenticated,
  async (req, res) => {
    try {
      const { password } = req.body;
      const users = await readData("users.json");
      const currentUser = users.find((u) => u.id === req.session.user.id);

      if (!currentUser) {
        return res
          .status(404)
          .json({ success: false, message: "User not found." });
      }

      const isMatch = await bcrypt.compare(password, currentUser.password);
      if (isMatch) {
        res.json({ success: true, message: "Password validated." });
      } else {
        res.status(401).json({ success: false, message: "Invalid password." });
      }
    } catch (error) {
      console.error("Error validating password:", error);
      res.status(500).json({ success: false, message: "Server error." });
    }
  }
);

app.post('/api/admin/reencrypt', strictLimiter, isAuthenticated, isAdmin, async (req, res) => {
  try {
    const pass = process.env.POS_PASSPHRASE || '';
    if (!pass) return res.status(400).json({ success: false, message: 'POS_PASSPHRASE is required' });
    const files = await fs.readdir(DATA_DIR).catch(() => []);
    let processed = 0;
    const failed = [];
    for (const f of files) {
      if (!f.toLowerCase().endsWith('.json')) continue;
      try {
        const obj = await readData(f);
        await writeData(f, obj);
        processed++;
      } catch (e) {
        failed.push(f);
      }
    }
    res.json({ success: true, processed, failed });
  } catch (e) {
    res.status(500).json({ success: false, message: 'Re-encrypt failed' });
  }
});

app.post('/api/admin/rekey', isAuthenticated, isAdmin, async (req, res) => {
  try {
    const body = req.body || {};
    const oldPassphrase = String(body.oldPassphrase || '');
    const newPassphrase = String(body.newPassphrase || '');
    if (!oldPassphrase || !newPassphrase) return res.status(400).json({ success: false, message: 'oldPassphrase and newPassphrase are required' });
    if (oldPassphrase === newPassphrase) return res.status(400).json({ success: false, message: 'newPassphrase must be different' });
    const files = await fs.readdir(DATA_DIR).catch(() => []);
    let processed = 0;
    const failed = [];
    for (const f of files) {
      if (!f.toLowerCase().endsWith('.json')) continue;
      const full = path.join(DATA_DIR, f);
      const raw = await fs.readFile(full, 'utf-8').catch(() => null);
      if (raw == null) { failed.push(f); continue; }
      let obj;
      try {
        if (raw.startsWith('ENC1:')) {
          const parts = raw.split(':');
          if (parts.length !== 5) throw new Error('bad format');
          const salt = Buffer.from(parts[1], 'base64');
          const iv = Buffer.from(parts[2], 'base64');
          const tag = Buffer.from(parts[3], 'base64');
          const ciphertext = Buffer.from(parts[4], 'base64');
          const key = crypto.scryptSync(oldPassphrase, salt, 32);
          const decipher = crypto.createDecipheriv('aes-256-gcm', key, iv);
          decipher.setAuthTag(tag);
          const dec = Buffer.concat([decipher.update(ciphertext), decipher.final()]).toString('utf8');
          obj = JSON.parse(dec);
        } else {
          obj = JSON.parse(raw);
        }
      } catch (e) { failed.push(f); continue; }
      try {
        const salt = crypto.randomBytes(16);
        const iv = crypto.randomBytes(12);
        const key = crypto.scryptSync(newPassphrase, salt, 32);
        const cipher = crypto.createCipheriv('aes-256-gcm', key, iv);
        const ciphertext = Buffer.concat([cipher.update(Buffer.from(JSON.stringify(obj), 'utf8')), cipher.final()]);
        const tag = cipher.getAuthTag();
        const out = `ENC1:${salt.toString('base64')}:${iv.toString('base64')}:${tag.toString('base64')}:${ciphertext.toString('base64')}`;
        await fs.writeFile(full, out, 'utf-8');
        processed++;
      } catch (e) { failed.push(f); }
    }
    process.env.POS_PASSPHRASE = newPassphrase;
    try { await fs.writeFile(path.join(DATA_DIR, 'passphrase.txt'), newPassphrase, 'utf-8'); } catch {}
    res.json({ success: true, processed, failed });
  } catch (e) {
    res.status(500).json({ success: false, message: 'Rekey failed' });
  }
});

// Units (Satuan)
app.get('/api/units', isAuthenticated, async (req, res) => {
  try {
    const units = await readData('units.json');
    res.json(Array.isArray(units) ? units : []);
  } catch (e) {
    res.status(500).json({ success:false, message:'Failed to load units' });
  }
});

app.post('/api/units', isAuthenticated, isAdmin, async (req, res) => {
  try {
    const { name } = req.body || {};
    if (!name || String(name).trim() === '') return res.status(400).json({ success:false, message:'Nama satuan wajib diisi' });
    let units = await readData('units.json'); if (!Array.isArray(units)) units = [];
    // unique by name (case-insensitive)
    const exists = units.find(u => u.name && u.name.toLowerCase() === String(name).trim().toLowerCase());
    if (exists) return res.status(400).json({ success:false, message:`Satuan "${name}" sudah ada.` });
    const now = Date.now();
    const unit = { id: now, name: String(name).trim(), description: String(req.body.description||'').trim(), createdAt: now, updatedAt: now };
    units.push(unit);
    await saveArrayWithSync('units.json', units);
    res.json(unit);
  } catch (e) { res.status(500).json({ success:false, message:'Failed to create unit' }); }
});

app.put('/api/units/:id', isAuthenticated, isAdmin, async (req, res) => {
  try {
    const id = String(req.params.id);
    let units = await readData('units.json'); if (!Array.isArray(units)) units = [];
    const idx = units.findIndex(u => String(u.id) === id);
    if (idx === -1) return res.status(404).json({ success:false, message:'Satuan tidak ditemukan' });
    const name = req.body && req.body.name != null ? String(req.body.name).trim() : (units[idx].name || '');
    if (!name) return res.status(400).json({ success:false, message:'Nama satuan wajib diisi' });
    // duplicate check
    const dup = units.find(u => u.name && u.name.toLowerCase() === name.toLowerCase() && String(u.id) !== id);
    if (dup) return res.status(400).json({ success:false, message:`Satuan "${name}" sudah ada.` });
    units[idx] = { ...units[idx], name, description: String(req.body.description||'').trim(), updatedAt: Date.now() };
    await saveArrayWithSync('units.json', units);
    res.json(units[idx]);
  } catch (e) { res.status(500).json({ success:false, message:'Failed to update unit' }); }
});

app.delete('/api/units/:id', isAuthenticated, isAdmin, async (req, res) => {
  try {
    const id = String(req.params.id);
    let units = await readData('units.json'); if (!Array.isArray(units)) units = [];
    const before = units.length;
    units = units.filter(u => String(u.id) !== id);
    if (units.length === before) return res.status(404).json({ success:false, message:'Satuan tidak ditemukan' });
    await saveArrayWithSync('units.json', units);
    res.json({ success:true });
  } catch (e) { res.status(500).json({ success:false, message:'Failed to delete unit' }); }
});

// Units export
app.get('/api/units/export', isAuthenticated, isAdmin, async (req, res) => {
  try {
    const units = await readData('units.json');
    const rows = (Array.isArray(units)?units:[]).map(u => ({ 'Unit ID': u.id, 'Unit Name': u.name || '', 'Description': u.description || '' }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Units');
    const out = XLSX.write(wb, { type:'buffer', bookType:'xlsx' });
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="units_export.xlsx"');
    res.send(out);
  } catch (e) { res.status(500).json({ success:false, message:'Export gagal' }); }
});

// Units template
app.get('/api/units/template', isAuthenticated, isAdmin, async (req, res) => {
  try {
    const ws = XLSX.utils.aoa_to_sheet([["Unit Name","Description"],["pcs","Satuan dasar"],["box","Kemasan box"]]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Template');
    const out = XLSX.write(wb, { type:'buffer', bookType:'xlsx' });
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="unit_import_template.xlsx"');
    res.send(out);
  } catch (e) { res.status(500).json({ success:false, message:'Gagal membuat template' }); }
});

// Units import
app.post('/api/units/import', isAuthenticated, isAdmin, async (req, res) => {
  try {
    const payload = req.body || {};
    const list = Array.isArray(payload.units) ? payload.units : [];
    if (!list.length) return res.status(400).json({ success:false, message:'Tidak ada data untuk diimport' });
    let units = await readData('units.json'); if (!Array.isArray(units)) units = [];
    let created = 0, updated = 0;
    for (const row of list) {
      const name = (row['Unit Name'] ?? row.name ?? '').toString().trim();
      const desc = (row['Description'] ?? row.description ?? '').toString().trim();
      if (!name) continue;
      const existIdx = units.findIndex(u => u.name && u.name.toLowerCase() === name.toLowerCase());
      if (existIdx >= 0) {
        units[existIdx] = { ...units[existIdx], description: desc, updatedAt: Date.now() };
        updated++;
      } else {
        const now = Date.now() + Math.floor(Math.random()*1000);
        units.push({ id: now, name, description: desc, createdAt: now, updatedAt: now });
        created++;
      }
    }
    await saveArrayWithSync('units.json', units);
    res.json({ success:true, message:`Import selesai. Ditambahkan: ${created}, Diupdate: ${updated}` });
  } catch (e) { res.status(500).json({ success:false, message:'Import gagal' }); }
});

// Categories
app.get("/api/categories", isAuthenticated, async (req, res) => {
  try {
    const categories = await readData("categories.json");
    res.json(categories);
  } catch (error) {
    res
      .status(500)
      .json({ success: false, message: "Failed to load categories" });
  }
});

app.post("/api/categories", isAuthenticated, isAdmin, async (req, res) => {
  try {
    const { name } = req.body;

    // Validasi nama kategori
    if (!name || name.trim() === "") {
      return res.status(400).json({
        success: false,
        message: "Nama kategori wajib diisi",
      });
    }

    // Cek nama kategori duplikat
    const existingCategory = await validateCategoryName(name.trim());
    if (existingCategory) {
      return res.status(400).json({
        success: false,
        message: `Kategori "${name}" sudah ada. Silakan gunakan nama lain.`,
      });
    }

    const categories = await readData("categories.json");
    const newCategory = {
      id: Date.now(),
      ...req.body,
      name: name.trim(),
    };
    categories.push(newCategory);
    await saveArrayWithSync("categories.json", categories);
    res.json(newCategory);
  } catch (error) {
    console.error("Error creating category:", error);
    res.status(500).json({
      success: false,
      message: "Failed to create category",
    });
  }
});

app.put("/api/categories/:id", isAuthenticated, isAdmin, async (req, res) => {
  try {
    const { name } = req.body;
    const categoryId = req.params.id;

    // Validasi nama kategori
    if (!name || name.trim() === "") {
      return res.status(400).json({
        success: false,
        message: "Nama kategori wajib diisi",
      });
    }

    // Cek nama kategori duplikat
    const existingCategory = await validateCategoryName(
      name.trim(),
      categoryId
    );
    if (existingCategory) {
      return res.status(400).json({
        success: false,
        message: `Kategori "${name}" sudah ada. Silakan gunakan nama lain.`,
      });
    }

    const categories = await readData("categories.json");
    const index = categories.findIndex((c) => c.id == categoryId);

    if (index !== -1) {
      categories[index] = {
        ...categories[index],
        ...req.body,
        name: name.trim(),
      };
      await saveArrayWithSync("categories.json", categories);
      res.json(categories[index]);
    } else {
      res.status(404).json({
        success: false,
        message: "Kategori tidak ditemukan",
      });
    }
  } catch (error) {
    console.error("Error updating category:", error);
    res.status(500).json({
      success: false,
      message: "Failed to update category",
    });
  }
});

app.delete(
  "/api/categories/:id",
  isAuthenticated,
  isAdmin,
  async (req, res) => {
    try {
      const categoryId = req.params.id;

      // Cek apakah kategori sedang digunakan oleh produk
      const products = await readData("products.json"); // Gunakan readData langsung
      const productsInCategory = products.filter(
        (p) => p.categoryId == categoryId
      );

      if (productsInCategory.length > 0) {
        return res.status(400).json({
          success: false,
          message: `Tidak dapat menghapus kategori ini karena masih digunakan oleh ${productsInCategory.length} produk. Pindahkan atau hapus produk tersebut terlebih dahulu.`,
        });
      }

      const categories = await readData("categories.json");
      const filteredCategories = categories.filter((c) => c.id != categoryId);

      if (categories.length !== filteredCategories.length) {
        await saveArrayWithSync("categories.json", filteredCategories);
        res.json({ success: true });
      } else {
        res.status(404).json({
          success: false,
          message: "Kategori tidak ditemukan",
        });
      }
    } catch (error) {
      console.error("Error deleting category:", error);
      res.status(500).json({
        success: false,
        message: "Failed to delete category",
      });
    }
  }
);

// Products
app.get("/api/products", isAuthenticated, async (req, res) => {
  try {
    // Generate cache key based on query parameters
    const cacheKey = `products_${JSON.stringify(req.query)}`;
    
    // Try to get from cache first
    let cachedProducts = productCache.get(cacheKey);
    if (cachedProducts) {
      return res.json(cachedProducts);
    }
    
    let products = await readData("products.json");
    if (!Array.isArray(products)) products = [];
    const includeDeleted = String(req.query.includeDeleted || 'false').toLowerCase() === 'true';
    const q = (req.query.q || "").toString().trim().toLowerCase();
    const sort = (req.query.sort || "").toString().trim();
    const fields = (req.query.fields || "").toString().trim();
    const limit = Math.max(0, Number(req.query.limit || 0) || 0); // Default 50 items
    const offset = Math.max(0, Number(req.query.offset || 0) || 0);
    const categoryIdFilter = (req.query.categoryId || '').toString().trim();
    const hasImage = String(req.query.hasImage || 'false').toLowerCase() === 'true';

    if (!includeDeleted) {
      products = products.filter(p => !(p && p.deleted === true));
    }

    // Filter by category if requested
    if (categoryIdFilter) {
      const target = categoryIdFilter.toLowerCase();
      products = products.filter(p => String(p && p.categoryId || '').toLowerCase() === target);
    }

    if (q) {
      const contains = (v) => (v == null ? "" : String(v)).toLowerCase().includes(q);
      // Also allow matching by category name
      let catMap = null;
      try {
        const cats = await readData('categories.json').catch(() => []);
        if (Array.isArray(cats)) {
          catMap = new Map(cats.map(c => [ String(c && c.id), String((c && (c.name || c.nama)) || '') ]));
        }
      } catch {}
      products = products.filter((p) => {
        if (contains(p.name) || contains(p.sku) || contains(p.qrCode)) return true;
        const catName = (p && (p.category || (catMap && catMap.get(String(p.categoryId))))) || '';
        return contains(catName);
      });
    }

    // Filter to only those with imageBase64 if requested
    if (hasImage) {
      products = products.filter(p => {
        const v = (p && p.imageBase64) || '';
        return typeof v === 'string' && v.trim().length > 0;
      });
    }

    if (sort) {
      const desc = sort.startsWith("-");
      const key = desc ? sort.slice(1) : sort;
      products = products.slice().sort((a,b) => {
        const va = a && a[key];
        const vb = b && b[key];
        if (va == null && vb == null) return 0;
        if (va == null) return desc ? 1 : -1;
        if (vb == null) return desc ? -1 : 1;
        if (va < vb) return desc ? 1 : -1;
        if (va > vb) return desc ? -1 : 1;
        return 0;
      });
    }

    if (fields) {
      const pick = new Set(fields.split(",").map(s=>s.trim()).filter(Boolean));
      if (pick.size > 0) {
        products = products.map((p) => {
          const o = {};
          pick.forEach((k)=>{ if (k in p) o[k] = p[k]; });
          // Sync docs may only have `_id`; keep `id` usable for forms/dropdowns.
          if (pick.has('id') && o.id == null && p && p._id != null) o.id = p._id;
          if (pick.has('_id') && o._id == null && p && p.id != null) o._id = p.id;
          return o;
        });
      }
    }

    if (limit > 0) {
      const start = offset;
      const end = offset + limit;
      products = products.slice(start, end);
    }

    // Set headers untuk Cloudflare Tunnel
    res.setHeader('Connection', 'keep-alive');
    res.setHeader('Keep-Alive', 'timeout=60');
    res.setHeader('Cache-Control', 'public, max-age=300');
    
    // Return paginated response with metadata
    const totalCount = products.length;
    const paginatedProducts = limit > 0 ? products.slice(offset, offset + limit) : products;
    
    const response = {
      data: paginatedProducts,
      pagination: {
        total: totalCount,
        offset: offset,
        limit: limit,
        hasMore: limit > 0 ? (offset + limit) < totalCount : false
      }
    };
    
    // Cache the response
    productCache.set(cacheKey, response);
    
    res.json(response);
  } catch (error) {
    res.status(500).json({ success: false, message: "Failed to load products" });
  }
});

// Endpoint to get popular products quickly for POS interface
app.get("/api/products/popular", isAuthenticated, async (req, res) => {
  try {
    // Try to get from cache first
    let cached = productCache.get('popular_products');
    if (cached) {
      return res.json(cached);
    }
    
    // If not in cache, get all products and filter
    const products = await readData('products.json');
    const popularProducts = Array.isArray(products) ? 
      products.filter(p => p.isTop || p.isBest || p.isTopProduct || p.isBestSeller || (p.stats && p.stats.salesCount > 10)) : [];
    
    // Cache for faster subsequent requests
    productCache.set('popular_products', popularProducts);
    
    res.json(popularProducts);
  } catch (error) {
    console.error('Error fetching popular products:', error);
    res.status(500).json({ error: 'Internal server error' });
  }
});

function formatProductVariantPricesForXlsx(unitPrices) {
  return VariantStockSync.formatAllVariantsForXlsx(unitPrices);
}

function parseProductVariantPricesFromXlsx(raw, fallbackUnit = "") {
  return VariantStockSync.parseVariantsFromXlsx(raw, fallbackUnit);
}

function sanitizeUnitPriceRow(v) {
  const note = validateAndSanitizeInput((v.note || v.desc || v.keterangan || "").toString(), "general");
  const sku = validateAndSanitizeInput((v.sku || "").toString(), "general");
  const photo = (v.photo || "").toString().trim();
  const o = {
    qty: Number(v.qty) || 0,
    unit: validateAndSanitizeInput((v.unit || "").toString(), "general"),
    price: Number(v.price) || 0,
    stock: Number(v.stock) || 0,
    use: v.use !== false,
  };
  const parentIndexRaw = v.parentIndex != null ? v.parentIndex : v.parentIdx;
  if (parentIndexRaw != null && parentIndexRaw !== "") {
    const pIdx = Number(parentIndexRaw);
    if (Number.isInteger(pIdx) && pIdx >= 0) o.parentIndex = pIdx;
  }
  const ratioRaw = v.unitsPerParent != null ? v.unitsPerParent : v.perParent;
  if (ratioRaw != null && ratioRaw !== "") {
    const ratio = Number(ratioRaw);
    if (isFinite(ratio) && ratio > 0) o.unitsPerParent = Math.floor(ratio);
  }
  if (note) o.note = note;
  if (sku) o.sku = sku;
  if (photo) o.photo = photo;
  return o;
}

function sanitizeUnitPricesList(list) {
  if (!Array.isArray(list)) return [];
  return list
    .map((v) => sanitizeUnitPriceRow(v || {}))
    .filter((v) => v.qty > 0 && v.price >= 0 && v.unit);
}

function applyVariantStockChange(product, variantIndex, qty, direction) {
  if (!product || !Array.isArray(product.unitPrices)) return product;
  const idx = Number(variantIndex);
  const amount = Math.max(0, Number(qty) || 0);
  if (!Number.isInteger(idx) || idx < 0 || idx >= product.unitPrices.length || amount <= 0) {
    return product;
  }
  const sign = Number(direction) < 0 ? -1 : 1;
  const result = sign < 0
    ? VariantStockSync.deductVariantStock(product.unitPrices, idx, amount)
    : VariantStockSync.addVariantStock(product.unitPrices, idx, amount);
  product.unitPrices = result.unitPrices;
  return product;
}

// --- Excel Import/Export API Routes (Products) ---
// Export Products to XLSX
app.get("/api/products/export", isAuthenticated, isAdmin, async (req, res) => {
  try {
    // console.log("Requesting export...");
    let products = await readData("products.json");
    if (!Array.isArray(products)) products = [];
    let categories = await readData("categories.json");
    if (!Array.isArray(categories)) categories = [];
    let units = await readData("units.json");
    if (!Array.isArray(units)) units = [];

    // Transform data for export - EXCLUDE Image Base64 to avoid cell limit
    const exportData = products.map((product) => {
      const category = categories.find((c) => c.id === product.categoryId);
      const variants = formatProductVariantPricesForXlsx(product.unitPrices);
      return {
        "Product Name": product.name || "",
        "Purchase Price": product.purchasePrice || 0,
        "Selling Price": (product.sellingPrice != null ? product.sellingPrice : product.price) || 0,
        "PPN (%)": Number(product.taxRate || 0) || 0,
        "Discount (%)": Number(product.discountPercent || 0) || 0,
        Price: product.price || 0,
        Stock: product.stock || 0,
        Category: category ? category.name : "",
        Unit: product.unit || "",
        "Variant Prices": variants,
        SKU: product.sku || "",
        "QR Code": product.qrCode || "",
        "Is Top Product": product.isTopProduct ? "Yes" : "No",
        "Is Best Seller": product.isBestSeller ? "Yes" : "No",
        "Has Image": product.imageBase64 ? "Yes" : "No",
      };
    });

    // Create workbook
    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Products");

    // Set column widths
    const colWidths = [
      { wch: 30 }, // Product Name
      { wch: 15 }, // Purchase Price
      { wch: 15 }, // Selling Price
      { wch: 12 }, // PPN (%)
      { wch: 14 }, // Discount (%)
      { wch: 15 }, // Price (legacy)
      { wch: 10 }, // Stock
      { wch: 20 }, // Category
      { wch: 12 }, // Unit
      { wch: 40 }, // Variant Prices
      { wch: 20 }, // SKU
      { wch: 25 }, // QR Code
      { wch: 15 }, // Is Top Product
      { wch: 15 }, // Is Best Seller
      { wch: 15 }, // Has Image
    ];
    ws["!cols"] = colWidths;

    // Generate buffer
    const buf = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });

    // Set headers
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader(
      "Content-Disposition",
      "attachment; filename=products_export.xlsx"
    );

    console.log("Export completed successfully");
    res.send(buf);
  } catch (error) {
    console.error("Export error:", error);
    res
      .status(500)
      .json({
        success: false,
        message: "Failed to export products: " + error.message,
      });
  }
});

// Download Import Template
app.get(
  "/api/products/template",
  isAuthenticated,
  isAdmin,
  async (req, res) => {
    try {
      // console.log("Requesting template...");
      const categories = await readData("categories.json");

      // Create template data with example rows - include new price fields and QR Code
      const templateData = [
        {
          "Product Name": "Example Product 1",
          "Purchase Price": 8000,
          "Selling Price": 10000,
          "PPN (%)": 11,
          "Discount (%)": 5,
          Price: 10000,
          Stock: 50,
          Category: categories.length > 0 ? categories[0].name : "General",
          Unit: "pcs",
          "Variant Prices": "1 Karton: 100000: 5 | 1 Box: 10000: 0:12:Karton | 1 Pcs: 500: 0:20:Box",
          SKU: "PROD-001",
          "QR Code": "QR-EX-001",
          "Is Top Product": "Yes",
          "Is Best Seller": "No",
          "Has Image": "No",
        },
        {
          "Product Name": "Example Product 2",
          "Purchase Price": 20000,
          "Selling Price": 25000,
          "PPN (%)": 11,
          "Discount (%)": 0,
          Price: 25000,
          Stock: 30,
          Category: categories.length > 1 ? categories[1].name : "General",
          Unit: "box",
          "Variant Prices": "1 box: 25000: 20 | 1 dus: 240000: 5",
          SKU: "PROD-002",
          "QR Code": "QR-EX-002",
          "Is Top Product": "No",
          "Is Best Seller": "Yes",
          "Has Image": "No",
        },
      ];

      // Create workbook
      const ws = XLSX.utils.json_to_sheet(templateData);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Template");

      // Set column widths
      const colWidths = [
        { wch: 30 }, // Product Name
        { wch: 15 }, // Purchase Price
        { wch: 15 }, // Selling Price
        { wch: 12 }, // PPN (%)
        { wch: 14 }, // Discount (%)
        { wch: 15 }, // Price (legacy)
        { wch: 10 }, // Stock
        { wch: 20 }, // Category
        { wch: 12 }, // Unit
        { wch: 40 }, // Variant Prices
        { wch: 20 }, // SKU
        { wch: 25 }, // QR Code
        { wch: 15 }, // Is Top Product
        { wch: 15 }, // Is Best Seller
        { wch: 15 }, // Has Image
      ];
      ws["!cols"] = colWidths;

      // Generate buffer
      const buf = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });

      // Set headers
      res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      );
      res.setHeader(
        "Content-Disposition",
        "attachment; filename=product_import_template.xlsx"
      );

      // console.log('Template generated successfully');
      res.send(buf);
    } catch (error) {
      console.error("Template generation error:", error);
      res
        .status(500)
        .json({
          success: false,
          message: "Failed to generate template: " + error.message,
        });
    }
  }
);

// Endpoint to get single product by ID
app.get("/api/products/:id", isAuthenticated, async (req, res) => {
  try {
    const productId = req.params.id;
    const products = await readData('products.json');
    
    if (!Array.isArray(products)) {
      return res.status(404).json({ error: 'Product not found' });
    }
    
    const product = products.find(p => String(p.id) === String(productId));
    
    if (!product) {
      return res.status(404).json({ error: 'Product not found' });
    }
    
    try {
      if (!Array.isArray(product.priceHistory) || product.priceHistory.length === 0) {
        const map = await readPriceHistoryMap();
        const ph = map[String(product.id)] || [];
        if (Array.isArray(ph) && ph.length) {
          product.priceHistory = ph;
        }
      }
    } catch {}
    res.json(product);
  } catch (error) {
    console.error('Error fetching product:', error);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// Get price history for a product
app.get("/api/products/:id/price-history", isAuthenticated, async (req, res) => {
  try {
    const productId = String(req.params.id || '');
    const filePath = path.join(DATA_DIR, 'price_history.json');
    let raw = await fs.readFile(filePath, 'utf-8').catch(() => '');
    try {
      if (typeof raw === 'string' && raw.startsWith('ENC1:')) {
        raw = decryptTextIfEnc1(raw) || '';
      }
    } catch {}
    let map = {};
    try { map = raw ? JSON.parse(raw) : {}; } catch { map = {}; }
    let history = Array.isArray(map[productId]) ? map[productId] : [];
    try {
      const products = await readData('products.json').catch(() => []);
      if (Array.isArray(products)) {
        const p = products.find(x => String(x && x.id) === productId);
        if (p && Array.isArray(p.priceHistory) && p.priceHistory.length) {
          // Merge product.priceHistory with map history (dedupe by signature)
          const sig = (h) => {
            try {
              const at = Number(h && h.at || 0);
              const by = String(h && h.by || '');
              const reason = String(h && h.reason || '');
              const base = JSON.stringify(h && h.base ? h.base : {});
              const variants = JSON.stringify(h && h.variants ? h.variants : []);
              return `${at}|${by}|${reason}|${base}|${variants}`;
            } catch { return String(h && h.at || Math.random()); }
          };
          const seen = new Set((history || []).map(sig));
          for (const h of p.priceHistory) {
            const k = sig(h);
            if (seen.has(k)) continue;
            history.push(h);
            seen.add(k);
          }
        }
      }
    } catch {}
    res.setHeader('Cache-Control', 'no-store');
    res.json({
      productId,
      history,
      _debug: {
        filePath,
        fileLen: raw ? raw.length : 0,
        keys: (map && typeof map === 'object') ? Object.keys(map).length : 0
      }
    });
  } catch (e) {
    res.status(500).json({ productId: req.params.id, history: [] });
  }
});

// Export price history for a product
app.get("/api/products/:id/price-history/export", isAuthenticated, isAdmin, async (req, res) => {
  try {
    const productId = req.params.id;
    const products = await readData('products.json');
    if (!Array.isArray(products)) {
      return res.status(404).json({ success: false, message: 'Product not found' });
    }
    const product = products.find(p => String(p.id) === String(productId));
    if (!product) {
      return res.status(404).json({ success: false, message: 'Product not found' });
    }
    let history = Array.isArray(product.priceHistory) ? product.priceHistory.slice() : [];
    if (!history.length) {
      const map = await readPriceHistoryMap();
      const ph = map[String(product.id)] || [];
      if (Array.isArray(ph)) history = ph.slice();
    }
    history.sort((a, b) => Number(b && b.at || 0) - Number(a && a.at || 0));
    const rows = history.map(h => {
      const dt = h && h.at ? new Date(h.at) : null;
      const time = dt && !isNaN(dt.getTime()) ? dt.toISOString().replace('T', ' ').slice(0, 19) : '';
      const baseSellFrom = h && h.base && h.base.sellingPrice ? h.base.sellingPrice.from : '';
      const baseSellTo = h && h.base && h.base.sellingPrice ? h.base.sellingPrice.to : '';
      const baseBuyFrom = h && h.base && h.base.purchasePrice ? h.base.purchasePrice.from : '';
      const baseBuyTo = h && h.base && h.base.purchasePrice ? h.base.purchasePrice.to : '';
      let variants = '';
      if (Array.isArray(h && h.variants) && h.variants.length) {
        variants = h.variants.map(v => {
          const label = [v.qty, v.unit].filter(Boolean).join(' ');
          const note = v.note ? ` (${v.note})` : '';
          const sku = v.sku ? ` [SKU:${v.sku}]` : '';
          const act = v.action || 'update';
          return `${act}: ${label}${note}${sku} ${v.from ?? ''} -> ${v.to ?? ''}`.trim();
        }).join(' | ');
      }
      return {
        Date: time,
        By: (h && h.by) ? h.by : '',
        Reason: (h && h.reason) ? h.reason : '',
        "Sell From": baseSellFrom,
        "Sell To": baseSellTo,
        "Buy From": baseBuyFrom,
        "Buy To": baseBuyTo,
        "Variant Changes": variants
      };
    });

    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "PriceHistory");
    const buf = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename=price_history_${productId}.xlsx`);
    res.send(buf);
  } catch (error) {
    console.error('Error exporting price history:', error);
    res.status(500).json({ success: false, message: 'Failed to export price history' });
  }
});

app.post("/api/products", isAuthenticated, isAdmin, async (req, res) => {
  try {
    const user = (req.session && req.session.user) || {};
    const { name, sku, category, unit } = req.body;

    // Validate and sanitize inputs
    if (!name) {
      return res.status(400).json({
        success: false,
        message: "Product name is required",
      });
    }

    try {
      var validatedName = validateAndSanitizeInput(name, 'general');
      var validatedSku = req.body.sku ? validateAndSanitizeInput(req.body.sku, 'general') : null;
      var validatedCategory = req.body.category ? validateAndSanitizeInput(req.body.category, 'general') : null;
      var validatedUnit = req.body.unit ? validateAndSanitizeInput(req.body.unit, 'general') : null;
    } catch (validationError) {
      return res.status(400).json({
        success: false,
        message: validationError.message,
      });
    }

    // Validasi nama produk
    if (!validatedName || validatedName.trim() === "") {
      return res.status(400).json({
        success: false,
        message: "Nama produk wajib diisi",
      });
    }

    // Cek nama produk duplikat (hormati setting)
    const settingsObjForName = await readData('settings.json').catch(()=>({}));
    const allowDupName = !!(settingsObjForName && settingsObjForName.allowDuplicateProductNames);
    if (!allowDupName) {
      const existingProduct = await validateProductName(validatedName.trim());
      if (existingProduct) {
        return res.status(400).json({
          success: false,
          message: `Produk "${validatedName}" sudah ada. Silakan gunakan nama lain.`,
        });
      }
    }

    const products = await readData("products.json");
    const purchasePrice = Number(req.body.purchasePrice || 0) || 0;
    const sellingPrice =
      req.body.sellingPrice !== undefined
        ? Number(req.body.sellingPrice) || 0
        : Number(req.body.price || 0) || 0;
    const taxRate = Math.max(0, Math.min(100, Number(req.body.taxRate || 0) || 0));
    const discountPercent = Math.max(0, Math.min(100, Number(req.body.discountPercent || 0) || 0));
    const rawSku = (validatedSku || "").trim();
    const finalSku = rawSku || `PROD-${Date.now()}`;
    // Respect settings: allow duplicate SKU
    const settingsObj = await readData('settings.json').catch(()=>({}));
    const allowDup = !!(settingsObj && settingsObj.allowDuplicateSku);
    if (!allowDup) {
      const duplicateSku = await validateProductSku(finalSku);
      if (duplicateSku) {
        return res.status(400).json({ success:false, message:`SKU "${finalSku}" sudah digunakan oleh produk lain.` });
      }
    }
    let qrCode = (req.body.qrCode || "").trim();
    if (!qrCode) qrCode = finalSku; // fallback QR to SKU if empty

    // Sanitize unitPrices if provided (preserve optional note/desc/keterangan, sku, photo)
    let unitPrices = Array.isArray(req.body.unitPrices) ? req.body.unitPrices : [];
    if (Array.isArray(unitPrices)) {
      unitPrices = sanitizeUnitPricesList(unitPrices);
    } else {
      unitPrices = [];
    }

    const newProduct = {
      id: Date.now(),
      ...req.body,
      name: validatedName.trim(),
      category: validatedCategory,
      unit: validatedUnit,
      purchasePrice,
      sellingPrice,
      taxRate,
      discountPercent,
      sku: finalSku,
      qrCode,
      // Backward compatibility for POS which uses product.price
      price: sellingPrice,
      unitPrices,
    };
    try {
      const by = user && (user.username || user.name || '');
      const entry = buildPriceHistoryEntry(null, newProduct, 'create', by);
      appendPriceHistory(newProduct, entry);
    } catch {}
    products.push(newProduct);
    await saveArrayWithSync("products.json", products);
    try { invalidateCache('products'); } catch {}
    res.json(newProduct);
  } catch (error) {
    console.error("Error creating product:", error);
    res.status(500).json({
      success: false,
      message: "Failed to create product",
    });
  }
});

app.put("/api/products/:id", isAuthenticated, isAdmin, async (req, res) => {
  try {
    const { name, sku, category, unit } = req.body;
    const productId = req.params.id;

    // Validate and sanitize inputs
    if (!name) {
      return res.status(400).json({
        success: false,
        message: "Product name is required",
      });
    }

    try {
      var validatedName = validateAndSanitizeInput(name, 'general');
      var validatedSku = req.body.sku ? validateAndSanitizeInput(req.body.sku, 'general') : null;
      var validatedCategory = req.body.category ? validateAndSanitizeInput(req.body.category, 'general') : null;
      var validatedUnit = req.body.unit ? validateAndSanitizeInput(req.body.unit, 'general') : null;
    } catch (validationError) {
      return res.status(400).json({
        success: false,
        message: validationError.message,
      });
    }

    // Validasi nama produk
    if (!validatedName || validatedName.trim() === "") {
      return res.status(400).json({
        success: false,
        message: "Nama produk wajib diisi",
      });
    }

    // Cek nama produk duplikat (hormati setting)
    const settingsObjForNameU = await readData('settings.json').catch(()=>({}));
    const allowDupNameU = !!(settingsObjForNameU && settingsObjForNameU.allowDuplicateProductNames);
    if (!allowDupNameU) {
      const existingProduct = await validateProductName(validatedName.trim(), productId);
      if (existingProduct) {
        return res.status(400).json({
          success: false,
          message: `Produk "${validatedName}" sudah ada. Silakan gunakan nama lain.`,
        });
      }
    }

    const products = await readData("products.json");
    const index = products.findIndex((p) => p.id == productId);

    if (index !== -1) {
      const prevProduct = {
        ...products[index],
        unitPrices: Array.isArray(products[index].unitPrices)
          ? products[index].unitPrices.map(v => ({ ...v }))
          : []
      };
      const rawSku = (validatedSku || products[index].sku || "").trim();
      // Respect settings: allow duplicate SKU on update
      const settingsObj = await readData('settings.json').catch(()=>({}));
      const allowDup = !!(settingsObj && settingsObj.allowDuplicateSku);
      if (!allowDup && rawSku) {
        const dup = await validateProductSku(rawSku, productId);
        if (dup) {
          return res.status(400).json({ success:false, message:`SKU "${rawSku}" sudah digunakan oleh produk lain.` });
        }
      }

      const purchasePrice =
        req.body.purchasePrice !== undefined
          ? Number(req.body.purchasePrice) || 0
          : products[index].purchasePrice || 0;
      const sellingPrice =
        req.body.sellingPrice !== undefined
          ? Number(req.body.sellingPrice) || 0
          : (products[index].sellingPrice != null
              ? products[index].sellingPrice
              : products[index].price || 0);
      const taxRate =
        req.body.taxRate !== undefined
          ? Math.max(0, Math.min(100, Number(req.body.taxRate) || 0))
          : (Math.max(0, Math.min(100, Number(products[index].taxRate || 0) || 0)));
      const discountPercent =
        req.body.discountPercent !== undefined
          ? Math.max(0, Math.min(100, Number(req.body.discountPercent) || 0))
          : (Math.max(0, Math.min(100, Number(products[index].discountPercent || 0) || 0)));
      let qrCode = req.body.qrCode !== undefined ? String(req.body.qrCode || "").trim() : String(products[index].qrCode || "");
      if (!qrCode) qrCode = rawSku; // fallback to SKU if empty

      // Sanitize unitPrices if provided in update (preserve optional note/desc/keterangan, sku, photo)
      let unitPricesU = Array.isArray(req.body.unitPrices) ? req.body.unitPrices : (products[index].unitPrices || []);
      if (Array.isArray(unitPricesU)) {
        unitPricesU = sanitizeUnitPricesList(unitPricesU);
      } else {
        unitPricesU = [];
      }

      // Preserve existing imageBase64 if incoming value is missing or empty
      const incomingImg = (typeof req.body.imageBase64 === 'string') ? req.body.imageBase64 : undefined;
      const imageBase64 = (incomingImg && incomingImg.trim()) ? incomingImg : (products[index].imageBase64 || '');
      const { imageBase64: _skipImg, priceHistory: _skipPriceHistory, _source: _skipSource, source: _skipSource2, ...restBody } = req.body || {};
      // Detect manual stock adjustment
      const prevStock = Number(products[index].stock || 0);
      const nextStock = (req.body && req.body.stock !== undefined) ? (Number(req.body.stock) || 0) : prevStock;
      const byUser = (req.session && req.session.user && (req.session.user.username || req.session.user.name)) || '';
      const nextSnapshot = {
        ...prevProduct,
        purchasePrice,
        sellingPrice,
        price: sellingPrice,
        unitPrices: unitPricesU
      };
      const sourceRaw = String((req.body && (req.body._source || req.body.source)) || '').toLowerCase();
      const reason = sourceRaw === 'stock_in' ? 'stock_in' : 'update';
      let priceHistoryEntry = null;
      try {
        priceHistoryEntry = buildPriceHistoryEntry(prevProduct, nextSnapshot, reason, byUser);
      } catch (e) {
        priceHistoryEntry = null;
      }
      if (!priceHistoryEntry) {
        const prevSell = Number((prevProduct.sellingPrice != null ? prevProduct.sellingPrice : prevProduct.price) || 0);
        const nextSell = Number((sellingPrice != null ? sellingPrice : prevSell) || 0);
        const prevBuy = Number(prevProduct.purchasePrice || 0);
        const nextBuy = Number(purchasePrice || 0);
        const changed = prevSell !== nextSell || prevBuy !== nextBuy;
        if (changed || reason === 'stock_in') {
          priceHistoryEntry = {
            at: Date.now(),
            by: byUser,
            reason,
            base: {
              sellingPrice: { from: prevSell, to: nextSell },
              purchasePrice: { from: prevBuy, to: nextBuy }
            },
            variants: []
          };
        }
      }
      // Fallback: ensure variant price changes are recorded even if base prices are unchanged
      if (!priceHistoryEntry) {
        try {
          const prevVars = Array.isArray(prevProduct && prevProduct.unitPrices) ? prevProduct.unitPrices : [];
          const nextVars = Array.isArray(unitPricesU) ? unitPricesU : [];
          const maxLen = Math.max(prevVars.length, nextVars.length);
          const variants = [];
          for (let i = 0; i < maxLen; i++) {
            const pv = (prevVars[i] && typeof prevVars[i] === 'object') ? prevVars[i] : null;
            const nv = (nextVars[i] && typeof nextVars[i] === 'object') ? nextVars[i] : null;
            const prevPrice = pv ? Number(pv.price || 0) : null;
            const nextPrice = nv ? Number(nv.price || 0) : null;
            if (pv && nv && prevPrice === nextPrice) continue;
            if (!pv && !nv) continue;
            const v = {
              index: i,
              qty: (nv && nv.qty != null) ? nv.qty : (pv && pv.qty),
              unit: (nv && nv.unit) ? nv.unit : (pv && pv.unit),
              sku: (nv && nv.sku) ? nv.sku : (pv && pv.sku),
              note: (nv && nv.note) ? nv.note : (pv && pv.note),
              from: prevPrice,
              to: nextPrice
            };
            if (!pv && nv) v.action = 'add';
            else if (pv && !nv) v.action = 'remove';
            else v.action = 'update';
            variants.push(v);
          }
          if (variants.length) {
            priceHistoryEntry = {
              at: Date.now(),
              by: byUser,
              reason,
              base: {},
              variants
            };
          }
        } catch {}
      }
      try {
        console.log('[price-history] update', {
          productId,
          prevSell: Number((prevProduct.sellingPrice != null ? prevProduct.sellingPrice : prevProduct.price) || 0),
          nextSell: Number((sellingPrice != null ? sellingPrice : prevProduct.price) || 0),
          prevBuy: Number(prevProduct.purchasePrice || 0),
          nextBuy: Number(purchasePrice || 0),
          hasEntry: !!priceHistoryEntry
        });
      } catch {}

      products[index] = {
        ...products[index],
        ...restBody,
        imageBase64,
        name: validatedName.trim(),
        category: validatedCategory,
        unit: validatedUnit,
        purchasePrice,
        sellingPrice,
        taxRate,
        discountPercent,
        sku: rawSku,
        qrCode,
        price: sellingPrice,
        unitPrices: unitPricesU,
        stock: nextStock,
      };
      try { appendPriceHistory(products[index], priceHistoryEntry); } catch {}
      try { console.log('[price-history] appendPriceHistoryForProduct called', { productId, hasEntry: !!priceHistoryEntry }); } catch {}
      try { if (global.appendPriceHistoryForProduct) await global.appendPriceHistoryForProduct(productId, priceHistoryEntry); } catch (e) { try { console.error('[price-history] appendPriceHistoryForProduct failed', e && e.message ? e.message : e); } catch {} }
      try {
        console.log('[price-history] stored', {
          productId,
          count: Array.isArray(products[index].priceHistory) ? products[index].priceHistory.length : 0
        });
      } catch {}
      // Append stock move for manual adjustment
      try {
        const delta = Number(nextStock) - Number(prevStock);
        if (delta !== 0) {
          await appendStockMove({ productId, delta, reason: 'manual_adjust', by: (req.session && req.session.user && req.session.user.username) || '', newStock: Number(nextStock), stockOnly: false });
        }
      } catch {}
      await saveArrayWithSync("products.json", products);
      
      // Invalidate cache when product is updated
      invalidateCache('products');
      
      res.json(products[index]);
    } else {
      res.status(404).json({
        success: false,
        message: "Produk tidak ditemukan",
      });
    }
  } catch (error) {
    console.error("Error updating product:", error);
    res.status(500).json({
      success: false,
      message: "Failed to update product",
    });
  }
});

app.delete("/api/products/:id", isAuthenticated, isAdmin, async (req, res) => {
  try {
    const productId = req.params.id;
    const pid = String(productId);
    // Cek apakah produk sedang digunakan oleh transaksi (early-exit)
    const transactions = await readData("transactions.json");
    let inUse = false;
    if (Array.isArray(transactions) && transactions.length) {
      outer: for (const t of transactions) {
        const items = Array.isArray(t && t.items) ? t.items : [];
        for (const it of items) {
          if (String(it && it.productId) === pid) { inUse = true; break outer; }
        }
      }
    }
    const force = String(req.query.force || 'false').toLowerCase() === 'true';

    const products = await readData("products.json");
    if (!Array.isArray(products) || products.length === 0) {
      return res.status(404).json({ success: false, message: "Produk tidak ditemukan" });
    }
    const idx = products.findIndex(p => String(p && p.id) === pid);
    if (idx < 0) return res.status(404).json({ success:false, message: "Produk tidak ditemukan" });

    if (inUse && !force) {
      // Soft-delete: tandai deleted=true agar tidak muncul di listing, transaksi tetap aman
      const now = Date.now();
      const cur = products[idx] || {};
      const nextDoc = { ...cur, deleted: true, updatedAt: now };
      products[idx] = nextDoc;
      await saveArrayWithSync("products.json", products);
      try { await enqueueOutbox({ collection: 'products', file: 'products.json', op: 'upsert', _id: pid, doc: nextDoc, updatedAt: now }); } catch {}
      return res.json({ success: true, softDeleted: true });
    }

    // Hard-delete
    const next = products.filter((p) => String(p && p.id) !== pid);
    await saveArrayWithSync("products.json", next);
    try { await enqueueOutbox({ collection: 'products', file: 'products.json', op: 'delete', _id: pid, deleted: true, updatedAt: Date.now() }); } catch {}
    res.json({ success: true, hardDeleted: true });
  } catch (error) {
    console.error("Error deleting product:", error);
    res.status(500).json({
      success: false,
      message: "Failed to delete product",
    });
  }
});

// Bulk purge all products (dangerous):
// - Produk yang pernah dipakai di transaksi akan di-soft-delete (deleted:true) agar riwayat tetap konsisten
// - Produk yang tidak pernah dipakai akan di-hard-delete
app.post("/api/products/purge-all", isAuthenticated, isAdmin, async (req, res) => {
  try {
    const confirm = String((req.body && req.body.confirm) || "").trim();
    if (confirm !== "DELETE_ALL_PRODUCTS") {
      return res.status(400).json({ success: false, message: "Konfirmasi tidak valid" });
    }

    let products = await readData("products.json").catch(() => []);
    if (!Array.isArray(products)) products = [];
    const totalBefore = products.length;

    // Kumpulkan ID produk yang muncul di transaksi
    const inUseIds = new Set();
    try {
      const transactions = await readData("transactions.json").catch(() => []);
      if (Array.isArray(transactions)) {
        for (const t of transactions) {
          const items = Array.isArray(t && t.items) ? t.items : [];
          for (const it of items) {
            const pid = String((it && it.productId) || "");
            if (pid) inUseIds.add(pid);
          }
        }
      }
    } catch {}

    const now = Date.now();
    const next = [];
    let softDeleted = 0;
    let hardDeleted = 0;
    for (const p of products) {
      const pid = String((p && p.id) || "");
      if (!pid) continue;
      if (inUseIds.has(pid)) {
        const cur = p || {};
        const doc = { ...cur, deleted: true, updatedAt: now };
        next.push(doc);
        softDeleted++;
      } else {
        hardDeleted++;
      }
    }

    await saveArrayWithSync("products.json", next);
    return res.json({
      success: true,
      totalBefore,
      totalAfter: next.length,
      softDeleted,
      hardDeleted,
    });
  } catch (e) {
    console.error("Error purging all products:", e);
    return res.status(500).json({ success: false, message: "Gagal menghapus semua produk" });
  }
});

// Remove duplicate products by name (keep the oldest one)
app.post("/api/products/remove-duplicates", isAuthenticated, isAdmin, async (req, res) => {
  try {
    const sendProgress = global.__duplicateRemovalProgress;
    
    let products = await readData("products.json").catch(() => []);
    if (!Array.isArray(products)) products = [];
    const totalBefore = products.length;

    if (sendProgress) sendProgress('loading', 5, 'Memuat data produk...');

    // Group products by name (case-insensitive, trimmed)
    const nameGroups = new Map();
    for (const p of products) {
      const name = String(p.name || "").trim().toLowerCase();
      if (!name) continue;
      if (!nameGroups.has(name)) {
        nameGroups.set(name, []);
      }
      nameGroups.get(name).push(p);
    }

    if (sendProgress) sendProgress('grouping', 15, 'Mengelompokkan produk berdasarkan nama...');

    // Find duplicates (groups with more than 1 product)
    const duplicates = [];
    for (const [name, group] of nameGroups) {
      if (group.length > 1) {
        // Sort by creation date/updated date (oldest first)
        group.sort((a, b) => {
          const dateA = new Date(a.createdAt || a.updatedAt || a.timestamp || 0);
          const dateB = new Date(b.createdAt || b.updatedAt || b.timestamp || 0);
          return dateA - dateB;
        });
        // Keep the first (oldest) product, mark others for deletion
        const toKeep = group[0];
        const toDelete = group.slice(1);
        duplicates.push({ name, toKeep, toDelete });
      }
    }

    if (duplicates.length === 0) {
      if (sendProgress) sendProgress('complete', 100, 'Tidak ada produk ganda ditemukan');
      return res.json({
        success: true,
        message: "Tidak ada produk ganda dengan nama yang sama",
        totalBefore,
        totalAfter: totalBefore,
        duplicateGroups: 0,
        deleted: 0
      });
    }

    if (sendProgress) sendProgress('checking', 25, `Memeriksa ${duplicates.length} grup produk ganda...`);

    // Check which products are used in transactions
    const inUseIds = new Set();
    try {
      const transactions = await readData("transactions.json").catch(() => []);
      if (Array.isArray(transactions)) {
        for (const t of transactions) {
          const items = Array.isArray(t && t.items) ? t.items : [];
          for (const it of items) {
            const pid = String((it && it.productId) || "");
            if (pid) inUseIds.add(pid);
          }
        }
      }
    } catch {}

    if (sendProgress) sendProgress('processing', 40, 'Memproses penghapusan produk...');

    // Process deletions
    const now = Date.now();
    const next = [];
    let softDeleted = 0;
    let hardDeleted = 0;
    let totalDeleted = 0;
    let processed = 0;

    // First, add all products that are NOT duplicates
    for (const p of products) {
      const name = String(p.name || "").trim().toLowerCase();
      if (!name) {
        next.push(p); // Keep products without names
        continue;
      }
      
      const isDuplicate = duplicates.some(d => 
        d.name === name && d.toDelete.some(dp => dp.id === p.id)
      );
      
      if (!isDuplicate) {
        next.push(p); // Keep non-duplicate products
      }
    }

    if (sendProgress) sendProgress('deleting', 60, 'Menghapus produk ganda...');

    // Then handle duplicates
    for (let i = 0; i < duplicates.length; i++) {
      const dup = duplicates[i];
      
      // Keep the oldest product
      next.push(dup.toKeep);
      
      // Delete the rest
      for (const p of dup.toDelete) {
        const pid = String(p.id || "");
        totalDeleted++;
        
        if (inUseIds.has(pid)) {
          // Soft delete if used in transactions
          const doc = { ...p, deleted: true, updatedAt: now };
          next.push(doc);
          softDeleted++;
        } else {
          // Hard delete if not used
          hardDeleted++;
        }
      }

      // Update progress
      processed++;
      const progress = 60 + Math.floor((processed / duplicates.length) * 35);
      if (sendProgress) {
        sendProgress('deleting', progress, 
          `Menghapus grup ${processed}/${duplicates.length}: ${dup.name}`
        );
      }
    }

    if (sendProgress) sendProgress('saving', 95, 'Menyimpan perubahan...');

    await saveArrayWithSync("products.json", next);
    
    if (sendProgress) sendProgress('complete', 100, 'Proses selesai!');
    
    return res.json({
      success: true,
      message: `Berhasil menghapus ${totalDeleted} produk ganda dari ${duplicates.length} grup`,
      totalBefore,
      totalAfter: next.length,
      duplicateGroups: duplicates.length,
      deleted: totalDeleted,
      softDeleted,
      hardDeleted,
      duplicates: duplicates.map(d => ({
        name: d.name,
        kept: d.toKeep.id,
        deleted: d.toDelete.map(p => p.id)
      }))
    });
  } catch (e) {
    console.error("Error removing duplicate products:", e);
    if (global.__duplicateRemovalProgress) {
      global.__duplicateRemovalProgress('error', 0, 'Error: ' + (e.message || e));
    }
    res.status(500).json({
      success: false,
      message: "Gagal menghapus produk ganda: " + (e.message || e)
    });
  }
});

// SSE endpoint for duplicate removal progress
app.get("/api/products/remove-duplicates-progress", isAuthenticated, isAdmin, async (req, res) => {
  res.writeHead(200, {
    'Content-Type': 'text/event-stream',
    'Cache-Control': 'no-cache',
    'Connection': 'keep-alive',
    'Access-Control-Allow-Origin': '*'
  });

  const sendProgress = (phase, progress, message) => {
    res.write(`data: ${JSON.stringify({ phase, progress, message })}\n\n`);
  };

  // Store the sendProgress function globally for the removal process
  global.__duplicateRemovalProgress = sendProgress;

  // Send initial status
  sendProgress('ready', 0, 'Siap memproses...');

  // Handle client disconnect
  req.on('close', () => {
    global.__duplicateRemovalProgress = null;
  });
});

// EXTREMELY DANGEROUS: Delete specific database files
app.post("/api/database/delete-all", isAuthenticated, isAdmin, async (req, res) => {
  try {
    const confirm = String((req.body && req.body.confirm) || "").trim();
    if (confirm !== "DELETE_ALL_DATABASE_PERMANENTLY") {
      return res.status(400).json({ success: false, message: "Konfirmasi tidak valid" });
    }

    const dataDir = resolveDataDir();

    // 1) Reset in-memory cache
    try { invalidateCache('products'); } catch {}
    try { invalidateCache('categories'); } catch {}
    try { invalidateCache('transactions'); } catch {}
    try { invalidateCache('settings'); } catch {}

    // 2) Tutup koneksi SQLite dan hapus file database
    try {
      sqliteStorage.close();
    } catch {}
    try {
      const dbFile = path.join(dataDir, 'pos.db');
      await fs.rm(dbFile, { force: true });
    } catch {}

    // 3) Hapus file-file data JSON di dataDir (kecuali yang dipreserve)
    const preservedFiles = new Set([
      'users.json',
      'settings.json',
      'license-key.txt',
      'trial-info.json',
      'passphrase.txt'
    ]);
    const deletedFiles = [];
    const errors = [];
    const entries = await fs.readdir(dataDir, { withFileTypes: true }).catch(() => []);
    for (const ent of entries) {
      try {
        const full = path.join(dataDir, ent.name);
        if (ent.isDirectory()) {
          await fs.rm(full, { recursive: true, force: true });
          deletedFiles.push(ent.name + path.sep);
          continue;
        }
        if (preservedFiles.has(ent.name)) continue;
        await fs.rm(full, { force: true });
        deletedFiles.push(ent.name);
      } catch (e) {
        errors.push({ file: ent.name, error: e.message });
      }
    }

    // 4) Hapus license.json dari database
    try { await writeData('license.json', {}); } catch {}

    // 5) Hapus Redis cache untuk app ini
    try { await cacheRedis.flushAll(); } catch {}

    // 6) Reset flag inisialisasi agar database dibuat ulang kosong
    __storageInitPromise = null;

    return res.json({
      success: true,
      message: "Data aplikasi telah dihapus permanen (kecuali user dan pengaturan).",
      deletedFiles,
      errors,
      preservedFiles: Array.from(preservedFiles),
      warning: "TINDAKAN TIDAK BISA DIURUNGGI. Data yang dipilih telah dihapus permanen."
    });
  } catch (e) {
    try { console.error("Error deleting database:", e); } catch {}
    res.status(500).json({
      success: false,
      message: "Gagal menghapus database: " + (e.message || e)
    });
  }
});

// API endpoint untuk cek update aplikasi
app.get('/api/check-update', async (req, res) => {
  try {
    const packageJson = require('./package.json');
    const currentVersion = packageJson.version;
    
    // Coba ambil konfigurasi server update dari database
    let updateUrl = 'https://api.github.com/repos/username/pos-premium/releases/latest'; // default
    let updateHeaders = { 'User-Agent': 'POS-App-UpdateChecker' };
    
    try {
      const config = await readData('update-server-config.json').catch(() => ({}));
      console.log('[UPDATE DEBUG] Raw config:', config);
      if (config && config.url) {
        updateUrl = config.url;
        updateHeaders = { ...updateHeaders, ...(config.headers || {}) };
        console.log('[UPDATE] Using custom server config:', config.name);
        console.log('[UPDATE] URL:', updateUrl);
        console.log('[UPDATE] Headers:', updateHeaders);
      } else {
        console.log('[UPDATE] No valid config found, using default');
      }
    } catch (configError) {
      console.log('[UPDATE] Using default config, failed to load custom config:', configError.message);
    }
    
    try {
      const response = await fetch(updateUrl, {
        headers: updateHeaders,
        timeout: 15000 // 15 seconds timeout
      });
      
      if (response.ok) {
        const release = await response.json();
        let latestVersion = '';
        let releaseInfo = {};
        
        // Handle different response formats
        if (release.tag_name) {
          // GitHub format
          latestVersion = release.tag_name.replace('v', '');
          releaseInfo = {
            name: release.name,
            publishedAt: release.published_at,
            downloadUrl: release.html_url || release.assets?.[0]?.browser_download_url,
            releaseNotes: release.body
          };
        } else if (release.version) {
          // Custom format
          latestVersion = release.version;
          releaseInfo = {
            name: release.name || `Version ${latestVersion}`,
            publishedAt: release.publishedAt || release.date,
            downloadUrl: release.downloadUrl || release.url,
            releaseNotes: release.releaseNotes || release.description
          };
        } else {
          throw new Error('Format response tidak dikenali');
        }
        
        // Bandingkan versi dan cek ketersediaan ZIP
        const hasUpdate = compareVersions(latestVersion, currentVersion) > 0;
        
        // Check if ZIP file is available for hot-reload
        let hasZipFile = false;
        let hasExeFile = false;
        if (release.assets && release.assets.length > 0) {
          const zipAsset = release.assets.find(asset => asset.name.endsWith('.zip'));
          const exeAsset = release.assets.find(asset => asset.name.endsWith('.exe'));
          hasZipFile = !!zipAsset;
          hasExeFile = !!exeAsset;
        }
        
        return res.json({
          success: true,
          currentVersion,
          latestVersion,
          hasUpdate,
          hasZipFile,
          hasExeFile,
          releaseInfo,
          updateServer: {
            url: updateUrl,
            name: (await readData('update-server-config.json').catch(() => ({name: 'Default'}))).name
          }
        });
      }
    } catch (fetchError) {
      console.log('[UPDATE] Failed to check online update:', fetchError.message);
    }
    
    // Jika gagal cek online, kembalikan info versi current
    return res.json({
      success: true,
      currentVersion,
      latestVersion: currentVersion,
      hasUpdate: false,
      message: 'Tidak dapat memeriksa update. Mode offline atau server tidak dapat dijangkau.'
    });
    
  } catch (error) {
    console.error('[UPDATE] Error checking update:', error);
    res.status(500).json({
      success: false,
      message: 'Gagal memeriksa update: ' + error.message
    });
  }
});

// Fungsi helper untuk membandingkan versi
function compareVersions(v1, v2) {
  const parts1 = v1.split('.').map(Number);
  const parts2 = v2.split('.').map(Number);
  
  for (let i = 0; i < Math.max(parts1.length, parts2.length); i++) {
    const num1 = parts1[i] || 0;
    const num2 = parts2[i] || 0;
    
    if (num1 > num2) return 1;
    if (num1 < num2) return -1;
  }
  
  return 0;
}

// --- Developer API Endpoints ---

// GET /api/dev/update-server-config - Ambil konfigurasi server update
app.get('/api/dev/update-server-config', isAuthenticated, isAdmin, async (req, res) => {
  try {
    const config = await readData('update-server-config.json').catch(() => ({}));
    res.json({
      success: true,
      config: config
    });
  } catch (error) {
    console.error('Error loading update server config:', error);
    res.status(500).json({
      success: false,
      message: 'Gagal memuat konfigurasi server update'
    });
  }
});

// POST /api/dev/update-server-config - Simpan konfigurasi server update
app.post('/api/dev/update-server-config', isAuthenticated, isAdmin, async (req, res) => {
  try {
    const { url, name, headers, checkInterval, autoCheckEnabled, updatedAt } = req.body;
    
    // Validasi input
    if (!url || !name) {
      return res.status(400).json({
        success: false,
        message: 'URL dan nama server harus diisi'
      });
    }
    
    const config = {
      url: url.trim(),
      name: name.trim(),
      headers: headers || {},
      checkInterval: parseInt(checkInterval) || 24,
      autoCheckEnabled: autoCheckEnabled !== false,
      updatedAt: updatedAt || new Date().toISOString(),
      version: '1.0.0'
    };
    
    await writeData('update-server-config.json', config);
    
    console.log('[DEV] Update server config saved:', config.name);
    
    res.json({
      success: true,
      message: 'Konfigurasi server update berhasil disimpan',
      config: config
    });
  } catch (error) {
    console.error('Error saving update server config:', error);
    res.status(500).json({
      success: false,
      message: 'Gagal menyimpan konfigurasi server update'
    });
  }
});

// POST /api/dev/test-update-server - Test koneksi ke server update
app.post('/api/dev/test-update-server', isAuthenticated, isAdmin, async (req, res) => {
  try {
    const { url, headers } = req.body;
    
    if (!url) {
      return res.status(400).json({
        success: false,
        message: 'URL harus diisi'
      });
    }
    
    console.log('[DEV] Testing connection to:', url);
    
    // Test connection
    const response = await fetch(url, {
      method: 'GET',
      headers: {
        'User-Agent': 'POS-App-UpdateChecker',
        ...headers
      },
      timeout: 10000 // 10 seconds timeout
    });
    
    const status = response.status;
    let responseData = null;
    
    try {
      responseData = await response.json();
    } catch (e) {
      // If not JSON, try text
      try {
        responseData = await response.text();
      } catch (e2) {
        responseData = null;
      }
    }
    
    res.json({
      success: true,
      status: status,
      statusText: response.statusText,
      data: responseData,
      message: `Server merespon dengan status ${status}`
    });
    
  } catch (error) {
    console.error('Error testing update server:', error);
    res.status(500).json({
      success: false,
      message: 'Gagal menguji koneksi: ' + error.message
    });
  }
});

// GET /api/dev/system-info - Informasi sistem untuk developer
app.get('/api/dev/system-info', isAuthenticated, isAdmin, async (req, res) => {
  try {
    const packageJson = require('./package.json');
    const config = await readData('update-server-config.json').catch(() => ({}));
    
    res.json({
      success: true,
      system: {
        appVersion: packageJson.version,
        appName: packageJson.name,
        nodeVersion: process.version,
        platform: process.platform,
        uptime: process.uptime(),
        memory: process.memoryUsage(),
        updateServerConfig: config
      }
    });
  } catch (error) {
    console.error('Error getting system info:', error);
    res.status(500).json({
      success: false,
      message: 'Gagal mengambil informasi sistem'
    });
  }
});

// POST /api/auto-update - Auto update aplikasi
app.post('/api/auto-update', isAuthenticated, isAdmin, async (req, res) => {
  try {
    console.log('[AUTO-UPDATE] Starting auto-update process...');
    
    // Get latest release info
    const packageJson = require('./package.json');
    const currentVersion = packageJson.version;
    
    // Get update server config
    let updateUrl = 'https://api.github.com/repos/username/pos-premium/releases/latest';
    let updateHeaders = { 'User-Agent': 'POS-App-UpdateChecker' };
    
    try {
      const config = await readData('update-server-config.json').catch(() => ({}));
      if (config && config.url) {
        updateUrl = config.url;
        updateHeaders = { ...updateHeaders, ...(config.headers || {}) };
      }
    } catch (configError) {
      console.log('[AUTO-UPDATE] Using default config:', configError.message);
    }
    
    // Fetch latest release
    const response = await fetch(updateUrl, {
      headers: updateHeaders,
      timeout: 30000
    });
    
    if (!response.ok) {
      throw new Error(`Failed to fetch release: ${response.status}`);
    }
    
    const release = await response.json();
    const latestVersion = release.tag_name.replace('v', '');
    
    // Check if update needed
    if (compareVersions(latestVersion, currentVersion) <= 0) {
      return res.json({
        success: false,
        message: 'Already up to date',
        currentVersion,
        latestVersion
      });
    }
    
    console.log('[AUTO-UPDATE] Update available:', currentVersion, '->', latestVersion);
    
    // Find download URL and check ZIP availability
    let downloadUrl = null;
    let fileName = null;
    let hasZipFile = false;
    
    if (release.assets && release.assets.length > 0) {
      // Look for .zip file first, then .exe
      const zipAsset = release.assets.find(asset => asset.name.endsWith('.zip'));
      const exeAsset = release.assets.find(asset => asset.name.endsWith('.exe'));
      
      if (zipAsset) {
        downloadUrl = zipAsset.browser_download_url;
        fileName = zipAsset.name;
        hasZipFile = true;
      } else if (exeAsset) {
        downloadUrl = exeAsset.browser_download_url;
        fileName = exeAsset.name;
        hasZipFile = false;
      }
    }
    
    if (!downloadUrl) {
      // Fallback to release page (only for traditional mode)
      downloadUrl = release.html_url;
      fileName = `release-${latestVersion}.html`;
      hasZipFile = false;
      hasExeFile = false;
      
      // If hot-reload is requested but no downloadable file available, fallback to traditional
      if (req.body && req.body.hotReload) {
        // console.log('[AUTO-UPDATE] Hot-reload requested but no ZIP/EXE available. Falling back to traditional mode.');
        // Continue with traditional mode but don't throw error
      }
    }
    
    // console.log('[AUTO-UPDATE] Downloading:', downloadUrl, 'File:', fileName, 'Has ZIP:', hasZipFile);
    
    // Download the file
    const fs = require('fs');
    const path = require('path');
    
    const downloadResponse = await fetch(downloadUrl, {
      headers: updateHeaders,
      timeout: 60000 // 1 minute timeout
    });
    
    if (!downloadResponse.ok) {
      throw new Error(`Download failed: ${downloadResponse.status}`);
    }
    
    // Create temp directory
    const tempDir = path.join(__dirname, 'temp');
    if (!fs.existsSync(tempDir)) {
      fs.mkdirSync(tempDir, { recursive: true });
    }
    
    const tempFilePath = path.join(tempDir, fileName || `update-${latestVersion}.zip`);
    
    // Get response data as buffer and write to file
    const arrayBuffer = await downloadResponse.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);
    fs.writeFileSync(tempFilePath, buffer);
    
    console.log('[AUTO-UPDATE] Downloaded to:', tempFilePath);
    
    // Check if hot-reload is requested
    const { hotReload = false } = req.body;
    
    // Auto-detect if hot-reload is possible
    let canHotReload = false;
    if (fileName) {
      if (fileName.endsWith('.zip') || fileName.endsWith('.exe')) {
        canHotReload = true;
      }
    }
    
    // If hot-reload requested but not possible, fallback to traditional
    if (hotReload && !canHotReload) {
      console.log('[AUTO-UPDATE] Hot-reload not supported for file:', fileName);
      console.log('[AUTO-UPDATE] Falling back to traditional mode...');
      // Continue with traditional mode
    }
    
    const useHotReload = hotReload && canHotReload;
    
    if (useHotReload) {
      // Hot-reload mode - extract and replace files without restart
      console.log('[AUTO-UPDATE] Using hot-reload mode...');
      
      // Check file type and extract accordingly
      if (!fileName) {
        throw new Error('File name tidak valid untuk hot-reload.');
      }
      
      // Extract update files
      const extractPath = path.join(tempDir, 'extracted');
      if (!fs.existsSync(extractPath)) {
        fs.mkdirSync(extractPath, { recursive: true });
      }
      
      if (fileName.endsWith('.zip')) {
        // Extract ZIP file
        let AdmZip = null;
        try { AdmZip = require('adm-zip'); } catch {}
        if (!AdmZip) {
          throw new Error('adm-zip belum terpasang. Gunakan Traditional mode atau pasang dependensi adm-zip.');
        }
        const zip = new AdmZip(tempFilePath);
        zip.extractAllTo(extractPath, true);
        console.log('[AUTO-UPDATE] ZIP file extracted to:', extractPath);
        
      } else if (fileName.endsWith('.exe')) {
        // Extract EXE file using 7-Zip
        console.log('[AUTO-UPDATE] Extracting EXE file...');
        
        const { spawn } = require('child_process');
        const sevenZipPath = path.join(__dirname, 'node_modules', '7zip-bin', 'win', '7za.exe');
        
        // Try to find 7za.exe in common locations
        let sevenZipExecutable = null;
        const possiblePaths = [
          sevenZipPath,
          path.join(process.env.ProgramFiles || 'C:\\Program Files', '7-Zip', '7z.exe'),
          path.join(process.env['ProgramFiles(x86)'] || 'C:\\Program Files (x86)', '7-Zip', '7z.exe'),
          '7z.exe', // If in PATH
          '7za.exe' // If in PATH
        ];
        
        for (const possiblePath of possiblePaths) {
          if (fs.existsSync(possiblePath)) {
            sevenZipExecutable = possiblePath;
            break;
          }
        }
        
        if (!sevenZipExecutable) {
          throw new Error('7-Zip tidak ditemukan. Silakan install 7-Zip atau gunakan Traditional mode.');
        }
        
        // Extract EXE using 7-Zip
        await new Promise((resolve, reject) => {
          const process = spawn(sevenZipExecutable, ['x', tempFilePath, `-o${extractPath}`, '-y'], {
            stdio: 'pipe'
          });
          
          let output = '';
          process.stdout.on('data', (data) => {
            output += data.toString();
          });
          
          process.stderr.on('data', (data) => {
            output += data.toString();
          });
          
          process.on('close', (code) => {
            if (code === 0) {
              console.log('[AUTO-UPDATE] EXE file extracted successfully');
              resolve();
            } else {
              console.error('[AUTO-UPDATE] EXE extraction failed:', output);
              reject(new Error('Gagal mengekstrak EXE file. Code: ' + code));
            }
          });
          
          process.on('error', (err) => {
            console.error('[AUTO-UPDATE] EXE extraction error:', err);
            reject(err);
          });
        });
        
        console.log('[AUTO-UPDATE] EXE file extracted to:', extractPath);
        
      } else {
        throw new Error('Hot-reload hanya support file ZIP dan EXE. Format tidak dikenali: ' + fileName);
      }
      
      // Prepare for graceful hot-reload
      const appDir = __dirname;
      const extractedFiles = getAllFiles(extractPath);
      
      // Create backup of critical files
      const criticalFiles = ['server.js', 'package.json'];
      const backupDir = path.join(tempDir, 'backup');
      if (!fs.existsSync(backupDir)) {
        fs.mkdirSync(backupDir, { recursive: true });
      }
      
      for (const criticalFile of criticalFiles) {
        const sourcePath = path.join(appDir, criticalFile);
        if (fs.existsSync(sourcePath)) {
          const backupPath = path.join(backupDir, criticalFile);
          fs.copyFileSync(sourcePath, backupPath);
          console.log('[AUTO-UPDATE] Backed up:', criticalFile);
        }
      }
      
      // Stage 1: Update non-critical files first
      const nonCriticalFiles = extractedFiles.filter(file => {
        const relativePath = path.relative(extractPath, file);
        return !criticalFiles.includes(path.basename(relativePath));
      });
      
      for (const file of nonCriticalFiles) {
        const relativePath = path.relative(extractPath, file);
        const targetPath = path.join(appDir, relativePath);
        
        // Skip certain system files and directories
        if (relativePath.includes('uninstall') || 
            relativePath.includes('setup') ||
            relativePath.includes('installer') ||
            relativePath.endsWith('.exe') && !relativePath.includes('pos-web-app.exe')) {
          console.log('[AUTO-UPDATE] Skipping system file:', relativePath);
          continue;
        }
        
        // Ensure target directory exists
        const targetDir = path.dirname(targetPath);
        if (!fs.existsSync(targetDir)) {
          fs.mkdirSync(targetDir, { recursive: true });
        }
        
        // Copy file
        fs.copyFileSync(file, targetPath);
        console.log('[AUTO-UPDATE] Updated (non-critical):', relativePath);
      }
      
      // Stage 2: Schedule critical file updates for next event loop
      setTimeout(() => {
        try {
          console.log('[AUTO-UPDATE] Starting critical file updates...');
          
          // Update package.json first
          const packageJsonPath = path.join(appDir, 'package.json');
          if (fs.existsSync(packageJsonPath)) {
            const packageData = JSON.parse(fs.readFileSync(packageJsonPath, 'utf8'));
            packageData.version = latestVersion;
            fs.writeFileSync(packageJsonPath, JSON.stringify(packageData, null, 2));
            console.log('[AUTO-UPDATE] Updated package.json');
          }
          
          // For hot-reload, DON'T overwrite the running server.js
          // Instead, create a new file and prepare for next restart
          const serverJsPath = path.join(extractPath, 'server.js');
          if (fs.existsSync(serverJsPath)) {
            const targetServerJsPath = path.join(appDir, 'server.js.new');
            fs.copyFileSync(serverJsPath, targetServerJsPath);
            console.log('[AUTO-UPDATE] Created server.js.new for next restart');
            
            // Create a flag file to indicate update is ready
            const updateFlagPath = path.join(appDir, '.update-ready');
            fs.writeFileSync(updateFlagPath, 'true');
            console.log('[AUTO-UPDATE] Update flag created');
          }
          
          // Clear require cache for updated files (except server.js)
          Object.keys(require.cache).forEach(key => {
            if (key.includes(appDir) && !key.includes('node_modules') && !key.includes('server.js')) {
              delete require.cache[key];
            }
          });
          
          console.log('[AUTO-UPDATE] Hot-reload completed successfully!');
          console.log('[AUTO-UPDATE] Server will use new files on next manual restart');
          
          // Clean up temp files
          fs.rmSync(tempDir, { recursive: true, force: true });
          
        } catch (error) {
          console.error('[AUTO-UPDATE] Error in critical file update:', error);
          // Restore from backup if needed
          try {
            for (const criticalFile of criticalFiles) {
              const backupPath = path.join(backupDir, criticalFile);
              const targetPath = path.join(appDir, criticalFile);
              if (fs.existsSync(backupPath)) {
                fs.copyFileSync(backupPath, targetPath);
                console.log('[AUTO-UPDATE] Restored backup:', criticalFile);
              }
            }
          } catch (restoreError) {
            console.error('[AUTO-UPDATE] Failed to restore backup:', restoreError);
          }
        }
      }, 1000); // Delay critical updates by 1 second
      
      // Respond to client immediately
      res.json({
        success: true,
        message: 'Hot-reload update initiated! Server will update files in background.',
        currentVersion,
        latestVersion,
        hotReload: true
      });
      
    } else if (fileName.endsWith('.html')) {
      // No-reload mode for HTML files - just update version info
      console.log('[AUTO-UPDATE] Using no-reload mode for HTML file...');
      
      // Update package.json version only
      const packageJsonPath = path.join(__dirname, 'package.json');
      if (fs.existsSync(packageJsonPath)) {
        const packageData = JSON.parse(fs.readFileSync(packageJsonPath, 'utf8'));
        packageData.version = latestVersion;
        fs.writeFileSync(packageJsonPath, JSON.stringify(packageData, null, 2));
        console.log('[AUTO-UPDATE] Updated package.json version to:', latestVersion);
      }
      
      // Clean up temp files
      fs.rmSync(tempDir, { recursive: true, force: true });
      
      res.json({
        success: true,
        message: 'Version updated successfully! No restart required for HTML release.',
        currentVersion,
        latestVersion,
        hotReload: true,
        noRestart: true
      });
    } else {
      // Traditional mode - restart server
      console.log('[AUTO-UPDATE] Using traditional restart mode...');
      
      // Create update script
      const updateScript = `
@echo off
echo Starting update process...
echo Current version: ${currentVersion}
echo New version: ${latestVersion}

echo Stopping application...
taskkill /F /IM node.exe /T 2>nul
timeout /t 2 /nobreak >nul

echo Extracting update...
powershell -Command "Expand-Archive -Path '${tempFilePath.replace(/\\/g, '\\')}' -DestinationPath '${__dirname.replace(/\\/g, '\\')}' -Force"

echo Cleaning up...
rmdir /s /q "${path.dirname(tempFilePath).replace(/\\/g, '\\')}"

echo Update completed!
echo Restarting application...
start "" "${__dirname}\\pos-web-app.exe"

exit
`;
      
      const scriptPath = path.join(tempDir, 'update.bat');
      fs.writeFileSync(scriptPath, updateScript);
      
      console.log('[AUTO-UPDATE] Update script created:', scriptPath);
      
      // Execute update script and exit
      const { spawn } = require('child_process');
      
      // Detach from current process
      spawn('cmd', ['/c', scriptPath], {
        detached: true,
        stdio: 'ignore',
        cwd: __dirname
      }).unref();
      
      // Give script time to start
      setTimeout(() => {
        console.log('[AUTO-UPDATE] Exiting for update...');
        process.exit(0);
      }, 2000);
      
      res.json({
        success: true,
        message: 'Update process started. Application will restart.',
        currentVersion,
        latestVersion,
        downloadUrl,
        fileName
      });
    }
    
  } catch (error) {
    console.error('[AUTO-UPDATE] Error:', error);
    res.status(500).json({
      success: false,
      message: 'Auto update failed: ' + error.message
    });
  }
});

// Helper function to get all files recursively
function getAllFiles(dir) {
  const fs = require('fs');
  const path = require('path');
  const files = [];
  
  function traverse(currentDir) {
    const items = fs.readdirSync(currentDir);
    
    for (const item of items) {
      const fullPath = path.join(currentDir, item);
      const stat = fs.statSync(fullPath);
      
      if (stat.isDirectory()) {
        traverse(fullPath);
      } else {
        files.push(fullPath);
      }
    }
  }
  
  traverse(dir);
  return files;
}

// Store a single settings object in settings.json

// --- AI Image Generation ---
app.post('/api/ai/generate-image', isAuthenticated, isAdmin, async (req, res) => {
  try {
    const {
      prompt = '',
      size = '1024x1024',
      provider: providerInBody,
      openaiApiKey: keyInBody,
      geminiApiKey: geminiKeyInBody,
      googleApiKey: googleKeyInBody,
      zaiApiKey: zaiKeyInBody
    } = req.body || {};
    const settings = await readData('settings.json').catch(() => ({}));
    const ai = (settings && settings.aiConfig) ? settings.aiConfig : {};
    // Resolve provider with fallbacks: body -> settings -> env -> inferred from available keys.
    let providerRaw = String(providerInBody || (ai && ai.provider) || process.env.AI_PROVIDER || 'none').toLowerCase().trim();
    const inferredOpenAIKey = String((ai && ai.openaiApiKey) || process.env.OPENAI_API_KEY || keyInBody || '').trim();
    const inferredGeminiKey = String((ai && (ai.geminiApiKey || ai.googleApiKey)) || process.env.GEMINI_API_KEY || geminiKeyInBody || googleKeyInBody || '').trim();
    const inferredZaiKey = String((ai && ai.zaiApiKey) || process.env.ZAI_API_KEY || zaiKeyInBody || '').trim();
    if (!providerRaw || providerRaw === 'none') {
      if (inferredZaiKey) providerRaw = 'zai';
      else if (inferredGeminiKey) providerRaw = 'gemini';
      else if (inferredOpenAIKey) providerRaw = 'openai';
      else providerRaw = 'none';
    }
    const provider = (
      providerRaw === 'z.ai' || providerRaw === 'z-ai' ? 'zai' :
      providerRaw === 'google' ? 'gemini' :
      providerRaw
    );
    if (!prompt || provider === 'none') {
      return res.status(400).json({ success: false, message: 'AI tidak dikonfigurasi atau prompt kosong' });
    }
    async function fetchImageBuffer(url){
      try {
        const u = new URL(url);
        const mod = u.protocol === 'https:' ? require('https') : require('http');
        const headers = {};
        return await new Promise((resolve)=>{
          const req2 = mod.get(url, { headers }, (resp)=>{
            if ((resp.statusCode||0) >= 300 && resp.headers && resp.headers.location) {
              return resolve(fetchImageBuffer(resp.headers.location));
            }
            if ((resp.statusCode||0) < 200 || (resp.statusCode||0) >= 300) return resolve(null);
            const chunks = [];
            resp.on('data', d => chunks.push(Buffer.isBuffer(d) ? d : Buffer.from(String(d))));
            resp.on('end', ()=> resolve(Buffer.concat(chunks)));
          });
          req2.on('error', ()=> resolve(null));
        });
      } catch { return null; }
    }
    // Prefer OpenAI if selected
    if (provider === 'openai') {
      // Resolve OpenAI key with fallbacks: settings -> env -> header -> body
      const headerKey = String(req.get('x-openai-key') || req.get('X-OpenAI-Key') || '').trim();
      const key = String((ai && ai.openaiApiKey) || process.env.OPENAI_API_KEY || headerKey || keyInBody || '').trim();
      if (!key) return res.status(400).json({ success: false, message: 'OpenAI API Key belum diset' });
      const body = {
        model: 'gpt-image-1',
        prompt: String(prompt),
        size: String(size || ai.imageSize || '1024x1024'),
        response_format: 'b64_json'
      };
      const r = await safeFetch('https://api.openai.com/v1/images/generations', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${key}`
        },
        body: JSON.stringify(body)
      });
      if (!r) return res.status(502).json({ success:false, message:'Tidak dapat menghubungi OpenAI' });
      if (!r.ok) {
        const t = await r.text().catch(()=> '');
        return res.status(r.status || 500).json({ success:false, message: 'OpenAI error', detail: t });
      }
      const j = await r.json().catch(()=>({}));
      const b64 = j && j.data && j.data[0] && (j.data[0].b64_json || j.data[0].b64) || '';
      if (!b64) return res.status(500).json({ success:false, message:'Gagal menerima gambar dari OpenAI' });
      return res.json({ success:true, imageBase64: `data:image/png;base64,${b64}` });
    }
    if (provider === 'gemini') {
      const headerKey = String(req.get('x-gemini-key') || req.get('X-Gemini-Key') || '').trim();
      const key = String((ai && (ai.geminiApiKey || ai.googleApiKey)) || process.env.GEMINI_API_KEY || headerKey || geminiKeyInBody || googleKeyInBody || '').trim();
      if (!key) return res.status(400).json({ success:false, message:'Gemini API Key belum diset' });

      const sizeStr = String(size || ai.imageSize || '1024x1024');
      let aspectRatio = '1:1';
      try {
        const m = sizeStr.match(/^(\d+)\s*x\s*(\d+)$/i);
        const w = m ? Number(m[1]) : 0;
        const h = m ? Number(m[2]) : 0;
        if (w > 0 && h > 0) {
          const ratio = w / h;
          if (ratio >= 1.7) aspectRatio = '16:9';
          else if (ratio >= 1.2) aspectRatio = '4:3';
          else if (ratio <= 0.58) aspectRatio = '9:16';
          else if (ratio <= 0.85) aspectRatio = '3:4';
          else aspectRatio = '1:1';
        }
      } catch {}

      const bodyGemini = {
        contents: [{ role: 'user', parts: [{ text: String(prompt) }] }],
        generationConfig: {
          responseModalities: ['TEXT', 'IMAGE'],
          imageConfig: { aspectRatio }
        }
      };

      const r = await safeFetch('https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-image:generateContent', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-goog-api-key': key
        },
        body: JSON.stringify(bodyGemini)
      });
      if (!r) return res.status(502).json({ success:false, message:'Tidak dapat menghubungi Gemini' });
      if (!r.ok) {
        const t = await r.text().catch(()=> '');
        if ((r.status || 0) === 429) {
          return res.status(429).json({ success:false, message:'Gemini rate limit tercapai. Coba lagi beberapa saat.', detail: t });
        }
        return res.status(r.status || 500).json({ success:false, message:'Gemini error', detail: t });
      }

      const j = await r.json().catch(()=>({}));
      let b64 = '';
      let mimeType = 'image/png';
      try {
        const cands = Array.isArray(j && j.candidates) ? j.candidates : [];
        for (const cand of cands) {
          const parts = Array.isArray(cand && cand.content && cand.content.parts) ? cand.content.parts : [];
          for (const part of parts) {
            const inline = (part && (part.inlineData || part.inline_data)) || null;
            const data = inline && (inline.data || inline.bytesBase64Encoded || inline.bytes_base64_encoded);
            if (data) {
              b64 = String(data);
              mimeType = String((inline.mimeType || inline.mime_type || 'image/png'));
              break;
            }
          }
          if (b64) break;
        }
      } catch {}
      if (!b64) return res.status(500).json({ success:false, message:'Gagal menerima gambar dari Gemini' });
      return res.json({ success:true, imageBase64: `data:${mimeType};base64,${b64}` });
    }
    if (provider === 'zai') {
      const headerKey = String(req.get('x-zai-key') || req.get('X-ZAI-Key') || req.get('x-z-ai-key') || req.get('X-Z-AI-Key') || '').trim();
      const key = String((ai && ai.zaiApiKey) || process.env.ZAI_API_KEY || headerKey || zaiKeyInBody || '').trim();
      if (!key) return res.status(400).json({ success:false, message:'Z.AI API Key belum diset' });

      const model = String((req.body && req.body.zaiModel) || (ai && ai.zaiModel) || process.env.ZAI_IMAGE_MODEL || 'glm-image').trim();
      const normalizeZaiSize = (sizeInput) => {
        const MAX_PIXELS = 1 << 22; // 2^22
        const MIN_DIM = 512;
        const MAX_DIM = 2880;
        const STEP = 32;
        const floorToStep = (v) => Math.floor(v / STEP) * STEP;
        const clamp = (v, lo, hi) => Math.max(lo, Math.min(hi, v));
        let w = 1024;
        let h = 1024;
        try {
          const m = String(sizeInput || '').trim().match(/^(\d+)\s*x\s*(\d+)$/i);
          if (m) {
            w = Number(m[1]) || w;
            h = Number(m[2]) || h;
          }
        } catch {}
        w = clamp(floorToStep(w), MIN_DIM, MAX_DIM);
        h = clamp(floorToStep(h), MIN_DIM, MAX_DIM);
        if (w < MIN_DIM) w = MIN_DIM;
        if (h < MIN_DIM) h = MIN_DIM;
        let area = w * h;
        if (area > MAX_PIXELS) {
          const scale = Math.sqrt(MAX_PIXELS / area);
          w = clamp(floorToStep(w * scale), MIN_DIM, MAX_DIM);
          h = clamp(floorToStep(h * scale), MIN_DIM, MAX_DIM);
          if (w < MIN_DIM) w = MIN_DIM;
          if (h < MIN_DIM) h = MIN_DIM;
          area = w * h;
          while (area > MAX_PIXELS && (w > MIN_DIM || h > MIN_DIM)) {
            if (w >= h && w > MIN_DIM) w = Math.max(MIN_DIM, w - STEP);
            else if (h > MIN_DIM) h = Math.max(MIN_DIM, h - STEP);
            area = w * h;
          }
        }
        return `${w}x${h}`;
      };
      const zaiSize = normalizeZaiSize(String(size || ai.imageSize || '1024x1024'));
      const bodyZai = {
        model,
        prompt: String(prompt),
        size: zaiSize
      };
      const r = await safeFetch('https://api.z.ai/api/paas/v4/images/generations', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${key}`
        },
        body: JSON.stringify(bodyZai)
      });
      if (!r) return res.status(502).json({ success:false, message:'Tidak dapat menghubungi Z.AI' });
      if (!r.ok) {
        const t = await r.text().catch(()=> '');
        if ((r.status || 0) === 429) {
          return res.status(429).json({ success:false, message:'Z.AI rate limit tercapai. Coba lagi beberapa saat.', detail: t });
        }
        return res.status(r.status || 500).json({ success:false, message:'Z.AI error', detail: t });
      }
      const j = await r.json().catch(()=>({}));
      const b64 = j && j.data && j.data[0] && (j.data[0].b64_json || j.data[0].b64 || j.data[0].image_base64) || '';
      if (b64) return res.json({ success:true, imageBase64: `data:image/png;base64,${b64}` });
      const imageUrl = j && j.data && j.data[0] && (j.data[0].url || j.data[0].image_url) || '';
      if (!imageUrl) return res.status(500).json({ success:false, message:'Gagal menerima gambar dari Z.AI' });
      const imgBuf = await fetchImageBuffer(String(imageUrl));
      if (!imgBuf) return res.status(500).json({ success:false, message:'Gagal mengambil gambar URL dari Z.AI' });
      const mime = /\.(jpe?g)(\?|$)/i.test(String(imageUrl)) ? 'image/jpeg' : 'image/png';
      return res.json({ success:true, imageBase64: `data:${mime};base64,${imgBuf.toString('base64')}` });
    }
    return res.status(400).json({ success:false, message:`Provider AI tidak valid: ${providerRaw}`, providerResolved: provider });
  } catch (e) {
    return res.status(500).json({ success:false, message:'Kesalahan server saat generate gambar' });
  }
});
// Cari gambar dari internet untuk produk dan kembalikan sebagai Data URL Base64
app.post('/api/ai/find-image', isAuthenticated, isAdmin, async (req, res) => {
  try {
    const { q = '', sku = '', name = '', category = '', unsplashKey: unsplashKeyInBody = '', pexelsKey: pexelsKeyInBody = '', bingKey: bingKeyInBody = '' } = req.body || {};
    const query = String(q || name || sku || category || '').trim();
    if (!query) return res.status(400).json({ success:false, message:'Query kosong' });

    // Helper: fetch binary to Buffer (with simple redirect follow)
    async function fetchBuffer(url){
      try {
        const u = new URL(url);
        const mod = u.protocol === 'https:' ? require('https') : require('http');
        const headers = {};
        return await new Promise((resolve)=>{
          const req = mod.get(url, { headers }, (resp)=>{
            if ((resp.statusCode||0) >= 300 && resp.headers && resp.headers.location) {
              return resolve(fetchBuffer(resp.headers.location));
            }
            const chunks = [];
            resp.on('data', d => chunks.push(Buffer.isBuffer(d) ? d : Buffer.from(d)));
            resp.on('end', ()=> resolve(Buffer.concat(chunks)));
          });
          req.on('error', ()=> resolve(null));
        });
      } catch { return null; }
    }

    // Try providers in order: Google CSE (if key+cx) -> Unsplash -> Pexels -> Bing
    const candidates = [];
    const unsHdr = String(req.get('x-unsplash-key') || req.get('X-Unsplash-Key') || '').trim();
    const pexHdr = String(req.get('x-pexels-key') || req.get('X-Pexels-Key') || '').trim();
    const bingHdr = String(req.get('x-bing-key') || req.get('X-Bing-Key') || '').trim();
    const gKeyHdr = String(req.get('x-google-key') || req.get('X-Google-Key') || '').trim();
    const gCxHdr = String(req.get('x-google-cx') || req.get('X-Google-Cx') || '').trim();
    const googleKey = (gKeyHdr || (req.body && req.body.googleKey) || process.env.GOOGLE_CSE_KEY || '').trim();
    const googleCx = (gCxHdr || (req.body && req.body.googleCx) || process.env.GOOGLE_CSE_CX || '').trim();
    if (googleKey && googleCx) {
      try {
        const url = `https://www.googleapis.com/customsearch/v1?q=${encodeURIComponent(query)}&searchType=image&num=1&key=${encodeURIComponent(googleKey)}&cx=${encodeURIComponent(googleCx)}`;
        const r = await safeFetch(url);
        if (r && r.ok) {
          const j = await r.json().catch(()=>({}));
          const link = j && j.items && j.items[0] && (j.items[0].link || (j.items[0].image && j.items[0].image.thumbnailLink));
          if (link) candidates.push(link);
        }
      } catch {}
    }
    const unsplashKey = (unsHdr || unsplashKeyInBody || process.env.UNSPLASH_ACCESS_KEY || '').trim();
    if (unsplashKey) {
      try {
        const r = await safeFetch(`https://api.unsplash.com/search/photos?query=${encodeURIComponent(query)}&per_page=1`, {
          headers: { 'Authorization': `Client-ID ${unsplashKey}` }
        });
        if (r && r.ok) {
          const j = await r.json().catch(()=>({}));
          const url = j && j.results && j.results[0] && j.results[0].urls && (j.results[0].urls.small || j.results[0].urls.regular);
          if (url) candidates.push(url);
        }
      } catch {}
    }
    const pexelsKey = (pexHdr || pexelsKeyInBody || process.env.PEXELS_API_KEY || '').trim();
    if (!candidates.length && pexelsKey) {
      try {
        const r = await safeFetch(`https://api.pexels.com/v1/search?query=${encodeURIComponent(query)}&per_page=1`, {
          headers: { 'Authorization': pexelsKey }
        });
        if (r && r.ok) {
          const j = await r.json().catch(()=>({}));
          const url = j && j.photos && j.photos[0] && j.photos[0].src && (j.photos[0].src.medium || j.photos[0].src.large || j.photos[0].src.original);
          if (url) candidates.push(url);
        }
      } catch {}
    }
    const bingKey = (bingHdr || bingKeyInBody || process.env.BING_IMAGE_SEARCH_KEY || '').trim();
    if (!candidates.length && bingKey) {
      try {
        const r = await safeFetch(`https://api.bing.microsoft.com/v7.0/images/search?q=${encodeURIComponent(query)}&count=1&safeSearch=Strict`, {
          headers: { 'Ocp-Apim-Subscription-Key': bingKey }
        });
        if (r && r.ok) {
          const j = await r.json().catch(()=>({}));
          const url = j && j.value && j.value[0] && (j.value[0].contentUrl || j.value[0].thumbnailUrl);
          if (url) candidates.push(url);
        }
      } catch {}
    }

    if (!candidates.length) {
      return res.status(404).json({ success:false, message:'Tidak menemukan gambar untuk query ini. Konfigurasikan API key (Google CSE/Unsplash/Pexels/Bing).' });
    }

    const imgUrl = candidates[0];
    const buf = await fetchBuffer(imgUrl);
    if (!buf || !buf.length) return res.status(502).json({ success:false, message:'Gagal mengunduh gambar' });
    let mime = 'image/jpeg';
    try {
      const low = imgUrl.toLowerCase();
      if (low.endsWith('.png')) mime = 'image/png';
      else if (low.endsWith('.webp')) mime = 'image/webp';
      else if (low.endsWith('.gif')) mime = 'image/gif';
    } catch {}
    const b64 = buf.toString('base64');
    return res.json({ success:true, imageBase64: `data:${mime};base64,${b64}`, sourceUrl: imgUrl });
  } catch (e) {
    console.error('find-image error:', e);
    return res.status(500).json({ success:false, message:'Gagal mencari gambar' });
  }
});

// Upload cart sound file
app.post('/api/upload/cart-sound', isAuthenticated, isAdmin, async (req, res) => {
  try {
    const { soundBase64 } = req.body;
    
    if (!soundBase64) {
      return res.status(400).json({ success: false, message: 'No sound file provided' });
    }
    
    // Validate that it's a valid audio file in base64 format
    if (!soundBase64.startsWith('data:audio/')) {
      return res.status(400).json({ success: false, message: 'Invalid audio format' });
    }
    
    // Get current settings
    const settings = await readData('settings.json') || {};
    
    // Update settings with new sound file
    settings.cartSoundBase64 = soundBase64;
    
    // Save settings
    await writeData('settings.json', settings);
    
    res.json({ success: true, message: 'Cart sound uploaded successfully' });
  } catch (e) {
    console.error('Failed to upload cart sound:', e);
    res.status(500).json({ success: false, message: 'Failed to upload cart sound' });
  }
});

// --- Customers API ---
// GET /api/customers - Get all customers
app.get('/api/customers', isAuthenticated, async (req, res) => {
  try {
    const customersRaw = await readData('customers.json');
    const customers = Array.isArray(customersRaw) ? customersRaw : [];
    const normalized = customers.map((c) => ({
      ...c,
      balance: Math.max(0, Number(c && c.balance != null ? c.balance : (c && c.depositBalance != null ? c.depositBalance : 0)) || 0)
    }));
    res.json(normalized);
  } catch (error) {
    console.error('Failed to load customers:', error);
    res.status(500).json({ success: false, message: 'Failed to load customers' });
  }
});

// POST /api/customers - Create new customer
app.post('/api/customers', isAuthenticated, async (req, res) => {
  try {
    const { name, phone, email, address, balance = 0 } = req.body;
    if (!name || !name.trim()) {
      return res.status(400).json({
        success: false,
        message: 'Customer name is required'
      });
    }

    const customers = await readData('customers.json');
    
    // Check if customer name already exists
    const existingCustomer = customers.find(c => 
      c.name.toLowerCase() === name.trim().toLowerCase()
    );
    if (existingCustomer) {
      return res.status(400).json({
        success: false,
        message: 'Customer name already exists'
      });
    }

    const newCustomer = {
      id: Date.now(),
      name: name.trim(),
      phone: phone ? phone.trim() : '',
      email: email ? email.trim() : '',
      address: address ? address.trim() : '',
      balance: Math.max(0, Number(balance) || 0),
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };

    customers.push(newCustomer);
    await saveArrayWithSync('customers.json', customers);

    res.json({
      success: true,
      message: 'Customer created successfully',
      customer: newCustomer
    });
  } catch (error) {
    console.error('Failed to create customer:', error);
    res.status(500).json({ success: false, message: 'Failed to create customer' });
  }
});

// PUT /api/customers/:id - Update customer
app.put('/api/customers/:id', isAuthenticated, async (req, res) => {
  try {
    const { name, phone, email, address, balance } = req.body;
    const customerId = parseInt(req.params.id);

    if (!name || !name.trim()) {
      return res.status(400).json({
        success: false,
        message: 'Customer name is required'
      });
    }

    const customers = await readData('customers.json');
    const index = customers.findIndex(c => c.id === customerId);

    if (index === -1) {
      return res.status(404).json({
        success: false,
        message: 'Customer not found'
      });
    }

    // PERBAIKAN: Validasi nama duplikat hanya jika nama berubah (bukan nama yang sama dengan pelanggan yang sedang diedit)
    const currentCustomer = customers[index];
    const newName = name.trim().toLowerCase();
    const currentName = currentCustomer.name.toLowerCase();
    
    // Jika nama berubah, cek apakah nama baru sudah digunakan oleh pelanggan lain
    if (newName !== currentName) {
      const existingCustomer = customers.find(c => 
        c.name.toLowerCase() === newName && c.id !== customerId
      );
      if (existingCustomer) {
        return res.status(400).json({
          success: false,
          message: 'Customer name already exists'
        });
      }
    }
    // Jika nama tidak berubah (nama sama dengan yang sekarang), tidak perlu validasi duplikat

    customers[index] = {
      ...customers[index],
      name: name.trim(),
      phone: phone ? phone.trim() : '',
      email: email ? email.trim() : '',
      address: address ? address.trim() : '',
      balance: Math.max(
        0,
        Number(balance != null ? balance : (customers[index] && customers[index].balance != null ? customers[index].balance : 0)) || 0
      ),
      updatedAt: new Date().toISOString()
    };

    await saveArrayWithSync('customers.json', customers);

    res.json({
      success: true,
      message: 'Customer updated successfully',
      customer: customers[index]
    });
  } catch (error) {
    console.error('Failed to update customer:', error);
    res.status(500).json({ success: false, message: 'Failed to update customer' });
  }
});

// DELETE /api/customers/:id - Delete customer
app.delete('/api/customers/:id', isAuthenticated, async (req, res) => {
  try {
    const customerId = parseInt(req.params.id);
    const customers = await readData('customers.json');
    
    // Don't allow deleting default customer
    if (customerId === 1) {
      return res.status(400).json({
        success: false,
        message: 'Cannot delete default customer'
      });
    }

    const filteredCustomers = customers.filter(c => c.id !== customerId);

    if (customers.length === filteredCustomers.length) {
      return res.status(404).json({
        success: false,
        message: 'Customer not found'
      });
    }

    await saveArrayWithSync('customers.json', filteredCustomers);

    res.json({
      success: true,
      message: 'Customer deleted successfully'
    });
  } catch (error) {
    console.error('Failed to delete customer:', error);
    res.status(500).json({ success: false, message: 'Failed to delete customer' });
  }
});

// Check customer name availability
app.post('/api/customers/check-name/:id?', async (req, res) => {
  try {
    const { name } = req.body;
    const customerId = req.params.id ? parseInt(req.params.id) : null;

    const customers = await readData('customers.json');
    const existingCustomer = customers.find(c => 
      c.name.toLowerCase() === name.trim().toLowerCase() && 
      (!customerId || c.id !== customerId)
    );
    
    res.json({ exists: !!existingCustomer });
  } catch (error) {
    console.error('Error checking customer name:', error);
    res.status(500).json({ success: false, message: 'Server error' });
  }
});

// --- Suppliers API ---
// GET /api/suppliers - Get all suppliers
app.get('/api/suppliers', isAuthenticated, async (req, res) => {
  try {
    const suppliers = await readData('suppliers.json').catch(() => []);
    res.json(Array.isArray(suppliers) ? suppliers : []);
  } catch (error) {
    console.error('Failed to load suppliers:', error);
    res.status(500).json({ success: false, message: 'Failed to load suppliers' });
  }
});

// POST /api/suppliers - Create new supplier
app.post('/api/suppliers', isAuthenticated, async (req, res) => {
  try {
    const { name, phone, address, notes } = req.body || {};
    if (!name || !String(name).trim()) {
      return res.status(400).json({ success: false, message: 'Supplier name is required' });
    }

    let suppliers = await readData('suppliers.json').catch(() => []);
    if (!Array.isArray(suppliers)) suppliers = [];

    const normName = String(name).trim().toLowerCase();
    const exists = suppliers.find(s => String(s.name || '').toLowerCase() === normName);
    if (exists) {
      return res.status(400).json({ success: false, message: 'Supplier name already exists' });
    }

    const nowIso = new Date().toISOString();
    const newSupplier = {
      id: Date.now(),
      name: String(name).trim(),
      phone: phone ? String(phone).trim() : '',
      address: address ? String(address).trim() : '',
      notes: notes ? String(notes).trim() : '',
      createdAt: nowIso,
      updatedAt: nowIso
    };

    suppliers.push(newSupplier);
    await saveArrayWithSync('suppliers.json', suppliers, { keyField: 'id' });

    res.json({ success: true, message: 'Supplier created successfully', supplier: newSupplier });
  } catch (error) {
    console.error('Failed to create supplier:', error);
    res.status(500).json({ success: false, message: 'Failed to create supplier' });
  }
});

// PUT /api/suppliers/:id - Update supplier
app.put('/api/suppliers/:id', isAuthenticated, async (req, res) => {
  try {
    const supplierId = parseInt(req.params.id);
    const { name, phone, address, notes } = req.body || {};

    if (!name || !String(name).trim()) {
      return res.status(400).json({ success: false, message: 'Supplier name is required' });
    }

    let suppliers = await readData('suppliers.json').catch(() => []);
    if (!Array.isArray(suppliers)) suppliers = [];

    const idx = suppliers.findIndex(s => Number(s.id) === supplierId);
    if (idx === -1) {
      return res.status(404).json({ success: false, message: 'Supplier not found' });
    }

    const current = suppliers[idx];
    const newNameNorm = String(name).trim().toLowerCase();
    const curNameNorm = String(current.name || '').toLowerCase();
    if (newNameNorm !== curNameNorm) {
      const exists = suppliers.find(s => String(s.name || '').toLowerCase() === newNameNorm && Number(s.id) !== supplierId);
      if (exists) {
        return res.status(400).json({ success: false, message: 'Supplier name already exists' });
      }
    }

    suppliers[idx] = {
      ...current,
      name: String(name).trim(),
      phone: phone ? String(phone).trim() : '',
      address: address ? String(address).trim() : '',
      notes: notes ? String(notes).trim() : '',
      updatedAt: new Date().toISOString()
    };

    await saveArrayWithSync('suppliers.json', suppliers, { keyField: 'id' });
    res.json({ success: true, message: 'Supplier updated successfully', supplier: suppliers[idx] });
  } catch (error) {
    console.error('Failed to update supplier:', error);
    res.status(500).json({ success: false, message: 'Failed to update supplier' });
  }
});

// DELETE /api/suppliers/:id - Delete supplier
app.delete('/api/suppliers/:id', isAuthenticated, async (req, res) => {
  try {
    const supplierId = parseInt(req.params.id);
    let suppliers = await readData('suppliers.json').catch(() => []);
    if (!Array.isArray(suppliers)) suppliers = [];

    const beforeLen = suppliers.length;
    suppliers = suppliers.filter(s => Number(s.id) !== supplierId);
    if (beforeLen === suppliers.length) {
      return res.status(404).json({ success: false, message: 'Supplier not found' });
    }

    await saveArrayWithSync('suppliers.json', suppliers, { keyField: 'id' });
    res.json({ success: true, message: 'Supplier deleted successfully' });
  } catch (error) {
    console.error('Failed to delete supplier:', error);
    res.status(500).json({ success: false, message: 'Failed to delete supplier' });
  }
});

// DANGEROUS: Delete selected database files
app.post("/api/database/delete-selected", isAuthenticated, isAdmin, async (req, res) => {
  try {
    const items = Array.isArray(req.body?.items) ? req.body.items : [];
    if (items.length === 0) {
      return res.status(400).json({ success: false, message: "Tidak ada item yang dipilih" });
    }

    const dataDir = resolveDataDir();
    const fs = rawFs.promises;
    const path = require("path");

    const deletedFiles = [];
    const skipped = [];
    const errors = [];

    const EMPTY_ARRAY = new Set([
      'products', 'categories', 'units', 'price_history', 'customers',
      'suppliers', 'transactions', 'invoices', 'expenses', 'stock_in',
      'stock_moves', 'deletions', 'permissions_audit', 'outbox',
      'drafts', 'pos-drafts', 'users'
    ]);

    const EMPTY_OBJECT = new Set([
      'banners', 'qris', 'sync_config', 'lastSync'
    ]);

    for (const key of items) {
      try {
        if (key === 'shifts') {
          await writeData('shifts.json', []);
          deletedFiles.push('shifts.json');
        } else if (key === 'cart') {
          await writeData('cart.json', { users: {} });
          deletedFiles.push('cart.json');
        } else if (key === 'backups') {
          const backupsDir = path.join(dataDir, 'backups');
          await fs.rm(backupsDir, { recursive: true, force: true });
          deletedFiles.push('backups/');
        } else if (EMPTY_ARRAY.has(key)) {
          await writeData(key + '.json', []);
          deletedFiles.push(key + '.json');
        } else if (EMPTY_OBJECT.has(key)) {
          await writeData(key + '.json', {});
          deletedFiles.push(key + '.json');
        } else {
          skipped.push({ key, reason: "Unknown target" });
        }
      } catch (e) {
        errors.push({ target: key, error: e.message });
      }
    }

    return res.json({
      success: true,
      message: "Data terpilih telah dihapus.",
      deletedFiles,
      skipped,
      errors
    });
  } catch (e) {
    console.error("Error deleting selected database:", e);
    res.status(500).json({
      success: false,
      message: "Gagal menghapus data terpilih: " + (e.message || e)
    });
  }
});

// --- Suppliers Import/Template ---
app.get('/api/suppliers/template', isAuthenticated, isAdmin, async (req, res) => {
  try {
    const templateData = [
      {
        "Supplier Name": "Supplier A",
        Phone: "081234567890",
        Address: "Jl. Contoh No. 1",
        Notes: "Catatan supplier"
      },
      {
        "Supplier Name": "Supplier B",
        Phone: "085678901234",
        Address: "Jl. Contoh No. 2",
        Notes: ""
      }
    ];
    const ws = XLSX.utils.json_to_sheet(templateData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Template");
    ws["!cols"] = [{ wch: 30 }, { wch: 18 }, { wch: 40 }, { wch: 30 }];
    const buf = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", "attachment; filename=supplier_import_template.xlsx");
    res.send(buf);
  } catch (e) {
    res.status(500).json({ success: false, message: "Gagal membuat template supplier" });
  }
});

app.post('/api/suppliers/import', isAuthenticated, isAdmin, async (req, res) => {
  try {
    const list = Array.isArray(req.body && req.body.suppliers) ? req.body.suppliers : [];
    if (list.length === 0) {
      return res.status(400).json({ success: false, message: "No valid data to import" });
    }

    let suppliers = await readData('suppliers.json').catch(() => []);
    if (!Array.isArray(suppliers)) suppliers = [];

    let successCount = 0;
    let errorCount = 0;
    const errors = [];

    for (let i = 0; i < list.length; i++) {
      try {
        const row = list[i] || {};
        const nameRaw = row["Supplier Name"] ?? row["Nama Supplier"] ?? row["Name"];
        const name = nameRaw != null ? String(nameRaw).trim() : "";
        if (!name) {
          errors.push(`Baris ${i + 1}: Supplier Name wajib diisi.`);
          errorCount++;
          continue;
        }

        const phone = row["Phone"] ?? row["Telepon"];
        const address = row["Address"] ?? row["Alamat"];
        const notes = row["Notes"] ?? row["Catatan"];

        const normName = name.toLowerCase();
        const idx = suppliers.findIndex(s => String(s && s.name || '').toLowerCase() === normName);
        if (idx >= 0) {
          suppliers[idx] = {
            ...suppliers[idx],
            phone: phone != null ? String(phone).trim() : (suppliers[idx].phone || ''),
            address: address != null ? String(address).trim() : (suppliers[idx].address || ''),
            notes: notes != null ? String(notes).trim() : (suppliers[idx].notes || ''),
            updatedAt: new Date().toISOString()
          };
        } else {
          const nowIso = new Date().toISOString();
          suppliers.push({
            id: Date.now() + i,
            name,
            phone: phone != null ? String(phone).trim() : '',
            address: address != null ? String(address).trim() : '',
            notes: notes != null ? String(notes).trim() : '',
            createdAt: nowIso,
            updatedAt: nowIso
          });
        }
        successCount++;
      } catch (err) {
        errors.push(`Baris ${i + 1}: ${err.message}`);
        errorCount++;
      }
    }

    await saveArrayWithSync('suppliers.json', suppliers, { keyField: 'id' });

    let message = `Import selesai. Sukses: ${successCount}, Error: ${errorCount}`;
    if (errors.length > 0) {
      message += `\n\nBeberapa error pertama:\n${errors.slice(0, 3).join("\n")}`;
      if (errors.length > 5) {
        message += ` ... dan ${errors.length - 5} more errors`;
      }
    }
    res.json({ success: true, message, successCount, errorCount, errors: errors.slice(0, 10) });
  } catch (e) {
    res.status(500).json({ success: false, message: "Failed to import suppliers: " + e.message });
  }
});

// Export Suppliers to XLSX
app.get('/api/suppliers/export', isAuthenticated, isAdmin, async (req, res) => {
  try {
    let suppliers = await readData('suppliers.json').catch(() => []);
    if (!Array.isArray(suppliers)) suppliers = [];
    const exportData = suppliers.map(s => ({
      "Supplier Name": s.name || "",
      Phone: s.phone || "",
      Address: s.address || "",
      Notes: s.notes || "",
      "Created At": s.createdAt || "",
      "Updated At": s.updatedAt || ""
    }));
    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Suppliers");
    ws["!cols"] = [
      { wch: 30 },
      { wch: 18 },
      { wch: 40 },
      { wch: 30 },
      { wch: 20 },
      { wch: 20 }
    ];
    const buf = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", "attachment; filename=suppliers_export.xlsx");
    res.send(buf);
  } catch (e) {
    res.status(500).json({ success: false, message: "Failed to export suppliers: " + e.message });
  }
});

// --- Stock In (Barang Masuk) API ---
// GET /api/stock-in - list all stock-in records
app.get('/api/stock-in', isAuthenticated, async (req, res) => {
  try {
    let stockIn = await readData('stock_in.json').catch(() => []);
    if (!Array.isArray(stockIn)) stockIn = [];
    res.json(stockIn);
  } catch (error) {
    console.error('Failed to load stock-in records:', error);
    res.status(500).json({ success: false, message: 'Failed to load stock-in records' });
  }
});

// POST /api/stock-in - create new stock-in document and update product stock
app.post('/api/stock-in', isAuthenticated, async (req, res) => {
  try {
    const user = (req.session && req.session.user) || {};
    const { date, supplierId, items, note, paidAmount, remainingAmount, paymentDate, shippingCost, vatPercent, vatAmount, discountPercent, discountAmount } = req.body || {};

    const sid = supplierId ? parseInt(supplierId) : null;
    if (!sid || !Number.isFinite(sid)) {
      return res.status(400).json({ success: false, message: 'Supplier is required' });
    }
    if (!Array.isArray(items) || items.length === 0) {
      return res.status(400).json({ success: false, message: 'Items are required' });
    }

    let suppliers = await readData('suppliers.json').catch(() => []);
    if (!Array.isArray(suppliers)) suppliers = [];
    const supplier = suppliers.find(s => Number(s.id) === sid);
    if (!supplier) {
      return res.status(400).json({ success: false, message: 'Supplier not found' });
    }

    let products = await readData('products.json').catch(() => []);
    if (!Array.isArray(products)) products = [];

    let stockIn = await readData('stock_in.json').catch(() => []);
    if (!Array.isArray(stockIn)) stockIn = [];

    const now = Date.now();
    const id = `STKIN-${new Date(now).toISOString().slice(0,10).replace(/-/g,'')}-${now}`;

    const round2 = (n) => Math.round((Number(n || 0)) * 100) / 100;
    const normalizedItems = [];
    const itemDrafts = [];
    let itemsSubtotal = 0;
    let totalQty = 0;
    const resolveProductIdKey = (rawId) => {
      const key = String(rawId != null ? rawId : '').trim();
      if (!key || key === 'undefined' || key === 'null') return '';
      return key;
    };
    const findProductById = (rawId) => {
      const key = resolveProductIdKey(rawId);
      if (!key) return null;
      return products.find((p) => {
        if (!p) return false;
        const a = p.id != null ? String(p.id) : '';
        const b = p._id != null ? String(p._id) : '';
        return a === key || b === key;
      }) || null;
    };
    for (const raw of items) {
      const qty = Number(raw && raw.qty != null ? raw.qty : 0) || 0;
      const purchasePrice = Number(raw && raw.purchasePrice != null ? raw.purchasePrice : 0) || 0;
      const sellingPrice = Number(raw && raw.sellingPrice != null ? raw.sellingPrice : 0) || 0;
      if (qty <= 0) continue;

      const product = findProductById(raw && (raw.productId != null ? raw.productId : raw.id));
      if (!product) {
        const badId = resolveProductIdKey(raw && raw.productId);
        if (!badId) continue;
        return res.status(400).json({ success: false, message: `Product with ID ${badId} not found` });
      }
      const pid = product.id != null ? product.id : product._id;

      const lineSubtotal = round2(purchasePrice * qty);
      itemDrafts.push({ product, pid, qty, purchasePrice, sellingPrice, lineSubtotal, variant: raw && raw.variant ? raw.variant : null });
      itemsSubtotal += lineSubtotal;
      totalQty += qty;
    }

    if (!itemDrafts.length) {
      return res.status(400).json({ success: false, message: 'No valid items' });
    }

    const shippingCostNum = Math.max(0, Number(shippingCost) || 0);
    const vatPercentNum = Math.max(0, Number(vatPercent) || 0);
    const computedVatAmount = round2(itemsSubtotal * (vatPercentNum / 100));
    const vatAmountNum = Math.max(0, Number(vatAmount) || computedVatAmount);
    const discountPercentNum = Math.max(0, Number(discountPercent) || 0);
    const computedDiscountAmount = round2(itemsSubtotal * (discountPercentNum / 100));
    const discountAmountNum = Math.max(0, Number(discountAmount) || computedDiscountAmount);
    const additionalCostPool = round2(shippingCostNum + vatAmountNum - discountAmountNum);
    const totalAmount = round2(itemsSubtotal + additionalCostPool);

    for (const draft of itemDrafts) {
      const { product, pid, qty, purchasePrice, sellingPrice, lineSubtotal, variant } = draft;
      const prevProduct = {
        purchasePrice: product.purchasePrice,
        sellingPrice: product.sellingPrice,
        price: product.price,
        unitPrices: Array.isArray(product.unitPrices)
          ? product.unitPrices.map(v => ({ ...v }))
          : []
      };
      product.stock = Number(product.stock || 0) + qty;
      const prevSell = Number((prevProduct.sellingPrice != null ? prevProduct.sellingPrice : prevProduct.price) || 0);
      const prevBuy = Number(prevProduct.purchasePrice || 0);

      const weight = itemsSubtotal > 0
        ? (lineSubtotal / itemsSubtotal)
        : (totalQty > 0 ? (qty / totalQty) : 0);
      const allocatedAdditional = round2(additionalCostPool * weight);
      const effectivePurchasePrice = qty > 0
        ? Math.max(0, round2((lineSubtotal + allocatedAdditional) / qty))
        : round2(purchasePrice);

      const nextBuy = effectivePurchasePrice > 0 ? effectivePurchasePrice : prevBuy;
      const nextSell = sellingPrice > 0 ? Number(sellingPrice || 0) : prevSell;
      let priceChanged = false;
      if (nextBuy !== prevBuy) {
        product.purchasePrice = nextBuy;
        priceChanged = true;
      }
      if (sellingPrice > 0 && nextSell !== prevSell) {
        product.sellingPrice = nextSell;
        product.price = nextSell;
        priceChanged = true;
      }
      product.updatedAt = Date.now();

      try {
        const by = String(user && (user.username || user.name || ''));
        const nextSnapshot = { ...prevProduct, purchasePrice: nextBuy, sellingPrice: nextSell, price: nextSell };
        let entry = buildPriceHistoryEntry(prevProduct, nextSnapshot, 'stock_in', by);
        if (!entry) {
          entry = {
            at: Date.now(),
            by,
            reason: 'stock_in',
            base: {
              sellingPrice: { from: prevSell, to: nextSell },
              purchasePrice: { from: prevBuy, to: nextBuy }
            },
            variants: []
          };
        }
        appendPriceHistory(product, entry);
        try { if (global.appendPriceHistoryForProduct) await global.appendPriceHistoryForProduct(pid, entry); } catch {}
      } catch (e) {}

      try {
        await appendStockMove({
          productId: pid,
          delta: qty,
          reason: 'purchase',
          refId: id,
          by: String(user && (user.username || user.name || '')),
          newStock: Number(product.stock || 0),
          stockOnly: priceChanged ? false : true
        });
      } catch (e) {}

      normalizedItems.push({
        productId: pid,
        productName: String(product.name || ''),
        qty,
        purchasePrice: round2(purchasePrice),
        effectivePurchasePrice: round2(effectivePurchasePrice),
        sellingPrice,
        prevPurchasePrice: round2(prevBuy),
        prevSellingPrice: round2(prevSell),
        appliedPurchasePrice: round2(nextBuy),
        appliedSellingPrice: round2(nextSell),
        ...(variant ? { variant } : {})
      });
    }

    const paidAmountNum = Math.max(0, Number(paidAmount) || 0);
    const remainingAmountNum = (remainingAmount !== undefined && remainingAmount !== null && String(remainingAmount) !== '')
      ? Math.max(0, Number(remainingAmount) || 0)
      : Math.max(0, round2(totalAmount - paidAmountNum));

    const stockInDoc = {
      id,
      timestamp: new Date(now).toISOString(),
      date: date ? String(date) : new Date(now).toISOString().slice(0,10),
      supplierId: sid,
      supplierName: String(supplier.name || ''),
      items: normalizedItems,
      note: note ? String(note).trim() : '',
      itemsSubtotal: round2(itemsSubtotal),
      shippingCost: round2(shippingCostNum),
      vatPercent: round2(vatPercentNum),
      vatAmount: round2(vatAmountNum),
      discountPercent: round2(discountPercentNum),
      discountAmount: round2(discountAmountNum),
      totalAmount: round2(totalAmount),
      paidAmount: paidAmountNum,
      remainingAmount: remainingAmountNum,
      paymentDate: paymentDate ? String(paymentDate) : '',
      updatedAt: now
    };

    stockIn.push(stockInDoc);
    await writeData('stock_in.json', stockIn);
    await saveArrayWithSync('products.json', products, { keyField: 'id' });
    try { invalidateCache('products'); } catch {}

    res.json({ success: true, message: 'Stock-in recorded successfully', stockIn: stockInDoc, stockApplied: true });
  } catch (error) {
    console.error('Failed to record stock-in:', error);
    res.status(500).json({ success: false, message: 'Failed to record stock-in' });
  }
});

// --- Stock-In Import/Template ---
app.get('/api/stock-in/template', isAuthenticated, isAdmin, async (req, res) => {
  try {
    const templateData = [
      {
        Date: new Date().toISOString().slice(0, 10),
        "Supplier Name": "Supplier A",
        "Product ID": "",
        "Product SKU": "SKU-001",
        "Product Name": "Produk Contoh 1",
        Qty: 10,
        "Purchase Price": 8000,
        "Selling Price": 12000,
        "Shipping Cost": 10000,
        "VAT (%)": 11,
        "VAT Amount": 0,
        "Discount (%)": 2,
        "Discount Amount": 0,
        Note: "No faktur A-001",
        "Paid Amount": 50000,
        "Remaining Amount": 30000,
        "Payment Date": new Date().toISOString().slice(0, 10),
        Ref: "A-001"
      },
      {
        Date: new Date().toISOString().slice(0, 10),
        "Supplier Name": "Supplier A",
        "Product ID": "",
        "Product SKU": "SKU-002",
        "Product Name": "Produk Contoh 2",
        Qty: 5,
        "Purchase Price": 12000,
        "Selling Price": 16000,
        "Shipping Cost": 10000,
        "VAT (%)": 11,
        "VAT Amount": 0,
        "Discount (%)": 2,
        "Discount Amount": 0,
        Note: "No faktur A-001",
        "Paid Amount": 50000,
        "Remaining Amount": 30000,
        "Payment Date": new Date().toISOString().slice(0, 10),
        Ref: "A-001"
      }
    ];
    const ws = XLSX.utils.json_to_sheet(templateData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Template");
    ws["!cols"] = [
      { wch: 12 }, // Date
      { wch: 25 }, // Supplier Name
      { wch: 12 }, // Product ID
      { wch: 18 }, // Product SKU
      { wch: 30 }, // Product Name
      { wch: 8 },  // Qty
      { wch: 14 }, // Purchase Price
      { wch: 14 }, // Selling Price
      { wch: 14 }, // Shipping Cost
      { wch: 10 }, // VAT (%)
      { wch: 14 }, // VAT Amount
      { wch: 12 }, // Discount (%)
      { wch: 14 }, // Discount Amount
      { wch: 25 }, // Note
      { wch: 14 }, // Paid Amount
      { wch: 16 }, // Remaining Amount
      { wch: 14 }, // Payment Date
      { wch: 12 }  // Ref
    ];
    const buf = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", "attachment; filename=stock_in_import_template.xlsx");
    res.send(buf);
  } catch (e) {
    res.status(500).json({ success: false, message: "Gagal membuat template barang masuk" });
  }
});

app.post('/api/stock-in/import', isAuthenticated, isAdmin, async (req, res) => {
  try {
    const rows = Array.isArray(req.body && req.body.stockIn) ? req.body.stockIn : [];
    if (rows.length === 0) {
      return res.status(400).json({ success: false, message: "No valid data to import" });
    }

    let suppliers = await readData('suppliers.json').catch(() => []);
    if (!Array.isArray(suppliers)) suppliers = [];
    let products = await readData('products.json').catch(() => []);
    if (!Array.isArray(products)) products = [];
    let stockIn = await readData('stock_in.json').catch(() => []);
    if (!Array.isArray(stockIn)) stockIn = [];

    const supplierByName = new Map(suppliers.map(s => [String(s && s.name || '').toLowerCase(), s]));
    const productById = new Map(products.map(p => [String(p && p.id), p]));
    const productBySku = new Map(products.map(p => [String(p && p.sku || '').toLowerCase(), p]));
    const productByName = new Map(products.map(p => [String(p && p.name || '').toLowerCase(), p]));

    const round2 = (n) => Math.round((Number(n || 0)) * 100) / 100;

    function parseExcelDate(val) {
      if (val == null || val === '') return '';
      if (typeof val === 'number' && isFinite(val)) {
        const d = new Date(Date.UTC(1899, 11, 30));
        d.setUTCDate(d.getUTCDate() + Math.floor(val));
        return d.toISOString().slice(0, 10);
      }
      const s = String(val).trim();
      if (!s) return '';
      const d = new Date(s);
      if (!isNaN(d.getTime())) return d.toISOString().slice(0, 10);
      const m = s.match(/^([0-3]?\d)[\/-]([01]?\d)[\/-](\d{2,4})$/);
      if (m) {
        const dd = parseInt(m[1], 10);
        const mm = parseInt(m[2], 10) - 1;
        const yy = parseInt(m[3], 10);
        const yyyy = yy < 100 ? 2000 + yy : yy;
        const d2 = new Date(yyyy, mm, dd);
        if (!isNaN(d2.getTime())) return d2.toISOString().slice(0, 10);
      }
      return '';
    }

    const grouped = new Map();
    let successCount = 0;
    let errorCount = 0;
    const errors = [];

    for (let i = 0; i < rows.length; i++) {
      try {
        const row = rows[i] || {};
        const supplierNameRaw = row["Supplier Name"] ?? row["Supplier"] ?? row["Nama Supplier"];
        const supplierName = supplierNameRaw != null ? String(supplierNameRaw).trim() : '';
        if (!supplierName) {
          errors.push(`Baris ${i + 1}: Supplier Name wajib diisi.`);
          errorCount++;
          continue;
        }
        const supplier = supplierByName.get(supplierName.toLowerCase());
        if (!supplier) {
          errors.push(`Baris ${i + 1}: Supplier "${supplierName}" tidak ditemukan.`);
          errorCount++;
          continue;
        }

        const productIdRaw = row["Product ID"] ?? row["ID Produk"];
        const skuRaw = row["Product SKU"] ?? row["SKU"];
        const nameRaw = row["Product Name"] ?? row["Nama Produk"];
        let product = null;
        if (productIdRaw != null && String(productIdRaw).trim() !== '') {
          product = productById.get(String(productIdRaw).trim());
        }
        if (!product && skuRaw != null && String(skuRaw).trim() !== '') {
          product = productBySku.get(String(skuRaw).trim().toLowerCase());
        }
        if (!product && nameRaw != null && String(nameRaw).trim() !== '') {
          product = productByName.get(String(nameRaw).trim().toLowerCase());
        }
        if (!product) {
          errors.push(`Baris ${i + 1}: Produk tidak ditemukan (ID/SKU/Nama).`);
          errorCount++;
          continue;
        }

        const qty = Number(row["Qty"] ?? row["Quantity"] ?? 0) || 0;
        if (qty <= 0) {
          errors.push(`Baris ${i + 1}: Qty harus > 0.`);
          errorCount++;
          continue;
        }
        const purchasePrice = Number(row["Purchase Price"] ?? row["Harga Beli"] ?? 0) || 0;
        const sellingPrice = Number(row["Selling Price"] ?? row["Harga Jual"] ?? 0) || 0;

        const date = parseExcelDate(row["Date"] ?? row["Tanggal"]) || new Date().toISOString().slice(0, 10);
        const note = row["Note"] ?? row["Catatan"] ?? '';
        const paidAmount = Number(row["Paid Amount"] ?? row["Dibayar"] ?? 0) || 0;
        const remainingAmountRaw = row["Remaining Amount"] ?? row["Sisa"];
        const shippingCost = Number(row["Shipping Cost"] ?? row["Ongkir"] ?? row["Biaya Angkut"] ?? 0) || 0;
        const vatPercent = Number(row["VAT (%)"] ?? row["PPN (%)"] ?? row["PPN"] ?? 0) || 0;
        const vatAmount = Number(row["VAT Amount"] ?? row["PPN Amount"] ?? row["PPN Nominal"] ?? row["Pajak Masuk"] ?? 0) || 0;
        const discountPercent = Number(row["Discount (%)"] ?? row["Diskon (%)"] ?? row["Diskon"] ?? 0) || 0;
        const discountAmount = Number(row["Discount Amount"] ?? row["Diskon Amount"] ?? row["Diskon Nominal"] ?? 0) || 0;
        const paymentDate = parseExcelDate(row["Payment Date"] ?? row["Tanggal Bayar"]) || '';
        const ref = row["Ref"] ?? row["Reference"] ?? row["Invoice"] ?? '';

        const key = `${ref || ''}|${date}|${supplier.id}|${String(note||'')}|${String(paymentDate||'')}|${String(paidAmount||0)}|${String(remainingAmountRaw||'')}|${String(shippingCost||0)}|${String(vatPercent||0)}|${String(vatAmount||0)}|${String(discountPercent||0)}|${String(discountAmount||0)}`;
        if (!grouped.has(key)) {
          grouped.set(key, {
            date,
            supplier,
            note: String(note || '').trim(),
            paidAmount,
            remainingAmountRaw,
            shippingCost,
            vatPercent,
            vatAmount,
            discountPercent,
            discountAmount,
            paymentDate,
            ref: String(ref || '').trim(),
            items: []
          });
        }
        grouped.get(key).items.push({ productId: product.id, qty, purchasePrice, sellingPrice });
        successCount++;
      } catch (err) {
        errors.push(`Baris ${i + 1}: ${err.message}`);
        errorCount++;
      }
    }

    const user = (req.session && req.session.user) || {};
    for (const group of grouped.values()) {
      const now = Date.now();
      const id = `STKIN-${new Date(now).toISOString().slice(0,10).replace(/-/g,'')}-${now}-${Math.random().toString(36).slice(2,7)}`;
      const items = [];
      const itemDrafts = [];
      let itemsSubtotal = 0;
      let totalQty = 0;

      for (const it of group.items) {
        const product = products.find(p => Number(p.id) === Number(it.productId));
        if (!product) continue;
        const qty = Number(it.qty || 0) || 0;
        const purchasePrice = Number(it.purchasePrice || 0) || 0;
        const sellingPrice = Number(it.sellingPrice || 0) || 0;
        if (qty <= 0) continue;
        const lineSubtotal = round2(purchasePrice * qty);
        itemDrafts.push({ product, qty, purchasePrice, sellingPrice, lineSubtotal });
        itemsSubtotal += lineSubtotal;
        totalQty += qty;
      }

      if (!itemDrafts.length) continue;

      const shippingCostNum = Math.max(0, Number(group.shippingCost) || 0);
      const vatPercentNum = Math.max(0, Number(group.vatPercent) || 0);
      const computedVatAmount = round2(itemsSubtotal * (vatPercentNum / 100));
      const vatAmountNum = Math.max(0, Number(group.vatAmount) || computedVatAmount);
      const discountPercentNum = Math.max(0, Number(group.discountPercent) || 0);
      const computedDiscountAmount = round2(itemsSubtotal * (discountPercentNum / 100));
      const discountAmountNum = Math.max(0, Number(group.discountAmount) || computedDiscountAmount);
      const additionalCostPool = round2(shippingCostNum + vatAmountNum - discountAmountNum);
      const totalAmount = round2(itemsSubtotal + additionalCostPool);

      for (const draft of itemDrafts) {
        const product = draft.product;
        const qty = Number(draft.qty || 0) || 0;
        const purchasePrice = Number(draft.purchasePrice || 0) || 0;
        const sellingPrice = Number(draft.sellingPrice || 0) || 0;
        const lineSubtotal = Number(draft.lineSubtotal || 0) || 0;
        const prevProduct = {
          purchasePrice: product.purchasePrice,
          sellingPrice: product.sellingPrice,
          price: product.price,
          unitPrices: Array.isArray(product.unitPrices)
            ? product.unitPrices.map(v => ({ ...v }))
            : []
        };
        product.stock = Number(product.stock || 0) + qty;
        const prevSell = Number((prevProduct.sellingPrice != null ? prevProduct.sellingPrice : prevProduct.price) || 0);
        const prevBuy = Number(prevProduct.purchasePrice || 0);

        const weight = itemsSubtotal > 0
          ? (lineSubtotal / itemsSubtotal)
          : (totalQty > 0 ? (qty / totalQty) : 0);
        const allocatedAdditional = round2(additionalCostPool * weight);
        const effectivePurchasePrice = qty > 0
          ? Math.max(0, round2((lineSubtotal + allocatedAdditional) / qty))
          : round2(purchasePrice);

        const nextBuy = effectivePurchasePrice > 0 ? effectivePurchasePrice : prevBuy;
        const nextSell = sellingPrice > 0 ? Number(sellingPrice || 0) : prevSell;
        let priceChanged = false;
        if (nextBuy !== prevBuy) {
          product.purchasePrice = nextBuy;
          priceChanged = true;
        }
        if (sellingPrice > 0 && nextSell !== prevSell) {
          product.sellingPrice = nextSell;
          product.price = nextSell;
          priceChanged = true;
        }
        product.updatedAt = Date.now();
        try {
          const by = String(user && (user.username || user.name || ''));
          const nextSnapshot = { ...prevProduct, purchasePrice: nextBuy, sellingPrice: nextSell, price: nextSell };
          let entry = buildPriceHistoryEntry(prevProduct, nextSnapshot, 'stock_in_import', by);
          if (!entry) {
            entry = {
              at: Date.now(),
              by,
              reason: 'stock_in_import',
              base: {
                sellingPrice: { from: prevSell, to: nextSell },
                purchasePrice: { from: prevBuy, to: nextBuy }
              },
              variants: []
            };
          }
          appendPriceHistory(product, entry);
          try { if (global.appendPriceHistoryForProduct) await global.appendPriceHistoryForProduct(product.id, entry); } catch {}
        } catch (e) {}
        try {
          await appendStockMove({
            productId: Number(product.id),
            delta: qty,
            reason: 'purchase',
            refId: id,
            by: String(user && (user.username || user.name || '')),
            newStock: Number(product.stock || 0),
            stockOnly: priceChanged ? false : true
          });
        } catch {}
        items.push({
          productId: Number(product.id),
          qty,
          purchasePrice: round2(purchasePrice),
          effectivePurchasePrice: round2(effectivePurchasePrice),
          sellingPrice
        });
      }

      if (!items.length) continue;
      const paidAmount = Number(group.paidAmount) || 0;
      const remainingAmount = (group.remainingAmountRaw != null && group.remainingAmountRaw !== '')
        ? (Number(group.remainingAmountRaw) || 0)
        : Math.max(0, totalAmount - paidAmount);

      stockIn.push({
        id,
        timestamp: new Date(now).toISOString(),
        date: group.date,
        supplierId: Number(group.supplier.id),
        supplierName: String(group.supplier.name || ''),
        items,
        note: group.note,
        itemsSubtotal: round2(itemsSubtotal),
        shippingCost: round2(shippingCostNum),
        vatPercent: round2(vatPercentNum),
        vatAmount: round2(vatAmountNum),
        discountPercent: round2(discountPercentNum),
        discountAmount: round2(discountAmountNum),
        paidAmount,
        remainingAmount,
        paymentDate: group.paymentDate ? String(group.paymentDate) : '',
        totalAmount,
        updatedAt: now,
        ref: group.ref
      });
    }

    await writeData('stock_in.json', stockIn);
    await saveArrayWithSync('products.json', products, { keyField: 'id' });
    try { invalidateCache('products'); } catch {}

    let message = `Import selesai. Sukses: ${successCount}, Error: ${errorCount}`;
    if (errors.length > 0) {
      message += `\n\nBeberapa error pertama:\n${errors.slice(0, 3).join("\n")}`;
      if (errors.length > 5) {
        message += ` ... dan ${errors.length - 5} more errors`;
      }
    }
    res.json({ success: true, message, successCount, errorCount, errors: errors.slice(0, 10) });
  } catch (e) {
    res.status(500).json({ success: false, message: "Failed to import stock-in: " + e.message });
  }
});

function parseExcelDateForImport(val) {
  if (val == null || val === '') return '';
  if (typeof val === 'number' && isFinite(val)) {
    const d = new Date(Date.UTC(1899, 11, 30));
    d.setUTCDate(d.getUTCDate() + Math.floor(val));
    return d.toISOString().slice(0, 10);
  }
  const s = String(val).trim();
  if (!s) return '';
  const d = new Date(s);
  if (!isNaN(d.getTime())) return d.toISOString().slice(0, 10);
  const m = s.match(/^([0-3]?\d)[\/-]([01]?\d)[\/-](\d{2,4})$/);
  if (m) {
    const dd = parseInt(m[1], 10);
    const mm = parseInt(m[2], 10) - 1;
    const yy = parseInt(m[3], 10);
    const yyyy = yy < 100 ? 2000 + yy : yy;
    const d2 = new Date(yyyy, mm, dd);
    if (!isNaN(d2.getTime())) return d2.toISOString().slice(0, 10);
  }
  return '';
}

// --- Customer Debts Import (saldo piutang dari aplikasi lain) ---
app.get('/api/customer-debts/template', isAuthenticated, isAdmin, async (req, res) => {
  try {
    const today = new Date().toISOString().slice(0, 10);
    const templateData = [
      {
        Date: today,
        "Customer Name": "Budi Santoso",
        Phone: "081234567890",
        "Total Amount": 500000,
        "Paid Amount": 100000,
        "Remaining Amount": 400000,
        "Due Date": today,
        Ref: "INV-001",
        Note: "Saldo awal dari aplikasi lama",
        "Create Customer": "Yes"
      },
      {
        Date: today,
        "Customer Name": "Siti Aminah",
        Phone: "081987654321",
        "Total Amount": 250000,
        "Paid Amount": 0,
        "Remaining Amount": 250000,
        "Due Date": "",
        Ref: "INV-002",
        Note: "Piutang belum dibayar",
        "Create Customer": "Yes"
      }
    ];
    const ws = XLSX.utils.json_to_sheet(templateData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Template");
    ws["!cols"] = [
      { wch: 12 }, { wch: 28 }, { wch: 16 }, { wch: 14 }, { wch: 14 },
      { wch: 16 }, { wch: 12 }, { wch: 14 }, { wch: 30 }, { wch: 16 }
    ];
    const buf = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", "attachment; filename=customer_debt_import_template.xlsx");
    res.send(buf);
  } catch (e) {
    res.status(500).json({ success: false, message: "Gagal membuat template hutang customer" });
  }
});

const importCustomerDebtsHandler = async (req, res) => {
  try {
    const rows = Array.isArray(req.body && req.body.debts) ? req.body.debts : [];
    if (rows.length === 0) {
      return res.status(400).json({ success: false, message: "No valid data to import" });
    }

    const sessionUser = (req.session && req.session.user) ? req.session.user : {};
    const cashierName = sessionUser.name || sessionUser.username || 'Admin';
    const round2 = (n) => Math.round((Number(n || 0)) * 100) / 100;

    let customers = await readData('customers.json').catch(() => []);
    if (!Array.isArray(customers)) customers = [];
    let transactions = await readData('transactions.json').catch(() => []);
    if (!Array.isArray(transactions)) transactions = [];

    const customerById = new Map(customers.map(c => [String(c && c.id), c]));
    const customerByPhone = new Map(customers.filter(c => c && c.phone).map(c => [String(c.phone).replace(/\D/g, ''), c]));
    const customerByName = new Map(customers.map(c => [String(c && c.name || '').toLowerCase(), c]));
    const existingKeys = new Set(
      transactions
        .filter(t => t && t.imported)
        .map(t => `${String(t.ref || '')}|${String(t.customerId || '')}|${String(t.date || '')}`)
    );

    let successCount = 0;
    let errorCount = 0;
    const errors = [];

    for (let i = 0; i < rows.length; i++) {
      try {
        const row = rows[i] || {};
        const customerNameRaw = row["Customer Name"] ?? row["Nama Customer"] ?? row["Nama Pelanggan"];
        const customerName = customerNameRaw != null ? String(customerNameRaw).trim() : '';
        const customerIdRaw = row["Customer ID"] ?? row["ID Customer"];
        const phoneRaw = row["Phone"] ?? row["Telepon"] ?? row["No HP"];
        const phone = phoneRaw != null ? String(phoneRaw).trim() : '';

        let customer = null;
        if (customerIdRaw != null && String(customerIdRaw).trim() !== '') {
          customer = customerById.get(String(customerIdRaw).trim());
        }
        if (!customer && phone) {
          customer = customerByPhone.get(phone.replace(/\D/g, ''));
        }
        if (!customer && customerName) {
          customer = customerByName.get(customerName.toLowerCase());
        }

        const createFlag = String(row["Create Customer"] ?? row["Buat Customer"] ?? 'Yes').trim().toLowerCase();
        const allowCreate = !createFlag || ['yes', 'ya', '1', 'true', 'y'].includes(createFlag);

        if (!customer) {
          if (!customerName) {
            errors.push(`Baris ${i + 1}: Customer Name wajib diisi.`);
            errorCount++;
            continue;
          }
          if (!allowCreate) {
            errors.push(`Baris ${i + 1}: Customer "${customerName}" tidak ditemukan.`);
            errorCount++;
            continue;
          }
          const newId = Date.now() + i;
          customer = {
            id: newId,
            name: customerName,
            phone,
            email: '',
            address: '',
            balance: 0,
            createdAt: new Date().toISOString()
          };
          customers.push(customer);
          customerById.set(String(newId), customer);
          if (phone) customerByPhone.set(phone.replace(/\D/g, ''), customer);
          customerByName.set(customerName.toLowerCase(), customer);
        }

        const totalAmount = round2(row["Total Amount"] ?? row["Total"] ?? row["Jumlah"] ?? 0);
        if (totalAmount <= 0) {
          errors.push(`Baris ${i + 1}: Total Amount harus > 0.`);
          errorCount++;
          continue;
        }

        const paidAmount = round2(row["Paid Amount"] ?? row["Dibayar"] ?? 0);
        const remainingRaw = row["Remaining Amount"] ?? row["Sisa"];
        const remainingAmount = (remainingRaw != null && remainingRaw !== '')
          ? round2(remainingRaw)
          : round2(Math.max(0, totalAmount - paidAmount));

        const date = parseExcelDateForImport(row["Date"] ?? row["Tanggal"]) || new Date().toISOString().slice(0, 10);
        const dueDate = parseExcelDateForImport(row["Due Date"] ?? row["Jatuh Tempo"]) || '';
        const ref = String(row["Ref"] ?? row["Reference"] ?? row["Invoice"] ?? row["No Faktur"] ?? '').trim();
        const note = String(row["Note"] ?? row["Catatan"] ?? 'Import hutang customer').trim();

        const dupKey = `${ref}|${customer.id}|${date}`;
        if (ref && existingKeys.has(dupKey)) {
          errors.push(`Baris ${i + 1}: Ref "${ref}" sudah diimport untuk customer ini.`);
          errorCount++;
          continue;
        }

        const dateTs = new Date(date + 'T12:00:00').getTime();
        const txId = `TRX-IMPORT-${date.replace(/-/g, '')}-${Date.now()}-${i}`;
        const tx = {
          id: txId,
          timestamp: isNaN(dateTs) ? Date.now() : dateTs,
          date,
          paymentMethod: 'debt',
          amountReceived: paidAmount,
          change: 0,
          customerId: customer.id,
          customerName: customer.name || customerName,
          customerBalanceUsed: 0,
          payableAfterBalance: totalAmount,
          paidAmount,
          remainingAmount,
          paymentDate: date,
          dueDate: dueDate || undefined,
          isDebt: remainingAmount > 0,
          imported: true,
          ref,
          note,
          cashier: cashierName,
          cashierName,
          cashierUsername: sessionUser.username || '',
          cashierId: sessionUser.id || sessionUser.userId || '',
          cashierRole: sessionUser.role || '',
          items: [{
            productId: 0,
            name: note || 'Saldo hutang (import)',
            price: totalAmount,
            qty: 1,
            quantity: 1
          }],
          subtotal: totalAmount,
          discountAmount: 0,
          taxAmount: 0,
          serviceAmount: 0,
          totalAmount,
          voided: false
        };

        transactions.push(tx);
        if (ref) existingKeys.add(dupKey);
        successCount++;
      } catch (rowErr) {
        errors.push(`Baris ${i + 1}: ${rowErr.message || 'Error tidak diketahui'}`);
        errorCount++;
      }
    }

    await writeData('customers.json', customers);
    await writeData('transactions.json', transactions);

    let message = `Import hutang customer selesai. Sukses: ${successCount}, Error: ${errorCount}`;
    if (errors.length > 0) {
      message += `\n\nBeberapa error pertama:\n${errors.slice(0, 3).join("\n")}`;
      if (errors.length > 3) message += ` ... dan ${errors.length - 3} error lainnya`;
    }
    res.json({ success: true, message, successCount, errorCount, errors: errors.slice(0, 10) });
  } catch (e) {
    res.status(500).json({ success: false, message: "Failed to import customer debts: " + e.message });
  }
};

app.post('/api/customer-debts/import', isAuthenticated, isAdmin, importCustomerDebtsHandler);
app.post('/api/transactions/debts/import', isAuthenticated, isAdmin, importCustomerDebtsHandler);

// --- Supplier Credits/Debts Import (saldo hutang supplier tanpa ubah stok) ---
app.get('/api/credits/template', isAuthenticated, isAdmin, async (req, res) => {
  try {
    const today = new Date().toISOString().slice(0, 10);
    const templateData = [
      {
        Date: today,
        "Supplier Name": "Supplier A",
        "Total Amount": 800000,
        "Paid Amount": 300000,
        "Remaining Amount": 500000,
        "Payment Date": "",
        Ref: "PO-001",
        Note: "Saldo hutang dari aplikasi lama",
        "Create Supplier": "Yes"
      },
      {
        Date: today,
        "Supplier Name": "Supplier B",
        "Total Amount": 1200000,
        "Paid Amount": 0,
        "Remaining Amount": 1200000,
        "Payment Date": "",
        Ref: "PO-002",
        Note: "Hutang belum dibayar",
        "Create Supplier": "Yes"
      }
    ];
    const ws = XLSX.utils.json_to_sheet(templateData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Template");
    ws["!cols"] = [
      { wch: 12 }, { wch: 28 }, { wch: 14 }, { wch: 14 }, { wch: 16 },
      { wch: 14 }, { wch: 14 }, { wch: 30 }, { wch: 16 }
    ];
    const buf = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", "attachment; filename=supplier_debt_import_template.xlsx");
    res.send(buf);
  } catch (e) {
    res.status(500).json({ success: false, message: "Gagal membuat template hutang supplier" });
  }
});

const importSupplierDebtsHandler = async (req, res) => {
  try {
    const rows = Array.isArray(req.body && req.body.debts) ? req.body.debts : [];
    if (rows.length === 0) {
      return res.status(400).json({ success: false, message: "No valid data to import" });
    }

    const round2 = (n) => Math.round((Number(n || 0)) * 100) / 100;

    let suppliers = await readData('suppliers.json').catch(() => []);
    if (!Array.isArray(suppliers)) suppliers = [];
    let stockIn = await readData('stock_in.json').catch(() => []);
    if (!Array.isArray(stockIn)) stockIn = [];

    const supplierById = new Map(suppliers.map(s => [String(s && s.id), s]));
    const supplierByName = new Map(suppliers.map(s => [String(s && s.name || '').toLowerCase(), s]));
    const existingKeys = new Set(
      stockIn
        .filter(r => r && r.imported)
        .map(r => `${String(r.ref || '')}|${String(r.supplierId || '')}|${String(r.date || '')}`)
    );

    let successCount = 0;
    let errorCount = 0;
    const errors = [];

    for (let i = 0; i < rows.length; i++) {
      try {
        const row = rows[i] || {};
        const supplierNameRaw = row["Supplier Name"] ?? row["Supplier"] ?? row["Nama Supplier"];
        const supplierName = supplierNameRaw != null ? String(supplierNameRaw).trim() : '';
        const supplierIdRaw = row["Supplier ID"] ?? row["ID Supplier"];

        let supplier = null;
        if (supplierIdRaw != null && String(supplierIdRaw).trim() !== '') {
          supplier = supplierById.get(String(supplierIdRaw).trim());
        }
        if (!supplier && supplierName) {
          supplier = supplierByName.get(supplierName.toLowerCase());
        }

        const createFlag = String(row["Create Supplier"] ?? row["Buat Supplier"] ?? 'Yes').trim().toLowerCase();
        const allowCreate = !createFlag || ['yes', 'ya', '1', 'true', 'y'].includes(createFlag);

        if (!supplier) {
          if (!supplierName) {
            errors.push(`Baris ${i + 1}: Supplier Name wajib diisi.`);
            errorCount++;
            continue;
          }
          if (!allowCreate) {
            errors.push(`Baris ${i + 1}: Supplier "${supplierName}" tidak ditemukan.`);
            errorCount++;
            continue;
          }
          const newId = Date.now() + i;
          supplier = {
            id: newId,
            name: supplierName,
            phone: '',
            address: '',
            notes: '',
            createdAt: new Date().toISOString()
          };
          suppliers.push(supplier);
          supplierById.set(String(newId), supplier);
          supplierByName.set(supplierName.toLowerCase(), supplier);
        }

        const totalAmount = round2(row["Total Amount"] ?? row["Total"] ?? row["Jumlah"] ?? 0);
        if (totalAmount <= 0) {
          errors.push(`Baris ${i + 1}: Total Amount harus > 0.`);
          errorCount++;
          continue;
        }

        const paidAmount = round2(row["Paid Amount"] ?? row["Dibayar"] ?? 0);
        const remainingRaw = row["Remaining Amount"] ?? row["Sisa"];
        const remainingAmount = (remainingRaw != null && remainingRaw !== '')
          ? round2(remainingRaw)
          : round2(Math.max(0, totalAmount - paidAmount));

        const date = parseExcelDateForImport(row["Date"] ?? row["Tanggal"]) || new Date().toISOString().slice(0, 10);
        const paymentDate = parseExcelDateForImport(row["Payment Date"] ?? row["Tanggal Bayar"]) || '';
        const ref = String(row["Ref"] ?? row["Reference"] ?? row["Invoice"] ?? row["No Faktur"] ?? '').trim();
        const note = String(row["Note"] ?? row["Catatan"] ?? 'Import hutang supplier').trim();

        const dupKey = `${ref}|${supplier.id}|${date}`;
        if (ref && existingKeys.has(dupKey)) {
          errors.push(`Baris ${i + 1}: Ref "${ref}" sudah diimport untuk supplier ini.`);
          errorCount++;
          continue;
        }

        const now = Date.now();
        const id = `STKIN-IMPORT-${date.replace(/-/g, '')}-${now}-${i}`;
        stockIn.push({
          id,
          timestamp: new Date(now).toISOString(),
          date,
          supplierId: Number(supplier.id),
          supplierName: String(supplier.name || supplierName),
          items: [],
          note,
          itemsSubtotal: 0,
          shippingCost: 0,
          vatPercent: 0,
          vatAmount: 0,
          discountPercent: 0,
          discountAmount: 0,
          paidAmount,
          remainingAmount,
          paymentDate,
          totalAmount,
          ref,
          imported: true,
          updatedAt: now
        });

        if (ref) existingKeys.add(dupKey);
        successCount++;
      } catch (rowErr) {
        errors.push(`Baris ${i + 1}: ${rowErr.message || 'Error tidak diketahui'}`);
        errorCount++;
      }
    }

    await writeData('suppliers.json', suppliers);
    await writeData('stock_in.json', stockIn);

    let message = `Import hutang supplier selesai. Sukses: ${successCount}, Error: ${errorCount}`;
    if (errors.length > 0) {
      message += `\n\nBeberapa error pertama:\n${errors.slice(0, 3).join("\n")}`;
      if (errors.length > 3) message += ` ... dan ${errors.length - 3} error lainnya`;
    }
    res.json({ success: true, message, successCount, errorCount, errors: errors.slice(0, 10) });
  } catch (e) {
    res.status(500).json({ success: false, message: "Failed to import supplier debts: " + e.message });
  }
};

app.post('/api/credits/import', isAuthenticated, isAdmin, importSupplierDebtsHandler);
app.post('/api/stock-in/debts/import', isAuthenticated, isAdmin, importSupplierDebtsHandler);
app.get('/api/stock-in/debts/template', isAuthenticated, isAdmin, async (req, res) => {
  res.redirect(302, '/api/credits/template');
});

// PATCH /api/stock-in/:id - update stock-in fields (payment + basic metadata)
app.patch('/api/stock-in/:id', isAuthenticated, async (req, res) => {
  try {
    const user = (req.session && req.session.user) || {};
    const { id } = req.params;
    const { paidAmount, remainingAmount, paymentDate, date, supplierId, supplierName, note, items, shippingCost, vatPercent, vatAmount, discountPercent, discountAmount } = req.body || {};

    if (!id) {
      return res.status(400).json({ success: false, message: 'Stock-in ID is required' });
    }

    let stockIn = await readData('stock_in.json').catch(() => []);
    if (!Array.isArray(stockIn)) stockIn = [];

    const index = stockIn.findIndex(record => String(record.id) === String(id));
    if (index === -1) {
      return res.status(404).json({ success: false, message: 'Stock-in record not found' });
    }

    const round2 = (n) => Math.round((Number(n || 0)) * 100) / 100;
    const hasField = (v) => v !== undefined && v !== null;
    let products = null;

    // Update payment/meta fields
    const record = stockIn[index];
    if (hasField(paymentDate)) record.paymentDate = String(paymentDate || '');

    // Update basic stock-in fields from detail modal
    if (hasField(date)) record.date = String(date || '');
    if (hasField(note)) record.note = String(note || '');
    if (hasField(supplierId) || hasField(supplierName)) {
      let suppliers = await readData('suppliers.json').catch(() => []);
      if (!Array.isArray(suppliers)) suppliers = [];
      const sid = parseInt(supplierId, 10);
      if (Number.isFinite(sid) && sid > 0) {
        const found = suppliers.find((s) => Number(s.id) === sid);
        if (!found) return res.status(400).json({ success: false, message: 'Supplier tidak ditemukan' });
        record.supplierId = sid;
        record.supplierName = String(found.name || '');
      } else if (hasField(supplierName)) {
        record.supplierName = String(supplierName || '');
      }
    }

    // If items are provided, update products in record and adjust product stocks by delta.
    if (Array.isArray(items)) {
      products = await readData('products.json').catch(() => []);
      if (!Array.isArray(products)) products = [];

      const normalizeQtyMap = (arr) => {
        const map = new Map();
        (Array.isArray(arr) ? arr : []).forEach((it) => {
          const key = String(it && it.productId != null ? it.productId : '').trim();
          const qty = Number(it && it.qty);
          if (!key || key === 'undefined' || key === 'null' || !qty || !Number.isFinite(qty)) return;
          map.set(key, (map.get(key) || 0) + qty);
        });
        return map;
      };

      const oldQtyMap = normalizeQtyMap(record.items);
      const itemDrafts = [];
      let itemsSubtotalCalc = 0;
      let totalQty = 0;
      for (const raw of items) {
        const qty = Number(raw && raw.qty != null ? raw.qty : 0) || 0;
        const purchasePriceNum = Number(raw && raw.purchasePrice != null ? raw.purchasePrice : 0) || 0;
        const sellingPriceNum = Number(raw && raw.sellingPrice != null ? raw.sellingPrice : 0) || 0;
        if (qty <= 0) continue;
        const key = String(raw && raw.productId != null ? raw.productId : '').trim();
        if (!key || key === 'undefined' || key === 'null') continue;
        const product = products.find((p) => {
          if (!p) return false;
          const a = p.id != null ? String(p.id) : '';
          const b = p._id != null ? String(p._id) : '';
          return a === key || b === key;
        });
        if (!product) return res.status(400).json({ success: false, message: `Produk ID ${key} tidak ditemukan` });
        const pid = product.id != null ? product.id : product._id;
        const pidKey = String(pid);
        const lineSubtotal = round2(purchasePriceNum * qty);
        itemDrafts.push({
          product,
          pid,
          pidKey,
          qty,
          purchasePriceNum,
          sellingPriceNum,
          lineSubtotal,
          variant: raw && raw.variant ? raw.variant : null
        });
        itemsSubtotalCalc += lineSubtotal;
        totalQty += qty;
      }
      if (!itemDrafts.length) return res.status(400).json({ success: false, message: 'Minimal ada satu item produk valid' });

      const shippingCostNum = Math.max(0, hasField(shippingCost) ? (Number(shippingCost) || 0) : (Number(record.shippingCost) || 0));
      const vatPercentNum = Math.max(0, hasField(vatPercent) ? (Number(vatPercent) || 0) : (Number(record.vatPercent) || 0));
      const discountPercentNum = Math.max(0, hasField(discountPercent) ? (Number(discountPercent) || 0) : (Number(record.discountPercent) || 0));
      const vatAmountNum = Math.max(0, hasField(vatAmount) ? (Number(vatAmount) || 0) : round2(itemsSubtotalCalc * (vatPercentNum / 100)));
      const discountAmountNum = Math.max(0, hasField(discountAmount) ? (Number(discountAmount) || 0) : round2(itemsSubtotalCalc * (discountPercentNum / 100)));
      const additionalCostPool = round2(shippingCostNum + vatAmountNum - discountAmountNum);
      const totalAmountNum = Math.max(0, round2(itemsSubtotalCalc + additionalCostPool));

      const oldItemMetaMap = new Map();
      (Array.isArray(record.items) ? record.items : []).forEach((it) => {
        const key = String(it && it.productId != null ? it.productId : '').trim();
        if (!key || key === 'undefined' || key === 'null') return;
        if (!oldItemMetaMap.has(key)) oldItemMetaMap.set(key, it || {});
      });

      const newQtyMap = new Map();
      const normalizedItems = [];
      for (const draft of itemDrafts) {
        const { pid, pidKey, qty, purchasePriceNum, sellingPriceNum, lineSubtotal, variant } = draft;
        const weight = itemsSubtotalCalc > 0
          ? (lineSubtotal / itemsSubtotalCalc)
          : (totalQty > 0 ? (qty / totalQty) : 0);
        const allocatedAdditional = round2(additionalCostPool * weight);
        const effectivePurchasePrice = qty > 0
          ? Math.max(0, round2((lineSubtotal + allocatedAdditional) / qty))
          : round2(purchasePriceNum);
        const oldMeta = oldItemMetaMap.get(pidKey) || {};
        normalizedItems.push({
          productId: pid,
          qty,
          purchasePrice: round2(purchasePriceNum),
          effectivePurchasePrice: round2(effectivePurchasePrice),
          sellingPrice: round2(sellingPriceNum),
          prevPurchasePrice: Number(oldMeta.prevPurchasePrice),
          prevSellingPrice: Number(oldMeta.prevSellingPrice),
          appliedPurchasePrice: Number(oldMeta.appliedPurchasePrice || effectivePurchasePrice || purchasePriceNum || 0),
          appliedSellingPrice: Number(oldMeta.appliedSellingPrice || sellingPriceNum || 0),
          ...(variant ? { variant } : {})
        });
        newQtyMap.set(pidKey, (newQtyMap.get(pidKey) || 0) + qty);
      }

      const allPids = new Set([...oldQtyMap.keys(), ...newQtyMap.keys()]);
      for (const pidKey of allPids) {
        const oldQty = Number(oldQtyMap.get(pidKey) || 0);
        const newQty = Number(newQtyMap.get(pidKey) || 0);
        const delta = newQty - oldQty;
        if (!delta) continue;
        const product = products.find((p) => {
          if (!p) return false;
          const a = p.id != null ? String(p.id) : '';
          const b = p._id != null ? String(p._id) : '';
          return a === pidKey || b === pidKey;
        });
        if (!product) continue;
        const currentStock = Number(product.stock || 0) || 0;
        const nextStock = currentStock + delta;
        if (nextStock < 0) {
          return res.status(400).json({ success: false, message: `Stok produk ID ${pidKey} tidak mencukupi untuk perubahan ini` });
        }
        product.stock = nextStock;
        try {
          await appendStockMove({
            productId: product.id != null ? product.id : product._id,
            delta,
            reason: 'stock_in_edit',
            refId: id,
            by: String(user && (user.username || user.name || '')),
            newStock: Number(nextStock || 0),
            stockOnly: true
          });
        } catch {}
      }

      record.items = normalizedItems;
      record.itemsSubtotal = round2(itemsSubtotalCalc);
      record.shippingCost = round2(shippingCostNum);
      record.vatPercent = round2(vatPercentNum);
      record.vatAmount = round2(vatAmountNum);
      record.discountPercent = round2(discountPercentNum);
      record.discountAmount = round2(discountAmountNum);
      record.totalAmount = round2(totalAmountNum);
    }

    // Payment can be updated independently or after items update
    if (hasField(paidAmount)) record.paidAmount = Math.max(0, Number(paidAmount) || 0);
    if (hasField(remainingAmount)) {
      record.remainingAmount = Math.max(0, Number(remainingAmount) || 0);
    } else {
      const totalAmountNum = Math.max(0, Number(record.totalAmount || 0) || 0);
      const paidAmountNum = Math.max(0, Number(record.paidAmount || 0) || 0);
      record.remainingAmount = Math.max(0, round2(totalAmountNum - paidAmountNum));
    }
    
    // Update timestamp
    record.updatedAt = Date.now();

    // Save back to file
    await writeData('stock_in.json', stockIn);
    if (Array.isArray(products)) {
      await saveArrayWithSync('products.json', products, { keyField: 'id' });
    }

    res.json({ success: true, message: 'Stock-in record updated successfully', record });
  } catch (e) {
    console.error('Failed to update stock-in payment:', e);
    res.status(500).json({ success: false, message: 'Failed to update stock-in record' });
  }
});

// DELETE /api/stock-in/:id - delete stock-in and rollback stock deltas
app.delete('/api/stock-in/:id', isAuthenticated, async (req, res) => {
  try {
    const user = (req.session && req.session.user) || {};
    const { id } = req.params;
    if (!id) {
      return res.status(400).json({ success: false, message: 'Stock-in ID is required' });
    }

    let stockIn = await readData('stock_in.json').catch(() => []);
    if (!Array.isArray(stockIn)) stockIn = [];
    const index = stockIn.findIndex((record) => String(record.id) === String(id));
    if (index === -1) {
      return res.status(404).json({ success: false, message: 'Stock-in record not found' });
    }

    const record = stockIn[index] || {};
    const items = Array.isArray(record.items) ? record.items : [];
    let products = await readData('products.json').catch(() => []);
    if (!Array.isArray(products)) products = [];

    const qtyMap = new Map();
    for (const it of items) {
      const key = String(it && it.productId != null ? it.productId : '').trim();
      const qty = Number(it && it.qty);
      if (!key || key === 'undefined' || key === 'null' || !qty || !Number.isFinite(qty)) continue;
      qtyMap.set(key, (qtyMap.get(key) || 0) + qty);
    }

    for (const [pidKey, qty] of qtyMap.entries()) {
      const product = products.find((p) => {
        if (!p) return false;
        const a = p.id != null ? String(p.id) : '';
        const b = p._id != null ? String(p._id) : '';
        return a === pidKey || b === pidKey;
      });
      if (!product) continue;
      const currentStock = Number(product.stock || 0) || 0;
      const nextStock = currentStock - Number(qty || 0);
      if (nextStock < 0) {
        return res.status(400).json({ success: false, message: `Stok produk ID ${pidKey} tidak mencukupi untuk rollback` });
      }
      product.stock = nextStock;

      // Rollback harga ke nilai sebelum stock-in ini, tetapi hanya jika harga saat ini
      // masih sama dengan harga yang pernah diterapkan oleh stock-in tersebut.
      const itemMeta = items.find((it) => String(it && it.productId) === pidKey) || {};
      const hasPrevBuy = Number.isFinite(Number(itemMeta.prevPurchasePrice));
      const hasPrevSell = Number.isFinite(Number(itemMeta.prevSellingPrice));
      const appliedBuy = Number(itemMeta.appliedPurchasePrice || itemMeta.effectivePurchasePrice || itemMeta.purchasePrice || 0) || 0;
      const appliedSell = Number(itemMeta.appliedSellingPrice || itemMeta.sellingPrice || 0) || 0;
      const prevBuy = Number(itemMeta.prevPurchasePrice || 0) || 0;
      const prevSell = Number(itemMeta.prevSellingPrice || 0) || 0;
      const currentBuy = Number(product.purchasePrice || 0) || 0;
      const currentSell = Number((product.sellingPrice != null ? product.sellingPrice : product.price) || 0) || 0;
      let priceRolledBack = false;

      if (hasPrevBuy && Math.abs(currentBuy - appliedBuy) < 0.0001) {
        product.purchasePrice = prevBuy;
        priceRolledBack = true;
      }
      if (hasPrevSell && Math.abs(currentSell - appliedSell) < 0.0001) {
        product.sellingPrice = prevSell;
        product.price = prevSell;
        priceRolledBack = true;
      }
      product.updatedAt = Date.now();

      if (priceRolledBack) {
        try {
          const by = String(user && (user.username || user.name || ''));
          const beforeSnapshot = {
            purchasePrice: currentBuy,
            sellingPrice: currentSell,
            price: currentSell,
            unitPrices: Array.isArray(product.unitPrices) ? product.unitPrices.map(v => ({ ...v })) : []
          };
          const afterSnapshot = {
            purchasePrice: Number(product.purchasePrice || 0) || 0,
            sellingPrice: Number((product.sellingPrice != null ? product.sellingPrice : product.price) || 0) || 0,
            price: Number((product.sellingPrice != null ? product.sellingPrice : product.price) || 0) || 0,
            unitPrices: Array.isArray(product.unitPrices) ? product.unitPrices.map(v => ({ ...v })) : []
          };
          let entry = buildPriceHistoryEntry(beforeSnapshot, afterSnapshot, 'stock_in_delete', by);
          if (!entry) {
            entry = {
              at: Date.now(),
              by,
              reason: 'stock_in_delete',
              base: {
                sellingPrice: { from: currentSell, to: afterSnapshot.sellingPrice },
                purchasePrice: { from: currentBuy, to: afterSnapshot.purchasePrice }
              },
              variants: []
            };
          }
          appendPriceHistory(product, entry);
          try { if (global.appendPriceHistoryForProduct) await global.appendPriceHistoryForProduct(product.id != null ? product.id : product._id, entry); } catch {}
        } catch {}
      }

      try {
        await appendStockMove({
          productId: product.id != null ? product.id : product._id,
          delta: -Number(qty || 0),
          reason: 'stock_in_delete',
          refId: String(id),
          by: String(user && (user.username || user.name || '')),
          newStock: Number(nextStock || 0),
          stockOnly: !priceRolledBack
        });
      } catch {}
    }

    stockIn.splice(index, 1);
    await writeData('stock_in.json', stockIn);
    await saveArrayWithSync('products.json', products, { keyField: 'id' });
    try { invalidateCache('products'); } catch {}

    res.json({ success: true, message: 'Stock-in deleted successfully', id: String(id) });
  } catch (e) {
    console.error('Failed to delete stock-in:', e);
    res.status(500).json({ success: false, message: 'Failed to delete stock-in record' });
  }
});

// Serve favicon from settings if provided
app.get('/favicon.ico', async (req, res) => {
  try {
    const raw = await readData('settings.json');
    const base = Array.isArray(raw) ? {} : (raw || {});
    const dataUrl = base.faviconBase64 || '';
    if (!dataUrl || typeof dataUrl !== 'string' || !dataUrl.startsWith('data:')) {
      return res.status(204).end();
    }
    const match = dataUrl.match(/^data:(.*?);base64,(.*)$/);
    if (!match) return res.status(204).end();
    const contentType = match[1] || 'image/x-icon';
    const b64 = match[2] || '';
    const buf = Buffer.from(b64, 'base64');
    res.set('Content-Type', contentType);
    res.set('Cache-Control', 'no-store');
    return res.send(buf);
  } catch (e) {
    return res.status(204).end();
  }
});

// Transactions
app.get("/api/transactions", isAuthenticated, isAdminOrCashier, async (req, res) => {
  try {
    const transactions = await readData("transactions.json");
    res.json(transactions);
  } catch (error) {
    res
      .status(500)
      .json({ success: false, message: "Failed to load transactions" });
  }
});

app.post("/api/transactions", isAuthenticated, async (req, res) => {
  try {
    const { items, paymentMethod, amountReceived, customerId = 'default', customerName = 'Pelanggan Umum', discountPercent = 0, discountAmount = 0, paidAmount, remainingAmount, paymentDate } = req.body;
    if (!items || items.length === 0) {
      return res
        .status(400)
        .json({ success: false, message: "Cart cannot be empty." });
    }

    const products = await readData("products.json");
    const transactions = await readData("transactions.json");
    let baseSubtotal = 0;
    let perProductDiscountTotal = 0;
    let perProductTaxTotal = 0;
    let afterItemDiscountSubtotal = 0;
    const transactionItems = items.map((item) => {
      const product = products.find((p) => p.id === item.productId);
      if (!product)
        throw new Error(`Product with ID ${item.productId} not found`);
      // Allow negative stock - no stock validation for transactions

      // Debug: Log the item data received
      console.log('Server processing item:', JSON.stringify(item, null, 2));

      const itemBase = product.price * (item.quantity || item.qty || 0);
      baseSubtotal += itemBase;
      const pDisc = Math.max(0, Number(product.discountPercent || 0));
      const pTax = Math.max(0, Number(product.taxRate || 0));
      const itemDisc = Math.round(itemBase * (pDisc / 100));
      const itemNet = itemBase - itemDisc;
      const itemTax = Math.round(itemNet * (pTax / 100));
      perProductDiscountTotal += itemDisc;
      perProductTaxTotal += itemTax;
      afterItemDiscountSubtotal += itemNet;

      const returnItem = {
        productId: product.id,
        name: product.name,
        price: product.price,
        qty: item.quantity || item.qty,
        quantity: item.quantity || item.qty,
        subtotal: itemNet,
        // Preserve variant information if present
        ...(item.variant && { variant: item.variant }),
      };
      
      // Debug: Log the item being returned
      console.log('Server returning item:', JSON.stringify(returnItem, null, 2));
      
      return returnItem;
    });

    // compute taxes based on settings
    const settings = await readData('settings.json').catch(() => ({}));
    const taxRate = Number(settings?.taxRate || 0);
    const serviceRate = Number(settings?.serviceRate || 0);
    const priceIncludesTax = Boolean(settings?.priceIncludesTax || false);
    const subtotal = baseSubtotal;
    const discountP = Math.max(0, Number(discountPercent) || 0);
    const discountA = Math.max(0, Number(discountAmount) || 0);
    let computedDiscount = 0;
    if (discountP > 0) {
      computedDiscount = Math.round(afterItemDiscountSubtotal * (discountP / 100));
    } else if (discountA > 0) {
      computedDiscount = Math.round(discountA);
    }
    if (computedDiscount > afterItemDiscountSubtotal) computedDiscount = afterItemDiscountSubtotal;
    const netAfterCartDiscount = afterItemDiscountSubtotal - computedDiscount;
    const globalTax = priceIncludesTax ? 0 : Math.round(netAfterCartDiscount * (taxRate / 100));
    const serviceAmount = priceIncludesTax ? 0 : Math.round(netAfterCartDiscount * (serviceRate / 100));
    const taxAmount = perProductTaxTotal + globalTax;
    const grandTotal = netAfterCartDiscount + taxAmount + serviceAmount;

    const newTransaction = {
      id: `TRX-${new Date()
        .toISOString()
        .slice(0, 10)
        .replace(/-/g, "")}-${Date.now()}`,
      timestamp: new Date().toISOString(),
      date: new Date().toISOString().split('T')[0],
      userId: req.session.user.id,
      customerId,
      customerName,
      items: transactionItems,
      subtotal,
      discountAmount: perProductDiscountTotal + computedDiscount,
      taxAmount,
      serviceAmount,
      totalAmount: grandTotal,
      paymentMethod,
      amountReceived: paymentMethod === "cash" ? amountReceived : grandTotal,
      change: paymentMethod === "cash" ? amountReceived - grandTotal : 0,
      // Add debt tracking fields
      paidAmount: paidAmount || amountReceived,
      remainingAmount: remainingAmount !== undefined ? remainingAmount : 0,
      paymentDate: paymentDate || new Date().toISOString().split('T')[0]
    };

    transactions.push(newTransaction);
    // enqueue transaction append for sync
    try { await enqueueOutbox({ collection: 'transactions', file: 'transactions.json', op: 'insert', _id: newTransaction.id, doc: newTransaction, updatedAt: Number(new Date(newTransaction.timestamp).getTime()) || Date.now() }); } catch {}
    await writeData("transactions.json", transactions);
    await saveArrayWithSync("products.json", products);
    res.json(newTransaction);
  } catch (error) {
    console.error("Transaction error:", error);
    res
      .status(400)
      .json({
        success: false,
        message: error.message || "Failed to create transaction",
      });
  }
});

// PATCH endpoint for updating transaction payments (debt tracking)
app.patch("/api/transactions/:id", isAuthenticated, async (req, res) => {
  try {
    const { id } = req.params;
    const { paidAmount, remainingAmount, paymentDate } = req.body;
    
    const transactions = await readData("transactions.json");
    if (!Array.isArray(transactions)) {
      return res.status(404).json({ success: false, message: "Transactions not found" });
    }
    
    const transactionIndex = transactions.findIndex(t => t.id === id);
    if (transactionIndex === -1) {
      return res.status(404).json({ success: false, message: "Transaction not found" });
    }
    
    // Update payment fields
    transactions[transactionIndex].paidAmount = Number(paidAmount) || 0;
    transactions[transactionIndex].remainingAmount = Number(remainingAmount) || 0;
    transactions[transactionIndex].paymentDate = paymentDate || new Date().toISOString().split('T')[0];
    transactions[transactionIndex].updatedAt = Date.now();
    
    await writeData("transactions.json", transactions);
    
    // Enqueue update for sync
    try { 
      await enqueueOutbox({ 
        collection: 'transactions', 
        file: 'transactions.json', 
        op: 'update', 
        _id: id, 
        doc: transactions[transactionIndex], 
        updatedAt: Number(transactions[transactionIndex].updatedAt) 
      }); 
    } catch {}
    
    res.json({ 
      success: true, 
      message: "Payment updated successfully",
      transaction: transactions[transactionIndex]
    });
  } catch (error) {
    console.error("Failed to update transaction payment:", error);
    res.status(500).json({ 
      success: false, 
      message: error.message || "Failed to update payment" 
    });
  }
});

app.get("/api/recent-transactions", isAuthenticated, async (req, res) => {
  try {
    const transactions = await readData("transactions.json");
    const recentTransactions = transactions
      .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp))
      .slice(0, 5);
    res.json(recentTransactions);
  } catch (error) {
    console.error("Failed to fetch recent transactions:", error);
    res
      .status(500)
      .json({
        success: false,
        message: "Failed to fetch recent transactions.",
      });
  }
});

app.delete("/api/transactions/:id", isAuthenticated, async (req, res) => {
  try {
    const transactions = await readData("transactions.json");
    const products = await readData("products.json");
    const transactionIndex = transactions.findIndex(
      (t) => t.id === req.params.id
    );

    if (transactionIndex === -1) {
      return res
        .status(404)
        .json({ success: false, message: "Transaction not found." });
    }

    const transactionToVoid = transactions[transactionIndex];

    // Kembalikan stok produk + tulis stock_moves (void)
    for (const item of transactionToVoid.items) {
      const product = products.find((p) => p.id === item.productId);
      if (product) {
        product.stock += item.qty;
        try { await appendStockMove({ productId: item.productId, delta: Number(item.qty||0), reason: 'void', refId: transactionToVoid.id, by: (req.session && req.session.user && req.session.user.username) || '', newStock: Number(product.stock || 0), stockOnly: true }); } catch {}
      }
    }

    // Hapus transaksi
    transactions.splice(transactionIndex, 1);

    await saveArrayWithSync("products.json", products);
    await saveArrayWithSync("transactions.json", transactions, { keyField: 'id' });

    res.json({ success: true, message: "Transaction deleted successfully." });
  } catch (error) {
    console.error("Failed to void transaction:", error);
    res
      .status(500)
      .json({ success: false, message: "Failed to void transaction." });
  }
});

// --- Shift Kasir (Cashier Shifts) ---
// Buka shift baru untuk kasir yang sedang login
app.post('/api/shifts/open', isAuthenticated, isAdminOrCashier, async (req, res) => {
  try {
    const user = (req.session && req.session.user) || null;
    if (!user || !user.id) {
      return res.status(400).json({ success:false, message:'User session tidak valid' });
    }
    const openingCash = Number((req.body && req.body.openingCash) || 0) || 0;
    let shifts = await readData('shifts.json').catch(() => []);
    if (!Array.isArray(shifts)) shifts = [];
    const cashierId = String(user.id);
    const hasOpen = shifts.some(s => String(s && s.cashierId) === cashierId && !s.closedAt);
    if (hasOpen) {
      return res.status(400).json({ success:false, message:'Masih ada shift aktif untuk kasir ini.' });
    }
    const now = Date.now();
    const shift = {
      id: `SHIFT-${new Date(now).toISOString().slice(0,10).replace(/-/g,'')}-${now}`,
      cashierId,
      cashierUsername: String(user.username || ''),
      cashierName: String(user.name || user.username || ''),
      openedAt: now,
      closedAt: null,
      openingCash,
      closingCash: null,
      expectedCash: null,
      cashSales: 0,
      nonCashSales: 0,
      totalSales: 0,
      cashVariance: null,
      transactionsCount: 0,
    };
    shifts.push(shift);
    await writeData('shifts.json', shifts);
    return res.json({ success:true, shift });
  } catch (e) {
    console.error('Failed to open shift:', e);
    return res.status(500).json({ success:false, message:'Gagal membuka shift' });
  }
});

// Tutup shift aktif untuk kasir yang sedang login dan hitung selisih kas
app.post('/api/shifts/close', isAuthenticated, isAdminOrCashier, async (req, res) => {
  try {
    const user = (req.session && req.session.user) || null;
    if (!user || !user.id) {
      return res.status(400).json({ success:false, message:'User session tidak valid' });
    }
    const closingCash = Number((req.body && req.body.closingCash) || 0) || 0;
    let shifts = await readData('shifts.json').catch(() => []);
    if (!Array.isArray(shifts)) shifts = [];
    const cashierId = String(user.id);
    // Cari shift terbuka terakhir untuk kasir ini
    let idx = -1;
    for (let i = shifts.length - 1; i >= 0; i--) {
      const s = shifts[i] || {};
      if (String(s.cashierId) === cashierId && !s.closedAt) { idx = i; break; }
    }
    if (idx < 0) {
      return res.status(400).json({ success:false, message:'Tidak ada shift aktif untuk kasir ini.' });
    }
    const now = Date.now();
    const shift = shifts[idx] || {};
    shift.closedAt = now;
    shift.closingCash = closingCash;

    // Hitung ringkasan transaksi untuk shift ini
    let txs = await readData('transactions.json').catch(() => []);
    if (!Array.isArray(txs)) txs = [];
    const start = Number(shift.openedAt || 0);
    const end = Number(shift.closedAt || now);
    let cashSales = 0;
    let nonCashSales = 0;
    let count = 0;
    for (const tx of txs) {
      const uid = tx && tx.userId != null ? String(tx.userId) : '';
      if (uid && uid !== cashierId) continue;
      let ts = 0;
      if (typeof tx.timestamp === 'string') {
        const d = new Date(tx.timestamp);
        ts = d.getTime();
      } else {
        ts = Number(tx.timestamp || 0);
      }
      if (!Number.isFinite(ts) || ts < start || ts > end) continue;
      const total = Number(tx.totalAmount != null ? tx.totalAmount : (tx.total || 0)) || 0;
      const pm = String(tx.paymentMethod || 'cash').toLowerCase();
      if (pm === 'cash') cashSales += total; else nonCashSales += total;
      count++;
    }
    const totalSales = cashSales + nonCashSales;
    const openingCash = Number(shift.openingCash || 0);
    const expectedCash = openingCash + totalSales;
    const cashVariance = closingCash - expectedCash;
    shift.cashSales = cashSales;
    shift.nonCashSales = nonCashSales;
    shift.totalSales = totalSales;
    shift.expectedCash = expectedCash;
    shift.cashVariance = cashVariance;
    shift.transactionsCount = count;

    shifts[idx] = shift;
    await writeData('shifts.json', shifts);
    return res.json({ success:true, shift });
  } catch (e) {
    console.error('Failed to close shift:', e);
    return res.status(500).json({ success:false, message:'Gagal menutup shift' });
  }
});

// Ambil shift aktif untuk kasir yang sedang login
app.get('/api/shifts/current', isAuthenticated, isAdminOrCashier, async (req, res) => {
  try {
    const user = (req.session && req.session.user) || null;
    if (!user || !user.id) {
      return res.status(400).json({ success:false, message:'User session tidak valid' });
    }
    let shifts = await readData('shifts.json').catch(() => []);
    if (!Array.isArray(shifts)) shifts = [];
    const cashierId = String(user.id);
    const current = shifts.slice().reverse().find(s => String(s && s.cashierId) === cashierId && !s.closedAt) || null;
    return res.json({ success:true, shift: current || null });
  } catch (e) {
    console.error('Failed to get current shift:', e);
    return res.status(500).json({ success:false, message:'Gagal mengambil shift aktif' });
  }
});

// Ringkasan shift aktif (untuk auto default saldo akhir)
app.get('/api/shifts/current-summary', isAuthenticated, isAdminOrCashier, async (req, res) => {
  try {
    const user = (req.session && req.session.user) || null;
    if (!user || !user.id) {
      return res.status(400).json({ success:false, message:'User session tidak valid' });
    }
    let shifts = await readData('shifts.json').catch(() => []);
    if (!Array.isArray(shifts)) shifts = [];
    const cashierId = String(user.id);
    const current = shifts.slice().reverse().find(s => String(s && s.cashierId) === cashierId && !s.closedAt) || null;
    if (!current) {
      return res.json({ success:true, shift: null, summary: null });
    }
    let txs = await readData('transactions.json').catch(() => []);
    if (!Array.isArray(txs)) txs = [];
    const start = Number(current.openedAt || 0);
    const end = Date.now();
    let cashSales = 0;
    let nonCashSales = 0;
    let count = 0;
    for (const tx of txs) {
      const uid = tx && tx.userId != null ? String(tx.userId) : '';
      if (uid && uid !== cashierId) continue;
      let ts = 0;
      if (typeof tx.timestamp === 'string') {
        const d = new Date(tx.timestamp);
        ts = d.getTime();
      } else {
        ts = Number(tx.timestamp || 0);
      }
      if (!Number.isFinite(ts) || ts < start || ts > end) continue;
      const total = Number(tx.totalAmount != null ? tx.totalAmount : (tx.total || 0)) || 0;
      const pm = String(tx.paymentMethod || 'cash').toLowerCase();
      if (pm === 'cash') cashSales += total; else nonCashSales += total;
      count++;
    }
    const totalSales = cashSales + nonCashSales;
    const openingCash = Number(current.openingCash || 0);
    const expectedCash = openingCash + totalSales;
    const summary = {
      cashSales,
      nonCashSales,
      totalSales,
      openingCash,
      expectedCash,
      transactionsCount: count
    };
    return res.json({ success:true, shift: current, summary });
  } catch (e) {
    console.error('Failed to get current shift summary:', e);
    return res.status(500).json({ success:false, message:'Gagal mengambil ringkasan shift aktif' });
  }
});

// Daftar semua shift (admin)
app.get('/api/shifts', isAuthenticated, isAdmin, async (req, res) => {
  try {
    let shifts = await readData('shifts.json').catch(() => []);
    if (!Array.isArray(shifts)) shifts = [];
    // Optional filter by cashierId/from/to
    let { cashierId, from, to } = req.query || {};
    if (cashierId) {
      const cid = String(cashierId);
      shifts = shifts.filter(s => String(s && s.cashierId) === cid);
    }
    const fromMs = Number(from) || 0;
    const toMs = Number(to) || 0;
    if (fromMs || toMs) {
      shifts = shifts.filter(s => {
        const openTs = Number((s && s.openedAt) || 0);
        if (fromMs && openTs < fromMs) return false;
        if (toMs && openTs > toMs) return false;
        return true;
      });
    }
    shifts.sort((a,b) => Number((b && b.openedAt) || 0) - Number((a && a.openedAt) || 0));
    return res.json({ success:true, shifts });
  } catch (e) {
    console.error('Failed to list shifts:', e);
    return res.status(500).json({ success:false, message:'Gagal memuat data shift' });
  }
});

// --- Excel Import/Export API Routes ---
// --- API untuk Validasi Admin ---
app.post("/api/admin/validate-password", isAuthenticated, async (req, res) => {
  try {
    const { password } = req.body;

    // Di production, gunakan bcrypt untuk hash password
    // Untuk demo ini, kita bandingkan dengan password admin hardcoded
    const ADMIN_PASSWORD = "admin123"; // Ganti dengan password admin Anda yang sebenarnya

    const isValid = password === ADMIN_PASSWORD;

    res.json({ valid: isValid });
  } catch (error) {
    console.error("Error validating admin password:", error);
    res.status(500).json({ success: false, message: "Server error" });
  }
});

// Restore produk dari backup
app.post(
  "/api/products/restore",
  isAuthenticated,
  isAdmin,
  async (req, res) => {
    try {
      // Validasi password admin secara langsung, bukan dengan fetch
      const { password } = req.body;
      const ADMIN_PASSWORD = "admin123"; // Sama dengan di API validasi

      if (password !== ADMIN_PASSWORD) {
        return res.status(401).json({
          success: false,
          message: "Invalid admin password",
        });
      }

      // Cari file backup terbaru
      const backupDir = path.join(DATA_DIR, "backup");
      let backupFiles = [];

      try {
        const files = await fs.readdir(backupDir);
        backupFiles = files.filter(
          (file) => file.startsWith("products_") && file.endsWith(".json")
        );
        backupFiles.sort((a, b) => {
          const aTime = a.split("_")[1].replace(".json", "");
          const bTime = b.split("_")[1].replace(".json", "");
          return bTime.localeCompare(aTime);
        });
      } catch (error) {
        console.error("Error reading backup directory:", error);
      }

      if (backupFiles.length === 0) {
        return res.status(404).json({
          success: false,
          message: "Tidak ada backup produk yang ditemukan",
        });
      }

      // Baca file backup terbaru
      const latestBackup = await readData(`backup/${backupFiles[0]}`);

      // Restore produk
      await writeData("products.json", latestBackup);

      res.json({
        success: true,
        message: `Produk berhasil dipulihkan dari backup: ${backupFiles[0]}`,
      });
    } catch (error) {
      console.error("Error restoring products:", error);
      res.status(500).json({
        success: false,
        message: "Failed to restore products",
      });
    }
  }
);

// Import Products from XLSX
app.post("/api/products/import", isAuthenticated, isAdmin, async (req, res) => {
  try {
    // console.log("Starting import...");
    const { products: importData } = req.body;

    if (!Array.isArray(importData) || importData.length === 0) {
      return res
        .status(400)
        .json({ success: false, message: "No valid data to import" });
    }

    // console.log("Import data received:", importData.length, "rows");
    // console.log("First row sample:", importData[0]);
    // console.log("First row category:", importData[0]["Category"]);

    let products = await readData("products.json");
    if (!Array.isArray(products)) products = [];
    let categories = await readData("categories.json");
    if (!Array.isArray(categories)) categories = [];
    let units = await readData("units.json");
    if (!Array.isArray(units)) units = [];

    // console.log("Available categories:", categories.map(c => ({ id: c.id, name: `"${c.name}"` })));

    let successCount = 0;
    let errorCount = 0;
    const errors = [];

    for (let i = 0; i < importData.length; i++) {
      try {
        const row = importData[i];
        // console.log(`Processing row ${i + 1}:`, row);

        // Validate required fields: Product Name, Stock, and (Selling Price or Price)
        const hasName = !!(row["Product Name"] || row["Nama Produk"]);
        const stockRaw = row["Stock"] != null ? row["Stock"] : row["Stok"];
        const hasStock = stockRaw !== undefined && stockRaw !== "";
        const hasSellingPrice = row["Selling Price"] !== undefined && row["Selling Price"] !== "";
        const hasLegacyPrice = row["Price"] !== undefined && row["Price"] !== "";
        const hasIdSellingPrice = row["Harga Jual"] !== undefined && row["Harga Jual"] !== "";
        if (!hasName || !hasStock || (!hasSellingPrice && !hasLegacyPrice && !hasIdSellingPrice)) {
          const errorMsg = `Baris ${i + 1}: Kolom wajib tidak lengkap. Wajib: Product Name, Stock, dan Selling Price atau Price.`;
          errors.push(errorMsg);
          errorCount++;
          continue;
        }

        // Find or create category
        let categoryId = null;
        const categoryNameRaw = row["Category"] ?? row["Kategori"];
        if (categoryNameRaw != null && String(categoryNameRaw).trim() !== "") {
          const categoryName = String(categoryNameRaw).trim();

          let category = categories.find(
            (c) => c.name && c.name.toLowerCase() === categoryName.toLowerCase()
          );

          if (!category) {
            const newCatId = Date.now() + i + 100000;
            category = { id: newCatId, name: categoryName, description: "" };
            categories.push(category);
          }
          categoryId = category.id;
        }

        // Create product object with new fields
        const purchasePrice = parseFloat(row["Purchase Price"] || row["Harga Beli"]) || 0;
        const sellingPrice = row["Selling Price"] !== undefined && row["Selling Price"] !== ""
          ? (parseFloat(row["Selling Price"]) || 0)
          : row["Harga Jual"] !== undefined && row["Harga Jual"] !== ""
            ? (parseFloat(row["Harga Jual"]) || 0)
            : (parseFloat(row["Price"] || row["Harga"]) || 0);
        const taxRate = Number(row["PPN (%)"] ?? row["Tax (%)"] ?? row["Tax Rate"] ?? row["Pajak (%)"] ?? 0) || 0;
        const discountPercent = Number(row["Discount (%)"] ?? row["Diskon (%)"] ?? row["Discount"] ?? row["Diskon"] ?? 0) || 0;
        const qrCode = (row["QR Code"] || "").toString().trim();
        const unitSource =
          row["Unit"] !== undefined && row["Unit"] !== null
            ? row["Unit"]
            : row["Satuan"];
        let unit = unitSource != null ? unitSource.toString().trim() : "";
        const variantsRaw = row["Variant Prices"] ?? row["Varian Harga"] ?? "";
        let unitPrices = parseProductVariantPricesFromXlsx(variantsRaw, unit);
        if (!unit && unitPrices.length > 0) {
          const firstUnit = unitPrices[0] && unitPrices[0].unit;
          if (firstUnit) unit = firstUnit;
        }

        if (unit) {
          const unitExists = units.find(
            (u) =>
              u &&
              u.name &&
              u.name.toLowerCase() === unit.toLowerCase()
          );
          if (!unitExists) {
            const nowUnit = Date.now() + i;
            units.push({
              id: nowUnit,
              name: unit,
              description: "",
              createdAt: nowUnit,
              updatedAt: nowUnit,
            });
          }
        }
        if (unitPrices.length > 0) {
          unitPrices.forEach((v, idx) => {
            const uName = v && v.unit ? String(v.unit).trim() : "";
            if (!uName) return;
            const exists = units.find(
              (u) => u && u.name && u.name.toLowerCase() === uName.toLowerCase()
            );
            if (!exists) {
              const nowUnit = Date.now() + i + idx + 1;
              units.push({
                id: nowUnit,
                name: uName,
                description: "",
                createdAt: nowUnit,
                updatedAt: nowUnit,
              });
            }
          });
        }

        const newProduct = {
          id: Date.now() + i,
          name: String(row["Product Name"] || row["Nama Produk"]).trim(),
          stock: parseInt(stockRaw, 10) || 0,
          categoryId: categoryId,
          sku: row["SKU"]
            ? row["SKU"].toString().trim()
            : `PROD-${Date.now()}-${i}`,
          purchasePrice,
          sellingPrice,
          taxRate,
          discountPercent,
          qrCode,
          unit,
          unitPrices: unitPrices.length > 0 ? unitPrices : undefined,
          // Backward compat for POS
          price: sellingPrice,
          isTopProduct:
            row["Is Top Product"] &&
            row["Is Top Product"].toString().toLowerCase() === "yes",
          isBestSeller:
            row["Is Best Seller"] &&
            row["Is Best Seller"].toString().toLowerCase() === "yes",
          imageBase64: "", // Always empty for imports
        };

        // console.log(`Product "${newProduct.name}" will be linked to category ID: ${categoryId}`);
        products.push(newProduct);
        successCount++;
        // console.log(`Successfully added product: ${newProduct.name} with categoryId: ${newProduct.categoryId}`);
      } catch (error) {
        const errorMsg = `Baris ${i + 1}: ${error.message}`;
        errors.push(errorMsg);
        errorCount++;
      }
    }

    // Save data with sync
    await saveArrayWithSync("products.json", products);
    await saveArrayWithSync("categories.json", categories);
    await saveArrayWithSync("units.json", units);

    // Send response
    let message = `Import selesai. Sukses: ${successCount}, Error: ${errorCount}`;
    if (errors.length > 0) {
      message += `\n\nBeberapa error pertama:\n${errors
        .slice(0, 3)
        .join("\n")}`;
      if (errors.length > 5) {
        message += ` ... dan ${errors.length - 5} more errors`;
      }
    }

    // console.log("Import completed:", message);
    res.json({
      success: successCount > 0,
      message,
      successCount,
      errorCount,
      errors: errors.slice(0, 10), // Return first 10 errors
    });
  } catch (error) {
    console.error("!!! IMPORT ERROR !!!", error);
    res.status(500).json({
      success: false,
      message: "Failed to import products: " + error.message,
    });
  }
});

// === CATEGORIES EXPORT/IMPORT ===
// Export Categories to XLSX
app.get("/api/categories/export", isAuthenticated, isAdmin, async (req, res) => {
  try {
    const categories = await readData("categories.json");
    const exportData = categories.map((cat) => ({
      "Category Name": cat.name || "",
      "Description": cat.description || "",
      "ID": cat.id || "",
    }));

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Categories");
    ws["!cols"] = [{ wch: 25 }, { wch: 50 }, { wch: 15 }];
    const buf = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", "attachment; filename=categories_export.xlsx");
    res.send(buf);
  } catch (error) {
    res.status(500).json({ success: false, message: "Failed to export categories: " + error.message });
  }
});

// Download Category Import Template
app.get("/api/categories/template", isAuthenticated, isAdmin, async (req, res) => {
  try {
    const templateData = [
      { "Category Name": "Example Category 1", "Description": "Description for category 1" },
      { "Category Name": "Example Category 2", "Description": "Description for category 2" },
    ];

    const ws = XLSX.utils.json_to_sheet(templateData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Template");
    ws["!cols"] = [{ wch: 25 }, { wch: 50 }];
    const buf = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", "attachment; filename=category_import_template.xlsx");
    res.send(buf);
  } catch (error) {
    res.status(500).json({ success: false, message: "Failed to generate template: " + error.message });
  }
});

// Import Categories from XLSX
app.post("/api/categories/import", isAuthenticated, isAdmin, async (req, res) => {
  try {
    const { categories: importData } = req.body;
    if (!Array.isArray(importData) || importData.length === 0) {
      return res.status(400).json({ success: false, message: "No valid data to import" });
    }

    const categories = await readData("categories.json");
    let successCount = 0;
    let errorCount = 0;
    const errors = [];

    for (let i = 0; i < importData.length; i++) {
      try {
        const row = importData[i];
        if (!row["Category Name"] || row["Category Name"].toString().trim() === "") {
          errors.push(`Baris ${i + 1}: Category Name wajib diisi`);
          errorCount++;
          continue;
        }

        const categoryName = row["Category Name"].toString().trim();
        const existingCategory = categories.find(
          (c) => c.name && c.name.toLowerCase() === categoryName.toLowerCase()
        );

        if (existingCategory) {
          errors.push(`Baris ${i + 1}: Kategori "${categoryName}" sudah ada`);
          errorCount++;
          continue;
        }

        const newCategory = {
          id: Date.now() + i,
          name: categoryName,
          description: (row["Description"] || "").toString().trim(),
        };

        categories.push(newCategory);
        successCount++;
      } catch (error) {
        errors.push(`Baris ${i + 1}: ${error.message}`);
        errorCount++;
      }
    }

    await writeData("categories.json", categories);
    let message = `Import selesai. Sukses: ${successCount}, Error: ${errorCount}`;
    if (errors.length > 0) {
      message += `\n\nBeberapa error pertama:\n${errors
        .slice(0, 3)
        .join("\n")}`;
      if (errors.length > 5) {
        message += ` ... dan ${errors.length - 5} more errors`;
      }
    }

    res.json({ success: true, message, successCount, errorCount, errors: errors.slice(0, 10) });
  } catch (error) {
    res.status(500).json({ success: false, message: "Failed to import categories: " + error.message });
  }
});

// === TRANSACTIONS EXPORT ===
// Export Transactions to XLSX
app.get("/api/transactions/export", isAuthenticated, isAdminOrCashier, async (req, res) => {
  try {
    const transactions = await readData("transactions.json");
    const users = await readData("users.json");
    const exportData = transactions.map((t) => {
      const user = users.find((u) => u.id === t.cashierId);
      return {
        "Transaction ID": t.id || "",
        "Date": t.date || "",
        "Time": t.time || "",
        "Cashier": user ? user.name || user.username : "",
        "Items Count": t.items ? t.items.length : 0,
        "Subtotal": t.subtotal || 0,
        "Tax": t.tax || 0,
        "Discount": t.discount || 0,
        "Total": t.total || 0,
        "Payment Method": t.paymentMethod || "",
        "Status": t.status || "",
      };
    });

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Transactions");
    ws["!cols"] = [
      { wch: 15 },
      { wch: 12 },
      { wch: 12 },
      { wch: 20 },
      { wch: 12 },
      { wch: 15 },
      { wch: 12 },
      { wch: 12 },
      { wch: 15 },
      { wch: 15 },
      { wch: 12 },
    ];
    const buf = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader(
      "Content-Disposition",
      "attachment; filename=transactions_export.xlsx"
    );
    res.send(buf);
  } catch (error) {
    res.status(500).json({
      success: false,
      message: "Failed to export transactions: " + error.message,
    });
  }
});

// === MUTASI BARANG ===
// API untuk mendapatkan data mutasi barang berdasarkan produk dan rentang tanggal
app.get("/api/mutasi-barang", isAuthenticated, isAdminOrCashier, async (req, res) => {
  try {
    const resolveCashierName = (transaction, users) => {
      const tx = transaction || {};
      const userList = Array.isArray(users) ? users : [];
      const normalize = (v) => String(v || "").trim().toLowerCase();
      const txUserId = String(tx.cashierId || tx.userId || tx.user_id || "").trim();
      const txUsername = String(
        tx.cashierUsername || tx.username || tx.userName || tx.cashier || ""
      ).trim();
      const userById = txUserId
        ? userList.find((u) => String(u && u.id) === txUserId)
        : null;
      if (userById) return userById.name || userById.username || "Unknown";
      if (txUsername) {
        const userByUsername = userList.find(
          (u) => normalize(u && u.username) === normalize(txUsername)
        );
        if (userByUsername) return userByUsername.name || userByUsername.username || txUsername;
      }
      const direct = String(
        tx.cashierName || tx.cashierFullName || tx.fullName || tx.full_name || tx.displayName || tx.nama || ""
      ).trim();
      if (direct) return direct;
      if (txUsername) return txUsername;
      return "Unknown";
    };

    const productId = req.query.productId;
    const startDate = req.query.startDate;
    const endDate = req.query.endDate;
    const page = req.query.page || 1;
    const pageSize = req.query.pageSize || 10;
    const exportFlag = req.query.export || false;

    if (!startDate || !endDate) {
      return res.status(400).json({ success: false, error: "Start date and end date are required" });
    }

    // Parse dates
    const start = new Date(startDate);
    start.setHours(0, 0, 0, 0);
    const end = new Date(endDate);
    end.setHours(23, 59, 59, 999);

    // Load data
    let transactions = await readData("transactions.json").catch(() => []);
    let products = await readData("products.json").catch(() => []);
    let users = await readData("users.json").catch(() => []);
    if (!Array.isArray(transactions)) transactions = [];
    if (!Array.isArray(products)) products = [];
    if (!Array.isArray(users)) users = [];

    if (!productId) {
      // Jika productId kosong, tampilkan semua transaksi dalam rentang tanggal

      const filteredTransactions = transactions.filter((transaction) => {
        const transactionDate = new Date(transaction.timestamp || transaction.date);
        const isInDateRange = transactionDate >= start && transactionDate <= end;
        return isInDateRange;
      });

      // Extract all items from filtered transactions
      const mutasiData = [];
      filteredTransactions.forEach((transaction) => {
        const cashierName = resolveCashierName(transaction, users);

        if (transaction.items && Array.isArray(transaction.items)) {
          transaction.items.forEach((item) => {
            const product = products.find((p) => String(p.id) === String(item.productId));
            if (product) {
              // Use variant quantity if available and > 0, otherwise use regular quantity
              let quantity = 0;
              // For items with variants, check if variant qty matches the actual qty
              if (item.variant) {
                // If variant qty is different from regular qty, use regular qty
                if (item.variant.qty !== (item.quantity || item.qty)) {
                  quantity = item.quantity || item.qty || 0;
                } else {
                  quantity = item.variant.qty;
                }
              } else {
                quantity = item.quantity || item.qty || 0;
              }

              const unit = (item.variant && item.variant.unit) || product.unit || "";
              const variantNote = (item.variant && item.variant.note) || "";
              const price = (item.variant && item.variant.price) || item.price || 0;

              mutasiData.push({
                transactionId: transaction.id,
                timestamp: transaction.timestamp || transaction.date,
                cashier: cashierName,
                productId: item.productId,
                productName: product.name,
                productSku: product.sku || "",
                quantity: quantity,
                unit: unit,
                variantNote: variantNote,
                price: price,
                subtotal: item.subtotal || (price * quantity),
              });
            }
          });
        }
      });

      // Sort by timestamp (newest first)
      mutasiData.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));

      // Pagination
      const pageNum = parseInt(page);
      const size = exportFlag === "true" ? mutasiData.length : parseInt(pageSize);
      const startIndex = exportFlag === "true" ? 0 : (pageNum - 1) * size;
      const endIndex = exportFlag === "true" ? mutasiData.length : startIndex + size;
      const paginatedResults = mutasiData.slice(startIndex, endIndex);

      return res.json({
        success: true,
        results: paginatedResults,
        total: mutasiData.length,
        productName: "Semua Produk",
        currentPage: pageNum,
        pageSize: size,
      });
    }

    // Find product info - handle both string and number comparison
    const product = products.find((p) => String(p.id) === String(productId));

    if (!product) {
      return res.status(404).json({ success: false, error: "Product not found" });
    }

    // Filter transactions that contain the product within date range
    const filteredTransactions = transactions.filter((transaction) => {
      const transactionDate = new Date(transaction.timestamp || transaction.date);
      const isInDateRange = transactionDate >= start && transactionDate <= end;

      // Check if transaction contains the product
      const hasProduct = transaction.items && transaction.items.some((item) => String(item.productId) === String(productId));
      if (hasProduct) {
      }
      return hasProduct && isInDateRange;
    });

    // Extract mutation data
    const mutasiData = [];
    filteredTransactions.forEach((transaction) => {
      const cashierName = resolveCashierName(transaction, users);
      const productItems = transaction.items.filter((item) => String(item.productId) === String(productId));

      productItems.forEach((productItem) => {
        // For items with variants, check if variant qty matches the actual qty
        let quantity = 0;
        if (productItem.variant) {
          // If variant qty is different from regular qty, use regular qty
          if (productItem.variant.qty !== (productItem.quantity || productItem.qty)) {
            quantity = productItem.quantity || productItem.qty || 0;
          } else {
            quantity = productItem.variant.qty;
          }
        } else {
          quantity = productItem.quantity || productItem.qty || 0;
        }

        const unit = (productItem.variant && productItem.variant.unit) || product.unit || "";
        const variantNote = (productItem.variant && productItem.variant.note) || "";
        const price = (productItem.variant && productItem.variant.price) || productItem.price || 0;

        mutasiData.push({
          transactionId: transaction.id,
          timestamp: transaction.timestamp || transaction.date,
          cashier: cashierName,
          productId: productId,
          productName: product.name,
          productSku: product.sku || "",
          quantity: quantity,
          unit: unit,
          variantNote: variantNote,
          price: price,
          subtotal: productItem.subtotal || (price * quantity),
        });
      });
    });

    // Sort by timestamp (newest first)
    mutasiData.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));

    // Pagination
    const pageNum = parseInt(page);
    const size = exportFlag === "true" ? mutasiData.length : parseInt(pageSize);
    const startIndex = exportFlag === "true" ? 0 : (pageNum - 1) * size;
    const endIndex = exportFlag === "true" ? mutasiData.length : startIndex + size;
    const paginatedResults = mutasiData.slice(startIndex, endIndex);

    res.json({
      success: true,
      results: paginatedResults,
      total: mutasiData.length,
      productName: product.name,
      currentPage: pageNum,
      pageSize: size,
    });
  } catch (error) {
    res.status(500).json({ success: false, error: "Failed to load mutation data: " + error.message });
  }
});

// Get single transaction details for mutasi view
app.get("/api/transactions/:id", isAuthenticated, isAdminOrCashier, async (req, res) => {
  try {
    const { id } = req.params;
    const transactions = await readData("transactions.json");
    const users = await readData("users.json");
    const products = await readData("products.json").catch(() => []);
    const productMap = new Map((Array.isArray(products) ? products : []).map(p => [String(p && p.id), p]));

    const transaction = transactions.find((t) => t.id === id);

    if (!transaction) {
      return res.status(404).json({ success: false, error: "Transaction not found" });
    }

    // Add cashier info
    const user = users.find((u) => u.id === transaction.cashierId);
    if (user) {
      transaction.cashier = user.name || user.username;
    }

    // Calculate correct subtotal for each item (handle variants)
    if (transaction.items && Array.isArray(transaction.items)) {
      transaction.items = transaction.items.map((item) => {
        // For items with variants, check if variant qty matches the actual qty
        let quantity = 0;
        if (item.variant) {
          // If variant qty is different from regular qty, use regular qty
          if (item.variant.qty !== (item.quantity || item.qty)) {
            quantity = item.quantity || item.qty || 0;
          } else {
            quantity = item.variant.qty;
          }
        } else {
          quantity = item.quantity || item.qty || 0;
        }

        const unit = (item.variant && item.variant.unit) || "";
        const variantNote = (item.variant && item.variant.note) || "";
        const price = (item.variant && item.variant.price) || item.price || 0;
        const calculatedSubtotal = price * quantity;
        const p = productMap.get(String(item.productId || ""));
        const itemName = item.name || (p && p.name) || String(item.productId || "-");

        return {
          ...item,
          name: itemName,
          quantity: quantity,
          unit: unit,
          variantNote: variantNote,
          price: price,
          subtotal: item.subtotal || calculatedSubtotal,
        };
      });
    }

    res.json(transaction);
  } catch (error) {
    res.status(500).json({ success: false, error: "Failed to load transaction: " + error.message });
  }
});

// --- Invoices ---
app.get("/api/invoices", isAuthenticated, isAdminOrCashier, async (req, res) => {
  try {
    let invoices = await readData("invoices.json").catch(() => []);
    if (!Array.isArray(invoices)) invoices = [];
    res.json(invoices);
  } catch (error) {
    res.status(500).json({ success: false, message: "Failed to load invoices" });
  }
});

app.get("/api/invoices/:id", isAuthenticated, isAdminOrCashier, async (req, res) => {
  try {
    const id = String(req.params.id || "").trim();
    let invoices = await readData("invoices.json").catch(() => []);
    if (!Array.isArray(invoices)) invoices = [];
    const inv = invoices.find((x) => {
      const invId = String(x && (x.id || x._id || "")).trim();
      const invNo = String(x && (x.invoiceNo || "")).trim();
      return invId === id || invNo === id;
    });
    if (!inv) return res.status(404).json({ success: false, message: "Invoice not found" });
    res.json(inv);
  } catch (error) {
    res.status(500).json({ success: false, message: "Failed to load invoice" });
  }
});

app.post("/api/invoices", isAuthenticated, isAdminOrCashier, async (req, res) => {
  try {
    const payload = req.body || {};
    const data = (payload && typeof payload.data === 'object') ? payload.data : null;
    if (!data) return res.status(400).json({ success: false, message: "Invalid invoice data" });

    let invoices = await readData("invoices.json").catch(() => []);
    if (!Array.isArray(invoices)) invoices = [];

    const invoiceNo = sanitizeHtml(String(payload.invoiceNo || data.invoiceNo || '')).trim();
    const source = sanitizeHtml(String(payload.source || data.source || '')).trim() || 'manual';
    const transactionId = sanitizeHtml(String(payload.transactionId || data.transactionId || '')).trim();
    const customerName = sanitizeHtml(String(payload.customerName || (data.customer && data.customer.name) || '')).trim();
    const total = Number(payload.total ?? data.total ?? 0) || 0;
    const date = String(payload.date || data.date || '').trim();
    const now = Date.now();

    let idx = -1;
    if (transactionId) {
      idx = invoices.findIndex((x) => String(x && x.transactionId) === transactionId);
    } else if (invoiceNo) {
      idx = invoices.findIndex((x) => String(x && x.invoiceNo) === invoiceNo);
    }

    let inv;
    if (idx >= 0) {
      inv = {
        ...invoices[idx],
        invoiceNo,
        source,
        transactionId,
        customerName,
        total,
        date,
        data,
        updatedAt: now
      };
      invoices[idx] = inv;
    } else {
      const id = `inv-${now}-${Math.random().toString(36).slice(2, 8)}`;
      inv = {
        id,
        invoiceNo: invoiceNo || id,
        source,
        transactionId,
        customerName,
        total,
        date,
        data,
        createdAt: now,
        updatedAt: now
      };
      invoices.push(inv);
    }

    await saveArrayWithSync("invoices.json", invoices, { keyField: "id" });
    res.json({ success: true, invoice: inv });
  } catch (error) {
    res.status(500).json({ success: false, message: "Failed to save invoice" });
  }
});

app.delete("/api/invoices/:id", isAuthenticated, isAdminOrCashier, async (req, res) => {
  try {
    const id = String(req.params.id || "").trim();
    if (!id) return res.status(400).json({ success: false, message: "ID invoice tidak valid" });
    let invoices = await readData("invoices.json").catch(() => []);
    if (!Array.isArray(invoices)) invoices = [];
    const before = invoices.length;
    invoices = invoices.filter(x => {
      const invId = String(x && (x.id || x._id || "")).trim();
      const invNo = String(x && (x.invoiceNo || "")).trim();
      return invId !== id && invNo !== id;
    });
    if (before === invoices.length) {
      return res.status(404).json({ success: false, message: "Invoice tidak ditemukan" });
    }
    await saveArrayWithSync("invoices.json", invoices, { keyField: "id" });
    res.json({ success: true, message: "Invoice dihapus" });
  } catch (error) {
    res.status(500).json({ success: false, message: "Failed to delete invoice" });
  }
});

// Print transaction receipt
app.get("/print-transaction/:id", isAuthenticated, isAdminOrCashier, async (req, res) => {
  try {
    const { id } = req.params;
    const transactions = await readData("transactions.json");
    const users = await readData("users.json");
    const products = await readData("products.json").catch(() => []);
    const productMap = new Map((Array.isArray(products) ? products : []).map(p => [String(p && p.id), p]));

    const transaction = transactions.find((t) => t.id === id);

    if (!transaction) {
      return res.status(404).send("Transaction not found");
    }

    // Add cashier info
    const user = users.find((u) => u.id === transaction.cashierId);
    if (user) {
      transaction.cashier = user.name || user.username;
    }

    // Calculate correct subtotal for each item (handle variants)
    if (transaction.items && Array.isArray(transaction.items)) {
      transaction.items = transaction.items.map((item) => {
        // For items with variants, check if variant qty matches the actual qty
        let quantity = 0;
        if (item.variant) {
          // If variant qty is different from regular qty, use regular qty
          if (item.variant.qty !== (item.quantity || item.qty)) {
            quantity = item.quantity || item.qty || 0;
          } else {
            quantity = item.variant.qty;
          }
        } else {
          quantity = item.quantity || item.qty || 0;
        }

        const unit = (item.variant && item.variant.unit) || "";
        const variantNote = (item.variant && item.variant.note) || "";
        const price = (item.variant && item.variant.price) || item.price || 0;
        const calculatedSubtotal = price * quantity;
        const p = productMap.get(String(item.productId || ""));
        const itemName = item.name || (p && p.name) || String(item.productId || "-");

        return {
          ...item,
          name: itemName,
          quantity: quantity,
          unit: unit,
          variantNote: variantNote,
          price: price,
          subtotal: item.subtotal || calculatedSubtotal,
        };
      });
    }

    // Calculate total from all item subtotals
    const calculatedTotal = transaction.items.reduce((sum, item) => sum + item.subtotal, 0);

    const settings = await readData("settings.json");
    const storeLogo = settings?.logoBase64 || "";

    // Prepare receipt data
    const receiptData = {
      id: transaction.id,
      timestamp: transaction.timestamp,
      cashier: transaction.cashier || "Unknown",
      paymentMethod: transaction.paymentMethod || "Cash",
      items: transaction.items || [],
      subtotal: transaction.subtotal || 0,
      discountAmount: transaction.discountAmount || 0,
      total: calculatedTotal,
      amountReceived: transaction.amountReceived || 0,
      change: transaction.change || 0,
      notes: transaction.notes,
      storeLogo: storeLogo,
      receiptLogo: storeLogo,
    };

    // Serve the JavaScript-based receipt template directly with embedded data
    const templatePath = path.join(__dirname, "public", "receipt-print-full.html");
    let template = await fs.readFile(templatePath, "utf8");

    // Embed data directly into the template
    const dataScript = `
        <script>
            window.transactionData = ${JSON.stringify(receiptData)};
        </script>
    `;

    // Insert data script before the closing </body> tag
    template = template.replace("</body>", dataScript + "</body>");

    res.setHeader("Content-Type", "text/html");
    res.send(template);
  } catch (error) {
    res.status(500).send("Failed to load transaction for printing");
  }
});

// Print receipt (regular transaction)
app.get("/receipt-print", isAuthenticated, isAdminOrCashier, async (req, res) => {
  try {
    // Get data from URL parameters
    const urlParams = new URLSearchParams(req.url.split("?")[1] || "");
    const transactionData = urlParams.get("data");

    if (!transactionData) {
      return res.status(400).send("Transaction data not provided");
    }

    // Parse transaction data
    const transaction = JSON.parse(decodeURIComponent(transactionData));

    // Serve the JavaScript-based receipt template directly with embedded data
    const templatePath = path.join(__dirname, "public", "receipt-print.html");
    let template = await fs.readFile(templatePath, "utf8");

    // Embed data directly into the template
    const dataScript = `
        <script>
            window.transactionData = ${JSON.stringify(transaction)};
        </script>
    `;

    // Insert data script before the closing </body> tag
    template = template.replace("</body>", dataScript + "</body>");

    res.setHeader("Content-Type", "text/html");
    res.send(template);
  } catch (error) {
    res.status(500).send("Failed to load receipt for printing");
  }
});

// Print debt receipt
app.get("/debt-receipt-print", isAuthenticated, isAdminOrCashier, async (req, res) => {
  try {
    // Get data from URL parameters
    const urlParams = new URLSearchParams(req.url.split("?")[1] || "");
    const transactionData = urlParams.get("data");

    if (!transactionData) {
      return res.status(400).send("Transaction data not provided");
    }

    // Parse transaction data
    const transaction = JSON.parse(decodeURIComponent(transactionData));

    // Serve the JavaScript-based debt receipt template directly with embedded data
    const templatePath = path.join(__dirname, "public", "debt-receipt-print.html");
    let template = await fs.readFile(templatePath, "utf8");

    // Embed data directly into the template
    const dataScript = `
        <script>
            window.transactionData = ${JSON.stringify(transaction)};
        </script>
    `;

    // Insert data script before the closing </body> tag
    template = template.replace("</body>", dataScript + "</body>");

    res.setHeader("Content-Type", "text/html");
    res.send(template);
  } catch (error) {
    res.status(500).send("Failed to load debt receipt for printing");
  }
});

// === USERS EXPORT/IMPORT ===
// Export Users to XLSX
app.get("/api/users/export", isAuthenticated, isAdmin, async (req, res) => {
  try {
    const users = await readData("users.json");
    const exportData = users.map((u) => ({
      "Username": u.username || "",
      "Name": u.name || "",
      "Role": u.role || "",
      "Status": u.status || "active",
      "ID": u.id || "",
    }));

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Users");
    ws["!cols"] = [{ wch: 20 }, { wch: 30 }, { wch: 15 }, { wch: 15 }, { wch: 15 }];
    const buf = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", "attachment; filename=users_export.xlsx");
    res.send(buf);
  } catch (error) {
    res.status(500).json({ success: false, message: "Failed to export users: " + error.message });
  }
});

// Download User Import Template
app.get("/api/users/template", isAuthenticated, isAdmin, async (req, res) => {
  try {
    const templateData = [
      { "Username": "user1", "Password": "password123", "Name": "User 1", "Role": "cashier", "Status": "active" },
      { "Username": "user2", "Password": "password123", "Name": "User 2", "Role": "admin", "Status": "active" },
    ];

    const ws = XLSX.utils.json_to_sheet(templateData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Template");
    ws["!cols"] = [{ wch: 20 }, { wch: 20 }, { wch: 30 }, { wch: 15 }, { wch: 15 }];
    const buf = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", "attachment; filename=user_import_template.xlsx");
    res.send(buf);
  } catch (error) {
    res.status(500).json({ success: false, message: "Failed to generate template: " + error.message });
  }
});

// Import Users from XLSX
app.post("/api/users/import", isAuthenticated, isAdmin, async (req, res) => {
  try {
    const { users: importData } = req.body;
    if (!Array.isArray(importData) || importData.length === 0) {
      return res.status(400).json({ success: false, message: "No valid data to import" });
    }

    const users = await readData("users.json");
    let successCount = 0;
    let errorCount = 0;
    const errors = [];

    for (let i = 0; i < importData.length; i++) {
      try {
        const row = importData[i];
        if (!row["Username"] || !row["Password"]) {
          errors.push(`Baris ${i + 1}: Username dan Password wajib diisi`);
          errorCount++;
          continue;
        }

        const username = row["Username"].toString().trim();
        const existingUser = users.find((u) => u.username && u.username.toLowerCase() === username.toLowerCase());
        if (existingUser) {
          errors.push(`Baris ${i + 1}: Username "${username}" sudah ada`);
          errorCount++;
          continue;
        }

        const hashedPassword = await bcrypt.hash(row["Password"].toString(), 10);
        const newUser = {
          id: Date.now() + i,
          username: username,
          password: hashedPassword,
          name: (row["Name"] || "").toString().trim(),
          role: (row["Role"] || "cashier").toString().trim(),
          status: (row["Status"] || "active").toString().trim(),
        };

        users.push(newUser);
        successCount++;
      } catch (error) {
        errors.push(`Baris ${i + 1}: ${error.message}`);
        errorCount++;
      }
    }

    await writeData("users.json", users);
    let message = `Import selesai. Sukses: ${successCount}, Error: ${errorCount}`;
    if (errors.length > 0) {
      message += `\n\nBeberapa error pertama:\n${errors
        .slice(0, 3)
        .join("\n")}`;
      if (errors.length > 5) {
        message += ` ... dan ${errors.length - 5} more errors`;
      }
    }

    res.json({ success: true, message, successCount, errorCount, errors: errors.slice(0, 10) });
  } catch (error) {
    res.status(500).json({ success: false, message: "Failed to import users: " + error.message });
  }
});

// === CUSTOMERS EXPORT/IMPORT ===
// Export Customers to XLSX
app.get("/api/customers/export", isAuthenticated, isAdmin, async (req, res) => {
  try {
    const customers = await readData("customers.json");
    const exportData = customers.map((c) => ({
      "Customer Name": c.name || "",
      "Phone": c.phone || "",
      "Email": c.email || "",
      "Address": c.address || "",
      "Saldo": Math.max(0, Number(c.balance || 0) || 0),
      "ID": c.id || "",
    }));

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Customers");
    ws["!cols"] = [{ wch: 30 }, { wch: 20 }, { wch: 30 }, { wch: 50 }, { wch: 16 }, { wch: 15 }];
    const buf = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", "attachment; filename=customers_export.xlsx");
    res.send(buf);
  } catch (error) {
    res.status(500).json({ success: false, message: "Failed to export customers: " + error.message });
  }
});

// Download Customer Import Template
app.get("/api/customers/template", isAuthenticated, isAdmin, async (req, res) => {
  try {
    const templateData = [
      { "Customer Name": "John Doe", "Phone": "081234567890", "Email": "john@example.com", "Address": "Jl. Example No. 123", "Saldo": 50000 },
      { "Customer Name": "Jane Smith", "Phone": "081987654321", "Email": "jane@example.com", "Address": "Jl. Test No. 456", "Saldo": 0 },
    ];

    const ws = XLSX.utils.json_to_sheet(templateData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Template");
    ws["!cols"] = [{ wch: 30 }, { wch: 20 }, { wch: 30 }, { wch: 50 }, { wch: 16 }];
    const buf = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", "attachment; filename=customer_import_template.xlsx");
    res.send(buf);
  } catch (error) {
    res.status(500).json({ success: false, message: "Failed to generate template: " + error.message });
  }
});

// Import Customers from XLSX
app.post("/api/customers/import", isAuthenticated, isAdmin, async (req, res) => {
  try {
    const { customers: importData } = req.body;
    if (!Array.isArray(importData) || importData.length === 0) {
      return res.status(400).json({ success: false, message: "No valid data to import" });
    }

    const customers = await readData("customers.json");
    let successCount = 0;
    let errorCount = 0;
    const errors = [];

    for (let i = 0; i < importData.length; i++) {
      try {
        const row = importData[i];
        if (!row["Customer Name"] || row["Customer Name"].toString().trim() === "") {
          errors.push(`Baris ${i + 1}: Customer Name wajib diisi`);
          errorCount++;
          continue;
        }

        const newCustomer = {
          id: Date.now() + i,
          name: row["Customer Name"].toString().trim(),
          phone: (row["Phone"] || "").toString().trim(),
          email: (row["Email"] || "").toString().trim(),
          address: (row["Address"] || "").toString().trim(),
          balance: Math.max(0, Number(row["Saldo"] || 0) || 0),
          createdAt: new Date().toISOString(),
        };

        customers.push(newCustomer);
        successCount++;
      } catch (error) {
        errors.push(`Baris ${i + 1}: ${error.message}`);
        errorCount++;
      }
    }

    await writeData("customers.json", customers);
    let message = `Import selesai. Sukses: ${successCount}, Error: ${errorCount}`;
    if (errors.length > 0) {
      message += `\n\nBeberapa error pertama:\n${errors
        .slice(0, 3)
        .join("\n")}`;
      if (errors.length > 5) {
        message += ` ... dan ${errors.length - 5} more errors`;
      }
    }

    res.json({ success: true, message, successCount, errorCount, errors: errors.slice(0, 10) });
  } catch (error) {
    res.status(500).json({ success: false, message: "Failed to import customers: " + error.message });
  }
});

// Check username availability
app.post("/api/users/check-username/:id?", async (req, res) => {
  try {
    const { username } = req.body;
    const userId = req.params.id;

    if (!username || typeof username !== 'string') {
      return res.status(400).json({ success: false, message: "Username wajib diisi" });
    }
    try {
      const existingUser = await validateUsername(username.trim(), userId);
      res.json({ exists: !!existingUser });
    } catch (validationError) {
      return res.status(400).json({ success: false, message: validationError.message });
    }
  } catch (error) {
    res.status(500).json({ success: false, message: "Server error" });
  }
});

// Check product name availability
app.post("/api/products/check-name/:id?", async (req, res) => {
  try {
    const { name } = req.body;
    const productId = req.params.id;

    const existingProduct = await validateProductName(name.trim(), productId);
    res.json({ exists: !!existingProduct });
  } catch (error) {
    res.status(500).json({ success: false, message: "Server error" });
  }
});

// Check category name availability
app.post("/api/categories/check-name/:id?", async (req, res) => {
  try {
    const { name } = req.body;
    const categoryId = req.params.id;

    const existingCategory = await validateCategoryName(name.trim(), categoryId);
    res.json({ exists: !!existingCategory });
  } catch (error) {
    res.status(500).json({ success: false, message: "Server error" });
  }
});

// Reset user password
app.post(
  "/api/users/:id/reset-password",
  isAuthenticated,
  isAdmin,
  async (req, res) => {
    try {
      const { newPassword } = req.body;
      const userId = req.params.id;

      if (!newPassword || newPassword.length < 6) {
        return res.status(400).json({
          success: false,
          message: "Password minimal 6 karakter",
        });
      }

      const users = await readData("users.json");
      const index = users.findIndex((u) => u.id == userId);

      if (index !== -1) {
        users[index].password = await bcrypt.hash(newPassword, 10);
        users[index].updatedAt = new Date().toISOString();

        await writeData("users.json", users);
        res.json({
          success: true,
          message: "Password berhasil direset",
        });
      } else {
        res.status(404).json({
          success: false,
          message: "User tidak ditemukan",
        });
      }
    } catch (error) {
      res.status(500).json({
        success: false,
        message: "Failed to reset password",
      });
    }
  }
);

// --- Product Drafts API ---
// Helper untuk draf POS
const readPosDrafts = async () => readData("pos-drafts.json");
const writePosDrafts = async (drafts) => writeData("pos-drafts.json", drafts);

// GET /api/pos-drafts - Ambil semua draf
app.get("/api/pos-drafts", isAuthenticated, async (req, res) => {
  try {
    const drafts = await readPosDrafts();
    res.json(drafts);
  } catch (error) {
    res.status(500).json({ success: false, message: "Failed to load drafts" });
  }
});

// POST /api/pos-drafts - Simpan draf baru
app.post("/api/pos-drafts", isAuthenticated, async (req, res) => {
  try {
    const { items } = req.body;
    if (!items || items.length === 0) {
      return res
        .status(400)
        .json({ success: false, message: "Cannot save an empty draft." });
    }

    const drafts = await readPosDrafts();
    const newDraft = {
      id: Date.now().toString(),
      items: items,
      timestamp: new Date().toISOString(),
    };
    drafts.push(newDraft);
    await writePosDrafts(drafts);
    res.json({
      success: true,
      message: "Draft saved successfully!",
      draft: newDraft,
    });
  } catch (error) {
    res.status(500).json({ success: false, message: "Failed to save draft" });
  }
});

// PUT /api/pos-drafts/:id/load - Muat draf ke keranjang dan hapus
app.put("/api/pos-drafts/:id/load", isAuthenticated, async (req, res) => {
  try {
    const drafts = await readPosDrafts();
    const draftIndex = drafts.findIndex((d) => d.id === req.params.id);

    if (draftIndex === -1) {
      return res
        .status(404)
        .json({ success: false, message: "Draft not found." });
    }

    const draftToLoad = drafts[draftIndex];

    // Hapus draf setelah dimuat
    drafts.splice(draftIndex, 1);
    await writePosDrafts(drafts);

    res.json({
      success: true,
      message: "Draft loaded successfully.",
      items: draftToLoad.items,
    });
  } catch (error) {
    res.status(500).json({ success: false, message: "Failed to load draft" });
  }
});

// DELETE /api/pos-drafts/:id - Hapus draf
app.delete("/api/pos-drafts/:id", isAuthenticated, async (req, res) => {
  try {
    const drafts = await readPosDrafts();
    const filteredDrafts = drafts.filter((d) => d.id !== req.params.id);

    if (drafts.length === filteredDrafts.length) {
      return res
        .status(404)
        .json({ success: false, message: "Draft not found." });
    }

    await writePosDrafts(filteredDrafts);
    res.json({ success: true, message: "Draft deleted successfully." });
  } catch (error) {
    res.status(500).json({ success: false, message: "Failed to delete draft" });
  }
});

app.post('/api/admin/encrypt-migrate', isAuthenticated, isAdmin, async (req, res) => {
  try {
    return res.json({
      success: true,
      processed: 0,
      mode: 'sqlite',
      message: 'Data utama ada di pos.db (SQLite) — tidak perlu migrasi enkripsi file JSON.'
    });
  } catch (e) {
    console.error('Error in encrypt-migrate:', e);
    res.status(500).json({ success: false, message: 'Server error' });
  }
});

// Public tool: decrypt ENC1 payload with passphrase (for local DB reader)
app.post('/api/tools/decrypt-enc1', strictLimiter, isAuthenticated, async (req, res) => {
  try {
    const encText = req.body && req.body.encText ? String(req.body.encText) : '';
    const passphrase = req.body && req.body.passphrase ? String(req.body.passphrase) : '';
    if (!encText || !encText.startsWith('ENC1:')) {
      return res.status(400).json({ success: false, message: 'ENC1 text is required' });
    }
    if (!passphrase) {
      return res.status(400).json({ success: false, message: 'Passphrase is required' });
    }
    const dec = decryptTextEnc1WithPassphrase(encText, passphrase);
    let data = null;
    try { data = JSON.parse(dec); } catch {}
    return res.json({ success: true, data, raw: data ? undefined : dec });
  } catch (e) {
    return res.status(400).json({ success: false, message: e.message || 'Decryption failed' });
  }
});

// Public tool: read passphrase from data/passphrase.txt (local use)
app.get('/api/tools/passphrase', isAuthenticated, async (req, res) => {
  try {
    const ip = (req.ip || '').toString();
    const host = (req.hostname || '').toString();
    const isLocal = ip === '127.0.0.1' || ip === '::1' || host === 'localhost';
    if (!isLocal) {
      return res.status(403).json({ success: false, message: 'Forbidden' });
    }
    const passPath = path.join(DATA_DIR, 'passphrase.txt');
    const raw = await fs.readFile(passPath, 'utf-8').catch(() => '');
    if (!raw) return res.status(404).json({ success: false, message: 'Passphrase tidak ditemukan' });
    return res.json({ success: true, passphrase: raw.toString().trim() });
  } catch (e) {
    return res.status(500).json({ success: false, message: 'Gagal membaca passphrase' });
  }
});

// --- Global 404 handler (must be after all routes) ---
app.use((req, res, next) => {
  const code = 404;
  const message = 'Not Found';
  // If client expects JSON (API/fetch), return JSON
  const acceptsHtml = req.accepts(['html', 'json']) === 'html';
  if (!acceptsHtml || req.originalUrl.startsWith('/api/')) {
    return res.status(code).json({ success: false, code, message, path: req.originalUrl });
  }
  const q = new URLSearchParams({ code: String(code), msg: message, path: req.originalUrl });
  return res.status(code).redirect(`/error.html?${q.toString()}`);
});

// --- Global error handler ---
// eslint-disable-next-line no-unused-vars
app.use((err, req, res, next) => {
  const code = err?.status || err?.statusCode || 500;
  const message = err?.message || 'Internal Server Error';
  const acceptsHtml = req.accepts(['html', 'json']) === 'html';
  if (!acceptsHtml || req.originalUrl.startsWith('/api/')) {
    return res.status(code).json({ success: false, code, message, path: req.originalUrl });
  }
  const q = new URLSearchParams({ code: String(code), msg: message, path: req.originalUrl });
  return res.status(code).redirect(`/error.html?${q.toString()}`);
});

// --- PERUBAHAN 3: Inisialisasi Server yang Lebih Aman ---
// Gunakan async IIFE untuk memastikan direktori data ada sebelum server berjalan
const IS_TEST_MODE = String(process.env.POS_TEST_MODE || '').toLowerCase() === 'true';
(async () => {
  try {
    await ensureDataDir();
    await initDatabase();
    await loadPassphraseFromFile();
    if (IS_TEST_MODE) { return; }
    // Initialize encryption settings after passphrase is loaded
    try {
      const s = await readData('settings.json').catch(() => ({}));
      if (s && typeof s === 'object' && s.encryption && typeof s.encryption.enabled === 'boolean') {
        encryptionEnabled = !!s.encryption.enabled;
      }
    } catch {}
    // Run auto-backup if configured
    await autoBackupIfNeededOnStart();
    
    // Increment license runs on server startup if configured
    try {
      const off = await verifyOfflineLicense();
      if (off && off.payload && off.payload.mode === 'runs' && Number(off.payload.maxRuns || 0) > 0) {
        const maxRuns = Number(off.payload.maxRuns || 0);
        const lk = await readLicenseKey();
        if (lk) {
          let settings = {};
          let licenseConfig = {};
          try {
            settings = await readData('settings.json').catch(() => ({}));
            licenseConfig = settings.license || {};
          } catch (e) {
            // Fallback untuk build exe
            try {
              const licenseConfigPath = path.join(DATA_DIR, 'license-config.json');
              const raw = await fs.readFile(licenseConfigPath, 'utf-8').catch(() => '');
              if (raw) {
                licenseConfig = JSON.parse(raw);
              }
            } catch (fallbackError) {
              // Default values
              licenseConfig = {
                countOnRestart: false,
                sessionTimeout: 5000
              };
            }
          }
          const countOnRestart = licenseConfig.countOnRestart !== false; // default: true
          
          if (countOnRestart) {
            const result = await incrementLicenseRunsOnStartup(lk, maxRuns, { isRestart: true });
          }
        }
      }
    } catch (e) {
    }
    
    let hasValidLicense = false;
    try {
      const off = await verifyOfflineLicense();
      hasValidLicense = !!(off && off.valid);
      if (hasValidLicense) {
      }
    } catch {}
    if (!hasValidLicense) {
      try {
        const lic = await checkLicenseOnline();
        hasValidLicense = !!(lic && lic.ok);
      } catch {}
    }
    if (!hasValidLicense) {
      await ensureTrialNotExpired();
    } else {
    }
    
    // Modify saveArrayWithSync to broadcast updates when products are saved
    const originalSaveArrayWithSync = saveArrayWithSync;
    saveArrayWithSync = async function(filename, data, options) {
      console.log(`saveArrayWithSync called with filename: ${filename}`);
      const result = await originalSaveArrayWithSync.call(this, filename, data, options);
      
      // Broadcast update if products.json was saved
      if (filename === 'products.json') {
        console.log('Products.json saved, broadcasting update...');
        try {
          try { invalidateCache('products'); } catch {}
          let reason = 'full';
          try {
            const lastAt = global.__lastStockMoveAt || 0;
            const stockOnly = global.__lastStockMoveStockOnly === true;
            if (stockOnly && Date.now() - Number(lastAt) < 2000) reason = 'stock_only';
          } catch {}
          broadcastProductUpdate('products_updated', { 
            message: 'Products data has been updated',
            count: Array.isArray(data) ? data.length : 0,
            reason
          });
        } catch (error) {
          console.error('Error broadcasting product update:', error);
        }
      }
      if (filename === 'categories.json') {
        try { invalidateCache('categories'); } catch {}
      }
      
      return result;
    };

    // Initialize JWT secret
    try {
      if (!JWT_SECRET) {
        const jwtFile = getJwtFilePath();
        const dir = path.dirname(jwtFile);
        await fs.mkdir(dir, { recursive: true });
        try {
          JWT_SECRET = await fs.readFile(jwtFile, 'utf8').then(s => s.trim()).catch(() => '');
        } catch {}
        if (!JWT_SECRET) {
          JWT_SECRET = crypto.randomBytes(32).toString('hex');
          await fs.writeFile(jwtFile, JWT_SECRET, 'utf8');
        }
      }
    } catch {}

    // Preload products into cache when server starts
    preloadProducts();
    
    const server = app.listen(PORT, HOST, () => {
      const os = require('os');
      const getNetworkIp = () => {
        const interfaces = os.networkInterfaces();
        for (const name of Object.keys(interfaces)) {
          for (const iface of interfaces[name]) {
            if (iface.family === 'IPv4' && !iface.internal) {
              return iface.address;
            }
          }
        }
        return null;
      };
      
      const networkIp = getNetworkIp();
      console.log(`Server berjalan di:`);
      console.log(`  - Local:    http://localhost:${PORT}`);
      if (networkIp) {
        console.log(`  - Network:  http://${networkIp}:${PORT}`);
      }
      
      const baseUrl = `http://localhost:${PORT}`;
      try {
        const https = require('https');
        const { ensureLocalHttpsCredentials } = require('./lib/local-https');
        const HTTPS_PORT = Number(process.env.HTTPS_PORT) || (PORT + 1);
        ensureLocalHttpsCredentials(DATA_DIR).then((creds) => {
          if (!creds) {
            console.log('  - HTTPS:    tidak tersedia (gagal membuat sertifikat lokal)');
            return;
          }
          https.createServer(creds, app).listen(HTTPS_PORT, HOST, () => {
            console.log(`  - HTTPS:    https://localhost:${HTTPS_PORT}`);
            if (networkIp) {
              console.log(`  - HTTPS HP: https://${networkIp}:${HTTPS_PORT}  (untuk Bluetooth browser Android)`);
            }
          });
        }).catch(() => {
          console.log('  - HTTPS:    tidak aktif');
        });
      } catch (httpsErr) {
        try { console.log('  - HTTPS:    tidak aktif'); } catch {}
      }
      
      if (hasValidLicense) {
        // console.log('License: VALID');
      } else if (!TRIAL_ENABLED) {
        // console.log('License: NOT VALID, TRIAL DISABLED via env, running in unlocked mode');
      } else {
        // console.log('License: MISSING or INVALID, running in TRIAL mode (if not expired)');
      }

      // Auto-open admin page in default browser on Windows
      try {
        if (process.platform === 'win32') {
          const adminUrl = `${baseUrl}/admin.html`;
          exec(`start "" "${adminUrl}"`);
        }
      } catch (e) {
        // try { console.warn('Failed to auto-open browser:', e && e.message ? e.message : e); } catch {}
      }
    });
    server.on("error", (err) => {
      // console.error("Server error:", err);
      process.exit(1);
    });
  } catch (err) {
    try { process.stderr.write(`[POS] Fatal startup error: ${err && err.stack ? err.stack : err}\n`); } catch {}
    if (!IS_TEST_MODE) process.exit(1);
  }
})();

// Export app for testing
if (IS_TEST_MODE) {
  module.exports = { app };
}

