/* One-time move from the old version, which kept its data inside Microsoft Edge (or Chrome).
   Reads the browser's saved-data files (a LevelDB folder) directly and returns the planner's
   keys for the file:// page. Read-only: the browser's own copy is left exactly as it was. */
'use strict';
const fs = require('fs');
const os = require('os');
const path = require('path');
const snappy = require('./vendor/snappyjs');

function varint(b, i) {
  let r = 0, s = 0, c;
  do { c = b[i++]; r += (c & 0x7f) * Math.pow(2, s); s += 7; } while (c & 0x80);
  return [r, i];
}
function* blockEntries(data) {
  const nres = data.readUInt32LE(data.length - 4), end = data.length - 4 - 4 * nres;
  let i = 0, last = Buffer.alloc(0);
  while (i < end) {
    let sh, ns, vl;
    [sh, i] = varint(data, i); [ns, i] = varint(data, i); [vl, i] = varint(data, i);
    const k = Buffer.concat([last.subarray(0, sh), data.subarray(i, i + ns)]); i += ns;
    const v = data.subarray(i, i + vl); i += vl; last = k;
    yield [k, v];
  }
}
function readBlock(f, off, size) {
  const raw = f.subarray(off, off + size);
  return f[off + size] === 1 ? Buffer.from(snappy.uncompress(raw)) : raw;
}
function* sst(file) {
  const f = fs.readFileSync(file), foot = f.subarray(f.length - 48);
  let i = 0, mo, ms, io, isz;
  [mo, i] = varint(foot, 0); [ms, i] = varint(foot, i); [io, i] = varint(foot, i); [isz, i] = varint(foot, i);
  for (const [, h] of blockEntries(readBlock(f, io, isz))) {
    let j, bo, bs; [bo, j] = varint(h, 0); [bs, j] = varint(h, j);
    for (const [ik, v] of blockEntries(readBlock(f, bo, bs))) {
      const tag = ik.readBigUInt64LE(ik.length - 8);
      yield [ik.subarray(0, ik.length - 8).toString('latin1'), tag >> 8n, Number(tag & 0xffn), v];
    }
  }
}
function* logFile(file) {
  const f = fs.readFileSync(file); let i = 0, buf = null; const recs = [];
  while (i + 7 <= f.length) {
    const left = 32768 - (i % 32768);
    if (left < 7) { i += left; continue; }
    const len = f.readUInt16LE(i + 4), t = f[i + 6], data = f.subarray(i + 7, i + 7 + len); i += 7 + len;
    if (t === 0 && len === 0) continue;
    if (t === 1) recs.push(data);
    else if (t === 2) buf = [data];
    else if (t === 3 && buf) buf.push(data);
    else if (t === 4 && buf) { buf.push(data); recs.push(Buffer.concat(buf)); buf = null; }
  }
  for (const r of recs) {
    if (r.length < 12) continue;
    const seq = r.readBigUInt64LE(0), cnt = r.readUInt32LE(8); let j = 12;
    for (let n = 0; n < cnt; n++) {
      const tg = r[j++]; let kl, vl; [kl, j] = varint(r, j);
      const k = r.subarray(j, j + kl).toString('latin1'); j += kl;
      if (tg === 1) { [vl, j] = varint(r, j); const v = r.subarray(j, j + vl); j += vl; yield [k, seq + BigInt(n), 1, v]; }
      else yield [k, seq + BigInt(n), 0, null];
    }
  }
}
function decode(b) {
  if (!b || !b.length) return '';
  return b[0] === 0 ? b.subarray(1).toString('utf16le') : b.subarray(1).toString('latin1');
}
function readFolder(dir) {
  /* copy first: the browser may have the folder open */
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'hufc-edge-'));
  const best = new Map();
  try {
    const names = fs.readdirSync(dir).filter(n => /\.(ldb|log)$/i.test(n));
    names.forEach(n => { try { fs.copyFileSync(path.join(dir, n), path.join(tmp, n)); } catch (e) {} });
    fs.readdirSync(tmp).forEach(n => {
      try {
        const it = /\.log$/i.test(n) ? logFile(path.join(tmp, n)) : sst(path.join(tmp, n));
        for (const [k, seq, t, v] of it) { const b = best.get(k); if (!b || b[0] < seq) best.set(k, [seq, t, v]); }
      } catch (e) { /* one damaged file shouldn't stop the rest */ }
    });
  } finally { try { fs.rmSync(tmp, { recursive: true, force: true }); } catch (e) {} }
  const out = {};
  best.forEach(([, t, v], k) => {
    if (t !== 1 || k[0] !== '_') return;
    const z = k.indexOf('\x00'); if (z < 0) return;
    if (k.slice(1, z) !== 'file://') return;
    const keyBuf = Buffer.from(k.slice(z + 1), 'latin1'), key = decode(keyBuf);
    if (key.indexOf('hufc-') === 0) out[key] = decode(v);
  });
  return out;
}
/* Edge first (that's what the old version opened in), then Chrome */
function readOldVersion() {
  const la = process.env.LOCALAPPDATA || '';
  const dirs = [
    path.join(la, 'Microsoft', 'Edge', 'User Data', 'Default', 'Local Storage', 'leveldb'),
    path.join(la, 'Google', 'Chrome', 'User Data', 'Default', 'Local Storage', 'leveldb')
  ];
  const merged = {};
  dirs.forEach(d => {
    if (!fs.existsSync(d)) return;
    try {
      const kv = readFolder(d);
      Object.keys(kv).forEach(k => {
        if (!(k in merged)) { merged[k] = kv[k]; return; }
        const j = s => { try { return JSON.parse(s || 'null'); } catch (e) { return null; } };
        const a = j(merged[k]), b = j(kv[k]);  /* both browsers had it: keep the newer */
        if (a && b && (b.savedAt || 0) > (a.savedAt || 0)) merged[k] = kv[k];
      });
    } catch (e) {}
  });
  return merged;
}
module.exports = { readOldVersion, readFolder };
