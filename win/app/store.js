/* The planner's saved data, kept as ordinary files in named folders (Sessions, Practices, Lineups, ...)
   instead of inside a browser. The page still talks to it like localStorage (see preload.js); this side
   turns each key into a readable .json file and back again. */
'use strict';
const fs = require('fs');
const path = require('path');

const F = {
  sessions: 'Sessions', pregames: 'Pre-games', lineups: 'Lineups', practices: 'Practices',
  squads: 'Squads', stats: 'Stats', pictures: 'Pictures', videos: 'Videos', powerpoint: 'PowerPoint', pdfs: 'PDFs', settings: 'Settings', backups: 'Backups'
};
const P_SESSION = 'hufc-session-plan:', P_LINEUP = 'hufc-lineup:', P_SQUAD = 'hufc-squad:', LIB_KEY = 'hufc-practice-library';
const SETTINGS_KEY = '\u0000settings';
const MEDIA_BACKUP_MAX = 300 * 1024 * 1024;   /* clips and pictures are backed up; full match videos aren't */
/* Firebase keeps its own short-lived notes in storage (which tab is open, sync progress): they're never worth saving to disk */
function volatileKey(k) { return /^(firestore_|firebase:|firebase-heartbeat|__sak)/.test(String(k)); }
/* each kind of file has its own type, so a double-click opens it in the planner (and other .json files are left alone) */
const EXT = { sessions: '.hufcsession', pregames: '.hufcpregame', lineups: '.hufclineup', practices: '.hufcpractice', squads: '.hufcsquad' };

const clean = v => (v == null ? '' : String(v)).replace(/\s+/g, ' ').trim();
const parse = s => { try { return JSON.parse(s); } catch (e) { return undefined; } };
function safeName(s, max) {
  s = clean(s).replace(/[\\/:*?"<>|\x00-\x1f]+/g, '-').replace(/[. ]+$/, '');
  if (s.length > (max || 70)) s = s.slice(0, max || 70).replace(/\s+\S*$/, '');
  return s || 'Untitled';
}
const idOf = file => { const m = /\[([^\]]+)\]\.(json|hufc[a-z]+)$/i.exec(file); return m ? m[1] : null; };

/* ---- what goes where, and what each file is called ---- */
function isPre(s) { return !!(s && s.doc === 'pregame'); }
function sessionName(s) {
  const m = s.meta || {};
  if (clean(m.title)) return clean(m.title);
  if (isPre(s)) return clean(m.opponent) ? 'vs ' + clean(m.opponent) : '';
  return clean(s.focus && s.focus.theme);
}
function sessionHasContent(s) {
  const m = s.meta || {};
  if (sessionName(s) || clean(m.date) || clean(m.venue)) return true;
  const ds = Array.isArray(s.drills) ? s.drills : [];
  if (isPre(s)) return ds.some(d => (d.clips || []).length || (d.pics || []).length || (d.notes || []).some(f => clean(f.value)));
  return ds.some(d => (d.items && d.items.length) || clean(d.focus) || (d.notes || []).some(f => clean(f.value)));
}
function lineupName(s) {
  const m = s.meta || {};
  return clean(s.title) || (clean(m.opponent) ? 'vs ' + clean(m.opponent) : '') || clean(m.round);
}
function lineupHasContent(s) {
  const m = s.meta || {};
  return !!(clean(s.title) || clean(m.opponent) || clean(m.round) || clean(m.date) || clean(m.competition) ||
    (s.squad || []).some(p => clean(p.name) || clean(p.num)));
}
const datePart = d => { d = clean(d).replace(/[\\/]+/g, '-'); return d ? d + ' ' : ''; };

class Store {
  constructor(root, log) {
    this.root = root; this.log = log || (() => {});
    this.mem = new Map();          /* key -> value (string), exactly what the page stored */
    this.files = new Map();        /* key -> file path (sessions, lineups, squads) */
    this.practiceFiles = new Map();/* practice id -> file path */
    this.settings = {};            /* everything else, in Settings/app-settings.json */
    this.timers = new Map();
    this.oldFiles = [];             /* files saved as .json by the first version: renamed to their own type on load */
  }
  dir(name) { return path.join(this.root, F[name]); }
  ensureFolders() {
    fs.mkdirSync(this.root, { recursive: true });
    Object.keys(F).forEach(k => fs.mkdirSync(this.dir(k), { recursive: true }));
    const readme = path.join(this.root, 'About these folders.txt');
    const text = [
      'HUFC Season Planner keeps everything here as it saves:',
      '',
      '  Sessions    training session plans',
      '  Pre-games   pre-game plans',
      '  Lineups     matchday lineups',
      '  Practices   every saved practice (one file each)',
      '  Squads      your saved squad for each team',
      '  Stats       game time & player stats for each team (opens in Excel)',
      '  Pictures    pictures added to pre-game boards',
      '  Videos      video clips added to pre-games with + Upload video',
      '  PowerPoint  slides made with Export > Open in PowerPoint',
      '  PDFs        PDFs made with Export > PDF',
      '  Settings    the app\'s own settings',
      '  Backups     a copy of everything, made each day you use the app (the last 14 days are kept);\r\n              pictures and video clips are copied once into Backups\\Pictures and Backups\\Videos\r\n              (whole match videos aren\'t, as they\'re so big)',
      '',
      'Double-click any of these files to open it in the planner. Please don\'t edit them by hand.',
      ''].join('\r\n');
    try { if (!fs.existsSync(readme) || fs.readFileSync(readme, 'utf8') !== text) fs.writeFileSync(readme, text); } catch (e) {}
  }

  /* ---------- reading everything in when the app opens ---------- */
  load() {
    this.ensureFolders();
    const readJson = p => parse(fs.readFileSync(p, 'utf8'));
    const each = (folder, fn) => {
      let names = [];
      try { names = fs.readdirSync(this.dir(folder)); } catch (e) { return; }
      names.filter(n => /\.(json|hufc[a-z]+)$/i.test(n)).forEach(n => {
        const p = path.join(this.dir(folder), n), id = idOf(n);
        if (!id) return;
        if (/\.json$/i.test(n)) this.oldFiles.push(folder);
        try { const v = readJson(p); if (v !== undefined) fn(id, v, p); } catch (e) { this.log('read failed', p, e.message); }
      });
    };
    const keep = (key, obj, p) => {
      const old = this.files.get(key);
      if (old && old !== p) {        /* two files for one plan (e.g. renamed while the other was open): keep the newer */
        const prev = parse(this.mem.get(key)) || {};
        if ((prev.savedAt || 0) >= (obj.savedAt || 0)) return;
      }
      this.mem.set(key, JSON.stringify(obj)); this.files.set(key, p);
    };
    each('sessions', (id, v, p) => keep(P_SESSION + id, v, p));
    each('pregames', (id, v, p) => keep(P_SESSION + id, v, p));
    each('lineups', (id, v, p) => keep(P_LINEUP + id, v, p));
    each('squads', (id, v, p) => keep(P_SQUAD + id, v, p));
    /* the practice library: one file per practice, plus the library's own bits in Settings */
    const items = {};
    each('practices', (id, v, p) => { items[id] = v; this.practiceFiles.set(id, p); });
    let lib = null;
    try { lib = readJson(path.join(this.dir('settings'), 'practice-library.json')); } catch (e) {}
    if (lib || Object.keys(items).length) {
      lib = Object.assign({ v: 1, removed: {} }, lib || {}); lib.items = items;
      this.mem.set(LIB_KEY, JSON.stringify(lib));
    }
    try { this.settings = readJson(path.join(this.dir('settings'), 'app-settings.json')) || {}; } catch (e) { this.settings = {}; }
    let stale = false;
    Object.keys(this.settings).forEach(k => { if (volatileKey(k)) { stale = true; return; } this.mem.set(k, String(this.settings[k])); });
    if (stale) this.schedule(SETTINGS_KEY);   /* (tidy away the sign-in system's own bookkeeping saved there by older versions) */
    if (this.oldFiles.length) {
      this.files.forEach((p, key) => { if (/\.json$/i.test(p)) this.write(key); });
      if (this.oldFiles.indexOf('practices') !== -1) this.write(LIB_KEY);
      this.oldFiles = [];
    }
    return this;
  }
  isEmpty() {
    return ![...this.mem.keys()].some(k => k.indexOf(P_SESSION) === 0 || k.indexOf(P_LINEUP) === 0 || k === LIB_KEY || k.indexOf(P_SQUAD) === 0);
  }
  all() { const o = {}; this.mem.forEach((v, k) => { o[k] = v; }); return o; }

  /* ---------- the page saved or removed something ---------- */
  /* the sign-in system's own bookkeeping (it changes every few seconds) stays in memory: writing it to disk kept the program busy */
  set(key, value) { this.mem.set(key, String(value)); if (!volatileKey(key)) this.schedule(key); }
  remove(key) { this.mem.delete(key); if (!volatileKey(key)) this.schedule(key); }
  schedule(key) {
    /* plans each get their own file; everything else shares the one settings file, written once for a burst of changes */
    const doc = key.indexOf(P_SESSION) === 0 || key.indexOf(P_LINEUP) === 0 || key.indexOf(P_SQUAD) === 0 || key === LIB_KEY;
    const k = doc ? key : SETTINGS_KEY;
    clearTimeout(this.timers.get(k));
    this.timers.set(k, setTimeout(() => { this.timers.delete(k); this.write(k); }, doc ? 250 : 800));
  }
  flush() {
    const was = this.syncWrites; this.syncWrites = true;
    if (this.settingsBusy || this.settingsAgain) { this.settingsAgain = false; this.writeSettings(); }   /* (a background save still going: the newest goes down now) */
    [...this.timers.keys()].forEach(k => { clearTimeout(this.timers.get(k)); this.timers.delete(k); this.write(k); });
    this.syncWrites = was;
  }

  writeFile(p, text) {
    const tmp = p + '.saving';
    fs.writeFileSync(tmp, text, 'utf8');
    fs.renameSync(tmp, p);
  }
  unlink(p) { try { if (p && fs.existsSync(p)) fs.unlinkSync(p); } catch (e) { this.log('delete failed', p, e.message); } }

  write(key) {
    try {
      if (key.indexOf(P_SESSION) === 0 || key.indexOf(P_LINEUP) === 0 || key.indexOf(P_SQUAD) === 0) return this.writeDoc(key);
      if (key === LIB_KEY) return this.writeLibrary();
      return this.writeSettings();
    } catch (e) { this.log('save failed', key, e.message); }
  }
  writeDoc(key) {
    const old = this.files.get(key), raw = this.mem.get(key);
    const obj = raw == null ? undefined : parse(raw);
    if (obj === undefined || obj === null || typeof obj !== 'object') {
      if (raw == null) { this.unlink(old); this.files.delete(key); return; }
      return this.writeSettings();          /* not a plan after all: keep it with the settings */
    }
    const id = key.slice(key.indexOf(':') + 1);
    let folder, name, keepIt = true;
    if (key.indexOf(P_SESSION) === 0) {
      folder = isPre(obj) ? 'pregames' : 'sessions';
      name = datePart(obj.meta && obj.meta.date) + (sessionName(obj) || (isPre(obj) ? 'Untitled pre-game' : 'Untitled session'));
      keepIt = sessionHasContent(obj);
    } else if (key.indexOf(P_LINEUP) === 0) {
      folder = 'lineups';
      name = datePart(obj.meta && obj.meta.date) + (lineupName(obj) || 'Untitled lineup');
      keepIt = lineupHasContent(obj);
    } else {
      folder = 'squads';
      name = (clean(obj.group) || 'Team') + ' squad';
      keepIt = Array.isArray(obj.players) && obj.players.length > 0;
    }
    /* an empty new plan isn't written until it has something in it (the app's home page hides those too) */
    if (!keepIt) { this.unlink(old); this.files.delete(key); return; }
    const p = path.join(this.dir(folder), safeName(name) + ' [' + safeName(id, 40) + ']' + EXT[folder]);
    this.writeFile(p, JSON.stringify(obj, null, 2));
    if (old && old !== p) this.unlink(old);
    this.files.set(key, p);
  }
  writeLibrary() {
    const lib = parse(this.mem.get(LIB_KEY) || 'null');
    const items = (lib && lib.items) || {};
    const seen = new Set();
    Object.keys(items).forEach(id => {
      const e = items[id], pr = (e && e.practice) || {};
      const p = path.join(this.dir('practices'), safeName(pr.name || 'Untitled practice') + ' [' + safeName(id, 40) + ']' + EXT.practices);
      const text = JSON.stringify(e, null, 2), old = this.practiceFiles.get(id);
      let same = false;
      if (old === p) { try { same = fs.readFileSync(p, 'utf8') === text; } catch (err) {} }
      if (!same) this.writeFile(p, text);
      if (old && old !== p) this.unlink(old);
      this.practiceFiles.set(id, p); seen.add(id);
    });
    [...this.practiceFiles.keys()].forEach(id => { if (!seen.has(id)) { this.unlink(this.practiceFiles.get(id)); this.practiceFiles.delete(id); } });
    const rest = Object.assign({}, lib || {}); delete rest.items;
    const lp = path.join(this.dir('settings'), 'practice-library.json');
    if (lib) this.writeFile(lp, JSON.stringify(rest, null, 2)); else this.unlink(lp);
  }
  writeSettings() {
    const s = {};
    this.mem.forEach((v, k) => {
      if (k.indexOf(P_SESSION) === 0 || k.indexOf(P_LINEUP) === 0 || k.indexOf(P_SQUAD) === 0 || k === LIB_KEY) {
        const o = parse(v); if (o && typeof o === 'object') return;
      }
      if (k === 'hufc-probe' || volatileKey(k)) return;
      s[k] = v;
    });
    this.settings = s;
    const p = path.join(this.dir('settings'), 'app-settings.json'), text = JSON.stringify(s, null, 2);
    const gen = this.settingsGen = (this.settingsGen || 0) + 1;
    if (this.syncWrites) return this.writeFile(p, text);   /* (closing: write it now) */
    /* in the background, so the window never waits on the disk; one at a time, and the newest always wins */
    if (this.settingsBusy) { this.settingsAgain = true; return; }
    this.settingsBusy = true;
    const tmp = p + '.saving-bg';
    const done = (e) => {
      this.settingsBusy = false;
      if (e) this.log('save failed', 'settings', e.message);
      if (this.settingsAgain) { this.settingsAgain = false; this.writeSettings(); }
    };
    fs.promises.writeFile(tmp, text, 'utf8').then(async () => {
      for (let i = 0; ; i++) {
        if (gen !== this.settingsGen) { try { await fs.promises.unlink(tmp); } catch (x) {} return; }   /* (a newer save has gone down since: this one's out of date) */
        try { await fs.promises.rename(tmp, p); return; }
        catch (e) { if (i >= 4 || !/EPERM|EBUSY|EACCES/.test(e.code || '')) throw e; await new Promise(r => setTimeout(r, 120 * (i + 1))); }   /* (antivirus or OneDrive holding the file a moment) */
      }
    }).then(() => done(), done);
  }

  /* ---------- Stats: a spreadsheet per team, made by the page from the lineups ---------- */
  writeStats(list) {
    if (!Array.isArray(list)) return;
    const keep = new Set();
    list.forEach(t => {
      if (!t || typeof t.csv !== 'string') return;
      const p = path.join(this.dir('stats'), safeName(t.team || 'Team') + ' game time & stats.csv');
      keep.add(p);
      const text = '﻿' + t.csv + '\r\n';
      try { if (fs.existsSync(p) && fs.readFileSync(p, 'utf8') === text) return; this.writeFile(p, text); }
      catch (e) { this.log('stats save failed (is it open in Excel?)', p, e.message); }
    });
    try {
      fs.readdirSync(this.dir('stats')).filter(n => / game time & stats\.csv$/.test(n)).map(n => path.join(this.dir('stats'), n))
        .forEach(p => { if (!keep.has(p)) this.unlink(p); });
    } catch (e) {}
  }

  /* ---------- a copy of everything once a day; the last 14 kept ---------- */
  dailyBackup() {
    const d = new Date(), pad = n => String(n).padStart(2, '0');
    const today = d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate());
    const dest = path.join(this.dir('backups'), today);
    if (fs.existsSync(dest)) return;
    if (this.isEmpty()) return;
    ['sessions', 'pregames', 'lineups', 'practices', 'squads', 'stats', 'settings'].forEach(k => {
      try { fs.cpSync(this.dir(k), path.join(dest, F[k]), { recursive: true }); } catch (e) { this.log('backup', k, e.message); }
    });
    this.backupVideos();
    try {
      const all = fs.readdirSync(this.dir('backups')).filter(n => /^\d{4}-\d{2}-\d{2}$/.test(n)).sort();
      all.slice(0, Math.max(0, all.length - 14)).forEach(n => fs.rmSync(path.join(this.dir('backups'), n), { recursive: true, force: true }));
    } catch (e) {}
  }

  /* pictures and videos are big: each one is copied into Backups once, not every day */
  backupVideos() { this.backupMedia(); }
  backupMedia() {
    this.tidyBigBackups();
    ['pictures', 'videos'].forEach(k => {
      try {
        const dest = path.join(this.dir('backups'), F[k]); fs.mkdirSync(dest, { recursive: true });
        fs.readdirSync(this.dir(k)).forEach(n => {
          const from = path.join(this.dir(k), n), to = path.join(dest, n);
          if (fs.existsSync(to)) return;
          /* a whole match video (hundreds of MB to several GB) isn't copied again: it would double the space it takes */
          try { if (fs.statSync(from).size > MEDIA_BACKUP_MAX) return; } catch (e) { return; }
          fs.copyFile(from, to, () => {});
        });
      } catch (e) { this.log('media backup', k, e.message); }
    });
  }
  /* older versions copied whole match videos into Backups\Videos: those copies go (the video itself stays in Videos) */
  tidyBigBackups() {
    try {
      const bdir = path.join(this.dir('backups'), F.videos), vdir = this.dir('videos');
      if (!fs.existsSync(bdir)) return;
      fs.readdirSync(bdir).forEach(n => {
        const b = path.join(bdir, n), v = path.join(vdir, n);
        let bs, vs; try { bs = fs.statSync(b); vs = fs.statSync(v); } catch (e) { return; }   /* (only when the same video is still in Videos) */
        if (bs.isFile() && vs.isFile() && bs.size > MEDIA_BACKUP_MAX && bs.size === vs.size) { fs.unlink(b, e => { if (e) this.log('tidy backup', n, e.message); else this.log('removed backup copy of a whole match video', n); }); }
      });
    } catch (e) { this.log('tidy backup', e.message); }
  }
  picturePath(name) {
    name = path.basename(String(name || ''));
    const p = path.join(this.dir('pictures'), name);
    return name && fs.existsSync(p) ? p : null;
  }
  async savePictureData(buf, id, title) {
    const name = safeName(title || 'Picture') + ' [' + safeName(id, 40) + '].jpg';
    await fs.promises.writeFile(path.join(this.dir('pictures'), name), buf);
    return name;
  }
  deletePicture(name) { const p = this.picturePath(name); if (p) this.unlink(p); }
  videoPath(name) {
    name = path.basename(String(name || ''));
    const p = path.join(this.dir('videos'), name);
    return name && fs.existsSync(p) ? p : null;
  }
  async saveVideo(src, id, title) {
    const ext = (path.extname(src) || '.mp4').toLowerCase();
    const name = safeName(title || path.basename(src, ext)) + ' [' + safeName(id, 40) + ']' + ext;
    await fs.promises.copyFile(src, path.join(this.dir('videos'), name));
    return name;
  }
  async saveVideoData(buf, original, id, title) {
    const ext = (path.extname(original) || '.mp4').toLowerCase();
    const name = safeName(title || path.basename(original, ext)) + ' [' + safeName(id, 40) + ']' + ext;
    await fs.promises.writeFile(path.join(this.dir('videos'), name), buf);
    return name;
  }
  deleteVideo(name) { const p = this.videoPath(name); if (p) this.unlink(p); }

  /* ---------- bringing in data from the old (browser) version ---------- */
  importFrom(kv) {
    let n = 0;
    Object.keys(kv).forEach(k => {
      if (k.indexOf('hufc-') !== 0 || k === 'hufc-probe' || k === 'hufc-status') return;
      if (this.mem.has(k)) return;
      this.mem.set(k, String(kv[k])); this.write(k); n++;
    });
    return n;
  }

  /* a planner file double-clicked (from these folders, or one someone sent you): which item it is, adding it if it's new */
  openFile(file) {
    const ext = path.extname(file).toLowerCase();
    let kind = Object.keys(EXT).find(k => EXT[k] === ext);
    if (!kind && ext === '.json') {
      const parent = path.basename(path.dirname(file));
      kind = Object.keys(F).find(k => F[k] === parent && EXT[k]);
    }
    if (!kind) return null;
    let obj; try { obj = JSON.parse(fs.readFileSync(file, 'utf8')); } catch (e) { return null; }
    if (!obj || typeof obj !== 'object') return null;
    const id = idOf(path.basename(file)) || obj.planId || obj.id || obj.libId;
    if (!id) return null;
    const changed = [];
    if (kind === 'practices') {
      const lib = parse(this.mem.get(LIB_KEY) || 'null') || { v: 1, items: {}, removed: {} };
      lib.items = lib.items || {}; lib.removed = lib.removed || {};
      if (!lib.items[id] || (obj.savedAt || 0) > (lib.items[id].savedAt || 0)) {
        lib.items[id] = obj; delete lib.removed[id];
        this.set(LIB_KEY, JSON.stringify(lib)); changed.push([LIB_KEY, this.mem.get(LIB_KEY)]);
      }
      return { kind: 'practice', id, changed };
    }
    const key = (kind === 'lineups' ? P_LINEUP : kind === 'squads' ? P_SQUAD : P_SESSION) + id;
    const cur = parse(this.mem.get(key) || 'null');
    if (!cur || (obj.savedAt || 0) > (cur.savedAt || 0)) { this.set(key, JSON.stringify(obj)); changed.push([key, this.mem.get(key)]); }
    return { kind: kind === 'lineups' ? 'lineup' : kind === 'squads' ? 'squad' : (isPre(obj) ? 'pregame' : 'session'), id, changed };
  }
}
module.exports = { Store, FOLDERS: F, EXT };
