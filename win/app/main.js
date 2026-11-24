/* HUFC Season Planner: the planner as its own Windows program.
   The planner page is the same one used online; this wraps it in its own window and keeps
   everything it saves as files in a folder (Documents\HUFC Planner by default). */
'use strict';
const { app, BrowserWindow, ipcMain, shell, dialog, Menu, session } = require('electron');
const fs = require('fs');
const path = require('path');
const { Store, FOLDERS, EXT } = require('./store');
const { execFile } = require('child_process');

const APP_NAME = 'HUFC Season Planner';
app.setName(APP_NAME);
if (process.platform === 'win32') app.setAppUserModelId('au.com.hufc.seasonplanner');

if (!app.requestSingleInstanceLock()) { app.quit(); }

const LOCAL = process.env.LOCALAPPDATA || app.getPath('appData');
const PLANNER_DIR = process.env.HUFC_PLANNER_DIR || path.join(LOCAL, 'HUFC Session Planner');
const PLANNER_FILE = path.join(PLANNER_DIR, 'HUFC Session Planner.html');   /* Claude drops planner updates here */
const BUNDLED = path.join(__dirname, 'planner.html');
const CONFIG = path.join(app.getPath('userData'), 'config.json');
const LOG = path.join(app.getPath('userData'), 'planner.log');

function log() {
  try { fs.appendFileSync(LOG, new Date().toISOString() + ' ' + Array.prototype.join.call(arguments, ' ') + '\n'); } catch (e) {}
}
function readConfig() { try { return JSON.parse(fs.readFileSync(CONFIG, 'utf8')); } catch (e) { return {}; } }
function saveConfig(c) { try { fs.mkdirSync(path.dirname(CONFIG), { recursive: true }); fs.writeFileSync(CONFIG, JSON.stringify(c, null, 2)); } catch (e) { log('config', e.message); } }

let config = readConfig();
const dataDir = () => process.env.HUFC_DATA_DIR || config.dataDir || path.join(app.getPath('documents'), 'HUFC Planner');
let store = null, win = null;

/* whichever copy of the planner is newer: the one Claude updates, or the one that came with this install */
/* the planner page: the copy Claude keeps up to date (always, when it's there); the one that came with the install otherwise */
function plannerPath() {
  try { if (fs.statSync(PLANNER_FILE).size > 100000) return PLANNER_FILE; } catch (e) {}
  return BUNDLED;
}
/* when an update arrives while the planner is open, tell the page (it offers Reload; your work is saved first) */
function watchPlanner() {
  let last = 0;
  try { last = fs.statSync(PLANNER_FILE).mtimeMs; } catch (e) {}
  fs.watchFile(PLANNER_FILE, { interval: 4000 }, cur => {
    if (!cur || !cur.mtimeMs || cur.mtimeMs === last || cur.size < 100000) return;
    last = cur.mtimeMs;
    setTimeout(() => { try { if (win && !win.isDestroyed()) win.webContents.send('planner:updated'); } catch (e) {} }, 1500);   /* let the file finish writing */
  });
}

function openStore() {
  store = new Store(dataDir(), log).load();
  if (store.isEmpty() && !config.imported) {
    let kv = {};
    try { kv = require('./edge-import').readOldVersion(); } catch (e) { log('old version read failed', e.message); }
    if (!Object.keys(kv).length) {
      try { kv = JSON.parse(fs.readFileSync(path.join(__dirname, 'edge-import.json'), 'utf8')); } catch (e) {}
    }
    const n = store.importFrom(kv);
    log('imported', n, 'items from the old version');
    config.imported = new Date().toISOString(); saveConfig(config);
  }
  try { store.dailyBackup(); } catch (e) { log('backup', e.message); }
}

ipcMain.on('store:all', e => { e.returnValue = store ? store.all() : {}; });
ipcMain.on('store:info', e => { e.returnValue = { dataDir: dataDir(), version: app.getVersion(), folders: FOLDERS }; });
ipcMain.on('store:set', (e, k, v) => { if (store) store.set(String(k), String(v)); });
ipcMain.on('store:remove', (e, k) => { if (store) store.remove(String(k)); });
ipcMain.on('stats:write', (e, list) => { if (store) store.writeStats(list); });
ipcMain.handle('video:save', async (e, src, id, title) => {
  try { if (!src || !fs.existsSync(src)) return { error: 'missing' }; const file = await store.saveVideo(src, String(id), String(title || '')); store.backupVideos(); return { file }; }
  catch (err) { log('video save', err.message); return { error: err.message }; }
});
ipcMain.handle('video:saveData', async (e, bytes, name, id, title) => {
  try { const file = await store.saveVideoData(Buffer.from(bytes), String(name), String(id), String(title || '')); store.backupVideos(); return { file }; }
  catch (err) { log('video save', err.message); return { error: err.message }; }
});
ipcMain.handle('picture:saveData', async (e, bytes, id, title) => {
  try { const file = await store.savePictureData(Buffer.from(bytes), String(id), String(title || '')); store.backupMedia(); return { file }; }
  catch (err) { log('picture save', err.message); return { error: err.message }; }
});
ipcMain.on('picture:path', (e, name) => { e.returnValue = store ? store.picturePath(name) : null; });
ipcMain.on('picture:delete', (e, name) => { if (store) store.deletePicture(name); });
ipcMain.on('file:data', (e, file) => {
  try {
    const f = path.resolve(String(file)), root = path.resolve(dataDir());
    if (f.toLowerCase().indexOf(root.toLowerCase() + path.sep) !== 0) { e.returnValue = null; return; }   /* only from the planner folder */
    const ext = path.extname(f).toLowerCase(), type = ext === '.png' ? 'image/png' : ext === '.gif' ? 'image/gif' : ext === '.webp' ? 'image/webp' : 'image/jpeg';
    e.returnValue = 'data:' + type + ';base64,' + fs.readFileSync(f).toString('base64');
  } catch (err) { e.returnValue = null; }
});
/* Export > Open in PowerPoint: save the slides in the PowerPoint folder, then open them */
ipcMain.handle('ppt:open', async (e, bytes, name) => {
  try {
    const dir = path.join(dataDir(), FOLDERS.powerpoint); fs.mkdirSync(dir, { recursive: true });
    const base = String(name || 'HUFC slides.pptx').replace(/[\\/:*?"<>|]+/g, '-'), ext = path.extname(base) || '.pptx', stem = path.basename(base, ext);
    let file = path.join(dir, stem + ext), n = 1;
    for (;;) {
      try { fs.writeFileSync(file, Buffer.from(bytes)); break; }
      catch (err) { if (++n > 20) throw err; file = path.join(dir, stem + ' (' + n + ')' + ext); }   /* already open in PowerPoint: use a new name */
    }
    const err = await shell.openPath(file);
    return err ? { ok: false, error: err, name: path.basename(file) } : { ok: true, name: path.basename(file) };
  } catch (err) { log('ppt open', err.message); return { ok: false, error: err.message }; }
});
/* Print / PDF: the plan as a PDF (laid out exactly as it prints), opened in the computer's PDF viewer, which has a proper
   print preview (the built-in print window here can't show one). keep: saved in the PDFs folder; otherwise a temporary copy */
ipcMain.handle('print:pdf', async (e, name, keep) => {
  try {
    const data = await e.sender.printToPDF({ printBackground: true, preferCSSPageSize: true });
    const dir = keep ? path.join(dataDir(), FOLDERS.pdfs || 'PDFs') : path.join(app.getPath('temp'), 'HUFC Planner print');
    fs.mkdirSync(dir, { recursive: true });
    const stem = String(name || 'HUFC plan').replace(/[\\/:*?"<>|\x00-\x1f]+/g, '-').replace(/[. ]+$/, '').slice(0, 120) || 'HUFC plan';
    let file = path.join(dir, stem + '.pdf'), n = 1;
    for (;;) {
      try { fs.writeFileSync(file, data); break; }
      catch (err) { if (++n > 20) throw err; file = path.join(dir, stem + ' (' + n + ').pdf'); }   /* already open in the PDF viewer: use a new name */
    }
    const err = process.env.HUFC_NO_OPEN ? '' : await shell.openPath(file);      /* (HUFC_NO_OPEN: testing without a PDF viewer) */
    return err ? { ok: false, error: err, file } : { ok: true, file, name: path.basename(file) };
  } catch (err) { log('print pdf', err.message); return { ok: false, error: err.message }; }
});
/* a PDF the planner made (the same maker as online): saved, then opened in the computer's PDF viewer to print */
ipcMain.handle('pdf:open', async (e, bytes, name, keep) => {
  try {
    const dir = keep ? path.join(dataDir(), FOLDERS.pdfs || 'PDFs') : path.join(app.getPath('temp'), 'HUFC Planner print');
    fs.mkdirSync(dir, { recursive: true });
    const base = String(name || 'HUFC plan.pdf').replace(/[\\/:*?"<>|\x00-\x1f]+/g, '-'), stem = path.basename(base, '.pdf').slice(0, 120) || 'HUFC plan';
    let file = path.join(dir, stem + '.pdf'), n = 1;
    for (;;) {
      try { fs.writeFileSync(file, Buffer.from(bytes)); break; }
      catch (err) { if (++n > 20) throw err; file = path.join(dir, stem + ' (' + n + ').pdf'); }
    }
    const err = process.env.HUFC_NO_OPEN ? '' : await shell.openPath(file);
    return err ? { ok: false, error: err, file } : { ok: true, file, name: path.basename(file) };
  } catch (err) { log('pdf open', err.message); return { ok: false, error: err.message }; }
});
/* Save as…: the usual Windows save box (PDFs start in the planner's PDFs folder); nothing is opened afterwards */
ipcMain.handle('file:saveAs', async (e, bytes, name) => {
  try {
    const base = String(name || 'HUFC plan.pdf').replace(/[\\/:*?"<>|\x00-\x1f]+/g, '-').slice(0, 150);
    const ext = (path.extname(base) || '.pdf').slice(1).toLowerCase();
    const dir = ext === 'pdf' ? path.join(dataDir(), FOLDERS.pdfs || 'PDFs') : app.getPath('documents');
    fs.mkdirSync(dir, { recursive: true });
    const r = await dialog.showSaveDialog(BrowserWindow.fromWebContents(e.sender), {
      title: ext === 'pdf' ? 'Save the PDF' : 'Save the file', defaultPath: path.join(dir, base),
      filters: [{ name: ({ pdf: 'PDF', hufcsession: 'HUFC session', hufcpregame: 'HUFC pre-game', hufclineup: 'HUFC lineup', hufcbackup: 'HUFC backup', pptx: 'PowerPoint' })[ext] || ext.toUpperCase() + ' file', extensions: [ext] }] });
    if (r.canceled || !r.filePath) return { canceled: true };
    fs.writeFileSync(r.filePath, Buffer.from(bytes));
    return { ok: true, file: r.filePath, name: path.basename(r.filePath) };
  } catch (err) { log('save as', err.message); return { ok: false, error: err.message }; }
});
/* sharing: the file is kept in the planner's Shared folder, ready to attach or drag into any app */
function keepShared(bytes, name) {
  const dir = path.join(dataDir(), 'Shared'); fs.mkdirSync(dir, { recursive: true });
  const base = String(name || 'HUFC plan.pdf').replace(/[\\/:*?"<>|\x00-\x1f]+/g, '-'), ext = path.extname(base) || '.pdf', stem = path.basename(base, ext).slice(0, 120) || 'HUFC plan';
  let file = path.join(dir, stem + ext), n = 1;
  for (;;) {
    try { fs.writeFileSync(file, Buffer.from(bytes)); return file; }
    catch (err) { if (++n > 20) throw err; file = path.join(dir, stem + ' (' + n + ')' + ext); }
  }
}
ipcMain.handle('share:keep', async (e, bytes, name) => {
  try { return { ok: true, file: keepShared(bytes, name) }; } catch (err) { return { ok: false, error: err.message }; }
});
/* put a file on the clipboard (like Copy in File Explorer), so Ctrl+V in WhatsApp, Teams, Messenger or an email attaches it */
ipcMain.handle('file:copy', async (e, file) => new Promise(resolve => {
  if (process.platform !== 'win32' || !file || !fs.existsSync(file)) return resolve({ ok: false });
  const lit = String(file).replace(/'/g, "''");
  execFile('powershell.exe', ['-NoProfile', '-NonInteractive', '-STA', '-Command', "Set-Clipboard -LiteralPath '" + lit + "'"], { windowsHide: true, timeout: 15000 },
    err => resolve(err ? { ok: false, error: err.message } : { ok: true }));
}));
/* drag a file out of the planner straight into WhatsApp, Teams, an email… */
const DRAG_ICON = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACgAAAAwCAYAAABjezibAAAAxElEQVR4nO3ZQQ6CMBSE4cG4Fe7j/eN9wAPowjRK1PaRSftmMf+KBUm/FMILKeC4pnJxmZfHqEXv2zq173q1A263tY/oo/m6AIgjT101laJPLA0IxJCpQKCNTAcCdaQEEPiPlAECv5FSQOAbKQcE9khJIPBGnkcvHJ1WZeLI7mDJQDYD2QxkM5BNHtgcdWXk9Kw2/uR30EC25js44me+lvwOGsiW8h088l7L76CBbP4OshnIZiCbgWwGsqWc1UU7cqaX1hMViigloS7llQAAAABJRU5ErkJggg==';
ipcMain.on('file:drag', (e, file) => {
  try {
    if (!file || !fs.existsSync(file)) return;
    const { nativeImage } = require('electron');
    e.sender.startDrag({ file, icon: nativeImage.createFromDataURL(DRAG_ICON) });
  } catch (err) { log('drag', err.message); }
});
ipcMain.on('file:reveal', (e, file) => { try { if (file && fs.existsSync(file)) shell.showItemInFolder(file); } catch (err) {} });
ipcMain.on('open:external', (e, url) => { url = String(url || ''); if (/^(mailto:|https:\/\/)/i.test(url)) shell.openExternal(url); });
/* email: a new Outlook message with the file already attached; without Outlook, the email program opens and the file's folder is shown */
function outlookPath() {
  return new Promise(resolve => {
    if (process.platform !== 'win32') return resolve(null);
    execFile('reg', ['query', 'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\OUTLOOK.EXE', '/ve'], { windowsHide: true }, (err, out) => {
      if (err || !out) return resolve(null);
      const m = /REG_SZ\s+(.+)/.exec(out); const p = m ? m[1].trim().replace(/^"|"$/g, '') : null;
      resolve(p && fs.existsSync(p) ? p : null);
    });
  });
}
ipcMain.handle('mail:attach', async (e, bytes, name, subject, body) => {
  try {
    const file = keepShared(bytes, name);
    const ol = await outlookPath();
    if (ol) {
      const { spawn } = require('child_process');
      const child = spawn(ol, ['/c', 'ipm.note', '/a', file], { detached: true, stdio: 'ignore', windowsHide: false });
      child.on('error', () => {});
      child.unref();
      return { ok: true, how: 'outlook', file };
    }
    shell.openExternal('mailto:?subject=' + encodeURIComponent(subject || '') + '&body=' + encodeURIComponent(body || ''));
    shell.showItemInFolder(file);
    return { ok: true, how: 'mailto', file };
  } catch (err) { log('mail', err.message); return { ok: false, error: err.message }; }
});
ipcMain.on('video:path', (e, name) => { e.returnValue = store ? store.videoPath(name) : null; });
/* where a downloaded video came from: Windows notes the web page and download address on every downloaded file (its "Zone.Identifier") */
ipcMain.handle('video:source', (e, p) => {
  try {
    const t = fs.readFileSync(String(p || '') + ':Zone.Identifier', 'utf8');
    const get = k => { const m = t.match(new RegExp('^' + k + '=(.+)$', 'mi')); return m ? m[1].trim() : ''; };
    const ok = u => /^https?:\/\//i.test(u) ? u : '';
    return { page: ok(get('ReferrerUrl')), file: ok(get('HostUrl')) };
  } catch (err) { return null; }
});
/* Google Drive for videos (see gdrive.js): connect once, then the planner puts its videos there so every device and coach can play them */
const drive = require('./gdrive')({ getConfig: () => config, saveConfig: c => { config = c; saveConfig(c); }, log });
const driveOut = p => p.then(v => ({ ok: true, v }), e => { log('drive', e && e.message); return { ok: false, error: (e && e.message) || 'error', code: (e && e.code) || '' }; });
ipcMain.handle('gdrive:status', () => drive.status());
ipcMain.handle('gdrive:connect', (e, client) => driveOut(drive.connect(client)));
ipcMain.handle('gdrive:disconnect', () => driveOut(drive.disconnect()));
ipcMain.handle('gdrive:remove', (e, id) => driveOut(drive.remove(String(id || ''))));
ipcMain.handle('gdrive:quota', () => driveOut(drive.quota()));
/* a video in the planner's Videos folder (by its file name), or sent as bytes */
/* a full match video, from wherever it is on this computer (it carries on where it got to if it was cut short) */
ipcMain.handle('gdrive:uploadPath', (e, job, filePath, fileName, mime, resumeKey) => {
  const p = String(filePath || '');
  if (!/\.(mp4|m4v|mov|webm|mkv|avi)$/i.test(p) || !fs.existsSync(p)) return { ok: false, error: 'That video isn’t on this computer any more', code: 'missing' };
  const sender = e.sender;
  return driveOut(drive.upload(p, String(fileName || path.basename(p)), String(mime || ''), f => { try { if (!sender.isDestroyed()) sender.send('gdrive:progress', job, f); } catch (x) {} }, String(resumeKey || '')));
});
ipcMain.handle('gdrive:upload', (e, job, localName, bytes, fileName, mime) => {
  const src = localName ? (store ? store.videoPath(String(localName)) : null) : (bytes ? Buffer.from(bytes) : null);
  if (!src) return { ok: false, error: 'That video isn’t on this computer', code: 'missing' };
  const sender = e.sender;
  return driveOut(drive.upload(src, String(fileName || 'clip.mp4'), String(mime || ''), f => { try { if (!sender.isDestroyed()) sender.send('gdrive:progress', job, f); } catch (x) {} }));
});
ipcMain.on('video:delete', (e, name) => { if (store) store.deleteVideo(name); });
/* the ☰ button on the title bar: the program's menu (folder, back up, zoom, about…) */
ipcMain.on('menu:popup', (e, x, y) => { const m = Menu.getApplicationMenu(); if (m && win && !win.isDestroyed()) m.popup({ window: win, x: Math.round(+x || 0), y: Math.round(+y || 0) }); });
ipcMain.on('open:data', (e, sub) => { shell.openPath(sub && FOLDERS[sub] ? path.join(dataDir(), FOLDERS[sub]) : dataDir()); });

/* ---------- double-clicking a planner file opens it here ---------- */
const TYPES = {
  sessions:  ['HUFC.Session',  'HUFC session plan'],
  pregames:  ['HUFC.PreGame',  'HUFC pre-game plan'],
  lineups:   ['HUFC.Lineup',   'HUFC matchday lineup'],
  practices: ['HUFC.Practice', 'HUFC practice'],
  squads:    ['HUFC.Squad',    'HUFC squad']
};
function registerFileTypes() {
  if (process.platform !== 'win32') return;
  const exe = process.execPath, ico = path.join(__dirname, 'icon.ico');
  const stamp = exe + '|' + app.getVersion();
  if (config.fileTypes === stamp) return;
  const cmds = [];
  Object.keys(TYPES).forEach(k => {
    const [progId, label] = TYPES[k], root = 'HKCU\\Software\\Classes\\';
    cmds.push([root + EXT[k], '/ve', '/d', progId, '/f']);
    cmds.push([root + progId, '/ve', '/d', label, '/f']);
    cmds.push([root + progId + '\\DefaultIcon', '/ve', '/d', '"' + ico + '",0', '/f']);
    cmds.push([root + progId + '\\shell\\open\\command', '/ve', '/d', '"' + exe + '" "%1"', '/f']);
  });
  let left = cmds.length, ok = true;
  cmds.forEach(c => execFile('reg', ['add'].concat(c), { windowsHide: true }, err => {
    if (err) { ok = false; log('file type', c[0], err.message); }
    if (--left === 0 && ok) {
      config.fileTypes = stamp; saveConfig(config);
      execFile('ie4uinit.exe', ['-show'], { windowsHide: true }, () => {});   /* refresh Explorer's icons */
    }
  }));
}
const fileArgs = argv => argv.slice(1).filter(a => a && a[0] !== '-' && /\.(hufc[a-z]+|json)$/i.test(a) && fs.existsSync(a));
let pageReady = false, pending = [];
function openFiles(list) {
  list.forEach(f => {
    let r = null;
    try { r = store.openFile(f); } catch (e) { log('open file', f, e.message); }
    if (!r) { dialog.showErrorBox(APP_NAME, 'That file couldn’t be opened in the planner:\n' + f); return; }
    if (!win || !pageReady) { pending.push(r); return; }
    r.changed.forEach(([k, v]) => win.webContents.send('store:external', k, v));
    win.webContents.send('open-item', r.kind, r.id);
    if (win.isMinimized()) win.restore(); win.focus();
  });
}

/* =====================================================================
   The online planner, linked: a hidden window keeps the online planner open (signed in once),
   and the planner here syncs through it — the same club library as online, phone and every computer.
   ===================================================================== */
const ONLINE_URL = process.env.HUFC_ONLINE_URL || 'https://claude.ai/artifact/QuqVxaWUqC2mmARNfwJYdR';
const ONLINE_OPS = ['status', 'hasDir', 'exists', 'readAll', 'read', 'write', 'remove', 'refresh', 'sharesList', 'sharesAdd', 'sharesRemove', 'names', 'plannerInfo', 'plannerFetch'];
let onlineWin = null, onlineFrame = null, onlineState = { state: 'off' }, onlinePolling = false, onlineSignin = false, quitting = false;
function setOnline(s) {
  if (JSON.stringify(s) === JSON.stringify(onlineState)) return;
  onlineState = s;
  try { if (win && !win.isDestroyed()) win.webContents.send('online:state', s); } catch (e) {}
}
function startOnline(show) {
  if (!config.online) { config.online = true; saveConfig(config); }
  if (!onlineWin || onlineWin.isDestroyed()) {
    onlineWin = new BrowserWindow({ show: false, width: 1150, height: 840, title: 'Sign in to Claude · HUFC online planner', icon: path.join(__dirname, 'icon.ico'), autoHideMenuBar: true,
      webPreferences: { partition: 'persist:hufc-online', contextIsolation: true, sandbox: true, backgroundThrottling: false } });
    /* look like an ordinary browser, so the sign-in pages work */
    try { onlineWin.webContents.setUserAgent(onlineWin.webContents.getUserAgent().replace(/\s*Electron\/\S+/i, '').replace(/\s*hufc-season-planner\/\S+/i, '').replace(/\s*HUFC Season Planner\/\S+/i, '')); } catch (e) {}
    onlineWin.on('close', e => { if (!quitting) { e.preventDefault(); onlineWin.hide(); onlineSignin = false; } });
    onlineWin.on('closed', () => { onlineWin = null; onlineFrame = null; });
    onlineWin.webContents.setWindowOpenHandler(({ url }) => ({ action: 'allow', overrideBrowserWindowOptions: { autoHideMenuBar: true, webPreferences: { partition: 'persist:hufc-online' } } }));
    onlineWin.webContents.on('did-finish-load', () => { onlineFrame = null; pollOnline(); });
    onlineWin.loadURL(ONLINE_URL).catch(err => { log('online load', err.message); setOnline({ state: 'offline' }); });
    setOnline({ state: 'connecting' });
  }
  if (show) { onlineSignin = true; onlineWin.show(); onlineWin.focus(); }
  pollOnline();
}
async function findBridge() {
  if (!onlineWin || onlineWin.isDestroyed()) return null;
  let frames = [];
  try { frames = onlineWin.webContents.mainFrame.framesInSubtree; } catch (e) { return null; }
  for (const f of frames) {
    try { if (await f.executeJavaScript('!!(window.__HUFC_BRIDGE && window.__HUFC_BRIDGE.call)')) return f; } catch (e) {}
  }
  return null;
}
async function pollOnline() {
  if (onlinePolling) return;
  onlinePolling = true;
  try {
    for (let i = 0; i < 900; i++) {          /* keeps trying (about half an hour) while you sign in */
      if (!onlineWin || onlineWin.isDestroyed()) return;
      const f = await findBridge();
      if (f) {
        let st = null;
        try { st = await f.executeJavaScript('window.__HUFC_BRIDGE.call("status", [])'); } catch (e) {}
        const v = st && st.v;
        if (v && v.ok && v.who && v.who.uid) {
          onlineFrame = f;
          setOnline({ state: 'on', who: v.who });
          setTimeout(checkPlannerUpdate, 4000);
          if (onlineSignin) { onlineSignin = false; onlineWin.hide(); }
          return;
        }
      }
      setOnline({ state: i > 6 ? 'signin' : 'connecting' });
      await new Promise(r => setTimeout(r, 2000));
    }
  } finally { onlinePolling = false; }
}
async function onlineCall(op, args) {
  if (ONLINE_OPS.indexOf(op) === -1) return { __err: 'unknown', code: 'unknown' };
  const js = 'window.__HUFC_BRIDGE.call(' + JSON.stringify(op) + ', ' + JSON.stringify(Array.isArray(args) ? args : []) + ')';
  for (let attempt = 0; attempt < 2; attempt++) {
    if (!onlineFrame) onlineFrame = await findBridge();
    if (!onlineFrame) break;
    try { return await onlineFrame.executeJavaScript(js); }
    catch (e) { onlineFrame = null; }
  }
  if (onlineState.state === 'on') { setOnline({ state: 'connecting' }); pollOnline(); }
  return { __err: 'Not connected to the online planner', code: 'offline' };
}
/* ---- planner updates, through the online planner: when a newer planner is there, fetch it, check it, swap it in
   (the page then offers Reload, as for any update). Checked when linked, and every half hour. ---- */
function buildOf(text) { const m = /<meta name="hufc-build" content="(\d{14})">/.exec(String(text || '').slice(0, 4000)); return m ? m[1] : ''; }
function currentBuild() { try { const fd = fs.openSync(plannerPath(), 'r'), b = Buffer.alloc(4000); fs.readSync(fd, b, 0, 4000, 0); fs.closeSync(fd); return buildOf(b.toString('utf8')); } catch (e) { return ''; } }
/* ---- updates from the club website (every program, linked to the online planner or not) ----
   win/planner.json says which planner is newest (build, sha256, size); win/planner.html is that planner.
   win/app/program.json lists the program's own files (name -> sha256); a changed file is fetched, checked and swapped in,
   and the page offers Restart. Only whole files whose sha256 matches are ever used. */
const crypto = require('crypto');
const vm = require('vm');
const WEB = 'https://4s45d77rv5dkambo.github.io/HUFC-Season-Planner/win/';
const sha256 = t => crypto.createHash('sha256').update(typeof t === 'string' ? Buffer.from(t, 'utf8') : t).digest('hex');
async function webGet(rel) {
  const { net } = require('electron');
  const r = await net.fetch(WEB + rel + (rel.indexOf('?') === -1 ? '?' : '&') + 't=' + Date.now(), { cache: 'no-store' });
  if (!r.ok) throw new Error(rel + ' ' + r.status);
  return r;
}
function swapPlanner(text, build, from) {
  fs.mkdirSync(PLANNER_DIR, { recursive: true });
  const tmp = PLANNER_FILE + '.new';
  fs.writeFileSync(tmp, text, 'utf8');
  fs.renameSync(tmp, PLANNER_FILE);
  log('planner updated to', build, 'from', from || '(none)');
}
function plannerOk(text, build) {
  return typeof text === 'string' && text.length > 1000000 && buildOf(text) === String(build) && /<\/html>\s*$/i.test(text) && !/__HUFC_SAVE/.test(text.slice(0, 200));
}
async function webPlanner() {
  const j = await (await webGet('planner.json')).json();
  if (!j || !/^\d{14}$/.test(String(j.build || ''))) return { state: 'error', error: 'no update information' };
  const have = currentBuild();
  if (have && String(j.build) <= have) return { state: 'latest', build: have };
  const text = await (await webGet('planner.html')).text();
  if (!plannerOk(text, j.build) || (j.sha256 && sha256(text) !== j.sha256)) { log('web update skipped: file not right', j.build); return { state: 'error', error: 'the download wasn’t complete' }; }
  swapPlanner(text, j.build, have);
  return { state: 'updated', build: String(j.build), from: have };
}
const PROGRAM_FILES = ['main.js', 'preload.js', 'store.js', 'gdrive.js', 'edge-import.js', 'package.json'];
async function webProgram() {
  const j = await (await webGet('app/program.json')).json();
  if (!j || !j.files || !j.version) return { state: 'error' };
  const newer = (x, y) => { const A = String(x).split('.').map(Number), B = String(y).split('.').map(Number); for (let i = 0; i < 3; i++) { if ((A[i] || 0) !== (B[i] || 0)) return (A[i] || 0) > (B[i] || 0); } return false; };
  if (!newer(j.version, app.getVersion())) return { state: 'latest', version: app.getVersion() };   /* only ever forward */
  const want = Object.keys(j.files).filter(n => PROGRAM_FILES.indexOf(n) !== -1);
  const todo = want.filter(n => { try { return sha256(fs.readFileSync(path.join(__dirname, n))) !== j.files[n]; } catch (e) { return true; } });
  if (!todo.length) return { state: 'latest', version: j.version };
  const got = {};
  for (const n of todo) {
    const buf = Buffer.from(await (await webGet('app/' + n)).arrayBuffer());
    if (sha256(buf) !== j.files[n]) { log('program update skipped: file not right', n); return { state: 'error', error: 'the download wasn’t complete' }; }
    if (/\.js$/.test(n)) { try { new vm.Script(buf.toString('utf8'), { filename: n }); } catch (e) { log('program update skipped: ', n, e.message); return { state: 'error', error: 'a file didn’t check out' }; } }
    if (/\.json$/.test(n)) { try { JSON.parse(buf.toString('utf8')); } catch (e) { return { state: 'error', error: 'a file didn’t check out' }; } }
    got[n] = buf;
  }
  /* all checked: swap them in together (the old ones are kept as .old) */
  for (const n of Object.keys(got)) {
    const f = path.join(__dirname, n);
    try { if (fs.existsSync(f)) fs.copyFileSync(f, f + '.old'); } catch (e) {}
    fs.writeFileSync(f + '.new', got[n]); fs.renameSync(f + '.new', f);
  }
  log('program updated to', j.version, todo.join(', '));
  return { state: 'updated', version: j.version };
}
let updating = false;
async function checkPlannerUpdate(manual) {
  if (updating) return { state: 'busy' };
  updating = true;
  const out = { planner: null, program: null };
  try {
    /* 1. the club website */
    try { out.planner = await webPlanner(); } catch (e) { out.planner = { state: 'error', error: /fetch|net::|ENOTFOUND|ERR_/i.test(e.message) ? 'offline' : e.message }; log('web update check', e.message); }
    try { out.program = await webProgram(); } catch (e) { out.program = { state: 'error', error: e.message }; log('program update check', e.message); }
    /* 2. linked to the online planner: it may have a newer one still */
    if (onlineState.state === 'on') {
      try {
        const info = await onlineCall('plannerInfo', []);
        const v = info && info.v;
        const have = currentBuild();
        if (v && /^\d{14}$/.test(String(v.build || '')) && v.asset && (!have || String(v.build) > have)) {
          const got = await onlineCall('plannerFetch', [v.asset]);
          if (plannerOk(got && got.v, v.build)) { swapPlanner(got.v, v.build, have); out.planner = { state: 'updated', build: String(v.build), from: have }; }
          else log('update skipped: file not right', v.build);
        }
      } catch (e) { log('update check', e.message); }
    }
  } finally { updating = false; checkPlannerUpdate.last = Date.now(); }
  if (out.planner && out.planner.state === 'latest') out.planner.build = currentBuild();
  return out;
}
setInterval(() => { checkPlannerUpdate(); }, 60 * 1000);   /* every minute: updates come in straight away, as before */
setTimeout(() => { checkPlannerUpdate(); }, 8000);              /* and soon after opening */
app.on('browser-window-focus', () => { if (Date.now() - (checkPlannerUpdate.last || 0) > 20000) checkPlannerUpdate(); });
/* Settings → Check for updates */
ipcMain.handle('update:check', async () => {
  for (let i = 0; i < 40 && updating; i++) await new Promise(r => setTimeout(r, 500));
  const r = await checkPlannerUpdate(true);
  return Object.assign({ have: currentBuild(), programVersion: app.getVersion() }, r);
});
ipcMain.handle('online:state', () => onlineState);
ipcMain.handle('online:connect', () => { startOnline(true); return onlineState; });
ipcMain.handle('online:show', () => { if (onlineWin && !onlineWin.isDestroyed()) { onlineWin.show(); onlineWin.focus(); } else startOnline(true); return onlineState; });
ipcMain.handle('online:disconnect', () => {
  config.online = false; saveConfig(config);
  if (onlineWin && !onlineWin.isDestroyed()) { quitting = true; onlineWin.destroy(); quitting = false; }
  onlineWin = null; onlineFrame = null; setOnline({ state: 'off' });
  return onlineState;
});
ipcMain.handle('online:call', (e, op, args) => onlineCall(String(op || ''), args));

/* ---- program updates: restart from inside the planner (the page offers it) ---- */
ipcMain.on('app:relaunch', () => {
  try { if (store) store.flush(); } catch (e) {}
  quitting = true;
  app.relaunch(); app.exit(0);
});
/* the program's own files changed (an update arrived): tell the page, which offers Restart */
function watchProgram() {
  const files = ['main.js', 'preload.js', 'store.js', 'package.json'].map(n => path.join(__dirname, n));
  const seen = {};
  files.forEach(f => { try { seen[f] = fs.statSync(f).mtimeMs; } catch (e) {} });
  files.forEach(f => fs.watchFile(f, { interval: 5000 }, cur => {
    if (!cur || !cur.mtimeMs || cur.mtimeMs === seen[f]) return;
    seen[f] = cur.mtimeMs;
    setTimeout(() => { try { if (win && !win.isDestroyed()) win.webContents.send('program:updated'); } catch (e) {} }, 1500);
  }));
}

function createWindow() {
  const b = config.bounds || {};
  win = new BrowserWindow({
    width: b.width || 1320, height: b.height || 960, x: b.x, y: b.y,
    minWidth: 360, minHeight: 480,
    title: APP_NAME, icon: path.join(__dirname, 'icon.ico'), backgroundColor: '#f3f3f0', show: false,
    /* its own title bar: the planner draws a black bar with the crest; Windows puts minimise / maximise / close on it */
    titleBarStyle: 'hidden', titleBarOverlay: { color: '#0d0e0f', symbolColor: '#ffffff', height: 38 },
    webPreferences: { preload: path.join(__dirname, 'preload.js'), contextIsolation: false, nodeIntegration: false, sandbox: false, spellcheck: true }
  });
  if (config.maximized) win.maximize();
  win.once('ready-to-show', () => win.show());
  win.loadFile(plannerPath());
  /* right-click: spelling suggestions for a misspelt word, Add to dictionary, and cut / copy / paste */
  win.webContents.on('context-menu', (e, p) => {
    const items = [];
    if (p.misspelledWord) {
      (p.dictionarySuggestions || []).slice(0, 6).forEach(s => items.push({ label: s, click: () => win.webContents.replaceMisspelling(s) }));
      if (!items.length) items.push({ label: 'No suggestions', enabled: false });
      items.push({ type: 'separator' }, { label: 'Add "' + p.misspelledWord + '" to dictionary', click: () => win.webContents.session.addWordToSpellCheckerDictionary(p.misspelledWord) }, { type: 'separator' });
    }
    if (p.isEditable) {
      const f = p.editFlags || {};
      items.push({ role: 'undo', enabled: f.canUndo !== false }, { role: 'redo', enabled: f.canRedo !== false }, { type: 'separator' },
        { role: 'cut', enabled: !!f.canCut }, { role: 'copy', enabled: !!f.canCopy }, { role: 'paste', enabled: !!f.canPaste }, { type: 'separator' }, { role: 'selectAll' });
    } else if (p.selectionText && p.selectionText.trim()) items.push({ role: 'copy' });
    if (items.length) Menu.buildFromTemplate(items).popup({ window: win });
  });
  win.webContents.on('did-finish-load', () => {
    pageReady = true;
    const p = pending; pending = [];
    p.forEach(r => { r.changed.forEach(([k, v]) => win.webContents.send('store:external', k, v)); win.webContents.send('open-item', r.kind, r.id); });
  });
  win.webContents.on('did-start-loading', () => { pageReady = false; });
  /* web links (YouTube, Veo, ...) open in the normal browser */
  win.webContents.setWindowOpenHandler(({ url }) => { if (/^https?:/i.test(url)) shell.openExternal(url); return { action: 'deny' }; });
  win.webContents.on('will-navigate', (e, url) => { if (!/^file:/i.test(url)) { e.preventDefault(); if (/^https?:/i.test(url)) shell.openExternal(url); } });
  win.on('page-title-updated', e => e.preventDefault());
  const remember = () => {
    if (!win || win.isDestroyed()) return;
    config.maximized = win.isMaximized();
    if (!config.maximized && !win.isMinimized()) config.bounds = win.getBounds();
    saveConfig(config);
  };
  win.on('close', () => { remember(); if (store) store.flush(); });
  win.on('closed', () => { win = null; quitting = true; if (onlineWin && !onlineWin.isDestroyed()) onlineWin.destroy(); });
}

async function changeDataFolder() {
  const r = await dialog.showOpenDialog(win, { title: 'Choose where to keep your planner data (e.g. your OneDrive)', properties: ['openDirectory', 'createDirectory'] });
  if (r.canceled || !r.filePaths[0]) return;
  let target = r.filePaths[0];
  if (path.basename(target).toLowerCase() !== 'hufc planner') target = path.join(target, 'HUFC Planner');
  const from = dataDir();
  if (path.resolve(target) === path.resolve(from)) return;
  store.flush();
  const hasData = fs.existsSync(path.join(target, 'Sessions')) && fs.readdirSync(path.join(target, 'Sessions')).length > 0;
  if (hasData) {
    const a = await dialog.showMessageBox(win, { type: 'question', buttons: ['Use that folder', 'Cancel'], defaultId: 0, cancelId: 1,
      message: 'That folder already has planner data in it.', detail: 'Switch to it? Your current data stays where it is (' + from + ').' });
    if (a.response !== 0) return;
  } else {
    try {
      fs.mkdirSync(target, { recursive: true });
      Object.keys(FOLDERS).forEach(k => { const s = path.join(from, FOLDERS[k]); if (fs.existsSync(s)) fs.cpSync(s, path.join(target, FOLDERS[k]), { recursive: true }); });
      ['About these folders.txt'].forEach(n => { const s = path.join(from, n); if (fs.existsSync(s)) fs.copyFileSync(s, path.join(target, n)); });
    } catch (e) { dialog.showErrorBox('Couldn’t copy your data', e.message); return; }
  }
  config.dataDir = target; config.imported = config.imported || 'moved'; saveConfig(config);
  await dialog.showMessageBox(win, { type: 'info', message: 'Your planner data is now kept in:', detail: target + (hasData ? '' : '\n\nA copy of everything was moved there. The old folder is still at ' + from + ' — you can delete it once you’re happy.') });
  app.relaunch(); app.exit(0);
}

function buildMenu() {
  const t = [
    { label: 'File', submenu: [
      { label: 'Open my planner folder', click: () => shell.openPath(dataDir()) },
      { label: 'Open the Stats folder', click: () => shell.openPath(path.join(dataDir(), FOLDERS.stats)) },
      { label: 'Back up now', click: () => {
        store.flush();
        const stamp = new Date().toISOString().slice(0, 16).replace(/[:T]/g, '-');
        const dest = path.join(dataDir(), FOLDERS.backups, stamp);
        ['sessions', 'pregames', 'lineups', 'practices', 'squads', 'stats', 'settings'].forEach(k => {
          try { fs.cpSync(path.join(dataDir(), FOLDERS[k]), path.join(dest, FOLDERS[k]), { recursive: true }); } catch (e) {}
        });
        store.backupMedia();
        dialog.showMessageBox(win, { type: 'info', message: 'Backed up', detail: dest });
      } },
      { label: 'Change where my data is kept…', click: changeDataFolder },
      { label: 'Link to the online planner…', click: () => startOnline(true) },
      { type: 'separator' },
      { role: 'quit', label: 'Exit' }
    ] },
    { label: 'Edit', submenu: [{ role: 'undo' }, { role: 'redo' }, { type: 'separator' }, { role: 'cut' }, { role: 'copy' }, { role: 'paste' }, { role: 'selectAll' }] },
    { label: 'View', submenu: [
      { role: 'reload', label: 'Reload' }, { type: 'separator' },
      { role: 'zoomIn', label: 'Bigger' }, { role: 'zoomOut', label: 'Smaller' }, { role: 'resetZoom', label: 'Actual size' },
      { type: 'separator' }, { role: 'togglefullscreen', label: 'Full screen' },
      { type: 'separator' }, { role: 'toggleDevTools', label: 'Developer tools' }
    ] },
    { label: 'Help', submenu: [
      { label: 'About ' + APP_NAME, click: () => dialog.showMessageBox(win, { type: 'info', title: APP_NAME, message: APP_NAME + ' ' + app.getVersion(),
        detail: 'Heidelberg United FC\n\nYour data is kept in:\n' + dataDir() + '\n\nPlanner page:\n' + plannerPath() }) }
    ] }
  ];
  Menu.setApplicationMenu(Menu.buildFromTemplate(t));
}

app.on('second-instance', (e, argv) => {
  if (win) { if (win.isMinimized()) win.restore(); win.focus(); }
  if (store) openFiles(fileArgs(argv));
});
app.whenReady().then(() => {
  try { openStore(); }
  catch (e) {
    log('store failed', e.stack || e.message);
    dialog.showErrorBox(APP_NAME, 'Couldn’t open your planner folder:\n' + dataDir() + '\n\n' + e.message);
    app.exit(1); return;
  }
  /* YouTube won't show its player to a page with no web address; tell it the planner's */
  session.defaultSession.webRequest.onBeforeSendHeaders(
    { urls: ['*://*.youtube.com/*', '*://*.youtube-nocookie.com/*', '*://youtube.com/*', '*://youtube-nocookie.com/*', 'https://www.googleapis.com/drive/v3/files/*'] },
    (d, cb) => {
      /* (match videos streamed from Google Drive: the planner's key only works for the planner's own web address) */
      d.requestHeaders['Referer'] = /googleapis\.com\/drive\//.test(d.url) ? 'https://4s45d77rv5dkambo.github.io/HUFC-Season-Planner/' : 'https://hufc-season-planner.app/';
      cb({ requestHeaders: d.requestHeaders });
    });
  try {
    const have = session.defaultSession.availableSpellCheckerLanguages || [];
    const want = ['en-AU', 'en-GB', 'en-US'].filter(l => !have.length || have.includes(l));
    if (want.length) session.defaultSession.setSpellCheckerLanguages(want);
  } catch (e) {}
  buildMenu();
  watchPlanner();
  watchProgram();
  createWindow();
  if (config.online) startOnline(false);   /* linked before: reconnect in the background */
  registerFileTypes();
  openFiles(fileArgs(process.argv));
});
app.on('before-quit', () => { quitting = true; if (store) store.flush(); });
app.on('window-all-closed', () => app.quit());
