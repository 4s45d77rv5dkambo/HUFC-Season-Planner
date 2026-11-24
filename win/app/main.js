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

/* ---- Veo: a match's stats, read the way the Veo website reads them ----
   A hidden window opens the match's Analytics page in Veo (signed in to Veo inside this program, once), and keeps
   the stats Veo's own page loads: team stats, each player's running numbers, the halves and Veo's goals and shots.
   Nothing is typed in for the coach: if Veo wants a sign-in, the window shows and the coach signs in there. */
const VEO_PART = 'persist:hufc-veo';
let veoJob = null, veoWin = null;
const veoWanted = u => /^https:\/\/app\.veo\.co\/api\/(app|mes)\//.test(u) && !/player-tracking|step-events|\/i18n\//.test(u);
function veoRun(slug) {
  const target = 'https://app.veo.co/matches/' + slug + '/#/analysis/';
  return new Promise(resolve => {
    const got = {}, reqs = {};
    let finished = false, settleT = null, quietT = null, overall = null, nudged = 0, signin = false, userStatus = 0;
    const w = new BrowserWindow({ show: false, width: 1400, height: 900, title: 'Sign in to Veo · ' + APP_NAME, icon: path.join(__dirname, 'icon.ico'), autoHideMenuBar: true,
      webPreferences: { partition: VEO_PART, contextIsolation: true, sandbox: true, backgroundThrottling: false } });
    veoWin = w;
    try { w.webContents.setUserAgent(w.webContents.getUserAgent().replace(/\s*Electron\/\S+/i, '').replace(/\s*hufc-season-planner\/\S+/i, '').replace(/\s*HUFC Season Planner\/\S+/i, '')); } catch (e) {}
    w.webContents.setWindowOpenHandler(() => ({ action: 'allow', overrideBrowserWindowOptions: { autoHideMenuBar: true, webPreferences: { partition: VEO_PART } } }));
    const dbg = w.webContents.debugger;
    const done = out => {
      if (finished) return; finished = true;
      clearTimeout(settleT); clearTimeout(quietT); clearTimeout(overall);
      try { dbg.detach(); } catch (e) {}
      try { if (!w.isDestroyed()) w.destroy(); } catch (e) {}
      veoWin = null;
      log('veo', slug, out.error || ('ok' + (out.partial ? ' (partial)' : '')), keys().map(u => u.replace(/^https:\/\/app\.veo\.co\/api\//, '').replace(/[0-9a-f-]{36}|\d{8}-[a-z0-9-]+/g, '…').slice(0, 60)).join(' | '));
      resolve(out);
    };
    const keys = () => Object.keys(got);
    const has = re => keys().some(u => re.test(u));
    const complete = () => has(/\/analysis\/stats\//) && has(/\/periods\/?$/) && has(/\/highlights\//) && has(/physical-metrics/);
    const finishSoon = ms => { clearTimeout(settleT); settleT = setTimeout(() => done(keys().length ? { ok: true, partial: !complete(), data: got } : { error: signin ? 'signin' : 'nothing' }), ms); };
    /* everything in: a moment for stragglers, then back to the planner */
    const check = () => { if (complete()) finishSoon(1200); };
    /* Veo's page loaded and went quiet without its Analytics panel: open it (as a coach would), then settle for what came */
    const quiet = () => {
      clearTimeout(quietT);
      quietT = setTimeout(async () => {
        if (finished || signin || complete()) return;
        if (nudged < 2) {
          nudged++;
          try { await w.webContents.executeJavaScript('(() => { const els = [...document.querySelectorAll("button, a, [role=tab], [role=button]")]; const b = els.find(e => /analytics|stats/i.test((e.getAttribute("aria-label") || "") + " " + (e.title || "") + " " + (e.textContent || "").trim().slice(0, 40))); if (b) { b.click(); return true; } return false; })()'); } catch (e) {}
          quiet();
        } else if (has(/\/api\/app\/matches\//)) finishSoon(500);
      }, 7000);
    };
    try { dbg.attach('1.3'); } catch (e) { log('veo debugger', e.message); try { w.destroy(); } catch (x) {} veoWin = null; resolve({ error: 'internal' }); return; }
    dbg.on('message', (ev, method, p) => {
      try {
        if (method === 'Network.requestWillBeSent' && p.request && /\/analysis\/stats\//.test(p.request.url) && p.request.postData) got.__statsRequest = String(p.request.postData).slice(0, 4000);
        else if (method === 'Network.responseReceived') {
          const u = p.response && p.response.url || '';
          if (/^https:\/\/app\.veo\.co\/api\/app\/user\/?(\?|$)/.test(u)) userStatus = p.response.status;
          if (veoWanted(u)) reqs[p.requestId] = { url: u, status: p.response.status };
        } else if (method === 'Network.loadingFinished') {
          const r = reqs[p.requestId]; if (!r) return; delete reqs[p.requestId];
          if (r.status !== 200) return;
          dbg.sendCommand('Network.getResponseBody', { requestId: p.requestId }).then(b => {
            const t = b.base64Encoded ? Buffer.from(b.body, 'base64').toString('utf8') : b.body;
            if (t && t.length < 5e6) { try { got[r.url] = JSON.parse(t); } catch (e) {} }
            check(); quiet();
          }).catch(() => {});
        }
      } catch (e) {}
    });
    dbg.sendCommand('Network.enable', { maxResourceBufferSize: 10e6, maxTotalBufferSize: 80e6 }).catch(() => {});
    /* not signed in to Veo (or Veo's sign-in page): show the window so the coach signs in; then back to the match */
    const needSignin = () => {
      if (signin || finished) return; signin = true;
      try { if (!w.isDestroyed()) { w.show(); w.focus(); } } catch (e) {}
      if (win && !win.isDestroyed()) win.webContents.send('veo:signin');
    };
    /* where the window goes (for the log: the address without its query, so nothing private is kept) */
    const bare = u => String(u || '').split(/[?#]/)[0].slice(0, 120);
    w.webContents.on('did-navigate', (e, url) => {
      log('veo page', bare(url), signin ? '(signing in)' : '');
      if (/^https:\/\/(auth|login|accounts?|id)\.veo\.co\//.test(url)) needSignin();
      /* signed in: Veo goes on to the coach's profile page, so back to this match's Analytics */
      else if (signin && /^https:\/\/profile\.veo\.co\//.test(url)) setTimeout(() => { if (!finished) w.loadURL(target).catch(() => {}); }, 1500);
    });
    w.webContents.on('did-redirect-navigation', (e, url) => { log('veo redirect', bare(url)); });
    w.webContents.on('did-fail-load', (e, code, desc, url, main) => { if (main) log('veo load failed', code, desc, bare(url)); });
    /* each Veo page: signed out (show the window to sign in), part-way through signing in (leave it be),
       or signed in (hide the window and go to this match's Analytics) */
    w.webContents.on('did-finish-load', () => {
      const u = w.webContents.getURL();
      if (!/^https:\/\/app\.veo\.co\//.test(u)) return;
      setTimeout(async () => {
        if (finished) return;
        let st = {};
        try { st = await w.webContents.executeJavaScript('(() => { const t = document.body ? document.body.innerText : ""; return { wall: /log in to access/i.test(t.slice(0, 20000)), login: [...document.querySelectorAll("button, a")].some(e => /^\\s*(log ?in|sign ?in)\\s*$/i.test(e.textContent || "")), full: t.length > 150 }; })()') || {}; } catch (e) {}
        log('veo check', bare(u), (st.wall ? 'wall ' : '') + (st.login ? 'login-button ' : '') + (st.full ? 'page ' : 'blank ') + userStatus);
        if (st.wall || (!st.full && (userStatus === 401 || userStatus === 403))) { needSignin(); return; }
        if (signin) {
          if (!st.full || st.login) return;          /* still on the way through Veo's sign-in */
          signin = false; nudged = 0;
          try { w.hide(); } catch (e) {}
          if (win && !win.isDestroyed()) win.webContents.send('veo:working');
          if (u !== target) { w.loadURL(target).catch(() => {}); return; }
        }
        quiet();
      }, 2500);
    });
    w.on('close', () => { if (!finished) done(keys().length && complete() ? { ok: true, data: got } : { error: signin ? 'signin' : 'closed' }); });
    /* never more than 10 minutes (signing in), or a minute and a half without */
    overall = setTimeout(() => finishSoon(0), 10 * 60 * 1000);
    setTimeout(() => { if (!signin && !finished) finishSoon(0); }, 90 * 1000);
    w.loadURL(target).catch(err => { log('veo load', err.message); done({ error: /ERR_INTERNET|ERR_NAME|ERR_CONNECTION|ERR_NETWORK/.test(err.message) ? 'offline' : 'load' }); });
  });
}
ipcMain.handle('veo:stats', async (e, url) => {
  const m = String(url || '').match(/^https?:\/\/(?:app\.)?veo\.co\/matches\/([A-Za-z0-9_-]+)/i);
  if (!m) return { error: 'link' };
  if (veoJob) { try { if (veoWin && !veoWin.isDestroyed() && veoWin.isVisible()) veoWin.focus(); } catch (x) {} return { error: 'busy' }; }
  veoJob = veoRun(m[1]);
  try { return await veoJob; } finally { veoJob = null; }
});
/* ---- Veo's player tracking (the mini map): Veo's page only loads it 24 seconds at a time, around where its video is.
   So a hidden Veo window steps its video through each half and keeps what the page loads; the positions are then
   thinned to one every 0.8 s and packed small (each: who, x, y). Nothing is typed into Veo or asked of it beyond its page. */
/* Veo's positions, thinned to every 0.8 s and packed: per moment a count, then for each one: who (1 us, 2 our keeper,
   3 them, 4 their keeper, 5 ball), its track number, x, y (each 0..65535 over -0.25..1.25 of the pitch) */
function veoPack(frames, ranges) {
  const times = Object.keys(frames).map(Number).sort((a, b) => a - b);
  if (!times.length) return null;
  const dt = 0.8, t0 = Math.max(times[0], ranges[0][0]), tEnd = Math.min(times[times.length - 1], ranges[ranges.length - 1][1]);
  const n = Math.max(0, Math.floor((tEnd - t0) / dt) + 1);
  const parts = []; let j = 0;
  const q = v => Math.max(0, Math.min(65535, Math.round((v + 0.25) / 1.5 * 65535)));
  for (let i = 0; i < n; i++) {
    const tt = t0 + i * dt;
    while (j + 1 < times.length && Math.abs(times[j + 1] - tt) <= Math.abs(times[j] - tt)) j++;
    const list = Math.abs(times[j] - tt) <= 0.5 ? frames[times[j].toFixed(2)] : null;
    const objs = [];
    (list || []).forEach(o => {
      if (!Array.isArray(o) || o.length < 8) return;
      const cls = o[1], team = o[7];
      let kind = 0;
      if (cls === 6) kind = 5; else if (team === 1) kind = (cls === 0 || cls === 2) ? 2 : 1; else if (team === 2) kind = (cls === 0 || cls === 2) ? 4 : 3;
      if (kind) objs.push([kind, (o[0] + 1) & 255, q(o[2]), q(o[3])]);
    });
    const c = Math.min(255, objs.length), b = Buffer.alloc(1 + c * 6); b[0] = c;
    objs.slice(0, c).forEach((o, k) => { const at = 1 + k * 6; b[at] = o[0]; b[at + 1] = o[1]; b.writeUInt16LE(o[2], at + 2); b.writeUInt16LE(o[3], at + 4); });
    parts.push(b);
  }
  return { t0: t0, dt: dt, n: n, data: Buffer.concat(parts).toString('base64') };
}
let veoTrackJob = null;
function veoTrackRun(slug, ranges) {
  const target = 'https://app.veo.co/matches/' + slug + '/';
  const WORKERS = 3;   /* three Veo pages at once, each taking its share of the game */
  return new Promise(resolve => {
    const frames = {}, starts = [], workers = [];
    let finished = false, overall = null;
    const t0run = Date.now();
    const sleep = ms => new Promise(r => setTimeout(r, ms));
    const covered = t => starts.some(s => s <= t + 0.5 && t < s + 23.5);
    const done = out => {
      if (finished) return; finished = true;
      clearTimeout(overall);
      workers.forEach(k => { try { k.dbg.detach(); } catch (e) {} try { if (!k.w.isDestroyed()) k.w.destroy(); } catch (e) {} });
      log('veo track', slug, (out.error || ('ok ' + starts.length + ' pieces')) + ' in ' + Math.round((Date.now() - t0run) / 1000) + 's');
      resolve(out);
    };
    /* one Veo page, drawn in memory (offscreen) so it runs as it does on screen, keeping the positions its page loads */
    const makeWorker = () => {
      const w = new BrowserWindow({ show: false, width: 1280, height: 800, title: 'Veo · ' + APP_NAME, icon: path.join(__dirname, 'icon.ico'), autoHideMenuBar: true,
        webPreferences: { partition: VEO_PART, contextIsolation: true, sandbox: true, backgroundThrottling: false, offscreen: true } });
      try { w.webContents.setFrameRate(15); } catch (e) {}
      try { w.webContents.setUserAgent(w.webContents.getUserAgent().replace(/\s*Electron\/\S+/i, '').replace(/\s*hufc-season-planner\/\S+/i, '').replace(/\s*HUFC Season Planner\/\S+/i, '')); } catch (e) {}
      try { w.webContents.setAudioMuted(true); } catch (e) {}
      const dbg = w.webContents.debugger, reqs = {};
      try { dbg.attach('1.3'); } catch (e) { try { w.destroy(); } catch (x) {} return null; }
      dbg.on('message', (ev, method, p) => {
        try {
          if (method === 'Network.responseReceived') {
            const u = p.response && p.response.url || '';
            if (/\/api\/mes\/v2\/[^/]+\/player-tracking\?start=/.test(u) && p.response.status === 200) reqs[p.requestId] = u;
          } else if (method === 'Network.loadingFinished') {
            const u = reqs[p.requestId]; if (!u) return; delete reqs[p.requestId];
            dbg.sendCommand('Network.getResponseBody', { requestId: p.requestId }).then(b => {
              const t = b.base64Encoded ? Buffer.from(b.body, 'base64').toString('utf8') : b.body;
              let j = null; try { j = JSON.parse(t); } catch (e) { return; }
              if (!j || typeof j !== 'object') return;
              Object.keys(j).forEach(k => { const tt = parseFloat(k); if (isFinite(tt) && Array.isArray(j[k])) frames[tt.toFixed(2)] = j[k]; });
              const st = parseFloat((u.match(/start=([\d.]+)/) || [])[1]); if (isFinite(st)) starts.push(st);
            }).catch(() => {});
          }
        } catch (e) {}
      });
      dbg.sendCommand('Network.enable', { maxResourceBufferSize: 10e6, maxTotalBufferSize: 100e6 }).catch(() => {});
      const js = c => w.isDestroyed() ? Promise.resolve(null) : w.webContents.executeJavaScript(c).catch(() => null);
      w.on('close', () => { if (!finished) done({ error: 'closed' }); });
      w.loadURL(target).catch(() => {});
      return { w: w, dbg: dbg, js: js };
    };
    const ready = async k => {
      for (let i = 0; i < 60 && !finished; i++) {
        await sleep(500);
        const st = await k.js('(() => { const v = document.querySelector("video"); const wall = /log in to access/i.test(document.body ? document.body.innerText.slice(0, 20000) : ""); return { wall: wall, ok: !!(v && v.readyState >= 1 && v.duration > 0) }; })()');
        if (st && st.wall) return 'signin';
        if (st && st.ok) { await k.js('(() => { const v = document.querySelector("video"); v.muted = true; v.pause(); return true; })()'); return 'ok'; }
      }
      return 'load';
    };
    const fetchAt = async (k, t, waitMs) => {
      if (covered(t)) return true;
      await k.js('(() => { const v = document.querySelector("video"); v.currentTime = ' + t.toFixed(2) + '; return true; })()');
      for (let i = 0; i < waitMs / 120 && !finished; i++) { await sleep(120); if (covered(t)) return true; }
      /* a moment of play nudges Veo's page to load it */
      await k.js('(() => { const v = document.querySelector("video"); v.play().catch(() => {}); return true; })()');
      let ok = false;
      for (let i = 0; i < 20 && !finished; i++) { await sleep(150); if (covered(t)) { ok = true; break; } }
      await k.js('(() => { const v = document.querySelector("video"); v.pause(); return true; })()');
      return ok;
    };
    overall = setTimeout(() => done({ error: 'timeout' }), 20 * 60 * 1000);
    (async () => {
      for (let i = 0; i < WORKERS; i++) { const k = makeWorker(); if (k) workers.push(k); }
      if (!workers.length) { done({ error: 'internal' }); return; }
      const st = await Promise.all(workers.map(ready));
      if (st.indexOf('signin') !== -1) { done({ error: 'signin' }); return; }
      const live = workers.filter((k, i) => st[i] === 'ok');
      if (!live.length) { done({ error: 'load' }); return; }
      /* every 24-second piece of each half, shared out: one queue, each page taking the next piece */
      const steps = []; ranges.forEach(r => { for (let t = r[0]; t < r[1]; t += 24) steps.push(t); });
      let doneN = 0;
      const runQueue = async (list, waitMs, pctFrom, pctTo) => {
        let next = 0, n = 0;
        await Promise.all(live.map(async k => {
          while (!finished && next < list.length) {
            const t = list[next++];
            await fetchAt(k, t, waitMs);
            n++;
            if (win && !win.isDestroyed()) win.webContents.send('veo:trackProgress', Math.min(99, Math.round(pctFrom + n / list.length * (pctTo - pctFrom))));
          }
        }));
      };
      await runQueue(steps, 4000, 0, 92);
      if (finished) return;
      if (!starts.length) { done({ error: 'nothing' }); return; }
      /* more tries at any piece that didn't come (up to three more rounds, each waiting longer) */
      for (let round = 1; round <= 3 && !finished; round++) {
        const left = steps.filter(t => !covered(t)); if (!left.length) break;
        await runQueue(left, 5000 + round * 3000, 92 + (round - 1) * 2, 92 + round * 2);
      }
      await sleep(800);
      if (finished) return;
      const cov = steps.filter(t => covered(t)).length / (steps.length || 1);
      const pk = veoPack(frames, ranges);
      if (!pk) { done({ error: 'nothing' }); return; }
      log('veo track coverage', Math.round(cov * 100) + '%');
      done(Object.assign({ ok: true, cov: Math.round(cov * 1000) / 1000 }, pk));
    })().catch(() => done({ error: 'internal' }));
  });
}
ipcMain.handle('veo:track', async (e, url, ranges) => {
  const m = String(url || '').match(/^https?:\/\/(?:app\.)?veo\.co\/matches\/([A-Za-z0-9_-]+)/i);
  if (!m) return { error: 'link' };
  const rs = (Array.isArray(ranges) ? ranges : []).filter(r => Array.isArray(r) && r.length === 2 && isFinite(r[0]) && isFinite(r[1]) && r[1] > r[0]).map(r => [+r[0], +r[1]]);
  if (!rs.length) return { error: 'nothing' };
  if (veoTrackJob) return { error: 'busy' };
  veoTrackJob = veoTrackRun(m[1], rs);
  try { return await veoTrackJob; } finally { veoTrackJob = null; }
});
ipcMain.handle('veo:signOut', async () => {
  try { await session.fromPartition(VEO_PART).clearStorageData(); return { ok: true }; } catch (e) { return { error: e.message }; }
});

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
  /* the planner closed: close everything else too (the online link, and Veo if it's still reading a match), so the program really exits */
  win.on('closed', () => {
    win = null; quitting = true;
    if (onlineWin && !onlineWin.isDestroyed()) onlineWin.destroy();
    BrowserWindow.getAllWindows().forEach(w => { try { if (!w.isDestroyed()) w.destroy(); } catch (e) {} });
    setTimeout(() => app.quit(), 200);
  });
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
  else if (store) { try { createWindow(); } catch (err) { log('reopen', err.message); } }   /* (still running without its window: open it again) */
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
