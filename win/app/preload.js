/* Runs before the planner page. Gives the page a localStorage that is really the planner folder on disk,
   and a small window.hufcDesktop helper (Stats spreadsheets, opening the folder). */
'use strict';
const { ipcRenderer, webUtils } = require('electron');
const { pathToFileURL, fileURLToPath } = require('url');

const data = ipcRenderer.sendSync('store:all') || {};
let keys = Object.keys(data);
const has = k => Object.prototype.hasOwnProperty.call(data, k);
const store = {
  getItem(k) { k = String(k); return has(k) ? data[k] : null; },
  setItem(k, v) {
    k = String(k); v = String(v);
    if (!has(k)) keys.push(k);
    if (data[k] === v) return;
    data[k] = v; ipcRenderer.send('store:set', k, v);
  },
  removeItem(k) {
    k = String(k); if (!has(k)) return;
    delete data[k]; keys = keys.filter(x => x !== k); ipcRenderer.send('store:remove', k);
  },
  key(i) { return i >= 0 && i < keys.length ? keys[i] : null; },
  get length() { return keys.length; },
  clear() { keys.slice().forEach(k => store.removeItem(k)); }
};
/* a file opened from Explorer brought in something new */
ipcRenderer.on('store:external', (e, k, v) => { if (!has(k)) keys.push(k); data[k] = v; });
/* a newer planner has arrived: the page offers to reload (see hufcPlannerUpdated in the page) */
ipcRenderer.on('planner:updated', () => { if (typeof window.hufcPlannerUpdated === 'function') window.hufcPlannerUpdated(); else location.reload(); });
ipcRenderer.on('open-item', (e, kind, id) => {
  const go = () => { if (typeof window.hufcOpenItem === 'function') window.hufcOpenItem(kind, id); };
  if (document.readyState === 'complete') go(); else window.addEventListener('load', go);
});
Object.defineProperty(window, 'localStorage', { configurable: true, enumerable: true, get() { return store; } });

const info = ipcRenderer.sendSync('store:info') || {};
/* the version these helpers were written for: if the program still running is older, it needs a restart */
const PROGRAM_VERSION = '1.6.8';
ipcRenderer.on('program:updated', () => { if (typeof window.hufcProgramUpdated === 'function') window.hufcProgramUpdated(); });
const onlineCbs = [];
ipcRenderer.on('online:state', (e, s) => onlineCbs.forEach(cb => { try { cb(s); } catch (err) {} }));
const driveCbs = [];
ipcRenderer.on('gdrive:progress', (e, job, f) => driveCbs.forEach(cb => { try { cb(job, f); } catch (err) {} }));
window.hufcDesktop = {
  version: info.version,
  /* this program draws its own title bar (the page shows the crest, status and ☰ menu button) */
  frame: info.version === PROGRAM_VERSION ? 'app' : '',
  menu(x, y) { ipcRenderer.send('menu:popup', x, y); },
  /* Settings → Check for updates: asks the club website (and the online planner, if linked) for a newer planner and program */
  checkUpdate() { return ipcRenderer.invoke('update:check'); },
  /* Google Drive for videos (main.js / gdrive.js): connect once; videos go to the coach's Drive so every device can play them */
  gdrive: {
    status() { return ipcRenderer.invoke('gdrive:status'); },
    connect(client) { return ipcRenderer.invoke('gdrive:connect', client); },
    disconnect() { return ipcRenderer.invoke('gdrive:disconnect'); },
    remove(id) { return ipcRenderer.invoke('gdrive:remove', id); },
    quota() { return ipcRenderer.invoke('gdrive:quota'); },
    upload(job, localName, bytes, fileName, mime) { return ipcRenderer.invoke('gdrive:upload', job, localName || '', bytes || null, fileName, mime); },
    uploadPath(job, filePath, fileName, mime, resumeKey) { return ipcRenderer.invoke('gdrive:uploadPath', job, filePath, fileName, mime, resumeKey || ''); },
    onProgress(cb) { if (typeof cb === 'function') driveCbs.push(cb); }
  },
  needsRestart: !!info.version && info.version !== PROGRAM_VERSION,
  restart() { ipcRenderer.send('app:relaunch'); },
  /* the online planner, linked (see main.js): the club library shared with the online planner and phones */
  online: {
    state() { return ipcRenderer.invoke('online:state'); },
    connect() { return ipcRenderer.invoke('online:connect'); },
    show() { return ipcRenderer.invoke('online:show'); },
    disconnect() { return ipcRenderer.invoke('online:disconnect'); },
    call(op, ...args) { return ipcRenderer.invoke('online:call', String(op), args); },
    onState(cb) { if (typeof cb === 'function') onlineCbs.push(cb); }
  },
  dataDir: info.dataDir,
  writeStats(list) { ipcRenderer.send('stats:write', list); },
  openFolder(sub) { ipcRenderer.send('open:data', sub || ''); },
  /* videos added with + Upload video are kept as files in the planner's Videos folder */
  /* where a downloaded video came from (the web page, and the download address) */
  videoSource(p) { return ipcRenderer.invoke('video:source', String(p || '')); },
  pathOf(file) { try { return webUtils.getPathForFile(file) || ''; } catch (e) { return ''; } },
  async saveVideo(filePath, id, title, file) {
    if (filePath) return ipcRenderer.invoke('video:save', String(filePath), id, title);
    if (!file) return { error: 'missing' };
    const bytes = new Uint8Array(await file.arrayBuffer());   /* no location on disk (e.g. dragged from somewhere): send the video itself */
    return ipcRenderer.invoke('video:saveData', bytes, file.name || 'video.mp4', id, title);
  },
  videoUrl(name) { const p = ipcRenderer.sendSync('video:path', name); return p ? pathToFileURL(p).href : null; },
  deleteVideo(name) { ipcRenderer.send('video:delete', name); },
  /* pictures on pre-game boards are kept as files in the planner's Pictures folder */
  savePicture(bytes, id, title) { return ipcRenderer.invoke('picture:saveData', bytes, id, title); },
  pictureUrl(name) { const p = ipcRenderer.sendSync('picture:path', name); return p ? pathToFileURL(p).href : null; },
  deletePicture(name) { ipcRenderer.send('picture:delete', name); },
  /* a picture from the planner folder as data (for PowerPoint) */
  openPowerPoint(bytes, name) { return ipcRenderer.invoke('ppt:open', bytes, name); },
  /* print / save as PDF: made here, opened in the computer's PDF viewer (which has the print preview) */
  printPdf(name, keep) { return ipcRenderer.invoke('print:pdf', String(name || ''), !!keep); },
  openPdf(bytes, name, keep) { return ipcRenderer.invoke('pdf:open', bytes, String(name || ''), !!keep); },
  /* Save as…, and sharing with other coaches */
  saveFileAs(bytes, name) { return ipcRenderer.invoke('file:saveAs', bytes, String(name || '')); },
  keepShared(bytes, name) { return ipcRenderer.invoke('share:keep', bytes, String(name || '')); },
  emailFile(bytes, name, subject, body) { return ipcRenderer.invoke('mail:attach', bytes, String(name || ''), String(subject || ''), String(body || '')); },
  revealFile(file) { ipcRenderer.send('file:reveal', String(file || '')); },
  startDrag(file) { ipcRenderer.send('file:drag', String(file || '')); },
  copyFile(file) { return ipcRenderer.invoke('file:copy', String(file || '')); },
  openExternal(url) { ipcRenderer.send('open:external', String(url || '')); },
  fileDataUrl(url) { try { return ipcRenderer.sendSync('file:data', fileURLToPath(url)); } catch (e) { return null; } }
};
