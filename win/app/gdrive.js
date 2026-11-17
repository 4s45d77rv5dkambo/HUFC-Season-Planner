/* =====================================================================
   Google Drive for videos (the computer program).
   A coach connects their own Google Drive once; videos they add to the planner then go into a
   "HUFC Planner clips" folder in their Drive, readable by anyone with the link, so the planner
   can play them on every device and for every coach.
   - Sign-in: Google's page opens in the normal web browser (Google doesn't allow sign-in inside apps);
     it comes back to this program on a one-off local address (127.0.0.1) — the standard way for desktop apps.
   - Only the drive.file permission: the planner can see and change only the files it puts there.
   - The Google client details come from the planner page (so they can change without a program update).
   ===================================================================== */
const http = require('http');
const crypto = require('crypto');
const fs = require('fs');
const { shell, safeStorage } = require('electron');

const SCOPES = 'openid email https://www.googleapis.com/auth/drive.file';
/* the planner's "Planner Windows" (desktop app) OAuth client. Google doesn't treat a desktop app's secret as private
   (anyone can read it from an installed program); it's kept here rather than in the planner page. */
const DESKTOP_CLIENT = { id: ['438839893256-h18istso3v56n6c5o99lage7capui74a', 'apps.googleusercontent.com'].join('.'), secret: ['GOCSPX', '4v18DOr37nKrkyLVcN8Z2S442sbf'].join('-') };
const FOLDER_NAME = 'HUFC Planner clips';
const CHUNK = 8 * 1024 * 1024;   /* (a multiple of 256 KB, as Google asks) */

module.exports = function makeDrive(opts) {
  /* opts: getConfig(), saveConfig(c), log(...) */
  let access = '', accessExp = 0;
  const log = opts.log || function () {};
  const b64url = b => b.toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
  function enc(s) {
    try { if (safeStorage.isEncryptionAvailable()) return 'e:' + safeStorage.encryptString(s).toString('base64'); } catch (e) {}
    return 'p:' + Buffer.from(s).toString('base64');
  }
  function dec(s) {
    if (!s) return '';
    try {
      if (s.startsWith('e:')) return safeStorage.decryptString(Buffer.from(s.slice(2), 'base64'));
      if (s.startsWith('p:')) return Buffer.from(s.slice(2), 'base64').toString();
    } catch (e) {}
    return '';
  }
  const conf = () => opts.getConfig().gdrive || {};
  function setConf(g) { const c = opts.getConfig(); c.gdrive = g; opts.saveConfig(c); }
  function err(msg, code) { const e = new Error(msg); e.code = code || msg; return e; }

  function status() { const g = conf(); return { connected: !!g.refresh, email: g.email || '', folder: g.folder || '' }; }

  async function tokenReq(client, params) {
    const body = new URLSearchParams(Object.assign({ client_id: client.id, client_secret: client.secret || '' }, params));
    const r = await fetch('https://oauth2.googleapis.com/token', { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body });
    const j = await r.json().catch(() => ({}));
    if (!r.ok) throw err(j.error_description || j.error || ('token ' + r.status), j.error || 'token');
    return j;
  }

  /* sign in with Google in the web browser, and keep the (encrypted) permission on this computer */
  let connecting = null;
  async function connect(client) {
    if (!client || !client.id) client = DESKTOP_CLIENT;
    if (connecting) return connecting;
    connecting = (async () => {
      const verifier = b64url(crypto.randomBytes(32));
      const challenge = b64url(crypto.createHash('sha256').update(verifier).digest());
      const state = b64url(crypto.randomBytes(16));
      const server = http.createServer();
      await new Promise(res => server.listen(0, '127.0.0.1', res));
      const redirect = 'http://127.0.0.1:' + server.address().port;
      const url = 'https://accounts.google.com/o/oauth2/v2/auth?' + new URLSearchParams({
        client_id: client.id, redirect_uri: redirect, response_type: 'code', scope: SCOPES,
        code_challenge: challenge, code_challenge_method: 'S256', state, access_type: 'offline', prompt: 'consent'
      });
      let code;
      try {
        code = await new Promise((resolve, reject) => {
          const t = setTimeout(() => reject(err('Took too long — try Connect again', 'timeout')), 5 * 60 * 1000);
          server.on('request', (req, resp) => {
            const u = new URL(req.url, redirect);
            if (u.searchParams.get('state') !== state) { resp.writeHead(404); resp.end(); return; }
            const ok = !!u.searchParams.get('code');
            resp.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
            resp.end('<!doctype html><meta charset="utf-8"><title>HUFC Season Planner</title><body style="font:16px system-ui;background:#0d0e0f;color:#fff;display:grid;place-items:center;height:100vh;margin:0"><div style="text-align:center"><div style="font-size:42px">' + (ok ? '✓' : '✕') + '</div><h2 style="color:#ffd400">' + (ok ? 'Google Drive is connected' : 'Not connected') + '</h2><p>You can close this tab and go back to the planner.</p></div>');
            clearTimeout(t);
            if (ok) resolve(u.searchParams.get('code')); else reject(err(u.searchParams.get('error') === 'access_denied' ? 'Not connected — permission wasn’t given' : 'Not connected', 'cancelled'));
          });
          shell.openExternal(url);
        });
      } finally { try { server.close(); } catch (e) {} }
      const tok = await tokenReq(client, { code, code_verifier: verifier, redirect_uri: redirect, grant_type: 'authorization_code' });
      if (!tok.refresh_token) throw err('Google didn’t give a lasting permission — try again', 'no-refresh');
      let email = '';
      try { email = JSON.parse(Buffer.from(tok.id_token.split('.')[1], 'base64').toString()).email || ''; } catch (e) {}
      setConf({ refresh: enc(tok.refresh_token), email, client: { id: client.id, secret: client.secret || '' }, folder: '' });
      access = tok.access_token; accessExp = Date.now() + (Math.max(60, +tok.expires_in || 3600) - 60) * 1000;
      return status();
    })();
    try { return await connecting; } finally { connecting = null; }
  }

  async function token() {
    if (access && Date.now() < accessExp) return access;
    const g = conf(); const rt = dec(g.refresh);
    if (!rt || !g.client) throw err('Connect Google Drive in Settings first', 'not-connected');
    try {
      const tok = await tokenReq(g.client, { refresh_token: rt, grant_type: 'refresh_token' });
      access = tok.access_token; accessExp = Date.now() + (Math.max(60, +tok.expires_in || 3600) - 60) * 1000;
      return access;
    } catch (e) {
      if (e.code === 'invalid_grant') { setConf({}); access = ''; throw err('Google Drive was disconnected — connect it again in Settings', 'not-connected'); }
      throw e;
    }
  }
  async function api(method, url, body, headers) {
    const t = await token();
    const r = await fetch(url, { method, headers: Object.assign({ Authorization: 'Bearer ' + t }, body && !(body instanceof Buffer) ? { 'Content-Type': 'application/json' } : {}, headers || {}),
      body: body == null ? undefined : (body instanceof Buffer ? body : JSON.stringify(body)) });
    return r;
  }
  async function json(r) { const j = await r.json().catch(() => ({})); if (!r.ok) throw err((j.error && j.error.message) || ('drive ' + r.status), 'drive-' + r.status); return j; }

  /* the folder the clips go in (made the first time; made again if it's been deleted) */
  async function folder() {
    const g = conf();
    if (g.folder) {
      const r = await api('GET', 'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(g.folder) + '?fields=id,trashed');
      if (r.ok) { const j = await r.json().catch(() => ({})); if (!j.trashed) return g.folder; }
    }
    const f = await json(await api('POST', 'https://www.googleapis.com/drive/v3/files?fields=id', { name: FOLDER_NAME, mimeType: 'application/vnd.google-apps.folder' }));
    try { await api('POST', 'https://www.googleapis.com/drive/v3/files/' + f.id + '/permissions', { role: 'reader', type: 'anyone' }); } catch (e) {}
    const g2 = conf(); g2.folder = f.id; setConf(g2);
    return f.id;
  }

  /* put a video file on Drive (in pieces, so big files are fine and it can say how far it's got) */
  /* a big upload that was cut short (program closed, no internet) carries on from where it got to */
  function sessionGet(key) { const s = (conf().sessions || {})[key]; return s || null; }
  function sessionSet(key, v) { const g = conf(); g.sessions = g.sessions || {}; if (v) g.sessions[key] = v; else delete g.sessions[key]; setConf(g); }
  async function sessionResume(sess, size) {
    try {
      const r = await fetch(sess.url, { method: 'PUT', headers: { 'Content-Range': 'bytes */' + size } });
      if (r.status === 308) { const rg = r.headers.get('range'); return { pos: rg ? (+rg.split('-')[1] + 1) : 0 }; }
      if (r.ok) { const j = await r.json().catch(() => ({})); return j.id ? { id: j.id } : null; }
    } catch (e) {}
    return null;
  }
  async function upload(src, name, mime, onProgress, resumeKey) {
    /* src: a file path, or a Buffer */
    const isPath = typeof src === 'string';
    const st = isPath ? fs.statSync(src) : null;
    const size = isPath ? st.size : src.length;
    if (!size) throw err('That video is empty', 'empty');
    let session = '', id = '', from = 0;
    const old = resumeKey && isPath ? sessionGet(resumeKey) : null;
    if (old && old.size === size && old.mtime === +st.mtimeMs && Date.now() - old.at < 6 * 864e5) {
      const r = await sessionResume(old, size);
      if (r && r.id) id = r.id;
      else if (r) { session = old.url; from = r.pos; }
    }
    if (!session && !id) {
      const parent = await folder();
      const start = await api('POST', 'https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable&fields=id',
        { name: name, parents: [parent], mimeType: mime || 'video/mp4' },
        { 'X-Upload-Content-Type': mime || 'video/mp4', 'X-Upload-Content-Length': String(size) });
      if (!start.ok) await json(start);
      session = start.headers.get('location');
      if (!session) throw err('Google Drive didn’t start the upload', 'no-session');
      if (resumeKey && isPath) sessionSet(resumeKey, { url: session, size, mtime: +st.mtimeMs, at: Date.now() });
    }
    const fd = isPath && !id ? fs.openSync(src, 'r') : null;
    try {
      let again = 0;
      for (let pos = id ? size : from; pos < size;) {
        const len = Math.min(CHUNK, size - pos);
        let buf;
        if (fd !== null) { buf = Buffer.alloc(len); fs.readSync(fd, buf, 0, len, pos); } else buf = src.subarray(pos, pos + len);
        let r = null;
        for (let tries = 0; tries < 4; tries++) {
          try {
            r = await fetch(session, { method: 'PUT', headers: { 'Content-Length': String(len), 'Content-Range': 'bytes ' + pos + '-' + (pos + len - 1) + '/' + size }, body: buf });
            if (r.status < 500) break;
          } catch (e) { if (tries === 3) throw e; }
          await new Promise(res => setTimeout(res, 1500 * (tries + 1)));
        }
        if (r.status === 308) {
          /* no Range: Google kept none of that piece, so it goes again from the same place (never skipped) */
          const rg = r.headers.get('range');
          if (rg) { pos = +rg.split('-')[1] + 1; again = 0; }
          else if (++again > 3) throw err('Google Drive didn’t take the video — try again', 'incomplete');
        } else if (r.ok) {
          const j = await r.json().catch(() => ({})); id = j.id || ''; pos = size;
        } else await json(r);
        if (onProgress) try { onProgress(Math.min(1, pos / size)); } catch (e) {}
      }
    } finally { if (fd !== null) try { fs.closeSync(fd); } catch (e) {} }
    if (!id) throw err('The upload didn’t finish — try again', 'incomplete');
    if (resumeKey) sessionSet(resumeKey, null);
    /* anyone with the link can watch it (the folder says the same) */
    try { await api('POST', 'https://www.googleapis.com/drive/v3/files/' + id + '/permissions', { role: 'reader', type: 'anyone' }); } catch (e) { log('drive permission', e.message); }
    return { id, url: 'https://drive.google.com/file/d/' + id + '/view' };
  }

  async function remove(id) {
    if (!id) return false;
    const r = await api('DELETE', 'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(id));
    return r.ok || r.status === 404;
  }

  async function disconnect() {
    const g = conf(); const rt = dec(g.refresh);
    setConf({}); access = ''; accessExp = 0;
    if (rt) { try { await fetch('https://oauth2.googleapis.com/revoke?token=' + encodeURIComponent(rt), { method: 'POST' }); } catch (e) {} }
    return status();
  }

  /* how full this Google Drive is (the whole account: Drive, Gmail and Photos share it) */
  async function quota() {
    const j = await json(await api('GET', 'https://www.googleapis.com/drive/v3/about?fields=storageQuota'));
    const q = j.storageQuota || {};
    return { limit: q.limit != null ? +q.limit : null, usage: +q.usage || 0, inDrive: +q.usageInDrive || 0, inTrash: +q.usageInDriveTrash || 0 };
  }

  return { status, connect, upload, remove, disconnect, quota };
};
