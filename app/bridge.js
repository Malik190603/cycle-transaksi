/*
 * Cycle Transaksi — jembatan aplikasi Android.
 * Menggantikan google.script.run: setiap panggilan "server" dikirim sebagai POST ke Web App
 * Apps Script (doPost, lihat backend/ApiBridge.gs). Database tetap Google Sheets; antrean dan
 * worker tetap berjalan di Apps Script. File ini juga menangani alamat server, sesi login,
 * tombol kembali Android, ekspor Excel, dan pembaruan aplikasi (update kilat / APK).
 */
(function () {
  'use strict';
  var CAP = window.Capacitor || null;
  var NATIVE = !!(CAP && CAP.isNativePlatform && CAP.isNativePlatform());
  var PL = (CAP && CAP.Plugins) || {};
  var CFG = window.CT_CONFIG || { version: 'dev', updateRepo: '', nativeBase: '', build: 0, serverUrl: '' };
  var LS = {
    get: function (k) { try { return JSON.parse(localStorage.getItem(k)); } catch (e) { return null; } },
    set: function (k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) {} },
    del: function (k) { try { localStorage.removeItem(k); } catch (e) {} }
  };
  function $(id) { return document.getElementById(id); }
  function esc(v) { return String(v == null ? '' : v).replace(/[&<>"']/g, function (m) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[m]; }); }
  if (NATIVE) document.documentElement.classList.add('ct-native');

  /* ------------------------------ alamat server ------------------------------ */
  var URL_RE = /^https:\/\/script\.google\.com\/(?:a\/macros\/[^/]+|macros)\/s\/[A-Za-z0-9_-]+\/exec$/;
  function cleanUrl(u) { return String(u || '').trim().replace(/[?#].*$/, ''); }
  function serverUrl() { return cleanUrl(LS.get('ct.server') || CFG.serverUrl || ''); }

  /* ------------------------------ panggilan ke server ------------------------------ */
  var TIMEOUT_MS = 390000; // impor data bisa mendekati batas eksekusi Apps Script (6 menit)
  function netError(e) {
    var m = String((e && e.message) || e || '');
    if (/timeout|timed out/i.test(m)) return new Error('Server terlalu lama merespons. Coba lagi.');
    if (!navigator.onLine || /failed to fetch|network|unable to resolve|connect|ssl|host/i.test(m)) return new Error('Tidak ada koneksi internet. Periksa jaringan lalu coba lagi.');
    return new Error(m || 'Terjadi kesalahan.');
  }
  async function post(url, body) {
    var headers = { 'Content-Type': 'text/plain;charset=utf-8' }; // text/plain: tanpa preflight CORS di browser
    if (NATIVE && PL.CapacitorHttp) {
      var r = await PL.CapacitorHttp.request({ method: 'POST', url: url, headers: headers, data: body, connectTimeout: 20000, readTimeout: TIMEOUT_MS });
      if (r.status < 200 || r.status >= 300) throw new Error('Server menjawab HTTP ' + r.status);
      return r.data;
    }
    var ctl = new AbortController(), t = setTimeout(function () { ctl.abort(); }, TIMEOUT_MS);
    try {
      var res = await fetch(url, { method: 'POST', headers: headers, body: body, redirect: 'follow', signal: ctl.signal });
      if (!res.ok) throw new Error('Server menjawab HTTP ' + res.status);
      return await res.text();
    } catch (e) { throw (e && e.name === 'AbortError') ? new Error('timeout') : e; }
    finally { clearTimeout(t); }
  }
  async function call(name, args) {
    var url = serverUrl();
    if (!url) throw new Error('Alamat server belum diatur. Ketuk "Atur server" di layar masuk.');
    // undefined → null supaya jumlah & posisi argumen tetap (sama seperti google.script.run)
    var body = JSON.stringify({ action: name, args: Array.prototype.map.call(args, function (a) { return a === undefined ? null : a; }), client: 'apk', v: CFG.version });
    var raw;
    try { raw = await post(url, body); } catch (e) { throw netError(e); }
    var j = raw;
    if (typeof raw === 'string') {
      try { j = JSON.parse(raw); }
      catch (e) { throw new Error(/<html|<!doctype/i.test(raw) ? 'Server menolak permintaan. Pastikan Web App di-deploy dengan akses "Siapa saja".' : 'Jawaban server tidak dikenali.'); }
    }
    if (j && j.ok === true) return j.result === undefined ? null : j.result;
    if (j && j.ok === false) throw new Error(j.error || 'Terjadi kesalahan di server.');
    throw new Error('Backend belum mendukung aplikasi ini. Pasang backend/ApiBridge.gs di project Apps Script lalu deploy ulang Web App.');
  }

  /* ------------------------------ pengganti google.script.run ------------------------------ */
  function runner(ok, fail, user) {
    return new Proxy({}, {
      get: function (_, name) {
        if (name === 'withSuccessHandler') return function (f) { return runner(f, fail, user); };
        if (name === 'withFailureHandler') return function (f) { return runner(ok, f, user); };
        if (name === 'withUserObject') return function (o) { return runner(ok, fail, o); };
        if (typeof name !== 'string') return undefined;
        return function () {
          call(name, arguments).then(
            function (res) { if (ok) ok(res, user); },
            function (err) { if (fail) fail(err, user); else console.error('[server] ' + name + ': ' + err.message); }
          );
        };
      }
    });
  }
  window.google = window.google || {};
  window.google.script = { run: runner(null, null, undefined), host: { close: function () {}, origin: '' } };

  /* ------------------------------ toast & pita offline ------------------------------ */
  var toastT = 0;
  function toast(msg) {
    var t = $('ctToast'); if (!t) { t = document.createElement('div'); t.id = 'ctToast'; t.className = 'ct-toast'; t.setAttribute('role', 'status'); document.body.appendChild(t); }
    t.textContent = msg; t.classList.add('show'); clearTimeout(toastT); toastT = setTimeout(function () { t.classList.remove('show'); }, 3200);
  }
  function offlineBar() {
    var b = $('ctOffline');
    if (navigator.onLine) { if (b) b.remove(); return; }
    if (!b) { b = document.createElement('div'); b.id = 'ctOffline'; b.className = 'ct-offline'; b.textContent = 'Tidak ada koneksi internet'; document.body.appendChild(b); }
  }
  window.addEventListener('online', offlineBar); window.addEventListener('offline', offlineBar);

  /* ------------------------------ layar masuk: atur server + ingat NIK ------------------------------ */
  function serverBox() {
    var login = $('screenLogin'); if (!login || $('ctServerBox')) return;
    var link = document.createElement('button');
    link.type = 'button'; link.className = 'link'; link.id = 'ctServerLink'; link.textContent = 'Atur server';
    var box = document.createElement('div');
    box.className = 'ct-box hidden'; box.id = 'ctServerBox';
    box.innerHTML = '<h4>Alamat server</h4><p>Tempel URL Web App Apps Script yang berakhiran <b>/exec</b>. Cukup diisi sekali di HP ini.</p>' +
      '<input type="url" id="ctServerInput" placeholder="https://script.google.com/macros/s/…/exec" autocomplete="off" autocapitalize="off" spellcheck="false">' +
      '<div class="ct-row"><button type="button" id="ctServerSave">Simpan &amp; uji</button><button type="button" class="secondary" id="ctServerCancel">Tutup</button></div>' +
      '<div class="ct-msg" id="ctServerMsg" aria-live="polite"></div>';
    var foot = document.createElement('div'); foot.className = 'ct-foot'; foot.id = 'ctLoginFoot'; foot.style.textAlign = 'center'; foot.style.marginTop = '14px';
    foot.textContent = 'Aplikasi v' + CFG.version;
    login.appendChild(link); login.appendChild(box); login.appendChild(foot);
    function open(show) { box.classList.toggle('hidden', !show); if (show) { $('ctServerInput').value = serverUrl(); } }
    link.addEventListener('click', function () { open(box.classList.contains('hidden')); });
    $('ctServerCancel').addEventListener('click', function () { open(false); });
    $('ctServerSave').addEventListener('click', async function () {
      var msg = $('ctServerMsg'), u = cleanUrl($('ctServerInput').value), btn = this;
      if (!URL_RE.test(u)) { msg.className = 'ct-msg err'; msg.textContent = 'URL harus berbentuk https://script.google.com/macros/s/…/exec'; return; }
      var before = LS.get('ct.server'); LS.set('ct.server', u);
      btn.disabled = true; msg.className = 'ct-msg'; msg.textContent = 'Menguji koneksi…';
      try { var v = await call('getAppVersion', []); msg.className = 'ct-msg ok'; msg.textContent = 'Terhubung. Versi server ' + v + '.'; setTimeout(function () { open(false); }, 900); }
      catch (e) { if (before) LS.set('ct.server', before); else LS.del('ct.server'); msg.className = 'ct-msg err'; msg.textContent = e.message; }
      btn.disabled = false;
    });
    if (!serverUrl()) open(true);
  }
  function hookSession() {
    if (typeof window.setupHomeScreen === 'function' && !window.setupHomeScreen.__ct) {
      var origSetup = window.setupHomeScreen;
      window.setupHomeScreen = function () {
        try { if (typeof currentUser !== 'undefined' && currentUser && currentUser.username) LS.set('ct.user', currentUser.username); } catch (e) {}
        var r = origSetup.apply(this, arguments); homeExtras(); return r;
      };
      window.setupHomeScreen.__ct = true;
    }
    if (typeof window.logout === 'function' && !window.logout.__ct) {
      var origLogout = window.logout;
      window.logout = function () { LS.del('ct.user'); return origLogout.apply(this, arguments); };
      window.logout.__ct = true;
    }
  }
  function autoLogin() {
    var saved = LS.get('ct.user');
    if (!saved || !serverUrl() || typeof window.doLogin !== 'function' || !$('loginUsername')) return;
    $('loginUsername').value = saved; window.doLogin();
  }

  /* ------------------------------ tombol kembali Android ------------------------------ */
  function visible(id) { var e = $(id); return !!e && !e.classList.contains('hidden'); }
  function onBack() {
    if (visible('screenLogin') || visible('screenHome')) { if (PL.App && PL.App.minimizeApp) PL.App.minimizeApp(); return; }
    if (typeof window.goHome_ === 'function') window.goHome_();
  }

  /* ------------------------------ ekspor Excel di Android ------------------------------ */
  // Di WebView Android, XLSX.writeFile (unduhan lewat tautan blob) tidak menghasilkan file.
  // Diganti: tulis ke folder cache aplikasi lalu buka dengan aplikasi spreadsheet di HP.
  function hookXlsx() {
    if (!NATIVE || !window.XLSX || window.XLSX.__ct || !PL.Filesystem || !PL.FileOpener) return;
    var X = window.XLSX, orig = X.writeFile;
    X.writeFile = function (wb, name, opts) {
      var file = String(name || 'export.xlsx').replace(/[\\/:*?"<>|]+/g, '_');
      var type = /\.csv$/i.test(file) ? 'csv' : 'xlsx';
      (async function () {
        try {
          var data = X.write(wb, Object.assign({}, opts || {}, { bookType: type, type: 'base64' }));
          await PL.Filesystem.writeFile({ path: file, data: data, directory: 'CACHE' });
          var uri = (await PL.Filesystem.getUri({ path: file, directory: 'CACHE' })).uri;
          await PL.FileOpener.open({ filePath: uri, contentType: type === 'csv' ? 'text/csv' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', openWithDefault: false });
        } catch (e) { toast('Tidak bisa membuka file ekspor: ' + ((e && e.message) || e)); }
      })();
    };
    X.writeFile.__orig = orig; X.__ct = true;
  }

  /* ------------------------------ pembaruan aplikasi ------------------------------ */
  var UPD = { rel: null, busy: false, file: '', checkedAt: 0 };
  function verParts(v) { return String(v || '').replace(/^v/i, '').split('.').map(function (x) { return parseInt(x, 10) || 0; }); }
  function verNewer(a, b) { var x = verParts(a), y = verParts(b); for (var i = 0; i < Math.max(x.length, y.length); i++) { if ((x[i] || 0) !== (y[i] || 0)) return (x[i] || 0) > (y[i] || 0); } return false; }
  function livePlugin() { try { return NATIVE && CAP.isPluginAvailable && CAP.isPluginAvailable('CapacitorUpdater') ? PL.CapacitorUpdater : null; } catch (e) { return null; } }
  function liveCompatible(rel) { var l = rel && rel.live; if (!l || !livePlugin() || !CFG.nativeBase || l.native !== CFG.nativeBase) return false; return (LS.get('ct.live.failed') || []).indexOf(rel.version) < 0; }
  async function getJson(url) {
    if (NATIVE && PL.CapacitorHttp) { var r = await PL.CapacitorHttp.get({ url: url, headers: { Accept: 'application/json' }, connectTimeout: 15000, readTimeout: 20000 }); if (r.status !== 200) throw new Error('HTTP ' + r.status); return typeof r.data === 'string' ? JSON.parse(r.data) : r.data; }
    var res = await fetch(url, { cache: 'no-store' }); if (!res.ok) throw new Error('HTTP ' + res.status); return res.json();
  }
  // latest.json adalah berkas rilis biasa (bukan GitHub API), jadi tidak kena batas 60 permintaan/jam
  // per alamat IP — penting karena puluhan HP gudang memakai satu jaringan WiFi.
  async function fetchLatest() {
    var base = 'https://github.com/' + CFG.updateRepo + '/releases/latest/download/';
    var j = await getJson(base + 'latest.json');
    var rel = { version: String(j.version || ''), notes: String(j.notes || ''), url: base + j.apk, size: Number(j.apkSize) || 0, live: null };
    if (j.live && /^[0-9a-f]{64}$/.test(j.live.sha256 || '')) rel.live = { native: String(j.live.native), sha256: j.live.sha256, size: Number(j.live.size) || 0, url: base + j.live.asset };
    return rel;
  }
  function notesHtml(md) {
    var s = String(md || '').replace(/\r/g, '').replace(/<!--[\s\S]*?-->/g, ''), i = s.search(/^## Catatan developer/m), u = i >= 0 ? s.slice(0, i) : s;
    return esc(u).replace(/^## (.*)$/gm, '<b>$1</b>').replace(/^[-*] (.*)$/gm, '• $1').split('\n').filter(function (l) { return l.trim() && !/Unduh file \.apk/i.test(l); }).slice(0, 14).join('<br>') || 'Perbaikan &amp; peningkatan.';
  }
  function updSet(pct, msg) {
    var p = $('ctUpdProg'), b = $('ctUpdBar'), s = $('ctUpdStatus');
    if (p) p.hidden = pct == null; if (b && pct != null) b.style.width = Math.max(3, Math.min(100, pct)) + '%'; if (s) s.innerHTML = msg || '';
  }
  function homeExtras() {
    var home = $('screenHome'); if (!home) return;
    var footer = home.querySelector('.app-footer');
    if (footer && !$('ctFoot')) {
      var f = document.createElement('div'); f.className = 'ct-foot'; f.id = 'ctFoot';
      f.innerHTML = 'Aplikasi v' + esc(CFG.version) + '<span id="ctFootDot"></span> <a href="#" id="ctUpdCheck">Periksa pembaruan</a>';
      footer.appendChild(f);
      $('ctUpdCheck').addEventListener('click', function (e) { e.preventDefault(); checkForUpdate(true); });
    }
    var dot = $('ctFootDot'); if (dot) dot.innerHTML = UPD.rel ? '<i class="ct-dot" title="Versi baru tersedia"></i>' : '';
    var old = $('ctUpdCard'); if (old) old.remove();
    var rel = UPD.rel; if (!rel) return;
    var card = document.createElement('div'); card.className = 'ct-box ct-upd'; card.id = 'ctUpdCard';
    var kilat = rel.live && liveCompatible(rel);
    card.innerHTML = '<h4>Versi ' + esc(rel.version) + ' tersedia</h4><div class="ct-notes">' + notesHtml(rel.notes) + '</div>' +
      '<div class="ct-prog" id="ctUpdProg" hidden><i id="ctUpdBar"></i></div><div class="ct-msg" id="ctUpdStatus" aria-live="polite"></div>' +
      '<button type="button" id="ctUpdGo">' + (kilat ? 'Perbarui sekarang' : 'Unduh &amp; pasang') + '</button>';
    home.insertBefore(card, home.firstChild);
    $('ctUpdGo').addEventListener('click', runUpdate);
  }
  async function checkForUpdate(manual) {
    if (!CFG.updateRepo || CFG.version === 'dev') { if (manual) toast('Pembaruan hanya tersedia di aplikasi Android'); return; }
    if (!manual && Date.now() - UPD.checkedAt < 15 * 60e3) return;
    UPD.checkedAt = Date.now();
    try {
      var rel = await fetchLatest(); LS.set('ct.upd.latest', rel);
      if (rel.version && verNewer(rel.version, CFG.version)) { var was = UPD.rel; UPD.rel = rel; homeExtras(); if (!was || was.version !== rel.version || manual) toast('Versi ' + rel.version + ' tersedia'); }
      else { UPD.rel = null; homeExtras(); if (manual) toast('Sudah versi terbaru (' + CFG.version + ')'); }
    } catch (e) {
      var c = LS.get('ct.upd.latest'); if (c && verNewer(c.version, CFG.version)) { UPD.rel = c; homeExtras(); }
      if (manual) toast('Tidak bisa memeriksa pembaruan: ' + netError(e).message);
    }
  }
  async function runUpdate() {
    var rel = UPD.rel; if (!rel || UPD.busy) return; var go = $('ctUpdGo'), h = null;
    if (rel.live && liveCompatible(rel)) {
      UPD.busy = true; if (go) go.disabled = true; var L = livePlugin();
      try {
        h = await L.addListener('download', function (e) { var p = Number(e && e.percent) || 0; updSet(p, p < 70 ? 'Mengunduh… ' + p + '%' : 'Memasang…'); });
        updSet(2, 'Mengunduh…');
        var b = await L.download({ url: rel.live.url, version: rel.version, checksum: rel.live.sha256 });
        updSet(100, 'Memuat ulang aplikasi…'); LS.set('ct.live.pending', { v: rel.version, from: CFG.version }); await L.set({ id: b.id });
      } catch (e) { LS.del('ct.live.pending'); UPD.busy = false; rel.live = null; homeExtras(); updSet(null, 'Update kilat gagal (' + esc((e && e.message) || e) + '). Ketuk lagi untuk mengunduh APK.'); }
      finally { try { h && h.remove(); } catch (x) {} }
      return;
    }
    var FT = PL.FileTransfer, FS = PL.Filesystem, FO = PL.FileOpener;
    if (!NATIVE || !FT || !FS || !FO) { window.open(rel.url, '_blank'); return; }
    if (UPD.file) return installApk();
    UPD.busy = true; if (go) { go.disabled = true; go.textContent = 'Mengunduh…'; }
    try {
      var name = 'CycleTransaksi-' + rel.version + '.apk'; try { await FS.deleteFile({ path: name, directory: 'CACHE' }); } catch (e) {}
      var uri = (await FS.getUri({ path: name, directory: 'CACHE' })).uri, path = String(uri).replace(/^file:\/\//, ''), total = rel.size || 0;
      h = await FT.addListener('progress', function (p) { var t = p.lengthComputable && p.contentLength ? p.contentLength : total; if (t) updSet(p.bytes / t * 100, 'Mengunduh ' + (p.bytes / 1048576).toFixed(1) + ' / ' + (t / 1048576).toFixed(1) + ' MB'); });
      updSet(1, 'Mengunduh…'); await FT.downloadFile({ url: rel.url, path: path, progress: true, connectTimeout: 20000, readTimeout: 60000 });
      UPD.file = uri; UPD.busy = false; if (go) { go.disabled = false; go.textContent = 'Pasang versi ' + rel.version; }
      updSet(100, 'Unduhan selesai. Pilih <b>Update</b> di layar berikutnya.'); await installApk();
    } catch (e) { UPD.busy = false; UPD.file = ''; if (go) { go.disabled = false; go.textContent = 'Coba lagi'; } updSet(null, 'Unduhan gagal: ' + esc(netError(e).message)); }
    finally { try { h && h.remove(); } catch (x) {} }
  }
  async function installApk() {
    try { await PL.FileOpener.open({ filePath: UPD.file, contentType: 'application/vnd.android.package-archive', openWithDefault: true }); updSet(100, 'Penginstal Android terbuka. Kalau diminta, izinkan <b>Instal aplikasi tidak dikenal</b> untuk Cycle Transaksi, lalu kembali dan ketuk tombol lagi.'); }
    catch (e) { updSet(100, 'Tidak bisa membuka penginstal (' + esc((e && e.message) || e) + ').'); }
  }
  // Update kilat yang gagal menampilkan aplikasi dibatalkan otomatis oleh plugin; versi itu ditandai
  // supaya berikutnya ditawarkan lewat APK.
  (function livePendingCheck() {
    var pend = LS.get('ct.live.pending'); if (!pend || !pend.v) return; LS.del('ct.live.pending'); if (pend.v === CFG.version) return;
    var f = LS.get('ct.live.failed') || []; if (f.indexOf(pend.v) < 0) f.push(pend.v); LS.set('ct.live.failed', f.slice(-10));
  })();

  /* ------------------------------ mulai ------------------------------ */
  function start() {
    hookSession(); serverBox(); hookXlsx(); offlineBar();
    var L = livePlugin(); if (L) L.notifyAppReady().catch(function () {});
    if (NATIVE && PL.App) {
      PL.App.addListener('backButton', onBack);
      PL.App.addListener('appStateChange', function (s) { if (s && s.isActive) { offlineBar(); checkForUpdate(false); } });
    }
    autoLogin();
    setTimeout(function () { checkForUpdate(false); }, 1500);
  }
  if (document.readyState !== 'loading') start(); else document.addEventListener('DOMContentLoaded', start);
  window.CT_BRIDGE = { call: call, checkForUpdate: checkForUpdate, serverUrl: serverUrl, version: CFG.version };
})();
