/*
 * Hubungan ke server: Web App Google Apps Script (doPost, lihat backend/src/ApiBridge.gs).
 * Database tetap Google Sheets. Alamat server TIDAK diisi pengguna: ditanam saat build
 * (app/server.json atau variabel GAS_URL) dan diperbarui otomatis dari berkas yang sama di repo,
 * jadi HP yang sudah terpasang ikut pindah bila alamat Web App berganti.
 */
const CFG = Object.assign({ version: 'dev', updateRepo: '', nativeBase: '', build: 0, serverUrl: '', demoSaja: false, xlsxUrl: '' }, window.CT_CONFIG || {});
const CAP = window.Capacitor || null;
const NATIVE = !!(CAP && CAP.isNativePlatform && CAP.isNativePlatform());
const PL = (CAP && CAP.Plugins) || {};

function galat(jenis, pesan) { const e = new Error(pesan); e.jenis = jenis; return e; }
function galatJaringan(e) {
  const m = String((e && e.message) || e || '');
  if (/timeout|timed out/i.test(m)) return galat('lambat', 'Sambungan terlalu lama. Coba lagi.');
  if (!navigator.onLine || /failed to fetch|network|unable to resolve|connect|ssl|host|load failed/i.test(m)) return galat('offline', 'Tidak ada koneksi internet. Periksa jaringan lalu coba lagi.');
  return galat('jaringan', 'Data belum bisa diambil' + (m ? ' (' + m + ')' : '') + '. Coba lagi; bila tetap gagal, hubungi admin.');
}

/* ---------- HTTP (CapacitorHttp di Android: tanpa CORS; fetch di browser) ---------- */
async function httpPost(url, body, batasMs) {
  const headers = { 'Content-Type': 'text/plain;charset=utf-8' }; // text/plain: tanpa preflight CORS di browser
  if (NATIVE && PL.CapacitorHttp) {
    const r = await PL.CapacitorHttp.request({ method: 'POST', url, headers, data: body, connectTimeout: 20000, readTimeout: batasMs });
    if (r.status < 200 || r.status >= 300) throw new Error('HTTP ' + r.status);
    return r.data;
  }
  const ctl = new AbortController(), t = setTimeout(() => ctl.abort(), batasMs);
  try {
    const res = await fetch(url, { method: 'POST', headers, body, redirect: 'follow', signal: ctl.signal });
    if (!res.ok) throw new Error('HTTP ' + res.status);
    return await res.text();
  } catch (e) { throw (e && e.name === 'AbortError') ? new Error('timeout') : e; }
  finally { clearTimeout(t); }
}
async function httpJson(url, batasMs) {
  if (NATIVE && PL.CapacitorHttp) {
    const r = await PL.CapacitorHttp.get({ url, headers: { Accept: 'application/json' }, connectTimeout: 15000, readTimeout: batasMs || 20000 });
    if (r.status !== 200) throw new Error('HTTP ' + r.status);
    return typeof r.data === 'string' ? JSON.parse(r.data) : r.data;
  }
  const ctl = new AbortController(), t = setTimeout(() => ctl.abort(), batasMs || 20000);
  try { const res = await fetch(url, { cache: 'no-store', signal: ctl.signal }); if (!res.ok) throw new Error('HTTP ' + res.status); return await res.json(); }
  finally { clearTimeout(t); }
}

/* ---------- alamat server ---------- */
const URL_SERVER = /^https:\/\/script\.google\.com\/(?:a\/macros\/[^/]+|macros)\/s\/[A-Za-z0-9_-]+\/exec$/;
const Server = {
  _janji: null,
  url() {
    if (CFG.version === 'dev') { const d = LS.get('ct.server.dev'); if (d && URL_SERVER.test(d)) return d; } // hanya build pengembangan
    const r = LS.get('ct.server.remote');
    if (r && URL_SERVER.test(r.url || '')) return r.url;
    if (URL_SERVER.test(CFG.serverUrl || '')) return CFG.serverUrl;
    // HP yang dulu diatur lewat versi 1.x tetap tersambung setelah diperbarui (tidak ada lagi layar pengaturannya).
    const lama = String(LS.get('ct.server') || '').trim().replace(/[?#].*$/, '');
    return URL_SERVER.test(lama) ? lama : '';
  },
  // Ambil alamat terbaru dari repo. Dipanggil saat aplikasi dibuka dan saat server tidak bisa dihubungi.
  sinkron(paksa) {
    if (!CFG.updateRepo || CFG.demoSaja) return Promise.resolve(this.url());
    const r = LS.get('ct.server.remote');
    if (!paksa && r && Date.now() - (r.at || 0) < 30 * 60e3) return Promise.resolve(this.url());
    if (this._janji) return this._janji;
    const sumber = 'https://raw.githubusercontent.com/' + CFG.updateRepo + '/main/app/server.json';
    this._janji = httpJson(sumber, 12000).then((j) => {
      const u = String((j && j.url) || '').trim().replace(/[?#].*$/, '');
      LS.set('ct.server.remote', { url: URL_SERVER.test(u) ? u : '', at: Date.now() });
    }).catch(() => { /* offline atau repo tidak terjangkau: pakai alamat yang sudah ada */ })
      .then(() => { this._janji = null; return this.url(); });
    return this._janji;
  }
};

/* ---------- mode demo: backend asli dijalankan di dalam aplikasi dengan data contoh ---------- */
const Demo = {
  aktif: false, info: null, _mesin: null,
  async mulai() {
    if (this._mesin) { this.aktif = true; return this.info; }
    if (!window.CT_DEMO) {
      await new Promise((ok, gagal) => {
        const s = document.createElement('script');
        s.src = 'demo/demo-backend.js'; s.onload = ok; s.onerror = () => gagal(galat('demo', 'Data demo tidak bisa dimuat.'));
        document.head.appendChild(s);
      });
    }
    await tunggu(30); // beri kesempatan layar menggambar tulisan "menyiapkan" sebelum data dibuat
    this._mesin = window.CT_DEMO.buat();
    this.info = this._mesin.info; this.aktif = true;
    return this.info;
  },
  berhenti() { this.aktif = false; },
  kirim(nama, args) { return this._mesin.kirim(nama, args); }
};

/* ---------- panggilan fungsi backend ---------- */
const GALAT_TEKNIS = /\b(Exception|Error|failed|Service|Spreadsheets?|timeout|timed out|undefined|null|Cannot|invoked|quota|exceeded|not defined|is not a function|We're sorry|permission|Authorization)\b/i;
const BATAS_BIASA = 75000, BATAS_PANJANG = 390000; // impor data bisa mendekati batas eksekusi Apps Script (6 menit)
const FUNGSI_PANJANG = { importRawData: 1, tambahFacility: 1, daftarkanFacilityExisting: 1, importLokasiAktif: 1, copyLokasiDariFacility: 1 };
async function panggil(nama, ...args) {
  // undefined → null supaya jumlah & posisi argumen tetap
  const argumen = args.map((a) => (a === undefined ? null : a));
  let j;
  if (Demo.aktif) {
    j = await Demo.kirim(nama, JSON.parse(JSON.stringify(argumen)));
  } else {
    let url = Server.url();
    if (!url) url = await Server.sinkron(true);
    if (!url) throw galat('tanpa-server', 'Aplikasi belum tersambung. Hubungi admin aplikasi.');
    const body = JSON.stringify({ action: nama, args: argumen, client: 'apk', v: CFG.version });
    let raw;
    try { raw = await httpPost(url, body, FUNGSI_PANJANG[nama] ? BATAS_PANJANG : BATAS_BIASA); }
    catch (e) { const g = galatJaringan(e); if (g.jenis !== 'offline') Server.sinkron(true); throw g; }
    j = raw;
    if (typeof raw === 'string') {
      try { j = JSON.parse(raw); }
      catch (e) { Server.sinkron(true); throw galat('ditolak', /<html|<!doctype/i.test(raw) ? 'Aplikasi belum diizinkan mengambil data. Hubungi admin aplikasi (akses Web App harus "Siapa saja").' : 'Data belum bisa dibaca. Coba lagi; bila tetap gagal, hubungi admin.'); }
    }
  }
  if (j && j.ok === true) return j.result === undefined ? null : j.result;
  if (j && j.ok === false) {
    const asli = String(j.error || ''), jenis = /^Fungsi /.test(asli) ? 'fungsi-tak-ada' : /^User tidak dikenali|login ulang/i.test(asli) ? 'sesi' : /^Akses ditolak/i.test(asli) ? 'akses' : 'server';
    let pesan = asli || 'Permintaan gagal diproses. Coba lagi; bila tetap gagal, hubungi admin.';
    // Pesan yang ditulis backend berbahasa Indonesia dan boleh tampil apa adanya. Galat dari layanan Google
    // sendiri (bahasa Inggris, berisi nama layanan atau id dokumen) diganti kalimat yang bisa ditindaklanjuti.
    if (jenis === 'server' && GALAT_TEKNIS.test(asli)) { console.warn('Galat server pada ' + nama + ': ' + asli); pesan = 'Sistem sedang terganggu. Coba lagi sebentar lagi.'; }
    const g = galat(jenis, pesan); g.asli = asli;
    throw g;
  }
  throw galat('backend-lama', 'Sistem perlu diperbarui. Hubungi admin aplikasi.');
}

/* ---------- sesi ---------- */
const Sesi = {
  user: null,
  muat() { const u = LS.get('ct.sesi'); this.user = (u && u.username) ? u : null; return this.user; },
  simpan(u) { this.user = u; if (!Demo.aktif) LS.set('ct.sesi', u); },
  hapus() { this.user = null; LS.del('ct.sesi'); },
  dari(res) {
    return {
      username: res.displayName, role: res.role, akses: !!res.aksesSettingOverflow,
      facilityId: res.facilityId || '', facilityName: res.facilityName || '', facilityCode: res.facilityCode || '', facilityStatus: res.facilityStatus || ''
    };
  }
};
const peran = () => (Sesi.user ? Sesi.user.role : '');
const Boleh = {
  admin: () => peran() === 'admin' || peran() === 'developer',
  validasi: () => peran() === 'inventory' || Boleh.admin(),
  config: () => !!Sesi.user && (Sesi.user.akses || peran() === 'developer'),
  developer: () => peran() === 'developer'
};

/* ---------- simpanan data terakhir per pengguna (tampil seketika, lalu diperbarui) ---------- */
const Simpanan = {
  kunci(k) { return 'ct.data.' + (Sesi.user ? Sesi.user.username : '-') + '.' + k; },
  get(k) { return Demo.aktif ? null : LS.get(this.kunci(k)); },
  set(k, v) { if (!Demo.aktif) LS.set(this.kunci(k), v); },
  bersih() { try { Object.keys(localStorage).forEach((k) => { if (k.indexOf('ct.data.') === 0) localStorage.removeItem(k); }); } catch (e) { /* abaikan */ } }
};
