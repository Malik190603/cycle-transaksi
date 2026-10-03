// Uji tampilan dari ujung ke ujung: aplikasi (www/) dijalankan di Chromium dan berbicara dengan backend
// Apps Script yang ASLI (backend/src, dijalankan di Node dengan layanan Google tiruan). Tidak ada yang
// ditiru di sisi aplikasi: alamat server diambil lewat app/server.json persis seperti di HP.
//
// Jalankan: npm run build && node test/ui.mjs      (butuh paket playwright di mesin ini)
import http from 'node:http';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';
import { createDemoBackend, createBackend } from './harness.mjs';

const require = createRequire(import.meta.url);
let chromium;
try { ({ chromium } = require('playwright')); } catch { ({ chromium } = require('/opt/npm-tools/node_modules/playwright')); }
const XLSX = require('xlsx');

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..'), WWW = path.join(ROOT, 'www');
const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'ct-ui-'));
const REPO = JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8')).updateRepo;
const EXEC = 'https://script.google.com/macros/s/AKfycbTEST_deploy-123/exec';
const MIME = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css', '.woff2': 'font/woff2' };

if (!fs.existsSync(path.join(WWW, 'app.js'))) { console.error('www/ belum ada. Jalankan: npm run build'); process.exit(1); }
// Panel admin: halaman yang disajikan doGet dari backend/Code.gs (file siap tempel), persis seperti di Apps Script.
const PANEL = createBackend({ single: true }).ctx.doGet({}).getContent();
const server = http.createServer((q, r) => {
  if (q.url === '/panel') { r.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' }); return r.end(PANEL); }
  const f = path.join(WWW, q.url === '/' ? 'index.html' : q.url.split('?')[0]);
  if (!f.startsWith(WWW) || !fs.existsSync(f)) { r.writeHead(404); return r.end(); }
  r.writeHead(200, { 'Content-Type': MIME[path.extname(f)] || 'application/octet-stream' }); r.end(fs.readFileSync(f));
}).listen(0);
const ORIGIN = 'http://127.0.0.1:' + server.address().port;

let gagal = 0, lulus = 0;
const cek = (nama, syarat, info) => { if (syarat) lulus++; else gagal++; console.log((syarat ? '  ✓ ' : '  ✗ ') + nama + (syarat || info === undefined ? '' : '  → ' + JSON.stringify(info))); };
const browser = await chromium.launch();
const semuaGalat = [];

/**
 * Buka aplikasi di konteks baru.
 *   be        backend yang melayani /exec (null = tidak ada server)
 *   remote    isi app/server.json di repo ('' = belum diisi)
 *   conf      timpa window.CT_CONFIG (mis. { version: '2.0.0' })
 *   latest    isi latest.json rilis terbaru (null = 404)
 *   ubah      (req) → jawaban pengganti untuk satu panggilan (uji backend lama)
 * Hasilnya punya `atur`: { ubah, tunda(req), tundaRilis } yang boleh diganti selagi uji berjalan, untuk meniru
 * server yang lambat atau menolak. `tunda` dan `tundaRilis` boleh berupa milidetik atau janji dari gerbang():
 * jawaban ditahan sampai uji membukanya, jadi urutan kejadian tidak bergantung pada kecepatan mesin.
 */
async function buka({ be = null, remote = EXEC, conf = null, latest = null, ubah = null, w = 393, h = 852, penyimpanan = null, gelap = false, jalur = '/' } = {}) {
  const ctx = await browser.newContext({ viewport: { width: w, height: h }, deviceScaleFactor: 1, hasTouch: true, isMobile: true, colorScheme: gelap ? 'dark' : 'light' });
  const panggilan = [], st = { offline: false, ubah, tunda: null, tundaRilis: 0 };
  await ctx.route('https://raw.githubusercontent.com/**', (r) => r.fulfill({ status: 200, contentType: 'application/json', headers: { 'Access-Control-Allow-Origin': '*' }, body: JSON.stringify({ url: remote }) }));
  await ctx.route('https://github.com/**', async (r) => {
    st.rilisDiminta = (st.rilisDiminta || 0) + 1;
    if (st.tundaRilis) await (typeof st.tundaRilis === 'number' ? tunggu(st.tundaRilis) : st.tundaRilis);
    const jawab = latest ? { status: 200, contentType: 'application/json', headers: { 'Access-Control-Allow-Origin': '*' }, body: JSON.stringify(latest) } : { status: 404, headers: { 'Access-Control-Allow-Origin': '*' }, body: '' };
    r.fulfill(jawab).catch(() => {});
  });
  await ctx.route('https://script.google.com/**', async (r) => {
    const q = r.request();
    if (st.offline) return r.abort('internetdisconnected');
    if (!be || q.method() !== 'POST') return r.fulfill({ status: 404, body: '' });
    const req = JSON.parse(q.postData()); panggilan.push(req);
    const lama = st.tunda ? st.tunda(req) : 0;
    if (lama) await (typeof lama === 'number' ? tunggu(lama) : lama);
    const ganti = st.ubah ? st.ubah(req) : undefined;
    // jawaban dihitung saat dikirim (bukan saat diminta), seperti server sungguhan yang lambat
    r.fulfill({ status: 200, contentType: 'application/json', headers: { 'Access-Control-Allow-Origin': '*' }, body: JSON.stringify(ganti !== undefined ? ganti : be.post(req)) }).catch(() => {});
  });
  if (conf) {
    const asli = fs.readFileSync(path.join(WWW, 'config.js'), 'utf8');
    const dasar = JSON.parse(asli.slice(asli.indexOf('{'), asli.lastIndexOf('}') + 1));
    await ctx.route('**/config.js', (r) => r.fulfill({ status: 200, contentType: 'text/javascript', body: 'window.CT_CONFIG = ' + JSON.stringify(Object.assign(dasar, conf)) + ';' }));
  }
  if (penyimpanan) await ctx.addInitScript((isi) => { if (!sessionStorage.getItem('ct-uji-awal')) { sessionStorage.setItem('ct-uji-awal', '1'); Object.keys(isi).forEach((k) => localStorage.setItem(k, isi[k])); } }, penyimpanan);
  const page = await ctx.newPage();
  const galat = [];
  page.on('pageerror', (e) => galat.push('pageerror: ' + e.message));
  page.on('console', (m) => { if (m.type() === 'error' && !/ERR_INTERNET_DISCONNECTED|Failed to load resource/.test(m.text())) galat.push('console: ' + m.text()); });
  if (jalur === '/panel') {
    // di Apps Script halaman ini memanggil backend lewat google.script.run; di sini diteruskan ke backend uji
    await ctx.route('https://fonts.googleapis.com/**', (r) => r.fulfill({ status: 200, contentType: 'text/css', body: '' }));
    await ctx.route('https://cdnjs.cloudflare.com/**', (r) => r.fulfill({ status: 200, contentType: 'text/javascript', body: fs.readFileSync(path.join(ROOT, 'node_modules/xlsx/dist/xlsx.full.min.js')) }));
    await ctx.addInitScript((exec) => {
      const buat = (ok, gagal) => ({
        withSuccessHandler: (f) => buat(f, gagal), withFailureHandler: (g) => buat(ok, g),
        panelApi: (aksi, argsJson) => { fetch(exec, { method: 'POST', body: JSON.stringify({ action: aksi, args: JSON.parse(argsJson) }) }).then((x) => x.text()).then(ok, gagal); }
      });
      window.google = { script: { run: buat(() => {}, () => {}) } };
    }, EXEC);
  }
  await page.goto(ORIGIN + jalur);
  await page.waitForSelector('.scr.is-on');
  // setOffline Playwright tidak memutus permintaan yang dijawab lewat route, jadi diputus di sini juga
  const offline = async (v) => { st.offline = v; await ctx.setOffline(v); };
  return { ctx, page, galat, panggilan, offline, atur: st, tutup: async () => { galat.forEach((g) => semuaGalat.push(g)); await ctx.close(); } };
}
const masuk = async (page, nik) => {
  await page.waitForFunction(() => /^Tersambung/.test(document.getElementById('lgStatus').textContent.trim()), null, { timeout: 15000 });
  await page.fill('#lgNik', nik); await page.click('#lgTombol');
  await page.waitForSelector('#scr-home.is-on .kpi, #scr-home.is-on .hm-tugas, #scr-home.is-on .kosong', { timeout: 15000 });
};
const teks = (page, sel) => page.$eval(sel, (n) => n.textContent.replace(/\s+/g, ' ').trim());
const angka = (s) => Number(String(s).replace(/[^\d-]/g, ''));
const kembali = (page) => page.evaluate(() => window.CT.Nav.kembali());
const antreKosong = (page) => page.waitForFunction(() => window.CT.Antrean.daftar.length === 0, null, { timeout: 15000 });
const tunggu = (ms) => new Promise((r) => setTimeout(r, ms));
// gerbang: jawaban server ditahan sampai uji memanggil buka()
const gerbang = () => { let buka; const janji = new Promise((r) => { buka = r; }); return { janji, buka }; };
// tunggu sampai syarat di sisi uji (bukan di halaman) terpenuhi
const sampai = async (syarat, maks = 15000) => { const t0 = Date.now(); while (!syarat()) { if (Date.now() - t0 > maks) throw new Error('sampai(): syarat tidak terpenuhi dalam ' + maks + ' ms'); await tunggu(40); } };
// HANYA=3b,7 node test/ui.mjs  → menjalankan bagian tertentu saja
const HANYA = (process.env.HANYA || '').split(',').map((x) => x.trim()).filter(Boolean);
async function bagian(nama, fn) {
  if (HANYA.length && HANYA.indexOf(nama.split('.')[0]) < 0) return;
  console.log('\n' + nama);
  try { await fn(); } catch (e) { gagal++; console.log('  ✗ berhenti: ' + String(e && e.message).split('\n')[0]); }
}

const be = createDemoBackend({ items: 260, historyDays: 2 });
const { admin: ADMIN, inventory: INV, petugas: PETUGAS, tertinggal: PETUGAS2 } = be.info;
const dataCount = () => be.rows('Data Count');

/* ------------------------------------------------------------------ */
await bagian('1. Layar masuk tanpa server: tidak ada pengaturan server', async () => {
  const s = await buka({ remote: '', conf: { serverUrl: '' } }); // build ini bisa sudah membawa alamat dari app/server.json
  await s.page.waitForFunction(() => /belum tersambung/.test(document.getElementById('lgStatus').textContent));
  cek('status menyebut aplikasi belum tersambung', true);
  const isiLayar = await teks(s.page, '#scr-login');
  cek('tidak ada "Atur server" atau kolom alamat server', !/atur server|alamat server|\/exec/i.test(isiLayar) && (await s.page.$$('#scr-login input')).length === 1, isiLayar.slice(0, 200));
  await s.page.fill('#lgNik', ADMIN); await s.page.click('#lgTombol');
  await s.page.waitForFunction(() => document.getElementById('lgPesan').textContent.length > 0);
  cek('masuk ditolak dengan pesan yang jelas', /belum tersambung/i.test(await teks(s.page, '#lgPesan')), await teks(s.page, '#lgPesan'));
  cek('tidak ada kata "server" di layar masuk', !/server/i.test(await teks(s.page, '#scr-login')), await teks(s.page, '#scr-login'));
  cek('tawaran mode demo tersedia', !!(await s.page.$('[data-aksi="demo-mulai"]')));
  // mode demo: backend asli berjalan di dalam aplikasi
  await s.page.click('[data-aksi="demo-mulai"]');
  await s.page.waitForSelector('[data-aksi="demo-masuk"][data-sebagai="admin"]', { timeout: 30000 });
  await s.page.click('[data-aksi="demo-masuk"][data-sebagai="admin"]');
  await s.page.waitForSelector('#scr-home.is-on .kpi', { timeout: 30000 });
  cek('mode demo: Home admin tampil dengan pita demo', /Mode demo/.test(await teks(s.page, '#pita')) && angka(await teks(s.page, '.kpi--biru .kpi__nilai')) > 0);
  cek('mode demo tidak menyimpan sesi', (await s.page.evaluate(() => localStorage.getItem('ct.sesi'))) === null);
  await s.page.click('[data-aksi="akun"]'); await s.page.click('[data-aksi="akun-keluar"]');
  await s.page.waitForSelector('#scr-login.is-on');
  cek('keluar dari demo kembali ke layar masuk tanpa pita', !(await s.page.$('#pita')));
  await s.tutup();
});

/* ------------------------------------------------------------------ */
await bagian('2. Masuk dengan NIK, Home admin sesuai data server', async () => {
  const s = await buka({ be });
  await s.page.waitForFunction(() => /^Tersambung/.test(document.getElementById('lgStatus').textContent.trim()));
  cek('alamat server diambil dari app/server.json di repo', (await s.page.evaluate(() => window.CT.Server.url())) === EXEC);
  await s.page.fill('#lgNik', '000000.tidakada'); await s.page.click('#lgTombol');
  await s.page.waitForFunction(() => document.getElementById('lgPesan').textContent.length > 0);
  cek('NIK tidak terdaftar ditolak', /belum terdaftar/.test(await teks(s.page, '#lgPesan')));
  await masuk(s.page, ADMIN.toLowerCase());
  const b = be.api('getHomeBundle', ADMIN);
  cek('nama dan facility tampil', (await teks(s.page, '.hm-nama')) === ADMIN && /RDC Tallo Makassar/.test(await teks(s.page, '.hm-tempat')));
  const kpi = await s.page.$$eval('.kpi__nilai', (n) => n.map((x) => x.textContent.replace(/\D/g, '')));
  cek('empat KPI sama dengan server', JSON.stringify(kpi.map(Number)) === JSON.stringify([b.totalCycleHariIni, b.outstanding, b.belumValidasi, b.selesaiHariIni]), { kpi, b: [b.totalCycleHariIni, b.outstanding, b.belumValidasi, b.selesaiHariIni] });
  cek('enam menu bawah untuk admin', (await s.page.$$eval('.nav__item span:not(.nav__lencana)', (n) => n.map((x) => x.textContent))).join(',') === 'Home,Cycle,Validasi,Verifikasi,Upload,Report');
  cek('Home dimuat dengan satu panggilan', s.panggilan.filter((p) => p.action === 'getHomeBundle').length === 1 && !s.panggilan.some((p) => p.action === 'getHomeSummary'));
  const persenTeks = await teks(s.page, '.hm-dua .hm-persen');
  cek('persen progres tidak dibulatkan menjadi 100% selama masih ada sisa', b.outstanding > 0 ? persenTeks !== '100%' : true, persenTeks);
  await s.page.click('.kpi--jingga');
  await s.page.waitForSelector('.lembar .orang');
  const sisaLembar = await s.page.$$eval('.lembar .orang__atas > span', (n) => n.reduce((a, x) => a + Number(x.textContent.replace(/\D/g, '')), 0));
  cek('rincian outstanding per petugas berjumlah sama dengan KPI', sisaLembar === b.outstanding, { sisaLembar, outstanding: b.outstanding });
  await kembali(s.page); await s.page.waitForSelector('.lembar-wadah', { state: 'detached' });
  await s.page.click('.pm--minus'); await s.page.waitForSelector('.lembar .pm-lokasi');
  cek('Summary plus minus bisa dirinci per SKU', (await s.page.$$('.lembar .list .row')).length === b.plusMinus.minus.totalSku);
  await kembali(s.page);
  // sesi tersimpan: buka ulang langsung ke Home
  await s.page.reload(); await s.page.waitForSelector('#scr-home.is-on .kpi');
  cek('sesi tersimpan, buka ulang langsung ke Home', (await teks(s.page, '.hm-nama')) === ADMIN);
  await s.tutup();
});

/* ------------------------------------------------------------------ */
await bagian('3. Petugas: cycle di mode fokus, urungkan, antrean kirim', async () => {
  const s = await buka({ be });
  await masuk(s.page, PETUGAS);
  const awal = be.api('getMyPendingTasks', PETUGAS);
  cek('menu bawah petugas hanya Home dan Cycle', (await s.page.$$('.nav__item')).length === 2);
  cek('Home petugas menampilkan sisa tugas sendiri', angka(await teks(s.page, '.hm-sisa b')) === awal.length, { tampil: await teks(s.page, '.hm-sisa'), server: awal.length });
  await s.page.click('[data-aksi="home-mulai-cycle"]');
  await s.page.waitForSelector('#fkPad');
  const pertama = await s.page.$eval('.rak-besar', (n) => n.getAttribute('aria-label').replace('Lokasi ', ''));
  const sku1 = await teks(s.page, '.fk__sku');
  const t1 = awal.find((t) => t.lokasi === pertama && String(t.article) === sku1);
  cek('lokasi pertama berasal dari daftar tugas server', !!t1, { pertama, sku1 });
  cek('hitungan buta: qty sistem tidak ada di layar', !new RegExp('(^|\\D)' + 'Qty sistem', 'i').test(await teks(s.page, '#scr-hitung')) && !(await s.page.$('#scr-hitung [data-sys]')));
  cek('tombol Simpan mati sebelum qty diisi', await s.page.$eval('#fkSimpan', (b) => b.disabled));
  const qtySalah = t1.qtySystem + 3;
  for (const d of String(qtySalah)) await s.page.click('[data-k="' + d + '"]');
  cek('keypad mengisi qty', angka(await teks(s.page, '#fkAngka')) === qtySalah);
  await s.page.click('#fkSimpan');
  await s.page.waitForSelector('.simpanan');
  cek('hasil tersimpan tampil dengan tombol Urungkan', new RegExp(pertama.replace(/\./g, '\\.') + ' tersimpan, qty ' + qtySalah).test(await teks(s.page, '.simpanan')));
  cek('selama bisa diurungkan, hasil belum dikirim ke server', !s.panggilan.some((p) => p.action === 'submitCount'));
  await s.page.click('[data-aksi="fokus-urung"]');
  await s.page.waitForFunction((l) => document.querySelector('.rak-besar').getAttribute('aria-label') === 'Lokasi ' + l, pertama);
  cek('Urungkan mengembalikan lokasi dan qty yang tadi', angka(await teks(s.page, '#fkAngka')) === qtySalah && !s.panggilan.some((p) => p.action === 'submitCount'));
  // perbaiki: hapus lalu isi qty yang cocok dengan stok
  for (let i = 0; i < String(qtySalah).length; i++) await s.page.click('[data-k="hapus"]');
  for (const d of String(t1.qtySystem)) await s.page.click('[data-k="' + d + '"]');
  await s.page.click('#fkSimpan');
  // item kedua: selisih (masuk validasi)
  await s.page.waitForFunction((l) => document.querySelector('.rak-besar').getAttribute('aria-label') !== 'Lokasi ' + l || document.querySelector('.fk__sku').textContent !== '', pertama);
  const kedua = await s.page.$eval('.rak-besar', (n) => n.getAttribute('aria-label').replace('Lokasi ', '')), sku2 = await teks(s.page, '.fk__sku');
  const t2 = awal.find((t) => t.lokasi === kedua && String(t.article) === sku2 && t.no !== t1.no);
  for (const d of String(t2.qtySystem + 1)) await s.page.click('[data-k="' + d + '"]');
  await s.page.click('#fkSimpan');
  await s.page.waitForFunction(() => window.CT.Antrean.daftar.length <= 1); // simpan berikutnya melepas hasil sebelumnya
  cek('hasil pertama dikirim begitu hasil berikutnya disimpan', s.panggilan.filter((p) => p.action === 'submitCount').length === 1 && s.panggilan.find((p) => p.action === 'submitCount').args[0] === t1.no);
  const kirim1 = s.panggilan.find((p) => p.action === 'submitCount').args;
  cek('hasil membawa jam hitung dari HP dan facility tempat hasil itu dibuat', Math.abs(kirim1[3] - Date.now()) < 120000 && kirim1[4] === be.info.facilityId, kirim1);
  await kembali(s.page); // keluar mode fokus: hasil yang masih ditahan ikut dikirim
  await antreKosong(s.page);
  be.worker();
  const baris = (no) => dataCount().find((r) => r[0] === no);
  cek('server mencatat hasil cocok', baris(t1.no)[10] === 'Selesai' && Number(baris(t1.no)[11]) === t1.qtySystem && baris(t1.no)[13] === 'HIT', baris(t1.no));
  cek('server mencatat hasil selisih', baris(t2.no)[10] === 'Selesai' && Number(baris(t2.no)[11]) === t2.qtySystem + 1 && baris(t2.no)[13] === 'DISCREPANCY', baris(t2.no));
  await s.page.waitForSelector('#scr-cycle.is-on .list');
  cek('daftar dan lencana berkurang dua', angka(await teks(s.page, '#scr-cycle .hd__ket')) === awal.length - 2 && angka(await teks(s.page, '#tab-cycle .nav__lencana')) === awal.length - 2, await teks(s.page, '#scr-cycle .hd__ket'));
  // riwayat sesi: hasil baru terlihat setelah terkirim
  await s.page.click('[data-aksi="tugas-mulai"]'); await s.page.waitForSelector('#fkPad');
  const ketiga = await s.page.$eval('.rak-besar', (n) => n.getAttribute('aria-label').replace('Lokasi ', '')), sku3 = await teks(s.page, '.fk__sku');
  const t3 = awal.find((t) => t.lokasi === ketiga && String(t.article) === sku3 && t.no !== t1.no && t.no !== t2.no);

  // ---- tanpa sinyal: hasil tetap tersimpan di HP dan terkirim saat tersambung lagi
  await s.offline(true);
  await s.page.waitForSelector('#pita');
  cek('pita "tidak ada internet" muncul dan menenangkan: hasil tetap tersimpan', /Tidak ada internet\. Hasil tetap tersimpan di HP/.test(await teks(s.page, '#pita')), await teks(s.page, '#pita'));
  for (const d of String(t3.qtySystem)) await s.page.click('[data-k="' + d + '"]');
  await s.page.click('#fkSimpan');
  await s.page.evaluate(() => window.CT.Antrean.lepas());
  await s.page.waitForFunction(() => window.CT.Antrean.daftar.length === 1 && window.CT.Antrean.daftar[0].coba >= 1, null, { timeout: 15000 });
  cek('offline: hasil menunggu di antrean HP, tidak hilang', (await s.page.evaluate(() => JSON.parse(localStorage.getItem('ct.antrean')).length)) === 1);
  cek('offline: penanda antrean tampil di kepala mode fokus', (await s.page.$eval('#fkAntre', (b) => b.classList.contains('is-tunggu'))) === true);
  await s.ctx.setOffline(false); // berkas aplikasi di uji ini disajikan lewat HTTP; server /exec tetap terputus
  await s.page.reload(); await s.page.waitForSelector('#scr-home.is-on');
  await tunggu(800);
  cek('aplikasi dibuka ulang saat server tidak terjangkau: antrean masih ada', (await s.page.evaluate(() => window.CT.Antrean.daftar.length)) === 1);
  cek('Home tetap tampil dari data terakhir', angka(await teks(s.page, '.hm-sisa b')) > 0);
  await s.offline(false);
  await s.page.evaluate(() => window.dispatchEvent(new Event('online')));
  await antreKosong(s.page);
  be.worker();
  cek('tersambung lagi: hasil terkirim dan tercatat', baris(t3.no)[10] === 'Selesai' && Number(baris(t3.no)[11]) === t3.qtySystem);
  cek('tidak ada kiriman ganda', be.rows('Riwayat').filter((r) => r[2] === t3.lokasi && String(r[3]) === String(t3.article) && r[7] === PETUGAS && String(r[1]).slice(0, 10) === be.info.today).length === 1);
  await s.tutup();
});

/* ------------------------------------------------------------------ */
await bagian('3b. Petugas: server sibuk dicoba ulang, penolakan mengembalikan tugas, item terakhir bisa diurungkan', async () => {
  // petugas dengan tugas terbanyak: kolom cari baru muncul bila tugasnya lebih dari delapan
  const s = await buka({ be });
  await masuk(s.page, PETUGAS2);
  await s.page.click('#tab-cycle'); await s.page.waitForSelector('#scr-cycle .list');
  const awal = be.api('getMyPendingTasks', PETUGAS2);
  const baris = (no) => dataCount().find((r) => r[0] === no);
  const lokasiLayar = () => s.page.$eval('.rak-besar', (n) => n.getAttribute('aria-label').replace('Lokasi ', ''));
  const itemLayar = async () => { const lok = await lokasiLayar(), sku = await teks(s.page, '.fk__sku'); return awal.find((t) => t.lokasi === lok && String(t.article) === sku); };
  const ketik = async (n) => { for (const d of String(n)) await s.page.click('[data-k="' + d + '"]'); };
  const toast = () => s.page.evaluate(() => { const n = document.querySelector('.toast.is-on'); return n ? n.textContent : ''; });
  await s.page.click('[data-aksi="tugas-mulai"]'); await s.page.waitForSelector('#fkPad');

  // ---- server sedang sibuk (pesan asli backend saat gagal mengunci antrean): hasil tidak boleh dibuang
  const a = await itemLayar();
  let sibuk = 0;
  s.atur.ubah = (req) => (req.action === 'submitCount' && req.args[0] === a.no && sibuk++ < 2 ? { ok: true, result: { success: false, message: 'Gagal masuk antrian submit: Lock timeout: another process was holding the lock for too long.' } } : undefined);
  await ketik(a.qtySystem); await s.page.click('#fkSimpan');
  await s.page.evaluate(() => window.CT.Antrean.lepas());
  await sampai(() => sibuk >= 1); // server sudah menjawab "sibuk" sekali
  await s.page.waitForFunction(() => window.CT.Antrean._sibuk === false, null, { timeout: 15000 });
  const antre1 = await s.page.evaluate(() => window.CT.Antrean.daftar.map((e) => ({ status: e.status, cobaServer: e.cobaServer || 0 })));
  cek('server sibuk: hasil tetap di antrean HP untuk dicoba lagi', antre1.length === 1 && antre1[0].status !== 'tahan' && antre1[0].cobaServer >= 1, antre1);
  cek('server sibuk: petugas tidak diberi tahu "belum tersimpan"', !/belum tersimpan/.test(await toast()), await toast());
  // dorong percobaan berikutnya sampai terkirim (percobaan otomatis menunggu beberapa detik)
  for (let i = 0; i < 40 && (await s.page.evaluate(() => window.CT.Antrean.daftar.length)) > 0; i++) { await s.page.evaluate(() => window.CT.Antrean.cobaLagi()); await tunggu(150); }
  await antreKosong(s.page);
  be.worker();
  cek('server pulih: hasil yang sama terkirim dan tercatat', baris(a.no)[10] === 'Selesai' && Number(baris(a.no)[11]) === a.qtySystem && s.panggilan.filter((p) => p.action === 'submitCount' && p.args[0] === a.no).length === 3, baris(a.no));

  // ---- server terganggu lama dan petugas berkali-kali menekan "kirim sekarang": hasil tidak boleh habis dicoba dalam hitungan detik
  const a2 = await itemLayar();
  let gagalTerus = 0;
  s.atur.ubah = (req) => (req.action === 'submitCount' && req.args[0] === a2.no ? (gagalTerus++, { ok: false, error: 'Service Spreadsheets failed while accessing document with id 1AbC.' }) : undefined);
  await ketik(a2.qtySystem); await s.page.click('#fkSimpan');
  await s.page.evaluate(() => window.CT.Antrean.lepas());
  for (let i = 0; i < 14; i++) { await sampai(() => gagalTerus > i).catch(() => {}); await s.page.waitForFunction(() => window.CT.Antrean._sibuk === false); await s.page.evaluate(() => window.CT.Antrean.cobaLagi()); await tunggu(60); }
  const antre2 = await s.page.evaluate(() => window.CT.Antrean.daftar.length);
  cek('server terganggu: 14 percobaan beruntun tidak membuang hasil', gagalTerus >= 12 && antre2 === 1 && !/belum tersimpan/.test(await toast()), { gagalTerus, antre2, toast: await toast() });
  // gangguan tidak kunjung pulih (lebih dari 15 menit sejak galat pertama): tugas dikembalikan, pesannya bisa ditindaklanjuti
  await s.page.waitForFunction(() => window.CT.Antrean._sibuk === false);
  await s.page.evaluate(() => { window.CT.Antrean.daftar[0].gagalSejak = Date.now() - 16 * 60e3; window.CT.Antrean.cobaLagi(); });
  await s.page.waitForFunction(() => /belum tersimpan/.test((document.querySelector('.toast.is-on') || {}).textContent || ''), null, { timeout: 15000 });
  cek('gangguan lebih dari 15 menit: tugas dikembalikan tanpa teks teknis server', (await toast()).indexOf(a2.lokasi + ' belum tersimpan: Sistem sedang sibuk. Hitung ulang lokasi ini nanti.') >= 0 && (await s.page.evaluate((no) => window.CT.Antrean.daftar.length === 0 && window.CT.Fokus.daftar.some((it) => it.no === no), a2.no)) && baris(a2.no)[10] !== 'Selesai', await toast());
  s.atur.ubah = null;

  // ---- penolakan yang pasti: item kembali menjadi tugas, dengan alasan yang terbaca
  const b = await itemLayar();
  s.atur.ubah = (req) => (req.action === 'submitCount' && req.args[0] === b.no ? { ok: true, result: { success: false, message: 'Qty hasil hitung tidak valid.' } } : undefined);
  await ketik(b.qtySystem); await s.page.click('#fkSimpan');
  await s.page.evaluate(() => window.CT.Antrean.lepas());
  // toast pengembalian item sebelumnya bisa masih tampil: tunggu pesan untuk lokasi ini
  await s.page.waitForFunction((awal) => ((document.querySelector('.toast.is-on') || {}).textContent || '').indexOf(awal) >= 0, b.lokasi + ' belum tersimpan: Qty', { timeout: 15000 }).catch(() => {});
  cek('penolakan: pesan menyebut lokasi dan alasannya', (await toast()).indexOf(b.lokasi + ' belum tersimpan: Qty hasil hitung tidak valid.') >= 0, await toast());
  cek('penolakan: item kembali ke urutan sesi ini dan antrean kosong', (await s.page.evaluate((no) => window.CT.Antrean.daftar.length === 0 && window.CT.Fokus.daftar.some((it) => it.no === no), b.no)) && baris(b.no)[10] !== 'Selesai');
  s.atur.ubah = null;

  // ---- item terakhir di daftar: layar selesai tetap memberi kesempatan mengurungkan
  await kembali(s.page); await s.page.waitForSelector('#scr-cycle.is-on .list');
  await s.page.evaluate(() => document.querySelector('[data-aksi="tugas-segar"]').click());
  await s.page.waitForFunction(() => !document.querySelector('#scr-cycle .iconbtn.is-putar'));
  const sisa = be.api('getMyPendingTasks', PETUGAS2);
  const c = sisa.find((t) => t.no !== b.no && t.no !== a2.no && t.no !== a.no && sisa.filter((x) => (x.lokasi + ' ' + x.article + ' ' + x.description).toLowerCase().indexOf(t.lokasi.toLowerCase()) >= 0).length === 1);
  await s.page.fill('#tgCari-cycle', c.lokasi); await tunggu(350);
  cek('pencarian: daftar tinggal satu item dan tombol mulai menyebut jumlahnya', (await s.page.$$('#tgDaftar-cycle .row')).length === 1 && /Mulai 1 item terpilih/.test(await teks(s.page, '.tg-mulai')), await teks(s.page, '.tg-mulai'));
  await s.page.click('#tgDaftar-cycle .row'); await s.page.waitForSelector('#fkPad');
  await ketik(c.qtySystem + 5); await s.page.click('#fkSimpan');
  await s.page.waitForSelector('.fk__tuntas');
  cek('item terakhir: layar selesai menampilkan hasil terakhir dengan tombol Urungkan', !!(await s.page.$('.fk__tuntas')) && !!(await s.page.$('.simpanan [data-aksi="fokus-urung"]')));
  await s.page.click('[data-aksi="fokus-urung"]'); await s.page.waitForSelector('#fkPad');
  cek('Urungkan dari layar selesai mengembalikan item dan qty, belum ada yang terkirim', (await lokasiLayar()) === c.lokasi && angka(await teks(s.page, '#fkAngka')) === c.qtySystem + 5 && !s.panggilan.some((p) => p.action === 'submitCount' && p.args[0] === c.no));
  for (let i = 0; i < String(c.qtySystem + 5).length; i++) await s.page.click('[data-k="hapus"]');
  await ketik(c.qtySystem); await s.page.click('#fkSimpan');
  await s.page.waitForSelector('.fk__tuntas');
  await kembali(s.page); await antreKosong(s.page);
  be.worker();
  cek('hasil yang diperbaiki tercatat dengan qty yang benar', baris(c.no)[10] === 'Selesai' && Number(baris(c.no)[11]) === c.qtySystem && baris(c.no)[13] === 'HIT', baris(c.no));
  await s.tutup();
});

/* ------------------------------------------------------------------ */
await bagian('4. Inventory: validasi ulang lalu verifikasi', async () => {
  const s = await buka({ be });
  await masuk(s.page, INV);
  const antre = be.api('getPendingValidasi', INV);
  cek('menu inventory: Home, Cycle, Validasi, Verifikasi', (await s.page.$$eval('.nav__item span:not(.nav__lencana)', (n) => n.map((x) => x.textContent))).join(',') === 'Home,Cycle,Validasi,Verifikasi');
  cek('lencana Validasi sama dengan server', angka(await teks(s.page, '#tab-validasi .nav__lencana')) === antre.length, antre.length);
  await s.page.click('#tab-validasi'); await s.page.waitForSelector('#scr-validasi .list');
  await s.page.click('#scr-validasi [data-aksi="tugas-mulai"]'); await s.page.waitForSelector('#fkPad');
  cek('mode fokus validasi memakai warna validasi', await s.page.$eval('#scr-hitung', (n) => n.classList.contains('fk--ungu')));
  const lok = await s.page.$eval('.rak-besar', (n) => n.getAttribute('aria-label').replace('Lokasi ', '')), sku = await teks(s.page, '.fk__sku');
  const it = antre.find((x) => x.lokasi === lok && String(x.article) === sku);
  const q = it.qtySystem + 2; // tetap selisih → menjadi task verifikasi
  for (const d of String(q)) await s.page.click('[data-k="' + d + '"]');
  await s.page.click('#fkSimpan'); await kembali(s.page); await antreKosong(s.page);
  be.worker();
  const task = be.api('getOpenTasks', INV).find((t) => t.lokasi === lok && String(t.article) === sku);
  const rw = be.rows('Riwayat').find((r) => r[0] === it.id);
  cek('validasi yang tetap selisih menjadi task verifikasi', !!task && task.id === it.id && Number(rw[13]) === q && rw[14] === 'DISCREPANCY' && rw[18] === 'Open', rw && rw.slice(11, 19));
  const besar = Math.abs(task.selisih); // besar selisih task = selisih hitungan pertama petugas

  // ---- verifikasi
  await s.page.click('#tab-verifikasi'); await s.page.evaluate(() => document.querySelector('[data-aksi="verif-segar"]').click());
  await s.page.waitForFunction((id) => !!document.querySelector('[data-aksi="verif-buka"][data-id="' + id + '"]'), task.id, { timeout: 15000 });
  const buka0 = be.api('getOpenTasks', INV).length;
  cek('jumlah task sama dengan server', angka(await teks(s.page, '#scr-verifikasi .hd__ket')) === buka0, await teks(s.page, '#scr-verifikasi .hd__ket'));
  await s.page.fill('#vfCari', lok); await tunggu(350);
  cek('pencarian menyaring daftar', (await s.page.$$eval('#vfDaftar .rak', (n, l) => n.length >= 1 && n.every((x) => x.textContent === l), lok)));
  await s.page.click('[data-aksi="verif-buka"][data-id="' + task.id + '"]'); await s.page.waitForSelector('#vfForm');
  await s.page.click('#vfSimpan');
  cek('menutup task tanpa catatan ditolak di tempat', /Catatan belum diisi/.test(await teks(s.page, '#vfPesan')));
  await s.page.selectOption('#vfKategori', task.selisih > 0 ? 'Lebih Picking' : 'Kurang Picking');
  cek('kategori dengan bukti menampilkan kolom bukti WMS', !(await s.page.$eval('#vfBuktiWadah', (n) => n.hidden)));
  await s.page.fill('#vfCatatan', 'Barang lebih karena picking kurang ambil.');
  await s.page.fill('#vfBukti', ['1', 'IDM', 'PICKING', sku, 'X', 'G', 'L', lok, '', 'STAGE', '', 'S1', String(besar + 4), '2026-10-01', 'wms.uji'].join('\t'));
  await s.page.waitForFunction((n) => new RegExp('Total qty di bukti ' + n).test(document.getElementById('vfBuktiCek').textContent), besar + 4);
  cek('bukti dengan qty salah langsung ditandai', true);
  await s.page.fill('#vfBukti', ['1', 'IDM', 'PICKING', sku, 'X', 'G', 'L', lok, '', 'STAGE', '', 'S1', String(besar), '2026-10-01', 'wms.uji'].join('\t'));
  await s.page.waitForFunction(() => /Bukti sesuai/.test(document.getElementById('vfBuktiCek').textContent));
  if (!(await s.page.$eval('#vfPic', (n) => n.value))) await s.page.selectOption('#vfPic', { index: 1 });
  await s.page.click('#vfSimpan');
  await s.page.waitForSelector('#vfForm', { state: 'detached' });
  be.worker();
  cek('task ditutup di server dengan bukti', !be.api('getOpenTasks', INV).some((t) => t.id === task.id) && be.rows('Riwayat').find((r) => r[0] === task.id)[18] === 'Selesai', be.rows('Riwayat').find((r) => r[0] === task.id).slice(17, 21));
  cek('task yang ditutup langsung hilang dari daftar', !(await s.page.$('[data-aksi="verif-buka"][data-id="' + task.id + '"]')));

  // ---- pasangan plus minus (dicari otomatis)
  await s.page.fill('#vfCari', ''); await tunggu(300);
  await s.page.click('[data-aksi="verif-tampilan"][data-v="pasangan"]');
  await s.page.waitForSelector('[data-aksi="verif-pasang"]');
  const p = be.info.pasangan;
  // semua calon pasangan yang sah ditampilkan; data contoh bisa punya pasangan lain yang kebetulan cocok
  const kartu = s.page.locator('.vf-pasang').filter({ hasText: p.minus }).filter({ hasText: p.plus });
  cek('pasangan plus minus ditemukan otomatis', (await kartu.count()) === 1, { p, kartu: await s.page.$$eval('.vf-pasang', (n) => n.map((x) => x.textContent.replace(/\s+/g, ' ').trim())) });
  await kartu.click(); await s.page.waitForSelector('#vpSimpan');
  await s.page.selectOption('#vpPic', { index: 1 });
  await s.page.fill('#vpCatatan', 'Barang salah taruh, sudah dipindahkan.');
  await s.page.fill('#vpBukti', ['1', 'IDM', 'MOVE', p.article, 'X', 'G', 'L', p.minus, '', p.plus, '', 'S1', '2', '2026-10-01', 'wms.uji'].join('\t'));
  await s.page.waitForFunction(() => /Bukti sesuai/.test(document.getElementById('vpCek').textContent));
  await s.page.click('#vpSimpan');
  await s.page.waitForSelector('#vpSimpan', { state: 'detached' });
  be.worker();
  cek('kedua task pasangan ditutup di server', !be.api('getOpenTasks', INV).some((t) => String(t.article) === String(p.article) && (t.lokasi === p.minus || t.lokasi === p.plus)));
  await s.page.click('[data-aksi="verif-tampilan"][data-v="riwayat"]');
  await s.page.waitForSelector('#scr-verifikasi [data-aksi="verif-log-buka"]');
  cek('riwayat memuat task yang baru ditutup', (await s.page.$$eval('#scr-verifikasi [data-aksi="verif-log-buka"] .rak', (n) => n.map((x) => x.textContent))).includes(lok));
  await s.tutup();
});

/* ------------------------------------------------------------------ */
await bagian('5. Admin: Upload Data dari file Excel', async () => {
  const s = await buka({ be });
  await masuk(s.page, ADMIN);
  await s.page.click('#tab-upload'); await s.page.waitForSelector('.up-berkas');
  const lokasi = be.rows('Lokasi_Aktif').map((r) => r[0]);
  const L = (i) => lokasi[(i * 37 + 11) % lokasi.length];
  const transaksi = [['Type', 'Article', 'Description', 'From Location', 'To Location', 'Qty', 'AddWho']], stok = [['Location', 'Article', 'Qty']];
  for (let i = 0; i < 12; i++) {
    const art = String(88000000 + i), move = i % 3 === 0;
    transaksi.push(move ? ['Move', art, 'BARANG UJI ' + i, 'STAGING', L(i), 2, 'wms.uji'] : ['Picking', art, 'BARANG UJI ' + i, L(i), '', 1, 'wms.uji']);
    stok.push([L(i), art, 10 + i]);
  }
  transaksi.push(['Adjustment', '88009999', 'TIPE LAIN', L(40), '', 1, 'wms.uji']);
  stok.push(['Z01.001.1', '87000000', 5]); // stok yang tidak ada transaksinya tidak perlu dikirim
  const tulis = (nama, aoa) => { const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), 'Sheet1'); const f = path.join(TMP, nama); XLSX.writeFile(wb, f); return f; };
  const fT = tulis('transaksi.xlsx', transaksi), fS = tulis('stok.xlsx', stok), fSalah = tulis('salah.xlsx', [['Kolom A', 'Kolom B'], [1, 2]]);
  await s.page.setInputFiles('input[data-berkas="transaksi"]', fSalah);
  await s.page.waitForSelector('.up-berkas.is-galat');
  cek('file dengan kolom salah langsung ditolak dengan nama kolomnya', /Kolom tidak ditemukan: type/.test(await teks(s.page, '.up-berkas.is-galat')));
  await s.page.setInputFiles('input[data-berkas="transaksi"]', fT);
  await s.page.waitForFunction(() => /13 baris/.test(document.querySelector('.up-berkas').textContent));
  cek('ringkasan file transaksi dihitung di HP', /13 baris: 4 Move, 8 Picking, 1 tipe lain dilewati/.test(await teks(s.page, '.up-berkas')), await teks(s.page, '.up-berkas'));
  await s.page.setInputFiles('input[data-berkas="stok"]', fS);
  await s.page.waitForFunction(() => document.querySelectorAll('.up-berkas.is-ok').length === 2);
  await s.page.click('#upProses');
  cek('proses tanpa petugas ditolak', /Pilih minimal satu petugas/.test(await teks(s.page, '#upPesan')));
  await s.page.click('[data-aksi="upload-peran"][data-peran="outbound"]');
  await s.page.click('[data-aksi="upload-peran"][data-peran="storing"]');
  await s.page.click('[data-aksi="upload-langkah"][data-n="tangga"][data-d="1"]');
  await s.page.click('[data-aksi="upload-langkah"][data-n="rt"][data-d="1"]');
  const sebelum = dataCount().length;
  await s.page.click('#upProses');
  await s.page.waitForSelector('.up-hasil', { timeout: 30000 });
  const req = s.panggilan.find((p) => p.action === 'importRawData');
  cek('stok yang dikirim hanya lokasi yang ada di transaksi', req.args[1].length === 12 && req.args[0].length === 13, { stok: req.args[1].length, transaksi: req.args[0].length });
  const ditugaskan = angka(await teks(s.page, '.up-hasil b'));
  cek('tugas baru tercatat di server', dataCount().length - sebelum === ditugaskan && ditugaskan > 0 && ditugaskan <= 12, { ditugaskan, tambah: dataCount().length - sebelum });
  cek('baris tipe lain dilaporkan dilewati', /tipe selain Move dan Picking/.test(await teks(s.page, '.lembar')));
  await s.page.click('.lembar [data-aksi="ke"]'); await s.page.waitForSelector('#scr-home.is-on .kpi');
  await s.page.waitForFunction(() => !document.querySelector('#hmSegar.is-putar'), null, { timeout: 15000 }); // Home menampilkan data lama dulu, lalu memperbarui
  cek('kembali ke Home dengan angka terbaru', angka((await s.page.$$eval('.kpi__nilai', (n) => n.map((x) => x.textContent)))[1]) === be.api('getHomeBundle', ADMIN).outstanding, { tampil: await s.page.$$eval('.kpi__nilai', (n) => n.map((x) => x.textContent)), server: be.api('getHomeBundle', ADMIN).outstanding });
  await s.page.click('#tab-upload'); await s.page.waitForSelector('.up-berkas');
  cek('pilihan petugas terakhir diingat untuk upload berikutnya', (await s.page.evaluate(() => Object.keys(JSON.parse(localStorage.getItem(Object.keys(localStorage).find((k) => k.indexOf('ct.petugas.') === 0)))).length)) === 6);

  // ---- file .csv: SKU berawalan nol dan SKU 19 digit tidak boleh ditebak sebagai angka
  const csvT = path.join(TMP, 'transaksi.csv'), csvS = path.join(TMP, 'stok.csv');
  // baris ketiga: SKU ditulis ="…" (cara ekspor memaksa teks di Excel) dan qty stok memakai pemisah ribuan
  fs.writeFileSync(csvT, 'Type,Article,Description,From Location,To Location,Qty,AddWho\nPicking,00880001,BARANG NOL,' + L(60) + ',,1,wms.uji\nMove,8899000000000000123,BARANG PANJANG,STAGING,' + L(61) + ',2,wms.uji\nPicking,="00990002",BARANG RUMUS,' + L(62) + ',,1,wms.uji\n');
  // file stok juga memuat baris yang tidak dipakai dan tidak boleh membuat file ditolak: judul yang terulang,
  // baris total, tanda kosong, dan qty bukan angka di lokasi yang tidak punya transaksi
  fs.writeFileSync(csvS, 'Location,Article,Qty\n' + L(60) + ',00880001,7\n' + L(61) + ',8899000000000000123,4\nLocation,Article,Qty\n' + L(62) + ',="00990002","1,234"\n' + L(63) + ',55500001,-\n' + L(64) + ',55500002,NULL\n' + L(65) + ',55500003,tiga\nTOTAL,,"1,245"\n');
  const csvSalah = path.join(TMP, 'stok-salah.csv');
  fs.writeFileSync(csvSalah, 'Location,Article,Qty\n' + L(60) + ',00880001,7\n' + L(61) + ',8899000000000000123,empat\n');
  await s.page.setInputFiles('input[data-berkas="transaksi"]', csvT);
  await s.page.waitForFunction(() => /3 baris: 1 Move, 2 Picking/.test(document.querySelector('.up-berkas').textContent));
  await s.page.setInputFiles('input[data-berkas="stok"]', csvSalah);
  await s.page.waitForFunction(() => { const u = document.querySelectorAll('.up-berkas')[1]; return u.classList.contains('is-galat') || u.classList.contains('is-ok'); });
  const pesanSalah = await s.page.$$eval('.up-berkas', (n) => n[1].textContent.replace(/\s+/g, ' '));
  cek('.csv: qty yang bukan angka ditolak di HP dengan nomor barisnya', /Baris 3/.test(pesanSalah) && /empat/.test(pesanSalah) && (await s.page.$$eval('.up-berkas', (n) => n[1].classList.contains('is-galat'))), pesanSalah);
  await s.page.setInputFiles('input[data-berkas="stok"]', csvS);
  await s.page.waitForFunction(() => document.querySelectorAll('.up-berkas.is-ok').length === 2);
  await s.page.click('#upProses'); await s.page.waitForSelector('.up-hasil', { timeout: 30000 });
  const req2 = s.panggilan.filter((p) => p.action === 'importRawData').pop();
  cek('.csv: SKU berawalan nol, SKU 19 digit, dan SKU ber-="…" dikirim sebagai teks aslinya', req2.args[0].map((r) => r[1]).join() === '00880001,8899000000000000123,00990002' && req2.args[1].map((r) => r[1]).join() === '00880001,8899000000000000123,00990002', { transaksi: req2.args[0].map((r) => r[1]), stok: req2.args[1].map((r) => r[1]) });
  cek('.csv: qty "1,234" dikirim sebagai 1234, bukan teks yang menjadi 0 di server', req2.args[1].map((r) => String(r[2])).join() === '7,4,1234', req2.args[1].map((r) => r[2]));
  await s.tutup();
});

/* ------------------------------------------------------------------ */
await bagian('6. Admin: Report (Dashboard, Productivity, Analytics)', async () => {
  const s = await buka({ be });
  await masuk(s.page, ADMIN);
  await s.page.click('#tab-report'); await s.page.waitForSelector('#rpTren rect');
  const d = be.api('getDashboardData', 'daily', be.info.today, 7, ADMIN), k = d.summary.kpi;
  const tiga = await s.page.$$eval('.rp-akurasi .rp-tiga b', (n) => n.map((x) => Number(x.textContent.replace(/\D/g, ''))));
  cek('Dashboard: total, hit, discrepancy sama dengan server', JSON.stringify(tiga) === JSON.stringify([k.total, k.hit, k.discrepancy]), { tiga, k });
  cek('grafik tren punya satu kolom per hari', (await s.page.$$('#rpTren rect[data-i]')).length === d.trend.length);
  const kemarin = d.trend.length - 2;
  await s.page.click('#rpTren rect[data-i="' + kemarin + '"]');
  cek('ketuk kolom grafik menampilkan angkanya', new RegExp('Hit ' + d.trend[kemarin].hit.toLocaleString('id-ID')).test(await teks(s.page, '#rpTrenInfo')) && d.trend[kemarin].hit > 0, await teks(s.page, '#rpTrenInfo'));
  await s.page.click('#rpTren rect[data-i="0"]');
  cek('hari tanpa data ditulis apa adanya', /Tidak ada data/.test(await teks(s.page, '#rpTrenInfo')));
  await s.page.click('.rp-rank[data-aksi="report-user"]'); await s.page.waitForSelector('.lembar .rp-empat');
  cek('ringkasan per petugas terbuka', (await s.page.$$('.lembar .rp-empat > div')).length >= 4);
  await kembali(s.page); await s.page.waitForSelector('.lembar-wadah', { state: 'detached' });
  await s.page.click('[data-aksi="report-tampilan"][data-v="productivity"]'); await s.page.waitForSelector('#rpJam rect');
  const pr = be.api('getProductivity', be.info.today, ADMIN, 0);
  cek('Productivity: jumlah item dan petugas aktif sama dengan server', angka(await teks(s.page, '.rp-tiga--kartu div:nth-child(1) b')) === pr.ringkasan.totalItem && angka(await teks(s.page, '.rp-tiga--kartu div:nth-child(2) b')) === pr.ringkasan.petugasAktif);
  cek('Productivity: satu baris per petugas', (await s.page.$$('#rpIsi .orang')).length === pr.petugas.length);
  await s.page.click('[data-aksi="report-tampilan"][data-v="analytics"]'); await s.page.waitForSelector('.rp-donat');
  const an = be.api('getAnalyticsRootCauseData', await s.page.$eval('#rpDari', (n) => n.value), await s.page.$eval('#rpSampai', (n) => n.value), ADMIN);
  cek('Analytics: jumlah discrepancy di donat sama dengan server', angka(await teks(s.page, '.rp-donat svg text')) === an.kpi.discrepancy);
  const akar = await s.page.$('.rp-akar:not([disabled])');
  if (akar) { await akar.click(); await s.page.waitForSelector('.lembar .rp-item, .lembar .kosong'); cek('rincian akar masalah terbuka', true); await kembali(s.page); }

  // ---- tanggal diganti lagi selagi laporan masih dimuat: yang tampil harus milik pilihan terakhir
  const geser = (t, n) => new Date(Date.parse(t + 'T00:00:00Z') + n * 86400000).toISOString().slice(0, 10);
  const h1 = geser(be.info.today, -1), h2 = geser(be.info.today, -2);
  const k1 = be.api('getDashboardData', 'daily', h1, 7, ADMIN).summary.kpi, k2 = be.api('getDashboardData', 'daily', h2, 7, ADMIN).summary.kpi;
  await s.page.click('[data-aksi="report-tampilan"][data-v="dashboard"]'); await s.page.waitForSelector('#rpTren rect');
  const gH1 = gerbang(); // jawaban untuk tanggal pertama ditahan sampai tanggal kedua selesai dimuat
  s.atur.tunda = (req) => (req.action === 'getDashboardData' && req.args[1] === h1 ? gH1.janji : 0);
  await s.page.fill('#rpTanggal', h1);
  await sampai(() => s.panggilan.some((p) => p.action === 'getDashboardData' && p.args[1] === h1));
  cek('data periode lama tidak ditampilkan di bawah tanggal yang baru', !(await s.page.$('.rp-akurasi')) && (await s.page.$eval('#rpTanggal', (n) => n.value)) === h1);
  await s.page.fill('#rpTanggal', h2);
  await s.page.waitForFunction(() => !document.querySelector('#scr-report .iconbtn.is-putar') && !!document.querySelector('.rp-akurasi, #rpIsi .kosong'), null, { timeout: 15000 }).catch(() => {});
  gH1.buka(); await tunggu(600); // jawaban untuk tanggal pertama tiba belakangan dan harus diabaikan
  const tiga2 = await s.page.$$eval('.rp-akurasi .rp-tiga b', (n) => n.map((x) => Number(x.textContent.replace(/\D/g, ''))));
  cek('tanggal diganti saat memuat: angka yang tampil milik tanggal terakhir', (await s.page.$eval('#rpTanggal', (n) => n.value)) === h2 && JSON.stringify(tiga2) === JSON.stringify([k2.total, k2.hit, k2.discrepancy]) && JSON.stringify([k1.total, k1.hit, k1.discrepancy]) !== JSON.stringify([k2.total, k2.hit, k2.discrepancy]), { tiga2, k1, k2 });
  s.atur.tunda = null;

  // ---- memperbarui periode yang sama lalu gagal: data lama tetap utuh, pemberitahuannya bisa dibaca orang awam
  const tigaAngka = () => s.page.$$eval('.rp-akurasi .rp-tiga b', (n) => n.map((x) => Number(x.textContent.replace(/\D/g, ''))));
  const adaOutstanding = async () => (await s.page.$$('[data-aksi="report-backlog"]')).length;
  await s.page.fill('#rpTanggal', be.info.today);
  await s.page.waitForFunction(() => !document.querySelector('#scr-report .iconbtn.is-putar') && !!document.querySelector('.rp-akurasi'), null, { timeout: 15000 });
  await s.page.click('[data-aksi="report-masalah"]'); await s.page.waitForSelector('#rpMasalah .rp-item');
  const sebelumSegar = { outstanding: await adaOutstanding(), tiga: await tigaAngka(), item: (await s.page.$$('#rpMasalah .rp-item')).length };
  const gSegar = gerbang();
  s.atur.tunda = (req) => (req.action === 'getDashboardData' || req.action === 'getPendingBacklog' ? gSegar.janji : 0);
  s.atur.ubah = (req) => (req.action === 'getDashboardData' || req.action === 'getPendingBacklog' ? { ok: false, error: 'Service Spreadsheets failed while accessing document with id 1AbC.' } : undefined);
  await s.page.click('[data-aksi="report-segar"]');
  await s.page.waitForFunction(() => !!document.querySelector('#scr-report .iconbtn.is-putar'));
  cek('selagi memperbarui: Outstanding Cycle dan daftar item bermasalah tetap tampil', sebelumSegar.outstanding > 0 && sebelumSegar.item > 0 && (await adaOutstanding()) === sebelumSegar.outstanding && (await s.page.$$('#rpMasalah .rp-item')).length === sebelumSegar.item, { sebelumSegar, outstanding: await adaOutstanding(), item: (await s.page.$$('#rpMasalah .rp-item')).length });
  gSegar.buka();
  await s.page.waitForFunction(() => !document.querySelector('#scr-report .iconbtn.is-putar'), null, { timeout: 15000 });
  const infoSegar = await s.page.$eval('#rpIsi .info', (n) => n.textContent).catch(() => '');
  cek('pembaruan gagal: angka, Outstanding Cycle, dan daftar item lama tetap utuh', /Belum bisa diperbarui/.test(infoSegar) && (await adaOutstanding()) === sebelumSegar.outstanding && JSON.stringify(await tigaAngka()) === JSON.stringify(sebelumSegar.tiga) && (await s.page.$$('#rpMasalah .rp-item')).length === sebelumSegar.item, { infoSegar, outstanding: await adaOutstanding() });
  cek('pemberitahuan gagal tidak memuat teks teknis dari server', infoSegar.length > 0 && !/Service|Spreadsheets|Exception|failed|document/i.test(infoSegar), infoSegar);
  s.atur.tunda = null; s.atur.ubah = null;
  await s.page.click('[data-aksi="report-segar"]');
  await s.page.waitForFunction(() => !document.querySelector('#scr-report .iconbtn.is-putar') && !document.querySelector('#rpIsi .info--jingga'), null, { timeout: 15000 });
  cek('pembaruan berikutnya berhasil: pemberitahuan hilang dan daftar item ikut diperbarui', (await s.page.$$('#rpMasalah .rp-item')).length === sebelumSegar.item && s.panggilan.filter((p) => p.action === 'getProblemItemsDetail').length >= 2);
  await s.tutup();
});

/* ------------------------------------------------------------------ */
await bagian('7. Config: user, akses, pembagian tugas, lokasi aktif', async () => {
  const s = await buka({ be });
  await masuk(s.page, ADMIN);
  // pindah tab selagi tab pertama masih dimuat: tab baru tetap harus terisi (dulu tertahan di kerangka)
  const gUser = gerbang();
  s.atur.tunda = (req) => (req.action === 'getDaftarUserMaster' ? gUser.janji : 0);
  await s.page.click('[data-aksi="akun"]'); await s.page.click('[data-aksi="akun-config"]');
  await s.page.waitForSelector('#scr-config.is-on .kerangka');
  await sampai(() => s.panggilan.some((p) => p.action === 'getDaftarUserMaster')); // daftar user sedang dimuat (ditahan)
  await s.page.click('[data-aksi="config-tab"][data-v="akses"]');
  let aksesTampil = true;
  try { await s.page.waitForSelector('#cfAksesNik', { timeout: 6000 }); } catch (e) { aksesTampil = false; }
  cek('pindah tab saat tab lain masih dimuat: tab baru tetap terisi', aksesTampil);
  await s.page.fill('#cfAksesNik', '123456.ketik').catch(() => {});
  gUser.buka(); s.atur.tunda = null; // daftar user (tab pertama) selesai dimuat di belakang
  await tunggu(700);
  cek('isian di tab yang sedang dibuka tidak terhapus oleh tab lain yang baru selesai dimuat', (await s.page.$eval('#cfAksesNik', (n) => n.value).catch(() => '')) === '123456.ketik');
  await s.page.click('[data-aksi="config-tab"][data-v="user"]');
  await s.page.waitForSelector('#cfUsers .list');
  cek('tab pertama sudah terisi tanpa dimuat dua kali', s.panggilan.filter((p) => p.action === 'getDaftarUserMaster').length === 1);
  await s.page.click('#cfUsers .row:has-text("' + ADMIN + '")'); await s.page.waitForSelector('#cfPeranUbah');
  cek('akun sendiri tidak punya tombol nonaktifkan', !(await s.page.$('#cfStatusUbah')) && /Akun sendiri tidak bisa dinonaktifkan/.test(await teks(s.page, '.lembar')));
  await kembali(s.page); await s.page.waitForSelector('.lembar-wadah', { state: 'detached' });

  // perubahan yang baru selesai setelah pindah tab: tab yang sedang dibuka tidak dimuat ulang, isiannya tidak hilang
  const gTambah = gerbang();
  s.atur.tunda = (req) => (req.action === 'tambahUserMaster' ? gTambah.janji : 0);
  await s.page.fill('#cfNik', '999002.UJI'); await s.page.selectOption('#cfPeran', 'inbound'); await s.page.click('[data-aksi="config-user-tambah"]');
  await sampai(() => s.panggilan.some((p) => p.action === 'tambahUserMaster'));
  await s.page.click('[data-aksi="config-tab"][data-v="akses"]'); await s.page.waitForSelector('#cfAksesNik');
  await s.page.fill('#cfAksesNik', '777777.ketik');
  const muatAkses = s.panggilan.filter((p) => p.action === 'getDaftarAksesSetting').length, muatUser = s.panggilan.filter((p) => p.action === 'getDaftarUserMaster').length;
  gTambah.buka(); s.atur.tunda = null;
  await sampai(() => s.panggilan.filter((p) => p.action === 'getDaftarUserMaster').length > muatUser, 5000).catch(() => {});
  await tunggu(500);
  cek('perubahan selesai setelah pindah tab: isian tetap ada dan tab yang dibuka tidak dimuat ulang', (await s.page.$eval('#cfAksesNik', (n) => n.value).catch(() => '')) === '777777.ketik' && s.panggilan.filter((p) => p.action === 'getDaftarAksesSetting').length === muatAkses, { isi: await s.page.$eval('#cfAksesNik', (n) => n.value).catch(() => '(hilang)'), muatAkses: s.panggilan.filter((p) => p.action === 'getDaftarAksesSetting').length - muatAkses });
  await s.page.click('[data-aksi="config-tab"][data-v="user"]'); await s.page.waitForSelector('#cfUsers .list');
  await s.page.waitForFunction(() => [...document.querySelectorAll('#cfUsers .row__t')].some((n) => n.textContent === '999002.UJI'), null, { timeout: 8000 }).catch(() => {});
  cek('tab yang datanya berubah sudah diperbarui saat dibuka lagi', (await s.page.$$eval('#cfUsers .row__t', (n) => n.map((x) => x.textContent))).includes('999002.UJI') && be.api('getUserRole', '999002.uji').role === 'inbound');

  // akun Developer (termasuk ejaan lama "dewa") tidak bisa diubah oleh yang bukan Developer
  be.sheet('Master_User').appendRow(['999003.DEWA', 'dewa', 'Aktif', be.info.facilityId]); be.ctx.clearMasterCache_();
  await s.page.click('[data-aksi="config-segar"]');
  await s.page.waitForFunction(() => [...document.querySelectorAll('#cfUsers .row__t')].some((n) => n.textContent === '999003.DEWA'), null, { timeout: 15000 });
  await s.page.click('#cfUsers .row:has-text("999003.DEWA")'); await s.page.waitForSelector('.lembar.lembar, .lembar');
  await tunggu(300);
  cek('akun Developer tidak menawarkan ubah peran atau nonaktifkan kepada yang bukan Developer', !(await s.page.$('#cfPeranSimpan')) && !(await s.page.$('#cfStatusUbah')) && /hanya bisa diubah oleh Developer/.test(await teks(s.page, '.lembar')), await teks(s.page, '.lembar'));
  await kembali(s.page); await s.page.waitForSelector('.lembar-wadah', { state: 'detached' });
  await s.page.fill('#cfNik', '999001.UJI'); await s.page.click('[data-aksi="config-user-tambah"]');
  await s.page.waitForSelector('.toast.is-on');
  cek('tambah user tanpa peran ditolak', /Pilih peran/.test(await teks(s.page, '.toast')));
  await s.page.selectOption('#cfPeran', 'outbound'); await s.page.click('[data-aksi="config-user-tambah"]');
  await s.page.waitForFunction(() => [...document.querySelectorAll('#cfUsers .row__t')].some((n) => n.textContent === '999001.UJI'), null, { timeout: 15000 });
  const u1 = be.api('getUserRole', '999001.uji');
  cek('user baru tersimpan di server dan masuk ke facility admin', !!u1 && u1.role === 'outbound' && u1.facilityName === 'RDC Tallo Makassar', u1);
  await s.page.click('#cfUsers .row:has-text("999001.UJI")'); await s.page.waitForSelector('#cfPeranUbah');
  await s.page.selectOption('#cfPeranUbah', 'lp'); await s.page.click('#cfPeranSimpan');
  await s.page.waitForSelector('#cfPeranUbah', { state: 'detached' });
  const u2 = be.api('getUserRole', '999001.uji');
  cek('ganti peran tersimpan dan facility tidak terhapus', u2.role === 'lp' && u2.facilityId === u1.facilityId, u2);
  await s.page.waitForFunction(() => [...document.querySelectorAll('#cfUsers .row')].some((n) => /999001\.UJI/.test(n.textContent) && /LP/.test(n.textContent)));
  await s.page.click('#cfUsers .row:has-text("999001.UJI")'); await s.page.waitForSelector('#cfStatusUbah');
  await s.page.click('#cfStatusUbah'); await s.page.waitForSelector('[data-j="1"]'); await s.page.click('[data-j="1"]');
  await s.page.waitForFunction(() => [...document.querySelectorAll('#cfUsers .row')].some((n) => /999001\.UJI/.test(n.textContent) && /Nonaktif/.test(n.textContent)), null, { timeout: 15000 });
  cek('user dinonaktifkan: tidak bisa masuk lagi', be.api('getUserRole', '999001.uji') === null);

  await s.page.click('[data-aksi="config-tab"][data-v="akses"]'); await s.page.waitForSelector('#cfAksesNik');
  await s.page.fill('#cfAksesNik', INV); await s.page.fill('#cfAksesNama', 'Inventory'); await s.page.click('[data-aksi="config-akses-tambah"]');
  await s.page.waitForSelector('[data-aksi="config-akses-cabut"]', { timeout: 15000 });
  cek('akses Config diberikan', be.api('getUserRole', INV).aksesSettingOverflow === true);
  await s.page.click('[data-aksi="config-akses-cabut"]'); await s.page.waitForSelector('[data-j="1"]'); await s.page.click('[data-j="1"]');
  await s.page.waitForSelector('[data-aksi="config-akses-cabut"]', { state: 'detached', timeout: 15000 });
  cek('akses Config dicabut', be.api('getUserRole', INV).aksesSettingOverflow === false);

  await s.page.click('[data-aksi="config-tab"][data-v="tugas"]'); await s.page.waitForSelector('.cf-matriks');
  const sebelum = be.api('getLevelAssignmentConfig', ADMIN).matrix.find((m) => m.grupAlat === 'bawah');
  await s.page.click('input[data-m-role="storing"][data-m-grup="bawah"]');
  await s.page.click('[data-aksi="config-cara"][data-grup="bawah"][data-v="prioritas"]');
  cek('urutan prioritas ditampilkan, peran baru di urutan terakhir', /lalu Storing\./.test(await teks(s.page, '#cfCara')), await teks(s.page, '#cfCara'));
  await s.page.click('[data-aksi="config-matriks-simpan"]');
  await s.page.waitForFunction(() => /berhasil|tersimpan|disimpan/i.test((document.querySelector('.toast') || {}).textContent || ''), null, { timeout: 15000 });
  const sesudah = be.api('getLevelAssignmentConfig', ADMIN).matrix.find((m) => m.grupAlat === 'bawah');
  cek('matriks tersimpan: mode prioritas, urutan peran lama dipertahankan', sesudah.mode === 'prioritas' && JSON.stringify(sesudah.roles) === JSON.stringify(sebelum.roles.concat(['storing'])), sesudah);

  await s.page.click('[data-aksi="config-tab"][data-v="facility"]'); await s.page.waitForSelector('[data-aksi="config-fac-lokasi"]');
  await s.page.click('[data-aksi="config-fac-lokasi"]'); await s.page.waitForSelector('#cfLokSimpan');
  const nLok = be.rows('Lokasi_Aktif').length;
  cek('jumlah lokasi aktif tampil', angka(await teks(s.page, '.lembar .up-hasil b')) === nLok);
  await s.page.click('#cfLokSumber [data-v="tempel"]');
  await s.page.fill('#cfLokTeks', 'D01.001.1\nD01.001.2\n' + be.rows('Lokasi_Aktif')[0][0]);
  await s.page.click('#cfLokSimpan'); await s.page.waitForSelector('#cfLokSimpan', { state: 'detached', timeout: 15000 });
  cek('impor lokasi (tambahkan) menambah dua lokasi baru saja', be.rows('Lokasi_Aktif').length === nLok + 2, be.rows('Lokasi_Aktif').length - nLok);
  const fLok = path.join(TMP, 'lokasi.csv'); fs.writeFileSync(fLok, 'Lokasi\nD02.001.1\nD02.001.2\nD02.001.2\nD02.001.3\n0107\n12E3\n');
  await s.page.click('[data-aksi="config-fac-lokasi"]'); await s.page.waitForSelector('#cfLokSimpan');
  await s.page.setInputFiles('#cfLokBerkas', fLok);
  await s.page.waitForSelector('#cfLokUbin.is-ok');
  cek('file lokasi dibaca begitu dipilih (baris judul dan duplikat dilewati)', /5 lokasi terbaca/.test(await teks(s.page, '#cfLokUbin')), await teks(s.page, '#cfLokUbin'));
  await s.page.click('#cfLokSimpan'); await s.page.waitForSelector('#cfLokSimpan', { state: 'detached', timeout: 15000 });
  cek('impor lokasi dari file tersimpan', be.rows('Lokasi_Aktif').length === nLok + 7);
  const kirimLok = s.panggilan.filter((p) => p.action === 'importLokasiAktif').pop().args[2];
  cek('.csv lokasi: kode "0107" dan "12E3" dikirim apa adanya, bukan sebagai angka', kirimLok.join() === 'D02.001.1,D02.001.2,D02.001.3,0107,12E3', kirimLok);
  await s.tutup();

  // Developer: "Buat facility" diketuk berkali-kali hanya membuat satu facility
  be.sheet('Master_User').appendRow(['999004.DEV', 'developer', 'Aktif', be.info.facilityId]); be.ctx.clearMasterCache_();
  const d = await buka({ be });
  await masuk(d.page, '999004.dev');
  await d.page.click('[data-aksi="akun"]'); await d.page.click('[data-aksi="akun-config"]');
  await d.page.click('[data-aksi="config-tab"][data-v="facility"]'); await d.page.waitForSelector('[data-aksi="config-fac-baru"]');
  await d.page.click('[data-aksi="config-fac-baru"]'); await d.page.waitForSelector('#cfBaruSimpan');
  await d.page.fill('#cfBaruNama', 'RDC Uji Tiga'); await d.page.fill('#cfBaruKode', 'RDC-UJI3'); await d.page.fill('#cfBaruSs', 'Cycle Count RDC Uji Tiga');
  await d.page.click('#cfBaruSumber [data-v="tempel"]'); await d.page.fill('#cfBaruTeks', 'Z01.001.1\nZ01.001.2');
  await d.page.evaluate(() => { const b = document.getElementById('cfBaruSimpan'); b.click(); b.click(); b.click(); });
  await d.page.waitForSelector('#cfBaruSimpan', { state: 'detached', timeout: 60000 });
  await tunggu(500);
  cek('Buat facility diketuk tiga kali: satu permintaan dan satu facility baru', d.panggilan.filter((p) => p.action === 'tambahFacility').length === 1 && be.api('getDaftarFacility', '999004.dev').facilities.filter((f) => f.kode === 'RDC-UJI3').length === 1, { permintaan: d.panggilan.filter((p) => p.action === 'tambahFacility').length });
  await d.tutup();
});

/* ------------------------------------------------------------------ */
await bagian('8. Akun dicabut, tema, pembaruan, backend lama, pindahan dari versi 1', async () => {
  // akun dinonaktifkan saat aplikasi sedang dipakai
  let s = await buka({ be });
  await masuk(s.page, PETUGAS2);
  const users = be.api('getDaftarUserMaster', ADMIN).users, u = users.find((x) => x.username === PETUGAS2);
  be.api('setStatusUserMaster', ADMIN, u.rowIndex, u.username, 'Nonaktif');
  await s.page.click('[data-aksi="home-segar"]');
  await s.page.waitForSelector('#scr-login.is-on', { timeout: 15000 });
  cek('akun yang dinonaktifkan dikeluarkan dengan penjelasan', /tidak lagi terdaftar/.test(await teks(s.page, '#lgPesan')) && (await s.page.evaluate(() => localStorage.getItem('ct.sesi'))) === null);
  be.api('setStatusUserMaster', ADMIN, u.rowIndex, u.username, 'Aktif');
  await s.tutup();

  // tema gelap tersimpan
  s = await buka({ be });
  await masuk(s.page, ADMIN);
  await s.page.click('[data-aksi="akun"]'); await s.page.click('[data-aksi="akun-tema"]');
  cek('tema gelap aktif', (await s.page.evaluate(() => document.documentElement.getAttribute('data-tema'))) === 'gelap');
  await s.page.reload(); await s.page.waitForSelector('#scr-home.is-on');
  cek('tema gelap tetap setelah dibuka ulang', (await s.page.evaluate(() => document.documentElement.getAttribute('data-tema'))) === 'gelap');
  await s.tutup();

  // pembaruan tersedia
  s = await buka({ be, conf: { version: '2.0.0' }, latest: { version: '2.1.0', notes: '- Perbaikan A\n- Perbaikan B\n\n## Catatan developer\nabc', apk: 'CycleTransaksi-v2.1.0.apk', apkSize: 5000000 } });
  await masuk(s.page, ADMIN);
  await s.page.waitForSelector('[data-aksi="pembaruan-buka"]', { timeout: 15000 });
  cek('kartu versi baru muncul di Home', /Versi 2\.1\.0 tersedia/.test(await teks(s.page, '#hmPembaruan')));
  await s.page.click('[data-aksi="pembaruan-buka"]'); await s.page.waitForSelector('.upd__catatan');
  cek('catatan rilis tampil tanpa bagian developer', (await s.page.$$('.upd__catatan li')).length === 2);
  cek('lembar pembaruan tidak menyebut server', !/server/i.test(await teks(s.page, '.lembar')), await teks(s.page, '.lembar'));
  // "Periksa pembaruan" diketuk dua kali selagi pemeriksaan berjalan: satu permintaan, satu lembar
  await kembali(s.page); await s.page.waitForSelector('.lembar-wadah', { state: 'detached' });
  await s.page.click('[data-aksi="akun"]'); await s.page.waitForSelector('#akVersi');
  const sebelumCek = s.atur.rilisDiminta, gCek = gerbang(); s.atur.tundaRilis = gCek.janji;
  await s.page.click('.lembar [data-aksi="cek-versi"]');
  cek('baris versi menunjukkan sedang memeriksa', /Memeriksa/.test(await teks(s.page, '#akVersi')), await teks(s.page, '#akVersi'));
  await s.page.click('.lembar [data-aksi="cek-versi"]');
  await sampai(() => s.atur.rilisDiminta > sebelumCek); await tunggu(250); // beri waktu permintaan kedua (bila ada) ikut tercatat
  const diminta = s.atur.rilisDiminta - sebelumCek;
  gCek.buka(); s.atur.tundaRilis = 0;
  await s.page.waitForSelector('.upd__catatan', { timeout: 15000 });
  await tunggu(400);
  cek('ketuk dua kali: hanya satu pemeriksaan dan satu lembar pembaruan', diminta === 1 && s.atur.rilisDiminta - sebelumCek === 1 && (await s.page.$$('.upd__catatan')).length === 1 && (await s.page.$$('.lembar-wadah.is-on')).length === 1, { diminta, lembar: (await s.page.$$('.lembar-wadah.is-on')).length });

  // jawaban pemeriksaan yang datang terlambat tidak boleh menutup lembar lain yang sedang dipakai
  await kembali(s.page); await s.page.waitForSelector('.lembar-wadah', { state: 'detached' });
  await s.page.click('[data-aksi="akun"]'); await s.page.waitForSelector('#akVersi');
  const gTelat = gerbang(); s.atur.tundaRilis = gTelat.janji;
  await s.page.click('.lembar [data-aksi="cek-versi"]');
  await kembali(s.page); await s.page.waitForSelector('.lembar-wadah', { state: 'detached' });
  await s.page.click('[data-aksi="akun"]'); await s.page.waitForSelector('#akVersi');
  cek('lembar Akun yang dibuka lagi selagi memeriksa tetap menunjukkan sedang memeriksa', /Memeriksa/.test(await teks(s.page, '#akVersi')), await teks(s.page, '#akVersi'));
  await kembali(s.page); await s.page.waitForSelector('.lembar-wadah', { state: 'detached' });
  await s.page.click('.kpi--jingga'); await s.page.waitForSelector('.lembar .orang');
  gTelat.buka(); s.atur.tundaRilis = 0;
  await s.page.waitForFunction(() => /Versi 2\.1\.0 tersedia/.test((document.querySelector('.toast.is-on') || {}).textContent || '') || !!document.querySelector('.upd__catatan'), null, { timeout: 15000 });
  await tunggu(300);
  cek('jawaban terlambat: lembar yang sedang dibuka tetap terbuka, pembaruan ditawarkan lewat toast', !!(await s.page.$('.lembar .orang')) && !(await s.page.$('.upd__catatan')) && !!(await s.page.$('.toast.is-on .toast__aksi')), { lembarLain: !!(await s.page.$('.lembar .orang')), lembarPembaruan: !!(await s.page.$('.upd__catatan')) });
  if (await s.page.$('.toast.is-on .toast__aksi')) { await s.page.click('.toast__aksi'); await s.page.waitForSelector('.upd__catatan'); cek('toast membuka lembar pembaruan', true); }
  await s.tutup();

  // akun dipindah ke facility lain selagi masih ada hasil yang belum terkirim di HP: saat aplikasi dibuka lagi,
  // hasil itu tidak boleh dikirim (nomor tugasnya milik facility lama), dan pengguna diberi tahu
  {
    const r = be.api('getUserRole', PETUGAS2), tugas = be.api('getMyPendingTasks', PETUGAS2)[0];
    const sesi = { username: r.displayName, role: r.role, akses: !!r.aksesSettingOverflow, facilityId: r.facilityId, facilityName: r.facilityName, facilityCode: r.facilityCode || '', facilityStatus: r.facilityStatus || '' };
    const entri = { id: 'uji1', jenis: 'cycle', kunci: tugas.no, fac: r.facilityId, args: [tugas.no, r.displayName, 5], info: { lokasi: tugas.lokasi, article: tugas.article, description: tugas.description, qty: 5, sys: tugas.qtySystem }, at: Date.now() - 60000, kirimAt: Date.now() - 1000, coba: 1, status: 'antre' };
    s = await buka({ be, penyimpanan: { 'ct.sesi': JSON.stringify(sesi), 'ct.antrean': JSON.stringify([entri]) },
      ubah: (req) => { if (req.action !== 'getUserRole') return undefined; const jw = be.post(req); if (jw.ok && jw.result) { jw.result.facilityId = 'FAC-LAIN'; jw.result.facilityName = 'RDC Lain'; } return jw; } });
    await s.page.waitForFunction(() => /dipindahkan admin/.test((document.querySelector('.toast.is-on') || {}).textContent || ''), null, { timeout: 15000 }).catch(() => {});
    const pesanPindah = await s.page.evaluate(() => (document.querySelector('.toast.is-on') || {}).textContent || '');
    await tunggu(1200); // beri kesempatan antrean berjalan (bila hasil itu keliru masih dikirim)
    cek('pindah facility: hasil lama tidak dikirim, dibuang dari antrean, dan pengguna diberi tahu jumlahnya',
      /dipindahkan admin ke facility lain\. 1 hasil yang belum terkirim tidak berlaku lagi/.test(pesanPindah) && (await s.page.evaluate(() => window.CT.Antrean.daftar.length)) === 0 && !s.panggilan.some((p) => p.action === 'submitCount') && /RDC Lain/.test(await teks(s.page, '.hm-tempat')) && dataCount().find((x) => x[0] === tugas.no)[10] !== 'Selesai',
      { pesanPindah, antre: await s.page.evaluate(() => window.CT.Antrean.daftar.length), kirim: s.panggilan.filter((p) => p.action === 'submitCount').length });
    await s.tutup();
  }

  // server menolak hasil karena akun dipindah selagi aplikasi terbuka: akun langsung dicek dan aplikasi masuk ulang
  // ke facility baru (tanpa ini setiap hasil berikutnya ikut ditolak sampai aplikasi ditutup)
  {
    s = await buka({ be });
    await masuk(s.page, PETUGAS2);
    await s.page.click('#tab-cycle'); await s.page.waitForSelector('#scr-cycle .list');
    await s.page.click('[data-aksi="tugas-mulai"]'); await s.page.waitForSelector('#fkPad');
    s.atur.ubah = (req) => {
      if (req.action === 'submitCount') return { ok: true, result: { success: false, kode: 'FACILITY_BERUBAH', message: 'Facility akun ini berubah sebelum hasil terkirim. Hitung ulang dari daftar tugas yang baru.' } };
      if (req.action === 'getUserRole') { const jw = be.post(req); if (jw.ok && jw.result) { jw.result.facilityId = 'FAC-LAIN'; jw.result.facilityName = 'RDC Lain'; } return jw; }
      return undefined;
    };
    await s.page.click('[data-k="5"]'); await s.page.click('#fkSimpan');
    await s.page.evaluate(() => window.CT.Antrean.lepas());
    await s.page.waitForFunction(() => document.getElementById('scr-home').classList.contains('is-on') && /RDC Lain/.test((document.querySelector('.hm-tempat') || {}).textContent || ''), null, { timeout: 15000 }).catch(() => {});
    cek('ditolak karena pindah facility selagi aplikasi terbuka: akun langsung dicek dan aplikasi masuk ulang', /RDC Lain/.test(await teks(s.page, '.hm-tempat').catch(() => '')) && (await s.page.evaluate(() => JSON.parse(localStorage.getItem('ct.sesi')).facilityId)) === 'FAC-LAIN' && (await s.page.evaluate(() => window.CT.Antrean.daftar.length)) === 0);
    s.atur.ubah = null;
    await s.tutup();
  }

  // peran diubah admin selagi aplikasi tertutup: saat dibuka lagi menu dan data mengikuti peran baru
  s = await buka({ be });
  await masuk(s.page, PETUGAS2);
  cek('sebelum diubah: menu petugas', (await s.page.$$('.nav__item')).length === 2);
  const up = be.api('getDaftarUserMaster', ADMIN).users.find((x) => x.username === PETUGAS2);
  be.api('updateRoleUserMaster', ADMIN, up.rowIndex, up.username, 'inventory');
  await s.page.reload();
  await s.page.waitForFunction(() => document.querySelectorAll('.nav__item').length === 4, null, { timeout: 15000 });
  await s.page.waitForFunction(() => /diperbarui admin/.test((document.querySelector('.toast.is-on') || {}).textContent || ''), null, { timeout: 15000 });
  cek('peran diubah admin: menu bertambah dan pengguna diberi tahu', (await s.page.$$eval('.nav__item span:not(.nav__lencana)', (n) => n.map((x) => x.textContent))).join(',') === 'Home,Cycle,Validasi,Verifikasi' && (await s.page.evaluate(() => JSON.parse(localStorage.getItem('ct.sesi')).role)) === 'inventory');
  be.api('updateRoleUserMaster', ADMIN, up.rowIndex, up.username, up.role);
  await s.tutup();
  s = await buka({ be, conf: { version: '2.1.0' }, latest: { version: '2.1.0', notes: '', apk: 'x.apk', apkSize: 1 } });
  await masuk(s.page, ADMIN); await tunggu(2500);
  cek('tidak ada kartu pembaruan bila sudah versi terbaru', !(await s.page.$('[data-aksi="pembaruan-buka"]')));
  await s.tutup();

  // backend lama (belum punya fungsi baru): Home tetap tampil
  s = await buka({ be, ubah: (req) => (['getHomeBundle', 'getUploadFormData', 'getProductivity'].includes(req.action) ? { ok: false, error: 'Fungsi tidak dikenal: ' + req.action } : undefined) });
  await masuk(s.page, ADMIN);
  const hs = be.api('getHomeSummary', ADMIN);
  cek('backend lama: Home dirakit dari fungsi lama', angka((await s.page.$$eval('.kpi__nilai', (n) => n.map((x) => x.textContent)))[0]) === hs.totalCycleHariIni);
  await s.page.click('#tab-upload'); await s.page.waitForSelector('.up-berkas');
  cek('backend lama: Upload tetap bisa dibuka', (await s.page.$$('[data-aksi="upload-peran"]')).length > 0);
  await s.page.click('#tab-report'); await s.page.waitForSelector('#rpTren rect');
  await s.page.click('[data-aksi="report-tampilan"][data-v="productivity"]'); await s.page.waitForSelector('#rpIsi .kosong');
  cek('backend lama: Productivity menjelaskan laporan belum tersedia, tanpa istilah teknis', /Belum tersedia/.test(await teks(s.page, '#rpIsi .kosong')) && /admin aplikasi memperbarui sistem/.test(await teks(s.page, '#rpIsi .kosong')) && !/server|Code\.gs|backend/i.test(await teks(s.page, '#rpIsi .kosong')), await teks(s.page, '#rpIsi .kosong'));
  await s.tutup();

  // pindahan dari versi 1.x: alamat server & NIK yang tersimpan tetap dipakai
  s = await buka({ be, remote: '', penyimpanan: { 'ct.server': JSON.stringify(EXEC + '?x=1'), 'ct.user': JSON.stringify(PETUGAS) } });
  await s.page.waitForSelector('#scr-home.is-on .hm-tugas, #scr-home.is-on .kosong', { timeout: 15000 });
  cek('pengguna versi 1 langsung masuk tanpa mengetik ulang', (await teks(s.page, '.hm-nama')) === PETUGAS && (await s.page.evaluate(() => localStorage.getItem('ct.user'))) === null);
  await s.tutup();
});

/* ------------------------------------------------------------------ */
await bagian('9. Muat di layar kecil dan tema gelap', async () => {
  for (const [w, h] of [[360, 568], [360, 640], [412, 915]]) {
    const s = await buka({ be, w, h });
    await masuk(s.page, PETUGAS);
    const lebarHome = await s.page.evaluate(() => document.documentElement.scrollWidth - window.innerWidth);
    await s.page.click('#tab-cycle'); await s.page.waitForSelector('#scr-cycle .list');
    await s.page.click('[data-aksi="tugas-mulai"]'); await s.page.waitForSelector('#fkPad');
    await s.page.click('[data-k="4"]'); await s.page.click('#fkSimpan'); await s.page.waitForSelector('.simpanan');
    const u = await s.page.evaluate(() => ({ tinggi: document.documentElement.scrollHeight, vh: window.innerHeight, pad: document.getElementById('fkPad').getBoundingClientRect().bottom, lebar: document.documentElement.scrollWidth - window.innerWidth }));
    cek(w + '×' + h + ': mode fokus muat tanpa menggulir, keypad terlihat utuh', u.tinggi <= u.vh && u.pad <= u.vh && u.lebar <= 0 && lebarHome <= 0, u);
    await s.page.click('[data-aksi="fokus-urung"]'); await tunggu(100);
    await s.tutup();
  }
  const s = await buka({ be, penyimpanan: { 'ct.tema': JSON.stringify('gelap') } });
  await masuk(s.page, ADMIN);
  const warna = await s.page.evaluate(() => ({ latar: getComputedStyle(document.body).backgroundColor, teks: getComputedStyle(document.body).color }));
  cek('tema gelap dipakai sejak aplikasi dibuka', warna.latar === 'rgb(15, 17, 21)', warna);
  await s.tutup();
});

/* ------------------------------------------------------------------ */
await bagian('10. Panel admin di browser (halaman Web App): Upload, Facility, Config', async () => {
  const s = await buka({ be, jalur: '/panel', w: 1280, h: 800 });
  await s.page.waitForFunction(() => /^Tersambung/.test(document.getElementById('lgStatus').textContent.trim()), null, { timeout: 15000 });
  cek('layar masuk panel menjelaskan isinya', /Panel admin/.test(await teks(s.page, '.lg__ket')) && !(await s.page.$('[data-aksi="cek-versi"]')));
  await s.page.fill('#lgNik', PETUGAS); await s.page.click('#lgTombol');
  await s.page.waitForFunction(() => document.getElementById('lgPesan').textContent.length > 0);
  cek('petugas ditolak dengan penjelasan', /khusus admin/.test(await teks(s.page, '#lgPesan')) && (await s.page.evaluate(() => localStorage.getItem('ct.sesi'))) === null, await teks(s.page, '#lgPesan'));
  await s.page.fill('#lgNik', ADMIN); await s.page.click('#lgTombol');
  await s.page.waitForSelector('#scr-upload.is-on .up-berkas', { timeout: 15000 });
  cek('menu panel: Upload, Config, Akun', (await s.page.$$eval('.nav__item span', (n) => n.map((x) => x.textContent))).join(',') === 'Upload,Config,Akun');
  // upload dari file .csv di komputer
  const lok = be.rows('Lokasi_Aktif').map((r) => r[0]), LP = lok.filter((l) => /\.1$/.test(String(l)))[41]; // Level 1: tidak butuh alat bantu
  const fT = path.join(TMP, 'panel-transaksi.csv'), fS = path.join(TMP, 'panel-stok.csv');
  fs.writeFileSync(fT, 'Type,Article,Description,From Location,To Location,Qty,AddWho\nPicking,77700001,BARANG PANEL,' + LP + ',,1,wms.uji\n');
  fs.writeFileSync(fS, 'Location,Article,Qty\n' + LP + ',77700001,9\n');
  await s.page.setInputFiles('input[data-berkas="transaksi"]', fT);
  await s.page.setInputFiles('input[data-berkas="stok"]', fS);
  await s.page.waitForFunction(() => document.querySelectorAll('.up-berkas.is-ok').length === 2, null, { timeout: 15000 });
  if (!(await s.page.$eval('.bag small', (n) => /^[1-9]/.test(n.textContent)))) await s.page.click('[data-aksi="upload-peran"][data-peran="outbound"]');
  const sebelum = dataCount().length;
  await s.page.click('#upProses'); await s.page.waitForSelector('.up-hasil', { timeout: 30000 });
  cek('upload dari panel masuk ke database yang sama dengan aplikasi', dataCount().length === sebelum + 1 && dataCount().some((r) => String(r[4]) === '77700001'), { tambah: dataCount().length - sebelum, lembar: await teks(s.page, '.lembar'), kirim: s.panggilan.filter((p) => p.action === 'importRawData').pop() });
  await s.page.click('.lembar__kaki [data-tutup]'); await s.page.waitForSelector('.lembar-wadah', { state: 'detached' });
  // config dan facility
  await s.page.click('#tab-config'); await s.page.waitForSelector('#cfUsers .list', { timeout: 15000 });
  cek('Config tampil sebagai menu (tanpa tombol kembali)', !(await s.page.$('#scr-config [data-aksi="kembali"]')));
  await s.page.click('[data-aksi="config-tab"][data-v="facility"]'); await s.page.waitForSelector('.cf-fac__db');
  cek('Facility: tautan ke database Google Sheets tiap gudang', /^https:\/\/docs\.google\.com\/spreadsheets\/d\//.test(await s.page.$eval('.cf-fac__db', (n) => n.href)));
  await s.page.click('#tab-akun'); await s.page.click('[data-aksi="akun-keluar"]');
  await s.page.waitForSelector('#scr-login.is-on');
  cek('keluar dari panel kembali ke layar masuk', true);
  await s.tutup();
});

await browser.close(); server.close();
fs.rmSync(TMP, { recursive: true, force: true });
const unik = [...new Set(semuaGalat)];
console.log('\n' + (unik.length ? '✗ Galat JavaScript selama uji:\n  ' + unik.join('\n  ') : '✓ Tidak ada galat JavaScript selama uji'));
console.log(`\n${lulus} lulus, ${gagal} gagal`);
process.exit(gagal || unik.length ? 1 : 0);
