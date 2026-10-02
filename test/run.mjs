// Uji jembatan di Chromium dengan backend tiruan (meniru doPost + ApiBridge.gs).
// Jalankan: npm run build && npm test   (butuh paket playwright terpasang di mesin ini)
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';
const require = createRequire(import.meta.url);
let chromium;
try { ({ chromium } = require('playwright')); } catch { ({ chromium } = require('/opt/npm-tools/node_modules/playwright')); }

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..'), WWW = path.join(ROOT, 'www'), TMP = path.join(ROOT, 'test', '.tmp');
fs.mkdirSync(TMP, { recursive: true });
const MIME = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css' };
const server = http.createServer((q, r) => {
  const f = path.join(WWW, q.url === '/' ? 'index.html' : q.url.split('?')[0]);
  if (!f.startsWith(WWW) || !fs.existsSync(f)) { r.writeHead(404); return r.end(); }
  r.writeHead(200, { 'Content-Type': MIME[path.extname(f)] || 'application/octet-stream' }); r.end(fs.readFileSync(f));
}).listen(0);
const ORIGIN = 'http://127.0.0.1:' + server.address().port, EXEC = 'https://script.google.com/macros/s/TEST_deploy-123/exec';

// ---------- backend tiruan ----------
const calls = [];
let pending = [
  { no: 1, lokasi: 'A03.23.02', article: 'SKU-001', description: 'Barang Satu', area: 'A', levelGroupKey: 'tanpa_alat', levelGroupLabel: 'Tanpa Alat Bantu', qtySystem: 10 },
  { no: 2, lokasi: 'A03.24.01', article: 'SKU-002', description: 'Barang Dua', area: 'A', levelGroupKey: 'tanpa_alat', levelGroupLabel: 'Tanpa Alat Bantu', qtySystem: 5 }
];
const API = {
  getAppVersion: () => 'v8.29.3',
  getUserRole: (u) => String(u).toLowerCase() === '1001.tes' ? { role: 'inventory', displayName: '1001.TES', aksesSettingOverflow: false, facilityId: 'FAC-1', facilityName: 'RDC Uji' } : null,
  getHomeSummary: () => ({ totalCycleHariIni: 7, selesaiHariIni: 5, belumValidasi: 2, lastUpdated: '02/10/2026 09:00', version: 'v8.29.3' }),
  getMyPendingCount: () => pending.length,
  getPendingValidasiCount: () => 2,
  getPlusMinusSummary: () => ({ plus: { totalSku: 1, totalQty: 3, items: [{ article: 'SKU-9', description: 'Plus', qty: 3, lokasi: 'B01.01.01', jumlahLokasi: 1 }] }, minus: { totalSku: 0, totalQty: 0, items: [] } }),
  getMyPendingTasks: () => pending,
  submitCount: (no, nama, qty) => { pending = pending.filter((p) => p.no !== no); return { success: true, queued: true, message: 'Submission diterima.' }; },
  getPendingValidasi: () => { throw new Error('Akses ditolak untuk role "outbound".'); }
};
function doPost(body) {
  const req = JSON.parse(body); calls.push(req);
  if (!Array.isArray(req.args)) return { success: false, message: 'Unknown action: ' + req.action };
  const fn = API[req.action]; if (!fn) return { ok: false, error: 'Fungsi tidak dikenal: ' + req.action };
  try { const result = fn(...req.args); return { ok: true, result: result === undefined ? null : result }; } catch (e) { return { ok: false, error: e.message }; }
}

let failed = 0;
const check = (name, cond, extra) => { if (!cond) failed++; console.log((cond ? '✓ ' : '✗ ') + name + (cond || !extra ? '' : ' → ' + extra)); };

const browser = await chromium.launch();
const ctx = await browser.newContext({ viewport: { width: 390, height: 844 }, deviceScaleFactor: 2 });
const page = await ctx.newPage();
const errors = []; page.on('pageerror', (e) => errors.push(e.message));
await ctx.route(EXEC, (route) => {
  const q = route.request();
  if (q.method() !== 'POST') return route.fulfill({ status: 405, body: '' });
  route.fulfill({ status: 200, contentType: 'application/json', headers: { 'Access-Control-Allow-Origin': '*' }, body: JSON.stringify(doPost(q.postData())) });
});

// 1) Buka pertama kali: belum ada server → kotak "Alamat server" terbuka
await page.goto(ORIGIN + '/');
check('layar masuk tampil', await page.isVisible('#screenLogin'));
check('kotak alamat server terbuka saat belum diatur', await page.isVisible('#ctServerBox'));
await page.fill('#ctServerInput', 'https://contoh.com/exec'); await page.click('#ctServerSave');
check('URL bukan Apps Script ditolak', /harus berbentuk/.test(await page.textContent('#ctServerMsg')));
await page.fill('#ctServerInput', EXEC + '?usp=sharing'); await page.click('#ctServerSave');
await page.waitForFunction(() => /Terhubung/.test(document.getElementById('ctServerMsg').textContent));
check('uji koneksi server berhasil', true);
await page.waitForSelector('#ctServerBox', { state: 'hidden' });
await page.screenshot({ path: path.join(TMP, '1-login.png') });

// 2) NIK tidak terdaftar, lalu NIK benar
await page.fill('#loginUsername', 'tidak.ada'); await page.click('#btnLogin');
await page.waitForFunction(() => /tidak terdaftar/.test(document.getElementById('loginStatus').textContent));
check('NIK tidak terdaftar ditolak', true);
await page.fill('#loginUsername', '1001.tes'); await page.click('#btnLogin');
await page.waitForSelector('#screenHome:not(.hidden)');
await page.waitForFunction(() => document.getElementById('hkTotal').textContent === '7');
check('masuk → Home dengan KPI dari server', (await page.textContent('#homeUsername')) === '1001.TES' && (await page.textContent('#hkSelesai')) === '5');
await page.waitForFunction(() => document.getElementById('badgeCycle').textContent === '2');
check('badge tugas terisi', true);
check('versi aplikasi & tombol pembaruan di footer', /Aplikasi vdev/.test(await page.textContent('#ctFoot')));
await page.screenshot({ path: path.join(TMP, '2-home.png'), fullPage: true });

// 3) Cycle Transaksi: argumen dikirim berurutan seperti google.script.run
await page.click('#screenHome button.home-btn >> nth=0');
await page.waitForSelector('#screenCount:not(.hidden)');
check('item pertama tampil', /SKU-001/.test(await page.textContent('#articleText')), await page.textContent('#articleText'));
await page.screenshot({ path: path.join(TMP, '3-count.png') });
await page.fill('#qtyInput', '10'); await page.click('#btnSubmit');
await page.waitForFunction(() => /SKU-002/.test(document.getElementById('articleText').textContent), null, { timeout: 8000 });
const sub = calls.find((c) => c.action === 'submitCount');
check('submitCount(no, nama, qty) terkirim dengan urutan benar', sub && sub.args[0] === 1 && sub.args[1] === '1001.TES' && Number(sub.args[2]) === 10, JSON.stringify(sub && sub.args));

// 4) Error dari server sampai ke penangan kegagalan
const err = await page.evaluate(() => new Promise((res) => google.script.run.withSuccessHandler(() => res('sukses?')).withFailureHandler((e) => res(e.message)).getPendingValidasi('1001.TES')));
check('pesan error server diteruskan', /Akses ditolak/.test(err), err);
const dyn = await page.evaluate(() => new Promise((res) => { const n = 'getAppVersion'; google.script.run.withSuccessHandler(res).withFailureHandler((e) => res('ERR ' + e.message))[n](); }));
check('pemanggilan dengan nama dinamis (cincin progres) jalan', dyn === 'v8.29.3', dyn);
const unk = await page.evaluate(() => window.CT_BRIDGE.call('resetCount', []).then(() => 'lolos', (e) => e.message));
check('fungsi di luar daftar izin ditolak', /tidak dikenal/.test(unk), unk);

// 5) Sesi diingat: muat ulang → langsung Home; keluar → kembali ke layar masuk
await page.reload(); await page.waitForSelector('#screenHome:not(.hidden)');
check('muat ulang tetap masuk (NIK diingat)', (await page.textContent('#homeUsername')) === '1001.TES');
await page.click('#screenHome button.link'); await page.reload();
check('setelah keluar, muat ulang kembali ke layar masuk', await page.isVisible('#screenLogin') && !(await page.isVisible('#screenHome')));

// 6) Backend lama (tanpa ApiBridge) → pesan yang jelas
await ctx.unroute(EXEC);
await ctx.route(EXEC, (route) => route.fulfill({ status: 200, contentType: 'application/json', headers: { 'Access-Control-Allow-Origin': '*' }, body: JSON.stringify({ success: false, message: 'Unknown action' }) }));
const old = await page.evaluate(() => window.CT_BRIDGE.call('getAppVersion', []).then(() => 'lolos', (e) => e.message));
check('backend tanpa ApiBridge dikenali', /ApiBridge/.test(old), old);
await ctx.unroute(EXEC); await ctx.route(EXEC, (route) => route.abort('internetdisconnected'));
const off = await page.evaluate(() => window.CT_BRIDGE.call('getAppVersion', []).then(() => 'lolos', (e) => e.message));
check('tanpa internet → pesan koneksi', /koneksi internet/.test(off), off);

check('tanpa error JavaScript di halaman', errors.length === 0, errors.join(' | '));
await browser.close(); server.close();
console.log(failed ? failed + ' pemeriksaan gagal' : 'Semua uji lolos');
process.exit(failed ? 1 : 0);
