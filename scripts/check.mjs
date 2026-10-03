// Pemeriksaan sebelum rilis (dijalankan di GitHub Actions setelah build-web). Gagal = tidak ada rilis.
//   1) berkas www/ lengkap dan sintaksnya benar
//   2) aplikasi tidak memuat apa pun dari luar (CDN) dan tidak punya layar pengaturan server
//   3) tiap aksi tombol punya penangan, tiap ikon ada, tiap fungsi server yang dipanggil memang ada di backend
//   4) konfigurasi Capacitor & izin Android
//   5) tidak ada kunci rahasia ikut ke APK
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const WWW = path.join(ROOT, 'www');
let bad = 0;
const fail = (m) => { bad++; console.error('✗ ' + m); };
const ok = (m) => console.log('✓ ' + m);
const baca = (p) => fs.readFileSync(path.join(ROOT, p), 'utf8');
const ada = (p) => fs.existsSync(path.join(ROOT, p));

if (!ada('www/index.html')) { console.error('www/index.html tidak ada. Jalankan npm run build'); process.exit(1); }

/* ---------- 1. berkas & sintaks ---------- */
const WAJIB = ['index.html', 'app.css', 'app.js', 'config.js', 'vendor/xlsx.full.min.js', 'demo/demo-backend.js',
  'fonts/barlow-400.woff2', 'fonts/barlow-500.woff2', 'fonts/barlow-600.woff2', 'fonts/barlow-700.woff2', 'fonts/barlow-condensed-600.woff2', 'fonts/barlow-condensed-700.woff2'];
const hilang = WAJIB.filter((f) => !ada('www/' + f) || fs.statSync(path.join(WWW, f)).size === 0);
if (hilang.length) fail('berkas www/ tidak lengkap: ' + hilang.join(', ')); else ok(WAJIB.length + ' berkas www/ lengkap');

const html = baca('www/index.html'), js = ada('www/app.js') ? baca('www/app.js') : '', css = ada('www/app.css') ? baca('www/app.css') : '', confJs = ada('www/config.js') ? baca('www/config.js') : '';
for (const f of ['app.js', 'config.js', 'demo/demo-backend.js']) {
  if (!ada('www/' + f)) continue;
  try { new vm.Script(baca('www/' + f), { filename: f }); ok('sintaks ' + f); } catch (e) { fail(f + ': ' + e.message); }
}
let n = 0, m;
const reInline = /<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g;
while ((m = reInline.exec(html))) { n++; try { new vm.Script(m[1], { filename: 'index.html#' + n }); } catch (e) { fail('index.html skrip ke-' + n + ': ' + e.message); } }
const urut = ['config.js', 'app.js'].map((s) => html.indexOf('src="' + s + '"'));
if (urut.some((i) => i < 0) || urut[0] > urut[1]) fail('index.html harus memuat config.js lalu app.js'); else ok('index.html memuat config.js lalu app.js');
if (!/href="app\.css"/.test(html)) fail('index.html tidak memuat app.css');
const fontDipakai = [...css.matchAll(/url\(['"]?(fonts\/[^'")]+)['"]?\)/g)].map((x) => x[1]);
const fontHilang = fontDipakai.filter((f) => !ada('www/' + f));
if (!fontDipakai.length || fontHilang.length) fail('font yang dirujuk app.css tidak ada: ' + (fontHilang.join(', ') || 'tidak ada @font-face')); else ok(fontDipakai.length + ' font dimuat dari dalam aplikasi');

/* ---------- 2. tanpa CDN, tanpa pengaturan server ---------- */
const luar = [...(html + css).matchAll(/(?:src|href)=["'](https?:)?\/\/[^"']+|url\(\s*['"]?https?:[^)]+/gi)].map((x) => x[0]).filter((u) => !/www\.w3\.org/.test(u));
if (/<\?/.test(html) || luar.length) fail('index.html atau app.css memuat sumber dari luar: ' + luar.join(', ')); else ok('tidak ada CDN atau sumber luar');
const HOST_BOLEH = ['script.google.com', 'raw.githubusercontent.com', 'github.com', 'docs.google.com', 'www.w3.org'];
const hostJs = [...new Set([...js.matchAll(/https?:\/\/([a-z0-9.-]+)/gi)].map((x) => x[1].toLowerCase()))].filter((h) => HOST_BOLEH.indexOf(h) < 0);
if (hostJs.length) fail('app.js menyebut host yang tidak dikenal: ' + hostJs.join(', ')); else ok('app.js hanya berbicara dengan Apps Script dan GitHub');
// Alamat server ditanam saat build / diambil dari repo; pengguna tidak pernah mengisinya (uji tampilan memeriksa teks layarnya).
if (/type="url"|placeholder="https:\/\/script\.google|id="[^"]*server[^"]*"|>\s*Atur server/i.test(js + html)) fail('aplikasi tidak boleh punya isian alamat server'); else ok('tidak ada isian alamat server');

let conf = null;
try { conf = JSON.parse(confJs.slice(confJs.indexOf('{'), confJs.lastIndexOf('}') + 1)); } catch (e) { fail('config.js tidak terbaca: ' + e.message); }
if (conf) {
  if (!/^(dev|\d+\.\d+\.\d+)$/.test(conf.version)) fail('config.js: versi tidak sah (' + conf.version + ')');
  if (conf.serverUrl && !/^https:\/\/script\.google\.com\/(?:a\/macros\/[^/]+|macros)\/s\/[A-Za-z0-9_-]+\/exec$/.test(conf.serverUrl)) fail('config.js: serverUrl bukan URL Web App Apps Script');
  if (conf.demoSaja && process.env.GITHUB_ACTIONS) fail('build rilis tidak boleh memakai --demo-saja');
  if (!conf.updateRepo) fail('config.js: updateRepo kosong (pembaruan & alamat server tidak bisa diambil)');
  ok('config.js: versi ' + conf.version + ', server ' + (conf.serverUrl ? 'tertanam' : 'diambil dari app/server.json di repo'));
}
try {
  const sj = JSON.parse(baca('app/server.json')), u = String(sj.url || '');
  if (u && !/^https:\/\/script\.google\.com\/.+\/exec$/.test(u)) fail('app/server.json: url harus URL Web App yang berakhiran /exec'); else ok('app/server.json ' + (u ? 'berisi alamat server' : 'masih kosong (aplikasi menawarkan mode demo)'));
} catch (e) { fail('app/server.json: ' + e.message); }

/* ---------- 3. aksi, ikon, fungsi server ---------- */
const aksiDipakai = new Set([...js.matchAll(/data-aksi="([a-z0-9-]+)"/g)].map((x) => x[1]));
// aksi yang namanya diteruskan sebagai teks (kpi(), gagalMuat(), kartu "berikutnya")
[...js.matchAll(/'((?:home|tugas|verif|upload|report|config|akun|fokus|pembaruan)-[a-z0-9-]+)'/g)].forEach((x) => aksiDipakai.add(x[1]));
const aksiAda = new Set([...js.matchAll(/Aksi\['([a-z0-9-]+)'\]\s*=/g)].map((x) => x[1]));
['cycle', 'validasi'].forEach((mode) => { if (/Aksi\['tugas-segar-' \+ mode\]/.test(js)) aksiAda.add('tugas-segar-' + mode); });
const aksiYatim = [...aksiDipakai].filter((a) => !aksiAda.has(a) && !/-$/.test(a));
if (aksiYatim.length) fail('tombol tanpa penangan aksi: ' + aksiYatim.join(', ')); else ok(aksiDipakai.size + ' aksi tombol punya penangan');

const blokIkon = js.slice(js.indexOf('const IKON = {'), js.indexOf('function ikon('));
const ikonAda = new Set([...blokIkon.matchAll(/^\s*([a-z0-9_]+):\s*'/gm)].map((x) => x[1]));
const ikonDipakai = new Set([...js.matchAll(/\bikon\('([a-z0-9_]+)'/g)].map((x) => x[1]));
[...js.matchAll(/ikon:\s*'([a-z0-9_]+)'/g)].forEach((x) => ikonDipakai.add(x[1]));
const ikonHilang = [...ikonDipakai].filter((i) => !ikonAda.has(i));
if (!ikonAda.size || ikonHilang.length) fail('ikon tidak ada di IKON: ' + ikonHilang.join(', ')); else ok(ikonDipakai.size + ' ikon tersedia');

const srcGs = fs.readdirSync(path.join(ROOT, 'backend', 'src')).filter((f) => f.endsWith('.gs')).map((f) => baca('backend/src/' + f)).join('\n');
const izin = new Set([...(srcGs.match(/const API_BRIDGE_ALLOW_ = \[([\s\S]*?)\];/) || ['', ''])[1].matchAll(/'([A-Za-z0-9_]+)'/g)].map((x) => x[1]));
const dipanggil = new Set([...js.matchAll(/panggil\('([A-Za-z0-9_]+)'/g)].map((x) => x[1]));
[...js.matchAll(/fungsi:\s*'([A-Za-z0-9_]+)'/g), ...js.matchAll(/(?:cycle|validasi):\s*'(submit[A-Za-z]+)'/g)].forEach((x) => dipanggil.add(x[1]));
const takDiizinkan = [...dipanggil].filter((f) => !izin.has(f));
const takAda = [...izin].filter((f) => !new RegExp('^function\\s+' + f + '\\s*\\(', 'm').test(srcGs));
if (takDiizinkan.length) fail('aplikasi memanggil fungsi yang tidak ada di daftar izin ApiBridge.gs: ' + takDiizinkan.join(', '));
else if (takAda.length) fail('daftar izin ApiBridge.gs menyebut fungsi yang tidak ada di backend: ' + takAda.join(', '));
else ok(dipanggil.size + ' fungsi server yang dipanggil aplikasi tersedia di backend');
if (ada('backend/Code.gs')) {
  const satu = baca('backend/Code.gs');
  const kurang = [...dipanggil].filter((f) => !new RegExp('^function\\s+' + f + '\\s*\\(', 'm').test(satu));
  if (kurang.length) fail('backend/Code.gs belum memuat: ' + kurang.join(', ') + '. Jalankan npm run backend'); else ok('backend/Code.gs (satu file) sudah memuat semuanya');
}

/* ---------- 4. Capacitor & Android ---------- */
for (const f of ['package.json', 'capacitor.config.json']) { try { JSON.parse(baca(f)); } catch (e) { fail(f + ': ' + e.message); } }
const cap = JSON.parse(baca('capacitor.config.json'));
if (!(cap.plugins && cap.plugins.CapacitorHttp && cap.plugins.CapacitorHttp.enabled)) fail('CapacitorHttp harus aktif (tanpa itu panggilan ke Apps Script diblokir CORS)'); else ok('CapacitorHttp aktif');
const cu = (cap.plugins || {}).CapacitorUpdater;
if (!cu || cu.autoUpdate !== false || cu.statsUrl !== '' || cu.updateUrl !== '' || cu.channelUrl !== '') fail('CapacitorUpdater harus mode manual tanpa server pihak ketiga'); else ok('update kilat: mode manual, tanpa statistik pihak ketiga');
if (!baca('android/app/src/main/AndroidManifest.xml').includes('REQUEST_INSTALL_PACKAGES')) fail('izin REQUEST_INSTALL_PACKAGES tidak ada'); else ok('izin pasang update');
if (!/notifyAppReady\(\)/.test(js)) fail('app.js harus memanggil notifyAppReady() (tanpa itu update kilat dibatalkan otomatis)'); else ok('update kilat dikonfirmasi saat aplikasi siap');

/* ---------- 5. kunci rahasia ---------- */
const semua = html + css + js + confJs;
const POLA = [[/-----BEGIN (?:RSA |EC )?PRIVATE KEY-----/, 'private key'], [/AIza[0-9A-Za-z_-]{30,}/, 'Google API key'], [/gh[pousr]_[0-9A-Za-z]{30,}/, 'GitHub token'], [/ya29\.[0-9A-Za-z_-]{20,}/, 'token OAuth Google']];
const bocor = POLA.filter(([p]) => p.test(semua)).map(([, nama]) => nama);
if (bocor.length) fail('kunci rahasia ikut ke aplikasi: ' + bocor.join(', ')); else ok('tidak ada kunci rahasia di www/');

if (bad) { console.error(bad + ' masalah. Rilis dibatalkan.'); process.exit(1); }
console.log('Semua pemeriksaan lolos');
