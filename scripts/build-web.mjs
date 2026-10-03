// Menyusun folder www/ untuk Capacitor dari src/.
//   src/index.html   → www/index.html
//   src/css/*.css    → www/app.css   (digabung berurutan nama)
//   src/js/*.js      → www/app.js    (digabung berurutan nama, dibungkus satu fungsi)
//   font Barlow, SheetJS, dan backend demo disalin lokal: aplikasi tidak memuat apa pun dari CDN.
// Opsi: --demo-saja  → build pratinjau yang selalu memakai data contoh (tanpa server).
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { demoBundle } from './demo-bundle.mjs';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const SRC = path.join(ROOT, 'src'), APP = path.join(ROOT, 'app'), OUT = path.join(ROOT, 'www');
const pkg = JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8'));
const VERSION = process.env.APP_VERSION || 'dev';
const DEMO_SAJA = process.argv.includes('--demo-saja');
const read = (p) => fs.readFileSync(p, 'utf8');
const need = (p) => { if (!fs.existsSync(p)) { console.error('File tidak ditemukan: ' + path.relative(ROOT, p)); process.exit(1); } return p; };
const daftar = (dir, ext) => fs.readdirSync(need(path.join(SRC, dir))).filter((f) => f.endsWith(ext)).sort();

// Alamat Web App: variabel GAS_URL di GitHub Actions, atau app/server.json. Aplikasi juga mengambil
// app/server.json terbaru dari repo saat dibuka, jadi alamat bisa diganti tanpa memasang ulang APK.
let serverUrl = (process.env.GAS_URL || '').trim();
if (!serverUrl && fs.existsSync(path.join(APP, 'server.json'))) serverUrl = String(JSON.parse(read(path.join(APP, 'server.json'))).url || '').trim();
if (serverUrl && !/^https:\/\/script\.google\.com\/.+\/exec$/.test(serverUrl)) { console.error('Alamat server harus URL Web App Apps Script yang berakhiran /exec'); process.exit(1); }

fs.rmSync(OUT, { recursive: true, force: true });
for (const d of ['vendor', 'fonts', 'demo']) fs.mkdirSync(path.join(OUT, d), { recursive: true });

const css = daftar('css', '.css').map((f) => '/* ' + f + ' */\n' + read(path.join(SRC, 'css', f))).join('\n');
const jsFiles = daftar('js', '.js');
const js = '/* Cycle Transaksi ' + VERSION + ' */\n(function () {\n\'use strict\';\n' + jsFiles.map((f) => '\n/* ===== ' + f + ' ===== */\n' + read(path.join(SRC, 'js', f))).join('\n') + '\n})();\n';
fs.writeFileSync(path.join(OUT, 'app.css'), css);
fs.writeFileSync(path.join(OUT, 'app.js'), js);
fs.copyFileSync(need(path.join(SRC, 'index.html')), path.join(OUT, 'index.html'));

const FONT = [
  ['@fontsource/barlow/files/barlow-latin-400-normal.woff2', 'barlow-400.woff2'],
  ['@fontsource/barlow/files/barlow-latin-500-normal.woff2', 'barlow-500.woff2'],
  ['@fontsource/barlow/files/barlow-latin-600-normal.woff2', 'barlow-600.woff2'],
  ['@fontsource/barlow/files/barlow-latin-700-normal.woff2', 'barlow-700.woff2'],
  ['@fontsource/barlow-condensed/files/barlow-condensed-latin-600-normal.woff2', 'barlow-condensed-600.woff2'],
  ['@fontsource/barlow-condensed/files/barlow-condensed-latin-700-normal.woff2', 'barlow-condensed-700.woff2']
];
for (const [dari, ke] of FONT) fs.copyFileSync(need(path.join(ROOT, 'node_modules', dari)), path.join(OUT, 'fonts', ke));
fs.copyFileSync(need(path.join(ROOT, 'node_modules', '@fontsource/barlow/LICENSE')), path.join(OUT, 'fonts', 'LISENSI-Barlow.txt')); // SIL Open Font License: lisensi ikut bersama font
fs.copyFileSync(need(path.join(ROOT, 'node_modules', 'xlsx', 'dist', 'xlsx.full.min.js')), path.join(OUT, 'vendor', 'xlsx.full.min.js'));

const demo = await demoBundle();
fs.writeFileSync(path.join(OUT, 'demo', 'demo-backend.js'), demo);

const conf = {
  version: VERSION, updateRepo: process.env.GITHUB_REPOSITORY || pkg.updateRepo || '', nativeBase: process.env.NATIVE_BASE || VERSION,
  build: Number(process.env.VERSION_CODE || process.env.GITHUB_RUN_NUMBER || 0), serverUrl, demoSaja: DEMO_SAJA
};
fs.writeFileSync(path.join(OUT, 'config.js'), '// Dibuat otomatis saat build\nwindow.CT_CONFIG = ' + JSON.stringify(conf, null, 2) + ';\n');
const kb = (p) => (fs.statSync(path.join(OUT, p)).size / 1024).toFixed(0) + ' KB';
console.log(`www/ siap: versi ${VERSION}, ${jsFiles.length} modul JS (${kb('app.js')}), CSS ${kb('app.css')}, demo ${kb('demo/demo-backend.js')}, server ${serverUrl ? 'tertanam' : 'diambil dari repo saat dibuka'}${DEMO_SAJA ? ', MODE DEMO SAJA' : ''}`);
