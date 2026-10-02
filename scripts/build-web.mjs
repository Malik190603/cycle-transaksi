// Menyusun folder www/ untuk Capacitor dari file tampilan Apps Script di src/.
// - src/Index.html + semua <?!= include('X'); ?> digabung jadi www/index.html
// - google.script.run diganti jembatan (app/bridge.js) yang memanggil doPost Web App
// - SheetJS dibundel lokal (bukan CDN) supaya unggah/ekspor Excel tidak bergantung jaringan pihak ketiga
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const SRC = path.join(ROOT, 'src'), APP = path.join(ROOT, 'app'), OUT = path.join(ROOT, 'www');
const pkg = JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8'));
const VERSION = process.env.APP_VERSION || 'dev';
const read = (p) => fs.readFileSync(p, 'utf8');
const need = (p) => { if (!fs.existsSync(p)) { console.error('File tidak ditemukan: ' + path.relative(ROOT, p)); process.exit(1); } return p; };

// Alamat Web App: variabel GAS_URL di GitHub Actions, atau app/server.json; boleh kosong (diisi di layar masuk).
let serverUrl = (process.env.GAS_URL || '').trim();
if (!serverUrl && fs.existsSync(path.join(APP, 'server.json'))) serverUrl = String(JSON.parse(read(path.join(APP, 'server.json'))).url || '').trim();
if (serverUrl && !/^https:\/\/script\.google\.com\/.+\/exec$/.test(serverUrl)) { console.error('GAS_URL harus URL Web App Apps Script yang berakhiran /exec'); process.exit(1); }

fs.rmSync(OUT, { recursive: true, force: true });
fs.mkdirSync(path.join(OUT, 'vendor'), { recursive: true });

let html = read(need(path.join(SRC, 'Index.html')));
let includes = 0;
html = html.replace(/<\?!=\s*include\('([A-Za-z0-9_]+)'\);?\s*\?>/g, (_, name) => { includes++; return read(need(path.join(SRC, name + '.html'))); });
if (/<\?/.test(html)) { console.error('Masih ada scriptlet Apps Script (<? … ?>) yang belum diganti'); process.exit(1); }

const cdn = /<script src="https:\/\/cdnjs\.cloudflare\.com\/ajax\/libs\/xlsx\/[^"]+"><\/script>/;
if (!cdn.test(html)) { console.error('Tag SheetJS tidak ditemukan di Index.html'); process.exit(1); }
html = html.replace(cdn, '<script src="vendor/xlsx.full.min.js"></script>\n  <script src="config.js"></script>\n  <script src="bridge.js"></script>\n  <link rel="stylesheet" href="bridge.css">');
html = html.replace('<meta name="viewport" content="width=device-width, initial-scale=1">', '<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n  <meta charset="utf-8">\n  <title>Cycle Transaksi</title>');
html = html.replace(/\s*<base target="_top">/, '');
// bridge.css harus menimpa Style.html → pindahkan tautannya ke akhir <head>
html = html.replace('\n  <link rel="stylesheet" href="bridge.css">', '').replace('</head>', '  <link rel="stylesheet" href="bridge.css">\n</head>');

fs.writeFileSync(path.join(OUT, 'index.html'), html);
fs.copyFileSync(need(path.join(APP, 'bridge.js')), path.join(OUT, 'bridge.js'));
fs.copyFileSync(need(path.join(APP, 'bridge.css')), path.join(OUT, 'bridge.css'));
fs.copyFileSync(need(path.join(ROOT, 'node_modules', 'xlsx', 'dist', 'xlsx.full.min.js')), path.join(OUT, 'vendor', 'xlsx.full.min.js'));

const conf = { version: VERSION, updateRepo: process.env.GITHUB_REPOSITORY || pkg.updateRepo || '', nativeBase: process.env.NATIVE_BASE || VERSION, build: Number(process.env.VERSION_CODE || process.env.GITHUB_RUN_NUMBER || 0), serverUrl };
fs.writeFileSync(path.join(OUT, 'config.js'), '// Dibuat otomatis saat build\nwindow.CT_CONFIG = ' + JSON.stringify(conf, null, 2) + ';\n');
console.log(`www/ siap: ${includes} bagian digabung, versi ${VERSION}, server ${serverUrl ? 'tertanam' : 'diisi di layar masuk'}`);
