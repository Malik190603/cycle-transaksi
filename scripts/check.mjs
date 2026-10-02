// Pemeriksaan sebelum rilis (dijalankan di GitHub Actions setelah build-web). Gagal = tidak ada rilis.
// 1) sintaks semua skrip  2) konfigurasi valid  3) jembatan terpasang  4) tidak ada kunci rahasia ikut ke APK
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
let bad = 0;
const fail = (m) => { bad++; console.error('✗ ' + m); };
const ok = (m) => console.log('✓ ' + m);

const htmlPath = path.join(ROOT, 'www', 'index.html');
if (!fs.existsSync(htmlPath)) { console.error('www/index.html tidak ada — jalankan npm run build'); process.exit(1); }
const html = fs.readFileSync(htmlPath, 'utf8');
const re = /<script>([\s\S]*?)<\/script>/g;
let m, n = 0;
while ((m = re.exec(html))) {
  n++;
  try { new vm.Script(m[1], { filename: 'index.html#script' + n }); }
  catch (e) { fail('index.html skrip ke-' + n + ': ' + e.message); }
}
if (!n) fail('tidak ada skrip di index.html'); else ok(n + ' skrip di index.html');
for (const f of ['bridge.js', 'config.js', 'vendor/xlsx.full.min.js', 'bridge.css']) {
  const p = path.join(ROOT, 'www', f);
  if (!fs.existsSync(p)) { fail('www/' + f + ' tidak ada'); continue; }
  if (f.endsWith('.js') && !f.startsWith('vendor/')) { try { new vm.Script(fs.readFileSync(p, 'utf8'), { filename: f }); ok(f); } catch (e) { fail(f + ': ' + e.message); } }
}
if (/<\?/.test(html) || /cdnjs\.cloudflare\.com/.test(html)) fail('index.html masih memuat scriptlet Apps Script atau CDN'); else ok('tanpa scriptlet & tanpa CDN');
const order = ['vendor/xlsx.full.min.js', 'config.js', 'bridge.js'].map((s) => html.indexOf('src="' + s + '"'));
if (order.some((i) => i < 0) || order[0] > order[1] || order[1] > order[2] || order[2] > html.indexOf('<script>')) fail('urutan skrip salah: jembatan harus dimuat sebelum skrip aplikasi'); else ok('jembatan dimuat sebelum skrip aplikasi');
if (!/google\.script\.run/.test(html)) fail('tidak ada pemanggilan google.script.run — sumber tampilan salah?');
for (const f of ['package.json', 'capacitor.config.json']) {
  try { JSON.parse(fs.readFileSync(path.join(ROOT, f), 'utf8')); ok(f); } catch (e) { fail(f + ': ' + e.message); }
}
const cap = JSON.parse(fs.readFileSync(path.join(ROOT, 'capacitor.config.json'), 'utf8'));
if (!(cap.plugins && cap.plugins.CapacitorHttp && cap.plugins.CapacitorHttp.enabled)) fail('CapacitorHttp harus aktif (tanpa itu panggilan ke Apps Script diblokir CORS)'); else ok('CapacitorHttp aktif');
const cu = (cap.plugins || {}).CapacitorUpdater;
if (!cu || cu.autoUpdate !== false || cu.statsUrl !== '' || cu.updateUrl !== '' || cu.channelUrl !== '') fail('CapacitorUpdater harus mode manual tanpa server pihak ketiga'); else ok('update kilat: mode manual, tanpa statistik pihak ketiga');
if (!fs.readFileSync(path.join(ROOT, 'android/app/src/main/AndroidManifest.xml'), 'utf8').includes('REQUEST_INSTALL_PACKAGES')) fail('izin REQUEST_INSTALL_PACKAGES tidak ada'); else ok('izin pasang update');
const all = html + fs.readFileSync(path.join(ROOT, 'www', 'bridge.js'), 'utf8');
const PATTERNS = [[/-----BEGIN (?:RSA )?PRIVATE KEY-----/, 'private key'], [/AIza[0-9A-Za-z_-]{30,}/, 'Google API key'], [/ghp_[0-9A-Za-z]{30,}/, 'GitHub token']];
for (const [p, name] of PATTERNS) if (p.test(all)) fail('kunci rahasia ikut ke aplikasi: ' + name);
ok('tidak ada kunci rahasia di www/');
if (bad) { console.error(bad + ' masalah — rilis dibatalkan'); process.exit(1); }
console.log('Semua pemeriksaan lolos');
