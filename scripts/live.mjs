// Update kilat (sama dengan Active Coach dan IMM).
//   node scripts/live.mjs base  → tentukan basis native rilis ini (NATIVE_BASE, LIVE_HASH ke $GITHUB_ENV)
//   node scripts/live.mjs pack  → zip isi web dari proyek Android + tulis info ke release-notes.md
// Basis native = versi APK terakhir yang mengubah bagian native (plugin, izin, ikon, konfigurasi Capacitor,
// kunci tanda tangan). Paket web hanya dipasang di HP yang basis native-nya sama; selain itu HP diarahkan
// mengunduh APK di dalam aplikasi.
import { execSync } from 'node:child_process';
import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const mode = process.argv[2];
const VERSION = process.env.APP_VERSION || JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8')).version;
const REPO = process.env.GITHUB_REPOSITORY || 'Malik190603/cycle-transaksi';
const env = (k, v) => { console.log(k + '=' + v); if (process.env.GITHUB_ENV) fs.appendFileSync(process.env.GITHUB_ENV, k + '=' + v + '\n'); };
const out = (k, v) => { if (process.env.GITHUB_OUTPUT) fs.appendFileSync(process.env.GITHUB_OUTPUT, k + '=' + v + '\n'); };
const walk = (d) => fs.existsSync(d) ? fs.readdirSync(d, { withFileTypes: true }).sort((a, b) => a.name.localeCompare(b.name)).flatMap((e) => e.isDirectory() ? walk(path.join(d, e.name)) : [path.join(d, e.name)]) : [];

function nativeFingerprint() {
  const h = crypto.createHash('sha256');
  const pkg = JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8'));
  h.update('deps:' + JSON.stringify(Object.entries(pkg.dependencies || {}).sort()));
  const files = ['capacitor.config.json', 'android/app/build.gradle', 'android/build.gradle', 'android/variables.gradle', 'android/app/cycle-transaksi.keystore', 'android/app/src/main/AndroidManifest.xml'];
  for (const f of walk(path.join(ROOT, 'android/app/src/main/res'))) files.push(path.relative(ROOT, f));
  for (const f of walk(path.join(ROOT, 'android/app/src/main/java'))) files.push(path.relative(ROOT, f));
  for (const f of files) { const p = path.join(ROOT, f); h.update('\n' + f + ':'); if (fs.existsSync(p) && fs.statSync(p).isFile()) h.update(fs.readFileSync(p)); }
  return h.digest('hex').slice(0, 24);
}
const marker = (body) => { const m = String(body || '').match(/<!--\s*ct-live\s+(\{[\s\S]*?\})\s*-->/); try { return m ? JSON.parse(m[1]) : null; } catch { return null; } };

async function base() {
  const hash = nativeFingerprint();
  let prev = null;
  try {
    const headers = { Accept: 'application/vnd.github+json', 'User-Agent': 'cycle-transaksi-ci' };
    if (process.env.GITHUB_TOKEN) headers.Authorization = 'Bearer ' + process.env.GITHUB_TOKEN;
    const r = await fetch('https://api.github.com/repos/' + REPO + '/releases?per_page=15', { headers });
    if (r.ok) for (const rel of await r.json()) { if (rel.draft) continue; const m = marker(rel.body); if (m) { prev = { ...m, tag: rel.tag_name }; break; } }
    else console.log('GitHub API ' + r.status + ' — anggap rilis native');
  } catch (e) { console.log('Tidak bisa membaca rilis sebelumnya: ' + e.message); }
  const same = prev && prev.hash === hash && prev.native;
  const nb = same ? prev.native : VERSION;
  console.log(same ? `Tanpa perubahan native sejak ${prev.tag} → update kilat (basis v${nb})` : `Perubahan native (atau rilis pertama) → APK baru jadi basis v${nb}`);
  env('NATIVE_BASE', nb); env('LIVE_HASH', hash); out('native', same ? 'false' : 'true'); out('base', nb);
}

function pack() {
  const pub = path.join(ROOT, 'android/app/src/main/assets/public');
  if (!fs.existsSync(path.join(pub, 'index.html'))) { console.error('assets/public belum ada — jalankan npx cap sync android'); process.exit(1); }
  const cfg = fs.readFileSync(path.join(pub, 'config.js'), 'utf8');
  if (!cfg.includes('"version": "' + VERSION + '"')) { console.error('config.js tidak berisi versi ' + VERSION); process.exit(1); }
  const name = 'CycleTransaksi-web-' + VERSION + '.zip', zip = path.join(ROOT, name);
  fs.rmSync(zip, { force: true });
  execSync('zip -q -r -9 -X ' + JSON.stringify(zip) + ' .', { cwd: pub, stdio: 'inherit' });
  const buf = fs.readFileSync(zip), sha256 = crypto.createHash('sha256').update(buf).digest('hex');
  const info = { v: VERSION, native: process.env.NATIVE_BASE || VERSION, hash: process.env.LIVE_HASH || nativeFingerprint(), asset: name, sha256, size: buf.length };
  const notesPath = path.join(ROOT, 'release-notes.md');
  const notes = fs.existsSync(notesPath) ? fs.readFileSync(notesPath, 'utf8') : '';
  fs.appendFileSync(notesPath, '\n<!-- ct-live ' + JSON.stringify(info) + ' -->\n');
  // latest.json: dibaca aplikasi lewat releases/latest/download/ (bukan GitHub API, jadi tanpa batas 60 permintaan/jam per IP)
  const apk = 'CycleTransaksi-v' + VERSION + '.apk', apkPath = path.join(ROOT, apk);
  if (!fs.existsSync(apkPath)) { console.error(apk + ' belum ada'); process.exit(1); }
  fs.writeFileSync(path.join(ROOT, 'latest.json'), JSON.stringify({ version: VERSION, notes, apk, apkSize: fs.statSync(apkPath).size, live: { native: info.native, sha256, size: buf.length, asset: name } }, null, 2));
  console.log('Paket web ' + name + ' · ' + (buf.length / 1048576).toFixed(2) + ' MB · basis native v' + info.native);
}

if (mode === 'base') await base();
else if (mode === 'pack') pack();
else { console.error('Pakai: node scripts/live.mjs base|pack'); process.exit(1); }
