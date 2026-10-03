// Menggabungkan backend/src/*.gs menjadi satu file backend/Code.gs yang siap ditempel ke editor
// Apps Script (komentar dibuang supaya ringan; kode tidak diubah atau dipendekkan).
// Jalankan: npm run backend
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { minify } from 'terser';
import { panelHtml } from './panel.mjs';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const SRC = path.join(ROOT, 'backend', 'src');
const files = fs.readdirSync(SRC).filter((f) => f.endsWith('.gs')).sort();
const all = files.map((f) => fs.readFileSync(path.join(SRC, f), 'utf8')).join('\n');
const out = await minify(all, { compress: false, mangle: false, format: { beautify: true, indent_level: 2, comments: false, max_line_len: 160, quote_style: 3 } });
const version = (all.match(/const APP_VERSION = '([^']+)'/) || [])[1] || '';
const head = `/**
 * Cycle Transaksi - backend Apps Script ${version} (satu file, gabungan ${files.length} file di backend/src).
 * JANGAN diedit di sini: ubah file di backend/src lalu jalankan "npm run backend".
 *
 * PASANG PERTAMA KALI:
 * 1. Di spreadsheet: Extensions > Apps Script. Hapus isi Code.gs, tempel SELURUH file ini, simpan.
 * 2. Pilih fungsi "setupAwal" di bilah atas, klik Run, izinkan akses. Semua sheet, facility,
 *    admin pertama, dan trigger antrean dibuat otomatis.
 * 3. Deploy > New deployment > Web app. Execute as: Me. Who has access: Anyone. Kirim URL /exec
 *    ke pengembang aplikasi: alamat itu ditanam di aplikasi (app/server.json), bukan diisi pengguna.
 *
 * PERBARUI BACKEND: tempel ulang file ini, lalu Deploy > Manage deployments > Edit > Version: New version.
 * URL /exec tidak berubah, jadi aplikasi tidak perlu disetel ulang.
 */
`;
// Panel admin ditanam sebagai teks: Code.gs tetap satu file yang cukup ditempel.
const panel = await panelHtml();
let body = head + out.code + '\nconst PANEL_HTML_ = ' + JSON.stringify(panel).replace(/\u2028/g, '\\u2028').replace(/\u2029/g, '\\u2029') + ';\n';
// Salinan pribadi (tidak masuk repo): SETUP_NAMA, SETUP_KODE, SETUP_NIK + OUT=<path>
const { SETUP_NAMA, SETUP_KODE, SETUP_NIK, OUT } = process.env;
if (OUT) {
  if (!SETUP_NAMA || !SETUP_KODE || !SETUP_NIK) { console.error('OUT butuh SETUP_NAMA, SETUP_KODE, SETUP_NIK'); process.exit(1); }
  const q = (v) => JSON.stringify(String(v)).slice(1, -1).replace(/'/g, "\\'");
  const n = body.split("'ISI NAMA FACILITY'").length + body.split("'ISI-KODE'").length + body.split("'ISI.NIK.ADMIN'").length - 3;
  if (n !== 3) { console.error('Placeholder setup tidak ditemukan tepat sekali (' + n + ')'); process.exit(1); }
  body = body.replace("'ISI NAMA FACILITY'", "'" + q(SETUP_NAMA) + "'").replace("'ISI-KODE'", "'" + q(SETUP_KODE) + "'").replace("'ISI.NIK.ADMIN'", "'" + q(SETUP_NIK) + "'");
  fs.writeFileSync(OUT, body);
  console.log('Salinan pribadi ditulis ke ' + OUT);
} else {
  fs.writeFileSync(path.join(ROOT, 'backend', 'Code.gs'), body);
}
console.log(`backend/Code.gs: ${files.length} file, ${(Buffer.byteLength(head + out.code) / 1024).toFixed(0)} KB (sumber ${(Buffer.byteLength(all) / 1024).toFixed(0)} KB)`);
