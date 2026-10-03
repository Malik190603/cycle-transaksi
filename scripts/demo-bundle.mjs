// Menyusun "backend demo": backend Apps Script yang asli + tiruan layanan Google + data contoh,
// dibungkus jadi satu skrip browser (www/demo/demo-backend.js). Dipakai mode demo di aplikasi dan
// pratinjau tampilan; tidak pernah dimuat kecuali pengguna memilih mode demo.
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { minify } from 'terser';
import { DEMO } from '../demo/seed.mjs';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const read = (p) => fs.readFileSync(path.join(ROOT, p), 'utf8');

export async function demoBundle() {
  const SRC = path.join(ROOT, 'backend', 'src');
  let backend = fs.readdirSync(SRC).filter((f) => f.endsWith('.gs')).sort().map((f) => fs.readFileSync(path.join(SRC, f), 'utf8')).join('\n');
  const isi = (cari, nilai) => { if (!backend.includes(cari)) throw new Error('Placeholder ' + cari + ' tidak ditemukan'); backend = backend.replace(cari, JSON.stringify(nilai)); };
  isi("'ISI NAMA FACILITY'", DEMO.facility); isi("'ISI-KODE'", DEMO.kode); isi("'ISI.NIK.ADMIN'", DEMO.admin);
  // apiBridge_ mencari fungsi lewat globalThis; di dalam pembungkus, globalThis diganti objek lokal
  // yang berisi semua fungsi tingkat atas backend.
  const names = [...backend.matchAll(/^function\s+([A-Za-z0-9_$]+)\s*\(/gm)].map((m) => m[1]);
  const mock = read('test/gas-mock.mjs').replace(/^export /gm, '');
  const seed = read('demo/seed.mjs').replace(/^export /gm, '');
  const code = `(function () {
${mock}
${seed}
function buatBackend(G) {
  var SpreadsheetApp = G.SpreadsheetApp, CacheService = G.CacheService, LockService = G.LockService, PropertiesService = G.PropertiesService,
    Utilities = G.Utilities, ScriptApp = G.ScriptApp, ContentService = G.ContentService, HtmlService = G.HtmlService, Logger = G.Logger, Session = G.Session;
  var globalThis = {};
${backend}
${[...new Set(names)].map((n) => 'globalThis.' + n + ' = ' + n + ';').join('\n')}
  return globalThis;
}
window.CT_DEMO = {
  buat: function () {
    var gas = createGas();
    var B = buatBackend(gas.globals);
    var post = function (action, args) { return JSON.parse(B.doPost({ postData: { contents: JSON.stringify({ action: action, args: args }) } }).getContent()); };
    var api = function (action) { var j = post(action, Array.prototype.slice.call(arguments, 1)); if (!j.ok) throw new Error(action + ': ' + j.error); return j.result; };
    var info = seedDemo({ api: api, fn: function (n) { return B[n].apply(null, Array.prototype.slice.call(arguments, 1)); }, sheet: function (n) { return gas.active.getSheetByName(n); } });
    // Di server asli antrean diproses trigger tiap menit; di demo tiap beberapa detik supaya hasilnya cepat terlihat.
    var timer = setInterval(function () { try { B.workerSemua(); } catch (e) { console.error(e); } }, 4000);
    return {
      info: info,
      kirim: function (nama, args) {
        return new Promise(function (ok, gagal) {
          setTimeout(function () { try { ok(post(nama, args)); } catch (e) { gagal(e); } }, 140 + Math.random() * 240);
        });
      },
      proses: function () { B.workerSemua(); },
      // dua "file" contoh untuk mencoba layar Upload Data tanpa file sungguhan
      contoh: function () {
        var lok = gas.active.getSheetByName('Lokasi_Aktif'), L = lok.getRange(2, 1, lok.getLastRow() - 1, 1).getValues().map(function (r) { return r[0]; });
        var dc = gas.active.getSheetByName('Data Count'), D = dc.getRange(2, 1, Math.min(300, dc.getLastRow() - 1), 16).getValues();
        var transaksi = [], stok = [];
        for (var i = 0; i < 80; i++) {
          var d = D[(i * 7) % D.length], l = L[Math.floor(Math.random() * L.length)], move = i % 2 === 1;
          transaksi.push({ type: move ? 'Move' : 'Picking', article: String(d[4]), description: String(d[5]), lokasiAwal: move ? 'STAGING' : l, lokasiTujuan: move ? l : '', qty: String(1 + (i % 5)), addWho: 'wms.demo' });
          stok.push({ lokasi: l, article: String(d[4]), qty: String(3 + ((i * 5) % 30)) });
        }
        transaksi.push({ type: 'Adjustment', article: String(D[0][4]), description: String(D[0][5]), lokasiAwal: L[0], lokasiTujuan: '', qty: '1', addWho: 'wms.demo' });
        return { transaksi: transaksi, stok: stok };
      },
      hentikan: function () { clearInterval(timer); }
    };
  }
};
})();`;
  const out = await minify(code, { compress: { passes: 1 }, mangle: true, format: { comments: false } });
  return '/* Cycle Transaksi: backend demo (dibuat otomatis oleh scripts/demo-bundle.mjs) */\n' + out.code;
}
