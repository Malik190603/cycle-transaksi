// Uji ujung-ke-ujung backend Apps Script di Node (layanan Google ditiru test/gas-mock.mjs).
// Semua panggilan lewat doPost dengan format yang sama seperti yang dikirim APK.
// Jalankan: node test/backend.mjs
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';
import { createGas } from './gas-mock.mjs';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const SRC = path.join(ROOT, 'backend', 'src');
const single = process.argv.includes('--single'); // uji backend/Code.gs (satu file siap tempel)
let code = single ? fs.readFileSync(path.join(ROOT, 'backend', 'Code.gs'), 'utf8')
  : fs.readdirSync(SRC).filter((f) => f.endsWith('.gs')).sort().map((f) => fs.readFileSync(path.join(SRC, f), 'utf8')).join('\n');
// Repo menyimpan placeholder; uji memakai nilai contoh.
const belumDiisi = vm.createContext({ ...createGas().globals }); vm.runInContext(code, belumDiisi);
let tolak = ''; try { belumDiisi.setupAwal(); } catch (e) { tolak = e.message; }
code = code.replace(/'ISI NAMA FACILITY'/, "'RDC Tallo Makassar'").replace(/'ISI-KODE'/, "'RDC-MKS'").replace(/'ISI\.NIK\.ADMIN'/, "'9001.admin'");
const gas = createGas();
const ctx = vm.createContext({ ...gas.globals });
vm.runInContext(code, ctx, { filename: single ? 'Code.gs' : 'backend.js' });

let failed = 0;
const check = (name, cond, extra) => { if (!cond) failed++; console.log((cond ? '✓ ' : '✗ ') + name + (cond || extra === undefined ? '' : ' → ' + (typeof extra === 'string' ? extra : JSON.stringify(extra)))); };
const raw = (action, ...args) => JSON.parse(ctx.doPost({ postData: { contents: JSON.stringify({ action, args }) } }).getContent());
const api = (action, ...args) => { const j = raw(action, ...args); if (!j.ok) throw new Error(action + ': ' + j.error); return j.result; };
const sheet = (n) => gas.active.getSheetByName(n);
const rows = (n) => { const s = sheet(n); const lr = s.getLastRow(); return lr < 2 ? [] : s.getRange(2, 1, lr - 1, Math.max(1, s.getLastColumn())).getValues(); };
const worker = () => ctx.workerSemua();
const isDate = (v) => Object.prototype.toString.call(v) === '[object Date]';
const ADMIN = '9001.admin';

// ---------- 1. Setup awal ----------
check('setupAwal menolak jalan bila konstanta belum diisi', /Isi dulu/.test(tolak), tolak);
ctx.setupAwal(); ctx.setupAwal(); // dua kali: harus idempoten
const names = gas.active.getSheets().map((s) => s.getName());
const wajib = ['Master', 'Master_User', 'Facility', 'User_Facility_Assignment', 'Config_Akses_Setting', 'Config_Sistem', 'Config_Level_Assignment', 'Log_Perubahan_Config', 'Log_Sistem',
  'Data Count', 'Riwayat', 'Riwayat_Archive', 'Lokasi_Aktif', 'Ringkasan_Harian', 'Antrian_Aktif', 'Queue_Ringkasan_Delta', 'Queue_Antrian_Counter', 'Queue_Submit_Cycle', 'Queue_Validasi', 'Queue_Task_Investigasi',
  'Log_Akses', 'Log_Anomali', 'Log_Unassigned', 'Log_Bukti_Investigasi'];
check('setupAwal membuat ' + wajib.length + ' sheet', wajib.every((n) => names.includes(n)) && !names.includes('Sheet1'), names);
check('setupAwal idempoten (1 facility, 1 user, 1 trigger)', rows('Facility').length === 1 && rows('Master_User').length === 1 && rows('Config_Akses_Setting').length === 1 && gas.triggers.length === 1 && gas.triggers[0].fn === 'workerSemua', [rows('Facility').length, rows('Master_User').length, gas.triggers.length]);
check('header Data Count & Riwayat sesuai', sheet('Data Count').getRange(1, 1, 1, 16).getValues()[0][15] === 'AddWho_Transaksi' && sheet('Riwayat').getRange(1, 1, 1, 27).getValues()[0][26] === 'Pasangan_Task_ID');
const FAC = rows('Facility')[0][0];

// ---------- 2. Login & user ----------
check('server menjawab versi', /^v\d/.test(api('getAppVersion')));
const me = api('getUserRole', ADMIN);
check('login admin: peran, facility, akses Config', me && me.role === 'admin' && me.facilityId === FAC && me.facilityName === 'RDC Tallo Makassar' && me.aksesSettingOverflow === true, me);
check('NIK tidak terdaftar → null', api('getUserRole', 'tidak.ada') === null);
for (const [u, r] of [['1001.storing', 'storing'], ['1002.out', 'outbound'], ['1003.out', 'outbound'], ['1004.inv', 'inventory']]) {
  const res = api('tambahUserMaster', ADMIN, u, r); if (!res.success) check('tambah user ' + u, false, res);
}
check('4 user ditambah & otomatis masuk facility admin', rows('Master_User').length === 5 && rows('Master_User').every((r) => r[3] === FAC), rows('Master_User'));
check('user baru bisa login dengan facility', api('getUserRole', '1004.inv').facilityId === FAC);
check('non-admin ditolak di menu admin', /Akses ditolak/.test(raw('getAssignableUsers', '1002.out').error || ''));
check('fungsi di luar daftar izin ditolak', /tidak dikenal/.test(raw('resetCount').error || ''));

// ---------- 3. Lokasi aktif & unggah ----------
const LOK = ['A01.01.01', 'A01.02.03', 'A01.03.05', 'B02.04.02', 'B02.05.01', 'A03.23.K09'];
const imp = api('importLokasiAktif', ADMIN, FAC, LOK.concat(['A01.01.01', '']), true);
check('impor lokasi aktif (duplikat & kosong dibuang)', imp.success && rows('Lokasi_Aktif').length === 6, imp);
const trans = [ // [tipe, article, deskripsi, lokasi awal, lokasi tujuan, qty, addwho]
  ['Picking', 'SKU1', 'Barang 1', 'A01.01.01', '', 5, 'wmsA'],
  ['Move', 'SKU2', 'Barang 2', 'X99.99.99', 'A01.02.03', 3, 'wmsB'],
  ['Move', 'SKU3', 'Barang 3', 'X99.99.99', 'A01.03.05', 2, 'wmsB'],
  ['Picking', 'SKU4', 'Barang 4', 'B02.04.02', '', 1, 'wmsC'],
  ['Picking', 'SKU5', 'Barang 5', 'B02.05.01', '', 1, 'wmsC'],
  ['Move', 'SKU5', 'Barang 5', 'B02.04.02', 'B02.05.01', 4, 'wmsA'],
  ['Picking', 'SKU6', 'Barang 6', 'A03.23.K09', '', 1, 'wmsA'],
  ['Picking', 'SKU10', 'Barang 10', 'B02.04.02', '', 2, 'wmsC'],
  ['Move', 'SKU10', 'Barang 10', 'X99.99.99', 'B02.05.01', 2, 'wmsC'],
  ['Adjustment', 'SKU7', 'Barang 7', 'A01.01.01', '', 1, 'wmsA'],
  ['Picking', 'SKU8', 'Barang 8', 'Z09.09.09', '', 1, 'wmsA'],
  ['Picking', 'SKU9', 'Barang 9', '', '', 1, 'wmsA']
];
const stock = [['A01.01.01', 'SKU1', 10], ['A01.02.03', 'SKU2', 7], ['A01.03.05', 'SKU3', 4], ['B02.04.02', 'SKU4', 6], ['B02.05.01', 'SKU5', 8], ['A03.23.K09', 'SKU6', 2], ['B02.04.02', 'SKU10', 5], ['B02.05.01', 'SKU10', 1]];
const petugas = ['1001.storing', '1002.out', '1003.out'];
check('daftar petugas facility tanpa developer', api('getAssignableUsers', ADMIN).length === 5);
const up = api('importRawData', trans, stock, petugas, ADMIN, 1, 1, 'job_1');
check('unggah: 8 item ditugaskan, 3 jenis baris dilewati dihitung terpisah', up.success && up.ditugaskan === 8 && up.skippedOtherType === 1 && up.skippedInactive === 1 && up.skippedBlankLokasi === 1, up);
const dc = rows('Data Count');
const who = (lok, art) => (dc.find((r) => r[2] === lok && r[4] === art) || [])[9];
check('Level 5-6 → Storing; Level 3-4 → petugas tangga; rollcage K → tanpa alat', who('A01.03.05', 'SKU3') === '1001.storing' && who('A01.02.03', 'SKU2') === '1002.out' && who('A03.23.K09', 'SKU6') === '1003.out', dc.map((r) => [r[2], r[4], r[9]]));
check('satu lokasi satu orang', new Set(dc.filter((r) => r[2] === 'B02.05.01').map((r) => r[9])).size === 1 && new Set(dc.filter((r) => r[2] === 'B02.04.02').map((r) => r[9])).size === 1);
const sku5 = dc.find((r) => r[4] === 'SKU5');
check('qty transaksi dijumlah & user WMS dipisah per jenis', sku5[6] === 5 && sku5[15] === 'MOVE:wmsA;PICKING:wmsC', [sku5[6], sku5[15]]);
check('progres impor tersimpan', api('getImportProgress', 'job_1').percent >= 95);
const up2 = api('importRawData', trans, stock, petugas, ADMIN, 1, 1, 'job_2');
check('unggah ulang tidak membuat tugas ganda (digabung)', up2.success && up2.merged === 8 && rows('Data Count').length === 8 && rows('Data Count').find((r) => r[4] === 'SKU5')[6] === 10, up2);
check('Outstanding Cycle menampilkan 8 pending', api('getPendingBacklog', ADMIN).reduce((a, b) => a + b.pending, 0) === 8);
const hb0 = api('getHomeBundle', ADMIN), hbP = api('getHomeBundle', '1002.out');
check('Home (satu panggilan) admin: 8 outstanding, rincian sisa per petugas', hb0.outstanding === 8 && hb0.totalCycleHariIni === 0 && hb0.perPetugas.length === 3 && hb0.perPetugas.reduce((a, p) => a + p.sisa, 0) === 8 && hb0.plusMinus && hb0.serverTime > 0, hb0);
check('Home (satu panggilan) petugas: hanya tugasnya, tanpa data admin', hbP.myPending > 0 && hbP.myPending < 8 && hbP.outstanding === hbP.myPending && hbP.perPetugas.length === 0 && hbP.plusMinus === null, hbP);
const uf = api('getUploadFormData', ADMIN);
check('formulir Upload: petugas, alat bantu, jumlah lokasi aktif, mode', uf.users.length === 5 && uf.lokasiAktif === 6 && uf.modeAssignment === 'legacy' && uf.equipment.reachTruck === 0 && uf.facilityName === 'RDC Tallo Makassar' && /Akses ditolak/.test(raw('getUploadFormData', '1002.out').error || ''), uf);

// ---------- 4. Hitung ----------
const hitung = { SKU1: 10, SKU2: 6, SKU3: 4, SKU4: 5, SKU5: 9, SKU6: 2 }; // SKU2, SKU4, SKU5 selisih
const hitung10 = { 'B02.04.02': 3, 'B02.05.01': 3 };                      // SKU10: -2 dan +2
let tugasTotal = 0;
for (const u of petugas) {
  const t = api('getMyPendingTasks', u); tugasTotal += t.length;
  for (const it of t) { const q = it.article === 'SKU10' ? hitung10[it.lokasi] : hitung[it.article]; const r = api('submitCount', it.no, u, q); if (!r.queued) check('submit ' + it.article, false, r); }
}
check('tiap petugas hanya melihat tugasnya (total 8)', tugasTotal === 8);
check('submit hanya masuk antrean (Data Count belum berubah)', rows('Queue_Submit_Cycle').length === 8 && rows('Data Count').every((r) => r[10] === 'Pending'));
api('submitCount', dc.find((r) => r[4] === 'SKU1')[0], who('A01.01.01', 'SKU1'), 10); // submit ganda
worker();
const dc2 = rows('Data Count'), rw = rows('Riwayat');
check('worker: semua tugas Selesai, antrean kosong', dc2.every((r) => r[10] === 'Selesai') && rows('Queue_Submit_Cycle').length === 0);
check('submit ganda tidak membuat Riwayat kedua', rw.length === 8, rw.length);
check('3 HIT tanpa validasi, 5 DISCREPANCY menunggu validator bergilir', rw.filter((r) => r[10] === 'HIT' && r[11] === 'Tidak Perlu' && r[18] === 'Selesai').length === 3 && rw.filter((r) => r[10] === 'DISCREPANCY' && r[11] === 'Pending' && r[22] === '1004.inv').length === 5, rw.map((r) => [r[3], r[10], r[11], r[22]]));
check('Waktu_Input terisi', dc2.every((r) => isDate(r[14])));
check('submit ganda tercatat di Log_Sistem', rows('Log_Sistem').some((r) => /duplikat/i.test(r[2])));
check('qty tidak valid ditolak', api('submitCount', 1, petugas[0], -1).success === false);

// ---------- 5. Validasi ----------
check('petugas tidak bisa membuka validasi', /Akses ditolak/.test(raw('getPendingValidasi', '1002.out').error || ''));
const pv = api('getPendingValidasi', '1004.inv');
check('validator melihat 5 item miliknya; admin melihat semua', pv.length === 5 && api('getPendingValidasi', ADMIN).length === 5 && api('getPendingValidasiCount', '1004.inv') === 5);
const qv = { SKU2: 7, SKU4: 5, SKU5: 9 };
for (const it of pv) api('submitValidasi', it.id, '1004.inv', it.article === 'SKU10' ? hitung10[it.lokasi] : qv[it.article]);
worker();
const rw2 = rows('Riwayat'), by = (art, lok) => rw2.find((r) => r[3] === art && (!lok || r[2] === lok));
check('validasi cocok → HIT, Salah Hitung, tertutup otomatis', by('SKU2')[14] === 'HIT' && by('SKU2')[17] === 'Salah Hitung' && by('SKU2')[18] === 'Selesai', by('SKU2').slice(11, 19));
check('validasi tetap selisih → task Open dengan PIC awal', ['SKU4', 'SKU5'].every((a) => by(a)[14] === 'DISCREPANCY' && by(a)[18] === 'Open' && by(a)[23]), [by('SKU4').slice(11, 24)]);
const pm = api('getPlusMinusSummary', '1004.inv');
check('Summary Plus Minus: plus 3 qty/2 SKU, minus 3 qty/2 SKU', pm.plus.totalQty === 3 && pm.plus.totalSku === 2 && pm.minus.totalQty === 3 && pm.minus.totalSku === 2, pm);

// ---------- 6. Verifikasi ----------
const tasks = api('getOpenTasks', '1004.inv');
check('4 task aktif dengan validator & user WMS', tasks.length === 4 && tasks.every((t) => t.namaValidator === '1004.inv') && tasks.some((t) => /PICKING:wmsC/.test(t.addWhoTransaksi)), tasks.map((t) => [t.article, t.selisih, t.addWhoTransaksi]));
const form = api('getInvestigasiFormData', '1004.inv');
check('form verifikasi: 9 kategori, 4 tahap, user facility', form.kategoriList.length === 9 && form.statusFlow.length === 4 && form.allUsers.length === 5);
const T = (a, l) => tasks.find((t) => t.article === a && (!l || t.lokasi === l));
const bukti = (tipe, sku, from, to, qty) => [['1', 'HCI', tipe, sku, 'Barang', 'GRP', 'LOT1', from, 'ID1', to, 'ID2', 'SRC0001', String(qty), '2026-10-01 08:00', 'wmsC']];
const tutup = (t, kat, b) => raw('updateTaskStatus', t.id, '1004.inv', 'Selesai', 'Selesai dicek', kat, ADMIN, b, t.lokasi, t.article, t.selisih).result;
const t4 = T('SKU4'), t5 = T('SKU5');
check('tanpa bukti ditolak', tutup(t4, 'Kurang Picking', []).success === false);
check('bukti dengan qty tidak sama ditolak', /tidak sama/.test(tutup(t4, 'Kurang Picking', bukti('PICKING', 'SKU4', 'B02.04.02', '', 3)).message));
check('catatan wajib', raw('updateTaskStatus', t4.id, '1004.inv', 'Sedang Dicari', '', '', '', [], t4.lokasi, t4.article, t4.selisih).result.success === false);
check('tahap antara diterima', raw('updateTaskStatus', t4.id, '1004.inv', 'Sedang Dicari', 'Mulai dicari', '', '', [], t4.lokasi, t4.article, t4.selisih).result.queued === true);
worker();
check('status task maju ke Sedang Dicari dengan catatan bercap waktu', rows('Riwayat').find((r) => r[0] === t4.id)[18] === 'Sedang Dicari' && /Sedang Dicari\] Mulai dicari/.test(rows('Riwayat').find((r) => r[0] === t4.id)[19]));
check('bukti PICKING qty cocok diterima', tutup(t4, 'Kurang Picking', bukti('PICKING', 'SKU4', 'B02.04.02', '', 1)).queued === true);
check('Adjustment Plus dengan qty negatif ditolak', tutup(t5, 'Adjustment Plus', bukti('ADJUSTMENT', 'SKU5', '', 'B02.05.01', -1)).success === false);
check('Adjustment Plus dengan TOLOC lain ditolak', /TOLOC/.test(tutup(t5, 'Adjustment Plus', bukti('ADJUSTMENT', 'SKU5', '', 'A01.01.01', 1)).message));
check('Adjustment Plus yang benar diterima', tutup(t5, 'Adjustment Plus', bukti('ADJUSTMENT', 'SKU5', '', 'B02.05.01', 1)).queued === true);
worker();
const rw3 = rows('Riwayat'), r4 = rw3.find((r) => r[0] === t4.id), r5 = rw3.find((r) => r[0] === t5.id);
check('task tertutup: kategori, PIC, penutup, waktu', r4[18] === 'Selesai' && r4[17] === 'Kurang Picking' && r4[23] === ADMIN && r4[20] === '1004.inv' && isDate(r4[21]) && r5[17] === 'Adjustment Plus', [r4.slice(17, 25)]);
check('bukti tersimpan di Log_Bukti_Investigasi (QTY di kolom QTY)', rows('Log_Bukti_Investigasi').length === 2 && rows('Log_Bukti_Investigasi')[0][17] === 1 && rows('Log_Bukti_Investigasi')[0][7] === 'PICKING', rows('Log_Bukti_Investigasi')[0]);
check('status tidak bisa mundur (ditolak worker, tercatat)', (raw('updateTaskStatus', t4.id, '1004.inv', 'Sedang Dicari', 'mundur', '', '', [], t4.lokasi, t4.article, t4.selisih), worker(), rows('Riwayat').find((r) => r[0] === t4.id)[18] === 'Selesai'));

// ---------- 7. Penyelesaian Plus Minus ----------
const cand = api('getPlusMinusCandidates', '1004.inv');
check('kandidat pasangan: 1 minus, 1 plus (SKU10)', cand.minus.length === 1 && cand.plus.length === 1 && cand.minus[0].article === 'SKU10', cand);
const mT = cand.minus[0], pT = cand.plus[0];
check('bukti MOVE tanpa rute kedua lokasi ditolak', /FROMLOC/.test(api('closePlusMinusPair', mT.id, pT.id, '1004.inv', 'Salah taruh', ADMIN, bukti('MOVE', 'SKU10', 'A01.01.01', 'B02.05.01', 2)).message));
const pair = api('closePlusMinusPair', mT.id, pT.id, '1004.inv', 'Salah taruh', ADMIN, bukti('MOVE', 'SKU10', 'B02.04.02', 'B02.05.01', 2));
const rw4 = rows('Riwayat'), rm = rw4.find((r) => r[0] === mT.id), rp = rw4.find((r) => r[0] === pT.id);
check('pasangan tertutup bersama dan saling menunjuk', pair.success && rm[18] === 'Selesai' && rp[18] === 'Selesai' && rm[17] === 'Pindah Lokasi (Plus-Minus)' && rm[26] === pT.id && rp[26] === mT.id, [pair, rm.slice(17, 27)]);
check('bukti pasangan tersimpan untuk kedua task', rows('Log_Bukti_Investigasi').length === 4);
worker();
const pm2 = api('getPlusMinusSummary', ADMIN);
check('Summary Plus Minus kosong setelah semua task tertutup', pm2.plus.totalSku === 0 && pm2.minus.totalSku === 0, pm2);
check('tidak ada task aktif tersisa', api('getOpenTasks', ADMIN).length === 0);

// ---------- 8. Home, Dashboard, Analytics ----------
const today = ctx.Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd');
const home = api('getHomeSummary', ADMIN);
check('Home admin: 8 dihitung, 8 selesai, 0 belum validasi', home.totalCycleHariIni === 8 && home.selesaiHariIni === 8 && home.belumValidasi === 0, home);
const homeP = api('getHomeSummary', '1003.out');
check('Home petugas menghitung miliknya saja', homeP.totalCycleHariIni > 0 && homeP.totalCycleHariIni < 8 && homeP.selesaiHariIni === homeP.totalCycleHariIni, homeP);
check('Log_Akses satu baris per user per hari', rows('Log_Akses').length === 3 && rows('Log_Akses')[0][0] === today && rows('Log_Akses')[0][7] >= 2, rows('Log_Akses'));
const hb = api('getHomeBundle', ADMIN);
check('Home (satu panggilan) admin setelah semua selesai', hb.totalCycleHariIni === 8 && hb.selesaiHariIni === 8 && hb.outstanding === 0 && hb.openTasks === 0 && hb.pendingValidasi === 0 && hb.perPetugas.reduce((a, p) => a + p.dihitung, 0) === 8 && hb.perPetugas.every((p) => p.sisa === 0), hb);
const pr = api('getProductivity', today, ADMIN, 480);
check('Productivity: 8 item oleh 3 petugas, 5 validasi oleh 1 validator, sebaran per jam', pr.hariIni === true && pr.ringkasan.totalItem === 8 && pr.ringkasan.petugasAktif === 3 && pr.petugas.reduce((a, p) => a + p.total, 0) === 8 && pr.petugas.every((p) => p.mulai > 0 && p.akhir >= p.mulai && p.jamAktif >= 1 && p.perJam > 0 && p.role) && pr.validator.length === 1 && pr.validator[0].total === 5 && pr.ringkasan.totalValidasi === 5 && pr.perJam.reduce((a, j) => a + j.total, 0) === 8, pr);
check('Productivity tanggal lain kosong; petugas ditolak', api('getProductivity', '2020-01-01', ADMIN, 480).ringkasan.totalItem === 0 && /Akses ditolak/.test(raw('getProductivity', today, '1002.out', 480).error || ''));
const dash = api('getDashboardData', 'daily', today, 7, ADMIN);
check('Dashboard KPI: total 8, hit 4, discrepancy 4', dash.summary.kpi.total === 8 && dash.summary.kpi.hit === 4 && dash.summary.kpi.discrepancy === 4, dash.summary.kpi);
check('leaderboard petugas & validator terisi; tren 7 hari', dash.summary.leaderboard.length === 3 && dash.summary.leaderboardValidator.length === 1 && dash.trend.length === 7 && dash.trend[6].total === 8, [dash.summary.leaderboard, dash.summary.leaderboardValidator]);
const kat = Object.fromEntries(dash.errorAnalysis.map((e) => [e.alasan, e.jumlah]));
check('Analisis Kesalahan per kategori', kat['Kurang Picking'] === 1 && kat['Adjustment Plus'] === 1 && kat['Pindah Lokasi (Plus-Minus)'] === 2 && kat['Salah Hitung'] === 1, kat);
ctx.rebuildRingkasanHarian_(ADMIN); ctx.rebuildAntrianAktif_(ADMIN);
const dash2 = api('getDashboardData', 'daily', today, 7, ADMIN);
check('rebuild ringkasan menghasilkan angka yang sama', JSON.stringify(dash2.summary.kpi) === JSON.stringify(dash.summary.kpi) && JSON.stringify(dash2.errorAnalysis) === JSON.stringify(dash.errorAnalysis), [dash2.summary.kpi, dash2.errorAnalysis]);
check('bulanan juga terbaca', api('getDashboardData', 'monthly', today.slice(0, 7), 6, ADMIN).summary.kpi.total === 8);
check('detail item bermasalah, detail kategori, detail user', api('getProblemItemsDetail', 'daily', today, ADMIN).length === 4 && api('getErrorAnalysisDetail', 'Kurang Picking', 'daily', today, ADMIN).length === 1 && api('getUserDashboardDetail', '1003.out', 'daily', today, ADMIN).summary.total > 0);
const an = api('getAnalyticsRootCauseData', today, today, ADMIN);
const rc = Object.fromEntries(an.rootCause.map((r) => [r.label, r.jumlah]));
check('Analytics: akurasi 50%, akar masalah terpetakan', an.kpi.akurasi === 50 && rc['Salah Picking / Move'] === 1 && rc['Salah Putaway'] === 2 && rc['Master Data / Location'] === 1, [an.kpi, rc]);
check('Analytics detail per akar masalah', api('getAnalyticsRootCauseDetail', 'Salah Putaway', today, today, ADMIN).length === 2);
check('Riwayat Log task: 4 task (Salah Hitung otomatis disembunyikan)', api('getTaskLog', 200, ADMIN, today, today).length === 4);

// ---------- 9. Config ----------
check('Kelola User & Akses', api('getDaftarUserMaster', ADMIN).users.length === 5 && api('getDaftarAksesSetting', ADMIN).daftar.length === 1 && api('getDaftarUserMaster', '1002.out').success === false);
const la = api('getLevelAssignmentConfig', ADMIN);
check('Pembagian Tugas: mode lama, 3 grup alat', la.success && la.modeAssignment === 'legacy' && la.matrix.length === 3, la);
check('Facility & lokasi aktif terbaca', api('getDaftarFacility', ADMIN).facilities.length === 1 && api('getDaftarLokasiAktif', ADMIN, FAC).lokasi.length === 6, api('getDaftarLokasiAktif', ADMIN, FAC));
check('Setting Overflow bawaan mati', api('getSettingOverflow', ADMIN).overflowEnabled === false);
check('nonaktifkan user → tidak bisa login', (api('setStatusUserMaster', ADMIN, 4, '1002.out', 'Nonaktif'), api('getUserRole', '1002.out') === null), rows('Master_User'));
check('ubah peran langsung berlaku dan facility user tidak terhapus', (api('updateRoleUserMaster', ADMIN, 5, '1003.out', 'inbound'), api('getUserRole', '1003.out').role === 'inbound' && api('getUserRole', '1003.out').facilityId === FAC && rows('Master_User')[3][3] === FAC), rows('Master_User'));
api('setModeAssignment', ADMIN, 'matrix');
const up3 = api('importRawData', [['Picking', 'SKU20', 'Barang 20', 'A01.01.01', '', 1, 'wmsA'], ['Move', 'SKU21', 'Barang 21', 'X', 'A01.03.05', 1, 'wmsA']], [['A01.01.01', 'SKU20', 3], ['A01.03.05', 'SKU21', 2]], ['1001.storing', '1003.out'], ADMIN, 0, 0, 'job_3');
const dcM = rows('Data Count').filter((r) => r[10] === 'Pending');
check('Mode Matrix: pembagian mengikuti matriks peran x grup alat', up3.success && up3.modeAssignment === 'matrix' && dcM.length === 2 && dcM.find((r) => r[4] === 'SKU21')[9] === '1001.storing' && dcM.find((r) => r[4] === 'SKU20')[9] === '1003.out', [up3, dcM.map((r) => [r[4], r[9]])]);
check('item tanpa petugas yang berhak tidak dipaksakan & dicatat', (() => { const r = api('importRawData', [['Move', 'SKU22', 'Barang 22', 'X', 'A01.02.03', 1, 'wmsA']], [['A01.02.03', 'SKU22', 1]], ['1001.storing'], ADMIN, 0, 0, 'job_4'); return r.perluKonfirmasi === true && rows('Log_Unassigned').length === 1 && Array.isArray(r.peringatan) && r.peringatan.length === 1 && Array.isArray(r.catatan); })());
check('hasil unggah membawa peringatan & catatan sebagai daftar', Array.isArray(up.peringatan) && up.peringatan.length === 0 && Array.isArray(up.catatan) && Array.isArray(up2.catatan) && up2.ditugaskan === 0, [up.peringatan, up.catatan, up2.catatan]);
check('akses Config: langsung berlaku saat diberi dan dicabut (cache info user dibuang)', (() => {
  const sebelum = api('getUserRole', '1003.out').aksesSettingOverflow;
  const beri = api('tambahAksesSetting', ADMIN, '1003.out', 'Uji');
  const sesudah = api('getUserRole', '1003.out').aksesSettingOverflow;
  const daftar = api('getDaftarAksesSetting', ADMIN).daftar, baris = daftar.find((d) => d.nik === '1003.out');
  const cabut = api('hapusAksesSetting', ADMIN, baris.rowIndex, '1003.out');
  return sebelum === false && beri.success && sesudah === true && /^\d{4}-\d{2}-\d{2}/.test(baris.tanggal) && cabut.success && api('getUserRole', '1003.out').aksesSettingOverflow === false;
})());
const lp = api('tambahUserMaster', ADMIN, '1005.lp', 'lp'), mt = api('tambahUserMaster', ADMIN, '1006.mtc', 'maintenance');
check('peran LP & Maintenance: terdaftar, bisa login, boleh membuka Cycle', lp.success && mt.success && api('getUserRole', '1005.lp').role === 'lp' && Array.isArray(api('getMyPendingTasks', '1005.lp')) && api('getHomeBundle', '1006.mtc').role === 'maintenance' && /Akses ditolak/.test(raw('getPendingValidasi', '1005.lp').error || ''), [lp, mt]);

// ---------- 10. Perbaikan dari tinjauan ulang ----------
// lokasi aktif & daftar petugas di-cache 5 menit: perubahan harus langsung terlihat di formulir Upload
check('formulir Upload langsung mengikuti impor lokasi dan user baru (cache dibuang)', (() => {
  const a = api('getUploadFormData', ADMIN);
  const imp2 = api('importLokasiAktif', ADMIN, FAC, ['C09.001.1', 'C09.001.2'], false);
  const b = api('getUploadFormData', ADMIN);
  return imp2.success && b.lokasiAktif === a.lokasiAktif + 2 && b.users.some((u) => u.username === '1005.lp') && b.users.some((u) => u.username === '1006.mtc');
})(), api('getUploadFormData', ADMIN));
// jam hitung = jam petugas menyimpan di HP (hasil bisa baru terkirim lama kemudian)
check('jam hitung memakai jam dari HP bila masuk akal, jam server bila tidak', (() => {
  api('setModeAssignment', ADMIN, 'legacy');
  const r = api('importRawData', [['Picking', 'SKU30', 'B30', 'C09.001.1', '', 1, 'w'], ['Picking', 'SKU31', 'B31', 'C09.001.2', '', 1, 'w'], ['Picking', 'SKU32', 'B32', 'A01.01.01', '', 1, 'w']],
    [['C09.001.1', 'SKU30', 4], ['C09.001.2', 'SKU31', 4], ['A01.01.01', 'SKU32', 4]], ['1003.out'], ADMIN, 0, 0, 'job_5');
  const t = api('getMyPendingTasks', '1003.out'), no = (a) => t.find((x) => x.article === a).no;
  // "tadi": 3 jam lalu, tetapi tidak sebelum 00.01 WIB hari ini. Lewat tengah malam WIB, 3 jam lalu adalah
  // kemarin dan hitungan itu tidak masuk laporan hari ini, sehingga uji ini bergantung pada jam dijalankan.
  const awalHariWib = Math.floor((Date.now() + 7 * 3600000) / 86400000) * 86400000 - 7 * 3600000;
  const jamLalu = Math.max(awalHariWib + 60000, Date.now() - 3 * 3600000);
  api('submitCount', no('SKU30'), '1003.out', 4, jamLalu);                    // tadi: dipakai
  api('submitCount', no('SKU31'), '1003.out', 4, Date.now() + 86400000);      // besok: ditolak, pakai jam server
  api('submitCount', no('SKU32'), '1003.out', 4, Date.now() - 40 * 3600000);  // 40 jam lalu: ditolak
  worker();
  const w = (a) => rows('Riwayat').find((x) => x[3] === a)[15].getTime();
  return r.success && Math.abs(w('SKU30') - jamLalu) < 1000 && Math.abs(w('SKU31') - Date.now()) < 60000 && Math.abs(w('SKU32') - Date.now()) < 60000 && rows('Data Count').find((x) => x[4] === 'SKU30')[14].getTime() === w('SKU30');
})());
// Nomor tugas hanya unik di dalam satu facility. Hasil yang dibuat saat akun masih di facility lain
// (tertahan di HP, lalu akun dipindah) tidak boleh mendarat di tugas facility yang baru.
check('hasil hitung & validasi dari facility lain ditolak sebelum masuk antrean', (() => {
  api('importRawData', [['Picking', 'SKU50', 'B50', 'C09.001.1', '', 1, 'w'], ['Picking', 'SKU51', 'B51', 'C09.001.2', '', 1, 'w']],
    [['C09.001.1', 'SKU50', 4], ['C09.001.2', 'SKU51', 4]], ['1003.out'], ADMIN, 0, 0, 'job_7');
  const t = api('getMyPendingTasks', '1003.out'), no = (a) => t.find((x) => x.article === a).no;
  const antre = (n) => rows(n).length, c0 = antre('Queue_Submit_Cycle');
  const tolak = api('submitCount', no('SKU50'), '1003.out', 4, Date.now(), 'FAC-LAIN');
  const c1 = antre('Queue_Submit_Cycle');
  const terima = api('submitCount', no('SKU50'), '1003.out', 4, Date.now(), FAC);
  const lama = api('submitCount', no('SKU51'), '1003.out', 9, Date.now()); // aplikasi versi lama tidak mengirim facility: tetap diterima
  const c2 = antre('Queue_Submit_Cycle');
  worker();
  const v = api('getPendingValidasi', ADMIN).find((x) => x.article === 'SKU51');
  const v0 = antre('Queue_Validasi');
  const tolakV = api('submitValidasi', v.id, ADMIN, 4, Date.now(), 'FAC-LAIN');
  const v1 = antre('Queue_Validasi');
  const terimaV = api('submitValidasi', v.id, ADMIN, 4, Date.now(), FAC);
  worker();
  return tolak.success === false && tolak.kode === 'FACILITY_BERUBAH' && /facility/i.test(tolak.message) && c1 === c0 && terima.success && lama.success && c2 === c0 + 2 &&
    tolakV.success === false && tolakV.kode === 'FACILITY_BERUBAH' && v1 === v0 && terimaV.success;
})());
check('Productivity: jam berurutan sesuai kejadian dan mengikuti zona waktu HP', (() => {
  const p7 = api('getProductivity', today, ADMIN, 420), p8 = api('getProductivity', today, ADMIN, 480);
  const urut = p8.perJam.every((j, i) => i === 0 || j.jam === (p8.perJam[i - 1].jam + 1) % 24);
  // dua kelompok jam hanya ada bila hitungan "tadi" jatuh di jam yang berbeda dengan sekarang (tidak pada menit-menit pertama hari WIB)
  const awalHari = Math.floor((Date.now() + 7 * 3600000) / 86400000) * 86400000 - 7 * 3600000, tadi = Math.max(awalHari + 60000, Date.now() - 3 * 3600000);
  const minimal = Math.floor(tadi / 3600000) !== Math.floor(Date.now() / 3600000) ? 2 : 1;
  return p7.perJam.length === p8.perJam.length && p8.perJam.length >= minimal && urut && p8.perJam[0].jam === (p7.perJam[0].jam + 1) % 24 && p8.perJam.reduce((a, j) => a + j.total, 0) === p8.ringkasan.totalItem;
})(), api('getProductivity', today, ADMIN, 480).perJam);
// Developer: dikenali dari Master_User, dan hanya Developer yang boleh memberi/mencabut peran itu
check('peran Developer tidak bisa diberikan oleh admin biasa', (() => {
  const a = api('tambahUserMaster', ADMIN, '1007.dev', 'developer');
  const u = api('getDaftarUserMaster', ADMIN).users.find((x) => x.username === '1003.out');
  const b = api('updateRoleUserMaster', ADMIN, u.rowIndex, u.username, 'developer');
  return a.success === false && b.success === false && api('getUserRole', '1003.out').role !== 'developer';
})());
sheet('Master_User').appendRow(['1007.dev', 'developer', 'Aktif', '']);
ctx.clearMasterCache_();
check('Developer di Master_User dikenali: akses Config, pemilih facility, rincian Report', (() => {
  const d = api('getUserRole', '1007.dev');
  const pindah = api('setDeveloperActiveFacility', '1007.dev', FAC);
  const beri = api('tambahUserMaster', '1007.dev', '1008.dev', 'developer');
  return d.role === 'developer' && d.aksesSettingOverflow === true && api('getDaftarFacility', '1007.dev').success && pindah.success && api('getUserRole', '1007.dev').facilityId === FAC &&
    Array.isArray(api('getProblemItemsDetail', 'daily', today, '1007.dev')) && Array.isArray(api('getErrorAnalysisDetail', 'Salah Hitung', 'daily', today, '1007.dev')) && beri.success;
})(), raw('getProblemItemsDetail', 'daily', today, '1007.dev'));
check('akun sendiri tidak bisa dinonaktifkan; akun Developer hanya oleh Developer', (() => {
  const users = api('getDaftarUserMaster', ADMIN).users, saya = users.find((x) => x.username === ADMIN), dev = users.find((x) => x.username === '1007.dev');
  const a = api('setStatusUserMaster', ADMIN, saya.rowIndex, saya.username, 'Nonaktif');
  const b = api('setStatusUserMaster', ADMIN, dev.rowIndex, dev.username, 'Nonaktif');
  return a.success === false && b.success === false && api('getUserRole', ADMIN) !== null && api('getUserRole', '1007.dev') !== null;
})());
// akun tanpa facility aktif tidak boleh melihat data spreadsheet pusat
// Akun dipindah ke facility lain: login berikutnya harus langsung melihat facility baru (bukan cache 5 menit),
// dan hasil yang dibuat di facility lama ditolak.
check('pindah facility langsung terlihat saat login dan hasil dari facility lama ditolak', (() => {
  const f2 = api('tambahFacility', '1007.dev', 'RDC Uji Dua', 'RDC-UJI2', 'Cycle Count RDC Uji Dua', ['Z01.001.1'], 'job_f2');
  const fac2 = (api('getDaftarFacility', '1007.dev').facilities || []).find((f) => f.kode === 'RDC-UJI2');
  if (!f2.success || !fac2) return false;
  const sebelum = api('getUserRole', '1003.out'); // mengisi cache peran
  const pindah = api('assignUserKeFacility', '1007.dev', '1003.out', fac2.id);
  const sesudah = api('getUserRole', '1003.out');
  const tolak = api('submitCount', 1, '1003.out', 1, Date.now(), FAC);
  const kembali = api('assignUserKeFacility', '1007.dev', '1003.out', FAC);
  return pindah.success && kembali.success && sebelum.facilityId === FAC && sesudah.facilityId === fac2.id && tolak.success === false && /facility/i.test(tolak.message) && api('getUserRole', '1003.out').facilityId === FAC;
})(), api('getUserRole', '1003.out'));
check('pesan tambah user menyebut nama facility, bukan kode internal', (() => {
  const r = api('tambahUserMaster', ADMIN, '1010.out', 'outbound');
  return r.success && /RDC Tallo Makassar/.test(r.message) && !/Auto-assign|FAC[-_]/i.test(r.message);
})(), 'pesan tidak sesuai');
sheet('Master_User').appendRow(['1009.tanpa', 'admin', 'Aktif', '']);
ctx.clearMasterCache_();
check('akun tanpa facility: Home, Upload, Productivity kosong dan ditandai', (() => {
  const h = api('getHomeBundle', '1009.tanpa'), u = api('getUploadFormData', '1009.tanpa'), p = api('getProductivity', today, '1009.tanpa', 420);
  return h.tanpaFacility === true && h.outstanding === 0 && h.totalCycleHariIni === 0 && h.perPetugas.length === 0 && u.tanpaFacility === true && u.users.length === 0 && p.tanpaFacility === true && p.petugas.length === 0;
})(), api('getHomeBundle', '1009.tanpa'));
check('Mode Matrix: petugas yang perannya belum dicentang dicatat', (() => {
  api('setModeAssignment', ADMIN, 'matrix');
  const m = api('getLevelAssignmentConfig', ADMIN).matrix.map((g) => ({ grupAlat: g.grupAlat, roles: g.roles.filter((r) => r !== 'lp'), mode: g.mode }));
  api('saveLevelAssignmentMatrix', ADMIN, m, 2);
  const r = api('importRawData', [['Picking', 'SKU40', 'B40', 'A01.01.01', '', 1, 'w']], [['A01.01.01', 'SKU40', 2]], ['1003.out', '1005.lp'], ADMIN, 0, 0, 'job_6');
  return r.success && r.catatan.some((c) => /1005\.lp/.test(c) && /belum dicentang/.test(c));
})());
const hbAkhir = api('getHomeBundle', ADMIN);
check('Home: tanggal kerja server dan item tanpa petugas dipisah dari daftar orang', hbAkhir.hariIni === today && hbAkhir.belumDitugaskan === 0 && hbAkhir.perPetugas.every((p) => p.nama && p.nama.charAt(0) !== '('), hbAkhir.perPetugas);
check('doGet menampilkan halaman server aktif', /server aktif/.test(ctx.doGet({}).getContent()));
check('router lama (tanpa args) tetap jalan', JSON.parse(ctx.doPost({ postData: { contents: JSON.stringify({ action: 'getAppVersion' }) } }).getContent()).success === true);
const errLog = rows('Log_Sistem').filter((r) => /ERROR/.test(r[2]));
check('tidak ada ERROR worker di Log_Sistem', errLog.length === 0, errLog);

console.log(failed ? failed + ' pemeriksaan gagal' : 'Semua uji backend lolos');
process.exit(failed ? 1 : 0);
