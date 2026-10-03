// Data contoh untuk mode demo aplikasi dan uji tampilan. Semua nama, NIK, SKU, dan angka di sini fiktif.
// Hari ini diisi lewat fungsi backend yang asli (unggah → hitung → validasi → verifikasi), jadi alurnya
// sama persis dengan pemakaian nyata; enam hari sebelumnya ditulis langsung ke sheet Riwayat supaya
// Dashboard, Productivity, dan Analytics punya riwayat.
//
// env: { api(nama, ...args), fn(nama, ...args), sheet(nama) }
//   api  → memanggil backend lewat doPost (format yang sama dengan aplikasi)
//   fn   → memanggil fungsi backend langsung (setupAwal, workerSemua, rebuild…)
//   sheet→ sheet tiruan di spreadsheet aktif (untuk riwayat & penyesuaian jam)

export const DEMO = {
  facility: 'RDC Tallo Makassar', kode: 'RDC-MKS', admin: '100101.RINA',
  users: [
    ['100102.BAYU', 'inventory'], ['100103.SARI', 'inventory'],
    ['100201.ANDI', 'outbound'], ['100202.DEWI', 'outbound'], ['100203.FAJAR', 'outbound'], ['100210.NANDA', 'outbound'],
    ['100204.HASAN', 'inbound'], ['100205.INTAN', 'inbound'], ['100211.OKTA', 'inbound'],
    ['100206.JOKO', 'storing'], ['100207.KIKI', 'storing'],
    ['100208.LUKMAN', 'lp'], ['100209.MAYA', 'maintenance']
  ]
};

export function seedDemo(env, opts) {
  const o = Object.assign({ seed: 20261002, items: 900, historyDays: 6, countShare: 0.93, now: new Date() }, opts || {});
  const { api, fn, sheet } = env;
  const ADMIN = DEMO.admin;

  // ---------- acak yang bisa diulang ----------
  let st = o.seed >>> 0;
  const rnd = () => { st |= 0; st = (st + 0x6D2B79F5) | 0; let t = Math.imul(st ^ (st >>> 15), 1 | st); t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t; return ((t ^ (t >>> 14)) >>> 0) / 4294967296; };
  const int = (a, b) => a + Math.floor(rnd() * (b - a + 1));
  const pick = (arr) => arr[Math.floor(rnd() * arr.length)];
  const uid = () => 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c) => { const r = Math.floor(rnd() * 16); return (c === 'x' ? r : (r & 3) | 8).toString(16); });
  const pad = (n, w) => String(n).padStart(w, '0');

  // ---------- tanggal (backend memakai tanggal Asia/Jakarta) ----------
  const JKT = 7 * 3600000;
  const nowMs = o.now.getTime();
  const dayStart = Math.floor((nowMs + JKT) / 86400000) * 86400000 - JKT; // 00:00 WIB hari ini
  const tag = (ms) => new Date(ms + JKT).toISOString().slice(0, 10);

  // ---------- 1. setup, user, lokasi ----------
  fn('setupAwal');
  const FAC = api('getUserRole', ADMIN).facilityId;
  DEMO.users.forEach(([u, r]) => { const res = api('tambahUserMaster', ADMIN, u, r); if (!res.success) throw new Error('seed user ' + u + ': ' + res.message); });
  const lokasi = [];
  ['A', 'B', 'C'].forEach((g) => { for (let l = 1; l <= 4; l++) for (let s = 1; s <= 24; s++) for (let lv = 1; lv <= 6; lv++) lokasi.push(g + pad(l, 2) + '.' + pad(s, 3) + '.' + lv); });
  api('importLokasiAktif', ADMIN, FAC, lokasi, true);

  // ---------- katalog barang ----------
  const NAMA = ['ARLO', 'NORA', 'KIRA', 'LUMA', 'VIGO', 'TARA', 'ELIO', 'MILA', 'ZENO', 'SENA', 'RUBI', 'DARA'];
  const JENIS = ['2 DOORS WARDROBE', '3 DOORS WARDROBE', 'COFFEE TABLE', 'TV CABINET', 'OFFICE CHAIR', 'DINING CHAIR', 'BEDSIDE TABLE', 'BOOKSHELF', 'SHOE RACK', 'SOFA 2 SEATER', 'QUEEN BED FRAME', 'STUDY DESK'];
  const WARNA = ['WHITE', 'BEIGE', 'OAK', 'WALNUT', 'GREY', 'BLACK'];
  const katalog = [];
  for (let i = 0; i < 140; i++) {
    const paket = rnd() < 0.3 ? ' (P' + int(1, 2) + '/2)' : '';
    katalog.push({ article: String(70100000 + i * 137 + int(0, 90)), description: pick(NAMA) + ' ' + pick(JENIS) + ' ' + pick(WARNA) + paket });
  }
  const WMS = ['wms.ahmad', 'wms.bella', 'wms.candra', 'wms.dian', 'wms.eko'];
  const petugas = DEMO.users.filter(([, r]) => r !== 'inventory').map(([u]) => u);
  const validator = DEMO.users.filter(([, r]) => r === 'inventory').map(([u]) => u);
  const KATEGORI = ['Lebih Picking', 'Kurang Picking', 'Lebih Move', 'Kurang Move', 'Barang ketemu di lokasi lain', 'Adjustment Plus', 'Adjustment Minus'];

  // ---------- 2. riwayat hari-hari sebelumnya (langsung ke sheet Riwayat) ----------
  const rw = sheet('Riwayat');
  const hist = [];
  for (let d = o.historyDays; d >= 1; d--) {
    const base = dayStart - d * 86400000;
    const n = int(780, 1020), akurasi = 0.93 + rnd() * 0.05;
    for (let i = 0; i < n; i++) {
      const b = pick(katalog), lok = pick(lokasi), sys = int(1, 36), who = pick(petugas);
      const wCycle = new Date(base + 6.5 * 3600000 + rnd() * 8.5 * 3600000);
      const addWho = (rnd() < 0.5 ? 'MOVE:' : 'PICKING:') + pick(WMS);
      const row = [uid(), tag(base), lok, b.article, b.description, int(1, 6), sys, who, sys, 0, 'HIT', 'Tidak Perlu', '', '', 'HIT', wCycle, '', '', 'Selesai', '', '', '', '', '', wCycle, addWho, ''];
      if (rnd() > akurasi) {
        const sel = pick([-3, -2, -1, -1, 1, 1, 2]), val = pick(validator), wVal = new Date(wCycle.getTime() + int(20, 120) * 60000);
        row[8] = Math.max(0, sys + sel); row[9] = row[8] - sys; row[10] = 'DISCREPANCY'; row[11] = 'Selesai'; row[12] = val; row[16] = wVal; row[22] = val;
        if (rnd() < 0.38) { // validasi ulang cocok: petugas yang salah hitung
          row[13] = sys; row[14] = 'HIT'; row[17] = 'Salah Hitung'; row[18] = 'Selesai'; row[20] = val; row[21] = wVal; row[24] = wVal;
          row[19] = 'Otomatis: hasil validasi ulang sesuai stok system, tidak ada discrepancy nyata (kesalahan hitung awal petugas).';
        } else {
          row[13] = row[8]; row[14] = 'DISCREPANCY'; row[23] = pick(validator.concat([ADMIN]));
          const masihBuka = rnd() < (d <= 2 ? 0.45 : 0.08);
          if (masihBuka) { row[18] = rnd() < 0.5 ? 'Open' : 'Sedang Dicari'; row[24] = wVal; }
          else {
            const wSelesai = new Date(wVal.getTime() + int(1, 20) * 3600000);
            row[17] = row[9] > 0 ? pick(['Lebih Picking', 'Lebih Move', 'Adjustment Plus', 'Barang ketemu di lokasi lain']) : pick(['Kurang Picking', 'Kurang Move', 'Adjustment Minus', 'Barang ketemu di lokasi lain']);
            row[18] = 'Selesai'; row[20] = pick(validator); row[21] = wSelesai; row[24] = wSelesai;
            row[19] = '[' + pad(wSelesai.getDate(), 2) + '/' + pad(wSelesai.getMonth() + 1, 2) + ' - Selesai] Sudah dicek ke lokasi dan dicocokkan dengan transaksi WMS.';
          }
        }
      }
      hist.push(row);
    }
  }
  if (hist.length) rw.getRange(rw.getLastRow() + 1, 1, hist.length, 27).setValues(hist);

  // ---------- 3. hari ini: unggah & bagi tugas ----------
  const trans = [], stock = [], pakai = {};
  const tambah = (tipe, b, lok, qty) => {
    trans.push(tipe === 'Move' ? ['Move', b.article, b.description, 'STAGING', lok, qty, pick(WMS)] : ['Picking', b.article, b.description, lok, '', qty, pick(WMS)]);
    const k = lok + '||' + b.article;
    if (!pakai[k]) { pakai[k] = int(2, 40); stock.push([lok, b.article, pakai[k]]); }
  };
  for (let i = 0; i < o.items; i++) tambah(rnd() < 0.55 ? 'Picking' : 'Move', pick(katalog), pick(lokasi), int(1, 6));
  // satu SKU salah taruh: kurang 2 di satu lokasi, lebih 2 di lokasi lain (kandidat pasangan plus-minus)
  const pasangan = katalog[7], lokMinus = 'B02.011.2', lokPlus = 'B03.016.2';
  tambah('Picking', pasangan, lokMinus, 2); tambah('Move', pasangan, lokPlus, 2);
  trans.push(['Adjustment', katalog[3].article, katalog[3].description, 'A01.001.1', '', 1, 'wms.eko']);     // tipe lain: dilewati
  trans.push(['Picking', katalog[4].article, katalog[4].description, 'Z99.001.1', '', 1, 'wms.eko']);        // lokasi tidak aktif
  const up = api('importRawData', trans, stock, petugas, ADMIN, 2, 3, 'seed');
  if (!up.success) throw new Error('seed unggah: ' + up.message);

  // ---------- 4. hitung ----------
  const hasilHitung = {};
  petugas.forEach((u, idx) => {
    const tugas = api('getMyPendingTasks', u);
    const porsi = idx === 0 ? 0.55 : idx === 1 ? 0.62 : (idx === 5 ? 1 : Math.min(1, o.countShare + (rnd() - 0.5) * 0.12)); // dua orang masih punya banyak tugas, satu orang tuntas
    const n = Math.min(tugas.length, Math.round(tugas.length * porsi));
    for (let i = 0; i < tugas.length; i++) {
      const t = tugas[i];
      if (i >= n && t.article !== pasangan.article) continue; // SKU pasangan selalu dihitung supaya contohnya ada
      let q = t.qtySystem;
      if (t.article === pasangan.article && t.lokasi === lokMinus) q = t.qtySystem - 2;
      else if (t.article === pasangan.article && t.lokasi === lokPlus) q = t.qtySystem + 2;
      else if (rnd() < 0.085) q = Math.max(0, t.qtySystem + pick([-3, -2, -1, -1, 1, 1, 2]));
      hasilHitung[t.lokasi + '||' + t.article] = q;
      api('submitCount', t.no, u, q);
    }
  });
  fn('workerSemua');

  // jam hitung disebar sepanjang hari (worker mencatat "sekarang" untuk semuanya)
  const awal = Math.max(dayStart + 60000, nowMs - 7.5 * 3600000), rentang = Math.max(60000, nowMs - 120000 - awal);
  const today = tag(nowMs), urut = {}, jamHitung = {};
  const dc = sheet('Data Count');
  const dcRows = dc.getLastRow() >= 2 ? dc.getRange(2, 1, dc.getLastRow() - 1, 16).getValues() : [];
  const total = {};
  dcRows.forEach((r) => { if (r[10] === 'Selesai') total[r[9]] = (total[r[9]] || 0) + 1; });
  dcRows.forEach((r, i) => {
    if (r[10] !== 'Selesai') return;
    const u = r[9]; urut[u] = (urut[u] || 0) + 1;
    const t = new Date(awal + rentang * Math.min(1, (urut[u] - 1 + rnd() * 0.8) / Math.max(1, total[u])));
    jamHitung[r[2] + '||' + r[4]] = t;
    dc.getRange(i + 2, 15).setValue(t);
  });
  const rwRows = () => rw.getRange(2, 1, rw.getLastRow() - 1, 27).getValues();
  rwRows().forEach((r, i) => {
    const t = jamHitung[r[2] + '||' + r[3]];
    if (r[1] !== today || !t) return;
    rw.getRange(i + 2, 16).setValue(t); rw.getRange(i + 2, 25).setValue(t);
  });

  // ---------- 5. validasi ulang (sebagian masih menunggu) ----------
  validator.forEach((v) => {
    const antre = api('getPendingValidasi', v);
    antre.forEach((it, i) => {
      const k = it.lokasi + '||' + it.article, punyaPasangan = it.article === pasangan.article;
      if (!punyaPasangan && i % 3 === 2) return; // dibiarkan menunggu
      const q = punyaPasangan || rnd() < 0.62 ? hasilHitung[k] : it.qtySystem;
      api('submitValidasi', it.id, v, q);
    });
  });
  fn('workerSemua');

  // ---------- 6. verifikasi: satu task sedang dicari, satu ditutup ----------
  const tasks = api('getOpenTasks', validator[0]).filter((t) => t.tanggal === today && t.article !== pasangan.article);
  if (tasks[0]) api('updateTaskStatus', tasks[0].id, validator[0], 'Sedang Dicari', 'Sudah cek lokasi kiri dan kanan, belum ketemu. Lanjut cek area staging.', '', '', [], tasks[0].lokasi, tasks[0].article, tasks[0].selisih);
  if (tasks[1]) api('updateTaskStatus', tasks[1].id, validator[0], 'Selesai', 'Barang sudah di picking, discrepancy terkonfirmasi.', 'Barang Sudah di Picking', ADMIN, [], tasks[1].lokasi, tasks[1].article, tasks[1].selisih);
  fn('workerSemua');

  // ---------- 7. ringkasan dihitung ulang dari Riwayat (mencakup riwayat yang ditulis langsung) ----------
  fn('rebuildRingkasanHarian_', ADMIN);
  fn('rebuildAntrianAktif_', ADMIN);

  return { admin: ADMIN, inventory: validator[0], petugas: petugas[1], tertinggal: petugas[0], tuntas: petugas[5], facilityId: FAC, today, pasangan: { article: pasangan.article, minus: lokMinus, plus: lokPlus }, ditugaskan: up.ditugaskan };
}
