/**
 * PerfTest.gs
 * Tools pengujian performa untuk mensimulasikan beban tinggi submit count.
 * HANYA untuk developer -- tidak perlu dideploy ke production.
 *
 * Cara pakai:
 * 1. Di Apps Script editor, pilih fungsi testKecepatan50User, klik Run.
 * 2. Lihat hasil di Execution Log (View > Logs).
 * 3. Untuk uji 50 user concurrent: jalankan testKecepatan50User di beberapa tab
 *    browser secara BERSAMAAN -- setiap tab = 1 eksekusi GAS terpisah, TIDAK perlu akun
 *    Google/user login yang berbeda-beda (USERNAME_TEST cuma string business-layer buat
 *    lookup facility, tidak terikat sesi login browser).
 *
 * v8.29.3: OFFSET_SLOT ditambahkan supaya tiap tab bisa digilir ambil IRISAN item Pending
 * yang BERBEDA (bukan rebutan item yang sama). Kalau semua tab pakai OFFSET_SLOT yang sama
 * (default 0), mereka akan mengincar 10 item PERTAMA yang identik -- begitu ada klaim
 * per-item (lihat CycleCount.gs v8.29.3), sebagian besar akan gagal dengan "Item ini sedang
 * diproses / sudah diproses sebelumnya". Itu BUKAN bukti lock lambat, itu proteksi race
 * condition bekerja sesuai desain -- tapi hasilnya jadi tidak representatif (aslinya 30-50
 * user beda orang mengerjakan barang masing-masing yang TIDAK saling tabrok). Ini
 * OFFSET_SLOT beda-beda tiap tab (0, 10, 20, 30, ...) supaya tiap tab dapat 10 item unik,
 * baru hasilnya murni mengukur throughput lock/slot di beban paralel yang sebenarnya.
 */

/**
 * Mensimulasikan N submit berurutan. Mengukur overhead per-submit.
 */
function testKecepatan50User() {
  const USERNAME_TEST = SETUP_ADMIN_NIK_; // NIK yang aktif & punya facility
  const JUMLAH_ITERASI = 10;
  const OFFSET_SLOT = 0; // Ganti per tab: 0, 10, 20, 30, ... supaya tiap tab pegang item Pending yang BEDA

  Logger.log('=== PERF TEST: ' + JUMLAH_ITERASI + ' sequential submit (offset ' + OFFSET_SLOT + ') ===');

  const facInfo = getUserFacility(USERNAME_TEST);
  if (!facInfo) {
    Logger.log('ERROR: User tidak punya facility. Ganti USERNAME_TEST.');
    return;
  }
  Logger.log('Facility: ' + facInfo.nama + ' (' + facInfo.id + ')');

  const sheet = getSheet_(USERNAME_TEST);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('Tidak ada data.'); return; }

  const allRows = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
  const pendingNos = allRows
    .filter(function (r) {
      return String(r[9] || '').trim().toLowerCase() === USERNAME_TEST.toLowerCase()
        && r[10] === 'Pending';
    })
    .slice(OFFSET_SLOT, OFFSET_SLOT + JUMLAH_ITERASI)
    .map(function (r) { return r[0]; });

  if (!pendingNos.length) {
    Logger.log('Tidak ada item Pending untuk user ini di offset ' + OFFSET_SLOT + '. Upload data lebih banyak, atau kecilkan OFFSET_SLOT.');
    return;
  }

  const waktuMulai = Date.now();
  let berhasil = 0, gagal = 0;
  const detailMs = [];

  pendingNos.forEach(function (no, idx) {
    const t0 = Date.now();
    try {
      const res = submitCount(no, USERNAME_TEST, Math.floor(Math.random() * 100));
      const ms = Date.now() - t0;
      detailMs.push(ms);
      if (res && res.success) { berhasil++; Logger.log('[' + (idx+1) + '] No ' + no + ' : OK (' + ms + ' ms)'); }
      else { gagal++; Logger.log('[' + (idx+1) + '] No ' + no + ' : FAIL ' + ms + ' ms - ' + (res ? res.message : '')); }
    } catch (e) {
      gagal++;
      detailMs.push(Date.now() - t0);
      Logger.log('[' + (idx+1) + '] No ' + no + ' : ERROR - ' + e.message);
    }
  });

  const totalMs = Date.now() - waktuMulai;
  const rataMs = detailMs.length ? Math.round(detailMs.reduce(function(a,b){ return a+b; },0)/detailMs.length) : 0;
  Logger.log('');
  Logger.log('=== HASIL ===');
  Logger.log('Berhasil: ' + berhasil + ' / ' + pendingNos.length + ' | Gagal: ' + gagal);
  Logger.log('Total: ' + totalMs + ' ms | Rata-rata: ' + rataMs + ' ms/submit');
  Logger.log('Min: ' + (detailMs.length ? Math.min.apply(null, detailMs) : 0) + ' ms | Max: ' + (detailMs.length ? Math.max.apply(null, detailMs) : 0) + ' ms');
  Logger.log('Throughput: ~' + (berhasil / (totalMs/1000)).toFixed(2) + ' submit/detik (sequential)');
  Logger.log('Estimasi 50 user paralel (10-slot lock): ~' + Math.ceil(50/10 * rataMs / 1000) + ' detik');
}

/**
 * Ukur overhead lock saja tanpa submit nyata.
 */
function testLockPerformance() {
  const FACILITY_ID = null; // null = 'default', atau isi ID facility
  const JUMLAH = 20;
  Logger.log('=== LOCK TEST: ' + JUMLAH + ' iterasi, sleep 200ms/iterasi ===');
  const t0 = Date.now();
  let ok = 0;
  for (let i = 0; i < JUMLAH; i++) {
    const lt = Date.now();
    const lk = acquireFacilityLock_(FACILITY_ID, 5000);
    if (lk) {
      ok++;
      Utilities.sleep(200);
      lk.release();
      Logger.log('#' + (i+1) + ' OK (' + (Date.now()-lt) + ' ms)');
    } else {
      Logger.log('#' + (i+1) + ' TIMEOUT (' + (Date.now()-lt) + ' ms)');
    }
  }
  Logger.log('Total: ' + (Date.now()-t0) + ' ms | OK: ' + ok + ' / ' + JUMLAH);
}
