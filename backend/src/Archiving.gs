/**
 * Archiving.gs
 * Fase 22 - performa: pindahkan baris Riwayat yang sudah Selesai & tua ke Riwayat_Archive,
 * supaya sheet Riwayat utama (yang di-scan penuh oleh Dashboard, Task Investigasi, Home
 * widget Plus/Minus, Validasi) tetap ramping seiring waktu.
 *
 * v8.26.0 (Facility Management - Strict Isolation):
 * SELURUH fungsi di file ini SEKARANG menerima parameter opsional `username`.
 * Kalau disediakan, arsip dibaca/ditulis ke spreadsheet facility user tersebut.
 * Fungsi `getRiwayatArchiveSheet_()` TIDAK LAGI didefinisikan di sini -- memakai versi
 * facility-aware dari SheetHelpers.gs yang menerima parameter username.
 *
 * DESAIN PENTING:
 * 1) Kriteria archivable: Status_Task === 'Selesai' DAN Tanggal (tanggal cycle count,
 *    bukan Waktu_Selesai_Task -- field itu cuma keisi untuk task yang lewat alur investigasi
 *    manual, task HIT otomatis tidak pernah mengisinya) sudah lebih tua dari
 *    ARCHIVE_AGE_MONTHS. Task yang masih aktif (Open/Sedang Dicari/Menunggu Konfirmasi)
 *    TIDAK PERNAH diarsipkan meskipun tanggalnya sudah lama -- itu justru task yang paling
 *    perlu kelihatan/butuh tindak lanjut.
 * 2) Ditulis dalam 2 batch operation (bukan appendRow/deleteRows satu-satu per baris) --
 *    aman dari limit eksekusi 6 menit meskipun datanya sudah puluhan ribu baris.
 * 3) Dibungkus LockService.getScriptLock() yang SAMA dengan submitCount/submitValidasi/
 *    updateTaskStatus/closePlusMinusPair/importRawData -- supaya tidak ada race condition
 *    (baris baru ke-overwrite/ke-skip) kalau ada user submit transaksi persis pas archiving
 *    jalan.
 * 4) Cache findRiwayatRow_ (lihat SheetHelpers.gs) SUDAH self-healing by design: dia selalu
 *    verifikasi ID di baris hasil cache sebelum dipakai, kalau tidak cocok (nomor baris
 *    bergeser akibat archive) otomatis fallback ke scan ulang & re-cache -- jadi TIDAK perlu
 *    flush manual setelah archive. ID yang sudah diarsipkan otomatis dapat -1 (tidak
 *    ketemu) kalau dicari lagi, yang memang seharusnya begitu karena task Selesai tidak
 *    pernah disentuh ulang oleh alur aplikasi manapun.
 */
// ---------- Sheet accessor ----------
//
// PENTING v8.26.0: getRiwayatArchiveSheet_() TIDAK didefinisikan di sini lagi.
// Memakai versi facility-aware dari SheetHelpers.gs: getRiwayatArchiveSheet_(username)
//
function hitungCutoffArchive_() {
  const now = new Date();
  return new Date(now.getFullYear(), now.getMonth() - ARCHIVE_AGE_MONTHS, now.getDate());
}

/**
 * dryRun=true: cuma hitung & laporkan, TIDAK menulis apapun (aman dipanggil kapan saja).
 * dryRun=false: benar-benar memindahkan baris ke Riwayat_Archive & merapikan Riwayat.
 *
 * v8.26.0: Menerima parameter opsional `username` untuk archive per-facility.
 */
function archiveOldRiwayat_(dryRun, username) {
  // v8.29.0: lock PER-FACILITY (bukan global) -- HARUS pakai key facility yang SAMA dengan
  // submitCount/submitValidasi/dst (lihat acquireFacilityLock_ di FacilityManagement.gs)
  // supaya archive tetap mutually-exclusive terhadap submit di facility yang sama (mencegah
  // race clearContent/rewrite Riwayat pas ada user submit persis bersamaan), tapi archive
  // facility lain tidak ikut ke-block.
  const facInfoForLock_ = getUserFacility(username);
  const facLock = acquireFacilityLock_(facInfoForLock_ ? facInfoForLock_.id : null, 20000);
  if (!facLock) throw new Error('Sistem sedang sibuk, coba lagi nanti.');
  try {
    // v8.26.0: Baca Riwayat dari facility user
    const sheet = getRiwayatSheet_(username);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, archived: 0, sisa: 0, dryRun: !!dryRun };
    const data = sheet.getRange(2, 1, lastRow - 1, RIWAYAT_HEADERS.length).getValues();
    const cutoff = hitungCutoffArchive_();
    const keep = [];
    const toArchive = [];
    data.forEach(function (row) {
      const statusTask = row[18]; // Status_Task
      // pakai rowDateTag_() (berbasis Waktu_Cycle, fallback ke Tanggal) -- SAMA persis dengan
      // konvensi yang dipakai semua fungsi Dashboard (fillTrendFromData_, computeSummaryFromData_,
      // dst) supaya kriteria "usia" archive konsisten dengan cara Dashboard mengelompokkan tanggal.
      const tglStr = rowDateTag_(row);
      const tgl = tglStr ? new Date(String(tglStr) + 'T00:00:00') : null;
      const eligible = statusTask === 'Selesai' && tgl && !isNaN(tgl.getTime()) && tgl < cutoff;
      if (eligible) toArchive.push(row); else keep.push(row);
    });
    if (toArchive.length === 0) {
      return { success: true, archived: 0, sisa: keep.length, dryRun: !!dryRun };
    }
    if (dryRun) {
      return { success: true, archived: toArchive.length, sisa: keep.length, dryRun: true };
    }
    // 1) Tulis ke Riwayat_Archive milik facility user -- 1x batch write
    const archiveSheet = getRiwayatArchiveSheet_(username);
    const archiveLastRow = archiveSheet.getLastRow();
    archiveSheet.getRange(archiveLastRow + 1, 1, toArchive.length, RIWAYAT_HEADERS.length).setValues(toArchive);
    // 2) Rapikan Riwayat -- hapus SEMUA baris data lalu tulis ulang yang "keep" saja
    //    (1x clear + 1x write, jauh lebih cepat & aman dari limit eksekusi dibanding
    //    deleteRows() satu-satu untuk ribuan baris)
    sheet.getRange(2, 1, lastRow - 1, RIWAYAT_HEADERS.length).clearContent();
    if (keep.length > 0) {
      sheet.getRange(2, 1, keep.length, RIWAYAT_HEADERS.length).setValues(keep);
    }
    return { success: true, archived: toArchive.length, sisa: keep.length, dryRun: false };
  } finally {
    facLock.release();
  }
}

/**
 * Menu "Cycle Count > Cek & Jalankan Archive Riwayat Lama" -- cek dulu (dry run), tampilkan
 * ringkasannya, baru minta konfirmasi sebelum benar-benar memindahkan data. Aman dijalankan
 * berkali-kali (baris yang sudah diarsipkan otomatis tidak akan ke-proses ulang, karena
 * sudah tidak ada lagi di sheet Riwayat begitu dipindah).
 */
function cekArchiveRiwayatLama_() {
  const ui = SpreadsheetApp.getUi();
  // v8.26.0: Menu dipanggil dari spreadsheet aktif, pakai null/fallback.
  const cek = archiveOldRiwayat_(true, null);
  if (!cek.archived) {
    ui.alert('Tidak ada baris Riwayat yang perlu diarsipkan saat ini (kriteria: Status_Task=Selesai & usia > ' + ARCHIVE_AGE_MONTHS + ' bulan).');
    return;
  }
  const pesan = 'Ditemukan ' + cek.archived + ' baris Riwayat (Status_Task=Selesai, usia > ' + ARCHIVE_AGE_MONTHS + ' bulan) siap diarsipkan.\n\n' +
    'Sisa yang TETAP di Riwayat (masih aktif atau masih baru): ' + cek.sisa + ' baris.\n\n' +
    'Baris yang diarsipkan akan dipindah ke sheet "' + RIWAYAT_ARCHIVE_SHEET_NAME + '" -- data TIDAK dihapus, hanya dipindah, tetap bisa dibuka manual kapan saja.\n\n' +
    'Lanjutkan?';
  const resp = ui.alert(pesan, ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;
  // v8.26.0: Menu dipanggil dari spreadsheet aktif, pakai null/fallback.
  const hasil = archiveOldRiwayat_(false, null);
  ui.alert('Selesai. ' + hasil.archived + ' baris dipindah ke "' + RIWAYAT_ARCHIVE_SHEET_NAME + '". Sisa ' + hasil.sisa + ' baris aktif di Riwayat.');
}
