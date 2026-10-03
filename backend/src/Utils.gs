/**
 * Utils.gs
 * Fungsi utilitas murni (parsing lokasi, encode/decode AddWho, tanggal) --
 * tidak akses Sheet/Cache
 * (Hasil pemecahan dari Code.gs -- semua file .gs berbagi global scope yang
 * sama di Apps Script,
 * jadi urutan file tidak masalah, fungsi di file lain tetap bisa saling
 * panggil.)
 */

// --------- AddWho_Transaksi: histori user WMS dipecah per jenis transaksi ---------
// Format string yang disimpan di kolom AddWho_Transaksi (Data Count kol.16 /
// Riwayat kol.26):
// "MOVE:nama1,nama2;PICKING:nama3"
// LAINNYA dipakai untuk menampung data lama (sebelum breakdown per jenis
// ada, formatnya cuma
// flat "nama1,nama2") supaya saat digabung (merge upload) datanya tidak
// hilang/ketimpa.
function encodeAddWhoByType_(byType) {
  const order = ['move', 'picking', 'lainnya'];
  const parts = [];
  order.forEach(function (t) {
    const names = Object.keys((byType && byType[t]) || {}).filter(function (n) { return n; }).sort();
    if (names.length) parts.push(t.toUpperCase() + ':' + names.join(','));
  });
  return parts.join(';');
}

function decodeAddWhoByType_(str) {
  const s = String(str || '').trim();
  const result = { move: {}, picking: {}, lainnya: {} };
  if (!s) return result;
  if (/(MOVE|PICKING|LAINNYA):/i.test(s)) {
    s.split(';').forEach(function (part) {
      const idx = part.indexOf(':');
      if (idx === -1) return;
      const type = part.substring(0, idx).trim().toLowerCase();
      const names = part.substring(idx + 1).split(',').map(function (x) { return x.trim(); }).filter(function (x) { return x; });
      if (!result[type]) result[type] = {};
      names.forEach(function (n) { result[type][n] = true; });
    });
  } else {
    // Format lama (sebelum ada breakdown per jenis) -- masukkan ke LAINNYA,
    // bukan dibuang.
    s.split(',').map(function (x) { return x.trim(); }).filter(function (x) { return x; })
      .forEach(function (n) { result.lainnya[n] = true; });
  }
  return result;
}

/**
 * Menggabungkan AddWho yang sudah tersimpan (string, format baru atau lama) dengan user WMS
 * dari upload baru (object { move:{}, picking:{}, lainnya:{} }). Dipakai importRawData saat
 * item digabung ke tugas Pending yang sudah ada, supaya nama lama tidak hilang.
 */
function mergeAddWho_(existingStr, newByType) {
  const merged = decodeAddWhoByType_(existingStr);
  ['move', 'picking', 'lainnya'].forEach(function (t) {
    Object.keys((newByType && newByType[t]) || {}).forEach(function (n) {
      if (!merged[t]) merged[t] = {};
      merged[t][n] = true;
    });
  });
  return encodeAddWhoByType_(merged);
}

/**
 * v8.29.0: Ambil Spreadsheet ID dari input yang bisa berupa ID mentah ATAU
 * URL Google Sheets
 * lengkap (mis. 'https://docs.google.com/spreadsheets/d/ABC123.../edit#gid=0'). Dipakai saat
 * mendaftarkan facility dari spreadsheet yang sudah ada
 * (daftarkanFacilityExisting) supaya
 * admin bisa tinggal paste link, gak perlu potong-potong ID manual.
 */
function extractSpreadsheetId_(idAtauUrl) {
  const s = String(idAtauUrl || '').trim();
  if (!s) return '';
  const match = s.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  if (match) return match[1];
  // Bukan URL -- anggap sudah berupa ID mentah, cuma bersihkan whitespace
  return s;
}

function extractLevel_(lokasi) {
  const parts = String(lokasi || '').split('.');
  const last = parts[parts.length - 1] || '';
  const match = last.match(/\d+/);
  if (!match) return null;
  return parseInt(match[0], 10);
}

/**
 * v8.10.0 -- Rollcage: format lokasi levelnya bukan angka biasa, ada huruf K
 * (mis.
 * 'A03.23.K09'). Rollcage TIDAK butuh alat bantu (setara Level 1-2 / Tanpa
 * Alat Bantu).
 * Kalau dibiarkan lewat extractLevel_() biasa, angka setelah K (mis. 9)
 * kebaca sebagai
 * level tinggi, dan getLevelGroup_() punya fallback diam-diam ke grup
 * TERAKHIR (Reach
 * Truck) kalau levelnya gak match range manapun -- jadi rollcage bisa salah
 * ke-assign
 * sebagai tugas Reach Truck kalau tidak di-exception-kan di sini.
 */
function isRollcageLocation_(lokasi) {
  const parts = String(lokasi || '').split('.');
  const last = parts[parts.length - 1] || '';
  return /K/i.test(last);
}

/**
 * v8.10.0 -- Pengganti pola "getLevelGroup_(extractLevel_(lokasi))" yang
 * dipakai berulang
 * di ImportData.gs. PAKAI FUNGSI INI (bukan getLevelGroup_ langsung) di
 * manapun grouping
 * alat/effort dihitung dari sebuah lokasi, supaya lokasi Rollcage selalu
 * konsisten masuk
 * grup Tanpa Alat Bantu di seluruh codebase, bukan cuma di satu tempat.
 */
function getLevelGroupByLocation_(lokasi) {
  if (isRollcageLocation_(lokasi)) return LEVEL_GROUPS[0]; // Tanpa Alat Bantu
  return getLevelGroup_(extractLevel_(lokasi));
}

/**
 * v8.10.0 -- Area gudang murni dari HURUF AWAL kode lokasi, dipakai buat
 * dropdown filter
 * Area di app counting (lihat getMyPendingTasks di CycleCount.gs).
 * Pengecualian: prefix
 * "WA"/"WB"/dst ("Wall" -- racking yang nempel ke tembok) BUKAN area
 * tersendiri, tapi
 * ikut gudang yang hurufnya mengikuti W (WA->Area A, WB->Area B). Lokasi
 * Rollcage
 * (A03.23.K09) tetap kebaca Area A seperti biasa -- huruf gudang di depan
 * tidak berubah,
 * cuma segmen level-nya yang formatnya beda.
 */
function extractArea_(lokasi) {
  const raw = String(lokasi || '').trim().toUpperCase();
  const m = raw.match(/^([A-Z]+)/);
  if (!m) return '';
  const alpha = m[1];
  if (alpha.length >= 2 && alpha.charAt(0) === 'W') return alpha.charAt(1);
  return alpha.charAt(0);
}

function parseLocation_(lokasi) {
  const raw = String(lokasi || '').trim().toUpperCase();
  const parts = raw.split('.');
  const first = parts[0] || '';
  const firstMatch = first.match(/^([A-Z]+)?(\d+)?/);
  const gudang = firstMatch && firstMatch[1] ? firstMatch[1] : '';
  const lorong = firstMatch && firstMatch[2] ? parseInt(firstMatch[2], 10) : 999999;

  const sectionRaw = parts.length > 1 ? parts[1] : '';
  const sectionMatch = sectionRaw.match(/\d+/);
  const section = sectionMatch ? parseInt(sectionMatch[0], 10) : 999999;
  const level = extractLevel_(raw);

  return {
    raw: raw, gudang: gudang, lorong: lorong, section: section,
    level: level === null ? 999999 : level
  };
}

/**
 * Urutan: gudang -> lorong -> section (naik lurus, TIDAK dipisah ganjil/
 * genap dulu) -> level.
 * Section yang berhadapan (mis. 066 & 067) jadi otomatis berurutan, jadi 1
 * user bisa cycle
 * kedua sisi lorong sambil jalan maju terus tanpa mondar-mandir balik ke
 * section awal.
 */
function compareLocation_(a, b) {
  const la = a.locationSort || parseLocation_(a.lokasi);
  const lb = b.locationSort || parseLocation_(b.lokasi);
  if (la.gudang !== lb.gudang) return la.gudang < lb.gudang ? -1 : 1;
  if (la.lorong !== lb.lorong) return la.lorong - lb.lorong;
  if (la.section !== lb.section) return la.section - lb.section;
  if (la.level !== lb.level) return la.level - lb.level;
  return la.raw < lb.raw ? -1 : la.raw > lb.raw ? 1 : 0;
}

/**
 * Sama seperti compareLocation_ (gudang -> lorong -> section -> level), TAPI
 * khusus urutan
 * section-nya: GENAP dulu naik (2,4,6,...), baru GANJIL naik (1,3,5,...) --
 * dipakai untuk
 * Level 5-6 (Reach Truck) DAN Level 3-4 (Tangga Pesawat), lihat compareFn di
 * buildAssignedRows_/Tahap 1 & Tahap 2 di buildAssignedRowsEquipmentAware_.
 * Alasannya:
 * kedua alat ini lebih efisien jalan LURUS dulu di satu sisi lorong (mis.
 * semua section
 * genap), baru balik ke sisi satunya (ganjil) -- bukan zig-zag kiri-kanan
 * tiap section
 * kayak compareLocation_ biasa (yang masih dipakai buat Level 1-2/Bawah,
 * petugas jalan
 * kaki tanpa alat yang perlu manuver).
 * v8.13.0: sebelumnya cuma dipakai Reach Truck (nama lama:
 * compareLocationReachTruck_).
 * sekarang dipakai juga buat Tangga.
 */
function compareLocationEvenOddSection_(a, b) {
  const la = a.locationSort || parseLocation_(a.lokasi);
  const lb = b.locationSort || parseLocation_(b.lokasi);
  if (la.gudang !== lb.gudang) return la.gudang < lb.gudang ? -1 : 1;
  if (la.lorong !== lb.lorong) return la.lorong - lb.lorong;
  const parityA = la.section % 2 === 0 ? 0 : 1;
  const parityB = lb.section % 2 === 0 ? 0 : 1;
  if (parityA !== parityB) return parityA - parityB;
  if (la.section !== lb.section) return la.section - lb.section;
  if (la.level !== lb.level) return la.level - lb.level;
  return la.raw < lb.raw ? -1 : la.raw > lb.raw ? 1 : 0;
}

function getLevelGroup_(level) {
  if (level === null || level === undefined || isNaN(level) || level < 1)
    return LEVEL_GROUPS[0];
  for (let i = 0; i < LEVEL_GROUPS.length; i++) {
    const g = LEVEL_GROUPS[i];
    if (level >= g.min && level <= g.max) return g;
  }
  return LEVEL_GROUPS[LEVEL_GROUPS.length - 1];
}

/**
 * Jam dari HP (epoch ms) → Date, atau null bila tidak masuk akal: di masa depan (lebih dari 2 menit)
 * atau lebih tua dari 12 jam. Jam HP yang salah setel tidak boleh mengacaukan tanggal hitung.
 */
function waktuKlienSah_(ms) {
  const t = Number(ms);
  if (!t || isNaN(t)) return null;
  const now = new Date().getTime();
  if (t > now + 2 * 60000 || t < now - 12 * 3600000) return null;
  return new Date(t);
}

/** Waktu kejadian untuk satu baris antrean: jam dari HP bila ada, kalau tidak jam baris itu diterima server. */
function waktuBarisAntrean_(waktuHitung, timestampAntrean) {
  const isDate = function (v) { return Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime()); };
  if (isDate(waktuHitung)) return waktuHitung;
  if (isDate(timestampAntrean)) return timestampAntrean;
  return new Date();
}

function rowDateTag_(row) {
  const waktuCycle = row[15];
  if (Object.prototype.toString.call(waktuCycle) === '[object Date]') {
    return Utilities.formatDate(waktuCycle, 'Asia/Jakarta', 'yyyy-MM-dd');
  }
  const tglRaw = row[1];
  return Object.prototype.toString.call(tglRaw) === '[object Date]'
    ? Utilities.formatDate(tglRaw, 'Asia/Jakarta', 'yyyy-MM-dd')
    : String(tglRaw);
}

/**
 * Fase 22: parsing JSON yang aman -- dipakai untuk kolom Petugas_JSON/
 * Validator_JSON/
 * Kategori_JSON di sheet Ringkasan_Harian. Kalau string kosong/rusak (mis.
 * sel sempat
 * diedit manual), balik ke 'fallback' daripada bikin seluruh request gagal.
 */
function parseJsonSafe_(str, fallback) {
  if (!str) return fallback;
  try {
    return JSON.parse(str);
  } catch (e) {
    return fallback;
  }
}

/**
 * v8.27.0: Progress reporting sederhana pakai CacheService -- dipakai proses
 * panjang yang
 * client-nya butuh circular progress ring dengan persentase (Import Data
 * Admin, Tambah
 * Facility Baru). Client generate jobId unik per klik, backend lapor % di
 * titik-titik
 * checkpoint lewat reportProgress_(), client polling ringan tiap ~700ms lewat
 * getImportProgress(jobId). Murni instrumentasi tambahan -- TIDAK mengubah
 * logic/hasil
 * proses utama sama sekali; kalau CacheService gagal (jarang terjadi),
 * diamkan saja
 * tidak sampai menggagalkan proses utama.
 */
function reportProgress_(jobId, percent, label) {
  if (!jobId) return;
  try {
    CacheService.getScriptCache().put(
      'progress_' + jobId,
      JSON.stringify({ percent: percent, label: label || '' }),
      300 // expire 5 menit, cukup untuk proses upload/tambah facility terlambat sekalipun
    );
  } catch (e) { /* sengaja diabaikan */ }
}

function getImportProgress(jobId) {
  if (!jobId) return { percent: 0, label: '' };
  try {
    const raw = CacheService.getScriptCache().get('progress_' + jobId);
    return raw ? JSON.parse(raw) : { percent: 0, label: '' };
  } catch (e) {
    return { percent: 0, label: '' };
  }
}

/**
 * FASE 5 (v8.3): data tren untuk grafik Dashboard -- perbandingan beberapa
 * periode terakhir
 * SEKALIGUS (default 7 hari / 6 bulan tergantung periodType), bukan cuma 1
 * snapshot seperti
 * getSummaryByPeriod. PENTING soal performa (concern yang sudah dibahas
 * sebelumnya): ini
 * SATU KALI scan Riwayat yang langsung mengelompokkan ke semua periode dalam
 * satu pass
 * (pakai object lookup per baris), BUKAN loop yang memanggil
 * getSummaryByPeriod N kali --
 * jadi biayanya sama seperti 1x getSummaryByPeriod, tidak peduli count
 * berapa.
 */

/**
 * v8.31.0: helper generik 'tulis ulang sheet queue dengan HANYA baris yang
 * belum sempat
 * diproses' -- dipakai bersama oleh SubmitQueue.gs, ValidasiQueue.gs,
 * TaskQueue.gs supaya
 * tidak duplikasi logic clear+rewrite di 3 tempat. 1x clearContent + 1x
 * write (kalau ada
 * sisa), jauh lebih murah dibanding deleteRows satu-satu.
 */

/**
 * v8.31.1: SUDAH TIDAK DIPAKAI -- rawan menghapus baris baru yang masuk PAS
 * Worker sedang
 * jalan (lihat bug fix di deleteProcessedQueueRows_ di bawah). Dibiarkan ada
 * (bukan dihapus)
 * sesuai prinsip perubahan additive, TAPI JANGAN dipanggil lagi di kode baru.
 */
function rewriteQueueSheetGeneric_(sheet, keepRows, numCols) {
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, numCols).clearContent();
  }
  if (keepRows.length) {
    sheet.getRange(2, 1, keepRows.length, numCols).setValues(keepRows);
  }
}

/**
 * v8.31.1 (FIX bug konkurensi @ 15 akun bersamaan): appendRow() BUKAN
 * operasi atomic di Apps
 * Script -- 2 eksekusi simultan bisa sama-sama menyimpulkan 'baris kosong
 * berikutnya = row X'
 * dan saling timpa, salah satu appendRow-nya HILANG tanpa error. Ini
 * penyebab bug 'task sudah
 * disubmit tapi tidak pernah tersembunyi' -- payload-nya gak pernah sampai
 * ke antrian sama
 * sekali. Helper ini membungkus appendRow ke SATU baris dengan LockService.
 * getScriptLock()
 * (lock SANGAT SEBENTAR -- cuma durasi 1x appendRow ke spreadsheet BOUND
 * cepat, BUKAN durasi buka spreadsheet facility -- jadi TIDAK mengembalikan
 * masalah lambat
 * yang sudah diperbaiki sebelumnya). LockService.getScriptLock() ini SAMA
 * dengan yang dipakai
 * acquireFacilityLock_ (FacilityManagement.gs) -- aman dipakai bersama
 * karena keduanya cuma
 * menggenggam sebentar lalu langsung release.
 */
function appendRowLocked_(sheet, rowValues) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    throw new Error('Sistem sedang sibuk, coba lagi sebentar.');
  }
  try {
    sheet.appendRow(rowValues);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Versi batch dari appendRowLocked_ -- dipakai kalau perlu tulis >1 baris
 * sekaligus dalam
 * SATU genggaman lock (mis. queueCounterIncrements_ yang bisa nitip 2 baris
 * counter sekaligus).
 */
function appendRowsLocked_(sheet, rowsArray) {
  if (!rowsArray || !rowsArray.length) return;
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    throw new Error('Sistem sedang sibuk, coba lagi sebentar.');
  }
  try {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rowsArray.length, rowsArray[0].length).setValues(rowsArray);
  } finally {
    lock.releaseLock();
  }
}

/**
 * v8.31.1 (FIX bug konkurensi, lihat catatan appendRowLocked_ di atas):
 * pengganti
 * rewriteQueueSheetGeneric_ yang AMAN terhadap baris baru yang masuk SAAT
 * Worker sedang
 * memproses. rewriteQueueSheetGeneric_ (LAMA) baca-semua -> clear-semua ->
 * tulis-ulang-sisa --
 * kalau ada appendRow BARU masuk di antara baca & clear, baris baru itu ikut
 * kehapus tanpa
 * sempat diproses (data hilang). Versi ini HANYA menghapus baris-baris
 * SPESIFIK yang memang
 * sudah dipastikan selesai diproses (identifikasi by ACTUAL row number di
 * sheet, dihapus dari
 * BAWAH ke ATAS supaya nomor baris lain tidak ikut geser di tengah proses),
 * TIDAK PERNAH
 * menyentuh baris yang belum diproses ATAU baris baru yang masuk belakangan.
 * @param {Sheet} sheet
 * @param {Array<number>} processedRowNumbers - nomor baris AKTUAL (1-based,
 * sesuai sheet) yang
 *   sudah selesai diproses (sukses ATAU gagal permanen) dan boleh dihapus
 *   dari antrian.
 */
function deleteProcessedQueueRows_(sheet, processedRowNumbers) {
  if (!processedRowNumbers || !processedRowNumbers.length) return;
  // Urutkan DESCENDING (baris terbawah dulu) -- supaya deleteRows()
  // satu-satu TIDAK menggeser
  // nomor baris lain yang belum dihapus di iterasi berikutnya.
  const sorted = processedRowNumbers.slice().sort(function (a, b) { return b - a; });
  sorted.forEach(function (rowNum) {
    try { sheet.deleteRow(rowNum); } catch (e) { /* baris mungkin sudah tergeser/terhapus proses lain -- abaikan */ }
  });
}
