/**
 * TaskInvestigasi.gs
 * Alur Task Investigasi (status bertahap, prioritas, lsg)
 * (Hasil pemecahan dari Code.gs -- semua file .gs berbagi global scope yang
 * sama di Apps Script,
 * jadi urutan file tidak masalah, fungsi di file lain tetap bisa saling
 * panggil.)
 */

// --------- Task Investigasi (role: inventory/admin) ---------

function canManageTasks_(_username) {
  const info = getUserRole(_username);
  // v8.26.0: Developer juga bisa manage tasks (bisa cek semua keadaan)
  return !!info && (info.role === 'admin' || info.role === 'inventory' || info.role === 'developer');
}

/**
 * Isi getOpenTasks TANPA cek role -- dipisah supaya bisa dipakai ulang oleh
 * getPrioritasHariIni (v8) tanpa duplikasi logika aging/critical/sort.
 * 
 * v8.26.0: Menerima parameter opsional username untuk baca dari facility
 * yang benar.
 */
function getOpenTasksData_(username) {
  // v8.26.0: Meneruskan username agar baca dari spreadsheet facility user
  const sheet = getRiwayatSheet_(username);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  // Semua status KECUALI 'Selesai' dianggap masih aktif/terbuka.
  const activeStatuses = TASK_STATUS_FLOW.slice(0, TASK_STATUS_FLOW.length - 1);
  const now = new Date();

  const data = sheet.getRange(2, 1, lastRow - 1, RIWAYAT_HEADERS.length).getValues();
  const tasks = [];
  data.forEach(function (row) {
    const statusTask = row[18];
    if (activeStatuses.indexOf(statusTask) === -1) return;

    const waktuCycle = row[15];
    // Aging dihitung LANGSUNG dari Waktu_Cycle (kapan discrepancy pertama terjadi) --
    // TIDAK ada kolom Due_Date tersendiri, jadi SLA cukup ganti TASK_SLA_HARI di atas.
    const umurHari = (Object.prototype.toString.call(waktuCycle) === '[object Date]')
      ? Math.floor((now - waktuCycle) / 86400000) : 0;
    const selisihAbs = Math.abs(Number(row[9]) || 0);
    // FASE 4: skor gabungan umur+qty -- dipakai sebagai tie-breaker di dalam tier Critical/
    // Non-Critical yang sama, supaya task dengan qty selisih besar tidak terkubur di bawah
    // task lain yang cuma unggul di umur 1 hari.
    const skorPrioritas = (umurHari * PRIORITAS_BOBOT_UMUR) + selisihAbs;

    tasks.push({
      id: row[0], tanggal: row[1], lokasi: row[2], article: row[3],
      description: row[4],
      selisih: row[9], namaPetugas: row[7], kategori: row[17], alasan: row[17],
      statusTask: statusTask, picInvestigasi: row[23], addWhoTransaksi: row[25],
      namaValidator: row[12], nameValidator: row[12], // nameValidator: ejaan lama, dipertahankan untuk klien lama
      umurHari: umurHari, isCritical: umurHari > TASK_SLA_HARI,
      skorPrioritas: skorPrioritas
    });
  });

  // Critical dulu (masih SLA-based, ini tetap penting), lalu di dalam tier yang sama
  // diurutkan dari skorPrioritas tertinggi (kombinasi umur+qty, bukan cuma umur).
  tasks.sort(function (a, b) {
    if (a.isCritical !== b.isCritical) return a.isCritical ? -1 : 1;
    return b.skorPrioritas - a.skorPrioritas;
  });
  return tasks;
}

function getOpenTasks(requesterUsername) {
  requireRole_(requesterUsername, ['inventory', 'admin', 'developer']);
  // v8.26.0: Meneruskan requesterUsername agar baca dari facility yang benar
  return getOpenTasksData_(requesterUsername);
}

function getOpenTaskCount(requesterUsername) {
  requireRole_(requesterUsername, ['inventory', 'admin', 'developer']);
  // v8.26.0: Meneruskan requesterUsername agar baca dari facility user
  const sheet = getRiwayatSheet_(requesterUsername);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  const activeStatuses = TASK_STATUS_FLOW.slice(0, TASK_STATUS_FLOW.length - 1);

  const col = sheet.getRange(2, 19, lastRow - 1, 1).getValues();
  let count = 0;
  col.forEach(function (r) { if (activeStatuses.indexOf(r[0]) !== -1) count++; });
  return count;
}

// NOTE (v8.18.0): getPrioritasHariIni() (widget Home "🚨 Prioritas Hari Ini") dihapus --
// digantikan getPlusMinusSummary() (lihat HomeSummary.gs) untuk widget baru "Summary Plus Minus".
// getOpenTasksData_() TETAP ada, masih dipakai getOpenTasks() untuk list utama layer Verifikasi.

/**
 * v8.18.0: memvalidasi bukti transaksi WMS (Move/Picking) yang di-paste user
 * -- WAJIB saat
 * * menutup task dengan kategori BUKAN 'Salah Hitung' (kalau Salah Hitung, tidak perlu bukti
 * * fisik karena memang bukan discrepancy nyata). Aturan (kategori Move/Picking, spt sebelumnya):
 * 1. Harus ada minimal 1 baris dengan SKU yang cocok dengan Article task ini, DAN jenis
 * * transaksinya MOVE atau PICKING (jenis lain diabaikan, tidak dianggap bukti valid).
 * 2. Total Qty dari baris yang cocok HARUS SAMA PERSIS dengan besar selisih task ini
 * * (dibulatkan ke bilangan bulat terdekat -- Qty barang selalu bilangan bulat).
 * 3. Minimal 1 baris yang cocok harus mencantumkan nomor lokasi (FROMLOC atau TOLOC).
 * 
 * v8.21: kalau kategori-nya 'Adjustment Plus' / 'Adjustment Minus' (lihat KATEGORI_ADJUSTMENT_) --
 * aturannya BEDA -- ini mengonfirmasi barang PLUS/MINUS BENERAN (bukan salah lokasi):
 * 1. Jenis transaksi bukti WAJIB 'ADJUSTMENT' (bukan Move/Picking).
 * 2. Qty dicek SESUAI TANDA (bukan abs) -- 'Adjustment Minus' wajib total qty = -selisihAbs
 *    (mis. minus 1 -> bukti harus -1), 'Adjustment Plus' wajib total qty = +selisihAbs.
 * 3. Minimal 1 baris yang cocok harus mencantumkan TOLOC yang SAMA dengan lokasi task ini
 *    (adjustment TIDAK PUNYA FROMLOC di WMS -- beda dari Move/Picking yang punya keduanya --
 *    jadi cuma TOLOC yang dicek, bukan FROMLOC/TOLOC seperti kategori Move/Picking).
 * 
 * Urutan kolom buktiRows (per baris, array), SESUAI urutan LOG_BUKTI_HEADERS setelah 5 kolom
 * konteks: [NO, STORERKEY, TRANTYPE, SKU, Description, SKUGROUP, LOT, FROMLOC, TOLOC,
 * TOID, SOURCEKEY, QTY, ADDATE, ADOWHO] -- index 0..14.
 */
function validateBuktiTransaksi_(buktiRows, article, selisihAbs, kategori, lokasiTask) {
  const isAdjustment = KATEGORI_ADJUSTMENT_.indexOf(String(kategori || '').trim()) !== -1;
  const jenisLabel = isAdjustment ? 'Adjustment' : 'Move/Picking';

  if (!buktiRows || !buktiRows.length) {
    return { valid: false, message: 'Bukti transaksi WMS (' + jenisLabel + ') wajib di-paste untuk menutup task ini (kecuali kategori Salah Hitung / Barang Sudah di Picking).' };
  }

  const articleNorm = String(article || '').trim().toUpperCase();

  if (isAdjustment) {
    const relevant = buktiRows.filter(function (r) {
      const sku = String(r[3] || '').trim().toUpperCase();
      const tipe = String(r[2] || '').trim().toUpperCase();
      return sku === articleNorm && tipe === 'ADJUSTMENT';
    });

    if (!relevant.length) {
      return { valid: false, message: 'Tidak ada baris bukti dengan SKU ' + article + ' dan jenis transaksi ADJUSTMENT. Kategori ini wajib bukti transaksi ADJUSTMENT dari WMS.' };
    }

    const totalQtySigned = relevant.reduce(function (s, r) { return s + (Number(r[12]) || 0); }, 0);
    const expectedSigned = (String(kategori).trim() === 'Adjustment Plus') ? selisihAbs : -selisihAbs;

    if (Math.round(totalQtySigned) !== Math.round(expectedSigned)) {
      return {
        valid: false,
        message: 'Total Qty adjustment (' + totalQtySigned + ') harus tepat ' + expectedSigned + ' sesuai kategori ' + kategori + ' & selisih task ini. Qty di data WMS harus sesuai tanda (minus utk Adjustment Minus, plus utk Adjustment Plus).'
      };
    }

    const lokasiNorm = String(lokasiTask || '').trim().toUpperCase();
    const hasLokasiMatch = relevant.some(function (r) {
      return String(r[9] || '').trim().toUpperCase() === lokasiNorm;
    });

    if (!hasLokasiMatch) {
      return { valid: false, message: 'Bukti adjustment harus mencantumkan TOLOC yang sama dengan lokasi task ini (' + lokasiTask + ').' };
    }

    return { valid: true };
  }

  const relevant = buktiRows.filter(function (r) {
    const sku = String(r[3] || '').trim().toUpperCase();
    const tipe = String(r[2] || '').trim().toUpperCase();
    return sku === articleNorm && (tipe === 'MOVE' || tipe === 'PICKING');
  });

  if (!relevant.length) {
    return { valid: false, message: 'Tidak ada baris bukti dengan SKU ' + article + ' dan jenis transaksi MOVE/PICKING. Cek lagi data yang di-paste.' };
  }

  const totalQty = relevant.reduce(function (s, r) { return s + Math.abs(Number(r[12]) || 0); }, 0);
  if (Math.round(totalQty) !== Math.round(selisihAbs)) {
    return { valid: false, message: 'Total Qty di bukti transaksi (' + totalQty + ') tidak sama dengan besar selisih task ini (' + selisihAbs + '). Qty harus sama persis.' };
  }

  const hasLokasi = relevant.some(function (r) { return String(r[7] || '').trim() || String(r[9] || '').trim(); });
  if (!hasLokasi) {
    return { valid: false, message: 'Bukti transaksi harus mencantumkan nomor lokasi (FROMLOC/TOLOC).' };
  }

  return { valid: true };
}

/**
 * Menggantikan closeTask (v6) -- sekarang task punya beberapa tahap (TASK_STATUS_FLOW),
 * bukan cuma Open/Selesai. Dipanggil tiap kali user memajukan status task 1 langkah (atau
 * lebih, boleh loncat, TIDAK BOLEH mundur/ulang status yang sama).
 * 
 * newStatus = WAJIB, salah satu dari 'Sedang Dicari' / 'Menunggu Konfirmasi' / 'Selesai'.
 * dan index-nya di TASK_STATUS_FLOW harus lebih besar dari status saat ini.
 * catatan = WAJIB di SETIAP transisi (bukan cuma saat Selesai) -- dipakai sebagai log
 * progres akumulatif di kolom Catatan_Penyelesaian (setiap entri baru ditambahkan,
 * bukan menimpa, jadi riwayat pencarian barang tetap kebaca).
 * kategori/picUsername = WAJIB kalau newStatus = 'Selesai' (final close, sama seperti
 * closeTask lama). Untuk transisi ANTARA (Sedang Dicari/Menunggu Konfirmasi), keduanya
 * OPSIONAL -- kalau diisi, dipakai untuk mengoreksi kategori/PIC begitu root-cause
 * sebenarnya ketahuan (PIC awal cuma tebakan round-robin, lihat assignInitialPIC..).
 * buktiRows = v8.18.0 -- WAJIB kalau newStatus='Selesai' DAN kategori BUKAN 'Salah Hitung'.
 * Array baris bukti transaksi WMS yang di-paste user (lihat validateBuktiTransaksi_ di atas
 * untuk urutan kolom & aturan validasinya). Diabaikan kalau kategori='Salah Hitung'.
 * 
 * v8.31.0 (Asynchronous Queue Processing): signature bertambah 3 parameter opsional di akhir
 * (lokasiTask, articleTask, selisihTask) -- dikirim CLIENT dari data yang SUDAH ada di layar
 * (hasil getOpenTasks() sebelumnya), supaya validasi bukti WMS (validateBuktiTransaksi_) --
 * fungsi MURNI, tidak baca sheet sama sekali) tetap bisa jalan SINKRON & INSTAN di sini
 * TANPA perlu buka spreadsheet facility dulu. Penulisan sesungguhnya (update status, tulis
 * Log_Bukti_Investigasi, Ringkasan delta) dipindah ke Worker (TaskQueue.gs) -- yang JUGA
 * validasi ULANG bukti pakai data SEGAR dari sheet sebelum benar-benar menutup task
 * (defense-in-depth, client tidak dipercaya buta untuk keputusan final).
 */
function updateTaskStatus(id, namaUser, newStatus, catatan, kategori, picUsername, buktiRows, lokasiTask, articleTask, selisihTask) {
  const catatanTrim = String(catatan || '').trim();
  const newStatusTrim = String(newStatus || '').trim();
  const kategoriTrim = String(kategori || '').trim();
  const picTrim = String(picUsername || '').trim();

  if (!canManageTasks_(namaUser)) {
    return { success: false, message: 'Tidak punya akses update task (khusus role Inventory/Admin).' };
  }

  const targetIdx = TASK_STATUS_FLOW.indexOf(newStatusTrim);
  if (targetIdx <= 0) {
    return { success: false, message: 'Status tujuan tidak valid.' };
  }

  if (!catatanTrim) return { success: false, message: 'Catatan progres wajib diisi.' };

  const isFinal = newStatusTrim === 'Selesai';
  if (isFinal) {
    if (TASK_KATEGORI_LIST.indexOf(kategoriTrim) === -1) {
      return { success: false, message: 'Kategori selisih wajib dipilih (Lebih/Kurang Picking, Lebih/Kurang Move, Adjustment Plus/Minus, Salah Hitung, atau Barang Sudah di Picking).' };
    }
    if (!picTrim) return { success: false, message: 'User PIC wajib dipilih.' };
    if (!getUserRole(picTrim)) {
      return { success: false, message: 'User PIC "' + picTrim + '" tidak dikenali di Master.' };
    }

    // v8.18.0 (Feedback tetap instan meski penulisan sekarang async): bukti WMS divalidasi
    // di sini pakai lokasi/article/selisih yang dikirim client (sudah dimiliki client dari
    // getOpenTasks() -- lihat catatan signature di atas).
    if (KATEGORI_TANPA_BUKTI_.indexOf(kategoriTrim) === -1) {
      const selisihAbs = Math.abs(Number(selisihTask) || 0);
      const cek = validateBuktiTransaksi_(buktiRows, articleTask, selisihAbs, kategoriTrim, lokasiTask);
      if (!cek.valid) return { success: false, message: cek.message };
    }
  }

  const facInfo = requireUserFacility_(namaUser);
  if (!facInfo) {
    return { success: false, message: PESAN_TANPA_FACILITY_ };
  }

  // v8.31.0 (Asynchronous Queue Processing, lihat TaskQueue.gs): penulisan sesungguhnya
  // dipindah ke Worker -- di sini cukup titip payload ke Queue_Task_Investigasi (spreadsheet
  // BOUND, bukan openById) lalu langsung return.
  try {
    queueTaskUpdatePayload_(id, namaUser, newStatusTrim, catatanTrim, kategoriTrim, picTrim, buktiRows, facInfo.id);
  } catch (e) {
    return { success: false, message: 'Gagal masuk antrian: ' + e.message };
  }

  return { success: true, queued: true, message: 'Update diterima.' };
}

// --------- v8.21: Penyelesaian Plus Minus (pairing lintas lokasi) ---------
// 
// Kasus: task Minus di Lokasi A & task Plus di Lokasi B, SKU & qty sama persis -- ini kejadian
// fisik yang SAMA (barang misplace/salah taruh), bukan 2 kejadian terpisah.
// Alih-alih user paste bukti MOVE yang sama 2x di 2 form Task Investigasi berbeda (rawan
// tidak konsisten), menu ini memasangkan KEDUA task sekaligus & menutupnya BERSAMA dengan
// HANYA 1 bukti MOVE yang eksplisit menghubungkan lokasi A (FROMLOC) -> lokasi B (TOLOC).
// Pairing WAJIB 1-ke-1: |selisih| task Minus harus SAMA PERSIS dengan |selisih| task Plus.

/**
 * Daftar task terbuka (belum Selesai, Hasil_Final=DISCREPANCY) yang jadi kandidat pairing.
 * Dipisah jadi 2 sisi: minus[] (selisih < 0) & plus[] (selisih > 0). Dipakai layar
 * "Penyelesaian Plus Minus" supaya user tinggal pilih 1 dari tiap sisi (SKU sama & qty sama).
 */
function getPlusMinusCandidates(requesterUsername) {
  requireRole_(requesterUsername, ['inventory', 'admin', 'developer']);
  // v8.26.2: Meneruskan requesterUsername agar baca dari spreadsheet facility user
  const all = getOpenTasksData_(requesterUsername); // sudah exclude Status_Task='Selesai'
  const sheet = getRiwayatSheet_(requesterUsername);
  const lastRow = sheet.getLastRow();
  const plus = [], minus = [];
  if (lastRow < 2) return { plus: plus, minus: minus };

  // getOpenTasksData_ tidak menyertakan Hasil_Final, jadi ambil ulang kolom itu sekalian
  // supaya kandidat pairing cuma yang benar-benar discrepancy (bukan task lain yang kebetulan
  // masih berstatus aktif tapi bukan hasil count, kalaspun ada).
  const hasilFinalById = {};
  const data = sheet.getRange(2, 1, lastRow - 1, RIWAYAT_HEADERS.length).getValues();
  data.forEach(function (row) { hasilFinalById[row[0]] = row[14]; });

  all.forEach(function (t) {
    if (hasilFinalById[t.id] !== 'DISCREPANCY') return;
    const s = Number(t.selisih) || 0;
    if (s > 0) plus.push(t);
    else if (s < 0) minus.push(t);
  });
  return { plus: plus, minus: minus };
}

/**
 * Menutup SEKALIGUS 2 task (1 Minus + 1 Plus) yang dipasangkan, dengan 1 bukti MOVE yang
 * menghubungkan lokasi keduanya. Aturan:
 * * Kedua task masih aktif (belum Selesai), SKU (Article) sama persis.
 * * |selisih| kedua task HARUS sama persis (pairing 1-ke-1, tidak boleh selisih qty).
 * * Bukti WAJIB TRANTYPE MOVE, SKU cocok, total Qty = |selisih|, DAN ada baris yang FROMLOC-nya
 *   = lokasi task Minus SEKALIGUS TOLOC-nya = lokasi task Plus (memastikan buktinya benar2
 *   menghubungkan lokasi A ke lokasi B, bukan bukti MOVE acak yang qty-nya kebetulan sama).
 * * Kategori_Selisih kedua baris di-set ke KATEGORI_PLUS_MINUS_PAIR_, dan kolom baru
 *   Pasangan_Task_ID saling menunjuk ID satu sama lain (audit trail).
 */
function closePlusMinusPair(minusTaskId, plusTaskId, namaUser, catatan, picUsername, buktiRows) {
  const catatanTrim = String(catatan || '').trim();
  const picTrim = String(picUsername || '').trim();

  if (!canManageTasks_(namaUser)) {
    return { success: false, message: 'Tidak punya akses (khusus role Inventory/Admin).' };
  }

  if (!minusTaskId || !plusTaskId) {
    return { success: false, message: 'Task Minus & task Plus wajib dipilih.' };
  }

  if (String(minusTaskId) === String(plusTaskId)) {
    return { success: false, message: 'Task Minus & Plus tidak boleh task yang sama.' };
  }

  if (!catatanTrim) return { success: false, message: 'Catatan penyelesaian wajib diisi.' };
  if (!picTrim) return { success: false, message: 'User PIC wajib dipilih.' };
  if (!getUserRole(picTrim)) {
    return { success: false, message: 'User PIC "' + picTrim + '" tidak dikenali di Master.' };
  }

  if (!buktiRows || !buktiRows.length) {
    return { success: false, message: 'Bukti transaksi MOVE dari WMS wajib di-paste untuk menutup pasangan task ini.' };
  }

  // v8.29.0: wajib punya facility valid SEBELUM masuk antrean lock (lihat requireUserFacility_).
  const facInfo = requireUserFacility_(namaUser);
  if (!facInfo) {
    return { success: false, message: PESAN_TANPA_FACILITY_ };
  }

  // v8.29.0: lock PER-FACILITY (bukan lock global) -- proses dari DC lain gak ikut ke-block.
  const facLock = acquireFacilityLock_(facInfo.id);
  if (!facLock) {
    return { success: false, message: 'Sistem sedang sibuk, coba lagi.' };
  }

  try {
    // v8.26.2: Meneruskan namaUser agar baca dari spreadsheet facility user
    const sheet = getRiwayatSheet_(namaUser);
    const rowMinusIdx = findRiwayatRow_(sheet, minusTaskId);
    const rowPlusIdx = findRiwayatRow_(sheet, plusTaskId);
    if (rowMinusIdx === -1 || rowPlusIdx === -1) {
      return { success: false, message: 'Salah satu task tidak ditemukan, kemungkinan sudah diproses.' };
    }

    const rowMinus = sheet.getRange(rowMinusIdx, 1, 1, RIWAYAT_HEADERS.length).getValues()[0];
    const rowPlus = sheet.getRange(rowPlusIdx, 1, 1, RIWAYAT_HEADERS.length).getValues()[0];

    if (String(rowMinus[18]) === 'Selesai' || String(rowPlus[18]) === 'Selesai') {
      return { success: false, message: 'Salah satu task sudah Selesai, tidak bisa dipasangkan lagi.' };
    }

    const selisihMinus = Number(rowMinus[9]) || 0;
    const selisihPlus = Number(rowPlus[9]) || 0;
    if (selisihMinus >= 0) return { success: false, message: 'Task Minus yang dipilih ternyata bukan selisih minus.' };
    if (selisihPlus <= 0) return { success: false, message: 'Task Plus yang dipilih ternyata bukan selisih plus.' };

    const articleMinus = String(rowMinus[3] || '').trim().toUpperCase();
    const articlePlus = String(rowPlus[3] || '').trim().toUpperCase();
    if (articleMinus !== articlePlus) {
      return { success: false, message: 'SKU/Article kedua task harus sama (Minus: ' + rowMinus[3] + ', Plus: ' + rowPlus[3] + ').' };
    }

    if (Math.abs(selisihMinus) !== Math.abs(selisihPlus)) {
      return { success: false, message: 'Qty selisih kedua task harus sama persis (pairing 1-ke-1). Minus: ' + Math.abs(selisihMinus) + ', Plus: ' + Math.abs(selisihPlus) + '.' };
    }

    const qtyAbs = Math.abs(selisihMinus);
    const lokasiMinus = String(rowMinus[2] || '').trim().toUpperCase();
    const lokasiPlus = String(rowPlus[2] || '').trim().toUpperCase();

    const relevant = buktiRows.filter(function (r) {
      const sku = String(r[3] || '').trim().toUpperCase();
      const tipe = String(r[2] || '').trim().toUpperCase();
      return sku === articleMinus && tipe === 'MOVE';
    });

    if (!relevant.length) {
      return { success: false, message: 'Tidak ada baris bukti dengan SKU ' + rowMinus[3] + ' dan jenis transaksi MOVE.' };
    }

    const totalQty = relevant.reduce(function (s, r) { return s + Math.abs(Number(r[12]) || 0); }, 0);
    if (Math.round(totalQty) !== Math.round(qtyAbs)) {
      return { success: false, message: 'Total Qty di bukti MOVE (' + totalQty + ') tidak sama dengan qty pasangan ini (' + qtyAbs + ').' };
    }

    const hasLinkedRoute = relevant.some(function (r) {
      return String(r[7] || '').trim().toUpperCase() === lokasiMinus && String(r[9] || '').trim().toUpperCase() === lokasiPlus;
    });

    if (!hasLinkedRoute) {
      return {
        success: false,
        message: 'Bukti MOVE harus punya baris dengan FROMLOC=' + rowMinus[2] + ' (lokasi Minus) dan TOLOC=' + rowPlus[2] + ' (lokasi Plus) -- memastikan bukti ini benar menghubungkan kedua lokasi.'
      };
    }

    const waktu = new Date();
    const stamp = Utilities.formatDate(waktu, 'Asia/Jakarta', 'dd/MM HH:mm');
    const noteSuffix = '\n[' + stamp + ' - Selesai] ' + catatanTrim + '\n[' + stamp + '] Dipasangkan & ditutup via Penyelesaian Plus Minus, bukti MOVE terlampir -- lihat sheet Log_Bukti_Investigasi.';

    function closeRow(rowIdx, pairId, rowSnapshot) {
      const prevCatatan = String(rowSnapshot[19] || '');
      sheet.getRange(rowIdx, 19).setValue('Selesai'); // Status_Task
      sheet.getRange(rowIdx, 20).setValue(prevCatatan + noteSuffix); // Catatan_Penyelesaian
      sheet.getRange(rowIdx, 18).setValue(KATEGORI_PLUS_MINUS_PAIR_); // Kategori_Selisih
      sheet.getRange(rowIdx, 21).setValue(namaUser); // Diselesaikan_Oleh
      sheet.getRange(rowIdx, 22).setValue(waktu); // Waktu_Selesai_Task
      sheet.getRange(rowIdx, 24).setValue(picTrim); // PIC_Investigasi
      sheet.getRange(rowIdx, 25).setValue(waktu); // Waktu_Update_Status
      sheet.getRange(rowIdx, RIWAYAT_COL_PASANGAN_TASK_ID_).setValue(pairId); // Pasangan_Task_ID

      // Fase 22: kembalikan snapshot "newRow" (dipakai buat delta Ringkasan_Harian) -- kloning
      // rowSnapshot lalu terapkan mutasi yang SAMA seperti yang barusan ditulis ke sheet.
      const newRow = rowSnapshot.slice();
      newRow[17] = KATEGORI_PLUS_MINUS_PAIR_; newRow[18] = 'Selesai'; newRow[19] = prevCatatan + noteSuffix;
      newRow[20] = namaUser; newRow[21] = waktu; newRow[23] = picTrim; newRow[24] = waktu; newRow[26] = pairId;
      return newRow;
    }

    const newRowMinus = closeRow(rowMinusIdx, plusTaskId, rowMinus);
    const newRowPlus = closeRow(rowPlusIdx, minusTaskId, rowPlus);

    // v8.26.0: Teruskan namaUser sebagai konteks facility
    try {
      applyRingkasanDelta_(rowDateTag_(rowMinus), rowMinus, newRowMinus, namaUser);
      applyRingkasanDelta_(rowDateTag_(rowPlus), rowPlus, newRowPlus, namaUser);
    } catch (e) { /* lihat rebuildRingkasanHarian_ */ }

    // Fase 23: kedua task (Minus & Plus) keluar dari tracker Plus/Minus aktif -- keduanya
    // PASTI DISCREPANCY (syarat jadi kandidat pairing), jadi tidak perlu cek Hasil_Final lagi.
    try {
      adjustPlusMinusTracker_(rowMinus[9], rowMinus[3], rowMinus[4], rowMinus[2], false, namaUser);
      adjustPlusMinusTracker_(rowPlus[9], rowPlus[3], rowPlus[4], rowPlus[2], false, namaUser);
    } catch (e) { /* lihat rebuildAntrianAktif_ */ }

    // Simpan bukti ke Log_Bukti_Investigasi utk KEDUA task id, supaya audit trail masing-masing
    // task tetap lengkap kalau ditelusuri sendiri-sendiri (bukan cuma nyambung lewat Pasangan_Task_ID).
    // v8.26.0: Meneruskan namaUser agar tulis ke spreadsheet facility user
    const buktiSheet = getLogBuktiSheet_(namaUser);
    const buktiStartRow = buktiSheet.getLastRow() + 1;
    const buktiOut = [];
    buktiRows.forEach(function (r) {
      buktiOut.push([waktu, minusTaskId, rowMinus[2], rowMinus[3], namaUser].concat(r.slice(0, 15)));
    });
    buktiRows.forEach(function (r) {
      buktiOut.push([waktu, plusTaskId, rowPlus[2], rowPlus[3], namaUser].concat(r.slice(0, 15)));
    });
    buktiSheet.getRange(buktiStartRow, 1, buktiOut.length, LOG_BUKTI_HEADERS.length).setValues(buktiOut);

    return { success: true };
  } finally {
    facLock.release();
  }
}

/**
 * v8.20: dateFrom/dateTo (string 'YYYY-MM-DD', opsional) memfilter berdasarkan Waktu_Selesai_Task
 * (kapan task ditutup) -- filter diterapkan SEBELUM slice(limit), supaya rentang tanggal yang
 * dipilih tidak terpotong gara-gara limit. Dipakai oleh Riwayat Log di Dashboard (v8.20 pindah
 * dari screen Task Investigasi, collapsed default).
 */
function getTaskLog(limit, requesterUsername, dateFrom, dateTo) {
  requireRole_(requesterUsername, ['inventory', 'admin', 'developer']);

  // v8.26.2: Meneruskan requesterUsername agar baca dari spreadsheet facility user
  const sheet = getRiwayatSheet_(requesterUsername);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  // Rentang tanggal dibaca sebagai hari kerja WIB, sama seperti semua tanggal lain di sistem ini. Tanpa
  // "+07:00" batasnya mengikuti zona waktu project Apps Script, sehingga task yang ditutup dini hari
  // bisa jatuh ke hari yang salah (bergeser 1 jam di WITA, 7 jam bila zonanya UTC).
  const fromDate = dateFrom ? new Date(dateFrom + 'T00:00:00+07:00') : null;
  const toDate = dateTo ? new Date(dateTo + 'T23:59:59+07:00') : null;

  const data = sheet.getRange(2, 1, lastRow - 1, RIWAYAT_HEADERS.length).getValues();
  const logs = [];
  data.forEach(function (row) {
    const statusTask = row[18];
    const alasan = row[17];
    const hasilFinal = row[14];
    // 'Salah Hitung' AUTO (Hasil_Final=HIT, petugas yang salah hitung, tidak pernah lewat Task
    // Investigasi beneran) tetap disembunyikan dari log ini. 'Salah Hitung' MANUAL (Hasil_Final=
    // DISCREPANCY, validator yang salah hitung, ketahuan & ditutup lewat Task Investigasi
    // beneran) TETAP ditampilkan -- itu investigasi nyata yang perlu tercatat di log.
    if (statusTask !== 'Selesai' || !alasan) return;
    if (alasan === 'Salah Hitung' && hasilFinal === 'HIT') return;

    const waktuSelesaiDate = Object.prototype.toString.call(row[21]) === '[object Date]' ? row[21] : null;
    if (fromDate && (!waktuSelesaiDate || waktuSelesaiDate < fromDate)) return;
    if (toDate && (!waktuSelesaiDate || waktuSelesaiDate > toDate)) return;

    logs.push({
      tanggal: row[1], lokasi: row[2], article: row[3], description: row[4],
      selisih: row[9],
      namaPetugas: row[7], alasan: alasan, kategori: alasan, picUsername: row[23], catatan: row[19], diselesaikanOleh: row[20],
      addWhoTransaksi: row[25], namaValidator: row[12],
      waktuSelesai: waktuSelesaiDate
        ? Utilities.formatDate(waktuSelesaiDate, 'Asia/Jakarta', 'dd/MM/yyyy HH:mm')
        : String(row[21] || '')
    });
  });

  logs.sort(function (a, b) { return String(b.waktuSelesai).localeCompare(String(a.waktuSelesai)); });
  const max = limit ? Number(limit) : 200;
  return logs.slice(0, max);
}
