/**
 * HomeSummary.gs
 * Ringkasan Home & fungsi test
 * (Hasil pemecahan dari Code.gs -- semua file .gs berbagi global scope yang
 * sama di Apps Script.
 * Jadi urutan file tidak masalah, fungsi di file lain tetap bisa saling
 * panggil.)
 */

// ---------- Home summary & badge (ringan, dipakai di UI) ----------

/**
 * Fase 23: DIUBAH -- dulu scan penuh Riwayat, SEKARANG baca 1 baris
 * Ringkasan_Harian (hari
 * ini) + 1-2 baris Antrian_Aktif (counter live). Kecepatan buka Home TIDAK
 * LAGI tergantung
 * berapa banyak baris Riwayat.
 *
 * Untuk role admin: KPI_Total hari ini = totalCycleHariIni (mencakup SEMUA
 * petugas, sudah
 * pas artinya "semua item dihitung hari ini"), KPI_Hit+KPI_Disc =
 * selesaiHariIni (Pending
 * tidak masuk keduanya). Untuk role lain (petugas): perlu breakdown PER
 * ORANG, makanya baca
 * Petugas_Raw_JSON (lihat RingkasanHarian.gs) yang MEMANG didesain utk ini
 * -- beda dari
 * Petugas_JSON yang cuma hitung item final (dipakai leaderboard Dashboard).
 */
function getHomeSummary(username) {
  const info = requireRole_(username, ALL_CYCLE_LIKE_ROLES);
  catatAksesLog_(username, info);
  const today = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd');

  let totalCycleHariIni = 0, selesaiHariIni = 0;
  // v8.26.0: Teruskan username agar baca Ringkasan_Harian dari facility user
  const todayRows = readRingkasanHarianRows_('daily', today, username);

  if (todayRows.length) {
    const r = todayRows[0];
    if (info.role === 'admin' || info.role === 'developer') {
      totalCycleHariIni = Number(r[1]) || 0; // KPI_Total
      selesaiHariIni = (Number(r[2]) || 0) + (Number(r[3]) || 0); // KPI_Hit + KPI_Disc
    } else {
      const petugasRawMap = parseJsonSafe_(r[7], {});
      const mine = petugasRawMap[info.displayName];
      if (mine) {
        totalCycleHariIni = mine.rawTotal || 0;
        selesaiHariIni = mine.rawSelesai || 0;
      }
    }
  }

  let belumValidasi = 0;
  if (info.role === 'admin' || info.role === 'developer') {
    // v8.26.0: Teruskan username agar baca Antrian_Aktif dari facility user
    belumValidasi = (readAntrianValue_('pending_total', { count: 0 }, username)).count || 0;
  } else if (info.role === 'inventory') {
    // konsisten dengan getPendingValidasi/getPendingValidasiCount: cuma hitung yang
    // ditugaskan ke validator ini sendiri, bukan antrian Pending org-wide.
    belumValidasi = (readAntrianValue_('pending_validator:' + info.displayName, { count: 0 }, username)).count || 0;
  }

  return {
    totalCycleHariIni: totalCycleHariIni,
    selesaiHariIni: selesaiHariIni,
    belumValidasi: belumValidasi,
    lastUpdated: Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm'),
    version: APP_VERSION
  };
}

/**
 * Catat akses ke Log_Akses: 1 baris per user per hari (bukan per klik),
 * supaya sheet-nya
 * tidak membengkak tapi tetap kelihatan siapa pakai versi berapa & kapan
 * terakhir buka.
 * Kalau user itu sudah tercatat hari ini, tinggal update Waktu_Terakhir +
 * tambah Jumlah_Buka.
 * Dibungkus try/catch supaya kalau logging gagal (mis. sheet kebetulan lagi
 * dikunci proses lain),
 * TIDAK sampai bikin getHomeSummary ikut gagal -- ini cuma fungsi pelengkap,
 * bukan fungsi inti.
 */
function catatAksesLog_(username, info) {
  try {
    // v8.26.0: Meneruskan username agar tulis ke spreadsheet facility user
    const sheet = getLogAksesSheet_(username);
    const tz = 'Asia/Jakarta';
    const now = new Date();
    const tanggal = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    const jamStr = Utilities.formatDate(now, tz, 'HH:mm:ss');
    const uname = String(username || '').trim().toLowerCase();
    if (!uname) return;

    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      // scan dari bawah -- akses hari ini pasti ada di baris-baris terakhir,
      // jadi lebih cepat
      // ketemu dibanding scan dari atas begitu Log_Akses sudah punya riwayat
      // banyak hari.
      const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues(); // Tanggal, Username
      for (let i = data.length - 1; i >= 0; i--) {
        if (String(data[i][0]) === tanggal && String(data[i][1]) === uname) {
          const rowIndex = i + 2;
          const jumlahCell = sheet.getRange(rowIndex, 8);
          sheet.getRange(rowIndex, 7).setValue(jamStr); // Waktu_Terakhir
          jumlahCell.setValue(Number(jumlahCell.getValue() || 0) + 1); // Jumlah_Buka
          return;
        }
      }
    }

    // belum ada baris untuk user ini hari ini -> tambah baris baru
    sheet.appendRow([tanggal, uname, info.displayName || uname, info.role || '', APP_VERSION, jamStr, jamStr, 1]);
  } catch (e) {
    // sengaja diabaikan -- lihat catatan di atas
  }
}

/**
 * Fase 23: DIUBAH -- dulu scan penuh Riwayat tiap kali dipanggil. SEKARANG
 * baca 2 key kecil
 * di Antrian_Aktif (tracker live yang di-update incremental di
 * submitValidasi/updateTaskStatus/
 * closePlusMinusPair, lihat AntrianAktif.gs). Hasil & bentuk data PERSIS
 * SAMA seperti
 * sebelumnya (termasuk fix lokasi majemuk) -- cuma sumbernya beda, bukan
 * scan lagi.
 */
function getPlusMinusSummary(requesterUsername) {
  requireRole_(requesterUsername, ['inventory', 'admin', 'developer']);
  // v8.26.0: Teruskan requesterUsername agar baca dari facility user
  return getPlusMinusSummaryFromAntrian_(requesterUsername);
}

// ---------- Fungsi TEST ----------
function testSummary() {
  // v8.25.0.1: sebelumnya dipanggil dengan requesterUsername=null (+ 1
  // argumen nyasar) --
  // getSummaryByPeriod cuma 3 parameter, dan requireRole_(null, ['admin'])
  // pasti gagal.
  // Ganti SETUP_ADMIN_NIK_ di bawah kalau mau test pakai NIK lain yang role-nya
  // admin.
  const result = getSummaryByPeriod('daily', '2026-07-04', SETUP_ADMIN_NIK_);
  Logger.log(JSON.stringify(result, null, 2));
}
