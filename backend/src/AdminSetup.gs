/**
 * AdminSetup.gs
 * Menu admin: setup Master (level & mapping PIC), perbaiki header, reset count
 * (Hasil pemecahan dari Code.gs -- semua file .gs berbagi global scope yang sama di Apps Script,
 *  jadi urutan file tidak masalah, fungsi di file lain tetap bisa saling panggil.)
 */

function setupMasterLevelSettings() {
  const ui = SpreadsheetApp.getUi();
  const sheet = getMasterSheet_();
  if (!sheet) {
    ui.alert('Sheet "Master" belum ada. Buat dulu sheet Master (kolom A=ACTIVE LOCATION, C=USER, D=ROLE).');
    return;
  }
  if (!sheet.getRange('F1').getValue()) sheet.getRange('F1').setValue('LEVEL GROUP');
  if (!sheet.getRange('G1').getValue()) sheet.getRange('G1').setValue('PRODUKTIVITAS (qty/orang)');

  LEVEL_GROUPS.forEach(function (g, idx) {
    const row = idx + 2;
    const cellF = sheet.getRange(row, 6);
    const cellG = sheet.getRange(row, 7);
    if (!cellF.getValue()) {
      cellF.setValue('Level ' + g.min + '-' + g.max + ' (' + g.label + ')');
      cellG.setValue(g.defaultProd);
    }
  });

  sheet.getRange('F1:G1').setFontWeight('bold');
  clearMasterCache_();
  ui.alert('Setting level ditambahkan di sheet Master kolom F-G. Sesuaikan angka produktivitas, lalu upload data lagi.');
}

/**
 * Menambahkan area database mapping "Kategori Selisih -> User PIC -> Role PIC" di sheet
 * Master (kolom I, J, K) kalau belum ada -- konsep sama seperti kolom C/D (User/Role):
 * diisi & di-maintain MANUAL oleh admin. Baris kategori (kolom I) di-prefill otomatis dari
 * TASK_KATEGORI_LIST; admin tinggal isi kolom J (username PIC) & K (role PIC, mis. Outbound)
 * untuk tiap kategori. Boleh lebih dari 1 baris per kategori kalau PIC-nya lebih dari satu
 * orang -- semuanya akan muncul sebagai pilihan di form Task Investigasi.
 */
function setupMasterAssignmentMapping() {
  const ui = SpreadsheetApp.getUi();
  const sheet = getMasterSheet_();
  if (!sheet) {
    ui.alert('Sheet "Master" belum ada. Buat dulu sheet Master (kolom A=ACTIVE LOCATION, C=USER, D=ROLE).');
    return;
  }
  if (!sheet.getRange('I1').getValue()) sheet.getRange('I1').setValue('KATEGORI SELISIH');
  if (!sheet.getRange('J1').getValue()) sheet.getRange('J1').setValue('USER PIC');
  if (!sheet.getRange('K1').getValue()) sheet.getRange('K1').setValue('ROLE PIC');

  const existingKategori = {};
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    sheet.getRange(2, 9, lastRow - 1, 1).getValues().forEach(function (r) {
      const v = String(r[0] || '').trim();
      if (v) existingKategori[v] = true;
    });
  }

  let nextRow = Math.max(lastRow + 1, 2);
  TASK_KATEGORI_LIST.forEach(function (k) {
    if (existingKategori[k]) return; // sudah ada barisnya, jangan duplikat
    sheet.getRange(nextRow, 9).setValue(k);
    nextRow++;
  });

  sheet.getRange('I1:K1').setFontWeight('bold');
  clearMasterCache_();
  ui.alert('Kolom I/J/K (Kategori Selisih / User PIC / Role PIC) sudah disiapkan di Master. Isi kolom J & K manual untuk tiap baris kategori -- boleh lebih dari 1 baris per kategori kalau PIC-nya lebih dari 1 orang.');
}

function perbaikiHeader() {
  fixHeaderForSheet_(getSheet_(), HEADERS);
  fixHeaderForSheet_(getRiwayatSheet_(), RIWAYAT_HEADERS);
  SpreadsheetApp.getUi().alert('Header sudah diperiksa/diperbaiki di sheet "Data Count" dan "Riwayat".');
}

function resetCount() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert('Reset semua hasil count di sheet Data Count (Riwayat tidak terhapus)?', ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, HEADERS.length).clearContent();
  ui.alert('Selesai direset.');
}

/**
 * v8.17.2: repair Waktu_Input yang kosong di Data Count (baris Status='Selesai' tapi kolom
 * Waktu_Input kosong -- lihat investigasi Log_Anomali, ternyata jumlah baris yang beneran
 * kosong JAUH lebih banyak dari yang sempat ke-log, karena mekanisme pengecekannya sendiri
 * tidak reliable sebelum fix flush() di v8.17.1). Riwayat.Waktu_Cycle dipakai sebagai sumber
 * cadangan -- nilainya SAMA PERSIS (ditulis dari variabel `waktu` yang sama saat submitCount),
 * tapi ditulis dengan cara APPEND baris baru (bukan update baris lama seperti Data Count),
 * jadi kemungkinan besar tidak kena bug yang sama.
 *
 * Pencocokan pakai kombinasi Lokasi+Article+Nama_Petugas+Qty_Count+Selisih+Hasil_Awal (semua
 * kolom yang ada versinya di kedua sheet), dipasangkan FIFO per kombinasi -- kalau kombinasi
 * yang sama persis muncul lebih dari 1x (jarang tapi mungkin), tetap dipasangkan urut sesuai
 * urutan baris, bukan asal ambil salah satu.
 *
 * dryRun=true (default): CUMA hitung & laporkan lewat Logger, TIDAK menulis apa pun ke sheet.
 * dryRun=false: benar-benar menulis hasil backfill ke kolom Waktu_Input.
 * Dipanggil lewat menu "Cycle Count > Cek & Perbaiki Waktu_Input Kosong" (lihat
 * cekPerbaikiWaktuInputKosong di bawah) -- tidak perlu dijalankan manual dari sini.
 */
function repairWaktuInputDariRiwayat_(dryRun) {
  dryRun = dryRun === undefined ? true : dryRun;

  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { totalBlank: 0, matched: 0, unmatched: 0, unmatchedNos: [] };

  const dcData = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  // index (0-based, sesuai HEADERS): No0 Tanggal1 Lokasi2 Level3 Article4 Description5
  // QtyTransaksi6 QtySystem7 Batch8 NamaPetugas9 Status10 QtyCount11 Selisih12 HasilAwal13
  // WaktuInput14 AddWho15

  const riwayatSheet = getRiwayatSheet_();
  const rLastRow = riwayatSheet.getLastRow();
  const rData = rLastRow >= 2 ? riwayatSheet.getRange(2, 1, rLastRow - 1, RIWAYAT_HEADERS.length).getValues() : [];
  // index (0-based, sesuai RIWAYAT_HEADERS): Lokasi2 Article3 NamaPetugas7 QtyCount8 Selisih9
  // HasilAwal10 WaktuCycle15

  const riwayatQueueByKey = {};
  rData.forEach(function (row) {
    const key = [row[2], row[3], row[7], row[8], row[9], row[10]].join('||');
    if (!riwayatQueueByKey[key]) riwayatQueueByKey[key] = [];
    riwayatQueueByKey[key].push(row[15]);
  });

  let totalBlank = 0, matched = 0;
  const unmatchedNos = [];

  dcData.forEach(function (row) {
    const status = row[10];
    const waktuInput = row[14];
    const valid = waktuInput instanceof Date && !isNaN(waktuInput.getTime());
    if (status !== 'Selesai' || valid) return;
    totalBlank++;

    const key = [row[2], row[4], row[9], row[11], row[12], row[13]].join('||');
    const queue = riwayatQueueByKey[key];
    if (queue && queue.length) {
      row[14] = queue.shift(); // konsumsi 1 match, biar baris berikutnya dgn kombinasi sama gak dobel ambil ini juga
      matched++;
    } else {
      unmatchedNos.push(row[0]);
    }
  });

  Logger.log('Total baris Selesai dengan Waktu_Input kosong: ' + totalBlank);
  Logger.log('Berhasil dicocokkan ke Riwayat: ' + matched);
  Logger.log('TIDAK ketemu pasangannya di Riwayat: ' + unmatchedNos.length);
  if (unmatchedNos.length) {
    Logger.log('No yang tidak ketemu (cek manual): ' + unmatchedNos.slice(0, 100).join(', ') +
      (unmatchedNos.length > 100 ? ' ...(+' + (unmatchedNos.length - 100) + ' lagi)' : ''));
  }

  if (!dryRun && matched) {
    // Tulis balik SELURUH kolom Waktu_Input (O) sekaligus dalam 1x setValues -- jauh lebih
    // ringan daripada 1x setValue per baris yang di-repair, apalagi kalau jumlahnya ribuan.
    const waktuInputColumn = dcData.map(function (row) { return [row[14]]; });
    sheet.getRange(2, 15, waktuInputColumn.length, 1).setValues(waktuInputColumn);
    SpreadsheetApp.flush();
  }

  return { totalBlank: totalBlank, matched: matched, unmatched: unmatchedNos.length, unmatchedNos: unmatchedNos };
}

/**
 * Menu "Cycle Count > Cek & Perbaiki Waktu_Input Kosong" -- cek dulu (dry run), tampilkan
 * ringkasannya, baru minta konfirmasi sebelum benar-benar menulis. Aman dijalankan berkali-
 * kali (baris yang sudah terisi otomatis dilewati, tidak akan ketimpa/dobel proses).
 */
function cekPerbaikiWaktuInputKosong() {
  const ui = SpreadsheetApp.getUi();
  const cek = repairWaktuInputDariRiwayat_(true);

  if (!cek.totalBlank) {
    ui.alert('Tidak ada baris Waktu_Input kosong yang perlu diperbaiki. ');
    return;
  }

  const pesan = 'Ditemukan ' + cek.totalBlank + ' baris Selesai dengan Waktu_Input kosong.\n\n' +
    '- Bisa dicocokkan & diperbaiki dari Riwayat: ' + cek.matched + '\n' +
    '- TIDAK ketemu pasangannya (perlu cek manual, lihat Execution Log): ' + cek.unmatched + '\n\n' +
    'Lanjutkan perbaikan untuk ' + cek.matched + ' baris yang cocok?';

  const resp = ui.alert(pesan, ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;

  const hasil = repairWaktuInputDariRiwayat_(false);
  ui.alert('Selesai. ' + hasil.matched + ' baris Waktu_Input berhasil diperbaiki dari Riwayat.' +
    (hasil.unmatched ? (' ' + hasil.unmatched + ' baris tidak ketemu pasangannya, cek manual (lihat No-nya di menu Extensions > Apps Script > Executions, atau View > Logs).') : ''));
}

