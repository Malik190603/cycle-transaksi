/**
 * ImportData.gs
 * Import data mentah transaksi/stock & pembagian tugas equipment-aware
 * (Hasil pemecahan dari Code.gs -- semua file .gs berbagi global scope yang
 * sama di Apps Script.
 * Jadi urutan file tidak masalah, fungsi di file lain tetap bisa saling
 * panggil.)
 *
 * v8.26.0 (Facility Management - Strict Isolation):
 * - Semua pemanggilan getSheet_(), getLogUnassignedSheet_() SEKARANG
 * meneruskan parameter
 *   * username agar membaca/menulis ke spreadsheet facility user tersebut.
 * - importRawData() memvalidasi bahwa user yang dipilih HANYA dari facility
 * yang sama
 *   * dengan admin yang melakukan upload.
 * - getActiveLocations_() juga menerima username agar membaca lokasi aktif
 * per facility.
 */

// ---------- Import data (raw transaksi + stock by date) ----------

/**
 * Menghitung effort (1/produktivitas) yang SUDAH menjadi beban tiap user
 * dari tugas
 * Pending yang ada saat ini -- dipakai supaya pembagian tugas baru
 * memperhitungkan
 * backlog existing (user yang sudah keteteran tidak otomatis dapat beban
 * baru yang sama
 * besarnya dengan user yang backlog-nya kosong).
 *
 * v8.26.0: Menerima parameter opsional username untuk baca dari facility
 * yang benar.
 */
function getExistingBacklogEffort_(username) {
  const sheet = getSheet_(username);
  const lastRow = sheet.getLastRow();
  const prod = getLevelSettings_();
  const result = {};
  if (lastRow < 2) return result;
  const data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  data.forEach(function (row) {
    const status = row[10];
    if (status !== 'Pending') return;
    const namaPetugas = String(row[9] || '').trim();
    if (!namaPetugas) return;
    const group = getLevelGroupByLocation_(row[2]);
    const productivity = prod[group.key] > 0 ? prod[group.key] : 1;
    result[namaPetugas] = (result[namaPetugas] || 0) + (1 / productivity);
  });
  return result;
}

/**
 * Mengambil map Lokasi+Article -> nomor baris untuk SEMUA tugas berstatus
 * Pending yang
 * sudah ada di sheet -- dipakai untuk dedup lintas-upload (importRawData
 * tidak lagi
 * menciptakan tugas duplikat kalau lokasi+article yang sama masih menunggu
 * dihitung).
 *
 * v8.26.0: Menerima parameter opsional username untuk baca dari facility
 * yang benar.
 */
function getExistingPendingMap_(username) {
  const sheet = getSheet_(username);
  const lastRow = sheet.getLastRow();
  const map = {};
  if (lastRow < 2) return map;
  const data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  data.forEach(function (row, idx) {
    if (row[10] !== 'Pending') return;
    const key = String(row[2]).trim() + '||' + String(row[4]).trim();
    map[key] = { row: idx + 2, qtyTransaksi: Number(row[6]) || 0, addWho: row[15] };
  });
  return map;
}

function buildAssignedRows_(items, startingNo, tanggalUpload, usernames, existingBacklogEffort, compareFn) {
  const prod = getLevelSettings_();
  items.forEach(function (it) {
    const level = extractLevel_(it.lokasi);
    it.level = level;
    const group = getLevelGroupByLocation_(it.lokasi);
    it.levelGroupLabel = group.label;
    const productivity = prod[group.key] > 0 ? prod[group.key] : 1;
    it.effort = 1 / productivity;
    it.locationSort = parseLocation_(it.lokasi);
  });

  items.sort(compareFn || compareLocation_);
  const locGroups = [];
  items.forEach(function (it) {
    const last = locGroups[locGroups.length - 1];
    if (last && last.lokasi === it.lokasi) {
      last.items.push(it);
      last.effort += it.effort;
    } else {
      locGroups.push({ lokasi: it.lokasi, items: [it], effort: it.effort });
    }
  });

  const userCount = usernames.length;
  const totalNewEffort = items.reduce(function (s, it) { return s + it.effort; }, 0);
  const totalBacklogEffort = usernames.reduce(function (s, u) { return s + (existingBacklogEffort[u] || 0); }, 0);
  const targetPerUser = userCount > 0 ? (totalNewEffort + totalBacklogEffort) / userCount : totalNewEffort;
  const output = [];
  let no = startingNo;
  let userIdx = 0;
  let accEffort = existingBacklogEffort[usernames[0]] || 0;
  let chunk = [];

  function flushChunk() {
    if (!chunk.length) return;
    const username = usernames[Math.min(userIdx, userCount - 1)];
    const firstLoc = chunk[0].lokasi;
    const lastLoc = chunk[chunk.length - 1].lokasi;
    const mixLabel = getGroupMixLabel_(chunk);
    const batchLabel = 'Tugas - ' + tanggalUpload + ' - ' + username + ' (' + mixLabel + ') [' + firstLoc + ' s/d ' + lastLoc + ']';
    chunk.forEach(function (it) {
      output.push([
        no, tanggalUpload, it.lokasi, it.level !== null ? it.level : '', it.article, it.description,
        it.qtyTransaksi, it.qtySystem, batchLabel, username, 'Pending', '',
        '', '', '', it.addWhoList || ''
      ]);
      no++;
    });
    chunk = [];
  }

  locGroups.forEach(function (grp, idx) {
    chunk = chunk.concat(grp.items);
    accEffort += grp.effort;
    const isLastGroup = idx === locGroups.length - 1;
    const isLastUser = userIdx >= userCount - 1;
    if (isLastGroup) {
      flushChunk();
    } else if (!isLastUser && accEffort >= targetPerUser) {
      flushChunk();
      userIdx++;
      accEffort = existingBacklogEffort[usernames[userIdx]] || 0;
    }
  });
  return output;
}

function getEquipmentReadyDefaults_() {
  const result = { reachTruck: 0, tangga: 0 };
  const sheet = getMasterSheet_();
  if (!sheet) return result;
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return result;
  const data = sheet.getRange(1, 6, lastRow, 2).getValues();
  data.forEach(function (row) {
    const label = String(row[0] || '').trim().toLowerCase();
    const value = Number(row[1]) || 0;
    if (!label) return;
    if (label.indexOf('reach truck') !== -1) result.reachTruck = value;
    else if (label.indexOf('tangga') !== -1) result.tangga = value;
  });
  return result;
}

function getEquipmentReadyDefaults(requesterUsername) {
  requireRole_(requesterUsername, ['admin', 'developer']);
  return getEquipmentReadyDefaults_();
}

function buildAssignedRowsEquipmentAware_(items, startingNo, tanggalUpload, reachTruckUsers, tanggaUsers, bawahUsers, existingBacklogEffort, username) {
  const warnings = [];
  const reachTruckItems = [];
  const tanggaItems = [];
  const bawahItems = [];
  items.forEach(function (it) {
    const group = getLevelGroupByLocation_(it.lokasi);
    if (group.key === 'reach_truck') reachTruckItems.push(it);
    else if (group.key === 'tangga') tanggaItems.push(it);
    else bawahItems.push(it);
  });

  // v8.26.4: Teruskan username agar baca Setting Overflow dari facility yang benar
  const overflowResult = applyOverflowStoringKeTangga_(reachTruckItems.length, reachTruckUsers, tanggaUsers, username);
  reachTruckUsers = overflowResult.reachTruckUsers;
  tanggaUsers = overflowResult.tanggaUsers;
  const overflowPesan = overflowResult.pesan;
  let no = startingNo;
  let output = [];

  if (reachTruckItems.length) {
    if (!reachTruckUsers.length) {
      warnings.push(reachTruckItems.length + ' item Level 5-6 (Reach Truck) TIDAK dibagi karena tidak ada user role Storing yang dicentang.');
    } else {
      const rows1 = buildAssignedRows_(reachTruckItems, no, tanggalUpload, reachTruckUsers, existingBacklogEffort, compareLocationEvenOddSection_);
      output = output.concat(rows1);
      no += rows1.length;
    }
  }

  if (tanggaItems.length) {
    if (!tanggaUsers.length) {
      warnings.push(tanggaItems.length + ' item Level 3-4 (Tangga Pesawat) TIDAK dibagi karena Tangga Ready = 0.');
    } else {
      const rows2 = buildAssignedRows_(tanggaItems, no, tanggalUpload, tanggaUsers, existingBacklogEffort, compareLocationEvenOddSection_);
      output = output.concat(rows2);
      no += rows2.length;
    }
  }

  if (bawahItems.length) {
    if (!bawahUsers.length) {
      warnings.push(bawahItems.length + ' item Level 1-2 (Bawah) TIDAK dibagi karena tidak ada petugas tersisa di luar slot Reach Truck/Tangga.');
    } else {
      const rows3 = buildAssignedRows_(bawahItems, no, tanggalUpload, bawahUsers, existingBacklogEffort);
      output = output.concat(rows3);
      no += rows3.length;
    }
  }

  return { rows: output, warnings: warnings, overflowPesan: overflowPesan, overflowCount: overflowResult.overflowCount };
}

function getGroupMixLabel_(items) {
  const present = {};
  items.forEach(function (it) { present[it.levelGroupLabel] = true; });
  const labels = Object.keys(present);
  return labels.length > 1 ? 'Campuran' : labels[0];
}

/**
 * v8.26.0: Menerima parameter tambahan `username` agar menulis ke
 * spreadsheet facility
 * yang benar. Semua pemanggilan getSheet_() dan getExistingBacklogEffort_()
 * sekarang
 * meneruskan username ini.
 */
function writeItemsToSheet_(items, reachTruckUsers, tanggaUsers, bawahUsers, selectedUsernames, usernameRoleMap, username) {
  const sheet = getSheet_(username);
  const lastRow = sheet.getLastRow();
  const startingNo = lastRow >= 2 ? Number(sheet.getRange(lastRow, 1).getValue() || 1) + 1 : 1;
  const tanggalUpload = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd');
  const existingBacklogEffort = getExistingBacklogEffort_(username);

  // v8.26.3: mode_assignment SEKARANG per-facility (baris meta di
  // Config_Level_Assignment
  // milik facility `username`, bukan lagi Config_System terpusat.
  const modeAssignment = getLevelAssignmentMeta_(getLevelAssignmentSheet_(username)).mode;
  const result = (modeAssignment === 'matrix')
    ? buildAssignedRowsMatrixAware_(items, startingNo, tanggalUpload, selectedUsernames, usernameRoleMap, existingBacklogEffort, username)
    : buildAssignedRowsEquipmentAware_(items, startingNo, tanggalUpload, reachTruckUsers, tanggaUsers, bawahUsers, existingBacklogEffort, username);

  const rows = result.rows;
  if (rows.length > 0) {
    const startRow = lastRow + 1;
    // Tanggal_Upload (kol. B) & Article (kol. E) harus tetap teks -- diformat SEBELUM ditulis,
    // supaya Sheets tidak mengubah "2026-10-02" jadi sel tanggal atau "00123" jadi angka.
    sheet.getRange(startRow, 2, rows.length, 1).setNumberFormat('@');
    sheet.getRange(startRow, 5, rows.length, 1).setNumberFormat('@');
    sheet.getRange(startRow, 1, rows.length, HEADERS.length).setValues(rows);
  }
  return { total: rows.length, warnings: result.warnings, overflowPesan: result.overflowPesan, overflowCount: result.overflowCount, modeAssignment: modeAssignment };
}

/**
 * v8.26.0: Menerima parameter tambahan `username` agar menulis ke facility
 * yang benar.
 */
function catatLogUnassigned_(adminUsername, warningMessages, username) {
  try {
    const sheet = getLogUnassignedSheet_(username);
    const now = new Date();
    warningMessages.forEach(function (msg) {
      sheet.appendRow([now, adminUsername, msg]);
    });
  } catch (e) {
    // sengaja diabaikan -- logging tidak boleh menggagalkan importRawData()
  }
}

/**
 * requesterUsername wajib admin/developer. Item yang Lokasi+Article-nya SAMA
 * dengan tugas
 * Pending yang sudah ada (dari upload sebelumnya) akan DIGABUNG (qty
 * dijumlahkan ke baris lama)
 * alih-alih membuat baris tugas baru -- mencegah duplikasi task
 * lintas-upload.
 *
 * v8.26.0 (Strict Isolation):
 * - User yang dipilih (selectedUsernames) divalidasi HANYA dari facility
 * yang sama dengan
 *   * requester (admin yang upload).
 * - Semua akses sheet meneruskan requesterUsername agar baca/tulis ke
 * facility yang benar.
 * - Lokasi aktif dibaca dari facility requester, bukan global.
 */
function importRawData(transaksiRawRows, stockRows, selectedUsernames, requesterUsername, reachTruckReady, tanggaReady, jobId) {
  requireRole_(requesterUsername, ['admin', 'developer']);
  reportProgress_(jobId, 5, 'Validasi input');

  if (!transaksiRawRows || !transaksiRawRows.length) {
    return { success: false, message: 'Data Transaksi (raw) kosong.' };
  }
  if (!stockRows || !stockRows.length) {
    return { success: false, message: 'Data Stock by Location kosong.' };
  }
  const usernames = (selectedUsernames || []).map(function (u) { return String(u).trim(); }).filter(function (u) { return u; });
  if (!usernames.length) {
    return { success: false, message: 'Pilih minimal 1 user yang bertugas.' };
  }

  // v8.26.0 STRICT ISOLATION: Validasi bahwa user yang dipilih HANYA dari facility yang sama
  // v8.29.0: sekarang wajib valid -- upload ditolak eksplisit kalau admin/developer belum
  // ter-assign ke facility manapun, bukan diam-diam jalan dengan facilityId null.
  const facInfo = requireUserFacility_(requesterUsername);
  if (!facInfo) {
    return {
      success: false,
      message: 'Akun Anda belum di-assign ke facility manapun (atau facility-nya nonaktif). Hubungi admin lain untuk set assignment facility Anda, atau kalau Anda Developer, pilih facility aktif dulu lewat Facility Switcher.'
    };
  }

  const facilityId = facInfo.id;
  const facilityName = facInfo.name;

  const assignableInFacility = getAssignableUsers_(facilityId);
  const assignableMap = {};
  assignableInFacility.forEach(function (u) { assignableMap[u.username.toLowerCase()] = true; });

  const userTidakValid = [];
  usernames.forEach(function (u) {
    if (!assignableMap[u.toLowerCase()]) {
      userTidakValid.push(u);
    }
  });

  if (userTidakValid.length > 0) {
    return {
      success: false,
      message: 'User berikut TIDAK TERDAFTAR di facility "' + facilityName + '" atau tidak bisa diberi tugas: ' + userTidakValid.join(', ') + '. Pilih hanya user yang sudah di-assign ke facility ini.'
    };
  }

  const roleMap = getUserRoleMap_();
  const reachTruckUsers = usernames.filter(function (u) { return roleMap[u.toLowerCase()] === 'storing'; });
  const nonStoring = usernames.filter(function (u) { return roleMap[u.toLowerCase()] !== 'storing'; });
  const m = Math.max(0, Math.min(Number(tanggaReady) || 0, nonStoring.length));
  const tanggaUsers = nonStoring.slice(0, m);
  const bawahUsers = nonStoring.slice(m);

  const reachTruckReadyNum = Number(reachTruckReady) || 0;
  // v8.26.3: mode_assignment SEKARANG per-facility, baca dari facility requesterUsername.
  const modeAssignmentSaatIni = getLevelAssignmentMeta_(getLevelAssignmentSheet_(requesterUsername)).mode;
  const reachTruckNote = (modeAssignmentSaatIni === 'legacy' && reachTruckReadyNum && reachTruckReadyNum !== reachTruckUsers.length)
    ? (' (Catatan: Input Reach Truck Ready (' + reachTruckReadyNum + ') beda dari jumlah user role Storing yang dicentang (' + reachTruckUsers.length + ') -- yang dipakai tetap semua user Storing yang dicentang, input cuma referensi.)')
    : '';

  // v8.26.0: Baca lokasi aktif DARI FACILITY admin yang upload, bukan global
  reportProgress_(jobId, 15, 'Memuat lokasi aktif facility');
  const activeLocations = getActiveLocations_(requesterUsername);
  if (Object.keys(activeLocations).length === 0) {
    return { success: false, message: 'Tidak ada lokasi aktif untuk facility "' + facilityName + '". Import lokasi aktif terlebih dahulu di Config > Facility Management.' };
  }

  // v8.29.0: lock PER-FACILITY (bukan lock global) -- upload admin DC lain gak skut ke-block.
  const facLock = acquireFacilityLock_(facilityId, 20000);
  if (!facLock) {
    return { success: false, message: 'Sistem sedang sibuk memproses upload lain, coba lagi sebentar.' };
  }

  try {
    reportProgress_(jobId, 25, 'Memproses data stock');
    const stockMap = {};
    stockRows.forEach(function (r) {
      const lokasi = (r[0] === undefined || r[0] === null) ? '' : String(r[0]).trim();
      const article = (r[1] === undefined || r[1] === null) ? '' : String(r[1]).trim();
      if (!lokasi && !article) return;
      const key = lokasi + '||' + article;
      stockMap[key] = (stockMap[key] || 0) + (Number(r[2]) || 0);
    });

    reportProgress_(jobId, 40, 'Memproses data transaksi');
    const itemMap = {};
    let skippedBlankLokasi = 0;
    let skippedInactive = 0;
    let skippedOtherType = 0;

    transaksiRawRows.forEach(function (r) {
      const type = String(r[0] || '').trim().toLowerCase();
      if (type !== 'move' && type !== 'picking') { if (type) skippedOtherType++; return; }

      const article = String(r[1] || '').trim();
      const description = String(r[2] || '').trim();
      const lokasiH = String(r[3] || '').trim();
      const lokasiJ = String(r[4] || '').trim();
      const qtyTransaksi = Number(r[5]) || 0;
      const addWho = String(r[6] || '').trim();

      const lokasi = type === 'move' ? lokasiJ : lokasiH;
      if (!lokasi) { skippedBlankLokasi++; return; }
      if (!activeLocations[lokasi]) { skippedInactive++; return; }

      const key = lokasi + '||' + article;
      if (!itemMap[key]) {
        const qtySystem = stockMap.hasOwnProperty(key) ? Number(stockMap[key]) : 0;
        itemMap[key] = { lokasi: lokasi, article: article, description: description, qtyTransaksi: 0, qtySystem: qtySystem, addWhoByType: { move: {}, picking: {}, lainnya: {} } };
      }

      itemMap[key].qtyTransaksi += qtyTransaksi;
      if (addWho) itemMap[key].addWhoByType[type][addWho] = true;
    });

    let items = Object.keys(itemMap).map(function (k) {
      const it = itemMap[k];
      it.addWhoList = encodeAddWhoByType_(it.addWhoByType);
      return it;
    });

    if (!items.length) {
      return {
        success: false,
        message: 'Tidak ada item yang valid untuk dibagi tugas. ' +
          (skippedBlankLokasi ? skippedBlankLokasi + ' baris lokasi kosong. ' : '') +
          (skippedInactive ? skippedInactive + ' baris lokasi tidak aktif di facility "' + facilityName + '". ' : '') +
          (skippedOtherType ? skippedOtherType + ' baris tipe selain Move/Picking. ' : '')
      };
    }

    // v8.26.0: Meneruskan requesterUsername agar baca dari facility yang benar
    reportProgress_(jobId, 60, 'Menggabungkan tugas existing (dedup)');
    const existingMap = getExistingPendingMap_(requesterUsername);
    const mergedKeys = {};
    let mergedCount = 0;

    items = items.filter(function (it) {
      const key = it.lokasi + '||' + it.article;
      const existing = existingMap[key];
      if (!existing) return true;
      mergedKeys[key] = existing;
      mergedCount++;
      return false;
    });

    if (mergedCount > 0) {
      const sheet = getSheet_(requesterUsername);
      const mergedRows = Object.keys(mergedKeys).map(function (k) { return mergedKeys[k]; });
      mergedRows.sort(function (a, b) { return a.row - b.row; });
      const minRow = mergedRows[0].row;
      const maxRow = mergedRows[mergedRows.length - 1].row;
      const numRows = maxRow - minRow + 1;

      const rangeQty = sheet.getRange(minRow, 7, numRows, 1);
      const rangeAddWho = sheet.getRange(minRow, 16, numRows, 1);
      const qtyValues = rangeQty.getValues();
      const addWhoValues = rangeAddWho.getValues();

      Object.keys(mergedKeys).forEach(function (key) {
        const existing = mergedKeys[key];
        const it = itemMap[key];
        const relRow = existing.row - minRow;
        const qtyLama = Number(qtyValues[relRow][0]) || 0;
        qtyValues[relRow][0] = qtyLama + it.qtyTransaksi;
        const addWhoLama = String(addWhoValues[relRow][0] || '');
        const addWhoBaru = mergeAddWho_(addWhoLama, it.addWhoByType);
        addWhoValues[relRow][0] = addWhoBaru;
      });

      rangeQty.setValues(qtyValues);
      rangeAddWho.setValues(addWhoValues);
    }

    if (!items.length) {
      return {
        success: true,
        message: mergedCount + ' item digabung dengan tugas Pending yang sudah ada (qty ditambahkan). Tidak ada tugas baru dibuat.',
        merged: mergedCount,
        ditugaskan: 0
      };
    }

    const tulis = writeItemsToSheet_(items, reachTruckUsers, tanggaUsers, bawahUsers, usernames, roleMap, requesterUsername);
    reportProgress_(jobId, 95, 'Menyimpan hasil pembagian tugas');

    let pesan = 'Sukses. ' + tulis.total + ' item ditugaskan ke ' + usernames.length + ' user.';
    if (mergedCount) pesan += ' ' + mergedCount + ' item digabung dengan tugas existing.';
    if (skippedInactive) pesan += ' ' + skippedInactive + ' lokasi tidak aktif di facility ini.';
    if (skippedBlankLokasi) pesan += ' ' + skippedBlankLokasi + ' lokasi kosong.';
    if (skippedOtherType) pesan += ' ' + skippedOtherType + ' tipe selain Move/Picking.';
    if (reachTruckNote) pesan += ' ' + reachTruckNote;
    if (tulis.overflowPesan) pesan += ' ' + tulis.overflowPesan;

    let warningGabungan = '';
    if (tulis.warnings && tulis.warnings.length) {
      warningGabungan = tulis.warnings.join(' ');
      pesan += ' ⚠️ ' + warningGabungan;
      catatLogUnassigned_(requesterUsername, tulis.warnings, requesterUsername);
    }

    return {
      success: true,
      message: pesan,
      ditugaskan: tulis.total,
      merged: mergedCount,
      skippedInactive: skippedInactive,
      skippedBlankLokasi: skippedBlankLokasi,
      skippedOtherType: skippedOtherType,
      warning: warningGabungan,
      perluKonfirmasi: !!(tulis.warnings && tulis.warnings.length),
      modeAssignment: tulis.modeAssignment,
      facilityName: facilityName
    };

  } finally {
    facLock.release();
  }
}
