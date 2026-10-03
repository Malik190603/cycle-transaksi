/*
 * Config (pemegang akses Config dan Developer): user, pembagian tugas, akses, facility.
 * Dibuka dari lembar Akun sebagai halaman tersendiri.
 */
// Server menganggap "dev" dan "dewa" (ejaan lama) sama dengan Developer; daftar user mengirim ejaan aslinya.
const peranDeveloper = (r) => /^(developer|dev|dewa)$/.test(String(r || '').trim().toLowerCase());
const Config = {
  tab: 'user', users: null, roleValid: [], cariUser: '', tugas: null, log: null, logSemua: false, akses: null, facility: null,
  jalan: {}, galat: {}, req: 0, // per tab: nomor permintaan yang sedang berjalan dan galat terakhirnya

  reset() { Object.assign(this, { tab: 'user', users: null, cariUser: '', tugas: null, log: null, logSemua: false, akses: null, facility: null, jalan: {}, galat: {} }); },
  saya() { return Sesi.user.username; },

  gambar() {
    const node = el('scr-config');
    const TABS = [['user', 'User'], ['tugas', 'Pembagian tugas'], ['akses', 'Akses'], ['facility', 'Facility']];
    isi(node, h`${(CFG.panel ? kepala : kepalaHalaman)('Config', Sesi.user.facilityName || '', h`<button type="button" class="iconbtn${this.jalan[this.tab] ? ' is-putar' : ''}" data-aksi="config-segar" aria-label="Perbarui">${ikon('segar')}</button>`)}
      <div class="chips" style="margin-bottom:14px" role="tablist">${TABS.map((t) => h`<button type="button" role="tab" class="chip${this.tab === t[0] ? ' is-on' : ''}" data-aksi="config-tab" data-v="${t[0]}">${t[1]}</button>`)}</div>
      <div id="cfIsi">${this.galat[this.tab] ? gagalMuat(this.galat[this.tab], 'config-segar') : this[this.tab + 'Html']()}</div>`);
    const c = el('cfCari');
    if (c) c.addEventListener('input', tunda(() => { this.cariUser = c.value; isi(el('cfUsers'), this.daftarUser()); }, 160));
    // Matriks peran: urutan centang = urutan prioritas, jadi keadaan disimpan saat dicentang, bukan dibaca dari tabel saat menyimpan.
    $$('input[data-m-role]', node).forEach((cb) => cb.addEventListener('change', () => {
      const g = this.grup(cb.dataset.mGrup), i = g.roles.indexOf(cb.dataset.mRole);
      if (cb.checked && i < 0) g.roles.push(cb.dataset.mRole); else if (!cb.checked && i >= 0) g.roles.splice(i, 1);
      isi(el('cfCara'), this.caraHtml());
    }));
    const maks = el('cfMaks');
    if (maks) maks.addEventListener('input', () => { this.tugas.maksGrupAlat = maks.value; });
  },
  gambarTetap() { const y = window.scrollY; this.gambar(); window.scrollTo(0, y); },
  ada(t) { return !!(t === 'user' ? this.users : t === 'tugas' ? this.tugas : t === 'akses' ? this.akses : this.facility); },
  // Tiap tab dimuat sendiri-sendiri. Pindah tab selagi tab lain masih dimuat tidak boleh membuat tab baru
  // tertahan di kerangka, dan hasil permintaan lama (atau milik pengguna sebelumnya) tidak boleh menimpa yang baru.
  // `tab` wajib diisi oleh pemanggil yang baru selesai mengubah data: perubahan bisa selesai setelah pengguna
  // pindah tab, dan yang perlu dimuat ulang adalah tab asal perubahan itu, bukan tab yang sedang dibuka.
  async muat(paksa, tab) {
    const t = tab || this.tab;
    if (!Sesi.user || (!paksa && (this.ada(t) || this.jalan[t]))) return;
    const no = this.jalan[t] = ++this.req, siapa = this.saya();
    delete this.galat[t];
    if (Nav.aktif() === 'config' && this.tab === t) this.gambar();
    const cek = (r) => { if (r && r.success === false) throw new Error(r.message); return r; };
    let hasil = null, salah = null;
    try {
      if (t === 'user') { const r = cek(await panggil('getDaftarUserMaster', siapa)); hasil = { users: r.users, roleValid: r.roleValid || [] }; }
      else if (t === 'tugas') { const r = await Promise.all([panggil('getLevelAssignmentConfig', siapa), panggil('getLogPerubahanConfig', siapa)]); hasil = { tugas: cek(r[0]), log: (r[1] && r[1].log) || [] }; }
      else if (t === 'akses') hasil = { akses: cek(await panggil('getDaftarAksesSetting', siapa)).daftar };
      else hasil = { facility: cek(await panggil('getDaftarFacility', siapa)).facilities };
    } catch (e) { salah = e; }
    if (this.jalan[t] !== no) return;
    delete this.jalan[t];
    if (salah) this.galat[t] = salah; else Object.assign(this, hasil);
    // Tab lain yang sedang dibuka tidak digambar ulang: isian yang sedang diketik di sana tidak boleh hilang.
    if (Nav.aktif() === 'config' && this.tab === t) this.gambar();
  },
  // Jalankan satu perubahan: tombol sibuk, pesan hasil lewat toast, lalu muat ulang tab.
  async ubah(tombol, janji, sesudah) {
    if (tombol) tombol.classList.add('is-sibuk');
    try {
      const res = await janji();
      if (res && res.success === false) { Toast.galat(new Error(res.message)); return false; }
      if (res && res.message) Toast.tampil(res.message, { lama: res.warning ? 6000 : 3200 });
      if (sesudah) sesudah(res);
      return true;
    } catch (e) { Toast.galat(e); return false; }
    finally { if (tombol) tombol.classList.remove('is-sibuk'); }
  },
  peranBoleh() { return this.roleValid.filter((r) => r !== 'developer' || Boleh.developer()); },

  /* ================= User ================= */
  userHtml() {
    if (!this.users) return kerangka(4, 62);
    return h`<section class="card"><h2 class="card__judul">Tambah user</h2>
        <label class="lbl" for="cfNik">NIK</label><input class="field" id="cfNik" type="text" autocapitalize="off" autocorrect="off" spellcheck="false" placeholder="contoh: 123456.nama" autocomplete="off">
        <label class="lbl" for="cfPeran">Peran</label><select class="field" id="cfPeran"><option value="">Pilih peran</option>${this.peranBoleh().map((r) => h`<option value="${r}">${namaPeran(r)}</option>`)}</select>
        <p class="ket">User baru otomatis masuk ke facility Anda.</p>
        <button type="button" class="btn" data-aksi="config-user-tambah" style="margin-top:14px">Tambah user</button></section>
      <h2 class="bag"><span>Daftar user</span><small>${angka(this.users.length)} orang</small></h2>
      ${this.users.length > 8 ? h`<div class="cari" style="margin-bottom:10px">${ikon('cari', 'ic--kecil')}<input class="field" type="search" id="cfCari" placeholder="Cari NIK atau peran" value="${this.cariUser}" autocomplete="off"></div>` : ''}
      <div id="cfUsers">${this.daftarUser()}</div>`;
  },
  daftarUser() {
    const q = this.cariUser.trim().toLowerCase(), daftar = this.users.filter((u) => !q || cocok(u.username + ' ' + u.role + ' ' + u.status, q));
    if (!daftar.length) return h`<p class="redup" style="text-align:center;padding:18px 0">Tidak ada user yang cocok.</p>`;
    return h`<div class="list">${daftar.map((u) => h`<button type="button" class="row" data-aksi="config-user-buka" data-row="${u.rowIndex}"><span class="row__isi"><span class="row__t">${u.username}</span><span class="row__s">${u.facilityNama || 'Belum punya facility'}</span></span>
      <span class="row__ekor"><span class="pill${u.status === 'Nonaktif' ? ' pill--merah' : ''}">${u.status === 'Nonaktif' ? 'Nonaktif' : namaPeran(peranDeveloper(u.role) ? 'developer' : u.role)}</span>${ikon('kanan', 'ic--kecil')}</span></button>`)}</div>`;
  },
  bukaUser(row) {
    const u = this.users.find((x) => String(x.rowIndex) === String(row)); if (!u) return;
    const nonaktif = u.status === 'Nonaktif', peranSekarang = peranDeveloper(u.role) ? 'developer' : String(u.role || '').trim().toLowerCase();
    const sendiri = String(u.username).toLowerCase() === this.saya().toLowerCase();
    // Aturan yang sama dijaga server: akun Developer hanya diubah Developer, dan akun sendiri tidak bisa dinonaktifkan.
    if (peranSekarang === 'developer' && !Boleh.developer()) {
      Lembar.buka({ judul: u.username, isi: h`<div class="rinci"><div class="rinci__baris"><span>Peran</span><b>${namaPeran(peranSekarang)}</b></div><div class="rinci__baris"><span>Status</span><b>${nonaktif ? 'Nonaktif' : 'Aktif'}</b></div></div>
        <p class="ket">Akun Developer hanya bisa diubah oleh Developer.</p>` });
      return;
    }
    const opsi = this.peranBoleh().indexOf(peranSekarang) >= 0 ? this.peranBoleh() : this.peranBoleh().concat([peranSekarang]);
    const l = Lembar.buka({
      judul: u.username,
      isi: h`<label class="lbl" for="cfPeranUbah">Peran</label><select class="field" id="cfPeranUbah">${opsi.map((r) => h`<option value="${r}" ${r === peranSekarang ? 'selected' : ''}>${namaPeran(r)}</option>`)}</select>
        <button type="button" class="btn" id="cfPeranSimpan" style="margin-top:14px">Simpan peran</button>
        ${sendiri ? h`<p class="ket">Ini akun Anda. Akun sendiri tidak bisa dinonaktifkan.</p>` : h`<button type="button" class="btn btn--garis" id="cfStatusUbah" style="margin-top:10px;${nonaktif ? '' : 'color:var(--merah-teks)'}">${nonaktif ? 'Aktifkan kembali' : 'Nonaktifkan user'}</button>
        <p class="ket">${nonaktif ? 'User ini sedang tidak bisa masuk.' : 'User nonaktif tidak bisa masuk dan tidak menerima tugas baru. Riwayatnya tetap tersimpan.'}</p>`}`
    });
    el('cfPeranSimpan').addEventListener('click', (ev) => this.ubah(ev.currentTarget, () => panggil('updateRoleUserMaster', this.saya(), u.rowIndex, u.username, el('cfPeranUbah').value), () => { l.tutup(); this.users = null; this.muat(true, 'user'); }));
    if (!sendiri) el('cfStatusUbah').addEventListener('click', async (ev) => {
      const b = ev.currentTarget;
      if (!nonaktif && !(await tanya({ judul: 'Nonaktifkan ' + u.username + '?', pesan: 'User ini tidak akan bisa masuk lagi sampai diaktifkan kembali.', ya: 'Nonaktifkan', bahaya: true }))) return;
      this.ubah(b, () => panggil('setStatusUserMaster', this.saya(), u.rowIndex, u.username, nonaktif ? 'Aktif' : 'Nonaktif'), () => { l.tutup(); this.users = null; this.muat(true, 'user'); });
    });
  },

  /* ================= Pembagian tugas ================= */
  PENDEK: { bawah: ['Bawah', 'Level 1-2'], tangga: ['Tangga', 'Level 3-4'], reach_truck: ['Reach truck', 'Level 5-6'] },
  grup(key) {
    const t = this.tugas; let g = t.matrix.find((m) => m.grupAlat === key);
    if (!g) { g = { grupAlat: key, roles: [], mode: 'pemerataan' }; t.matrix.push(g); }
    return g;
  },
  tugasHtml() {
    const t = this.tugas; if (!t) return kerangka(3, 110);
    const matrix = t.modeAssignment === 'matrix', P = this.PENDEK, log = this.log || [], BATAS = 5;
    return h`<section class="card"><h2 class="card__judul">${matrix ? 'Mode Matrix aktif' : 'Mode lama aktif'}</h2>
        <p class="card__ket">${matrix ? 'Tugas harian dibagi mengikuti tabel peran di bawah.' : 'Storing mengerjakan Level 5-6, peran lain mengerjakan Level 1-4. Tabel berikut belum berlaku sampai Mode Matrix diaktifkan.'}</p>
        <button type="button" class="btn ${matrix ? 'btn--garis' : ''}" data-aksi="config-mode" data-v="${matrix ? 'legacy' : 'matrix'}" style="margin-top:14px">${matrix ? 'Kembali ke mode lama' : 'Aktifkan Mode Matrix'}</button></section>
      <section class="card hm-kartu"><h2 class="card__judul">Peran per grup alat</h2><p class="card__ket">Centang peran yang boleh mengerjakan tiap grup alat.</p>
        <div class="cf-matriks"><table><thead><tr><th>Peran</th>${t.grupAlatList.map((g) => h`<th>${(P[g.key] || [g.label])[0]}<small>${(P[g.key] || ['', ''])[1]}</small></th>`)}</tr></thead>
          <tbody>${t.roleValid.map((r) => h`<tr><th scope="row">${namaPeran(r)}</th>${t.grupAlatList.map((g) => h`<td><input type="checkbox" class="cek" data-m-role="${r}" data-m-grup="${g.key}" ${this.grup(g.key).roles.indexOf(r) >= 0 ? 'checked' : ''} aria-label="${namaPeran(r)} di ${g.label}"></td>`)}</tr>`)}</tbody></table></div>
        <h3 class="lbl" style="margin-top:18px">Cara bagi</h3>
        <div class="cf-cara" id="cfCara">${this.caraHtml()}</div>
        <p class="ket">Merata: semua peran yang dicentang mendapat bagian yang sama. Prioritas: peran dipakai berurutan, peran berikutnya baru ikut bila yang sebelumnya tidak cukup.</p>
        <label class="lbl" for="cfMaks">Maksimal grup alat per peran</label><input class="field" id="cfMaks" type="number" inputmode="numeric" min="1" max="3" value="${t.maksGrupAlat}" style="max-width:120px">
        <p class="ket">Hanya peringatan saat dilampaui, tidak menghalangi penyimpanan.</p>
        <button type="button" class="btn" data-aksi="config-matriks-simpan" style="margin-top:14px">Simpan pembagian tugas</button></section>
      <h2 class="bag"><span>Riwayat perubahan</span>${log.length ? h`<small>${angka(log.length)} terakhir</small>` : ''}</h2>
      ${log.length ? h`<div class="list">${log.slice(0, this.logSemua ? 50 : BATAS).map((x) => h`<div class="row"><div class="row__isi"><div class="row__t row__t--lipat" style="font-weight:500">${x.detail}</div><div class="row__s">${x.nik}, ${this.waktu(x.waktu)}</div></div></div>`)}</div>
        ${log.length > BATAS ? h`<button type="button" class="btn btn--teks" data-aksi="config-log" style="margin:6px auto 0">${this.logSemua ? 'Tampilkan lebih sedikit' : 'Tampilkan semua'}</button>` : ''}` : h`<p class="redup">Belum ada perubahan.</p>`}`;
  },
  caraHtml() {
    const t = this.tugas, P = this.PENDEK;
    return t.grupAlatList.map((g) => {
      const gd = this.grup(g.key), prio = gd.mode === 'prioritas';
      return h`<div class="cf-cara__baris"><div class="cf-cara__atas"><span>${(P[g.key] || [g.label])[0]} <small>${(P[g.key] || ['', ''])[1]}</small></span>
          <div class="seg seg--kecil" role="group" aria-label="Cara bagi ${g.label}"><button type="button" class="${prio ? '' : 'is-on'}" aria-pressed="${prio ? 'false' : 'true'}" data-aksi="config-cara" data-grup="${g.key}" data-v="pemerataan">Merata</button><button type="button" class="${prio ? 'is-on' : ''}" aria-pressed="${prio ? 'true' : 'false'}" data-aksi="config-cara" data-grup="${g.key}" data-v="prioritas">Prioritas</button></div></div>
        ${prio ? h`<p class="ket">${gd.roles.length ? h`Urutan: <b>${gd.roles.map(namaPeran).join(', lalu ')}</b>. Untuk memindahkan satu peran ke urutan terakhir, hapus centangnya lalu centang lagi.` : 'Belum ada peran yang dicentang.'}</p>` : ''}</div>`;
    });
  },
  waktu(iso) { const d = new Date(iso); return isNaN(d) ? String(iso || '') : d.getDate() + ' ' + BULAN3[d.getMonth()] + ' ' + d.getFullYear() + ', ' + jam(d.getTime()); },
  simpanMatriks(tombol) {
    const t = this.tugas;
    const matrix = t.grupAlatList.map((g) => { const gd = this.grup(g.key); return { grupAlat: g.key, roles: gd.roles.slice(), mode: gd.mode === 'prioritas' ? 'prioritas' : 'pemerataan' }; });
    this.ubah(tombol, () => panggil('saveLevelAssignmentMatrix', this.saya(), matrix, Number(el('cfMaks').value) || 2), (res) => { if (res.warning) Toast.tampil(res.warning, { lama: 7000 }); this.tugas = null; this.muat(true, 'tugas'); });
  },
  async gantiMode(tombol, mode) {
    const ya = await tanya(mode === 'matrix'
      ? { judul: 'Aktifkan Mode Matrix?', pesan: 'Mulai upload berikutnya, tugas dibagi mengikuti tabel peran per grup alat.', ya: 'Aktifkan' }
      : { judul: 'Kembali ke mode lama?', pesan: 'Tugas kembali dibagi dengan aturan Storing dan non-Storing. Isi tabel tidak hilang dan bisa diaktifkan lagi.', ya: 'Kembali ke mode lama' });
    if (ya) this.ubah(tombol, () => panggil('setModeAssignment', this.saya(), mode), () => { this.tugas = null; Upload.form = null; this.muat(true, 'tugas'); });
  },

  /* ================= Akses ================= */
  aksesHtml() {
    if (!this.akses) return kerangka(3, 62);
    return h`<section class="card"><h2 class="card__judul">Beri akses Config</h2><p class="card__ket">NIK di daftar ini bisa membuka halaman Config.</p>
        <label class="lbl" for="cfAksesNik">NIK</label><input class="field" id="cfAksesNik" type="text" autocapitalize="off" autocorrect="off" spellcheck="false" autocomplete="off" placeholder="contoh: 123456.nama">
        <label class="lbl" for="cfAksesNama">Nama <small>(opsional)</small></label><input class="field" id="cfAksesNama" type="text" autocomplete="off">
        <button type="button" class="btn" data-aksi="config-akses-tambah" style="margin-top:14px">Beri akses</button></section>
      <h2 class="bag"><span>Pemegang akses</span><small>${angka(this.akses.length)} orang</small></h2>
      ${this.akses.length ? h`<div class="list">${this.akses.map((a) => { const sendiri = String(a.nik).toLowerCase() === this.saya().toLowerCase(); return h`<div class="row"><div class="row__isi"><div class="row__t">${a.nik}</div><div class="row__s">${a.nama || 'Tanpa nama'}${a.tanggal ? h`, sejak ${tanggalBebas(a.tanggal)}` : ''}</div></div>
        <div class="row__ekor">${sendiri ? h`<span class="pill">Anda</span>` : h`<button type="button" class="iconbtn iconbtn--polos" data-aksi="config-akses-cabut" data-row="${a.rowIndex}" data-nik="${a.nik}" aria-label="Cabut akses ${a.nik}" style="color:var(--merah-teks)">${ikon('sampah')}</button>`}</div></div>`; })}</div>` : h`<p class="redup">Belum ada NIK terdaftar.</p>`}`;
  },

  /* ================= Facility ================= */
  facilityHtml() {
    if (!this.facility) return kerangka(2, 150);
    return h`${Boleh.developer() ? h`<button type="button" class="btn btn--tenang" data-aksi="config-fac-baru" style="margin-bottom:12px">${ikon('tambah')}Tambah facility baru</button>` : ''}
      ${this.facility.length ? h`<div class="tumpuk">${this.facility.map((f) => { const aktif = f.status === 'Aktif'; const url = f.spreadsheetId ? 'https://docs.google.com/spreadsheets/d/' + encodeURIComponent(f.spreadsheetId) + '/edit' : f.url; return h`<section class="card">
        <div class="cf-fac__atas"><div><h2 class="card__judul">${f.nama}</h2><div class="card__ket">${f.kode}${f.tanggalDibuat ? h`, dibuat ${tanggalBebas(f.tanggalDibuat)}` : ''}</div></div><span class="pill ${aktif ? 'pill--hijau' : 'pill--merah'}">${f.status}</span></div>
        <div class="cf-fac__aksi">
          <button type="button" class="btn btn--kecil btn--garis" data-aksi="config-fac-lokasi" data-id="${f.id}">${ikon('pin', 'ic--kecil')}Lokasi aktif</button>
          <button type="button" class="btn btn--kecil btn--garis" data-aksi="config-fac-user" data-id="${f.id}">${ikon('orang', 'ic--kecil')}User</button>
          <button type="button" class="btn btn--kecil btn--garis" data-aksi="config-fac-ubah" data-id="${f.id}">${ikon('pensil', 'ic--kecil')}Ubah</button>
          <button type="button" class="btn btn--kecil btn--garis" data-aksi="config-fac-status" data-id="${f.id}">${aktif ? 'Nonaktifkan' : 'Aktifkan'}</button>
        </div>
        ${url ? h`<a class="tautan cf-fac__db" href="${url}" target="_blank" rel="noopener">Buka database di Google Sheets</a>` : ''}</section>`; })}</div>` : h`<section class="card">${kosong('gedung', 'Belum ada facility', 'Facility dibuat oleh Developer.')}</section>`}`;
  },
  fac(id) { return (this.facility || []).find((f) => f.id === id); },
  ubahFacility(id) {
    const f = this.fac(id); if (!f) return;
    const l = Lembar.buka({
      judul: 'Ubah facility',
      isi: h`<label class="lbl" for="cfFacNama">Nama facility</label><input class="field" id="cfFacNama" type="text" value="${f.nama}">
        <label class="lbl" for="cfFacKode">Kode</label><input class="field" id="cfFacKode" type="text" value="${f.kode}">
        <button type="button" class="btn" id="cfFacSimpan" style="margin-top:16px">Simpan</button>`
    });
    el('cfFacSimpan').addEventListener('click', (ev) => {
      const nama = el('cfFacNama').value.trim();
      if (!nama) { Toast.galat(new Error('Nama facility belum diisi.')); return; }
      this.ubah(ev.currentTarget, () => panggil('updateNamaFacility', this.saya(), id, nama, el('cfFacKode').value.trim()), () => { l.tutup(); this.facility = null; this.muat(true, 'facility'); });
    });
  },
  async statusFacility(tombol, id) {
    const f = this.fac(id); if (!f) return;
    const aktif = f.status === 'Aktif';
    if (aktif && !(await tanya({ judul: 'Nonaktifkan ' + f.nama + '?', pesan: 'User di facility ini tidak bisa mengirim hasil sampai facility diaktifkan kembali.', ya: 'Nonaktifkan', bahaya: true }))) return;
    this.ubah(tombol, () => panggil('setStatusFacility', this.saya(), id, aktif ? 'Nonaktif' : 'Aktif'), () => { this.facility = null; this.muat(true, 'facility'); });
  },
  async userFacility(id) {
    const f = this.fac(id); if (!f) return;
    const l = Lembar.buka({ judul: 'User di ' + f.nama, isi: memuat() });
    const gambarIsi = async () => {
      try {
        const r = await Promise.all([panggil('getDaftarUserMaster', this.saya()), panggil('getDaftarUserFacilityAssignment', this.saya())]);
        if (r[0].success === false) throw new Error(r[0].message);
        const semua = r[0].users.filter((u) => u.status !== 'Nonaktif'), anggota = semua.filter((u) => u.facilityId === id);
        const tgl = {}; ((r[1] && r[1].assignments) || []).forEach((a) => { if (a.facilityId === id) tgl[a.username] = a.tanggalDiassign; });
        isi(l.isi, h`<label class="lbl" for="cfAssign">Masukkan user ke facility ini</label>
          <select class="field" id="cfAssign"><option value="">Pilih user</option>${semua.filter((u) => u.facilityId !== id).map((u) => h`<option value="${u.username}">${u.username} (${namaPeran(u.role)}${u.facilityNama ? ', kini di ' + u.facilityNama : ''})</option>`)}</select>
          <button type="button" class="btn" id="cfAssignBtn" style="margin-top:12px">Masukkan</button>
          <h3 class="bag"><span>Sudah di facility ini</span><small>${angka(anggota.length)} orang</small></h3>
          ${anggota.length ? h`<div class="list">${anggota.map((u) => h`<div class="row"><div class="row__isi"><div class="row__t">${u.username}</div>${tgl[u.username] ? h`<div class="row__s">Sejak ${tanggalBebas(tgl[u.username])}</div>` : ''}</div><span class="pill">${namaPeran(u.role)}</span></div>`)}</div>` : h`<p class="redup">Belum ada user.</p>`}`);
        el('cfAssignBtn').addEventListener('click', (ev) => {
          const u = el('cfAssign').value;
          if (!u) { Toast.galat(new Error('Pilih user dulu.')); return; }
          this.ubah(ev.currentTarget, () => panggil('assignUserKeFacility', this.saya(), u, id), () => { this.users = null; gambarIsi(); });
        });
      } catch (e) { isi(l.isi, h`<p class="pesan pesan--galat">${e.message}</p>`); }
    };
    gambarIsi();
  },
  // Sumber daftar lokasi: file (kolom pertama), tempel teks, atau salin dari facility lain.
  sumberLokasiHtml(awalan, facLain) {
    return h`<div class="seg" id="${awalan}Sumber"><button type="button" class="is-on" data-v="file">File</button><button type="button" data-v="tempel">Tempel teks</button>${facLain.length ? h`<button type="button" data-v="salin">Salin</button>` : ''}</div>
      <div id="${awalan}File" style="margin-top:10px"><label class="up-berkas" id="${awalan}Ubin"><input type="file" class="sr" id="${awalan}Berkas" accept=".xlsx,.xls,.csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,text/csv">
        <span class="up-berkas__ic">${ikon('berkas')}</span><span class="up-berkas__isi"><span class="up-berkas__judul">File daftar lokasi</span><span class="up-berkas__ket">.xlsx atau .csv, kode lokasi di kolom pertama. Baris judul boleh ada.</span></span><span class="up-berkas__aksi">Pilih file</span></label></div>
      <div id="${awalan}Tempel" style="margin-top:10px" hidden><textarea class="field field--tempel" id="${awalan}Teks" rows="5" placeholder="A01.001.1&#10;A01.001.2&#10;A01.002.1" spellcheck="false" autocapitalize="characters"></textarea><p class="ket">Satu lokasi per baris.</p></div>
      <div id="${awalan}Salin" style="margin-top:10px" hidden><select class="field" id="${awalan}Dari"><option value="">Pilih facility sumber</option>${facLain.map((f) => h`<option value="${f.id}">${f.nama} (${f.kode})</option>`)}</select><p class="ket">Daftar lokasi aktif facility itu disalin ke sini.</p></div>`;
  },
  pasangSumber(awalan) {
    const tombol = $$('#' + awalan + 'Sumber button');
    tombol.forEach((b) => b.addEventListener('click', () => {
      tombol.forEach((x) => x.classList.toggle('is-on', x === b));
      ['File', 'Tempel', 'Salin'].forEach((n) => { const d = el(awalan + n); if (d) d.hidden = n.toLowerCase() !== b.dataset.v; });
    }));
    // File dibaca begitu dipilih, supaya jumlah lokasi (atau kesalahannya) terlihat sebelum tombol impor ditekan.
    const inp = el(awalan + 'Berkas'), ubin = el(awalan + 'Ubin');
    this._lokasiBerkas = this._lokasiBerkas || {}; delete this._lokasiBerkas[awalan];
    inp.addEventListener('change', async () => {
      const f = inp.files && inp.files[0]; if (!f) return;
      const ket = $('.up-berkas__ket', ubin), ic = $('.up-berkas__ic', ubin), tulis = (kelas, teks, ik) => { ubin.className = 'up-berkas' + (kelas ? ' ' + kelas : ''); ket.textContent = teks; isi(ic, ik); $('.up-berkas__aksi', ubin).textContent = 'Ganti'; };
      delete this._lokasiBerkas[awalan];
      tulis('', 'Membaca ' + f.name + '…', h`<span class="putar"></span>`);
      try {
        const daftar = await this.bacaBerkasLokasi(f);
        if (!daftar.length) throw new Error('Tidak ada kode lokasi di kolom pertama file ini.');
        this._lokasiBerkas[awalan] = daftar;
        tulis('is-ok', f.name + ': ' + angka(daftar.length) + ' lokasi terbaca.', ikon('cek'));
      } catch (e) { tulis('is-galat', e.message, ikon('berkas')); }
    });
    return () => { const on = tombol.find((b) => b.classList.contains('is-on')); return on ? on.dataset.v : 'file'; };
  },
  async bacaBerkasLokasi(file) {
    const X = await muatXlsx();
    const buf = await new Promise((ok, gagal) => { const r = new FileReader(); r.onload = () => ok(r.result); r.onerror = () => gagal(new Error('File tidak bisa dibaca.')); r.readAsArrayBuffer(file); });
    const wb = X.read(new Uint8Array(buf), { type: 'array', raw: true });
    const out = [], lihat = {};
    X.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 }).forEach((row, i) => {
      const v = teksSel(row && row[0]);
      if (!v || (i === 0 && /^(lokasi|location|loc)\b/i.test(v)) || lihat[v.toUpperCase()]) return;
      lihat[v.toUpperCase()] = true; out.push(v);
    });
    return out;
  },
  async bacaLokasi(awalan, sumber) {
    if (sumber === 'tempel') return el(awalan + 'Teks').value.split(/\r?\n/).map((x) => x.trim()).filter(Boolean);
    const siap = this._lokasiBerkas && this._lokasiBerkas[awalan];
    if (siap) return siap;
    const inp = el(awalan + 'Berkas');
    if (!inp.files || !inp.files[0]) throw new Error('Pilih file daftar lokasi dulu.');
    return this.bacaBerkasLokasi(inp.files[0]);
  },
  async bukaLokasi(id, nama, sesudah) {
    const l = Lembar.buka({ judul: 'Lokasi aktif', penuh: true, isi: memuat() });
    let jumlah = null;
    try { const r = await panggil('getDaftarLokasiAktif', this.saya(), id); if (r.success === false) throw new Error(r.message); jumlah = r.total || 0; }
    catch (e) { isi(l.isi, h`<p class="pesan pesan--galat">${e.message}</p>`); return; }
    if (!this.facility) { try { const r = await panggil('getDaftarFacility', this.saya()); if (r.success) this.facility = r.facilities; } catch (e) { /* daftar salin tidak tersedia */ } }
    const lain = (this.facility || []).filter((f) => f.id !== id && f.status === 'Aktif');
    isi(l.isi, h`<div class="up-hasil"><b class="angka">${angka(jumlah)}</b><span>lokasi aktif di ${nama}</span></div>
      <p class="redup" style="margin-bottom:14px">Hanya transaksi di lokasi aktif yang dijadikan tugas. Lokasi virtual dan lantai tidak perlu dimasukkan.</p>
      <span class="lbl">Sumber daftar lokasi</span>${this.sumberLokasiHtml('cfLok', lain)}
      <span class="lbl">Cara memasukkan</span>
      <div class="seg" id="cfLokCara"><button type="button" class="is-on" data-v="tambah">Tambahkan</button><button type="button" data-v="ganti">Ganti semua</button></div>
      <p class="ket">Tambahkan: lokasi yang sudah ada dibiarkan. Ganti semua: daftar lama dihapus dulu.</p>
      <div class="pesan pesan--galat" id="cfLokPesan" role="alert"></div>
      <button type="button" class="btn" id="cfLokSimpan" style="margin-top:14px">Impor lokasi</button>`);
    const sumber = this.pasangSumber('cfLok');
    const cara = $$('#cfLokCara button');
    cara.forEach((b) => b.addEventListener('click', () => cara.forEach((x) => x.classList.toggle('is-on', x === b))));
    el('cfLokSimpan').addEventListener('click', async (ev) => {
      const b = ev.currentTarget, caraOn = cara.find((x) => x.classList.contains('is-on')), ganti = !!caraOn && caraOn.dataset.v === 'ganti', pesan = (m) => { el('cfLokPesan').textContent = m; };
      pesan('');
      try {
        let janji;
        if (sumber() === 'salin') {
          const dari = el('cfLokDari').value; if (!dari) return pesan('Pilih facility sumber.');
          janji = () => panggil('copyLokasiDariFacility', this.saya(), dari, id, ganti);
        } else {
          const daftar = await this.bacaLokasi('cfLok', sumber());
          if (!daftar.length) return pesan('Daftar lokasi masih kosong.');
          if (ganti && !(await tanya({ judul: 'Ganti semua lokasi aktif?', pesan: angka(jumlah) + ' lokasi lama dihapus dan diganti ' + angka(daftar.length) + ' lokasi baru.', ya: 'Ganti semua', bahaya: true }))) return;
          janji = () => panggil('importLokasiAktif', this.saya(), id, daftar, ganti);
        }
        l.terkunci = true;
        const ok = await this.ubah(b, janji, () => { l.terkunci = false; l.tutup(); if (sesudah) sesudah(); });
        if (!ok) l.terkunci = false;
      } catch (e) { pesan(e.message); l.terkunci = false; }
    });
  },
  bukaFacilityBaru() {
    const lain = (this.facility || []).filter((f) => f.status === 'Aktif');
    const l = Lembar.buka({
      judul: 'Tambah facility baru', penuh: true,
      isi: h`<p class="redup" style="margin-bottom:6px">Tiap facility punya spreadsheet sendiri. Spreadsheet baru dibuat otomatis di Google Drive akun yang memasang aplikasi ini.</p>
        <label class="lbl" for="cfBaruNama">Nama facility</label><input class="field" id="cfBaruNama" type="text" placeholder="contoh: RDC Manado">
        <label class="lbl" for="cfBaruKode">Kode</label><input class="field" id="cfBaruKode" type="text" placeholder="contoh: RDC-MDC">
        <label class="lbl" for="cfBaruSs">Nama spreadsheet</label><input class="field" id="cfBaruSs" type="text" placeholder="contoh: Cycle Count RDC Manado">
        <span class="lbl">Daftar lokasi aktif</span>${this.sumberLokasiHtml('cfBaru', lain)}
        <div class="pesan pesan--galat" id="cfBaruPesan" role="alert"></div>
        <button type="button" class="btn" id="cfBaruSimpan" style="margin-top:16px">Buat facility</button>`
    });
    const sumber = this.pasangSumber('cfBaru');
    let jalan = false;
    el('cfBaruSimpan').addEventListener('click', async (ev) => {
      const b = ev.currentTarget, pesan = (m) => { el('cfBaruPesan').textContent = m; };
      const nama = el('cfBaruNama').value.trim(), kode = el('cfBaruKode').value.trim(), ss = el('cfBaruSs').value.trim();
      if (!nama || !kode || !ss) return pesan('Nama facility, kode, dan nama spreadsheet wajib diisi.');
      if (jalan) return;
      jalan = true; l.terkunci = true; b.classList.add('is-sibuk'); pesan('');
      try {
        let daftar;
        if (sumber() === 'salin') {
          const dari = el('cfBaruDari').value; if (!dari) return pesan('Pilih facility sumber.');
          const r = await panggil('getDaftarLokasiAktif', this.saya(), dari); if (r.success === false) throw new Error(r.message); daftar = r.lokasi || [];
        } else daftar = await this.bacaLokasi('cfBaru', sumber());
        if (!daftar.length) return pesan('Daftar lokasi aktif masih kosong.');
        const jobId = 'job_' + Date.now() + '_' + Math.floor(Math.random() * 1e5);
        Sibuk.tampil('Membuat facility baru', angka(daftar.length) + ' lokasi'); Sibuk.pantau(jobId);
        const res = await panggil('tambahFacility', this.saya(), nama, kode, ss, daftar, jobId);
        Sibuk.sembunyi();
        if (res.success === false) return pesan(res.message);
        l.terkunci = false; l.tutup(); Toast.tampil(res.message || 'Facility dibuat.'); this.facility = null; this.muat(true, 'facility');
      } catch (e) { Sibuk.sembunyi(); pesan(e.message); }
      finally { jalan = false; l.terkunci = false; b.classList.remove('is-sibuk'); }
    });
  }
};

Layar.config = { masuk() { if (!Boleh.config()) { Nav.ke(Nav.awal()); return; } Config.gambar(); Config.muat(false); } };
Aksi['config-segar'] = () => Config.muat(true);
Aksi['config-tab'] = (t) => { Config.tab = t.dataset.v; delete Config.galat[Config.tab]; Config.gambar(); Config.muat(false); };
Aksi['config-user-tambah'] = (t) => {
  const nik = el('cfNik').value.trim();
  if (!nik) { Toast.galat(new Error('NIK belum diisi.')); el('cfNik').focus(); return; }
  if (!el('cfPeran').value) { Toast.galat(new Error('Pilih peran untuk user ini.')); el('cfPeran').focus(); return; }
  Config.ubah(t, () => panggil('tambahUserMaster', Config.saya(), nik, el('cfPeran').value), () => { Config.users = null; Upload.form = null; Config.muat(true, 'user'); });
};
Aksi['config-user-buka'] = (t) => Config.bukaUser(t.dataset.row);
Aksi['config-mode'] = (t) => Config.gantiMode(t, t.dataset.v);
Aksi['config-matriks-simpan'] = (t) => Config.simpanMatriks(t);
Aksi['config-cara'] = (t) => { Config.grup(t.dataset.grup).mode = t.dataset.v; isi(el('cfCara'), Config.caraHtml()); };
Aksi['config-log'] = () => { Config.logSemua = !Config.logSemua; Config.gambarTetap(); };
Aksi['config-akses-tambah'] = (t) => {
  const nik = el('cfAksesNik').value.trim();
  if (!nik) { Toast.galat(new Error('NIK belum diisi.')); el('cfAksesNik').focus(); return; }
  Config.ubah(t, () => panggil('tambahAksesSetting', Config.saya(), nik, el('cfAksesNama').value.trim()), () => { Config.akses = null; Config.muat(true, 'akses'); });
};
Aksi['config-akses-cabut'] = async (t) => {
  if (!(await tanya({ judul: 'Cabut akses ' + t.dataset.nik + '?', pesan: 'NIK ini tidak bisa lagi membuka Config.', ya: 'Cabut akses', bahaya: true }))) return;
  Config.ubah(null, () => panggil('hapusAksesSetting', Config.saya(), Number(t.dataset.row), t.dataset.nik), () => { Config.akses = null; Config.muat(true, 'akses'); });
};
Aksi['config-fac-baru'] = () => Config.bukaFacilityBaru();
Aksi['config-fac-ubah'] = (t) => Config.ubahFacility(t.dataset.id);
Aksi['config-fac-status'] = (t) => Config.statusFacility(t, t.dataset.id);
Aksi['config-fac-user'] = (t) => Config.userFacility(t.dataset.id);
Aksi['config-fac-lokasi'] = (t) => { const f = Config.fac(t.dataset.id); if (f) Config.bukaLokasi(f.id, f.nama, null); };
