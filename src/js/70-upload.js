/*
 * Upload Data (admin): dua file dari WMS → pilih petugas → alat bantu → bagi tugas.
 * File dibaca di HP (SheetJS) dan langsung diperiksa begitu dipilih, jadi kolom yang salah ketahuan
 * sebelum menekan tombol proses. Hanya baris stok yang lokasinya ada di file transaksi yang dikirim,
 * supaya unggahan tetap kecil di jaringan seluler.
 */
const ALIAS_TRANSAKSI = {
  type: ['type', 'trantype', 'jenis transaksi', 'transaction type', 'tipe'],
  article: ['article', 'sku', 'kode barang', 'item code'],
  description: ['description', 'deskripsi', 'nama barang', 'item description'],
  lokasiAwal: ['from location', 'fromloc', 'lokasi awal', 'source location', 'lokasi picking'],
  lokasiTujuan: ['to location', 'toloc', 'lokasi tujuan', 'destination location', 'lokasi move'],
  qty: ['qty', 'quantity', 'qty transaksi', 'jumlah'],
  addWho: ['addwho', 'add who', 'added by', 'user wms', 'created by'] // opsional
};
const ALIAS_STOK = {
  lokasi: ['location', 'loc', 'lokasi'],
  article: ['article', 'sku', 'kode barang', 'item code'],
  qty: ['qty', 'quantity', 'stock', 'jumlah']
};
const URUT_PERAN = ['outbound', 'inbound', 'storing', 'lp', 'maintenance', 'inventory', 'admin'];

// Ekspor yang ditujukan untuk Excel sering menulis kode sebagai ="00123" supaya tidak dianggap angka.
const teksSel = (v) => { const s = v === undefined || v === null ? '' : String(v).trim(); const m = /^="(.*)"$/.exec(s); return m ? m[1].replace(/""/g, '"').trim() : s; };
// Qty dari .csv dibaca sebagai teks: "1,234" (pemisah ribuan) menjadi 1234; tanda kosong ("-", NULL) menjadi
// kosong. Teks lain yang bukan angka ditandai `salah`: di server ia diam-diam menjadi 0.
const QTY_KOSONG = /^(-|–|—|null|nan|n\/a|#n\/a)$/i;
function angkaQty(teks) {
  const s = teks.replace(/\s+/g, '');
  if (s === '' || QTY_KOSONG.test(s)) return { nilai: '' };
  const t = /^-?\d{1,3}(,\d{3})+(\.\d+)?$/.test(s) ? s.replace(/,/g, '') : s;
  return /^-?\d+(\.\d+)?$/.test(t) ? { nilai: String(Number(t)) } : { nilai: teks, salah: true };
}

async function bacaBerkasTabel(file, alias, opsional) {
  const X = await muatXlsx();
  const buf = await new Promise((ok, gagal) => { const r = new FileReader(); r.onload = () => ok(r.result); r.onerror = () => gagal(new Error('File tidak bisa dibaca.')); r.readAsArrayBuffer(file); });
  // raw: isi .csv dibaca sebagai teks apa adanya. Tanpa ini SKU "00123" menjadi 123 dan SKU 19 digit
  // kehilangan angka belakangnya, sehingga tidak lagi cocok dengan file stok.
  const wb = X.read(new Uint8Array(buf), { type: 'array', raw: true });
  const rows = X.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 });
  if (!rows.length) throw new Error('File kosong.');
  const norm = (v) => String(v || '').trim().toLowerCase().replace(/[_\s-]+/g, ' ');
  const kepalaBaris = rows[0].map(norm), idx = {}, hilang = [];
  Object.keys(alias).forEach((f) => { idx[f] = kepalaBaris.findIndex((x) => alias[f].indexOf(x) >= 0); if (idx[f] < 0 && (opsional || []).indexOf(f) < 0) hilang.push(alias[f][0]); });
  if (hilang.length) throw new Error('Kolom tidak ditemukan: ' + hilang.join(', ') + '. Periksa judul kolom di baris pertama.');
  const kunci = Object.keys(alias), out = [];
  rows.forEach((r, i) => {
    if (i === 0 || !r || !r.length) return;
    const o = {};
    kunci.forEach((f) => { o[f] = teksSel(idx[f] >= 0 ? r[idx[f]] : ''); });
    // qty yang bukan angka tidak langsung menolak file: baris total, judul yang terulang, atau lokasi tanpa
    // transaksi tidak pernah dikirim. Yang dipersoalkan hanya baris yang benar-benar dipakai (Upload.periksaStok).
    if ('qty' in o) { const q = angkaQty(o.qty); o.qty = q.nilai; if (q.salah) { o.qtySalah = true; o.baris = i + 1; } }
    out.push(o);
  });
  return out;
}

const Upload = {
  form: null, sibuk: false, galat: null, transaksi: null, stok: null, pilih: {}, rt: 0, tangga: 0, terbuka: false, lama: false, req: 0,

  // Dipanggil tiap pengguna atau facility berganti; jawaban permintaan yang masih berjalan diabaikan.
  reset() { Object.assign(this, { form: null, sibuk: false, galat: null, transaksi: null, stok: null, pilih: {}, terbuka: false, req: this.req + 1 }); },
  kunciPilih() { return 'ct.petugas.' + (Sesi.user.facilityId || '-'); },

  async muat(paksa) {
    if (this.sibuk) return;
    if (this.form && !paksa) return;
    const id = ++this.req, siapa = Sesi.user.username;
    this.sibuk = true; this.galat = null; this.gambar();
    try {
      let f;
      try { f = await panggil('getUploadFormData', siapa); this.lama = false; }
      catch (e) {
        if (e.jenis !== 'fungsi-tak-ada') throw e;
        const r = await Promise.all([panggil('getAssignableUsers', siapa), panggil('getEquipmentReadyDefaults', siapa)]);
        f = { users: r[0], equipment: r[1], lokasiAktif: null, modeAssignment: 'legacy', facilityId: Sesi.user.facilityId, facilityName: Sesi.user.facilityName }; this.lama = true;
      }
      if (id !== this.req) return;
      this.form = f; this.rt = Number(f.equipment.reachTruck) || 0; this.tangga = Number(f.equipment.tangga) || 0;
      // pilihan petugas terakhir dipakai lagi: biasanya tim yang bertugas sama dari hari ke hari
      const terakhir = Demo.aktif ? null : LS.get(this.kunciPilih());
      this.pilih = {};
      if (terakhir && terakhir.length) f.users.forEach((u) => { if (terakhir.indexOf(u.username) >= 0) this.pilih[u.username] = true; });
      this.adaTerakhir = Object.keys(this.pilih).length > 0;
    } catch (e) { if (id !== this.req) return; this.galat = e; }
    finally { if (id === this.req) { this.sibuk = false; if (Nav.aktif() === 'upload') this.gambar(); } }
  },
  terpilih() { return this.form.users.filter((u) => this.pilih[u.username]); },
  peranUser(u) { return String(u.role || '').trim().toLowerCase(); },

  gambar() {
    const node = el('scr-upload'), f = this.form;
    const kanan = h`<button type="button" class="iconbtn${this.sibuk ? ' is-putar' : ''}" data-aksi="upload-segar" aria-label="Perbarui daftar petugas">${ikon('segar')}</button>`;
    if (!f) { isi(node, h`${kepala('Upload Data', '', kanan)}${this.galat ? gagalMuat(this.galat, 'upload-segar') : kerangka(4, 96)}`); return; }
    if (f.tanpaFacility) { isi(node, h`${kepala('Upload Data', '', kanan)}<section class="card">${kosong('gedung', 'Akun ini belum punya facility', 'Tugas dibagi per facility. Minta pemegang akses Config memasukkan akun Anda ke facility lebih dulu.')}</section>`); return; }
    const pilih = this.terpilih(), matrix = f.modeAssignment === 'matrix';
    const nonStoring = pilih.filter((u) => this.peranUser(u) !== 'storing').length;
    const perPeran = {};
    f.users.forEach((u) => { const r = this.peranUser(u); if (!perPeran[r]) perPeran[r] = { total: 0, pilih: 0 }; perPeran[r].total++; if (this.pilih[u.username]) perPeran[r].pilih++; });
    const urutPeran = Object.keys(perPeran).sort((a, b) => (URUT_PERAN.indexOf(a) < 0 ? 99 : URUT_PERAN.indexOf(a)) - (URUT_PERAN.indexOf(b) < 0 ? 99 : URUT_PERAN.indexOf(b)));
    const ubin = (jenis, nomor, judul, ket) => {
      const b = this[jenis];
      const status = !b ? h`<span class="up-berkas__ket">${ket}</span>`
        : b.sibuk ? h`<span class="up-berkas__ket">Membaca ${b.nama}…</span>`
          : b.galat ? h`<span class="up-berkas__ket up-berkas__ket--galat">${b.galat}</span>`
            : h`<span class="up-berkas__ket up-berkas__ket--ok">${b.nama}</span><span class="up-berkas__ket">${b.ringkas}</span>`;
      return h`<label class="up-berkas${b && !b.galat && !b.sibuk ? ' is-ok' : ''}${b && b.galat ? ' is-galat' : ''}">
        <input type="file" class="sr" accept=".xlsx,.xls,.csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,text/csv" data-berkas="${jenis}">
        <span class="up-berkas__ic">${b && b.sibuk ? h`<span class="putar"></span>` : ikon(b && !b.galat ? 'cek' : 'berkas')}</span>
        <span class="up-berkas__isi"><span class="up-berkas__judul">${nomor}. ${judul}</span>${status}</span>
        <span class="up-berkas__aksi">${b ? 'Ganti' : 'Pilih file'}</span></label>`;
    };
    isi(node, h`${kepala('Upload Data', f.facilityName ? 'Bagi tugas cycle count untuk ' + f.facilityName : '', kanan)}
      ${f.lokasiAktif === 0 ? h`<div class="info info--merah" style="margin-bottom:12px">${ikon('awas')}<div><b>Daftar lokasi aktif masih kosong.</b> Tanpa daftar ini semua baris akan dilewati. ${Boleh.config() ? h`<button type="button" class="tautan" data-aksi="upload-lokasi">Impor lokasi aktif</button>` : 'Minta pemegang akses Config mengimpornya.'}</div></div>` : ''}
      <div class="tumpuk">
        ${ubin('transaksi', 1, 'Data transaksi', 'File ekspor dari WMS, .xlsx atau .csv')}
        ${ubin('stok', 2, 'Stock by location', 'Stok per lokasi, .xlsx atau .csv')}
      </div>
      ${Demo.aktif ? h`<button type="button" class="btn btn--teks" data-aksi="upload-contoh" style="margin:6px auto 0">Pakai file contoh</button>` : ''}
      <h2 class="bag"><span>3. Petugas yang bertugas</span><small>${angka(pilih.length)} dari ${angka(f.users.length)}</small></h2>
      <section class="card up-petugas">
        ${f.users.length ? h`
        <div class="chips up-peran" role="group" aria-label="Pilih per peran">${urutPeran.map((r) => h`<button type="button" class="chip${perPeran[r].pilih === perPeran[r].total ? ' is-on' : ''}" data-aksi="upload-peran" data-peran="${r}">${namaPeran(r)} <b>${perPeran[r].pilih}/${perPeran[r].total}</b></button>`)}</div>
        <div class="up-petugas__baris">
          <button type="button" class="tautan" data-aksi="upload-semua" data-v="1">Pilih semua</button>
          <button type="button" class="tautan" data-aksi="upload-semua" data-v="0">Kosongkan</button>
          <button type="button" class="tautan up-petugas__buka" data-aksi="upload-daftar">${this.terbuka ? 'Tutup daftar' : 'Lihat daftar'}</button>
        </div>
        ${this.adaTerakhir && !this.terbuka ? h`<p class="ket" style="margin-top:2px">Pilihan terakhir dipakai lagi. Ubah bila tim hari ini berbeda.</p>` : ''}
        ${this.terbuka ? h`<div class="up-daftar">${f.users.map((u) => h`<label class="row"><input type="checkbox" class="cek" data-user="${u.username}" ${this.pilih[u.username] ? 'checked' : ''}><span class="row__isi"><span class="row__t">${u.username}</span></span><span class="pill">${namaPeran(u.role)}</span></label>`)}</div>` : ''}`
      : h`<p class="redup">Belum ada petugas di facility ini. Tambahkan lewat Config.</p>`}
      </section>
      ${matrix ? h`<div class="info info--biru" style="margin-top:12px">${ikon('info')}<div>Pembagian tugas mengikuti <b>Mode Matrix</b> (peran per grup alat). Jumlah alat bantu tidak perlu diisi.</div></div>` : h`
      <h2 class="bag"><span>4. Alat bantu siap hari ini</span></h2>
      <section class="card">
        <div class="up-alat">
          <div><span class="lbl">Reach truck</span>${this.langkah('rt', this.rt)}</div>
          <div><span class="lbl">Tangga pesawat</span>${this.langkah('tangga', this.tangga)}</div>
        </div>
        <p class="ket">Petugas Storing mengerjakan Level 5-6 dengan reach truck. Petugas lain mengerjakan Level 3-4 sebanyak jumlah tangga; sisanya Level 1-2.</p>
        ${this.tangga > nonStoring && pilih.length ? h`<div class="pesan pesan--galat">Tangga (${this.tangga}) melebihi petugas terpilih di luar Storing (${nonStoring}).</div>` : ''}
      </section>`}
      <div class="pesan pesan--galat" id="upPesan" role="alert" style="margin-top:12px"></div>
      <button type="button" class="btn up-proses" id="upProses" data-aksi="upload-proses">Proses dan bagi tugas</button>
    `);
    $$('input[data-berkas]', node).forEach((inp) => inp.addEventListener('change', () => { if (inp.files && inp.files[0]) this.pilihBerkas(inp.dataset.berkas, inp.files[0]); }));
    $$('input[data-user]', node).forEach((cb) => cb.addEventListener('change', () => { if (cb.checked) this.pilih[cb.dataset.user] = true; else delete this.pilih[cb.dataset.user]; this.adaTerakhir = false; this.gambarSimpanGulir(); }));
  },
  langkah(nama, nilai) {
    return h`<div class="langkah"><button type="button" data-aksi="upload-langkah" data-n="${nama}" data-d="-1" aria-label="Kurangi">${ikon('kurang')}</button><output>${angka(nilai)}</output><button type="button" data-aksi="upload-langkah" data-n="${nama}" data-d="1" aria-label="Tambah">${ikon('tambah')}</button></div>`;
  },
  gambarSimpanGulir() { const y = window.scrollY; this.gambar(); window.scrollTo(0, y); },

  // Lokasi yang dihitung: tujuan untuk Move, asal untuk Picking (sama dengan backend).
  ringkasTransaksi(nama, rows) {
    let move = 0, picking = 0;
    const kunci = {};
    rows.forEach((r) => { const t = r.type.toLowerCase(); if (t === 'move') { move++; kunci[r.lokasiTujuan + '||' + r.article] = 1; } else if (t === 'picking') { picking++; kunci[r.lokasiAwal + '||' + r.article] = 1; } });
    if (!move && !picking) throw new Error('Tidak ada baris bertipe Move atau Picking di file ini.');
    return { nama, rows, kunci, ringkas: angka(rows.length) + ' baris: ' + angka(move) + ' Move, ' + angka(picking) + ' Picking' + (rows.length - move - picking ? ', ' + angka(rows.length - move - picking) + ' tipe lain dilewati' : '') };
  },
  async pilihBerkas(jenis, file) {
    this[jenis] = { nama: file.name, sibuk: true }; this.gambarSimpanGulir();
    try {
      if (jenis === 'transaksi') this.transaksi = this.ringkasTransaksi(file.name, await bacaBerkasTabel(file, ALIAS_TRANSAKSI, ['addWho']));
      else { const rows = await bacaBerkasTabel(file, ALIAS_STOK, null); this.stok = { nama: file.name, rows, ringkas: angka(rows.length) + ' baris stok' }; }
    } catch (e) { this[jenis] = { nama: file.name, galat: e.message }; }
    this.periksaStok();
    this.gambarSimpanGulir();
  },
  // Qty stok (qty sistem) harus angka untuk baris yang dipakai, yaitu lokasi + SKU yang ada di file transaksi:
  // qty sistem yang salah membuat setiap hitungan item itu tampak selisih. Dicek begitu kedua file ada.
  periksaStok() {
    const t = this.transaksi, s = this.stok;
    if (!s || !s.rows) return;
    delete s.galat;
    if (!t || !t.kunci) return;
    const x = s.rows.find((r) => r.qtySalah && t.kunci[r.lokasi + '||' + r.article]);
    if (x) s.galat = 'Baris ' + x.baris + ': qty "' + x.qty + '" bukan angka. Perbaiki file lalu pilih lagi.';
  },
  contoh() {
    const c = Demo._mesin.contoh();
    this.transaksi = this.ringkasTransaksi('transaksi-contoh.xlsx', c.transaksi);
    this.stok = { nama: 'stok-contoh.xlsx', rows: c.stok, ringkas: angka(c.stok.length) + ' baris stok' };
    if (!this.terpilih().length) this.form.users.forEach((u) => { if (this.peranUser(u) !== 'inventory' && this.peranUser(u) !== 'admin') this.pilih[u.username] = true; });
    this.gambarSimpanGulir();
  },

  async proses() {
    const f = this.form, pesan = (m) => { el('upPesan').textContent = m; };
    const pilih = this.terpilih();
    const belumSiap = (b, nama) => (!b ? 'Pilih file ' + nama + ' dulu.' : b.sibuk ? 'File ' + nama + ' masih dibaca. Tunggu sebentar.' : b.galat ? 'File ' + nama + ' bermasalah. Pilih file lain.' : '');
    this.periksaStok();
    const salah = belumSiap(this.transaksi, 'data transaksi') || belumSiap(this.stok, 'stock by location');
    if (salah) return pesan(salah);
    if (!pilih.length) return pesan('Pilih minimal satu petugas.');
    const matrix = f.modeAssignment === 'matrix', nonStoring = pilih.filter((u) => this.peranUser(u) !== 'storing').length;
    if (!matrix && this.tangga > nonStoring) return pesan('Jumlah tangga melebihi petugas terpilih di luar Storing (' + nonStoring + ').');
    pesan('');
    const transaksi = this.transaksi.rows.map((r) => [r.type, r.article, r.description, r.lokasiAwal, r.lokasiTujuan, r.qty, r.addWho]);
    let stok = this.stok.rows.filter((r) => this.transaksi.kunci[r.lokasi + '||' + r.article]).map((r) => [r.lokasi, r.article, r.qty]);
    if (!stok.length) stok = this.stok.rows.slice(0, 1).map((r) => [r.lokasi, r.article, r.qty]); // semua lokasi transaksi tanpa stok: qty sistem 0
    const nama = pilih.map((u) => u.username);
    if (!Demo.aktif) LS.set(this.kunciPilih(), nama);
    const jobId = 'job_' + Date.now() + '_' + Math.floor(Math.random() * 1e5);
    Sibuk.tampil('Memproses dan membagi tugas', angka(transaksi.length) + ' baris transaksi');
    Sibuk.pantau(jobId);
    try {
      const res = await panggil('importRawData', transaksi, stok, nama, Sesi.user.username, this.rt, this.tangga, jobId);
      Sibuk.persen(100); await tunggu(250); Sibuk.sembunyi();
      if (!res.success) { pesan(res.message); el('upPesan').scrollIntoView({ block: 'center' }); return; }
      this.transaksi = null; this.stok = null; Home.at = 0; Tugas.cycle.at = 0;
      this.gambar(); this.hasil(res, nama.length);
    } catch (e) { Sibuk.sembunyi(); pesan(e.message); el('upPesan').scrollIntoView({ block: 'center' }); }
  },
  hasil(res, jumlahPetugas) {
    // Server baru mengirim peringatan & catatan sebagai daftar; server lama hanya punya `warning` dan `message`.
    const peringatan = res.peringatan || (res.warning ? [res.warning] : []);
    const catatan = res.catatan || (res.message ? [String(res.message).split('⚠')[0].trim()] : []);
    const baris = (teks, n, warna) => (n > 0 ? h`<div class="row"><span class="row__isi"><span class="row__t row__t--lipat" style="font-weight:500">${teks}</span></span><span class="row__angka" style="${warna ? 'color:var(--' + warna + '-teks)' : ''}">${angka(n)}</span></div>` : '');
    Lembar.buka({
      judul: res.ditugaskan ? 'Tugas sudah dibagi' : 'Tidak ada item baru',
      isi: h`<div class="up-hasil"><b class="angka">${angka(res.ditugaskan || 0)}</b><span>item baru dibagi ke ${angka(jumlahPetugas)} petugas</span></div>
        ${res.merged || res.skippedInactive || res.skippedBlankLokasi || res.skippedOtherType ? h`<div class="list" style="margin-bottom:14px">
          ${baris('Digabung ke item yang masih menunggu', res.merged)}
          ${baris('Dilewati: lokasi tidak ada di daftar lokasi aktif', res.skippedInactive)}
          ${baris('Dilewati: kolom lokasi kosong', res.skippedBlankLokasi, 'jingga')}
          ${baris('Dilewati: tipe selain Move dan Picking', res.skippedOtherType)}
        </div>` : ''}
        ${peringatan.length ? h`<div class="info info--jingga" style="margin-bottom:12px">${ikon('awas')}<div><b>Ada item yang tidak dibagi.</b> ${peringatan.join(' ')} Rinciannya tercatat di sheet Log_Unassigned.</div></div>` : ''}
        ${catatan.map((c) => h`<div class="info" style="margin-bottom:10px">${ikon('info')}<div>${c}</div></div>`)}`,
      kaki: CFG.panel ? h`<button type="button" class="btn" data-tutup>${peringatan.length ? 'Mengerti' : 'Selesai'}</button>` : h`<button type="button" class="btn" data-aksi="ke" data-tab="home">${peringatan.length ? 'Mengerti, kembali ke Home' : 'Kembali ke Home'}</button>`
    });
  }
};

Layar.upload = { masuk() { Upload.gambar(); Upload.muat(false); } };
Aksi['upload-segar'] = () => Upload.muat(true);
Aksi['upload-daftar'] = () => { Upload.terbuka = !Upload.terbuka; Upload.gambarSimpanGulir(); };
Aksi['upload-semua'] = (t) => { Upload.pilih = {}; if (t.dataset.v === '1') Upload.form.users.forEach((u) => { Upload.pilih[u.username] = true; }); Upload.adaTerakhir = false; Upload.gambarSimpanGulir(); };
Aksi['upload-peran'] = (t) => {
  const r = t.dataset.peran, anggota = Upload.form.users.filter((u) => Upload.peranUser(u) === r), semua = anggota.every((u) => Upload.pilih[u.username]);
  anggota.forEach((u) => { if (semua) delete Upload.pilih[u.username]; else Upload.pilih[u.username] = true; });
  Upload.adaTerakhir = false; Upload.gambarSimpanGulir();
};
Aksi['upload-langkah'] = (t) => { const n = t.dataset.n; Upload[n] = Math.max(0, Math.min(99, Upload[n] + Number(t.dataset.d))); Upload.gambarSimpanGulir(); };
Aksi['upload-proses'] = () => Upload.proses();
Aksi['upload-contoh'] = () => Upload.contoh();
Aksi['upload-lokasi'] = () => { Config.bukaLokasi(Upload.form.facilityId, Upload.form.facilityName, () => Upload.muat(true)); };
