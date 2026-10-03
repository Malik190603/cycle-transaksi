/*
 * Antrean kirim. Hasil hitung dan validasi disimpan dulu di HP, lalu dikirim satu per satu di latar
 * belakang. Petugas tidak menunggu server untuk lanjut ke lokasi berikutnya, dan hasil tidak hilang
 * saat sinyal putus di antara rak.
 *
 * Tiap entri: { id, jenis: 'cycle' | 'validasi', kunci, fac, args, info, at, kirimAt, coba, status }
 *   fac: facility pemilik tugas (nomor tugas hanya unik di dalam satu facility)
 *   status 'tahan' → masih bisa diurungkan (beberapa detik pertama)
 *   status 'antre' → menunggu giliran / menunggu jaringan
 *   status 'kirim' → sedang dikirim
 *
 * Server memproses antreannya sendiri sekitar semenit sekali, jadi item yang baru terkirim masih
 * berstatus Pending di sheet selama itu. Antrean.sembunyi() dipakai layar tugas supaya item
 * tersebut tidak muncul lagi.
 */
const Antrean = {
  daftar: [], terkirim: {}, _pendengar: [], _timer: 0, _sibuk: false,
  FUNGSI: { cycle: 'submitCount', validasi: 'submitValidasi' },
  LAMA_SEMBUNYI: 4 * 60e3,

  muat() {
    this.daftar = (LS.get('ct.antrean') || []).map((e) => Object.assign(e, { status: e.status === 'kirim' ? 'antre' : e.status }));
    this.terkirim = LS.get('ct.terkirim') || {};
    this._rapikan();
  },
  _simpan() { if (Demo.aktif) return; LS.set('ct.antrean', this.daftar); LS.set('ct.terkirim', this.terkirim); },
  _rapikan() { const batas = Date.now() - this.LAMA_SEMBUNYI; Object.keys(this.terkirim).forEach((k) => { if (this.terkirim[k] < batas) delete this.terkirim[k]; }); },
  kosongkan() { this.daftar = []; this.terkirim = {}; clearTimeout(this._timer); },

  dengar(fn) { this._pendengar.push(fn); },
  _kabar(tipe, entri, pesan) { this._pendengar.forEach((fn) => { try { fn({ tipe, entri, pesan }); } catch (e) { console.error(e); } }); },

  jumlah(jenis) { return this.daftar.filter((e) => !jenis || e.jenis === jenis).length; },
  _fac() { return (Sesi.user && Sesi.user.facilityId) || '-'; },
  sembunyi(jenis, kunci) {
    const fac = this._fac(), k = jenis + ':' + fac + ':' + kunci;
    if (this.daftar.some((e) => e.jenis === jenis && (e.fac || '-') === fac && String(e.kunci) === String(kunci))) return true;
    return !!this.terkirim[k] && Date.now() - this.terkirim[k] < this.LAMA_SEMBUNYI;
  },

  tambah(entri, tahanMs) {
    // entri sebelumnya tidak bisa diurungkan lagi begitu ada yang baru
    this.daftar.forEach((e) => { if (e.status === 'tahan') { e.status = 'antre'; e.kirimAt = Date.now(); } });
    const e = Object.assign({ id: Date.now().toString(36) + Math.random().toString(36).slice(2, 7), at: Date.now(), coba: 0, fac: this._fac() }, entri,
      { status: tahanMs ? 'tahan' : 'antre', kirimAt: Date.now() + (tahanMs || 0) });
    this.daftar.push(e); this._simpan(); this._kabar('ubah', e); this.jalan();
    return e;
  },
  urung(id) {
    const i = this.daftar.findIndex((e) => e.id === id);
    if (i < 0 || this.daftar[i].status !== 'tahan') return null;
    const e = this.daftar.splice(i, 1)[0];
    this._simpan(); this._kabar('ubah', e);
    return e;
  },
  lepas() { this.daftar.forEach((e) => { if (e.status === 'tahan') { e.status = 'antre'; e.kirimAt = Date.now(); } }); this.jalan(); },
  cobaLagi() { this.daftar.forEach((e) => { if (e.status === 'antre') e.kirimAt = Math.min(e.kirimAt, Date.now()); }); this.jalan(); },

  // Kirim ditahan sebentar selagi akun dicek ke server: peran atau facility bisa berubah selagi aplikasi tertutup,
  // dan hasil milik facility lama tidak boleh dikirim atas nama facility baru. Ada batas waktunya supaya sinyal
  // yang buruk tidak menahan antrean (server tetap memeriksa facility tiap hasil).
  tahanSampai(janji, maksMs) {
    this._ditahan = true;
    const lepas = () => { if (this._ditahan) { this._ditahan = false; this.jalan(); } };
    Promise.resolve(janji).then(lepas, lepas); setTimeout(lepas, maksMs);
  },
  // Hasil pengguna ini yang dibuat di facility lain tidak bisa lagi dikirim. Mengembalikan jumlah yang dibuang.
  buangFacilityLain(username, fac) {
    const u = String(username || '').toLowerCase(), sebelum = this.daftar.length;
    this.daftar = this.daftar.filter((e) => e.status === 'kirim' || String(e.args[1] || '').toLowerCase() !== u || (e.fac || '-') === (fac || '-'));
    if (this.daftar.length !== sebelum) this._simpan();
    return sebelum - this.daftar.length;
  },
  jalan() {
    clearTimeout(this._timer);
    if (this._sibuk || this._ditahan || !this.daftar.length) return;
    const now = Date.now();
    const siap = this.daftar.find((e) => e.kirimAt <= now);
    if (!siap) { const dekat = Math.min.apply(null, this.daftar.map((e) => e.kirimAt)); this._timer = setTimeout(() => this.jalan(), Math.max(50, dekat - now)); return; }
    this._kirim(siap);
  },
  // Galat dari server (bukan jaringan) dicoba ulang; hasil baru dikembalikan menjadi tugas setelah sedikitnya
  // sekian percobaan DAN sekian lama sejak galat pertama. Batas waktu perlu karena "kirim sekarang", kembali
  // ke aplikasi, dan sinyal yang naik-turun semuanya memicu percobaan langsung: tanpa batas waktu, gangguan
  // server beberapa menit bisa menghabiskan jatah percobaan dalam hitungan detik.
  MAKS_COBA_SERVER: 8, SABAR_SERVER_MS: 15 * 60e3,
  async _kirim(e) {
    this._sibuk = true; e.status = 'kirim'; this._kabar('ubah', e);
    let hasil = null, err = null;
    const pemilikSama = Sesi.user && String(e.args[1] || '').toLowerCase() === Sesi.user.username.toLowerCase();
    if (pemilikSama && (e.fac || '-') !== this._fac()) {
      // Nomor tugas hanya berlaku di facility asalnya; setelah akun dipindah, hasil ini bisa menimpa item lain.
      err = galat('akses', 'Facility akun ini berubah sebelum hasil terkirim.');
    } else {
      // Argumen terakhir: jam petugas menekan Simpan. Hasil bisa baru terkirim lama kemudian (sinyal putus),
      // dan jam itulah yang dicatat sebagai jam hitung. Server lama mengabaikan argumen ini.
      // Lalu facility tempat hasil dibuat: server menolak bila akun sudah dipindah (nomor tugas hanya unik per facility).
      try { hasil = await panggil.apply(null, [this.FUNGSI[e.jenis]].concat(e.args, [e.at, e.fac && e.fac !== '-' ? e.fac : ''])); } catch (x) { err = x; }
    }
    this._sibuk = false;
    const buang = () => { const i = this.daftar.indexOf(e); if (i >= 0) this.daftar.splice(i, 1); };
    const pesan = (hasil && hasil.message) || (err && err.message) || '';
    // Gangguan sambungan: hasil disimpan dan dicoba terus. Galat dari server sering hanya sementara
    // (server sibuk, antrean terkunci: backend menjawab "Gagal masuk antrian ..."), jadi dicoba beberapa
    // kali dulu; penolakan yang pasti (akun dicabut, qty tidak sah) langsung dikembalikan menjadi tugas.
    const sambungan = !!err && ['offline', 'lambat', 'jaringan', 'ditolak', 'tanpa-server', 'backend-lama'].indexOf(err.jenis) >= 0;
    const sementara = err ? err.jenis === 'server' : /coba lagi|sibuk|sebentar|gagal masuk antri/i.test(pesan);
    const habis = sementara && (e.cobaServer || 0) >= this.MAKS_COBA_SERVER && Date.now() - (e.gagalSejak || Date.now()) >= this.SABAR_SERVER_MS;
    if (hasil && hasil.success) {
      buang(); this.terkirim[e.jenis + ':' + (e.fac || '-') + ':' + e.kunci] = Date.now(); this._rapikan(); this._simpan(); this._kabar('terkirim', e);
    } else if (sambungan || (sementara && !habis)) {
      if (!sambungan) { e.cobaServer = (e.cobaServer || 0) + 1; e.gagalSejak = e.gagalSejak || Date.now(); }
      e.status = 'antre'; e.coba++; e.kirimAt = Date.now() + Math.min(60000, 3000 * Math.pow(2, Math.min(e.coba - 1, 5)));
      this._simpan(); this._kabar('tertunda', e, pesan);
    } else {
      // Gangguan sementara yang tidak kunjung pulih: pesan teknis server diganti kalimat yang bisa ditindaklanjuti.
      buang(); this._simpan(); this._kabar('ditolak', e, sementara ? 'Sistem sedang sibuk. Hitung ulang lokasi ini nanti.' : (pesan || 'Hasil tidak diterima.'));
      // Server menolak karena akun sudah dipindah facility: sesi di HP basi. Akun dicek saat itu juga (tanpa menunggu
      // jeda 5 menit), supaya hasil berikutnya dicap facility yang benar dan tidak terus-menerus ditolak.
      if (hasil && hasil.kode === 'FACILITY_BERUBAH') this.tahanSampai(periksaSesi(), 4000);
    }
    this.jalan();
  }
};
window.addEventListener('online', () => Antrean.cobaLagi());
