/* Akun: tema, pembaruan aplikasi (update kilat / APK), lembar akun, keluar. */

/* ---------- tema ---------- */
const Tema = {
  gelap() { return document.documentElement.getAttribute('data-tema') === 'gelap'; },
  pasang() { this._bilah(); },
  atur(gelap) {
    if (gelap) document.documentElement.setAttribute('data-tema', 'gelap'); else document.documentElement.removeAttribute('data-tema');
    LS.set('ct.tema', gelap ? 'gelap' : 'terang');
    this._bilah();
  },
  // Ikon bilah status Android mengikuti tema aplikasi (terang di atas latar gelap, dan sebaliknya)
  _bilah() { try { if (NATIVE && PL.SystemBars && PL.SystemBars.setStyle) PL.SystemBars.setStyle({ style: this.gelap() ? 'DARK' : 'LIGHT' }).catch(() => {}); } catch (e) { /* abaikan */ } }
};

/* ---------- pembaruan aplikasi ---------- */
const Pembaruan = {
  rilis: null, sibuk: false, berkas: '', dicekAt: 0, lembar: null, mengecek: false, _manual: false,
  _bagian(v) { return String(v || '').replace(/^v/i, '').split('.').map((x) => parseInt(x, 10) || 0); },
  lebihBaru(a, b) { const x = this._bagian(a), y = this._bagian(b); for (let i = 0; i < Math.max(x.length, y.length); i++) { if ((x[i] || 0) !== (y[i] || 0)) return (x[i] || 0) > (y[i] || 0); } return false; },
  plugin() { try { return NATIVE && CAP.isPluginAvailable && CAP.isPluginAvailable('CapacitorUpdater') ? PL.CapacitorUpdater : null; } catch (e) { return null; } },
  bisaKilat(rel) { const l = rel && rel.live; if (!l || !this.plugin() || !CFG.nativeBase || l.native !== CFG.nativeBase) return false; return (LS.get('ct.live.failed') || []).indexOf(rel.version) < 0; },
  // latest.json adalah berkas rilis biasa (bukan GitHub API), jadi tidak kena batas 60 permintaan per jam
  // per alamat IP. Penting karena puluhan HP gudang memakai satu jaringan WiFi.
  async ambil() {
    const base = 'https://github.com/' + CFG.updateRepo + '/releases/latest/download/';
    const j = await httpJson(base + 'latest.json');
    const rel = { version: String(j.version || ''), notes: String(j.notes || ''), url: base + j.apk, size: Number(j.apkSize) || 0, live: null };
    if (j.live && /^[0-9a-f]{64}$/.test(j.live.sha256 || '')) rel.live = { native: String(j.live.native), sha256: j.live.sha256, size: Number(j.live.size) || 0, url: base + j.live.asset };
    return rel;
  },
  // Hanya satu pemeriksaan yang berjalan. Ketukan "Periksa pembaruan" saat pemeriksaan otomatis masih jalan
  // tidak memulai yang kedua: hasil pemeriksaan yang sedang berjalan itulah yang menjawabnya.
  async cek(manual) {
    if (!CFG.updateRepo || CFG.version === 'dev' || CFG.demoSaja) { if (manual) Toast.tampil('Pembaruan hanya tersedia di aplikasi Android.'); return; }
    const baris = () => el('akVersi'); // baris "Periksa pembaruan" di lembar Akun (bila sedang terbuka)
    if (manual) this._manual = true;
    if (this.mengecek) return;
    if (!manual && Date.now() - this.dicekAt < 15 * 60e3) return;
    this.mengecek = true; this.dicekAt = Date.now();
    if (baris()) baris().textContent = Akun.versiTeks();
    let salah = null;
    try {
      const rel = await this.ambil(); LS.set('ct.upd.latest', rel);
      this.rilis = rel.version && this.lebihBaru(rel.version, CFG.version) ? rel : null;
    } catch (e) {
      salah = e;
      const c = LS.get('ct.upd.latest'); if (c && this.lebihBaru(c.version, CFG.version)) this.rilis = c;
    }
    this.mengecek = false;
    const diminta = this._manual; this._manual = false;
    this.gambarKartu();
    if (baris()) baris().textContent = Akun.versiTeks();
    if (!diminta) return;
    if (this.rilis) {
      // Jawaban bisa datang lama setelah tombol ditekan. Lembar lain yang sedang dipakai (formulir verifikasi,
      // misalnya) tidak boleh ikut tertutup: hanya lembar Akun yang diganti, selain itu cukup ditawarkan.
      const atas = Lembar.tumpukan[Lembar.tumpukan.length - 1];
      if (!atas) this.buka();
      else if (Akun.lembar && atas === Akun.lembar) { atas.tutup(); this.buka(); }
      else Toast.tampil('Versi ' + this.rilis.version + ' tersedia.', { aksi: 'Lihat', onAksi: () => this.buka(), lama: 8000 });
    }
    else if (salah) Toast.galat(new Error('Tidak bisa memeriksa pembaruan: ' + galatJaringan(salah).message));
    else Toast.tampil('Sudah versi terbaru (' + CFG.version + ').');
  },
  gambarKartu() {
    const n = el('hmPembaruan'); if (!n) return;
    isi(n, this.rilis ? h`<button type="button" class="lanjut lanjut--biru" data-aksi="pembaruan-buka" style="margin-bottom:12px"><span class="lanjut__ic">${ikon('unduh')}</span><span class="lanjut__teks"><b>Versi ${this.rilis.version} tersedia</b><span>Ketuk untuk memperbarui aplikasi.</span></span>${ikon('kanan')}</button>` : '');
  },
  catatan(md) {
    const s = String(md || '').replace(/\r/g, '').replace(/<!--[\s\S]*?-->/g, ''), i = s.search(/^## Catatan developer/m), u = i >= 0 ? s.slice(0, i) : s;
    const baris = u.split('\n').map((l) => l.trim()).filter((l) => l && !/^## /.test(l) && !/Unduh file \.apk/i.test(l)).slice(0, 12);
    return baris.length ? baris.map((l) => h`<li>${l.replace(/^[-*]\s*/, '')}</li>`) : h`<li>Perbaikan dan peningkatan.</li>`;
  },
  buka() {
    const rel = this.rilis; if (!rel || this.lembar) return; // tidak menumpuk dua lembar pembaruan
    const kilat = this.bisaKilat(rel);
    this.lembar = Lembar.buka({
      judul: 'Versi ' + rel.version + ' tersedia',
      isi: h`<ul class="upd__catatan">${this.catatan(rel.notes)}</ul>
        <div class="bar bar--tipis bar--biru" id="updBar" hidden><i></i></div><div class="pesan" id="updStatus" aria-live="polite"></div>
        <button type="button" class="btn" id="updTombol" data-aksi="pembaruan-jalan" style="margin-top:14px">${kilat ? 'Perbarui sekarang' : 'Unduh dan pasang'}</button>
        <p class="ket" style="text-align:center">${kilat ? 'Langsung terpasang, tanpa mengunduh APK.' : 'Data tetap aman.'}</p>`,
      onTutup: () => { this.lembar = null; }
    });
    // Lembar ditutup lalu dibuka lagi selagi unduhan berjalan: tombolnya tidak boleh terlihat siap ditekan.
    if (this.sibuk) { const go = el('updTombol'); go.disabled = true; go.textContent = 'Mengunduh…'; }
  },
  _status(p, pesan) {
    const b = el('updBar'), s = el('updStatus');
    if (b) { b.hidden = p == null; if (p != null) $('i', b).style.width = Math.max(3, Math.min(100, p)) + '%'; }
    if (s) isi(s, pesan || '');
  },
  async jalan() {
    const rel = this.rilis; if (!rel || this.sibuk) return;
    const go = el('updTombol'); let hd = null;
    if (rel.live && this.bisaKilat(rel)) {
      this.sibuk = true; if (go) go.disabled = true; const L = this.plugin();
      try {
        hd = await L.addListener('download', (e) => { const p = Number(e && e.percent) || 0; this._status(p, p < 70 ? 'Mengunduh… ' + p + '%' : 'Memasang…'); });
        this._status(2, 'Mengunduh…');
        const b = await L.download({ url: rel.live.url, version: rel.version, checksum: rel.live.sha256 });
        this._status(100, 'Memuat ulang aplikasi…'); LS.set('ct.live.pending', { v: rel.version, from: CFG.version }); await L.set({ id: b.id });
      } catch (e) {
        LS.del('ct.live.pending'); this.sibuk = false; rel.live = null;
        if (go) { go.disabled = false; go.textContent = 'Unduh dan pasang'; }
        this._status(null, h`Pembaruan langsung gagal (${(e && e.message) || e}). Ketuk tombol untuk mengunduh APK.`);
      } finally { try { hd && hd.remove(); } catch (x) { /* abaikan */ } }
      return;
    }
    const FT = PL.FileTransfer, FS = PL.Filesystem, FO = PL.FileOpener;
    if (!NATIVE || !FT || !FS || !FO) { window.open(rel.url, '_blank'); return; }
    if (this.berkas) return this.pasang();
    this.sibuk = true; if (go) { go.disabled = true; go.textContent = 'Mengunduh…'; }
    try {
      const nama = 'CycleTransaksi-' + rel.version + '.apk'; try { await FS.deleteFile({ path: nama, directory: 'CACHE' }); } catch (e) { /* belum ada */ }
      const uri = (await FS.getUri({ path: nama, directory: 'CACHE' })).uri, path = String(uri).replace(/^file:\/\//, ''), total = rel.size || 0;
      hd = await FT.addListener('progress', (p) => { const t = p.lengthComputable && p.contentLength ? p.contentLength : total; if (t) this._status((p.bytes / t) * 100, 'Mengunduh ' + (p.bytes / 1048576).toFixed(1) + ' dari ' + (t / 1048576).toFixed(1) + ' MB'); });
      this._status(1, 'Mengunduh…'); await FT.downloadFile({ url: rel.url, path, progress: true, connectTimeout: 20000, readTimeout: 60000 });
      this.berkas = uri; this.sibuk = false; if (go) { go.disabled = false; go.textContent = 'Pasang versi ' + rel.version; }
      this._status(100, h`Unduhan selesai. Pilih <b>Update</b> di layar berikutnya.`); await this.pasang();
    } catch (e) { this.sibuk = false; this.berkas = ''; if (go) { go.disabled = false; go.textContent = 'Coba lagi'; } this._status(null, 'Unduhan gagal: ' + galatJaringan(e).message); }
    finally { try { hd && hd.remove(); } catch (x) { /* abaikan */ } }
  },
  async pasang() {
    try { await PL.FileOpener.open({ filePath: this.berkas, contentType: 'application/vnd.android.package-archive', openWithDefault: true }); this._status(100, h`Penginstal Android terbuka. Kalau diminta, izinkan <b>Instal aplikasi tidak dikenal</b> untuk Cycle Transaksi, lalu kembali dan ketuk tombol lagi.`); }
    catch (e) { this._status(100, 'Tidak bisa membuka penginstal (' + ((e && e.message) || e) + ').'); }
  },
  // Update kilat yang gagal menampilkan aplikasi dibatalkan otomatis oleh plugin; versi itu ditandai
  // supaya berikutnya ditawarkan lewat APK.
  periksaTertunda() {
    const pend = LS.get('ct.live.pending'); if (!pend || !pend.v) return; LS.del('ct.live.pending'); if (pend.v === CFG.version) return;
    const f = LS.get('ct.live.failed') || []; if (f.indexOf(pend.v) < 0) f.push(pend.v); LS.set('ct.live.failed', f.slice(-10));
  }
};
Aksi['cek-versi'] = () => Pembaruan.cek(true);
Aksi['pembaruan-buka'] = () => { Lembar.tutupSemua(); Pembaruan.buka(); };
Aksi['pembaruan-jalan'] = () => Pembaruan.jalan();

/* ---------- lembar akun ---------- */
const Akun = {
  lembar: null,
  versiTeks() { return Pembaruan.mengecek ? 'Memeriksa…' : 'Versi ' + CFG.version + (Home.data && Home.data.version ? ', sistem ' + Home.data.version : ''); },
  buka() {
    const u = Sesi.user; if (!u) return;
    const antre = Antrean.jumlah();
    this.lembar = Lembar.buka({
      judul: 'Akun',
      isi: h`<div class="ak__siapa"><div class="hm-avatar">${inisial(u.username)}</div><div><div class="ak__nama">${u.username}</div><div class="redup">${namaPeran(u.role)}${u.facilityName ? h`, ${u.facilityName}` : ''}</div></div></div>
        ${antre ? h`<button type="button" class="info info--jingga" data-aksi="akun-kirim" style="margin-bottom:12px">${ikon('awan_kirim')}<span><b>${angka(antre)} hasil belum terkirim.</b> Ketuk untuk mengirim sekarang.</span></button>` : ''}
        <div class="list">
          <div class="row"><span class="row__isi"><span class="row__t">Tema gelap</span><span class="row__s">Untuk area gudang yang redup</span></span><button type="button" class="saklar" role="switch" aria-checked="${Tema.gelap()}" aria-label="Tema gelap" data-aksi="akun-tema"></button></div>
          ${Boleh.developer() ? h`<div class="row"><span class="row__isi"><span class="row__t">Facility aktif</span><select class="field field--kecil" id="akFacility" style="margin-top:8px" aria-label="Facility aktif"><option value="">Memuat daftar facility…</option></select></span></div>` : ''}
          ${Boleh.config() ? h`<button type="button" class="row" data-aksi="akun-config">${ikon('atur')}<span class="row__isi"><span class="row__t">Config</span><span class="row__s">User, pembagian tugas, akses, facility</span></span><span class="row__ekor">${ikon('kanan', 'ic--kecil')}</span></button>` : ''}
          <button type="button" class="row" data-aksi="cek-versi">${ikon('unduh')}<span class="row__isi"><span class="row__t">Periksa pembaruan</span><span class="row__s" id="akVersi">${this.versiTeks()}</span></span>${Pembaruan.rilis ? h`<span class="pill pill--biru">Versi baru</span>` : ''}</button>
          <button type="button" class="row" data-aksi="akun-keluar" style="color:var(--merah-teks)">${ikon('keluar')}<span class="row__isi"><span class="row__t">${Demo.aktif && !CFG.demoSaja ? 'Keluar dari mode demo' : 'Keluar'}</span></span></button>
        </div>`,
      onTutup: () => { this.lembar = null; }
    });
    if (Boleh.developer()) this.muatFacility();
  },
  async muatFacility() {
    const sel = el('akFacility'); if (!sel) return;
    try {
      const res = await panggil('getDaftarFacility', Sesi.user.username);
      if (!res.success) throw new Error(res.message);
      isi(sel, h`<option value="">Assignment normal saya</option>${res.facilities.filter((f) => f.status === 'Aktif').map((f) => h`<option value="${f.id}" ${Sesi.user.facilityId === f.id ? 'selected' : ''}>${f.nama} (${f.kode})</option>`)}`);
      sel.addEventListener('change', () => this.gantiFacility(sel.value));
    } catch (e) { isi(sel, h`<option value="">Daftar facility tidak bisa dimuat</option>`); }
  },
  async gantiFacility(id) {
    const sel = el('akFacility'); if (sel) sel.disabled = true;
    try {
      const res = await panggil('setDeveloperActiveFacility', Sesi.user.username, id);
      if (!res.success) throw new Error(res.message);
      const info = await panggil('getUserRole', Sesi.user.username);
      if (info) { Sesi.simpan(Sesi.dari(info)); Simpanan.bersih(); Home.mulai(); Tugas.reset(); Verifikasi.reset(); Upload.reset(); Report.reset(); Config.reset(); }
      Lembar.tutupSemua(); Toast.tampil(res.message); Nav.ke('home'); Home.muat(true);
    } catch (e) { Toast.galat(e); if (sel) sel.disabled = false; }
  },
  async keluar() {
    const antre = Demo.aktif ? 0 : Antrean.jumlah();
    if (antre && !(await tanya({ judul: 'Keluar sekarang?', pesan: angka(antre) + ' hasil belum terkirim. Hasil itu tetap dikirim atas nama ' + Sesi.user.username + ' begitu jaringan tersedia.', ya: 'Keluar' }))) return;
    keluarAplikasi('');
  }
};
Aksi['akun'] = () => Akun.buka();
Aksi['akun-tema'] = (t) => { const g = !Tema.gelap(); Tema.atur(g); t.setAttribute('aria-checked', String(g)); };
Aksi['akun-config'] = () => { Lembar.tutupSemua(); Nav.dorong('config'); };
Aksi['akun-keluar'] = () => Akun.keluar();
Aksi['akun-kirim'] = () => { Antrean.cobaLagi(); Lembar.tutupSemua(); Toast.tampil('Mengirim hasil yang tertunda…'); };
