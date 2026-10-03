/* Layar masuk: cukup NIK. Alamat Web App ditanam di aplikasi, tidak pernah diisi pengguna. */

const GAMBAR_RAK = mentah(`<svg class="lg__rak" viewBox="0 0 360 208" role="img" aria-label="Rak gudang dengan satu label lokasi yang sudah dicentang">
  <line x1="8" y1="196" x2="352" y2="196" stroke="var(--garis-kuat)" stroke-width="2" stroke-linecap="round"/>
  <g fill="var(--kertas)" stroke="var(--garis-kuat)" stroke-width="1.5">
    <rect x="38" y="30" width="52" height="40" rx="5"/><rect x="96" y="40" width="40" height="30" rx="5"/><rect x="141" y="24" width="27" height="46" rx="5"/>
    <rect x="194" y="36" width="60" height="34" rx="5"/><rect x="260" y="26" width="60" height="44" rx="5"/>
    <rect x="38" y="90" width="44" height="42" rx="5"/><rect x="88" y="102" width="76" height="30" rx="5"/>
    <rect x="194" y="88" width="40" height="44" rx="5"/><rect x="240" y="96" width="40" height="36" rx="5"/><rect x="286" y="104" width="34" height="28" rx="5"/>
    <rect x="38" y="152" width="60" height="44" rx="5"/><rect x="104" y="160" width="60" height="36" rx="5"/><rect x="194" y="148" width="126" height="48" rx="5"/>
  </g>
  <g stroke="var(--garis-kuat)" stroke-width="1.5" stroke-linecap="round">
    <path d="M64 30v9M116 40v8M224 36v9M290 26v10M60 90v9M126 102v8M214 88v9M68 152v9M134 160v8M257 148v10"/>
  </g>
  <g fill="var(--jingga)"><rect x="22" y="70" width="316" height="7" rx="2"/><rect x="22" y="132" width="316" height="7" rx="2"/></g>
  <g fill="var(--biru)"><rect x="22" y="12" width="7" height="184" rx="2"/><rect x="176.5" y="12" width="7" height="184" rx="2"/><rect x="331" y="12" width="7" height="184" rx="2"/></g>
  <g class="lg__label">
    <rect x="52" y="118" width="118" height="36" rx="8" fill="var(--kuning)" stroke="var(--kuning-tepi)" stroke-width="1.5"/>
    <text x="111" y="144" text-anchor="middle" font-size="25" font-weight="700" fill="var(--tinta-label)" style="font-family:var(--huruf-rapat);letter-spacing:.02em">B03.016.2</text>
  </g>
  <g class="lg__centang">
    <circle cx="170" cy="118" r="14" fill="var(--biru)" stroke="var(--lantai)" stroke-width="3"/>
    <path d="m163.5 118.3 4.3 4.3 8.2-8.7" fill="none" stroke="var(--di-warna)" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round"/>
  </g>
</svg>`);

const Masuk = {
  status: 'cek',
  gambar() {
    isi(el('scr-login'), h`<div class="lg">
      <div class="lg__gambar">${GAMBAR_RAK}</div>
      <h1 class="lg__judul">Cycle Transaksi</h1>
      <p class="lg__ket">${CFG.panel ? 'Panel admin: upload data, facility gudang, dan config. Masuk dengan NIK admin.' : 'Masuk dengan NIK untuk melihat tugas hari ini.'}</p>
      <form id="lgForm" novalidate>
        <label class="lbl" for="lgNik">NIK</label>
        <input class="field" id="lgNik" type="text" inputmode="text" autocomplete="username" autocapitalize="off" autocorrect="off" spellcheck="false" enterkeyhint="go" placeholder="contoh: 123456.nama">
        <div class="pesan pesan--galat" id="lgPesan" role="alert"></div>
        <button class="btn lg__tombol" id="lgTombol" type="submit"><span>Masuk</span></button>
      </form>
      <div class="lg__status" id="lgStatus" aria-live="polite"></div>
      <div id="lgDemo"></div>
      <div class="lg__kaki"><span>Versi ${CFG.version}</span>${CFG.panel ? '' : h`<button type="button" class="tautan" data-aksi="cek-versi">Periksa pembaruan</button>`}</div>
    </div>`);
    el('lgForm').addEventListener('submit', (ev) => { ev.preventDefault(); this.kirim(); });
    el('lgNik').addEventListener('input', () => this.pesan(''));
    this.gambarStatus();
  },
  pesan(t) { el('lgPesan').textContent = t; el('lgNik').classList.toggle('is-galat', !!t); },
  setStatus(s) { this.status = s; this.gambarStatus(); },
  gambarStatus() {
    const n = el('lgStatus'); if (!n) return;
    const S = {
      cek: ['lg__titik--cek', 'Menyambungkan…'],
      siap: ['lg__titik--siap', 'Tersambung'],
      offline: ['lg__titik--awas', 'Tidak ada koneksi internet'],
      tanpa: ['lg__titik--awas', 'Aplikasi belum tersambung'],
      galat: ['lg__titik--awas', 'Belum bisa tersambung. Coba lagi nanti.'],
      demo: ['lg__titik--demo', 'Mode demo, memakai data contoh']
    }[this.status];
    isi(n, h`<i class="lg__titik ${S[0]}"></i><span>${S[1]}</span>`);
    const d = el('lgDemo');
    if (CFG.demoSaja || this.status === 'demo') {
      isi(d, h`<div class="lg__demo"><div class="lg__demo-judul">Lihat sebagai</div><div class="lg__demo-pilih">
        <button type="button" class="chip" data-aksi="demo-masuk" data-sebagai="admin">Admin</button>
        <button type="button" class="chip" data-aksi="demo-masuk" data-sebagai="inventory">Inventory</button>
        <button type="button" class="chip" data-aksi="demo-masuk" data-sebagai="petugas">Petugas</button></div></div>`);
    } else if (this.status === 'tanpa') {
      isi(d, h`<div class="info info--biru lg__info">${ikon('info')}<div>Aplikasi ini belum disambungkan ke data gudang, jadi NIK belum bisa dipakai. Hubungi admin aplikasi, atau lihat dulu tampilannya dengan data contoh.
        <button type="button" class="btn btn--kecil" data-aksi="demo-mulai" style="margin-top:10px">Coba mode demo</button></div></div>`);
    } else isi(d, '');
  },
  // Cek server begitu layar masuk tampil: pengguna tahu lebih dulu kalau ada masalah, dan panggilan pertama
  // "membangunkan" Apps Script sehingga login terasa lebih cepat.
  async cekServer() {
    if (CFG.demoSaja || Demo.aktif) { this.setStatus('demo'); return; }
    this.setStatus('cek');
    if (CFG.panel) { try { await panggil('getAppVersion'); this.setStatus('siap'); } catch (e) { this.setStatus('galat'); } return; }
    try {
      const url = Server.url() || await Server.sinkron(true);
      if (!url) { this.setStatus(navigator.onLine ? 'tanpa' : 'offline'); return; }
      await panggil('getAppVersion');
      this.setStatus('siap');
    } catch (e) { this.setStatus(e.jenis === 'offline' ? 'offline' : e.jenis === 'tanpa-server' ? 'tanpa' : 'galat'); }
  },
  async kirim(nikLangsung) {
    const nik = String(nikLangsung || el('lgNik').value || '').trim();
    if (!nik) { this.pesan('NIK belum diisi.'); el('lgNik').focus(); return; }
    const b = el('lgTombol');
    b.classList.add('is-sibuk'); this.pesan('');
    try {
      const res = await panggil('getUserRole', nik);
      if (!res) { this.pesan('NIK ini belum terdaftar atau sudah dinonaktifkan. Periksa ejaannya atau hubungi admin.'); return; }
      if (this.status !== 'demo') this.setStatus('siap');
      el('lgNik').value = '';
      masukAplikasi(Sesi.dari(res));
    } catch (e) {
      if (e.jenis === 'tanpa-server') this.setStatus('tanpa'); else if (e.jenis === 'offline') this.setStatus('offline');
      this.pesan(e.message);
    } finally { b.classList.remove('is-sibuk'); }
  },
  async demo(sebagai) {
    const b = el('lgTombol');
    b.classList.add('is-sibuk'); this.pesan('');
    try {
      const info = await Demo.mulai();
      Antrean.kosongkan();
      this.setStatus('demo');
      if (sebagai) await this.kirim(sebagai === 'admin' ? info.admin : sebagai === 'inventory' ? info.inventory : info.petugas);
    } catch (e) { this.pesan(e.message); } finally { b.classList.remove('is-sibuk'); }
  }
};
Layar.login = { masuk() { Masuk.cekServer(); } };
Aksi['demo-mulai'] = () => Masuk.demo('');
Aksi['demo-masuk'] = (t) => Masuk.demo(t.dataset.sebagai);
