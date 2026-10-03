/* Mulai: masuk/keluar aplikasi, tombol kembali Android, kembali dari latar belakang. */

function tampilLogin() {
  Lembar.tutupSemua();
  const lama = Nav.aktif();
  if (lama && Layar[lama] && Layar[lama].keluar) Layar[lama].keluar();
  Nav.tab = 'login'; Nav.halaman = [];
  el('nav').hidden = true;
  document.body.classList.add('tanpa-nav');
  Nav._tampil('login');
  Layar.login.masuk();
}

function masukAplikasi(user) {
  Sesi.simpan(user);
  // data layar milik pengguna sebelumnya tidak boleh terbawa
  Home.mulai(); Tugas.reset(); Verifikasi.reset(); Upload.reset(); Report.reset(); Config.reset();
  Nav.lencana = {};
  el('nav').hidden = false;
  Nav.bangun();
  if (Demo.aktif) Pita.pasang('demo', h`Mode demo. Data contoh, perubahan tidak tersimpan.`); else Pita.lepas('demo');
  Nav.ke('home');
}

function keluarAplikasi(pesan) {
  const demo = Demo.aktif;
  Sesi.hapus(); Simpanan.bersih();
  Pita.lepas('demo');
  if (demo) {
    Antrean.kosongkan();
    if (!CFG.demoSaja) { Demo.berhenti(); Antrean.muat(); Antrean.jalan(); } // antrean sungguhan (bila ada) dilanjutkan
  }
  tampilLogin();
  if (pesan) Masuk.pesan(pesan);
}

// Sesi tersimpan dipakai dulu supaya aplikasi langsung terbuka, lalu dicek ke server di belakang.
let sesiDicekAt = 0;
async function periksaSesi() {
  const u = Sesi.user; if (!u || Demo.aktif) return;
  sesiDicekAt = Date.now();
  try {
    const res = await panggil('getUserRole', u.username);
    if (Sesi.user !== u) return; // sesi sudah diganti selagi menunggu (keluar, ganti akun, ganti facility): jawaban ini basi
    if (!res) { keluarAplikasi('Akun ini tidak lagi terdaftar. Hubungi admin.'); return; }
    const baru = Sesi.dari(res);
    const berubah = baru.role !== u.role || baru.akses !== u.akses || baru.facilityId !== u.facilityId || baru.facilityName !== u.facilityName;
    if (!berubah) { Sesi.simpan(baru); return; }
    // Peran, akses, atau facility diubah admin: semua yang sudah dimuat (tugas, laporan, simpanan) milik keadaan lama.
    // Masuk ulang dari awal supaya tidak ada data facility atau menu lama yang tersisa di layar.
    Lembar.tutupSemua(); Simpanan.bersih();
    // Nomor tugas hanya berlaku di facility asalnya: hasil yang belum terkirim dari facility lama tidak bisa dipakai lagi.
    const gugur = baru.facilityId !== u.facilityId ? Antrean.buangFacilityLain(baru.username, baru.facilityId || '-') : 0;
    masukAplikasi(baru);
    Toast.tampil(gugur ? 'Akun Anda dipindahkan admin ke facility lain. ' + angka(gugur) + ' hasil yang belum terkirim tidak berlaku lagi.' : 'Akun Anda diperbarui admin. Data dimuat ulang.', { lama: gugur ? 8000 : 4500 });
  } catch (e) { /* offline: tetap pakai sesi tersimpan */ }
}

function saatAktif() {
  periksaJaringan(); Server.sinkron(false); Pembaruan.cek(false);
  if (!Sesi.user) { Antrean.cobaLagi(); if (Nav.tab === 'login' && Masuk.status !== 'siap' && Masuk.status !== 'demo') Masuk.cekServer(); return; }
  // Peran, facility, atau akses bisa berubah selagi aplikasi di latar belakang: akun dicek dulu, baru hasil yang tertahan dikirim.
  if (Date.now() - sesiDicekAt > 5 * 60e3) Antrean.tahanSampai(periksaSesi(), 4000);
  Antrean.cobaLagi();
  const a = Nav.aktif();
  if (a === 'home') Home.muat(false);
  else if (a === 'cycle' || a === 'validasi') Tugas.muat(a, false);
}

function mulai() {
  Tema.pasang();
  // area aman (bilah sistem Android) dipasang plugin sesaat setelah halaman dimuat: ukur lagi sesudahnya
  ukurLayar(); setTimeout(ukurLayar, 400); setTimeout(ukurLayar, 1500);
  Pembaruan.periksaTertunda();
  Antrean.muat();
  Masuk.gambar();
  periksaJaringan();

  const L = Pembaruan.plugin();
  if (L) L.notifyAppReady().catch(() => {});
  if (NATIVE && PL.App) {
    PL.App.addListener('backButton', () => { if (!Nav.kembali() && PL.App.minimizeApp) PL.App.minimizeApp(); });
    PL.App.addListener('appStateChange', (s) => { if (s && s.isActive) saatAktif(); });
  } else {
    document.addEventListener('visibilitychange', () => { if (!document.hidden) saatAktif(); });
  }

  const u = CFG.demoSaja ? null : Sesi.muat();
  if (u) { masukAplikasi(u); Antrean.tahanSampai(periksaSesi(), 4000); }
  else {
    tampilLogin();
    // Versi 1.x menyimpan NIK saja: pengguna yang sudah masuk tidak perlu mengetik ulang setelah diperbarui.
    const nikLama = CFG.demoSaja ? null : LS.get('ct.user');
    if (nikLama && typeof nikLama === 'string') { LS.del('ct.user'); el('lgNik').value = nikLama; if (Server.url()) Masuk.kirim(nikLama); }
  }
  Server.sinkron(false);
  Antrean.jalan();
  setTimeout(() => Pembaruan.cek(false), 1500);
}

window.CT = { panggil, Sesi, Antrean, Nav, Home, Tugas, Fokus, Demo, Server, Pembaruan, versi: CFG.version };
if (document.readyState !== 'loading') mulai(); else document.addEventListener('DOMContentLoaded', mulai);
