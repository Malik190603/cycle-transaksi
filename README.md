# Cycle Transaksi — Aplikasi Android

Aplikasi Android (APK) untuk cycle count gudang berbasis transaksi. **Database tetap Google Sheets** dan seluruh logika (pembagian tugas, antrean, worker, dashboard) berjalan di **Google Apps Script**. Aplikasi ini adalah tampilannya: dibuat khusus untuk HP, memanggil Web App lewat `doPost`.

```
HP Android (APK)                           Google
┌───────────────────────────┐   HTTPS    ┌──────────────────────────────────────┐
│ Tampilan  src/ (HTML/CSS/JS)│ ───────► │ Apps Script Web App (doPost)         │
│ Antrean kirim di HP        │           │   ApiBridge.gs → fungsi backend/src  │
│ Mode demo (backend di HP)  │ ◄──────── │ Google Sheets: Master DB, Config DB, │
└───────────────────────────┘            │   spreadsheet per facility           │
                                         └──────────────────────────────────────┘
```

## Yang ada di aplikasi

| Menu | Untuk | Isi |
|---|---|---|
| Home | semua | Satu saran tindakan berikutnya, empat KPI (Total Cycle, Outstanding Cycle, Belum Validasi, Selesai), Progress cycle hari ini, Summary plus minus. Petugas hanya melihat tugasnya sendiri. |
| Cycle | semua | Daftar tugas menurut urutan kerja (area lalu alat bantu), lalu **mode fokus**: satu lokasi per layar dengan keypad. Simpan langsung lanjut ke lokasi berikutnya, bisa diurungkan 5 detik. |
| Validasi | Inventory, Admin | Hitung ulang item yang selisih, dengan mode fokus yang sama. |
| Verifikasi | Inventory, Admin | Tindak lanjut selisih sampai ditutup dengan bukti WMS. Pasangan plus minus dicari otomatis. Riwayat task selesai. |
| Upload | Admin | Dua file WMS, pilih petugas, jumlah alat bantu, bagi tugas. File diperiksa di HP begitu dipilih. |
| Report | Admin | Dashboard (akurasi, ranking, analisis kesalahan), Productivity (per orang dan per jam), Analytics (akar masalah). |
| Config | pemegang akses | User, pembagian tugas, akses, facility dan lokasi aktif. Dibuka dari inisial nama di pojok kiri atas Home. |

Hal yang sengaja dibuat begini:

- **Tidak ada pengaturan server di aplikasi.** Pengguna hanya mengisi NIK. Alamat Web App ditanam saat build dan diperbarui dari `app/server.json` di repo ini (lihat Langkah 2).
- **Hitungan buta.** Qty sistem tidak pernah tampil saat menghitung. Hasil cocok atau selisih baru terlihat di riwayat sesi setelah hasil terkirim dan tidak bisa diurungkan.
- **Tahan sinyal putus.** Hasil hitung disimpan dulu di HP dan dikirim di latar belakang. Kalau jaringan hilang di antara rak, hasil menunggu di antrean HP dan terkirim sendiri saat tersambung lagi, juga setelah aplikasi ditutup.
- **Angka yang jujur.** Progres tidak ditulis 100% selama masih ada sisa.
- **Mode demo.** Selama server belum dipasang, layar masuk menawarkan mode demo: backend yang sama dijalankan di dalam aplikasi dengan data contoh (semua nama dan angka fiktif).

## Langkah 1 — Pasang backend di spreadsheet (sekali saja)

Satu spreadsheet Google menampung semuanya: user, facility, lokasi aktif, tugas hitung, riwayat, antrean, ringkasan, dan log.

1. Buka spreadsheet-nya di komputer, lalu **Extensions › Apps Script**.
2. Hapus isi `Code.gs`, tempel **seluruh** isi [`backend/Code.gs`](backend/Code.gs) (buka tombol *Raw*, pilih semua, salin), lalu simpan.
3. Isi tiga konstanta `SETUP_FACILITY_NAMA_`, `SETUP_FACILITY_KODE_`, `SETUP_ADMIN_NIK_` (cari `ISI.NIK.ADMIN` dengan Ctrl+F). NIK sengaja tidak disimpan di repo publik ini.
4. Pilih fungsi **`setupAwal`** di bilah atas, klik **Run**, dan izinkan akses. Fungsi ini membuat semua sheet beserta header-nya, facility pertama, admin pertama, dan trigger `workerSemua` yang memproses antrean tiap 1 menit. Aman dijalankan ulang.
5. **Deploy › New deployment › Web app.** Execute as: **Me**. Who has access: **Anyone**. Salin URL yang berakhiran `/exec`.

Setelah itu semuanya dari aplikasi: impor **lokasi aktif** lewat Config › Facility › Lokasi aktif (lokasi di luar daftar ini tidak dijadikan tugas), lalu tambah user lewat Config › User. Agar selisih bisa divalidasi oleh orang kedua, minimal harus ada satu user berperan **Inventory**.

Setiap kali `backend/Code.gs` berubah: tempel ulang, lalu **Deploy › Manage deployments › Edit › Version: New version** (URL tetap sama). Aplikasi tetap berjalan dengan backend versi lama, hanya saja laporan yang butuh fungsi baru (misalnya Productivity) menampilkan "Belum tersedia" sampai backend diperbarui.

**Catatan kuota:** trigger tiap menit memakai jatah waktu trigger Apps Script (90 menit/hari untuk akun Gmail biasa, 6 jam untuk Google Workspace). Karena itu semua antrean dijalankan oleh satu trigger.

## Langkah 2 — Sambungkan aplikasi ke server

Isi URL `/exec` dari Langkah 1 ke [`app/server.json`](app/server.json), lalu commit ke `main`:

```json
{ "url": "https://script.google.com/macros/s/XXXX/exec" }
```

- Rilis berikutnya membawa alamat itu di dalam APK.
- HP yang **sudah terpasang** mengambil berkas yang sama dari repo saat aplikasi dibuka, jadi ikut tersambung tanpa pasang ulang. Kalau suatu hari Web App di-deploy ulang dengan URL baru, cukup ubah berkas ini.
- Alternatif: variabel `GAS_URL` di **Settings › Secrets and variables › Actions › Variables**.

Karena repo dan APK-nya publik, URL ini bisa dibaca siapa pun. Keamanannya sama dengan web app Apps Script yang dibagikan lewat tautan: siapa pun yang tahu URL dan satu NIK terdaftar bisa masuk. Jangan menaruh data yang lebih sensitif dari itu.

## Cara pasang & update

1. Pertama kali: unduh `CycleTransaksi-vX.apk` dari halaman **Releases**, buka, izinkan "Instal aplikasi tidak dikenal".
2. Selanjutnya update dari dalam aplikasi: kartu **Versi X tersedia** muncul di Home (atau ketuk inisial nama di Home › Periksa pembaruan).
   - Perubahan tampilan/logika saja → **update kilat** (paket web kecil), langsung dipakai tanpa instal APK.
   - Perubahan native (plugin, izin, ikon) → APK diunduh di dalam aplikasi lalu dipasang menimpa versi lama. Tidak perlu uninstall.
3. Repo ini harus **publik** supaya aplikasi bisa membaca halaman Releases dan `app/server.json` tanpa login.

## Rilis ("gaspol")

Alurnya sama dengan IMM dan Active Coach:

1. Perubahan di-commit dan di-push ke `main`.
2. GitHub Actions (`.github/workflows/build-apk.yml`) menentukan versi, menguji backend, menyusun `www/` dari `src/`, memeriksa kode, membangun APK rilis, memastikan tanda tangannya sama dengan versi sebelumnya, lalu menerbitkannya di **Releases** bersama paket update kilat dan `latest.json`.
3. **Versi:** `package.json` berisi versi dasar. Perbaikan kecil otomatis naik PATCH (2.0.0 → 2.0.1 …). Untuk fitur besar, naikkan MINOR di `package.json`.
4. **Catatan rilis** diambil dari pesan commit. Pisahkan dengan baris `Untuk pengguna:` dan `Untuk developer:`.
5. Pesan commit berisi `[beta]` → rilis beta (prerelease), tidak menjadi versi terbaru. `[skip ci]` → tanpa rilis.
6. Rilis dibatalkan bila uji backend atau pemeriksaan kode gagal, APK ditandatangani kunci lain, atau APK masih bisa di-debug.

Aplikasi memeriksa versi lewat `releases/latest/download/latest.json`, bukan GitHub API, supaya puluhan HP di satu WiFi gudang tidak terkena batas 60 permintaan per jam.

## Mengubah tampilan

Tampilan ditulis tanpa framework dan tanpa CDN. `npm run build` menggabungkannya menjadi `www/`:

| Sumber | Hasil | Isi |
|---|---|---|
| `src/index.html` | `www/index.html` | Kerangka: satu `<section>` per layar dan navigasi bawah |
| `src/css/*.css` | `www/app.css` | `10-tokens` (warna, huruf, bentuk, tema gelap), `20-dasar` (komponen), `30-layar` (per layar) |
| `src/js/*.js` | `www/app.js` | Digabung berurutan nama dalam satu fungsi. `00-util`, `10-api`, `15-antrean`, `20-ui`, `25-grafik`, lalu satu file per layar |

Aturan singkatnya ada di [`docs/DESAIN.md`](docs/DESAIN.md). Yang paling sering dibutuhkan:

- Tombol memakai `data-aksi="nama"`; penangannya `Aksi['nama'] = (tombol) => …`. `npm run check` gagal bila ada tombol tanpa penangan.
- HTML dibuat dengan templat `` h`…` `` yang meng-escape semua nilai secara otomatis.
- Memanggil server: `await panggil('namaFungsi', arg1, arg2)`. Nama fungsi harus terdaftar di `API_BRIDGE_ALLOW_` (`backend/src/ApiBridge.gs`).

## Mengubah backend

Sumber backend ada di `backend/src/*.gs`. Setelah mengubahnya jalankan `npm run backend` untuk menyusun ulang `backend/Code.gs`, lalu tempel ke Apps Script dan deploy versi baru. Fungsi yang dibuat khusus untuk aplikasi (satu layar, satu panggilan) ada di `backend/src/AppApi.gs`.

## Uji lokal

`npm install && npm run build && npm test`

- `test/backend.mjs`: backend asli dijalankan di Node dengan tiruan layanan Google (`test/gas-mock.mjs`), dari `setupAwal`, unggah data, hitung, validasi, verifikasi dengan bukti WMS, penyelesaian plus minus, sampai Dashboard, Productivity, dan Analytics. Dijalankan dua kali: dari `backend/src` dan dari `backend/Code.gs`.
- `test/ui.mjs`: aplikasi dijalankan di Chromium (butuh paket `playwright`) dan berbicara dengan backend asli yang sama. Mencakup layar masuk tanpa pengaturan server, mode demo, cycle sampai tercatat di sheet, urungkan (termasuk item terakhir), antrean saat offline dan saat server sibuk, hasil yang ditolak, validasi, verifikasi dengan bukti, upload dari file Excel dan .csv, report (termasuk tanggal yang diganti selagi memuat), config, akun yang dinonaktifkan atau diubah perannya, pembaruan, backend versi lama, dan muat-tidaknya mode fokus di layar 360×568. `HANYA=3b,7 node test/ui.mjs` menjalankan bagian tertentu saja.
- `npm run demo` menyusun `www/` yang selalu memakai data contoh, untuk melihat tampilan di browser tanpa server (`npx serve www`).

## Isi repo

| Path | Isi |
|---|---|
| `src/` | Tampilan aplikasi (HTML, CSS, JS) |
| `app/server.json` | Alamat Web App Apps Script yang dipakai aplikasi |
| `demo/seed.mjs` | Data contoh untuk mode demo dan uji tampilan (fiktif) |
| `backend/Code.gs` | Backend Apps Script satu file, siap tempel |
| `backend/src/` | Sumber backend per modul, termasuk `Setup.gs` (setup awal), `ApiBridge.gs` (pintu masuk aplikasi), `AppApi.gs` |
| `scripts/` | Build web, bundel demo, pemeriksaan, versi, catatan rilis, update kilat, pembuat ikon |
| `android/` | Proyek Android (Capacitor). Kunci tanda tangan `android/app/cycle-transaksi.keystore` jangan dihapus atau diganti |
| `test/` | Uji backend dan uji tampilan ujung-ke-ujung |
| `docs/DESAIN.md` | Keputusan desain: warna, huruf, prinsip |
