# Cycle Transaksi — Aplikasi Android

Aplikasi Android (APK) untuk cycle count berbasis transaksi. **Database tetap Google Sheets** dan seluruh logika (pembagian tugas, antrean, worker, dashboard) tetap berjalan di **Google Apps Script**. APK ini hanya memindahkan tampilan ke HP dan memanggil Web App yang sama lewat `doPost`.

```
HP Android (APK)                          Google
┌──────────────────────────┐   HTTPS    ┌──────────────────────────────────────┐
│ Tampilan (src/*.html)    │ ─────────► │ Apps Script Web App (doPost)         │
│ Jembatan (app/bridge.js) │            │   ApiBridge.gs → fungsi .gs yang ada │
│ pengganti                │ ◄───────── │ Google Sheets: Master DB, Config DB, │
│ google.script.run        │            │   spreadsheet per facility           │
└──────────────────────────┘            └──────────────────────────────────────┘
```

## Langkah 1 — Pasang pintu masuk di Apps Script (sekali saja)

1. Buka project Apps Script Cycle Transaksi. Tambah file skrip baru bernama `ApiBridge`, tempel isi `backend/ApiBridge.gs`.
2. Di `Main.gs`, di dalam `doPost(e)`, tepat setelah baris `const req = JSON.parse(e.postData.contents || '{}');` tambahkan:
   ```js
   const viaApk = apiBridge_(req);
   if (viaApk) return viaApk;
   ```
3. **Deploy › Manage deployments › Edit › Version: New version.** Execute as: **Me**. Who has access: **Anyone**. Salin URL yang berakhiran `/exec`.

Tidak ada file `.gs` lain yang perlu diubah. Web app lama (lewat browser) tetap berjalan seperti biasa.

## Langkah 2 — Alamat server

Pilih salah satu:

- **Diisi di HP:** saat pertama dibuka, ketuk **Atur server** di layar masuk, tempel URL `/exec`, lalu **Simpan & uji**. Cukup sekali per HP.
- **Ditanam di APK:** di GitHub buka **Settings › Secrets and variables › Actions › tab Variables**, tambah `GAS_URL` berisi URL `/exec`, lalu jalankan ulang workflow. Karena repo dan APK-nya publik, URL ini bisa dibaca siapa pun yang mengunduh APK.

## Cara pasang & update

1. Pertama kali: unduh `CycleTransaksi-vX.apk` dari halaman **Releases**, buka, izinkan "Instal aplikasi tidak dikenal".
2. Selanjutnya update dari dalam aplikasi: kartu **Versi X tersedia** muncul di Home (atau ketuk **Periksa pembaruan** di bagian bawah Home).
   - Perubahan tampilan/logika saja → **update kilat** (±0,4 MB), langsung dipakai tanpa instal APK.
   - Perubahan native (plugin, izin, ikon) → APK diunduh di dalam aplikasi lalu dipasang menimpa versi lama. Tidak perlu uninstall.
3. Repo ini harus **publik** supaya aplikasi bisa membaca halaman Releases tanpa login.

## Yang berbeda dari web app

- **Tetap masuk.** NIK diingat di HP; aplikasi langsung membuka Home sampai pengguna menekan *Keluar / Ganti User*.
- **Tombol kembali Android** membawa ke Home; di Home atau layar masuk, aplikasi diperkecil.
- **Ekspor Excel** dibuka dengan aplikasi spreadsheet di HP (WebView Android tidak bisa mengunduh langsung).
- **SheetJS dibundel di APK**, tidak lagi diambil dari CDN.
- **Pita "Tidak ada koneksi internet"** muncul saat HP offline. Aplikasi butuh internet untuk semua data.

Versi di titlebar (mis. v8.29.3) adalah versi **backend** Apps Script. Versi **aplikasi** ada di bagian bawah Home.

## Rilis ("gaspol")

Alurnya sama dengan IMM dan Active Coach:

1. Perubahan di-commit dan di-push ke `main`.
2. GitHub Actions (`.github/workflows/build-apk.yml`) menentukan versi, menyusun `www/` dari `src/`, memeriksa kode, membangun APK rilis, memastikan tanda tangannya sama dengan versi sebelumnya, lalu menerbitkannya di **Releases** bersama paket update kilat dan `latest.json`.
3. **Versi:** `package.json` berisi versi dasar. Perbaikan kecil otomatis naik PATCH (1.0.0 → 1.0.1 …). Untuk fitur besar, naikkan MINOR di `package.json`.
4. **Catatan rilis** diambil dari pesan commit. Pisahkan dengan baris `Untuk pengguna:` dan `Untuk developer:`.
5. Pesan commit berisi `[beta]` → rilis beta (prerelease), tidak menjadi versi terbaru.
6. Rilis dibatalkan bila pemeriksaan kode gagal, APK ditandatangani kunci lain, atau APK masih bisa di-debug.

Aplikasi memeriksa versi lewat `releases/latest/download/latest.json`, bukan GitHub API, supaya puluhan HP di satu WiFi gudang tidak terkena batas 60 permintaan per jam.

## Mengubah tampilan

File di `src/` adalah file HTML yang sama dengan project Apps Script. Kalau tampilan diubah di Apps Script, salin file `.html` yang berubah ke `src/`, lalu commit. Kalau menambah fungsi server baru yang dipanggil `google.script.run`, tambahkan namanya ke `API_BRIDGE_ALLOW_` di `ApiBridge.gs` (dan perbarui di Apps Script).

Uji lokal: `npm install && npm run build && npm test` (memakai backend tiruan di Chromium).

## Isi repo

| Path | Isi |
|---|---|
| `src/` | Tampilan asli Apps Script (`Index.html`, `Body*.html`, `Js*.html`, `Style.html`) |
| `app/bridge.js`, `app/bridge.css` | Jembatan pengganti `google.script.run`, alamat server, sesi, update |
| `backend/ApiBridge.gs` | File yang ditambahkan ke project Apps Script |
| `scripts/` | Build web, pemeriksaan, versi, catatan rilis, update kilat, pembuat ikon |
| `android/` | Proyek Android (Capacitor). Kunci tanda tangan `android/app/cycle-transaksi.keystore` jangan dihapus atau diganti |
| `test/run.mjs` | Uji jembatan dari layar masuk sampai submit hitung |
