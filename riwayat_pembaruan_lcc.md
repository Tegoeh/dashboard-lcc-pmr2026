# 🏆 LAPORAN RIWAYAT PEMBARUAN & SINKRONISASI SISTEM LCC PMR 2026
*Dokumen ini mencatat seluruh riwayat pembaruan, solusi bug, dan panduan operasional LCC PMR Tingkat MADYA (SMP) & WIRA (SMA).*

---

## 🔍 1. Riwayat Perbaikan Bug & Integrasi Sistem

Berikut adalah ringkasan masalah teknis yang berhasil diselesaikan untuk memastikan tabulasi berjalan lancar:

### A. Pergeseran Indeks Skor di Live Dashboard Online
* **Masalah:** Website Live Dashboard salah membaca nomor undian peserta penyisihan sebagai total skor, serta salah menampilkan peringkat penyisihan sebagai lencana bintang.
* **Penyebab:** Pada sheet `PMR MADYA` dan `PMR WIRA`, terdapat kolom fisik tambahan yaitu **Kelompok** (Kolom D). Hal ini menggeser kolom **TOTAL SKOR** ke Kolom E (indeks 4) dan **Peringkat** ke Kolom F (indeks 5). Website lama masih menggunakan parser statis tingkat MULA (TOTAL SKOR di kolom D / indeks 3).
* **Solusi:** Kami memperbarui fungsi `parseAllRoundsFromTable` di berkas `index.html` dan `dashboard.html` agar mendeteksi indeks kolom secara dinamis berdasarkan kategori tingkat (`cat`):
  * **PMR MULA:** TOTAL SKOR = Kolom D (indeks 3), Peringkat = Kolom E (indeks 4).
  * **PMR MADYA & WIRA:** TOTAL SKOR = Kolom E (indeks 4), Peringkat = Kolom F (indeks 5).
* **Status:** Sukses di-push ke GitHub Pages menggunakan token baru. Live website sudah 100% normal.

### B. Sinkronisasi Urutan Nomor Undi di Panggung Semifinal
* **Masalah:** Tim Lolos Semifinal seperti SMPN 1 Negara (dapat nomor undi 1) tidak langsung menempati posisi nomor urut 1 di tabel rekapitulasi.
* **Penyebab:** 
  1. Operator memasukkan nomor undian secara tidak sengaja di **Kolom D (`Kelompok`)**, padahal sistem otomatis panggung semifinal di sheet `LIVE BEL MADYA` mencari nomor undian di **Kolom H (`No. Undi SF`)** yang saat itu kosong.
  2. Rumus tabel rekapitulasi semifinal `PMR MADYA` baris 28-37 bawaan sebelumnya diurutkan berdasarkan peringkat penyisihan, bukan nomor undian.
* **Solusi:**
  1. Memindahkan seluruh nomor undian secara otomatis dari Kolom D ke Kolom H menggunakan skrip bantu Python.
  2. Memperbarui logika rumus pada sel `B28:B37` di sheet `PMR MADYA` (dan `B24:B33` di sheet `PMR WIRA`) agar diurutkan secara penuh berdasarkan **No. Undi SF** (Kolom H).
* **Status:** Sukses. SMPN 1 Negara kini sudah berada di urutan No. 1 pada tabel klasemen semifinal dan langsung terpasang di Jalur Bel 1 Kelompok 1 pada sheet `LIVE BEL MADYA`.

### C. Penambahan Dropdown Opsi `0` pada Babak Final
* **Masalah:** Kebutuhan input jawaban netral/belum mulai agar tidak kosong atau error.
* **Solusi:** User telah menambahkan manual opsi dropdown `0` di sheet `LIVE BEL MADYA` Babak Final (E28:X32). Kami memverifikasi bahwa formula total skor final (`=IF(E28=1, 50, IF(E28=-1, -25, 0)) + ...`) sangat kompatibel dengan angka `0` maupun sel kosong, dan akan dihitung sebagai `0` poin (netral) secara aman.

### D. Duplikasi Nama Finalis saat Skor 0 (Babak Final)
* **Masalah:** Ketika seluruh skor final masih bernilai `0` (sebelum lomba dimulai), tabel final di Live Dashboard hanya menampilkan satu nama sekolah (Juara 1) sebanyak 5 kali. Ketika ada minimal 1 poin masuk, kelima finalis baru tampil dengan benar.
* **Penyebab:** Pada skrip inisialisasi formula `otomasi_tabulasi.js`, rumus cell fallback saat total skor panggung final bernilai `0` (`SUM(LIVE BEL ... = 0`) ditulis secara hardcoded mengarah ke `LIVE BEL MADYA!C28` untuk kelima baris finalis. Di sisi website, parser DOM me-reuse elemen card dengan `data-kode` yang sama sehingga hanya menampilkan satu baris di layar.
* **Solusi:** Kami memperbaiki logika penulisan formula di `otomasi_tabulasi.js` dengan membuat referensi fallback baris bersifat dinamis:
  * **PMR MULA:** Referensi fallback menggunakan `LIVE BEL MULA!D' + (r - 9)` (untuk baris 28-32).
  * **PMR MADYA:** Referensi fallback menggunakan `LIVE BEL MADYA!C' + (r - 13)` (untuk baris 28-32).
  * **PMR WIRA:** Referensi fallback menggunakan `LIVE BEL WIRA!C' + (r - 9)` (untuk baris 28-32).
  * Kami juga telah menjalankan skrip perbaikan langsung untuk memperbarui sel-sel formula final di Google Spreadsheet aktif (`PMR MULA B37:B41`, `PMR MADYA B41:B45`, `PMR WIRA B37:B41`) sehingga saat ini Google Sheets Anda sudah 100% normal dan siap digunakan.
* **Status:** Sukses diperbaiki di Google Sheets aktif dan kode repositori lokal.

---

## 🛠️ 2. Salinan Kode Apps Script Terbaru (`otomasi_tabulasi.js`)
Seluruh perbaikan rumus di atas telah disimpan secara permanen pada berkas Apps Script Proyek Anda di GitHub. Berikut adalah logika inisialisasi semifinal yang sudah diperbarui:

### Cuplikan Logika PMR MADYA (`initMadya`)
```javascript
  // 2. Babak Semifinal (Rows 28-37) - Diurutkan berdasarkan No. Undi SF (Kolom H)
  for (var r = 28; r <= 37; r++) {
    sheetPmr.getRange("B" + r).setFormula('=IFERROR(INDEX(SORT(FILTER($B$6:$B$23, $G$6:$G$23 = "LOLOS SEMIFINAL"), FILTER($H$6:$H$23, $G$6:$G$23 = "LOLOS SEMIFINAL"), TRUE), A' + r + '), "")');
    sheetPmr.getRange("C" + r).setFormula('=IFERROR(VLOOKUP(B' + r + ', $B$6:$C$23, 2, FALSE), "")');
    sheetPmr.getRange("D" + r).setFormula('=IF(COUNTIF(\'LIVE BEL MADYA\'!$C$8:$C$12, B' + r + ')>0, 1, IF(COUNTIF(\'LIVE BEL MADYA\'!$C$18:$C$22, B' + r + ')>0, 2, ""))');
    sheetPmr.getRange("E" + r).setFormula('=IF(B' + r + '="", "", SUM(IFERROR(FILTER(\'LIVE BEL MADYA\'!$U$8:$U$12, \'LIVE BEL MADYA\'!$C$8:$C$12 = B' + r + '), 0)) + SUM(IFERROR(FILTER(\'LIVE BEL MADYA\'!$U$18:$U$22, \'LIVE BEL MADYA\'!$C$18:$C$22 = B' + r + '), 0)))');
    sheetPmr.getRange("F" + r).setFormula('=IF(ISNUMBER(E' + r + '), RANK(E' + r + ', $E$28:$E$37), "")');
    sheetPmr.getRange("G" + r).setFormula('=IF(F' + r + '="", "", IF(F' + r + '<=5, "LOLOS FINAL", ""))');
  }
```

### Cuplikan Logika PMR WIRA (`initWira`)
```javascript
  // 2. Babak Semifinal (Rows 24-33) - Diurutkan berdasarkan No. Undi SF (Kolom H)
  for (var r = 24; r <= 33; r++) {
    sheetPmr.getRange("B" + r).setFormula('=IFERROR(INDEX(SORT(FILTER($B$6:$B$19, $G$6:$G$19 = "LOLOS SEMIFINAL"), FILTER($H$6:$H$19, $G$6:$G$19 = "LOLOS SEMIFINAL"), TRUE), A' + r + '), "")');
    sheetPmr.getRange("C" + r).setFormula('=IFERROR(VLOOKUP(B' + r + ', $B$6:$C$19, 2, FALSE), "")');
    sheetPmr.getRange("D" + r).setFormula('=IF(COUNTIF(\'LIVE BEL WIRA\'!$C$8:$C$12, B' + r + ')>0, 1, IF(COUNTIF(\'LIVE BEL WIRA\'!$C$18:$C$22, B' + r + ')>0, 2, ""))');
    sheetPmr.getRange("E" + r).setFormula('=IF(B' + r + '="", "", SUM(IFERROR(FILTER(\'LIVE BEL WIRA\'!$U$8:$U$12, \'LIVE BEL WIRA\'!$C$8:$C$12 = B' + r + '), 0)) + SUM(IFERROR(FILTER(\'LIVE BEL WIRA\'!$U$18:$U$22, \'LIVE BEL WIRA\'!$C$18:$C$22 = B' + r + '), 0)))');
    sheetPmr.getRange("F" + r).setFormula('=IF(ISNUMBER(E' + r + '), RANK(E' + r + ', $E$24:$E$33), "")');
    sheetPmr.getRange("G" + r).setFormula('=IF(F' + r + '="", "", IF(F' + r + '<=5, "LOLOS FINAL", ""))');
  }
```

---

## 📋 3. Panduan Operasional Lomba Esok Hari

Untuk memastikan jalannya lomba PMR MADYA & WIRA besok bebas dari hambatan, mohon ikuti daftar periksa (checklist) berikut sebelum acara dimulai:

1. **Hard Refresh Browser Operator LCC:**
   * Sebelum lomba dimulai, buka Live Dashboard online Anda di browser operator dan tekan tombol **`Ctrl + F5`** (atau **`Cmd + Shift + R`** pada Mac) beberapa kali untuk memastikan browser membuang cache lama dan memuat parser kolom dinamis yang baru dari GitHub Pages.
2. **Pengisian Nomor Undian Semifinal:**
   * Pengisian nomor undian 10 besar babak penyisihan **wajib** diisi langsung secara manual di spreadsheet utama pada **Kolom H (`No. Undi SF`)** sheet `PMR MADYA`/`PMR WIRA`. Jangan memasukkan nomor undian dari website.
3. **Sterilisasi Skor (Reset 0):**
   * Pastikan semua kolom skor rebutan semifinal (`F8:T12`, `F18:T22`) dan jawaban final (`E28:X32`) di sheet `LIVE BEL MADYA` sudah kosong atau bernilai `0` agar Live Dashboard dimulai dari skor awal 0 secara steril.
