# [🚀 Portofolio Case Study: Ambil AR — Lembar Tagihan & Sinkronisasi Google Sheets](https://github.com/ACC-TAX-REIGHTEEN/Lembar-Penagihan-Harian)

> **Otomasi Pipeline Python untuk Efisiensi Pemrosesan Piutang Accurate ke Lembar Tagihan Cetak & Rekap Cloud**

---

## 📌 Problem (Masalah)
* **Proses Manual Memakan Waktu**: Admin keuangan/piutang membutuhkan waktu sekitar **3 hingga 10 menit** setiap kali melakukan rekonsiliasi manual piutang usaha dari ekspor Accurate.
* **Risiko Human Error**: Proses pemisahan data per penagih/sales, perhitungan ulang umur piutang, penulisan ulang data ke spreadsheet rekap, dan cetak lembar tagihan secara manual sangat rentan terhadap kesalahan input, hilangnya formula, atau salah alokasi kode pelanggan.
* **Format Data Legacy yang Tidak Konsisten**: File ekspor dari Accurate (`ExportFile.xls`) memiliki posisi header yang bervariasi dan format angka yang tidak seragam (kombinasi pemisah titik/koma lokal & internasional).

---

## 👤 User (Pengguna Target)
1. **Admin Piutang & Finance**: Membutuhkan alat otomasi yang cepat, akurat, dan sekali klik tanpa perlu *copy-paste* data manual.
2. **Sales & Collector (Penagih Lapangan)**: Memerlukan lembar tagihan fisik (`Print_AR.xlsm`) yang rapi, lengkap dengan area tanda tangan dan total tagihan per penagih.
3. **Manajemen & Tim Eksekutif**: Memerlukan akses *real-time* ke rekap digital piutang di Google Sheets untuk monitoring status tagihan dan histori pembayaran.

---

## 💡 Solution (Solusi)
Pengembangan pipeline otomasi Python 5-langkah bernama **`Ambil AR`** yang mengintegrasikan pembersihan data, penyusunan template Excel berformat resmi, dan sinkronisasi cloud API.

Pipeline ini memangkas durasi pengerjaan dari **3–10 menit menjadi hitungan detik**, di mana pengguna hanya perlu menaruh file `ExportFile.xls` dan menjalankan satu skrip utama (`Ambil AR.py`).

---

## ✨ Key Features (Fitur Utama)
- **Pemetaan Penagih Dinamis**: Mengidentifikasi dan mengelompokkan kode pelanggan ke nama sales/penagih secara otomatis via berkas konfigurasi `piutang.conf`.
- **Kalkulasi Otomatis Umur JT & Terbayar**:
  - `Umur JT` dihitung ulang secara akurat berdasarkan selisih hari dari tanggal faktur ke hari ini.
  - `Terbayar` dihitung dari `Nilai Faktur − Sisa Piutang` (otomatis disembunyikan jika bernilai nol).
- **Preservasi Template Macro (`.xlsm`)**: Memanfaatkan `xlwings` untuk menyalin header, rumus sum `=SUM()`, footer tanda tangan, serta gambar/logo dari `TEMPLATE.xlsm`.
- **Helper Cleaning Sebelum Inject**: Membuka merge cell, memasukkan identitas penagih ke tiap baris data, dan membersihkan elemen non-data menggunakan `openpyxl`.
- **Sinkronisasi Google Sheets API**: Menyisipkan data baru pada posisi sebelum baris terakhir di Google Sheets dengan mempertahankan gaya format bawaan (`inherit_from_before=True`).
- **Validasi Integrity & Auto-Cleanup**: Verifikasi folder/file dependensi di awal serta pembersihan file sementara (`*temp.xlsx`) setelah eksekusi selesai.

---

## 🛠️ Challenge (Tantangan Teknis)
1. **Parsing Dynamic Header `.xls`**: Mengatasi variasi posisi kolom pada file ekspor Accurate dengan memindai 150 baris pertama untuk mendeteksi indeks kolom target secara otomatis.
2. **Manipulasi Excel COM via Python**: Menjaga integritas visual, shape (logo/gambar), dan format macro Excel (`.xlsm`) tanpa merusak struktur VBA bawaan.
3. **Dynamic Row Insertion di Google Sheets**: Memastikan data disisipkan dengan tepat sebelum baris footer/total pada spreadsheet cloud tanpa menimpa baris bawah atau merusak struktur sheet.

---

## 📊 Impact (Dampak & Hasil)
- **Efisiensi Waktu Signifikan**: Memangkas waktu pemrosesan harian dari **3–10 menit manual** menjadi **otomatis dalam kurun < 10 detik** (hemat waktu > 95%).
- **Akurasi Data 100%**: Menghilangkan risiko *human error* pada pemetaan penagih, perhitungan umur piutang, dan transfer data.
- **Transparansi & Rekapitulasi Real-Time**: Seluruh tim (keuangan, sales, manajemen) secara konsisten memiliki data tagihan yang sinkron antara lembar cetak fisik dan rekapitulasi digital cloud.

---

## 🛠️ Tech Choices (Pilihan Teknologi)
- **Python 3.8+**: Runtime otomasi utama.
- **Pandas & xlrd**: Membaca file legacy `.xls`, pembersihan data, dan transformasi dataframe.
- **xlwings**: Memanggil Excel COM engine untuk manipulasi template `.xlsm` dan preservasi shape/logo.
- **openpyxl**: Manipulasi struktur XML Excel, unmerging cell, dan pembentukan dataset siap upload.
- **gspread & google-auth**: Interaksi API Google Sheets dengan autentikasi Service Account.
- **XlsxWriter**: Pembuatan berkas spreadsheet temporer berformat angka presisi.

---

## 🖼️ Screenshot & Alur Sistem (Visual Architecture)

```
 [ Input ]               [ Processing Pipeline ]                     [ Output ]
 ExportFile.xls ───► [ 1_CleanerAcc.py ]
                        │
                     [ 2_FilterAR.py ]
                        │
                     [ 3_CalculateAR.py ] ──────► Print_AR.xlsm (Lembar Tagihan Cetak)
                        │
                     [ 4_HelperCleaning.py ]
                        │
                     [ 5_InjectDataToSS.py ] ────► Google Sheets (Rekap Cloud Tim)
```

- **File Output Utama**: `Print_AR.xlsm` (Siap cetak per penagih dengan area TTD Sales/Collector)
- **Rekap Cloud**: Terintegrasi langsung ke tab Google Sheets yang ditentukan pada `piutang.conf`.
