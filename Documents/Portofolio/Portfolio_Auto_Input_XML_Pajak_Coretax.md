# [📁 Portofolio Kasus Studi: Auto-Input XML Pajak Coretax](https://github.com/ACC-TAX-REIGHTEEN/Auto-Input-XML-Pajak-Coretax)

> **Kasus Otomasi Sistem Perpajakan**  
> *Transformasi Proses Pengolahan Data Faktur Akuntansi ke Format XML DJP Coretax berbasis Python*

---

## 1. 🎯 Problem (Masalah)
* **Proses Manual Memakan Waktu:** Penginputan dan rekapitulasi data dari sistem akuntansi (Accurate / e-Faktur) ke template XML DJP Coretax memerlukan waktu sekitar **1 hingga 3 jam per faktur/batch** jika dilakukan secara manual.
* **Risiko Human Error Tinggi:** Pengetikan ulang NPWP, kalkulasi DPP, PPN 12%, dan PPnBM, serta konversi tanggal secara manual sangat rentan terhadap kesalahan, terutama saat volume transaksi melonjak pada akhir bulan.
* **Inkonsistensi Format Data:** Format ekspor dari sistem akuntansi kerap tidak seragam (perbedaan nama/posisi kolom, format tanggal `DD/MM/YY`, hingga inkonsistensi nomor ID pembeli).

---

## 2. 👤 User (Pengguna Sasaran)
* **Tax Officer / Staf Perpajakan:** Membutuhkan alat kerja yang cepat dan presisi untuk menyiapkan file impor XML Coretax tanpa perlu mengedit file Excel baris demi baris.
* **Finance & Accounting Department:** Tim keuangan yang menggunakan sistem akuntansi seperti Accurate dan e-Faktur yang butuh menjembatani data ekspor lokal dengan portal resmi DJP Coretax.

---

## 3. 💡 Solution (Solusi)
* **Auto-Input XML Pajak Coretax (v1.2.3a):** Sebuah skrip otomasi berbasis Python yang mengonversi file ekspor akuntansi (`.xls`) menjadi file XML Coretax yang valid secara otomatis.
* **Eksekusi One-Click:** Mengubah pekerjaan manual yang memakan waktu 1 hingga 3 jam menjadi proses serba otomatis dalam hitungan detik.
* **Dua Mode Pemrosesan:** Mendukung skenario **Faktur Biasa** dan **Faktur dengan Diskon** per-item itemized.

---

## 4. ✨ Key Features (Fitur Utama)
* **Smart Header Auto-Detection:** Otomatis mendeteksi tata letak kolom pada file sumber ekspor, sehingga toleran terhadap variasi atau pergeseran posisi kolom.
* **Pembersihan & Normalisasi Data Otomatis:**
  * Standardisasi format tanggal dari `DD/MM/YY` ke `DD/MM/YYYY`.
  * Normalisasi NPWP dan ekstraksi ID TKU secara dinamis.
  * Pemetaan otomatis Jenis ID Pembeli ke *National ID* (NIK KTP via `KTP.txt`) atau *Other ID* (via `KTP-OTH.txt`).
* **Auto-Lookup Kode Barang/Jasa:** Pencarian otomatis kode komoditas Coretax beserta pengelompokannya (Grup A untuk Barang / Grup B untuk Jasa).
* **Kalkulasi Pajak Presisi:** Menghitung otomatis **DPP**, **DPP Nilai Lain** (`DPP × 11/12`), **PPN 12%**, dan **PPnBM**.
* **Filter & Clearing Otomatis:** Menghapus item bernilai negatif (retur) dan item non-faktur sesuai daftar `Helper_Del.txt`.
* **Auto-Update Reference via GitHub:** Otomatis memperbarui file referensi (`Helper_*.txt`) dari repositori GitHub saat terhubung ke internet.

---

## 5. 🧩 Challenge (TantanganTeknis)
* **Penanganan Diskon Bertingkat (Itemized Discount):** Menghitung porsi diskon per-item secara proposional (`DISC.ITEM = (DISC.TANPA / TOTAL QTY) × Qty`) agar nilai DPP dan PPN pada detail item tepat sesuai nilai faktur header.
* **Interoperabilitas Template Excel Coretax:** Pengisian template `.xlsx` resmi Coretax tanpa merusak struktur sel, rumus built-in, atau formatting asli Coretax. Solusi dicapai dengan memanfaatkan interop COM via `xlwings`.
* **Arsitektur Pipeline Terisolasi:** Memastikan file sementara (*temporary files*) yang diolah di folder `Dapur` dibersihkan secara otomatis setelah proses selesai agar tidak mengotori lingkungan kerja user.

---

## 6. 🚀 Impact (Dampak & Hasil)
* **Pangkas Waktu Hingga 95%:** Proses rekap dan pemrosesan data faktur yang semula membutuhkan **1 hingga 3 jam per batch**, kini selesai hanya dalam waktu **< 30 detik**.
* **Zero Error Rate pada Kalkulasi:** Menghilangkan *human error* pada perhitungan PPN 12% dan DPP Nilai Lain.
* **Akurasi Impor Coretax 100%:** Memastikan format data yang masuk ke portal DJP Coretax selalu valid dan sesuai skema XML Coretax versi 1.3.25.

---

## 7. 🛠️ Tech Choices (Teknologi yang Digunakan)
* **Python 3:** Bahasa pemrograman utama untuk orkestrasi skrip dan automasi.
* **Pandas:** Engine utama untuk *data cleaning*, manipulasi dataframe, dan transformasi logika tabel.
* **xlwings:** Otomasi Microsoft Excel via COM interface untuk mentransfer data hasil olahan ke template Excel Coretax secara aman.
* **openpyxl & xlsxwriter:** Modul pembacaan dan pembuatan file Excel intermediate dengan performa tinggi.
* **Requests:** Memuat dan membandingkan checksum MD5 untuk pembaruan otomatis file Helper dari repositori GitHub.

---

## 8. 🖼️ Screenshot & Workflow Blueprint

### Alur Kerja Internal (Data Pipeline Architecture)

```
 [INPUT FILE]              [ENGINE / DAPUR PIPELINE]                       [OUTPUT]
┌───────────────┐         ┌──────────────────────────────────────┐     ┌───────────────────┐
│ AccCtxFaktur  ├────────►│ 1. Cleaner (Normalisasi & Tanggal)   │     │                   │
└───────────────┘         │ 2. Item Calc (DPP, PPN 12%, Diskon) │     │                   │
                          │ 3. Lookup Helper & Mapping ID        │────►│ XML-DD-MMM-YY.xlsx│
┌───────────────┐         │ 4. Transfer via xlwings ke Template  │     │ (Siap Coretax)    │
│ AccEFaktur    ├────────►│ 5. Auto Cleanup Temp Files           │     │                   │
└───────────────┘         └──────────────────────────────────────┘     └───────────────────┘
```

### Struktur Repositori & Modul
```
Auto-Input-XML-Pajak-Coretax/
├── Buat Laporan XML Biasa.py       ← Entry Point (Faktur Biasa)
├── Buat Laporan XML Diskon.py      ← Entry Point (Faktur Diskon)
├── AccCtxFaktur.xls                ← Input Header
├── AccEFaktur.xls                  ← Input Detail
└── Dapur/                          ← Engine & Helper Referensi
    ├── 0_Ftch_github.py
    ├── 1_AccCtxFaktur_cleaner.py
    ├── 2_AccEFaktur_cleaner.py
    ├── ...
    └── TEMPLATE_1.3.25.xlsx        ← Template Resmi Coretax
```
