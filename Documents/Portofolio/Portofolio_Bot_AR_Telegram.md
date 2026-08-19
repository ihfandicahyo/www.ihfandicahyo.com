# Portofolio Proyek: Bot AR Telegram — Chatbot Piutang Real-Time

> **Ringkasan Portofolio Sistem Informasi Piutang Real-Time Berbasis Chatbot Telegram**

---

## 📌 1. Problem (Masalah)
* **Keterbatasan Akses Data:** Tim sales, operasional di lapangan dan beberapa tim lain yang membutuhkan data AR tidak dapat mengakses langsung laporan AR. Mereka harus bertanya melalui admin sebagai jembatan dan kemudian admin akan memproses penarikan data dari sistem atau jika dalam proyek generasi sebelumnya yang saya buat maka dapat membuka file Excel `ARVIEWER.xlsm` secara manual hanya untuk mengecek sisa piutang atau status pembayaran pelanggan.
* **Proses Lambat & Inefisien:** Pencarian manual memakan waktu, terutama saat berada di luar kantor tanpa akses komputer desktop.
* **Ambiguitas & Inkonsistensi Data:** Penulisan nama pelanggan, cabang, atau kode pelanggan sering kali tidak seragam, menyulitkan pencarian data yang akurat secara cepat.

---

## 👥 2. User (Pengguna Sasaran)
* **Tim Sales, Penjualan Lapangan dan tim terkait:** Memerlukan pengecekan sisa piutang pelanggan sebelum melakukan penagihan atau transaksi baru.
* **Tim Finance, Accounting & Admin AR:** Memantau keterlambatan jatuh tempo, histori pelunasan, dan verifikasi status pembayaran faktur secara *real-time*.
* **Manajemen Perusahaan:** Memperoleh visibilitas cepat atas kondisi piutang usaha (*Accounts Receivable*) dari mana saja.

---

## 💡 3. Solution (Solusi)
* **Bot Telegram Interaktif Real-Time:** Menghubungkan data piutang dari Excel (`ARVIEWER.xlsm`) langsung ke aplikasi Telegram yang dapat diakses dari smartphone maupun PC.
* **Pencarian Cepat Berbasis RAM (In-Memory Preloading):** Memuat seluruh data ke memori RAM saat startup dengan pembaruan otomatis di latar belakang (*background thread refresh*), menghasilkan respons query dalam hitungan detik.
* **Filter Interaktif 3-Langkah:** Menyediakan mekanisme filter bertahap via tombol inline Telegram (Produk/Depo → Jatuh Tempo → Data Fraud) untuk hasil laporan yang presisi.

---

## ✨ 5. Key Features (Fitur Utama)
* **Autentikasi Keamanan Sesi:** Menggunakan `secret_key` internal untuk mengunci akses bot dari pihak yang tidak berwenang.
* **Pencarian Fleksibel (Multi-Query & Multi-Kode):** Mendukung pencarian berdasar kode pelanggan, nama pelanggan/kontak, kata kunci grup, hingga penggabungan multi-kode dengan operator `&` (misal: `YY-2223 & MGL-1045`).
* **Resolusi Nama Pelanggan 6-Lapis:** Algoritma bertingkat menggunakan *branch rules*, kamus resolusi, hingga *fuzzy matching* (`rapidfuzz`) untuk menangani variasi penulisan nama cabang/pelanggan.
* **Informasi Faktur Terbayar BG:** Otomatis memberikan keterangan Jatuh Tempo BG jika faktur telah dibayar dengan metode BG.
* **Visualisasi Laporan Lunas & Piutang (PNG):** Otomatis meregenerasi tabel data piutang menjadi gambar PNG siap kirim dengan penanda warna status pembayaran per faktur (LUNAS, DICICIL, LEBIH BAYAR).
* **Toleransi Ukuran Gambar & Fallback:** Penyesuaian DPI otomatis untuk data berukuran besar dan fitur pengiriman *fallback* dari foto ke dokumen PNG jika melebihi limit Telegram.

---

## ⚠️ 5. Challenge (Tantangan & Penanganan)
* **Tantangan 1 — Inkonsistensi Nama Cabang & Pelanggan:**
  * *Penanganan:* Menerapkan *Branch Rules* khusus dan pencocokan bertingkat (Exact → ML (Machine Learning) Dictionary → Fuzzy Matching WRatio/Token Sort) dengan batas ambang (*threshold*) hingga 80%.
* **Tantangan 2 — Batasi Limit Piksel & Format Output Telegram:**
  * *Penanganan:* Mengatur batasan piksel maksimum (9.000 piksel/sisi), dinamisasi ukuran canvas Matplotlib (0.32 inci/baris), serta konversi otomatis dari pengiriman foto ke dokumen jika file terlalu besar.
* **Tantangan 3 — Keamanan Thread & Konsistensi Data (Concurrency):**
  * *Penanganan:* Penggunaan *thread lock* (`data_lock` dan `session_lock`) untuk mencegah *race condition* saat background thread memperbarui data RAM secara periodik sembari bot melayani query pengguna.

---

## 📈 6. Impact (Dampak & Hasil)
* **Efisiensi Waktu Signifikan:** Memangkas waktu pengecekan piutang dari beberapa menit menjadi **kurang dari 5 detik**.
* **Peningkatan Aksesibilitas Lapangan:** Tim sales dapat mengecek sisa piutang dan histori pembayaran secara langsung di lokasi pelanggan sebelum melakukan penagihan atau persetujuan order.
* **Reduksi Kesalahan Informasi:** Mengurangi kesalahan manusia (*human error*) dalam membaca status faktur berkat indikator warna visual (Biru = LUNAS, Oranye = DICICIL, Magenta = LEBIH BAYAR).

---

## 🛠️ 7. Tech Choices (Teknologi yang Digunakan)
* **Bahasa Pemrograman:** Python 3.8+
* **Telegram Bot Framework:** `pyTelegramBotAPI` (`telebot`)
* **Pengolahan & Transformasi Data:** `pandas`, `openpyxl`, `numpy`
* **Pencocokan Teks & Fuzzy Logic:** `rapidfuzz` (`token_sort_ratio`, `WRatio`)
* **Visualisasi & Rendering Gambar:** `matplotlib` (Backend `Agg` untuk rendering tabel ke PNG secara *headless*)
* **Manajemen Multithreading & Konfigurasi:** Standard Python Libraries (`threading`, `configparser`, `shutil`, `re`, `io`)

---

## 🖼️ 8. Screenshot & Workflows (Visualisasi & Alur Sistem)

### A. Alur Kerja Pengguna di Telegram
1. **Autentikasi:** `/start` ➔ Masukkan `secret_key` ➔ Akses Diterima.
2. **Pencarian Query:** Ketik `Wakid Kendal` atau `YY-2223 & MGL-1045`.
3. **Filter Interaktif (Inline Buttons):**
   * *Langkah 1:* Pilih Produk (`IRC` | `ZN` | `SEMUA`)
   * *Langkah 2:* Pilih Jatuh Tempo (`HANYA JT` | `SEMUA DATA`)
   * *Langkah 3:* Filter Fraud (`TANPA FRAUD` | `SERTAKAN FRAUD`)
4. **Hasil Laporan:** Bot mengirimkan Gambar Tabel PNG lengkap dengan header PT PRIMA TUNGGAL MANDIRI, warna status pembayaran per faktur, serta ringkasan total piutang di bagian footer.

### B. Diagram Arsitektur Data Pipeline
```text
[ARVIEWER.xlsm] + [File ML / Master Data]
       │
       ▼
[1_CopyData.py] ───► Ekstraksi Data ke Temp Files (.xlsx)
       │
       ▼
[2_AdjDateFormat.py] ───► Normalisasi Format Tanggal Indonesia
       │
       ▼
[3_ARBotTelegram.py] ───► Preload Data ke RAM (Global Memory)
       │                  ├── Background Thread Auto-Refresh (N Menit)
       │                  └── Lock Mechanism for Thread Safety
       ▼
[Telegram Client] ◄───► Filter Interaktif & Render Image (Matplotlib PNG)
```
