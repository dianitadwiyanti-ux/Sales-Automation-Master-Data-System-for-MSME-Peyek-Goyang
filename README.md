# 📊 Sales & Order Automation Dashboard for UMKM Peyek Goyang

![Project Status](https://img.shields.io/badge/Project%20Status-Completed-brightgreen)
![Tools](https://img.shields.io/badge/Tools-Microsoft%20Excel%20%7C%20Advanced%20Formulas%20%7C%20Pivot%20Table-blue)

## 📌 Business Overview & Problem Statement
**UMKM Peyek Goyang** merupakan usaha mikro di kuliner makanan ringan. Sebelum proyek ini dijalankan, seluruh alur rekapitulasi dan pencatatan transaksi harian masih dilakukan secara manual menggunakan **buku catatan berbasis kertas**. 

Proses manual ini menghadapi beberapa kendala utama:
- **Tinggi Risiko *Human Error*:** Penjumlahan nominal harga dan kuantitas pesanan yang dihitung secara manual rentan salah hitung.
- **Inefisiensi Waktu:** Membutuhkan waktu hingga 2–3 jam per hari hanya untuk mengkalkulasi total omzet dan merapikan buku rekap harian.
- **Kesulitan Menarik *Insight*:** Pemilik usaha sulit memantau total pendapatan harian/bulanan, tren waktu penjualan, varian produk terlaris secara *real-time*.

---

## 📸 Transformation: Before vs After

| Sebelum (Buku Catatan Manual) | Sesudah (Automated Excel Database & Dashboard) |
| :---: | :---: |
| Pencatatan fisik di buku rekap kertas, rentan rusak, tidak terstruktur, dan berisiko salah hitung. | Sistem database relasional terstruktur di Excel, otomatisasi formula, serta interactive dashboard visual. |
| <img width="720" height="1280" alt="WhatsApp Image 2026-10-31 at 18 15 37" src="https://github.com/user-attachments/assets/411dfdb0-4d10-40ec-acfb-245786a0d90f" />
 | :---: |

*(Tambahkan gambar preview screenshot perbandingan sebelum vs sesudah di bagian ini)*

---

## 🛠️ Project Architecture & Key Features

Proyek ini mentransformasi catatan fisik tak terstruktur (300+ entri transaksi fisik) menjadi skema basis data relasional terstruktur di Microsoft Excel yang terdiri dari **3 Lembar Kerja (Sheets)** utama:

### 1. `Dashboard` (Executive Summary & Visual Analysis)
- **Interactive Pivot Tables & Pivot Charts:** Menyajikan ringkasan visual terkait Total Pendapatan (*Revenue*), Total Pesanan (*Orders*), Kuantitas Produk Terjual (*Pcs*), serta Rata-rata Pembelian.
- **Dynamic Slicers & Timeline:** Tombol filter interaktif berdasarkan varian kemasan dan rentang waktu (Pagi/Sore/Bulan) untuk memadatkan laporan secara instan.
- **Key Metrics Highlighted:**
  - Total Pendapatan Per Bulan.
  - Pendapatan Per Varian Kemasan (Kemasan 2, 4, 12, 24).
  - Tren Transaksi Harian dan Sesi Operasional.

### 2. `Data_Pesanan` (Raw Data Transactional)
- **Automated Order ID:** Penomoran ID pesanan otomatis (`TRX-0001`, `TRX-0002`, dst.) yang bangkit secara otomatis saat kolom pesanan diisi.
- **Dynamic Day Name Extraction:** Menguraikan tanggal menjadi nama hari secara otomatis menggunakan fungsi `=TEXT(Tanggal, "dddd")`.
- **Automated Customer ID & Classification:** Standardisasi otomatis kualifikasi pelanggan (Pelanggan Rutin vs *Walk-in*). Jika terdaftar di master data, ID otomatis ditarik (`P001`, `P002`); jika pelanggan baru/umum, otomatis terkualifikasi sebagai `P000`.
  ```excel
  =IF(ISNA(XLOOKUP(E2; Master_Pelanggan!$A$2:$A$10; Master_Pelanggan!$B$2:$B$10)); "P000"; XLOOKUP(E2; Master_Pelanggan!$A$2:$A$10; Master_Pelanggan!$B$2:$B$10))
  ```
- **Dynamic Pricing (XLOOKUP):** Mengeliminasi kesalahan input nominal harga manual secara otomatis berdasarkan *Master Pricing*.
- **Data Validation & Flexible Dropdown:** Memastikan konsistensi penulisan nama varian dan waktu operasional tanpa membatasi fleksibilitas input nama pelanggan baru.

### 3. `File_Master` (Reference Tables)
- Menyimpan data referensi master varian kemasan beserta harga satuan, serta daftar *repeat customer* terdaftar.

---

## 📈 Key Insights

1. **Efisiensi Rekapitulasi Operasional:** Otomatisasi kalkulasi memangkas waktu operasional rekap harian secara signifikan dari hitungan jam ke hitungan menit.
2. **Pola Konsentrasi Operasional:** Teridentifikasi konsentrasi puncak pemesanan pada sesi waktu dan hari-hari tertentu yang berguna untuk optimalisasi persediaan stok bahan baku peyek.
3. **Segmentasi Produk & Pelanggan:** Varian kemasan tertentu (kategori 12 & 24) menjadi kontributor pendapatan terbesar, sekaligus memetakan profil *repeat customers* untuk efisiensi strategi promosi berulang.

---

## 🎯 Business Outcome & Impact

- ⏱️ **Memangkas Waktu Rekapitulasi:** Dari **2 jam menjadi kurang dari 10 menit per hari**.
- 🎯 **Akurasi Data 100%:** Mengeliminasi potensi *human error* pada penjumlahan nominal dan kalkulasi harga manual.
- 📊 **Pengambilan Keputusan Berbasis Data:** Menyajikan data penjualan secara *real-time* yang memudahkan pemilik usaha dalam menentukan tingkat produksi dan inventaris stok.

---

## 🧰 Tools & Technologies Used
- **Software:** Microsoft Excel
- **Features & Methods:** Advanced Formulas (`XLOOKUP`, `IF`, `ISNA`, `TEXT`, `ROW`), Data Validation Dropdowns, Pivot Table, Pivot Chart, Dynamic Slicers, Conditional Formatting, Database Normalization.

---

## 📂 File Structure
```text
├── 📄 README.md                                   <- Dokumentasi Proyek GitHub
└── 📊 Sales_Automation_Peyek_Goyang.xlsx         <- File Master Excel (Database + Dashboard)
```

---
*Dikembangkan oleh Dianita sebagai Portofolio Data Analyst.*
