# 📊 Sales Automation & Master Data System for MSME Peyek Goyang

![Project Status](https://img.shields.io/badge/Project%20Status-Completed-brightgreen)
![Tools](https://img.shields.io/badge/Tools-Microsoft%20Excel%20%7C%20Advanced%20Formulas%20%7C%20Pivot%20Table-blue)

## 📌 Business Overview & Problem Statement
**UMKM Peyek Goyang** is a micro-enterprise operating in the snack culinary sector. Prior to this project, all daily transaction recording was done manually and was not structured.

This manual process posed several critical bottlenecks:
- **No Single Source of Truth:** There was no Product Master & Customer Master, so product names and prices were prone to duplication and input errors.
- **No Input Control:** Manual inputs were not validated, increasing the risk of incorrect product names and price errors.
- **Difficulty in Monitoring:** The business owner struggled to monitor sales performance because data was not centralized in one structured file.

---

## 📸 Transformation: Before vs After

| Before (Manual Records) | After (Structured Excel System) |
| :---: | :---: |
| Manual, unstructured records, high risk of duplication & price errors. | Structured relational database system in Excel, with Product Master & Customer Master as Single Source of Truth, automated formulas, and interactive dashboard. |
|<img width="1280" height="720" alt="before" src="https://github.com/user-attachments/assets/b82cab3f-cbe8-4371-9d6c-e32914a9cf8f" />| https://github.com/user-attachments/assets/36f3f3bb-3a63-4247-ad18-437f43940213|

---

## 🛠️ Project Architecture & Key Features

This project transforms manual records into a structured database schema in Microsoft Excel consisting of 3 primary worksheets:

### 1. `Dashboard` (Executive Summary & Visual Analysis)
- **Interactive Pivot Tables & Pivot Charts:** Presents visual summaries of Total Revenue, Total Orders, Units Sold (Pcs), and Average Purchase Value.
- **Dynamic Slicers & Timeline:** Interactive filter buttons based on packaging variants and time ranges (Date/Month/Year).

### 2. `Data_Pesanan` (Raw Transactional Data)
- **Automated Order ID :**  Automatic order ID numbering (`TRX-0001`, `TRX-0002`, etc.) that is filled automatically when the order column is populated.
- **Dynamic Day Name Extraction:** Automatically converts dates into day names using the function `=TEXT(Date, "dddd")`.
- **Automated Customer ID & Classification:**  Automatic standardization of customer qualification (Regular Customer vs *Walk-in*). If registered in the master data, the ID is automatically pulled (`P001`, `P002`, ...); if it is a new / general customer, it is automatically qualified as `P000`.
  ```excel
  =IF(ISNA(XLOOKUP(E2; Master_Pelanggan!$A$2:$A$10; Master_Pelanggan!$B$2:$B$10)); "P000"; XLOOKUP(E2; Master_Pelanggan!$A$2:$A$10; Master_Pelanggan!$B$2:$B$10))
  ```
- **Dynamic Pricing (XLOOKUP):** Automatically eliminates manual price input errors based on *Master Pricing*.
- **Data Validation & Flexible Dropdown:** Ensures consistency in writing variant names and operational time without limiting the flexibility of new customer name inputs.

### 3. `Master Data` (Product Master & Customer Master)
- Centralized Product Master & Customer Master as the only reference for all transactions.

---


## 🎯 Business Outcome & Impact

- ⏱️ **Centralized & Structured Data:** Transformed manual records into one structured file with Product Master & Customer Master as Single Source of Truth.
- 🎯 **Prevented Input Errors:** Data Validation & Data Quality Check prevent product name duplication & price errors by rejecting inputs outside the master.
- 📊 **Faster Operational Monitoring:** Menyajikan data penjualan secara *real-time* yang memudahkan pemilik usaha dalam menentukan tingkat produksi dan inventaris stok.
Interactive Dashboard (Pivot + Slicer + Timeline) allows the owner to monitor sales without manual recapitulation.
---

## 🧰 Tools & Technologies Used
- **Software:** Microsoft Excel
- **Features & Methods:** Advanced Formulas (`XLOOKUP`, `IF`, `ISNA`, `TEXT`), Data Validation Dropdowns, Pivot Table, Pivot Chart, Dynamic Slicers, Conditional Formatting, Database Normalization.

---
