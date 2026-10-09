# Dynamic KPI Dashboard & Automated PDF Generator

Proyek otomatisasi pelaporan berbasis **VBA (Visual Basic for Applications)** di Microsoft Excel untuk menyelesaikan kebutuhan manajemen bisnis: pembuatan dan distribusi laporan ringkasan performa cabang/departemen bulanan dalam format PDF siap cetak, serta pembuatan draf email Microsoft Outlook secara otomatis.

---

## 📌 Ringkasan Masalah Bisnis & Solusi

| Aspek | Deskripsi |
| :--- | :--- |
| **Masalah Bisnis** | Manajemen membutuhkan ringkasan performa bulanan untuk setiap cabang dalam format PDF siap cetak. Jika dilakukan manual (filter satu per satu, save as PDF, buat email draf satu per satu), proses ini memakan waktu berjam-jam dan rentan *human error*. |
| **Solusi Otomatisasi** | Dashboard interaktif dengan tombol & dropdown yang terhubung ke Pivot Table via VBA. Dilengkapi tombol **Batch PDF Export** (1-klik untuk generate seluruh cabang) dan integrasi **Outlook Automation** untuk membuat draf email lengkap dengan lampiran PDF masing-masing cabang. |
| **Level** | Menengah (*Intermediate - Corporate Analytics Standard*) |
| **File Utama** | `Dynamic_KPI_Dashboard_Reporting.xlsm` |

---

## 🚀 Fitur Unggulan

1. **Interactive Filter Dashboard**
   - Dropdown seleksi cabang (*Branch*) dan bulan (*Month*) pada sel `C4` dan `C5`.
   - Event `Worksheet_Change` otomatis menyinkronkan visualisasi dan metrik saat dropdown dipilih.
   - Tombol **[Apply Filter]** dan **[Reset All]** untuk fleksibilitas pengguna.
2. **Batch PDF Export (Sekali Klik)**
   - Tombol **[Generate Monthly Reports]** mengiterasi daftar cabang aktif dari lembar konfigurasi.
   - Mengatur `PageSetup` (A4 Landscape, Fit to 1 Page Wide & Tall, margin proporsional).
   - Menghasilkan file PDF individual dengan format penamaan standar:
     `Report_[NamaCabang]_[Bulan].pdf` (contoh: `Report_Jakarta_YTD_2026.pdf`).
   - Menyimpan output rapi di folder `Reports/`.
3. **Outlook Automation (Late Binding)**
   - Tombol **[Create Outlook Drafts]** membuat draf email di Microsoft Outlook tanpa memerlukan referensi pustaka awal (*Late Binding* via `CreateObject("Outlook.Application")`).
   - Mencocokkan nama manajer cabang, alamat email penerima, dan CC dari tabel konfigurasi.
   - Otomatis melampirkan file PDF cabang terkait dan menyusun *body email* berformat HTML modern dengan ringkasan metrik KPI (Revenue, Target, Achievement).

---

## 🛠️ Struktur Proyek & File

```text
d:\Data Analyst Project\VBA\
│
├── Dynamic_KPI_Dashboard_Reporting.xlsm   # File Excel Macro-Enabled siap pakai
├── build_project.py                      # Skrip automasi pembangunan workbook & inject modul
├── README.md                             # Dokumentasi teknis & panduan operasional
│
├── Reports/                              # Direktori output PDF hasil batch export
│   ├── Report_Bali_YTD_2026.pdf
│   ├── Report_Bandung_YTD_2026.pdf
│   ├── Report_Jakarta_YTD_2026.pdf
│   ├── Report_Medan_YTD_2026.pdf
│   └── Report_Surabaya_YTD_2026.pdf
│
└── src/                                  # Source code VBA modular (.bas & .cls)
    ├── mod_Config.bas                    # Konfigurasi nama sheet, range, & path
    ├── mod_DashboardFilter.bas           # Logika manipulasi PivotFields & filtering
    ├── mod_BatchExportPDF.bas            # Engine Batch Export PDF & PageSetup
    ├── mod_OutlookAutomation.bas         # Integrasi Outlook Late Binding & HTML email
    ├── mod_Utils.bas                     # Utility: speed up runtime, folder I/O, sanitasi
    └── Sheet_Dashboard.cls               # Event handler Worksheet_Change
```

---

## 📊 Struktur Sheet dalam Workbook

1. **`Dashboard`**
   - Header korporat eksekutif.
   - Panel kontrol (Dropdown Cabang & Bulan, status Active View, timestamp update).
   - 4 Kartu KPI Utama: **Actual Revenue**, **Target Revenue**, **Achievement Rate**, dan **Gross Profit & Margin**.
   - Tabel Rincian Kinerja per Departemen (Enterprise Solutions, Cloud & Infrastructure, Hardware & Devices, Consulting Services).
   - Tabel Benchmark Kinerja Antar-Cabang (Jakarta, Surabaya, Bandung, Medan, Bali).
2. **`Data`**
   - Tabel data transaksi riil (180 baris data).
   - Kolom: `Date`, `Month_Year`, `Branch`, `Department`, `Sales_Rep`, `Target`, `Actual_Revenue`, `COGS`, `Gross_Profit`.
3. **`Pivot_Calculations`**
   - Mesin komputasi berbasis Pivot Table (`ptKPISummary`).
   - Menggunakan `Branch` dan `Month_Year` sebagai Report Filters (`xlPageField`).
4. **`Settings_Config`**
   - Pengaturan lokasi folder output laporan PDF.
   - Tabel pemetaan cabang: Nama Cabang, Nama Manajer, Email Penerima, dan Email CC.

---

## 💡 Pembahasan Teknis & Skill VBA yang Ditonjolkan

### 1. Manipulasi `PivotFields` & Filter Dinamis
```vba
Dim pt As PivotTable
Dim pfBranch As PivotField

Set pt = ThisWorkbook.Sheets("Pivot_Calculations").PivotTables("ptKPISummary")
Set pfBranch = pt.PivotFields("Branch")

' Mengunci refresh sementara untuk performa tinggi
pt.ManualUpdate = True

If branchVal = "All Branches" Then
    pfBranch.ClearAllFilters
Else
    pfBranch.ClearAllFilters
    pfBranch.CurrentPage = branchVal
End If

pt.ManualUpdate = False
```
* **Best Practice**: `pt.ManualUpdate = True` menghentikan rendering kalkulasi ulang Pivot Table di setiap pergantian properti sebelum seluruh filter siap, mencegah lag visual.
* **Event Guard**: Sebelum mengubah nilai sel dropdown via kode, `Application.EnableEvents = False` diaktifkan untuk mencegah loop rekursif pada event `Worksheet_Change`.

---

### 2. Batch Export PDF dengan `ExportAsFixedFormat` & `PageSetup`
```vba
' 1. Konfigurasi Halaman agar Pas 1 Halaman Landscape
With ws.PageSetup
    .PrintArea = "A1:N27"
    .Orientation = xlLandscape
    .PaperSize = xlPaperA4
    .Zoom = False
    .FitToPagesWide = 1
    .FitToPagesTall = 1
    .CenterHorizontally = True
End With

' 2. Ekspor Tampilan Sheet menjadi File PDF
ws.ExportAsFixedFormat _
    Type:=xlTypePDF, _
    Filename:=targetPdfPath, _
    Quality:=xlQualityStandard, _
    IncludeDocProperties:=True, _
    IgnorePrintAreas:=False, _
    OpenAfterPublish:=False
```
* **Kelebihan**: Format PDF yang dihasilkan proporsional (tidak terpotong antar halaman) dan siap cetak untuk rapat direksi.
* **Sanitasi File**: Nama cabang dan bulan disaring dengan fungsi `CleanFileName()` untuk mencegah error karakter terlarang Windows seperti `\ / : * ? " < > |`.

---

### 3. Integrasi Microsoft Outlook via Late Binding
```vba
Dim outlookApp As Object
Dim mailItem As Object

' Late Binding: Tidak memerlukan checklist manual di Tools > References
On Error Resume Next
Set outlookApp = GetObject(, "Outlook.Application")
If outlookApp Is Nothing Then
    Set outlookApp = CreateObject("Outlook.Application")
End If
On Error GoTo ErrHandler

' Membuat objek email baru (0 = olMailItem)
Set mailItem = outlookApp.CreateItem(0)
With mailItem
    .To = managerEmail
    .CC = ccEmail
    .Subject = "[CONFIDENTIAL] Laporan Kinerja Bulanan - Cabang " & branchName & " (" & selectedMonth & ")"
    .HTMLBody = GenerateEmailHTMLBody(...)
    .Attachments.Add pdfPath
    .Save ' Menyimpan ke folder Drafts Outlook untuk verifikasi sebelum dikirim
End With
```
* **Late Binding vs Early Binding**: Menggunakan `CreateObject("Outlook.Application")` memastikan workbook dapat dijalankan di komputer manapun tanpa risiko error `"Missing Reference: Microsoft Outlook Object Library"`.
* **Keamanan Operasional**: Mode default menggunakan `.Save` (masuk ke folder *Drafts*) sehingga pengguna memiliki kesempatan memeriksa draf sebelum terkirim ke alamat email sungguhan.

---

### 4. Optimalisasi Kinerja & Penanganan Error
```vba
Public Sub OptimizePerformance(ByVal enableSpeedMode As Boolean)
    With Application
        If enableSpeedMode Then
            .ScreenUpdating = False
            .DisplayAlerts = False
            .EnableEvents = False
            .Calculation = xlCalculationManual
            .Cursor = xlWait
        Else
            .ScreenUpdating = True
            .DisplayAlerts = True
            .EnableEvents = True
            .Calculation = xlCalculationAutomatic
            .StatusBar = False
            .Cursor = xlDefault
        End If
    End With
End Sub
```
* Memangkas waktu eksekusi batch loop hingga 70-80% dengan menonaktifkan *screen repaint* dan kalkulasi formula otomatis berulang.

---

## 📖 Panduan Penggunaan

1. **Membuka File**
   - Buka file `Dynamic_KPI_Dashboard_Reporting.xlsm` di Microsoft Excel.
   - Jika muncul peringatan keamanan *"Security Warning: Macros have been disabled"*, klik **Enable Content**.
2. **Mengubah Tampilan Dashboard**
   - Ubah pilihan pada dropdown **Pilih Cabang** (sel `C4`) atau **Pilih Bulan** (sel `C5`).
   - Dashboard akan otomatis terupdate. Anda juga dapat menekan tombol **[Apply Filter]** atau **[Reset All]**.
3. **Mengekspor Laporan PDF Seluruh Cabang**
   - Klik tombol **[Generate Monthly Reports]**.
   - Konfirmasi dialog yang muncul.
   - Status proses akan tampil di *Status Bar* bawah Excel.
   - Setelah selesai, dialog akan menampilkan konfirmasi dan menawarkan untuk langsung membuka folder `Reports/`.
4. **Membuat Draf Email Outlook**
   - Pastikan aplikasi Microsoft Outlook terpasang.
   - Klik tombol **[Create Outlook Drafts]**.
   - Script akan memverifikasi file PDF (membuat otomatis jika belum ada), menyusun email HTML, melampirkan file PDF, dan menyimpannya di folder **Drafts** Outlook Anda.

---

## 💼 Nilai Portofolio & Wawancara Data Analyst

Proyek ini mendemonstrasikan kompetensi komprehensif:
1. **Business Sense**: Memahami alur kerja eksekutif, kebutuhan pelaporan terstandardisasi, dan kepatuhan data rahasia (*confidentiality*).
2. **Advanced Automation**: Menguasai integrasi antar-aplikasi Microsoft 365 (Excel ke PDF ke Outlook).
3. **Clean Code & Architecture**: Kode tersusun rapi dalam modul terpisah (*Separation of Concerns*), penamaan variabel konsisten (*CamelCase*), serta error handling terstruktur.

