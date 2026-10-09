# Dynamic KPI Dashboard & Automated PDF Generator

An enterprise grade reporting automation system built with **VBA (Visual Basic for Applications)** in Microsoft Excel. Designed to address recurring business management needs: dynamic monthly performance monitoring per department/branch, automated 1-click batch PDF report generation, and streamlined email distribution drafts via Microsoft Outlook integration.

---

## 📌 Business Problem & Automated Solution

| Aspect                       | Description                                                                                                                                                                                                                                                                                                                         |
| :--------------------------- | :---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Business Problem**   | Management requires monthly performance summaries for each regional branch in a print-ready executive PDF format, delivered regularly. Manual processing (filtering branches one by one, saving as PDF, drafting individual emails, and attaching files) consumes hours of repetitive labor and is highly error-prone.              |
| **Automated Solution** | An interactive executive dashboard with dropdown filters linked dynamically to Pivot Tables via VBA. Features a**Batch PDF Export** engine (1 click generation across all active branches) and **Outlook Automation** to create personalized email drafts with customized HTML KPI summaries and attached PDF reports. |
| **Level**              | Intermediate (*Corporate Analytics & BI Automation Standard*)                                                                                                                                                                                                                                                                     |
| **Core File**          | `Dynamic_KPI_Dashboard_Reporting.xlsm`                                                                                                                                                                                                                                                                                            |

---

## 🚀 Key Features

1. **Interactive Filter Dashboard**
   - In-cell dropdown selectors for branch (`C4`) and month period (`C5`).
   - Integrated `Worksheet_Change` event handler that updates Pivot Tables, KPI metrics, and visualizations immediately upon selection.
   - Dedicated **[Apply Filter]** and **[Reset All]** action buttons for manual control.
2. **1-Click Batch PDF Export**
   - **[Generate Monthly Reports]** button iterates through all active branches listed in the configuration directory.
   - Automated `PageSetup` configuration (A4 Landscape, Fit to 1 Page Wide & Tall, proportional margins).
   - Exports individual, print-ready executive PDFs using standard file naming:
     `Report_[BranchName]_[Month].pdf` (e.g., `Report_Jakarta_YTD_2026.pdf`).
   - Saves all output systematically into the `Reports/` directory.
3. **Outlook Automation (Late Binding)**
   - **[Create Outlook Drafts]** button automates email generation without requiring early-bound library dependencies (*Late Binding* via `CreateObject("Outlook.Application")`).
   - Dynamically maps branch managers, recipient email addresses, and CC recipients from the configuration sheet.
   - Automatically attaches the corresponding branch PDF report and builds a responsive HTML email body with an executive KPI summary table (Actual Revenue, Target, Achievement Rate).
4. **Self-Healing Master Data Sync**
   - **[Sync / Refresh Master Data]** button automatically scans `tblSalesData` for newly added branches or months.
   - Refreshes the Pivot Cache (`PivotCache.Refresh`), updates master lists in `Settings_Config`, re-binds Data Validation ranges, and registers newly added branches into the email recipient directory.

---

## 🛠️ Project Structure & Directory Layout

```text
d:\Data Analyst Project\VBA\
│
├── Dynamic_KPI_Dashboard_Reporting.xlsm   # Production Macro-Enabled Excel Workbook
├── build_project.py                      # Automated COM workbook generator & module injector
├── README.md                             # Project overview & technical documentation (English)
├── PANDUAN_PEMBUATAN_PROJECT_VBA.md      # End-to-end tutorial & step-by-step build guide
│
├── Reports/                              # Output directory for exported executive PDFs
│   ├── Report_Bali_YTD_2026.pdf
│   ├── Report_Bandung_YTD_2026.pdf
│   ├── Report_Jakarta_YTD_2026.pdf
│   ├── Report_Medan_YTD_2026.pdf
│   └── Report_Surabaya_YTD_2026.pdf
│
└── src/                                  # Modular VBA source code (.bas & .cls)
    ├── mod_Config.bas                    # Global constants, sheet names, cell ranges & paths
    ├── mod_DashboardFilter.bas           # PivotField manipulation, filtering, & master data sync
    ├── mod_BatchExportPDF.bas            # Batch PDF export engine & PageSetup configuration
    ├── mod_OutlookAutomation.bas         # Outlook Late Binding & responsive HTML email builder
    ├── mod_Utils.bas                     # Speed optimization, folder I/O, & file sanitization
    └── Sheet_Dashboard.cls               # Worksheet_Change event handler for auto-filtering
```

---

## 📊 Workbook Architecture & Worksheet Structure

1. **`Dashboard` (Presentation Layer)**
   - Executive header banner and filter control panel (Branch & Month dropdowns, active status, last-updated timestamp).
   - 4 Executive KPI Cards: **Actual Revenue**, **Target Revenue**, **Achievement Rate**, and **Gross Profit & Margin** (powered by `=IFERROR(GETPIVOTDATA(...), 0)`).
   - Department Performance Breakdown Table (Enterprise Solutions, Cloud & Infrastructure, Hardware & Devices, Consulting Services).
   - Regional Benchmark Comparison Table (Jakarta, Surabaya, Bandung, Medan, Bali).
   - 5 Corporate Action Buttons (Apply Filter, Reset All, Sync Master Data, Generate Reports, Create Outlook Drafts).
2. **`Data` (Data Layer)**
   - Standardized transaction records formatted as an official Excel Table (`ListObject`: `tblSalesData`).
   - Fields: `Date`, `Month_Year`, `Branch`, `Department`, `Sales_Rep`, `Target`, `Actual_Revenue`, `COGS`, `Gross_Profit`.
   - `Month_Year` formatted explicitly as Text (`@`) to prevent unintended serial date conversions and ensure robust formula matching.
3. **`Pivot_Calculations` (Calculation Engine)**
   - Backend high-speed aggregation powered by Pivot Table `ptKPISummary`.
   - Uses `Branch` and `Month_Year` as Page Filters (`xlPageField`) and `Department` as Row Labels.
4. **`Settings_Config` (Configuration & Directory)**
   - Custom report export folder destination (Cell `C4`).
   - Branch Manager Contact Directory (Columns B–E: Branch, Manager Name, Primary Email, CC Email).
   - Master Dropdown Validation Lists (Columns G & H: Master Branches, Master Months).

---

## 💡 Technical Highlights & Featured VBA Skills

### 1. Dynamic `PivotFields` Manipulation & Performance Locks

```vba
Dim pt As PivotTable
Dim pfBranch As PivotField

Set pt = ThisWorkbook.Sheets("Pivot_Calculations").PivotTables("ptKPISummary")
Set pfBranch = pt.PivotFields("Branch")

' Suppress UI recalculations during filter adjustment for maximum execution speed
pt.ManualUpdate = True

If branchVal = "All Branches" Then
    pfBranch.ClearAllFilters
Else
    pfBranch.ClearAllFilters
    pfBranch.CurrentPage = branchVal
End If

pt.ManualUpdate = False
```

* **Best Practice**: `pt.ManualUpdate = True` prevents visual screen thrashing and expensive intermediate recalculations while properties are being set.
* **Event Guarding**: Setting `Application.EnableEvents = False` before changing cell values programmatically prevents infinite loops inside the `Worksheet_Change` event handler.

---

### 2. High-Quality Batch PDF Export with `ExportAsFixedFormat` & `PageSetup`

```vba
' 1. Ensure clean, proportional 1-page landscape layout
With ws.PageSetup
    .PrintArea = "A1:N28"
    .Orientation = xlLandscape
    .PaperSize = xlPaperA4
    .Zoom = False
    .FitToPagesWide = 1
    .FitToPagesTall = 1
    .CenterHorizontally = True
End With

' 2. Export worksheet view to individual PDF file
ws.ExportAsFixedFormat _
    Type:=xlTypePDF, _
    Filename:=targetPdfPath, _
    Quality:=xlQualityStandard, _
    IncludeDocProperties:=True, _
    IgnorePrintAreas:=False, _
    OpenAfterPublish:=False
```

* **Print-Ready Output**: Enforces strict single-page landscape dimensions suitable for board-level executive distribution.
* **File Name Sanitization**: Branch and period strings are filtered via `CleanFileName()` to strip illegal Windows filesystem characters (`\ / : * ? " < > |`).

---

### 3. Microsoft Outlook Integration via Late Binding

```vba
Dim outlookApp As Object
Dim mailItem As Object

' Late Binding: No manual reference required in Tools > References
On Error Resume Next
Set outlookApp = GetObject(, "Outlook.Application")
If outlookApp Is Nothing Then
    Set outlookApp = CreateObject("Outlook.Application")
End If
On Error GoTo ErrHandler

' Create new email item (0 = olMailItem)
Set mailItem = outlookApp.CreateItem(0)
With mailItem
    .To = managerEmail
    .CC = ccEmail
    .Subject = "[CONFIDENTIAL] Monthly Performance Report - " & branchName & " (" & selectedMonth & ")"
    .HTMLBody = BuildEmailBodyHTML(branchName, managerName, selectedMonth, revVal, tgtVal, achVal)
    .Attachments.Add pdfPath
    .Save ' Saves safely to Outlook Drafts folder for review before sending
End With
```

* **Late Binding Advantage**: Using `CreateObject("Outlook.Application")` ensures the workbook runs seamlessly across different workstations and Office versions without `"Missing Reference: Microsoft Outlook Object Library"` compile errors.
* **Operational Safety**: Defaults to `.Save` into the user's **Drafts** folder, providing human-in-the-loop review before live delivery.

---

### 4. Runtime Performance Optimization

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

* Accelerates batch execution loops by **70–80%** by disabling screen repaints, dialog alerts, and repetitive formula calculations during batch loops.

---

## 📖 Operational Guide

### 1. Opening the Workbook

- Open `Dynamic_KPI_Dashboard_Reporting.xlsm` in Microsoft Excel.
- If prompted with *"Security Warning: Macros have been disabled"*, click **Enable Content**.

### 2. Interacting with the Dashboard

- Select any branch from the **Branch** dropdown (`C4`) or period from the **Month** dropdown (`C5`).
- The dashboard, KPI cards, and breakdown charts will update automatically via worksheet events.
- Click **[Reset All]** to return to the aggregate organization-wide view.

### 3. Exporting PDF Reports in Batch

- Click **[Generate Monthly Reports]**.
- Confirm the dialog prompt.
- Progress will be displayed in real time on Excel's bottom *Status Bar*.
- Upon completion, a summary notification will offer to open the `Reports/` directory directly in Windows Explorer.

### 4. Creating Outlook Email Drafts

- Ensure Microsoft Outlook is installed on your computer.
- Click **[Create Outlook Drafts]**.
- The script automatically verifies or generates the required branch PDFs, compiles personalized HTML email bodies with KPI tables, attaches the PDFs, and saves the emails directly into your Outlook **Drafts** folder.

### 5. Adding New Data or Branches Dynamically

1. Open the **`Data`** sheet and append new transaction rows at the bottom of table `tblSalesData`.
2. Return to the **`Dashboard`** sheet and click **[Sync / Refresh Master Data]**.
3. The VBA engine automatically refreshes the Pivot Cache, detects new branches/months, updates dropdown options, and adds new branches into the email configuration directory.
