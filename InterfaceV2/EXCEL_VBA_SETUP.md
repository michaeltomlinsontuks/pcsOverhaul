# Excel VBA Setup Guide for InterfaceV2

## Important: Forms Must Be Recreated Manually

The `.frm` files in this directory contain **only the VBA code** - the form layouts must be recreated manually in Excel.

## Setup Steps

### 1. Import VBA Modules (.bas files)
In Excel VBA Editor:
1. Right-click VBA Project → Insert → Module
2. Copy/paste code from each `.bas` file:
   - `CacheManager.bas`
   - `DataTypes.bas`
   - `FileUtilities.bas`
   - `InterfaceLauncher.bas`
   - `PerformanceMonitor.bas`
   - `SearchEngineV2.bas`

### 2. Create UserForms Manually

#### MainV2 Form:
1. Insert → UserForm
2. Set properties:
   - Name: `MainV2`
   - Caption: `PCS Interface V2 - Enhanced Performance Dashboard`
   - Width: 16500
   - Height: 9000
3. Add controls (see original form layout)
4. Copy/paste code from `MainV2.frm` into form code module

#### frmSearchV2 Form:
1. Insert → UserForm
2. Set properties:
   - Name: `frmSearchV2`
   - Caption: `PCS Search V2 - Enhanced Search Interface`
   - Width: 12000
   - Height: 7200
3. Add controls (see original form layout)
4. Copy/paste code from `frmSearchV2.frm` into form code module

### 3. Required Controls

#### MainV2 Controls:
- `lstMain` (ListBox)
- `txtPreview` (TextBox - MultiLine, ScrollBars)
- `lblPerformance` (Label)
- `lblEnquiryCount`, `lblQuoteCount`, `lblWIPCount`, `lblJobCount` (Labels)
- `lblCacheStats`, `lblStatus` (Labels)
- `prgProgress` (ProgressBar or substitute)
- `chkNewEnquiries`, `chkQuotesToSubmit`, `chkWIPToSequence`, `chkJobsInWIP`, `chkShowArchived` (CheckBoxes)
- `btnRefresh`, `btnSearch`, `btnCacheStats`, `btnRebuildCache` (CommandButtons)

#### frmSearchV2 Controls:
- `txtSearch` (TextBox)
- `lstResults` (ListBox - MultiSelect, 5 columns)
- `txtResultPreview` (TextBox - MultiLine, ScrollBars)
- `lblSearchStats`, `lblSearchStatus` (Labels)
- `prgSearch` (ProgressBar or substitute)
- `btnOpenFile`, `btnCopyPath`, `btnShowInExplorer`, `btnAdvancedSearch` (CommandButtons)
- `btnNewEnquiry`, `btnConvertToQuote`, `btnCreateJob`, `btnClose` (CommandButtons)

### 4. Test
Run: `InterfaceLauncher.OpenMainInterface`

## Notes
- Form layouts were originally designed for VB6 - adapt control sizes/positions as needed for Excel
- Some controls (like ProgressBar) may need Excel-compatible substitutes
- All VBA code is now standard Excel VBA compatible