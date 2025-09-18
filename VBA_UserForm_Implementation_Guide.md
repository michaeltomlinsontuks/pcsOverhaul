# VBA UserForm Implementation Guide
## Complete Instructions for MainV2 and frmSearchV2 Forms

### Table of Contents
1. [General Setup](#general-setup)
2. [MainV2 Form Implementation](#mainv2-form-implementation)
3. [frmSearchV2 Form Implementation](#frmsearchv2-form-implementation)
4. [Event Handler Mappings](#event-handler-mappings)
5. [Common Properties Reference](#common-properties-reference)

---

## General Setup

### Prerequisites
1. Open VBA Editor (Alt + F11)
2. Right-click VBAProject → Insert → UserForm
3. Set UserForm properties to modern styling standards

### Modern UserForm Base Properties
```vb
' Set these properties for all forms:
Font.Name = "Segoe UI" (or "Calibri")
Font.Size = 11
BackColor = &H00F0F0F0& (Light Gray)
BorderColor = &H00808080& (Medium Gray)
BorderStyle = fmBorderStyleSingle
StartUpPosition = 0 (Manual - for precise positioning)
ShowModal = True
```

---

## MainV2 Form Implementation

### Form Dimensions and Layout
```vb
' MainV2 UserForm Properties:
Width = 1100 (approximately 16500 twips)
Height = 600 (approximately 9000 twips)
Caption = "PCS Interface V2 - Enhanced Performance Dashboard"
```

### Control Layout Structure

#### 1. TOP SECTION - Master Path and Search (Y: 10-50)

**TextBox: txtMasterPath**
```vb
Name = "txtMasterPath"
Left = 20
Top = 20
Width = 400
Height = 20
Text = "C:\YourPath\"
Font.Size = 11
BackColor = &H00FFFFFF&
BorderStyle = fmBorderStyleSingle
```

**TextBox: txtGotoSearch**
```vb
Name = "txtGotoSearch"
Left = 440
Top = 20
Width = 300
Height = 20
Text = ""
Font.Size = 11
BackColor = &H00FFFFFF&
BorderStyle = fmBorderStyleSingle
```

#### 2. FILTER TOGGLE SECTION (Y: 60-100)

**Frame: fraFilters**
```vb
Name = "fraFilters"
Left = 20
Top = 60
Width = 720
Height = 40
Caption = ""
BackColor = &H00E8E8E8&
BorderStyle = fmBorderStyleSingle
```

**CheckBox Controls (Inside fraFilters):**
```vb
' chkNewEnquiries (WIP Toggle)
Name = "chkNewEnquiries"
Left = 10
Top = 10
Width = 60
Height = 20
Caption = "WIP"
Value = True
BackColor = &H00F8F8F8&

' chkQuotesToSubmit (Enquiries Toggle)
Name = "chkQuotesToSubmit"
Left = 80
Top = 10
Width = 80
Height = 20
Caption = "Enquiries"
BackColor = &H00F8F8F8&

' chkWIPToSequence (Quotes Toggle)
Name = "chkWIPToSequence"
Left = 170
Top = 10
Width = 60
Height = 20
Caption = "Quotes"
BackColor = &H00F8F8F8&

' chkJobsInWIP (Archive Toggle)
Name = "chkJobsInWIP"
Left = 240
Top = 10
Width = 80
Height = 20
Caption = "Archive"
BackColor = &H00F8F8F8&

' chkShowArchived (Jobs in WIP Toggle)
Name = "chkShowArchived"
Left = 330
Top = 10
Width = 100
Height = 20
Caption = "Jobs in WIP"
BackColor = &H00F8F8F8&

' Add "Thirties" toggle
Name = "chkThirties"
Left = 440
Top = 10
Width = 70
Height = 20
Caption = "Thirties"
BackColor = &H00F8F8F8&
```

#### 3. STATUS COUNTERS SECTION (Y: 110-150)

**Label Controls for Counters:**
```vb
' lblEnquiryCount
Name = "lblEnquiryCount"
Left = 20
Top = 120
Width = 100
Height = 20
Caption = "WIP : 15"
BackColor = &H00FFF9C4&
BorderStyle = fmBorderStyleSingle
Font.Size = 10

' lblQuoteCount
Name = "lblQuoteCount"
Left = 130
Top = 120
Width = 100
Height = 20
Caption = "Enquiries : 8"
BackColor = &H00FFF9C4&
BorderStyle = fmBorderStyleSingle
Font.Size = 10

' lblWIPCount
Name = "lblWIPCount"
Left = 240
Top = 120
Width = 100
Height = 20
Caption = "Quotes : 12"
BackColor = &H00FFF9C4&
BorderStyle = fmBorderStyleSingle
Font.Size = 10

' lblJobCount
Name = "lblJobCount"
Left = 350
Top = 120
Width = 100
Height = 20
Caption = "Archive : 5"
BackColor = &H00FFF9C4&
BorderStyle = fmBorderStyleSingle
Font.Size = 10
```

#### 4. MAIN LIST SECTION (Y: 160-400)

**ListBox: lstMain**
```vb
Name = "lstMain"
Left = 20
Top = 160
Width = 500
Height = 240
MultiSelect = fmMultiSelectSingle
ListStyle = fmListStylePlain
BackColor = &H00FFFFFF&
ForeColor = &H00000000&
Font.Name = "Courier New"
Font.Size = 11
BorderStyle = fmBorderStyleSingle
```

#### 5. STATUS DISPLAY SECTION (Y: 410-500)

**Frame: fraStatus**
```vb
Name = "fraStatus"
Left = 20
Top = 410
Width = 500
Height = 90
Caption = ""
BackColor = &H00F8F8F8&
BorderStyle = fmBorderStyleSingle
```

**Status Labels (Inside fraStatus):**
```vb
' lblFileName
Name = "lblFileName"
Left = 10
Top = 10
Width = 480
Height = 15
Caption = "File Name: JOB001_Component_Analysis"
Font.Size = 10

' lblSystemStatus
Name = "lblSystemStatus"
Left = 10
Top = 25
Width = 480
Height = 15
Caption = "System Status: IN PROGRESS"
Font.Size = 10

' lblJobNumber
Name = "lblJobNumber"
Left = 10
Top = 40
Width = 240
Height = 15
Caption = "Job Number: J2024-001"
Font.Size = 10

' lblQuoteNumber
Name = "lblQuoteNumber"
Left = 250
Top = 40
Width = 240
Height = 15
Caption = "Quote Number: Q2024-001"
Font.Size = 10

' Performance and Cache Labels
Name = "lblPerformance"
Left = 10
Top = 55
Width = 240
Height = 15
Caption = "Performance: Ready"
ForeColor = &H00008000&
Font.Size = 10

Name = "lblCacheStats"
Left = 250
Top = 55
Width = 240
Height = 15
Caption = "Cache: Initializing..."
Font.Size = 10
```

#### 6. RIGHT PANEL - ACTION BUTTONS (X: 550-800)

**Frame: fraWorkflow**
```vb
Name = "fraWorkflow"
Left = 550
Top = 60
Width = 200
Height = 120
Caption = "Workflow Actions"
BackColor = &H00E8E8E8&
BorderStyle = fmBorderStyleSingle
```

**Workflow Buttons (Inside fraWorkflow):**
```vb
' btnAddEnquiry
Name = "btnAddEnquiry"
Left = 10
Top = 20
Width = 180
Height = 25
Caption = "Add Enquiry"
BackColor = &H00F8F8F8&
BorderStyle = fmBorderStyleSingle

' btnMakeQuote
Name = "btnMakeQuote"
Left = 10
Top = 50
Width = 180
Height = 25
Caption = "Make Quote"

' btnCreateJob
Name = "btnCreateJob"
Left = 10
Top = 80
Width = 180
Height = 25
Caption = "Create Job"
```

**Frame: fraFileOps**
```vb
Name = "fraFileOps"
Left = 550
Top = 190
Width = 200
Height = 100
Caption = "File Operations"
BackColor = &H00E8E8E8&
BorderStyle = fmBorderStyleSingle
```

**File Operation Buttons (Inside fraFileOps):**
```vb
' btnEditJC, btnPrint, btnOpenWIP, btnSearch
' Use 2x2 grid layout:
Left = 10 or 100 (alternating)
Top = 20 or 50 (alternating)
Width = 80
Height = 25
```

**Frame: fraContractTools**
```vb
Name = "fraContractTools"
Left = 550
Top = 300
Width = 200
Height = 100
Caption = "Contract Tools"
```

**Frame: fraSearchHistory**
```vb
Name = "fraSearchHistory"
Left = 550
Top = 410
Width = 200
Height = 90
Caption = "Search & History"
```

#### 7. PROGRESS CONTROL

**ProgressBar: prgProgress**
```vb
Name = "prgProgress"
Left = 20
Top = 510
Width = 500
Height = 15
Visible = False
Min = 0
Max = 100
BackColor = &H00F0F0F0&
BorderStyle = fmBorderStyleSingle
```

---

## frmSearchV2 Form Implementation

### Form Dimensions
```vb
' frmSearchV2 UserForm Properties:
Width = 800 (approximately 12000 twips)
Height = 480 (approximately 7200 twips)
Caption = "PCS Search V2 - Enhanced Search Interface"
```

### Control Layout Structure

#### 1. SEARCH INPUT SECTION (Y: 10-60)

**Frame: fraSearch**
```vb
Name = "fraSearch"
Left = 20
Top = 20
Width = 760
Height = 50
Caption = ""
BackColor = &H00E8E8E8&
BorderStyle = fmBorderStyleSingle
```

**TextBox: txtSearch**
```vb
Name = "txtSearch"
Left = 10
Top = 10
Width = 600
Height = 20
Text = ""
Font.Size = 12
BackColor = &H00FFFFFF&
BorderStyle = fmBorderStyleSingle
```

**Label: lblSearchStats**
```vb
Name = "lblSearchStats"
Left = 620
Top = 10
Width = 130
Height = 20
Caption = "Enter search term to begin"
Font.Size = 10
ForeColor = &H00646464&
```

**ProgressBar: prgSearch**
```vb
Name = "prgSearch"
Left = 10
Top = 35
Width = 740
Height = 10
Visible = False
Min = 0
Max = 100
```

#### 2. RESULTS SECTION (Y: 80-320)

**ListBox: lstResults**
```vb
Name = "lstResults"
Left = 20
Top = 80
Width = 520
Height = 240
MultiSelect = fmMultiSelectSingle
ListStyle = fmListStylePlain
BackColor = &H00FFFFFF&
ColumnCount = 5
ColumnWidths = "120;80;150;150;70"
Font.Name = "Courier New"
Font.Size = 11
BorderStyle = fmBorderStyleSingle
```

#### 3. PREVIEW SECTION (Y: 330-420)

**TextBox: txtResultPreview**
```vb
Name = "txtResultPreview"
Left = 20
Top = 330
Width = 520
Height = 90
MultiLine = True
ScrollBars = fmScrollBarsBoth
BackColor = &H00F8F8F8&
Locked = True
Font.Name = "Courier New"
Font.Size = 10
BorderStyle = fmBorderStyleSingle
```

#### 4. ACTION BUTTONS PANEL (X: 560-760)

**Frame: fraFileActions**
```vb
Name = "fraFileActions"
Left = 560
Top = 80
Width = 180
Height = 100
Caption = "File Operations"
BackColor = &H00E8E8E8&
BorderStyle = fmBorderStyleSingle
```

**File Action Buttons (Inside fraFileActions):**
```vb
' btnOpenFile
Name = "btnOpenFile"
Left = 10
Top = 20
Width = 160
Height = 25
Caption = "Open File"
BackColor = &H00B0D4F1&
BorderStyle = fmBorderStyleSingle

' btnCopyPath
Name = "btnCopyPath"
Left = 10
Top = 50
Width = 160
Height = 25
Caption = "Copy Path"

' btnShowInExplorer
Name = "btnShowInExplorer"
Left = 10
Top = 75
Width = 160
Height = 25
Caption = "Show in Explorer"
```

**Frame: fraQuickActions**
```vb
Name = "fraQuickActions"
Left = 560
Top = 190
Width = 180
Height = 120
Caption = "Quick Actions"
```

**Quick Action Buttons (Inside fraQuickActions):**
```vb
' 2x2 Grid Layout:
' btnNewEnquiry, btnConvertToQuote
' btnCreateJob, btnAdvancedSearch
Left = 10 or 95 (alternating)
Top = 20, 50, 80 (three rows)
Width = 75
Height = 25
```

**Frame: fraNavigation**
```vb
Name = "fraNavigation"
Left = 560
Top = 320
Width = 180
Height = 100
Caption = "Navigation"
```

---

## Event Handler Mappings

### MainV2 Form Event Handlers

```vb
' Form Events
Private Sub UserForm_Initialize()
    ' Maps to: InitializeInterface, LoadUserPreferences, RefreshListSmart
End Sub

Private Sub UserForm_Terminate()
    ' Maps to: SaveUserPreferences, CacheManager.SaveCacheToFile
End Sub

' Filter Events
Private Sub chkNewEnquiries_Click()
    ' Maps to: currentFilters.NewEnquiries = chkNewEnquiries.Value; RefreshListSmart
End Sub

Private Sub chkQuotesToSubmit_Click()
    ' Maps to: currentFilters.QuotesToSubmit = chkQuotesToSubmit.Value; RefreshListSmart
End Sub

Private Sub chkWIPToSequence_Click()
    ' Maps to: currentFilters.WIPToSequence = chkWIPToSequence.Value; RefreshListSmart
End Sub

Private Sub chkJobsInWIP_Click()
    ' Maps to: currentFilters.JobsInWIP = chkJobsInWIP.Value; RefreshListSmart
End Sub

Private Sub chkShowArchived_Click()
    ' Maps to: currentFilters.ShowArchived = chkShowArchived.Value; RefreshListSmart
End Sub

' List Events
Private Sub lstMain_Click()
    ' Maps to: ShowPreview
End Sub

' Button Events
Private Sub btnRefresh_Click()
    ' Maps to: RefreshListSmart
End Sub

Private Sub btnSearch_Click()
    ' Maps to: frmSearchV2.Show
End Sub

Private Sub btnCacheStats_Click()
    ' Maps to: MsgBox CacheManager.GetCacheStats()
End Sub

Private Sub btnRebuildCache_Click()
    ' Maps to: CacheManager.ClearCache; CacheManager.BuildCacheInBackground
End Sub
```

### frmSearchV2 Form Event Handlers

```vb
' Form Events
Private Sub UserForm_Initialize()
    ' Maps to: InitializeSearchInterface, CacheManager.InitializeCache
End Sub

' Search Events
Private Sub txtSearch_Change()
    ' Maps to: DelayedSearch mechanism with Timer
End Sub

' Results Events
Private Sub lstResults_Click()
    ' Maps to: ShowResultPreview(lstResults.ListIndex - 1)
End Sub

Private Sub lstResults_DblClick()
    ' Maps to: OpenSelectedFile(lstResults.ListIndex - 1)
End Sub

' Action Button Events
Private Sub btnOpenFile_Click()
    ' Maps to: OpenSelectedFile
End Sub

Private Sub btnCopyPath_Click()
    ' Maps to: Copy file path to clipboard functionality
End Sub

Private Sub btnShowInExplorer_Click()
    ' Maps to: Shell "explorer.exe /select," & result.FilePath
End Sub

Private Sub btnNewEnquiry_Click()
    ' Maps to: Create new enquiry workflow
End Sub

Private Sub btnConvertToQuote_Click()
    ' Maps to: Convert enquiry to quote workflow
End Sub

Private Sub btnCreateJob_Click()
    ' Maps to: Create job from quote workflow
End Sub

Private Sub btnAdvancedSearch_Click()
    ' Maps to: frmAdvancedSearch.Show vbModal
End Sub

Private Sub btnClose_Click()
    ' Maps to: Unload Me
End Sub
```

---

## Common Properties Reference

### Standard Button Properties
```vb
BackColor = &H00F8F8F8&
BorderStyle = fmBorderStyleSingle
Font.Name = "Segoe UI"
Font.Size = 11
Height = 25
Width = 80 (small) or 160 (wide)
```

### Standard Frame Properties
```vb
BackColor = &H00E8E8E8&
BorderStyle = fmBorderStyleSingle
Font.Name = "Segoe UI"
Font.Size = 11
```

### Standard TextBox Properties
```vb
BackColor = &H00FFFFFF&
BorderStyle = fmBorderStyleSingle
Font.Name = "Segoe UI"
Font.Size = 11
Height = 20
```

### Color Scheme
```vb
Background: &H00F0F0F0& (Light Gray)
Frame Background: &H00E8E8E8& (Medium Gray)
TextBox Background: &H00FFFFFF& (White)
Button Background: &H00F8F8F8& (Very Light Gray)
Notice Labels: &H00FFF9C4& (Light Yellow)
Border Color: &H00808080& (Medium Gray)
Text Color: &H00000000& (Black)
Performance Text: &H00008000& (Green)
Status Text: &H00646464& (Dark Gray)
```

### Implementation Steps

1. **Create UserForm**: Insert → UserForm in VBA Editor
2. **Set Form Properties**: Use dimensions and properties listed above
3. **Add Controls**: Insert controls in the order listed (frames first, then contents)
4. **Set Control Properties**: Apply all properties from the tables above
5. **Add Event Handlers**: Copy event handler signatures and map to existing code
6. **Test Layout**: Verify positioning matches mockup design
7. **Link to Existing Code**: Connect event handlers to existing MainV2.frm methods

### Tips for Implementation

- Use **Tab Order**: Set TabIndex properties for logical navigation
- **Group Related Controls**: Use frames to group logically related controls
- **Consistent Spacing**: Maintain 10-pixel margins and consistent spacing
- **Font Consistency**: Use Segoe UI 11pt for modern appearance
- **Error Handling**: Add proper error handling in all event procedures
- **Performance**: Use DoEvents sparingly and only when necessary
- **Testing**: Test with different screen resolutions and DPI settings

This guide provides complete specifications for implementing both UserForms to match the HTML mockups exactly, with all necessary control properties, positioning, and event handler mappings.