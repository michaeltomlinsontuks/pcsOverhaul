# PCS V2 Integration Patterns Documentation

## Overview

This document details how PCS V2 integrates with external systems, file structures, databases, and maintains compatibility with the existing 20081222/ directory architecture. It covers all system boundaries, data exchange patterns, and dependency management.

## Core Integration Architecture

### **System Boundaries**

```
PCS V2 Application Layer
├── Forms (UI Events Only)
├── Modules (Business Logic)
├── ─────────────────────────
├── Excel Application (Host Environment)
├── Windows File System (Data Storage)
├── 20081222/ Directory Structure (Legacy Compatibility)
└── External Dependencies (Templates, Databases)
```

---

## Integration Point 1: 20081222/ Directory Structure

### **Required Directory Architecture**

PCS V2 maintains complete integration with the existing 20081222/ directory structure containing 29,035+ files:

```
20081222/
├── Archive/           # 29,035 completed job files
├── Contracts/         # 129 job template files
├── Customers/         # 86 customer data files
├── Enquiries/         # 11 active enquiry files
├── Images/            # 127 technical drawings and documents
├── Job Templates/     # 41 reusable job template files
├── Quotes/            # 14 active quote files
├── Templates/         # 21 system template files
├── WIP/              # 7 work-in-progress files
├── Search.xls        # Master search database (CRITICAL)
├── WIP.xls           # Active jobs tracking database
├── _Interface.xls    # Main system file
└── [Various operational files]
```

### **Directory Validation and Creation**

**Integration Function**: `DataOperations.ValidateDirectoryStructure()`
```vba
Public Function ValidateDirectoryStructure() As Boolean
    Dim RequiredDirs As Variant
    RequiredDirs = Array("Enquiries", "Quotes", "WIP", "Archive", "Contracts", _
                        "Customers", "Templates", "Job Templates", "images", "Backups")

    For i = 0 To UBound(RequiredDirs)
        If Not DirExists(GetRootPath & "\" & RequiredDirs(i)) Then
            ' Log missing directory and create if needed
        End If
    Next i
End Function
```

**Integration Pattern**: System validates directory structure on startup and creates missing directories only when necessary, preserving existing file organization.

### **File Path Resolution**

**Root Path Discovery**:
```vba
Public Function GetRootPath() As String
    GetRootPath = ThisWorkbook.Path  ' Auto-detects system location
End Function
```

**File Access Pattern**:
```vba
' All file operations use root-relative paths
TemplatePath = DataOperations.GetRootPath & "\Templates\_Enq.xls"
EnquiryPath = DataOperations.GetRootPath & "\Enquiries\" & EnquiryNumber & ".xls"
CustomerPath = DataOperations.GetRootPath & "\Customers\" & CustomerName & ".xls"
```

---

## Integration Point 2: Excel File Format Compatibility

### **Standard Excel File Structure**

All PCS files follow a standardized Excel format for cross-system compatibility:

**Admin Sheet Layout**:
```
Column A: Field Names (Human Readable)    Column B: Field Values (Data)
─────────────────────────────────────────────────────────────────────────
Row 1:    [Header Information]           [System Metadata]
Row 2:    File_Name                      E00001
Row 3:    Customer                       ABC Company Ltd
Row 4:    Contact_Person                 John Smith
Row 5:    Component_Description          Precision Bearing
Row 6:    Component_Code                 BRG-001
Row 7:    Component_Quantity             50
...
```

### **File Template Integration**

**Template Loading Pattern**:
```vba
' Templates are loaded as base files for new records
TemplatePath = GetRootPath & "\Templates\_Enq.xls"
Set TemplateWB = DataOperations.SafeOpenWorkbook(TemplatePath)
' Populate template with data
PopulateEnquiryTemplate TemplateWB, EnquiryInfo
' Save as new file
TemplateWB.SaveAs NewFilePath
```

**Template Types and Usage**:
- **_Enq.xls**: Base template for all enquiry files
- **_Quote.xls**: Base template for quote generation
- **_Job.xls**: Base template for job creation
- **_Customer.xls**: Base template for customer records

### **Data Population Integration**

**Admin Sheet Data Mapping**:
```vba
Private Sub PopulateEnquiryTemplate(ByRef wb As Workbook, ByRef EnquiryInfo As EnquiryData)
    Dim ws As Worksheet
    Set ws = wb.Worksheets("Admin")

    ws.Cells(2, 2).Value = EnquiryInfo.EnquiryNumber        ' B2
    ws.Cells(3, 2).Value = EnquiryInfo.CustomerName         ' B3
    ws.Cells(4, 2).Value = EnquiryInfo.ContactPerson        ' B4
    ws.Cells(5, 2).Value = EnquiryInfo.ComponentDescription ' B5
    ' Continue for all fields...
End Sub
```

**Data Loading Integration**:
```vba
Public Function LoadEnquiry(ByVal FilePath As String) As EnquiryData
    Set EnquiryWB = DataOperations.SafeOpenWorkbook(FilePath)
    Set ws = EnquiryWB.Worksheets("Admin")

    With EnquiryInfo
        .EnquiryNumber = ws.Cells(2, 2).Value
        .CustomerName = ws.Cells(3, 2).Value
        .ContactPerson = ws.Cells(4, 2).Value
        ' Continue for all fields...
    End With
End Function
```

---

## Integration Point 3: Search Database System

### **Search.xls Database Integration**

**Database Schema**:
```
Column A: RecordType    (Enquiry/Quote/Job/Contract)
Column B: RecordNumber  (E00001, Q00001, J00001)
Column C: CustomerName  (ABC Company Ltd)
Column D: Description   (Component description)
Column E: Status        (Workflow status)
Column F: DateCreated   (Creation timestamp)
Column G: FilePath      (Full path to file)
Column H: Keywords      (Searchable terms)
```

**Search Integration Pattern**:
```vba
Public Sub UpdateSearchDatabase(ByRef SearchRec As SearchRecord)
    Dim SearchWB As Workbook
    Dim ws As Worksheet

    Set SearchWB = DataOperations.SafeOpenWorkbook(GetRootPath & "\Search.xls")
    Set ws = SearchWB.Worksheets(1)

    ' Find existing record or create new
    Dim FoundRow As Long
    FoundRow = FindExistingSearchRecord(ws, SearchRec.RecordNumber)

    If FoundRow = 0 Then
        ' Add new record
        FoundRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    End If

    ' Update record data
    ws.Cells(FoundRow, 1).Value = SearchRec.RecordType
    ws.Cells(FoundRow, 2).Value = SearchRec.RecordNumber
    ws.Cells(FoundRow, 3).Value = SearchRec.CustomerName
    ' Continue for all fields...

    DataOperations.SafeCloseWorkbook SearchWB
End Sub
```

**Real-Time Search Updates**:
Every workflow operation automatically updates the search database:
- **Enquiry Created** → New search record
- **Quote Generated** → Update record status
- **Job Created** → Update record type and status
- **Job Completed** → Update to "Archived" status

### **Search History Integration**

**Search History.xls Structure**:
- Tracks all search operations
- Maintains search analytics
- Supports system usage reporting
- Syncs with main search database via password-protected operations

---

## Integration Point 4: WIP Database System

### **WIP.xls Database Integration**

**Database Schema for Active Jobs**:
```
Job tracking fields for production management:
- JobNumber, CustomerName (Identification)
- StartDate, DueDate, WorkshopDueDate (Scheduling)
- Operation01-15 fields (Production operations)
- CurrentOperation, PercentComplete (Progress tracking)
- AssignedOperator, EstimatedHours (Resource allocation)
- OrderValue, ActualCosts (Financial tracking)
```

**WIP Integration Pattern**:
```vba
Public Function UpdateWIPDatabase(ByRef JobInfo As JobData) As Boolean
    Dim WIPWB As Workbook
    Dim ws As Worksheet

    Set WIPWB = DataOperations.SafeOpenWorkbook(GetRootPath & "\WIP.xls")
    Set ws = WIPWB.Worksheets(1)

    ' Find or create job record
    Dim JobRow As Long
    JobRow = FindWIPRecord(ws, JobInfo.JobNumber)

    If JobRow = 0 Then
        JobRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    End If

    ' Update job tracking information
    ws.Cells(JobRow, 1).Value = JobInfo.JobNumber
    ws.Cells(JobRow, 2).Value = JobInfo.CustomerName
    ws.Cells(JobRow, 3).Value = JobInfo.DueDate
    ws.Cells(JobRow, 4).Value = JobInfo.AssignedOperator
    ws.Cells(JobRow, 5).Value = JobInfo.Status
    ' Continue for operations and progress fields...

    DataOperations.SafeCloseWorkbook WIPWB
End Function
```

**WIP Reporting Integration**:
The reporting system directly integrates with WIP.xls to generate:
- Operation reports (jobs grouped by operation type)
- Operator reports (jobs grouped by assigned operator)
- Due date reports (jobs sorted by delivery dates)
- Customer reports (jobs organized by customer)

---

## Integration Point 5: Number Generation System

### **Sequential Number Tracking**

**Number Tracking Integration**: Uses template files to maintain sequential numbering:

```
Templates/
├── E - 00001.TXT    # Last enquiry number
├── Q - 00001.TXT    # Last quote number
├── J - 00001.TXT    # Last job number
└── [Other tracking files]
```

**Number Generation Pattern**:
```vba
Public Function GetNextEnquiryNumber() As String
    Dim TemplateDir As String
    Dim HighestNumber As Long
    Dim FilePattern As String

    TemplateDir = GetRootPath & "\Templates\"
    FilePattern = "E - *.TXT"

    ' Scan for highest existing number
    HighestNumber = ScanForHighestNumber(TemplateDir, FilePattern)

    ' Generate next number
    GetNextEnquiryNumber = "E" & Format(HighestNumber + 1, "00000")

    ' Create tracking file
    CreateTrackingFile TemplateDir & GetNextEnquiryNumber & ".TXT"
End Function
```

**Number Reservation System**:
- Numbers are reserved immediately when requested
- Prevents conflicts in multi-user scenarios
- Failed operations can be rolled back by removing tracking files
- Maintains strict sequential numbering (E00001, E00002, Q00001, etc.)

---

## Integration Point 6: Customer Database System

### **Customer File Integration**

**Customer Directory Structure**: `Customers/CustomerName.xls`

**Customer Record Integration**:
```vba
Public Function CreateNewCustomer(CustomerName As String) As Boolean
    Dim CustomerTemplate As String
    Dim CustomerPath As String
    Dim CustomerWB As Workbook

    CustomerTemplate = GetRootPath & "\Templates\_Customer.xls"
    CustomerPath = GetRootPath & "\Customers\" & CleanFileName(CustomerName) & ".xls"

    ' Create customer file from template
    Set CustomerWB = DataOperations.SafeOpenWorkbook(CustomerTemplate)
    PopulateCustomerTemplate CustomerWB, CustomerName
    CustomerWB.SaveAs CustomerPath
    DataOperations.SafeCloseWorkbook CustomerWB
End Function
```

**Customer Integration in Forms**:
- Customer dropdowns automatically populated from Customers directory
- New customers created on-demand during enquiry entry
- Customer contact information auto-loaded when customer selected
- Customer history accessible through file system integration

---

## Integration Point 7: Contract Template System

### **Job Templates Directory Integration**

**Contract Template Structure**: `Contracts/TemplateName.xls`

**Template Integration Pattern**:
```vba
Public Function LoadJobTemplate(TemplateName As String) As ContractData
    Dim TemplatePath As String
    Dim TemplateWB As Workbook
    Dim ContractInfo As ContractData

    TemplatePath = GetRootPath & "\Contracts\" & TemplateName & ".xls"
    Set TemplateWB = DataOperations.SafeOpenWorkbook(TemplatePath)

    ' Load template data
    With ContractInfo
        .ContractName = TemplateName
        .StandardOperations = LoadOperationsFromTemplate(TemplateWB)
        .LeadTime = TemplateWB.Worksheets("Admin").Cells(10, 2).Value
        .LastUsed = Now  ' Update usage tracking
    End With

    DataOperations.SafeCloseWorkbook TemplateWB
End Function
```

**Template Usage Tracking**:
- Templates track last usage date
- Usage statistics help identify popular templates
- Template maintenance based on usage patterns

---

## Integration Point 8: Technical Drawing Management

### **Images Directory Integration**

**Technical Drawing Storage**: `Images/` directory for technical drawings and documentation

**Image Integration Pattern**:
```vba
Public Function AttachDrawingToJob(JobNumber As String, ImagePath As String) As Boolean
    Dim JobWB As Workbook
    Dim JobPath As String

    JobPath = GetRootPath & "\WIP\" & JobNumber & ".xls"
    Set JobWB = DataOperations.SafeOpenWorkbook(JobPath)

    ' Update job record with image reference
    JobWB.Worksheets("Admin").Cells(15, 2).Value = ImagePath
    JobWB.Save
    DataOperations.SafeCloseWorkbook JobWB

    ' Copy image to images directory if not already there
    If Not IsImageInDirectory(ImagePath) Then
        CopyImageToDirectory ImagePath
    End If
End Function
```

**Image Reference Integration**:
- Job cards reference technical drawings by file path
- Images stored centrally in Images directory
- Job files contain references, not embedded images (performance)
- Image viewer integration through file path references

---

## Integration Point 9: Excel Application Integration

### **Host Application Dependency**

**Excel Version Compatibility**:
```vba
Public Function ValidateExcelCompatibility() As Boolean
    Dim ExcelVersion As String
    ExcelVersion = Application.Version

    Select Case Val(ExcelVersion)
        Case Is >= 16: ' Excel 2016+
            ValidateExcelCompatibility = True
        Case Is >= 14: ' Excel 2010+
            LogError 0, "Excel version may have compatibility issues", "ValidateExcelCompatibility", "SystemCore"
            ValidateExcelCompatibility = True
        Case Else:     ' Older versions
            LogError 0, "Excel version not supported", "ValidateExcelCompatibility", "SystemCore"
            ValidateExcelCompatibility = False
    End Select
End Function
```

**Excel Object Model Integration**:
- **Application Object**: System lifecycle management
- **Workbooks Collection**: File operations and management
- **Worksheet Objects**: Data access and manipulation
- **Range Objects**: Cell-level data operations
- **Events**: Form integration and user interaction

### **32-bit vs 64-bit Excel Integration**

**Dual Deployment Strategy**:
- **SystemCore.bas**: 64-bit Excel with PtrSafe declarations
- **SystemCore32.bas**: 32-bit Excel without PtrSafe

**API Integration Pattern**:
```vba
' 64-bit version (SystemCore.bas)
Private Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
    (ByVal lpBuffer As String, ByRef nSize As Long) As Long

' 32-bit version (SystemCore32.bas)
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long
```

---

## Integration Point 10: Windows File System

### **File System Operations Integration**

**Safe File Operations Pattern**:
```vba
Public Function SafeFileOperation(FilePath As String) As Boolean
    On Error GoTo Error_Handler

    ' Validate file path
    If Not ValidateFilePath(FilePath) Then
        SafeFileOperation = False
        Exit Function
    End If

    ' Perform operation with error handling
    ' ... file operation code ...

    SafeFileOperation = True
    Exit Function

Error_Handler:
    SystemCore.LogError Err.Number, Err.Description, "SafeFileOperation", "DataOperations"
    SafeFileOperation = False
End Function
```

**File Lock Management**:
- Exclusive file access during operations
- Read-only detection and handling
- Network file sharing compatibility
- Concurrent access prevention

**Backup Integration**:
```vba
Public Function CreateBackup(FilePath As String) As Boolean
    Dim BackupPath As String
    Dim BackupDir As String

    BackupDir = GetRootPath & "\Backups\"
    BackupPath = BackupDir & Format(Now, "yyyymmdd_hhmmss_") & Dir(FilePath)

    If Not DirExists(BackupDir) Then MkDir BackupDir

    FileCopy FilePath, BackupPath
    CreateBackup = True
End Function
```

---

## Error Handling Integration Patterns

### **Centralized Error Management**

**Error Integration Architecture**:
```vba
' All modules use consistent error handling
Public Function ModuleFunction() As Boolean
    On Error GoTo Error_Handler

    ' Function logic

    ModuleFunction = True
    Exit Function

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "ModuleFunction", "ModuleName"
    ModuleFunction = False
End Function
```

**Error Logging Integration**:
- **Log File**: `error_log.txt` in system root directory
- **Log Format**: Timestamp, Module.Function, Error Number, Description
- **Integration**: All modules log errors to centralized system

### **Recovery Integration Patterns**

**File Recovery**:
- Automatic backup creation before modifications
- Rollback capability for failed operations
- Orphaned file detection and cleanup

**Database Recovery**:
- Search database consistency checking
- WIP database synchronization
- Automatic data repair where possible

---

## Performance Integration Considerations

### **File Access Optimization**

**Caching Strategy**:
```vba
' Cache frequently accessed data
Private m_CustomerCache As Variant
Private m_ComponentCache As Variant
Private m_CacheExpiry As Date

Public Function GetCustomerList() As Variant
    If IsEmpty(m_CustomerCache) Or Now > m_CacheExpiry Then
        m_CustomerCache = LoadCustomerListFromDisk()
        m_CacheExpiry = DateAdd("n", 15, Now)  ' 15-minute cache
    End If

    GetCustomerList = m_CustomerCache
End Function
```

**Batch Operations**:
- Group file operations where possible
- Minimize workbook open/close cycles
- Batch database updates

**Resource Management**:
- Proper object cleanup (Set obj = Nothing)
- Memory management for large data sets
- Connection pooling for database operations

---

## Security Integration Patterns

### **Data Protection**

**File Security**:
- Read-only file detection and handling
- Permission validation before operations
- Secure file paths (no relative path traversal)

**User Authentication Integration**:
```vba
Public Function ValidateUserAccess() As Boolean
    Dim CurrentUser As String
    CurrentUser = SystemCore.GetCurrentUser()

    ' User validation logic
    If CurrentUser = "Unknown" Then
        ValidateUserAccess = False
        Exit Function
    End If

    ' Additional security checks as needed
    ValidateUserAccess = True
End Function
```

---

## Monitoring and Diagnostics Integration

### **System Health Monitoring**

**Health Check Integration**:
```vba
Public Function PerformSystemHealthCheck() As Boolean
    Dim HealthStatus As Boolean
    HealthStatus = True

    ' Directory structure validation
    If Not DataOperations.ValidateDirectoryStructure() Then
        LogError 0, "Directory structure validation failed", "PerformSystemHealthCheck", "SystemCore"
        HealthStatus = False
    End If

    ' Critical file existence check
    If Not DataOperations.FileExists(GetRootPath & "\Search.xls") Then
        LogError 0, "Critical file missing: Search.xls", "PerformSystemHealthCheck", "SystemCore"
        HealthStatus = False
    End If

    ' Database integrity check
    If Not ValidateDatabaseIntegrity() Then
        LogError 0, "Database integrity check failed", "PerformSystemHealthCheck", "SystemCore"
        HealthStatus = False
    End If

    PerformSystemHealthCheck = HealthStatus
End Function
```

### **Usage Analytics Integration**

**System Usage Tracking**:
- Form usage statistics
- Feature utilization metrics
- Performance benchmarking
- Error frequency analysis

The PCS V2 integration patterns ensure seamless interaction with all external dependencies while maintaining complete compatibility with the existing 20081222/ file architecture and supporting the full enquiry-to-archive workflow.