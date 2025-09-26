# PCS V2 Module API Reference

## 📋 **Module Overview**

This document provides detailed API reference for all PCS V2 modules. Each module has clear responsibilities and well-defined interfaces.

**Module Architecture**:
```
SystemCore.bas          → Infrastructure layer (error handling, validation)
DataOperations.bas      → Data access layer (files, templates, numbers)
WorkflowManagement.bas  → Business logic layer (processes, workflows)
BusinessLogic.bas       → Business logic layer (search, records)
UserInterface.bas       → Controller layer (form coordination)
ReportingSystem.bas     → Reporting layer (WIP reports, export)
```

---

## 🛠️ **SystemCore.bas** - Infrastructure Layer

### **Purpose**
Core infrastructure services including error handling, logging, validation, and system utilities.

### **Key Constants**
```vba
' Error handling constants
Private Const ERROR_LOG_FILE As String = "error_log.txt"
Private Const MAX_LOG_SIZE As Long = 1048576  ' 1MB
```

### **Public Functions**

#### **Error Management**
```vba
Sub HandleStandardErrors(ErrorNumber As Long, ProcedureName As String, ModuleName As String)
```
**Purpose**: Standardized error handling for all modules
**Parameters**:
- `ErrorNumber` (Long): VBA error number
- `ProcedureName` (String): Name of function where error occurred
- `ModuleName` (String): Name of module where error occurred
**Usage**: Call in every error handler across all modules

```vba
Sub LogError(ErrorNumber As Long, ErrorDescription As String, ProcedureName As String, ModuleName As String)
```
**Purpose**: Log errors to file with timestamp
**Side Effects**: Creates/appends to error_log.txt

#### **User Feedback**
```vba
Sub ShowError(Message As String, Title As String)
Sub ShowWarning(Message As String, Title As String)
Sub ShowInformation(Message As String, Title As String)
```
**Purpose**: Display user-friendly messages with appropriate icons

#### **Validation Functions**
```vba
Function ValidateReportSelection(ReportForm As Object) As Boolean
```
**Purpose**: Validate WIP report form selections
**Returns**: True if valid selections, False if validation fails

```vba
Function CleanFileName(FileName As String) As String
```
**Purpose**: Remove invalid characters from file names
**Returns**: Sanitized file name safe for Windows file system

#### **Utility Functions**
```vba
Function BackupFile(FilePath As String) As Boolean
```
**Purpose**: Create backup copy of file with timestamp
**Returns**: True if backup successful

```vba
Function GetCurrentUser() As String
```
**Purpose**: Get Windows username (32/64-bit compatible)
**Returns**: Current user name or "Unknown" if unable to retrieve

---

## 💾 **DataOperations.bas** - Data Access Layer

### **Purpose**
All file operations, template management, number generation, and WIP database interactions.

### **Key Constants**
```vba
Private Const ROOT_PATH As String = ""  ' Set at runtime
Private Const BACKUP_FOLDER As String = "\Backups\"
Private Const WIP_FILE As String = "WIP.xls"
Private Const SEARCH_FILE As String = "Search.xls"
```

### **Public Functions**

#### **File Operations**
```vba
Function SafeOpenWorkbook(FilePath As String, Optional ReadOnly As Boolean = False) As Workbook
```
**Purpose**: Safely open Excel workbook with error handling
**Returns**: Workbook object or Nothing if failed
**Usage**: Always use instead of Workbooks.Open

```vba
Sub SafeCloseWorkbook(WB As Workbook, Optional SaveChanges As Boolean = False)
```
**Purpose**: Safely close workbook with error handling
**Parameters**:
- `SaveChanges` (Boolean): True to save before closing

```vba
Function FileExists(FilePath As String) As Boolean
Function GetRootPath() As String
```

#### **Number Generation System**
```vba
Function GetNextEnquiryNumber() As String
Function GetNextQuoteNumber() As String
Function GetNextJobNumber() As String
```
**Purpose**: Generate next sequential numbers for each record type
**Returns**: Formatted number string (e.g., "E001234", "Q001234", "J001234")
**Thread-Safe**: Uses file locking to prevent conflicts

```vba
Function Calc_Next_Number(RecordType As String) As String
Sub Confirm_Next_Number(RecordType As String)
```
**Purpose**: Legacy-compatible number generation
**RecordType**: "Enquiry", "Quote", or "Job"

#### **Template Management**
```vba
Function PopulateEnquiryTemplate(EnquiryInfo As EnquiryRecord, TemplatePath As String) As Boolean
Function PopulateQuoteTemplate(QuoteInfo As QuoteRecord, TemplatePath As String) As Boolean
Function PopulateJobTemplate(JobInfo As JobRecord, TemplatePath As String) As Boolean
```
**Purpose**: Populate Excel templates with form data
**Returns**: True if successful, False if error
**Side Effects**: Creates new file with populated data

#### **Data Extraction**
```vba
Function GetRangeValues(FilePath As String, SheetName As String, RangeAddress As String) As Variant
Function GetComponentCodes() As Variant
Function GetMaterialGrades() As Variant
Function GetCustomerList() As Variant
```
**Purpose**: Extract data arrays from Excel files
**Returns**: Variant array of values or empty array if error

#### **WIP Database**
```vba
Function SaveInfoIntoWIP(JobData As JobRecord) As Boolean
```
**Purpose**: Save job information to WIP database
**Returns**: True if successful
**Side Effects**: Creates/updates WIP.xls file

#### **Date Formatting**
```vba
Function FormatDate(ByVal DateValue As Date) As String
Function FormatDateTime(ByVal DateValue As Date) As String
```
**Purpose**: Consistent date formatting across system
**Returns**: Formatted date strings ("dd/mm/yyyy" or "dd/mm/yyyy hh:mm")

---

## 🔄 **WorkflowManagement.bas** - Business Logic Layer

### **Purpose**
Core business processes, form initialization, and workflow coordination.

### **Public Functions**

#### **Form Initialization**
```vba
Function InitializeEnquiryForm(EnquiryForm As Object) As Boolean
Function InitializeQuoteForm(QuoteForm As Object) As Boolean
Sub InitializeJobGenerationForm(JobForm As Object)
```
**Purpose**: Set up forms with default values and populate dropdowns
**Returns**: Boolean for success/failure (where applicable)
**Side Effects**: Populates form controls with data

#### **Workflow Operations**
```vba
Function SaveEnquiry(EnquiryForm As Object) As Boolean
Function SaveQuote(QuoteForm As Object) As Boolean
Function SaveJob(JobForm As Object) As Boolean
```
**Purpose**: Save form data to appropriate files
**Returns**: True if save successful
**Side Effects**: Creates files in appropriate directories

```vba
Function ConvertToQuote(MainForm As Object) As Boolean
Function SubmitQuote(MainForm As Object) As Boolean
```
**Purpose**: Business process transitions
**Returns**: True if operation successful
**Side Effects**: Creates new files, updates statuses

#### **Job Card Operations**
```vba
Sub PrintJobCard(JobCardForm As Object)
Sub UpdateOperations(JobCardForm As Object)
```
**Purpose**: Job card printing and operation updates
**Side Effects**: Sends to printer, updates job records

#### **Event Handlers**
```vba
Sub HandleJobStatusChange(JobForm As Object, NewStatus As String)
Sub HandleFormValidation(FormObject As Object) As Boolean
```
**Purpose**: Respond to workflow events and validate form data

---

## 🔍 **BusinessLogic.bas** - Search & Records Layer

### **Purpose**
Search operations, record management, and data validation.

### **Public Functions**

#### **Search Operations**
```vba
Function SearchRecords(SearchTerm As String, Optional RecordTypeFilter As String) As Variant
Function SearchRecords_Optimized(SearchTerm As String, Optional RecordTypeFilter As String) As Variant
```
**Purpose**: Search across all record types
**Parameters**:
- `SearchTerm` (String): Text to search for
- `RecordTypeFilter` (String): Optional filter ("Enquiry", "Quote", "Job")
**Returns**: Array of matching records

```vba
Function Update_Search() As Boolean
Function SeachSYNC() As Boolean
```
**Purpose**: Update search database and synchronize with files
**Returns**: True if update successful
**Side Effects**: Updates Search.xls file

#### **Record Retrieval**
```vba
Function GetEnquiryInfo(EnquiryNumber As String) As EnquiryRecord
Function GetQuoteInfo(QuoteNumber As String) As QuoteRecord
Function GetJobInfo(JobNumber As String) As JobRecord
```
**Purpose**: Retrieve complete record information
**Returns**: Populated record structure or empty structure if not found

#### **History Functions**
```vba
Function GetJobHistory() As Variant
Function GetQuoteHistory() As Variant
```
**Purpose**: Retrieve historical data for display
**Returns**: Array of historical records

#### **Data Validation**
```vba
Function ValidateEnquiryData(EnquiryInfo As EnquiryRecord) As Boolean
Function ValidateQuoteData(QuoteInfo As QuoteRecord) As Boolean
Function ValidateJobData(JobInfo As JobRecord) As Boolean
```
**Purpose**: Validate business data before saving
**Returns**: True if data valid, False with error messages if invalid

---

## 🎯 **UserInterface.bas** - Controller Layer

### **Purpose**
Form coordination, navigation, and user interface management.

### **Public Functions**

#### **Form Management**
```vba
Function ShowForm(FormName As String, Optional InitializeData As Variant) As Boolean
Sub ShowMenu()
Sub DisplayStatusMessage(Message As String, MessageType As String)
```
**Purpose**: Control form display and initialization
**FormName**: Name of form to display
**MessageType**: "Info", "Warning", "Error"

#### **List Management**
```vba
Sub PopulateList(ListForm As Object, DataArray As Variant, ListTitle As String)
Function ExportListResults(ListForm As Object) As Boolean
```
**Purpose**: Manage list displays and export functionality
**Side Effects**: Populates list controls, creates export files

#### **Search Interface**
```vba
Sub ShowJobHistory(MainForm As Object)
Sub ShowQuoteHistory(MainForm As Object)
Function OpenSearchDatabase() As Boolean
```
**Purpose**: Interface with search functionality
**Side Effects**: Opens Search.xls file directly

---

## 📊 **ReportingSystem.bas** - Reporting Layer

### **Purpose**
WIP report generation, data export, and report formatting.

### **Key Constants**
```vba
Private Const DATE_FORMAT_DISPLAY As String = "dd/mm/yyyy"
Private Const DATE_FORMAT_DISPLAY_TIME As String = "dd/mm/yyyy hh:mm"
Private Const DATE_FORMAT_EXCEL_COLUMN As String = "dd/mm/yyyy"
Private Const DATE_FORMAT_FILE_TIMESTAMP As String = "yyyymmdd_hhmmss"
```

### **Public Functions**

#### **Report Generation**
```vba
Function GenerateWIPReports(ReportForm As Object) As Boolean
Sub GenerateOperationReports(Job() As Jobs, JobCount As Integer)
Sub GenerateOperatorReports(Job() As Jobs, JobCount As Integer)
```
**Purpose**: Generate various WIP report types
**Returns**: Boolean for success/failure
**Side Effects**: Creates Excel files with formatted reports

#### **Data Export**
```vba
Function ExportWIPData(Optional ExportPath As String) As Boolean
Function CreateWIPExport(ExportWB As Workbook, Job() As Jobs, JobCount As Integer) As Boolean
```
**Purpose**: Export WIP data to Excel format
**ExportPath**: Optional custom export location
**Side Effects**: Creates export files

#### **Statistics and Analysis**
```vba
Function GetWIPSummaryStatistics() As Variant
Function GetOldestJobDate(Job() As Jobs, JobCount As Integer) As Date
Function GetNewestJobDate(Job() As Jobs, JobCount As Integer) As Date
Function GetAverageJobAge(Job() As Jobs, JobCount As Integer) As Double
```
**Purpose**: Calculate WIP statistics for management reporting
**Returns**: Statistical data for display

---

## 📝 **Data Structures**

### **EnquiryRecord**
```vba
Type EnquiryRecord
    EnquiryNumber As String
    CustomerName As String
    Description As String
    DateCreated As Date
    Status As String
    ' Additional fields...
End Type
```

### **QuoteRecord**
```vba
Type QuoteRecord
    QuoteNumber As String
    EnquiryNumber As String
    CustomerName As String
    Description As String
    DateCreated As Date
    ValidUntil As Date
    Status As String
    ' Additional fields...
End Type
```

### **JobRecord**
```vba
Type JobRecord
    JobNumber As String
    QuoteNumber As String
    CustomerName As String
    Description As String
    DateCreated As Date
    Status As String
    ' Additional fields...
End Type
```

### **Jobs (WIP Reporting)**
```vba
Type Jobs
    Dat As Date                    ' Start date
    Cust As String                 ' Customer
    Job As String                  ' Job number
    JobD As Double                 ' Job number (numeric)
    Qty As String                  ' Quantity
    Cod As String                  ' Code
    Desc As String                 ' Description
    Remarks As String              ' Remarks
    DDat As String                 ' Due date
    OperatorN(1 To 15) As String   ' Operator names
    OperatorType(1 To 15) As String ' Operation types
End Type
```

---

## 🧪 **Testing Framework**

### **TestWorkflows.bas**
```vba
Sub TestSystemOperations()        ' Complete system validation
Sub TestDataOperations()          ' Test file operations
Sub TestBusinessLogic()           ' Test search and records
Sub TestWorkflowManagement()      ' Test business processes
Sub TestSearchFunctionality()     ' Test search operations
Sub TestWIPOperations()           ' Test WIP reporting
```

**Usage**: Run these functions to validate system components after changes.

---

## 🔧 **Development Guidelines**

### **Function Documentation Pattern**
```vba
' **Purpose**: Brief description of what function does
' **Original**: Interface_VBA/OriginalModule.bas.OriginalFunction (if applicable)
' **Parameters**:
'   - param1 (Type): Description
' **Returns**: Type - Description
' **Dependencies**: List of modules/functions required
' **Side Effects**: Any external effects (files created, global state changed)
' **Errors**: How errors are handled
```

### **Error Handling Pattern**
```vba
Public Function MyFunction(param As String) As Boolean
    On Error GoTo Error_Handler

    ' Function logic here
    MyFunction = True
    Exit Function

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "MyFunction", "ModuleName"
    MyFunction = False
End Function
```

### **Module Dependencies**
```
SystemCore.bas          ← No dependencies (base infrastructure)
DataOperations.bas      ← SystemCore
BusinessLogic.bas       ← SystemCore, DataOperations
WorkflowManagement.bas  ← SystemCore, DataOperations
UserInterface.bas       ← SystemCore, BusinessLogic, WorkflowManagement
ReportingSystem.bas     ← SystemCore, DataOperations
```

---

**This API reference provides complete coverage of all PCS V2 modules. Follow the established patterns when extending or modifying the system.**