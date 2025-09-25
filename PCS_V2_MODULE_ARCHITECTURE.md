# PCS V2 Module Architecture Documentation

## System Overview

PCS V2 refactors the original scattered VBA modules into 6 logical modules that maintain all original functionality while providing cleaner code organization. Each module has specific responsibilities and clear boundaries.

## Module Structure

### **SystemCore.bas** (Infrastructure Foundation)

**Primary Responsibility**: Core system infrastructure, data types, error handling, and validation framework

**Key Components**:
- **Windows API Functions**: User identification with 32/64-bit compatibility
- **Data Type Definitions**: Complete system data structures (EnquiryData, QuoteData, JobData, ContractData, SearchRecord, SystemConfig)
- **Error Handling Framework**: Centralized logging and standard error processing
- **Validation Framework**: Field validation functions with popup messages
- **Legacy Compatibility Functions**: Remove_Characters(), Insert_Characters()

**Replaces Original Modules**:
- CoreFramework.bas (error handling)
- ValidationFramework.bas (form validation)
- GetUserName32.bas / GetUserName64.bas (user identification)
- RemoveCharacters.bas (string utilities)
- Very_HiddenSheet.bas (worksheet management)
- Delete_Sheet.bas (worksheet operations)

**Key Functions**:
```vba
' Data Types
Public Type EnquiryData
Public Type QuoteData
Public Type JobData
Public Type ContractData

' Error Handling
Public Sub LogError(ErrorNumber, Description, Function, Module)
Public Sub HandleStandardErrors(ErrorNumber, Function, Module)

' Validation
Public Function ValidateRequired(Value, FieldName, FormObject) As Boolean
Public Function ValidateNumeric(Value, FieldName, FormObject) As Boolean
Public Function ValidateDate(Value, FieldName, FormObject) As Boolean

' Legacy Compatibility
Public Function Remove_Characters(Str As String) As String
Public Function Insert_Characters(Str As String) As String
```

**Dependencies**: None (base layer)

---

### **DataOperations.bas** (File System Operations)

**Primary Responsibility**: All file operations, Excel data access, and directory management

**Key Components**:
- **File System Operations**: Directory validation, file existence checking, path management
- **Excel Workbook Operations**: Safe workbook opening/closing with error handling
- **Directory Structure Management**: PCS directory validation and creation
- **File List Operations**: Directory enumeration and filtering

**Replaces Original Modules**:
- DataManager.bas (file management)
- DataUtilities.bas (data access utilities)
- Open_Book.bas (workbook operations)
- GetValue.bas (cell data retrieval)
- Check_Dir.bas (directory operations)

**Key Functions**:
```vba
' Directory Management
Public Function GetRootPath() As String
Public Function ValidateDirectoryStructure() As Boolean
Public Function CreateDirectoryStructure() As Boolean

' File Operations
Public Function FileExists(FilePath As String) As Boolean
Public Function DirExists(DirPath As String) As Boolean
Public Function GetFileList(DirectoryPath As String) As Variant

' Excel Operations
Public Function SafeOpenWorkbook(FilePath As String) As Workbook
Public Sub SafeCloseWorkbook(ByRef wb As Workbook)
Public Function GetCellValue(FilePath, SheetName, CellRef) As Variant

' Number Generation
Public Function GetNextEnquiryNumber() As String
Public Function GetNextQuoteNumber() As String
Public Function GetNextJobNumber() As String
```

**Dependencies**: SystemCore (for error handling)

---

### **BusinessLogic.bas** (Business Process Controller)

**Primary Responsibility**: Core business processes, workflow management, and search functionality

**Key Components**:
- **Enquiry Management**: Create and validate enquiries
- **Quote Management**: Generate quotes from enquiries
- **Job Management**: Create jobs from accepted quotes
- **Search Database Management**: Maintain search index and history
- **Customer Management**: Customer record operations
- **Data Validation**: Business rule enforcement

**Replaces Original Modules**:
- BusinessController.bas (business rules)
- SearchManager.bas (search operations)
- SaveSearchCode.bas (search database updates)
- SaveWIPCode.bas (WIP database management)

**Key Functions**:
```vba
' Core Business Processes
Public Function CreateEnquiry(ByRef EnquiryInfo As EnquiryData) As Boolean
Public Function CreateQuote(ByRef QuoteInfo As QuoteData) As Boolean
Public Function CreateJob(ByRef JobInfo As JobData) As Boolean

' Data Validation
Public Function ValidateEnquiryData(EnquiryInfo As EnquiryData) As String
Public Function ValidateQuoteData(QuoteInfo As QuoteData) As String
Public Function ValidateJobData(JobInfo As JobData) As String

' Search Operations
Public Function SearchRecords(SearchTerm As String) As Variant
Public Sub UpdateSearchDatabase(SearchRecord As SearchRecord)
Public Sub SynchronizeSearchHistory()

' Customer Management
Public Function CreateNewCustomer(CustomerName As String) As Boolean
Public Function GetCustomerList() As Variant
```

**Dependencies**: SystemCore (data types, validation), DataOperations (file operations)

---

### **WorkflowManagement.bas** (Document Lifecycle Management)

**Primary Responsibility**: Complete document lifecycle management and form processing

**Key Components**:
- **Form Processing**: Extract business logic from all forms
- **Workflow Orchestration**: Manage Enquiry → Quote → Job transitions
- **Form Initialization**: Load form data from templates and existing files
- **Data Population**: Transfer data between forms and files
- **Template Management**: Handle system templates and job cards

**Replaces Original Modules**:
- EnquiryManager.bas (enquiry processing)
- QuoteManager.bas (quote processing)
- QuoteAcceptanceManager.bas (quote-to-job transition)
- JobCardManager.bas (job card operations)
- JobGenerationManager.bas (direct job creation)
- **Form Business Logic** extracted from all .frm files

**Key Functions**:
```vba
' Enquiry Workflow
Public Function SaveEnquiry(EnquiryForm As Object) As Boolean
Public Function SaveEnquiryAndContinue(EnquiryForm As Object) As Boolean
Public Sub InitializeEnquiryForm(EnquiryForm As Object)
Public Sub ClearEnquiryForm(EnquiryForm As Object)

' Quote Workflow
Public Function ProcessQuote(QuoteForm As Object) As Boolean
Public Sub InitializeQuoteForm(QuoteForm As Object, EnquiryNumber As String)

' Job Workflow
Public Function AcceptQuote(JobForm As Object) As Boolean
Public Sub InitializeJobForm(JobForm As Object, QuoteNumber As String)

' Form Validation
Private Function ValidateEnquiryFormData(EnquiryForm As Object) As Boolean
Private Function ValidateQuoteFormData(QuoteForm As Object) As Boolean

' Customer Operations
Public Function CreateCustomerFromForm(FormObject As Object) As Boolean
```

**Dependencies**: SystemCore (data types, validation), BusinessLogic (business processes)

---

### **ReportingSystem.bas** (Reports and Analytics)

**Primary Responsibility**: WIP reports and system analytics

**Key Components**:
- **WIP Report Generation**: Operation and operator reports
- **Data Export Operations**: Export functionality for various formats
- **Report Validation**: Ensure report data integrity
- **System Analytics**: Generate system usage statistics

**Replaces Original Modules**:
- WIP reporting modules
- System reporting functions

**Key Functions**:
```vba
' Report Generation
Public Function GenerateWIPReports(ReportType As String, SortBy As String) As Boolean
Public Function GenerateOperationReport() As Boolean
Public Function GenerateOperatorReport() As Boolean

' Data Export
Public Function ExportWIPData(FilePath As String, Format As String) As Boolean

' Report Validation
Public Function ValidateReportParameters(ReportForm As Object) As Boolean
```

**Dependencies**: SystemCore (validation), DataOperations (file operations)

---

### **UserInterface.bas** (Interface Management)

**Primary Responsibility**: Main interface management, navigation, and application lifecycle

**Key Components**:
- **Application Lifecycle**: System startup, shutdown, initialization
- **Main Interface Management**: File listing, status updates, navigation
- **Form Coordination**: Launch and coordinate between forms
- **System Validation**: Monitor system health and file integrity

**Replaces Original Modules**:
- Main interface modules
- Navigation controls
- File listing systems
- a_Main.bas (system entry point)

**Key Functions**:
```vba
' Application Lifecycle
Public Sub ShowMenu()
Public Sub InitializeMainInterface(MainForm As Object)
Public Sub ShutdownSystem()

' Navigation Management
Public Sub AddEnquiry(MainForm As Object)
Public Sub ShowEnquiries(MainForm As Object)
Public Sub ShowQuotes(MainForm As Object)
Public Sub ShowWIPFiles(MainForm As Object)
Public Sub ShowArchiveFiles(MainForm As Object)

' Form Coordination
Public Sub AcceptQuote(MainForm As Object)
Public Sub CloseJob(MainForm As Object) As Boolean
Public Sub EditJobCard(MainForm As Object)

' File Management
Public Sub RefreshFileList(MainForm As Object)
Public Sub UpdateFileCounters(MainForm As Object)
```

**Dependencies**: All other modules (coordinates entire system)

---

## Form Architecture (Thin Wrapper Pattern)

### **Form Responsibilities**

All forms (.frm files) have been refactored to act as thin wrappers that delegate business logic to appropriate modules:

**Main.frm**:
- **Responsibility**: UI event handling only
- **Delegation**: All business logic → UserInterface module
- **Pattern**: Each button click calls corresponding UserInterface function

**FEnquiry.frm**:
- **Responsibility**: Enquiry form UI events
- **Delegation**: All business logic → WorkflowManagement module
- **Key Delegations**:
  - SaveQ_Click() → WorkflowManagement.SaveEnquiry()
  - AddMore_Click() → WorkflowManagement.SaveEnquiryAndContinue()
  - UserForm_Initialize() → WorkflowManagement.InitializeEnquiryForm()

**FQuote.frm**:
- **Responsibility**: Quote form UI events
- **Delegation**: All business logic → WorkflowManagement module

**FJobCard.frm**:
- **Responsibility**: Job card UI events
- **Delegation**: All business logic → WorkflowManagement module

**fwip.frm**:
- **Responsibility**: WIP report UI events
- **Delegation**: All business logic → ReportingSystem module

### **Form-Module Interaction Pattern**

```vba
' Standard Form Event Pattern
Private Sub ButtonName_Click()
    On Error GoTo Error_Handler

    ModuleName.AppropriateFunction Me
    Exit Sub

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "ButtonName_Click", "FormName"
End Sub
```

---

## Module Dependencies

### **Dependency Hierarchy**

```
SystemCore (Base - No dependencies)
    ↓
DataOperations (Depends on: SystemCore)
    ↓
BusinessLogic (Depends on: SystemCore, DataOperations)
    ↓
WorkflowManagement (Depends on: SystemCore, BusinessLogic)
    ↓
ReportingSystem (Depends on: SystemCore, DataOperations)
    ↓
UserInterface (Depends on: All other modules)
```

### **Cross-Module Communication**

- **SystemCore**: Provides foundation services to all modules
- **DataOperations**: Called by BusinessLogic, WorkflowManagement, ReportingSystem, UserInterface
- **BusinessLogic**: Called by WorkflowManagement and UserInterface
- **WorkflowManagement**: Called by UserInterface and form events
- **ReportingSystem**: Called by UserInterface
- **UserInterface**: Orchestrates all system operations

---

## 32/64-bit Compatibility Strategy

### **SystemCore.bas vs SystemCore32.bas**

The system maintains two identical versions differing only in Windows API declarations:

**SystemCore.bas (64-bit Excel)**:
```vba
Private Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
    (ByVal lpBuffer As String, ByRef nSize As Long) As Long
```

**SystemCore32.bas (32-bit Excel)**:
```vba
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long
```

### **Deployment Strategy**

- **Two Separate Systems**: One for 32-bit Excel, one for 64-bit Excel
- **Identical Functionality**: All other code identical between versions
- **API Isolation**: Only SystemCore module differs between deployments

---

## Benefits of V2 Architecture

### **Maintainability**
- **20+ scattered modules** consolidated into **6 logical modules**
- Related functions grouped together
- Clear separation of concerns
- Reduced code duplication

### **Reliability**
- Centralized error handling in SystemCore
- Consistent validation patterns across all forms
- Safe file operations in DataOperations
- Comprehensive logging and error recovery

### **Performance**
- Optimized file operations
- Reduced memory footprint through better organization
- Faster form processing through delegation
- Efficient search algorithms in BusinessLogic

### **Code Quality**
- **Thin Wrapper Pattern**: Forms contain only UI event handling
- **Clear Dependencies**: Each module has well-defined dependencies
- **Function Mapping**: Every V2 function maps to original V1 function
- **Legacy Compatibility**: All original workflows preserved exactly