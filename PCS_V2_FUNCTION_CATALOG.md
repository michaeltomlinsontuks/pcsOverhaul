# PCS V2 Function Catalog

## Overview

This comprehensive function catalog documents all public functions in the PCS V2 system. Functions are organized by module and functional area, providing exact VBA signatures, parameters, return types, dependencies, and usage patterns.

## Module 1: SystemCore.bas - Core Infrastructure

### **Windows API Functions**

#### `Get_User_Name() As String`
**Purpose**: Get current Windows username using Windows API
**Parameters**: None
**Returns**: String - Current Windows username
**Dependencies**: Windows API (advapi32.dll)
**Side Effects**: None
**Errors**: Returns "Unknown" if API call fails
**Usage**:
```vba
Dim currentUser As String
currentUser = SystemCore.Get_User_Name()
```

#### `GetCurrentUser() As String`
**Purpose**: Alternative user name function with validation
**Parameters**: None
**Returns**: String - Current Windows username (same as Get_User_Name)
**Dependencies**: Get_User_Name()
**Side Effects**: None
**Errors**: Returns "Unknown User" if username cannot be determined

---

### **Error Handling Functions**

#### `LogError(ErrorNumber As Long, ErrorDescription As String, ProcedureName As String, Optional ModuleName As String = "")`
**Purpose**: Centralized error logging with file output
**Parameters**:
- ErrorNumber (Long): The error number that occurred
- ErrorDescription (String): Description of the error
- ProcedureName (String): Name of procedure where error occurred
- ModuleName (String, Optional): Name of module where error occurred
**Returns**: None (Subroutine)
**Dependencies**: File system access for error logging
**Side Effects**: Creates/appends to error_log.txt file
**Usage**:
```vba
SystemCore.LogError Err.Number, Err.Description, "MyFunction", "MyModule"
```

#### `HandleStandardErrors(ErrorNumber As Long, ProcedureName As String, Optional ModuleName As String = "") As Boolean`
**Purpose**: Handle common errors with user-friendly messages and logging
**Parameters**:
- ErrorNumber (Long): The error number to handle
- ProcedureName (String): Name of procedure where error occurred
- ModuleName (String, Optional): Name of module where error occurred
**Returns**: Boolean - True if error was handled, False if unrecognized
**Dependencies**: LogError, MsgBox
**Side Effects**: Logs error to file, displays user message

#### `ClearError()`
**Purpose**: Clear the current error state
**Parameters**: None
**Returns**: None (Subroutine)
**Dependencies**: VBA Err object
**Side Effects**: Clears Err.Number and Err.Description

#### `GetLastErrorInfo() As String`
**Purpose**: Get formatted string of last error information
**Parameters**: None
**Returns**: String - Formatted error information
**Dependencies**: VBA Err object
**Side Effects**: None
**Errors**: Returns empty string if no error

---

### **User Interface Functions**

#### `ShowWarning(Message As String, Title As String)`
**Purpose**: Display warning message to user
**Parameters**:
- Message (String): Warning message to display
- Title (String): Title for message box
**Returns**: None (Subroutine)
**Dependencies**: VBA MsgBox
**Side Effects**: Displays message box to user

#### `ShowError(Message As String, Title As String)`
**Purpose**: Display error message to user
**Parameters**:
- Message (String): Error message to display
- Title (String): Title for message box
**Returns**: None (Subroutine)
**Dependencies**: VBA MsgBox
**Side Effects**: Displays message box to user

#### `ShowInformation(Message As String, Title As String)`
**Purpose**: Display information message to user
**Parameters**:
- Message (String): Information message to display
- Title (String): Title for message box
**Returns**: None (Subroutine)
**Dependencies**: VBA MsgBox
**Side Effects**: Displays message box to user

#### `ShowQuestion(Message As String, Title As String) As Long`
**Purpose**: Display question to user and get response
**Parameters**:
- Message (String): Question to display
- Title (String): Title for message box
**Returns**: Long - User response (vbYes, vbNo, etc.)
**Dependencies**: VBA MsgBox
**Side Effects**: Displays message box to user
**Errors**: Returns vbNo on error

---

### **System Utility Functions**

#### `ValidateSystemRequirements() As Boolean`
**Purpose**: Validate all system requirements for PCS operation
**Parameters**: None
**Returns**: Boolean - True if all requirements met, False if issues found
**Dependencies**: Excel application object, file system access
**Side Effects**: Logs validation results to error log

#### `GetSystemConfig() As SystemConfig`
**Purpose**: Get system configuration information
**Parameters**: None
**Returns**: SystemConfig - Current system configuration
**Dependencies**: GetCurrentUser, Application object, ThisWorkbook
**Side Effects**: None
**Usage**:
```vba
Dim config As SystemConfig
config = SystemCore.GetSystemConfig()
Debug.Print config.RootPath
```

#### `CleanFileName(FileName As String) As String`
**Purpose**: Clean filename for safe file system usage
**Parameters**:
- FileName (String): Raw filename to clean
**Returns**: String - Cleaned filename safe for file system
**Dependencies**: None
**Side Effects**: None
**Logic**: Removes invalid characters (\/:*?"<>|), limits to 50 characters

---

### **String Processing Functions**

#### `RemoveInvalidCharacters(InputString As String) As String`
**Purpose**: Remove invalid characters from string for data processing
**Parameters**:
- InputString (String): String to process
**Returns**: String - String with invalid characters removed
**Dependencies**: None
**Side Effects**: None
**Logic**: Removes "/" ":" and " " characters using legacy algorithm

#### `FormatDisplayText(InputString As String) As String`
**Purpose**: Format text for display by converting underscores and case changes to spaces
**Parameters**:
- InputString (String): String to format for display
**Returns**: String - Formatted string with improved readability
**Dependencies**: None
**Side Effects**: None
**Logic**: Converts underscores to spaces, adds spaces before uppercase letters

---

### **Directory Operations**

#### `CheckDir(Direc As String)`
**Purpose**: Check if directory exists, create if missing, and change to it
**Parameters**:
- Direc (String): Directory path to check/create
**Returns**: None (Subroutine)
**Dependencies**: VBA Dir, MkDir, ChDir functions
**Side Effects**: Creates directory if missing, changes current directory
**Legacy**: Exact replacement for legacy Check_Dir.bas CheckDir function

---

### **Validation Framework Functions**

#### `ValidateRequired(fieldValue As Variant, fieldName As String, Optional setFocusControl As Object = Nothing) As Boolean`
**Purpose**: Validates required field and shows popup if empty
**Parameters**:
- fieldValue (Variant): Value to validate
- fieldName (String): Display name for field in error message
- setFocusControl (Object, Optional): Control to focus on error
**Returns**: Boolean - True if field has value, False if empty
**Side Effects**: Shows MsgBox popup on validation failure

#### `ValidateNumeric(fieldValue As Variant, fieldName As String, Optional setFocusControl As Object = Nothing) As Boolean`
**Purpose**: Validates numeric field and shows popup if invalid
**Parameters**:
- fieldValue (Variant): Value to validate
- fieldName (String): Display name for field in error message
- setFocusControl (Object, Optional): Control to focus on error
**Returns**: Boolean - True if field is numeric, False if invalid
**Side Effects**: Shows MsgBox popup on validation failure

#### `ValidatePositiveNumber(fieldValue As Variant, fieldName As String, Optional setFocusControl As Object = Nothing) As Boolean`
**Purpose**: Validates positive number field and shows popup if invalid
**Parameters**:
- fieldValue (Variant): Value to validate
- fieldName (String): Display name for field in error message
- setFocusControl (Object, Optional): Control to focus on error
**Returns**: Boolean - True if field is positive number, False if invalid
**Dependencies**: ValidateNumeric

#### `ValidateListSelection(listControl As Object, fieldName As String) As Boolean`
**Purpose**: Validates list selection and shows popup if none selected
**Parameters**:
- listControl (Object): ListBox or ComboBox control to validate
- fieldName (String): Display name for field in error message
**Returns**: Boolean - True if selection made, False if no selection
**Side Effects**: Shows MsgBox popup on validation failure, sets focus to control

#### `ValidateDate(fieldValue As Variant, fieldName As String, Optional setFocusControl As Object = Nothing) As Boolean`
**Purpose**: Validates date field and shows popup if invalid
**Parameters**:
- fieldValue (Variant): Value to validate as date
- fieldName (String): Display name for field in error message
- setFocusControl (Object, Optional): Control to focus on error
**Returns**: Boolean - True if valid date, False if invalid
**Side Effects**: Shows MsgBox popup on validation failure

#### `ValidateFileExists(filePath As String, fieldName As String) As Boolean`
**Purpose**: Validates file exists and shows popup if missing
**Parameters**:
- filePath (String): Full path to file to check
- fieldName (String): Display name for field in error message
**Returns**: Boolean - True if file exists, False if missing
**Side Effects**: Shows MsgBox popup on validation failure

#### `ShowConfirmation(message As String, title As String) As Boolean`
**Purpose**: Shows confirmation dialog with Yes/No options
**Parameters**:
- message (String): Message to display
- title (String): Dialog title
**Returns**: Boolean - True if user clicked Yes, False if No
**Side Effects**: Shows MsgBox confirmation dialog

#### `ValidateSpecialDateCaption(dateCaption As String, fieldName As String) As Boolean`
**Purpose**: Validates special date caption (used in enquiry forms)
**Parameters**:
- dateCaption (String): Caption text to check
- fieldName (String): Display name for field
**Returns**: Boolean - True if user wants to continue without date, False to cancel
**Dependencies**: ShowConfirmation

#### `ValidateReportSelection(ReportForm As Object) As Boolean`
**Purpose**: Validates WIP report form selections
**Parameters**:
- ReportForm (Object): Form containing report selection controls
**Returns**: Boolean - True if at least one report type selected, False if none selected
**Side Effects**: Shows MsgBox popup if no selection made
**Usage**: Used by fwip.frm thin wrapper validation

---

### **Legacy Compatibility Functions**

#### `Remove_Characters(Str As String) As String`
**Purpose**: Legacy wrapper for RemoveInvalidCharacters function
**Parameters**:
- Str (String): String to process
**Returns**: String - String with invalid characters removed
**Dependencies**: RemoveInvalidCharacters

#### `Insert_Characters(Str As String) As String`
**Purpose**: Legacy wrapper for FormatDisplayText function
**Parameters**:
- Str (String): String to format for display
**Returns**: String - Formatted string with improved readability
**Dependencies**: FormatDisplayText

---

## Module 2: DataOperations.bas - Data Access & File Operations

### **File System Operations**

#### `GetRootPath() As String`
**Purpose**: Get PCS system root directory path
**Parameters**: None
**Returns**: String - Root directory path for PCS system
**Dependencies**: ThisWorkbook object
**Side Effects**: None
**Errors**: Returns empty string if workbook path unavailable

#### `ValidateDirectoryStructure() As Boolean`
**Purpose**: Check all required PCS directories exist
**Parameters**: None
**Returns**: Boolean - True if all directories exist, False if any missing
**Dependencies**: DirExists for individual directory checking
**Side Effects**: Logs missing directories to error log
**Required Dirs**: Enquiries, Quotes, WIP, Archive, Contracts, Customers, Templates, Job Templates, images, Backups

#### `CreateDirectoryStructure() As Boolean`
**Purpose**: Create missing PCS directories
**Parameters**: None
**Returns**: Boolean - True if all directories created successfully, False if failed
**Dependencies**: DirExists for checking, MkDir for creation
**Side Effects**: Creates missing directories in file system

#### `DirExists(DirPath As String) As Boolean`
**Purpose**: Check if directory exists
**Parameters**:
- DirPath (String): Directory path to check
**Returns**: Boolean - True if directory exists, False if not
**Dependencies**: VBA Dir function

#### `FileExists(FilePath As String) As Boolean`
**Purpose**: Check if file exists
**Parameters**:
- FilePath (String): File path to check
**Returns**: Boolean - True if file exists, False if not
**Dependencies**: VBA Dir function

#### `GetFileList(DirectoryName As String) As Variant`
**Purpose**: Get array of files in directory
**Parameters**:
- DirectoryName (String): Directory name relative to root
**Returns**: Variant - Array of filenames
**Dependencies**: GetRootPath, Dir function

#### `CountFilesInFolder(FolderPath As String, FilePattern As String) As Long`
**Purpose**: Count matching files in directory
**Parameters**:
- FolderPath (String): Directory path to search
- FilePattern (String): File pattern to match (e.g., "*.xls")
**Returns**: Long - Number of matching files

---

### **Workbook Operations**

#### `SafeOpenWorkbook(FilePath As String) As Workbook`
**Purpose**: Safely open Excel workbook with error handling
**Parameters**:
- FilePath (String): Full path to workbook file
**Returns**: Workbook - Opened workbook object, Nothing if failed
**Dependencies**: Excel Application object
**Side Effects**: Opens Excel workbook
**Usage**:
```vba
Dim wb As Workbook
Set wb = DataOperations.SafeOpenWorkbook("C:\Path\File.xls")
If Not wb Is Nothing Then
    ' Process workbook
    DataOperations.SafeCloseWorkbook wb
End If
```

#### `SafeCloseWorkbook(wb As Workbook, Optional SaveChanges As Boolean = True) As Boolean`
**Purpose**: Safely close workbook with error handling
**Parameters**:
- wb (Workbook): Workbook object to close
- SaveChanges (Boolean, Optional): Whether to save changes (default True)
**Returns**: Boolean - True if closed successfully, False if failed
**Side Effects**: Closes Excel workbook, optionally saves changes

---

### **Data Access Functions**

#### `GetValue(FilePath As String, SheetName As String, CellAddress As String) As Variant`
**Purpose**: Get single cell value from workbook
**Parameters**:
- FilePath (String): Full path to Excel file
- SheetName (String): Worksheet name
- CellAddress (String): Cell address (e.g., "A1")
**Returns**: Variant - Cell value
**Dependencies**: SafeOpenWorkbook, SafeCloseWorkbook

#### `GetValueFromClosedWorkbook(FilePath As String, SheetName As String, CellAddress As String) As Variant`
**Purpose**: Get cell value without opening workbook (using Excel4Macro)
**Parameters**:
- FilePath (String): Full path to Excel file
- SheetName (String): Worksheet name
- CellAddress (String): Cell address
**Returns**: Variant - Cell value
**Dependencies**: ExecuteExcel4Macro function

#### `GetRangeData(FilePath As String, SheetName As String, RangeAddress As String) As Variant`
**Purpose**: Get range data from workbook
**Parameters**:
- FilePath (String): Full path to Excel file
- SheetName (String): Worksheet name
- RangeAddress (String): Range address (e.g., "A1:C10")
**Returns**: Variant - Range values as array

---

### **Number Generation Functions**

#### `GetNextEnquiryNumber() As String`
**Purpose**: Get next enquiry number in E00001 format
**Parameters**: None
**Returns**: String - Next enquiry number
**Dependencies**: Number tracking system
**Side Effects**: Reserves number for use

#### `GetNextQuoteNumber() As String`
**Purpose**: Get next quote number in Q00001 format
**Parameters**: None
**Returns**: String - Next quote number
**Dependencies**: Number tracking system
**Side Effects**: Reserves number for use

#### `GetNextJobNumber() As String`
**Purpose**: Get next job number in J00001 format
**Parameters**: None
**Returns**: String - Next job number
**Dependencies**: Number tracking system
**Side Effects**: Reserves number for use

#### `Calc_Next_Number(Typ As String) As Variant`
**Purpose**: Calculate next number from templates (legacy compatibility)
**Parameters**:
- Typ (String): Type identifier ("E", "Q", "J")
**Returns**: Variant - Next number for type
**Dependencies**: Template directory scanning
**Legacy**: Maintains V1 compatibility

#### `Confirm_Next_Number(Typ As String) As Variant`
**Purpose**: Confirm and update template file (legacy compatibility)
**Parameters**:
- Typ (String): Type identifier
**Returns**: Variant - Confirmed number
**Side Effects**: Updates template file, deletes old template, creates new one

---

### **Data Utilities Functions**

#### `GetComponentCodes() As Variant`
**Purpose**: Get component codes from template
**Parameters**: None
**Returns**: Variant - Array of component codes
**Dependencies**: Component template file access

#### `GetMaterialGrades() As Variant`
**Purpose**: Get material grades from template
**Parameters**: None
**Returns**: Variant - Array of material grades
**Dependencies**: Material grades template file access

#### `GetCustomerList() As Variant`
**Purpose**: Get customer list from customer directory
**Parameters**: None
**Returns**: Variant - Array of customer names
**Dependencies**: Customers directory access

#### `GetComponentPrice(ComponentCode As String) As Variant`
**Purpose**: Look up component pricing from price list
**Parameters**:
- ComponentCode (String): Component code to look up
**Returns**: Variant - Component price, Empty if not found
**Dependencies**: Price list file access

---

### **Legacy Compatibility Functions**

#### `OpenBook(File As String, RO As Boolean)`
**Purpose**: Open workbook (exact legacy signature)
**Parameters**:
- File (String): Filename to open
- RO (Boolean): Read-only flag
**Returns**: None (Subroutine)
**Dependencies**: Excel Application
**Legacy**: Maintains exact V1 function signature

---

## Module 3: BusinessLogic.bas - Business Process Management

### **Enquiry Management Functions**

#### `CreateEnquiry(ByRef EnquiryInfo As EnquiryData) As Boolean`
**Purpose**: Create new enquiry following PCS business rules
**Parameters**:
- EnquiryInfo (EnquiryData): Complete enquiry information structure
**Returns**: Boolean - True if enquiry created successfully, False if failed
**Dependencies**: DataOperations.GetNextEnquiryNumber, DataOperations.SafeOpenWorkbook, UpdateSearchDatabase
**Side Effects**: Creates new enquiry Excel file in Enquiries directory, updates search database
**Business Logic**:
1. Validates enquiry data
2. Generates enquiry number
3. Creates file from template
4. Updates search database
5. Creates customer record if new

#### `LoadEnquiry(FilePath As String) As EnquiryData`
**Purpose**: Load enquiry data from file
**Parameters**:
- FilePath (String): Full path to enquiry file
**Returns**: EnquiryData - Populated enquiry structure, empty if failed
**Dependencies**: DataOperations.SafeOpenWorkbook
**Side Effects**: Opens and closes enquiry file

#### `UpdateEnquiry(ByRef EnquiryInfo As EnquiryData) As Boolean`
**Purpose**: Update existing enquiry with new data
**Parameters**:
- EnquiryInfo (EnquiryData): Updated enquiry information
**Returns**: Boolean - True if update successful, False if failed
**Dependencies**: DataOperations.SafeOpenWorkbook, PopulateEnquiryTemplate
**Side Effects**: Modifies enquiry file, saves changes, updates search database

#### `ValidateEnquiryData(ByRef EnquiryInfo As EnquiryData) As String`
**Purpose**: Validate enquiry data completeness and business rules
**Parameters**:
- EnquiryInfo (EnquiryData): Enquiry data to validate
**Returns**: String - Validation error messages, empty if valid
**Business Rules**:
- CustomerName and ComponentDescription required
- Quantity must be > 0
- Email format validation if provided

#### `CreateNewCustomer(CustomerName As String) As Boolean`
**Purpose**: Create new customer record file
**Parameters**:
- CustomerName (String): Customer name for new record
**Returns**: Boolean - True if customer created successfully, False if failed
**Dependencies**: Customer template file, DataOperations
**Side Effects**: Creates new customer file in Customers directory

---

### **Quote Management Functions**

#### `CreateQuote(ByRef QuoteInfo As QuoteData) As Boolean`
**Purpose**: Create quote from enquiry data
**Parameters**:
- QuoteInfo (QuoteData): Complete quote information structure
**Returns**: Boolean - True if quote created successfully, False if failed
**Dependencies**: DataOperations.GetNextQuoteNumber, template operations
**Side Effects**: Creates quote file, moves enquiry file, updates search database

#### `LoadQuote(FilePath As String) As QuoteData`
**Purpose**: Load quote data from file
**Parameters**:
- FilePath (String): Full path to quote file
**Returns**: QuoteData - Populated quote structure
**Dependencies**: DataOperations.SafeOpenWorkbook

#### `ValidateQuoteData(ByRef QuoteInfo As QuoteData) As String`
**Purpose**: Validate quote data completeness
**Parameters**:
- QuoteInfo (QuoteData): Quote data to validate
**Returns**: String - Validation error messages, empty if valid
**Business Rules**:
- Inherits enquiry validation rules
- UnitPrice must be > 0
- ValidUntil date must be future date
- TotalPrice automatically calculated

---

### **Job Management Functions**

#### `CreateJobFromQuote(ByRef QuoteInfo As QuoteData, ByRef JobInfo As JobData) As Boolean`
**Purpose**: Create job from accepted quote
**Parameters**:
- QuoteInfo (QuoteData): Source quote information
- JobInfo (JobData): Job information to populate and create
**Returns**: Boolean - True if job created successfully, False if failed
**Dependencies**: DataOperations.GetNextJobNumber, template operations
**Side Effects**: Creates job file, moves quote file, updates WIP database

#### `LoadJob(FilePath As String) As JobData`
**Purpose**: Load job data from file
**Parameters**:
- FilePath (String): Full path to job file
**Returns**: JobData - Populated job structure

#### `ValidateJobData(ByRef JobInfo As JobData) As String`
**Purpose**: Validate job data completeness
**Parameters**:
- JobInfo (JobData): Job data to validate
**Returns**: String - Validation error messages, empty if valid
**Business Rules**:
- Must reference valid quote
- CustomerOrderNumber required
- DueDate must be future date
- OrderValue must be > 0

#### `CloseJob(JobNumber As String) As Boolean`
**Purpose**: Close completed job and move to archive
**Parameters**:
- JobNumber (String): Job number to close (J00001 format)
**Returns**: Boolean - True if job closed successfully, False if failed
**Side Effects**: Moves job file from WIP to Archive, updates databases

---

### **Search System Functions**

#### `SearchRecords(SearchTerm As String, Optional RecordTypeFilter As String = "") As Variant`
**Purpose**: Search all PCS records (basic search)
**Parameters**:
- SearchTerm (String): Term to search for
- RecordTypeFilter (String, Optional): Filter by record type ("Enquiry", "Quote", "Job")
**Returns**: Variant - Array of matching search results
**Dependencies**: Search.xls database access

#### `SearchRecords_Optimized(SearchTerm As String, Optional RecordTypeFilter As String = "") As Variant`
**Purpose**: Enhanced search with recent file prioritization
**Parameters**:
- SearchTerm (String): Term to search for
- RecordTypeFilter (String, Optional): Filter by record type
**Returns**: Variant - Array of matching results, recent files prioritized
**Performance**: Optimized search algorithm, faster results

#### `CreateSearchRecord(RecordType As String, RecordNumber As String, CustomerName As String, Description As String, FilePath As String, Optional Keywords As String = "") As SearchRecord`
**Purpose**: Create search record structure
**Parameters**:
- RecordType (String): Type of record ("Enquiry", "Quote", "Job")
- RecordNumber (String): Record number (E00001, Q00001, J00001)
- CustomerName (String): Customer name
- Description (String): Component/description
- FilePath (String): Full path to record file
- Keywords (String, Optional): Additional searchable keywords
**Returns**: SearchRecord - Populated search record structure

#### `UpdateSearchDatabase(ByRef SearchRecord As SearchRecord) As Boolean`
**Purpose**: Update search database with new/modified record
**Parameters**:
- SearchRecord (SearchRecord): Search record to add/update
**Returns**: Boolean - True if update successful, False if failed
**Dependencies**: Search.xls file access
**Side Effects**: Updates or adds record to search database, sorts by date

#### `SaveRowIntoSearch(ByRef FormObject As Object) As Boolean`
**Purpose**: Save form data to search database (legacy compatibility)
**Parameters**:
- FormObject (Object): Form containing data to save to search
**Returns**: Boolean - True if save successful, False if failed
**Legacy**: Maintains compatibility with original form integration

---

## Module 4: WorkflowManagement.bas - Document Lifecycle Management

### **Enquiry Workflow Functions**

#### `SaveEnquiryAndContinue(EnquiryForm As Object) As Boolean`
**Purpose**: Save enquiry and prepare for new one
**Parameters**:
- EnquiryForm (Object): Form containing enquiry data
**Returns**: Boolean - True if save successful and ready for new enquiry, False if failed
**Dependencies**: SaveEnquiry, ClearEnquiryForm, InitializeEnquiryForm
**Side Effects**: Saves current enquiry, clears form for new enquiry
**Usage**: Used by FEnquiry.AddMore_Click() event

#### `SaveEnquiry(EnquiryForm As Object) As Boolean`
**Purpose**: Save current enquiry from form
**Parameters**:
- EnquiryForm (Object): Form containing enquiry data
**Returns**: Boolean - True if save successful, False if failed
**Dependencies**: ValidateEnquiryFormData, BusinessLogic.CreateEnquiry
**Side Effects**: Creates new enquiry file, updates search database
**Data Mapping**:
```vba
With EnquiryInfo
    .CustomerName = Trim(EnquiryForm.Customer.Value)
    .ContactPerson = Trim(EnquiryForm.Contact.Value)
    .ComponentDescription = Trim(EnquiryForm.Component_Description.Value)
    .Quantity = CLng(EnquiryForm.Component_Quantity.Value)
End With
```

#### `CreateCustomerFromForm(EnquiryForm As Object) As Boolean`
**Purpose**: Create customer record from enquiry form
**Parameters**:
- EnquiryForm (Object): Form containing customer data
**Returns**: Boolean - True if customer created successfully, False if failed
**Dependencies**: BusinessLogic.CreateNewCustomer
**Side Effects**: Creates new customer file in Customers directory

#### `SetEnquiryDate(EnquiryForm As Object)`
**Purpose**: Set enquiry date to current date
**Parameters**:
- EnquiryForm (Object): Form containing date control
**Returns**: None (Subroutine)
**Side Effects**: Updates date control with formatted current date
**Format**: "dd mmm yyyy" (e.g., "25 Sep 2023")

#### `InitializeEnquiryForm(EnquiryForm As Object)`
**Purpose**: Initialize enquiry form with default values and dropdowns
**Parameters**:
- EnquiryForm (Object): Form to initialize
**Returns**: None (Subroutine)
**Dependencies**: DataOperations.GetComponentCodes, DataOperations.GetMaterialGrades
**Side Effects**: Populates form dropdowns, sets default values
**Initialization**:
- Sets current date
- Loads component codes dropdown
- Loads material grades dropdown
- Sets default quantities and values

---

### **Quote Workflow Functions**

#### `SaveQuote(QuoteForm As Object) As Boolean`
**Purpose**: Save quote from form data
**Parameters**:
- QuoteForm (Object): Form containing quote data
**Returns**: Boolean - True if save successful, False if failed
**Dependencies**: BusinessLogic.CreateQuote
**Side Effects**: Creates quote file, updates search database

#### `CalculateQuoteTotalPrice(QuoteForm As Object)`
**Purpose**: Calculate total price from unit price and quantity
**Parameters**:
- QuoteForm (Object): Form containing pricing controls
**Returns**: None (Subroutine)
**Side Effects**: Updates TotalPrice control on form
**Calculation**: TotalPrice = UnitPrice × Quantity

#### `LoadComponentPricing(QuoteForm As Object)`
**Purpose**: Load pricing from component database
**Parameters**:
- QuoteForm (Object): Form containing component code control
**Returns**: None (Subroutine)
**Dependencies**: DataOperations.GetComponentPrice
**Side Effects**: Updates UnitPrice control with looked-up price

#### `SetQuoteValidUntilDate(QuoteForm As Object)`
**Purpose**: Set default validity date (30 days from current date)
**Parameters**:
- QuoteForm (Object): Form containing validity date control
**Returns**: None (Subroutine)
**Side Effects**: Updates ValidUntil control
**Default**: Current date + 30 days

#### `SearchComponentCode(QuoteForm As Object)`
**Purpose**: Search for component codes and populate form
**Parameters**:
- QuoteForm (Object): Form to populate with search results
**Returns**: None (Subroutine)
**Dependencies**: Component search functionality

#### `InitializeQuoteForm(QuoteForm As Object)`
**Purpose**: Initialize quote form with defaults
**Parameters**:
- QuoteForm (Object): Form to initialize
**Returns**: None (Subroutine)
**Initialization**:
- Sets current date
- Sets default validity period (30 days)
- Loads component codes and pricing
- Initializes calculation fields

#### `LoadQuoteFromEnquiry(QuoteForm As Object, EnquiryPath As String)`
**Purpose**: Populate quote form from enquiry data
**Parameters**:
- QuoteForm (Object): Form to populate
- EnquiryPath (String): Path to source enquiry file
**Returns**: None (Subroutine)
**Dependencies**: BusinessLogic.LoadEnquiry
**Side Effects**: Loads all enquiry data into quote form controls
**Data Transfer**: Copies customer, component, and contact information

---

### **Job Workflow Functions**

#### `AcceptQuote(JobForm As Object, QuotePath As String) As Boolean`
**Purpose**: Accept quote and convert to job
**Parameters**:
- JobForm (Object): Form containing job acceptance data
- QuotePath (String): Path to quote file being accepted
**Returns**: Boolean - True if quote accepted and job created, False if failed
**Dependencies**: BusinessLogic.CreateJobFromQuote
**Side Effects**: Creates job file, moves quote file, updates WIP database
**Required Data**: Customer order number, urgency level, delivery requirements

#### `LoadQuoteForAcceptance(JobForm As Object, QuotePath As String)`
**Purpose**: Load quote data into job acceptance form
**Parameters**:
- JobForm (Object): Form to populate
- QuotePath (String): Path to quote file
**Returns**: None (Subroutine)
**Dependencies**: BusinessLogic.LoadQuote
**Side Effects**: Populates job form with quote data

#### `SaveJobCard(JobCardForm As Object, CurrentJobPath As String) As Boolean`
**Purpose**: Save job card with form data
**Parameters**:
- JobCardForm (Object): Job card form containing production data
- CurrentJobPath (String): Path to current job file
**Returns**: Boolean - True if save successful, False if failed
**Side Effects**: Updates job file with production information
**Production Data**: Operations, operators, schedules, progress tracking

---

## Module 5: ReportingSystem.bas - Reports & Analytics

### **WIP Reporting Functions**

#### `GenerateWIPReports(ReportForm As Object) As Boolean`
**Purpose**: Generate WIP reports based on form selections
**Parameters**:
- ReportForm (Object): Form containing report configuration
**Returns**: Boolean - True if reports generated successfully, False if failed
**Dependencies**: WIP.xls database, Excel operations
**Side Effects**: Creates new workbook with formatted reports
**Report Types**:
- Operation reports (grouped by operation type)
- Operator reports (grouped by assigned operator)
- Due date reports (sorted by delivery dates)
- Customer reports (sorted by customer)

#### `ExportWIPData(Optional ExportPath As String = "") As Boolean`
**Purpose**: Export WIP data to Excel for external analysis
**Parameters**:
- ExportPath (String, Optional): Path for export file (default: generated path)
**Returns**: Boolean - True if export successful, False if failed
**Dependencies**: WIP.xls access, Excel operations
**Side Effects**: Creates Excel export file with WIP data

---

## Module 6: UserInterface.bas - UI Management & Application Lifecycle

### **Application Lifecycle Functions**

#### `ShowMenu()`
**Purpose**: Main system entry point - shows PCS main menu
**Parameters**: None
**Returns**: None (Subroutine)
**Dependencies**: Main.frm, system initialization
**Side Effects**: Displays main PCS interface, initializes system
**Usage**: Primary entry point called from a_Main.bas

#### `InitializeApplication() As Boolean`
**Purpose**: Initialize PCS application and validate system readiness
**Parameters**: None
**Returns**: Boolean - True if initialization successful, False if failed
**Dependencies**: DataOperations.ValidateDirectoryStructure, SystemCore validation
**Side Effects**: Validates system requirements, creates missing directories

#### `RefreshMainInterface() As Boolean`
**Purpose**: Refresh main form UI with current data
**Parameters**: None
**Returns**: Boolean - True if refresh successful, False if failed
**Side Effects**: Updates file lists, counters, and status indicators

---

### **Main Interface Management Functions**

#### `InitializeMainInterface(MainForm As Object)`
**Purpose**: Initialize main form interface
**Parameters**:
- MainForm (Object): Main form object to initialize
**Returns**: None (Subroutine)
**Side Effects**: Sets up form controls, loads initial data, sets event handlers

#### `AddEnquiry(MainForm As Object)`
**Purpose**: Open enquiry form to add new enquiry
**Parameters**:
- MainForm (Object): Main form calling this function
**Returns**: None (Subroutine)
**Side Effects**: Displays enquiry form (FEnquiry or FrmEnquiry)

#### `ShowArchiveFiles(MainForm As Object)`
**Purpose**: Show archive files in main list
**Parameters**:
- MainForm (Object): Main form to update
**Returns**: None (Subroutine)
**Dependencies**: DataOperations.GetFileList
**Side Effects**: Populates main list with archive files

#### `ShowEnquiries(MainForm As Object)`
**Purpose**: Show enquiry files in main list
**Parameters**:
- MainForm (Object): Main form to update
**Returns**: None (Subroutine)
**Side Effects**: Populates main list with enquiry files

#### `ShowQuotes(MainForm As Object)`
**Purpose**: Show quote files in main list
**Parameters**:
- MainForm (Object): Main form to update
**Returns**: None (Subroutine)
**Side Effects**: Populates main list with quote files

#### `ShowWIPFiles(MainForm As Object)`
**Purpose**: Show WIP files in main list
**Parameters**:
- MainForm (Object): Main form to update
**Returns**: None (Subroutine)
**Side Effects**: Populates main list with work-in-progress files

#### `AcceptQuote(MainForm As Object)`
**Purpose**: Accept selected quote and convert to job
**Parameters**:
- MainForm (Object): Main form with selected quote
**Returns**: None (Subroutine)
**Dependencies**: FAcceptQuote form, selected quote file
**Side Effects**: Opens quote acceptance form with selected quote data

#### `CloseJob(MainForm As Object) As Boolean`
**Purpose**: Close selected job
**Parameters**:
- MainForm (Object): Main form with selected job
**Returns**: Boolean - True if job closed successfully, False if failed
**Dependencies**: BusinessLogic.CloseJob
**Side Effects**: Moves job from WIP to Archive, updates databases

---

### **Specialized Management Functions**

#### `CreateContractTemplateItem()`
**Purpose**: Create new contract template
**Parameters**: None
**Returns**: None (Subroutine)
**Side Effects**: Opens contract template creation interface

#### `EditContractTemplateItem(MainForm As Object)`
**Purpose**: Edit existing contract template
**Parameters**:
- MainForm (Object): Main form with selected template
**Returns**: None (Subroutine)
**Side Effects**: Opens contract template for editing

#### `EditJobCard(MainForm As Object)`
**Purpose**: Open job card editing form
**Parameters**:
- MainForm (Object): Main form with selected job
**Returns**: None (Subroutine)
**Dependencies**: FJobCard form
**Side Effects**: Opens job card form with selected job data

#### `EditSearchDatabase(MainForm As Object)`
**Purpose**: Open search database for editing
**Parameters**:
- MainForm (Object): Main form calling this function
**Returns**: None (Subroutine)
**Dependencies**: Search.xls file access
**Side Effects**: Opens search database in Excel for manual editing

#### `ShowSearchHistory(MainForm As Object)`
**Purpose**: Display search history
**Parameters**:
- MainForm (Object): Main form calling this function
**Returns**: None (Subroutine)
**Dependencies**: Search History.xls file
**Side Effects**: Opens search history for review

#### `SortSearchDatabase()`
**Purpose**: Sort search database records
**Parameters**: None
**Returns**: None (Subroutine)
**Dependencies**: Search.xls file access
**Side Effects**: Sorts search database by date and customer name

---

## Function Usage Patterns

### **Error Handling Pattern**
Most functions follow this standard error handling pattern:
```vba
Public Function FunctionName() As Boolean
    On Error GoTo Error_Handler

    ' Function logic here

    FunctionName = True  ' Success
    Exit Function

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "FunctionName", "ModuleName"
    FunctionName = False  ' Failure
End Function
```

### **Form Delegation Pattern**
Forms delegate to modules using this pattern:
```vba
' In Form
Private Sub Button_Click()
    On Error GoTo Error_Handler

    ModuleName.AppropriateFunction Me
    Exit Sub

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "Button_Click", "FormName"
End Sub
```

### **Data Validation Pattern**
Data validation follows this pattern:
```vba
Public Function ValidateData(DataStructure As DataType) As String
    Dim ValidationErrors As String

    If [condition] Then
        ValidationErrors = ValidationErrors & "Error message." & vbCrLf
    End If

    ValidateData = ValidationErrors  ' Empty string = valid
End Function
```

### **File Operation Pattern**
Safe file operations follow this pattern:
```vba
Public Function FileOperation() As Boolean
    Dim wb As Workbook

    On Error GoTo Error_Handler

    Set wb = DataOperations.SafeOpenWorkbook(FilePath)
    If wb Is Nothing Then Exit Function

    ' Process file

    DataOperations.SafeCloseWorkbook wb
    FileOperation = True
    Exit Function

Error_Handler:
    If Not wb Is Nothing Then DataOperations.SafeCloseWorkbook wb, False
    SystemCore.HandleStandardErrors Err.Number, "FileOperation", "ModuleName"
    FileOperation = False
End Function
```

This function catalog provides complete reference documentation for all public functions in the PCS V2 system, enabling efficient system maintenance, development, and troubleshooting.