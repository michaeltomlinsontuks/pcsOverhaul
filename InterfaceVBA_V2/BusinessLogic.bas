Attribute VB_Name = "BusinessLogic"
' **Purpose**: Core business process controllers, workflow management, and search functionality
' **CLAUDE.md Compliance**: Maintains Enquiry → Quote → Jobs workflow, preserves all business logic
' **Consolidation**: Combines BusinessController.bas and SearchManager.bas
Option Explicit

' ===================================================================
' CONSTANTS AND PRIVATE VARIABLES
' ===================================================================

Private Const WIP_FILE As String = "WIP.xls"
Private Const SEARCH_FILE As String = "Search.xls"
Private Const SEARCH_HISTORY_FILE As String = "Search History.xls"
Private Const SYNC_PASSWORD As String = "KJB"

' ===================================================================
' ENQUIRY MANAGEMENT (CLAUDE.md: Preserve Enquiry → Quote → Jobs workflow)
' ===================================================================

' **Purpose**: Create new enquiry following PCS business rules
' **Parameters**:
'   - EnquiryInfo (EnquiryData): Complete enquiry information structure
' **Returns**: Boolean - True if enquiry created successfully, False if failed
' **Dependencies**: DataOperations.GetNextEnquiryNumber, DataOperations.SafeOpenWorkbook, UpdateSearchDatabase
' **Side Effects**: Creates new enquiry Excel file in Enquiries directory, updates search database
' **Errors**: Returns False on template missing, file creation failure, or validation errors
' **CLAUDE.md Compliance**: Maintains Enquiry → Quote → Jobs workflow
Public Function CreateEnquiry(ByRef EnquiryInfo As SystemCore.EnquiryData) As Boolean
    Dim EnquiryNumber As String
    Dim TemplatePath As String
    Dim NewFilePath As String
    Dim TemplateWB As Workbook
    Dim SearchRecord As SystemCore.SearchRecord

    On Error GoTo Error_Handler

    ' Validate enquiry data before processing
    If ValidateEnquiryData(EnquiryInfo) <> "" Then
        CreateEnquiry = False
        Exit Function
    End If

    EnquiryNumber = DataOperations.GetNextEnquiryNumber()
    If EnquiryNumber = "" Then
        CreateEnquiry = False
        Exit Function
    End If

    EnquiryInfo.EnquiryNumber = EnquiryNumber
    EnquiryInfo.DateCreated = Now

    TemplatePath = DataOperations.GetRootPath & "\Templates\_Enq.xls"
    NewFilePath = DataOperations.GetRootPath & "\Enquiries\" & EnquiryNumber & ".xls"

    If Not DataOperations.FileExists(TemplatePath) Then
        SystemCore.LogError SystemCore.ERR_FILE_NOT_FOUND, "Enquiry template not found: " & TemplatePath, "CreateEnquiry", "BusinessLogic"
        CreateEnquiry = False
        Exit Function
    End If

    Set TemplateWB = DataOperations.SafeOpenWorkbook(TemplatePath)
    If TemplateWB Is Nothing Then
        CreateEnquiry = False
        Exit Function
    End If

    PopulateEnquiryTemplate TemplateWB, EnquiryInfo

    TemplateWB.SaveAs NewFilePath
    DataOperations.SafeCloseWorkbook TemplateWB

    EnquiryInfo.FilePath = NewFilePath

    ' Update search database
    SearchRecord = CreateSearchRecord(SystemCore.rtEnquiry, EnquiryNumber, EnquiryInfo.CustomerName, EnquiryInfo.ComponentDescription, NewFilePath, EnquiryInfo.SearchKeywords)
    UpdateSearchDatabase SearchRecord

    ' Create customer record if new
    If Not DataOperations.FileExists(DataOperations.GetRootPath & "\Customers\" & SystemCore.CleanFileName(EnquiryInfo.CustomerName) & ".xls") Then
        CreateNewCustomer EnquiryInfo.CustomerName
    End If

    CreateEnquiry = True
    Exit Function

Error_Handler:
    If Not TemplateWB Is Nothing Then DataOperations.SafeCloseWorkbook TemplateWB, False
    SystemCore.HandleStandardErrors Err.Number, "CreateEnquiry", "BusinessLogic"
    CreateEnquiry = False
End Function

' **Purpose**: Load enquiry data from file
' **Parameters**:
'   - FilePath (String): Full path to enquiry file
' **Returns**: EnquiryData - Populated enquiry structure, empty if failed
' **Dependencies**: DataOperations.SafeOpenWorkbook for file access
' **Side Effects**: Opens and closes enquiry file
' **Errors**: Returns empty structure if file access fails
Public Function LoadEnquiry(ByVal FilePath As String) As SystemCore.EnquiryData
    Dim EnquiryWB As Workbook
    Dim ws As Worksheet
    Dim EnquiryInfo As SystemCore.EnquiryData

    On Error GoTo Error_Handler

    Set EnquiryWB = DataOperations.SafeOpenWorkbook(FilePath)
    If EnquiryWB Is Nothing Then
        Exit Function
    End If

    Set ws = EnquiryWB.Worksheets(1)

    With EnquiryInfo
        .EnquiryNumber = ws.Cells(2, 2).Value
        .CustomerName = ws.Cells(3, 2).Value
        .ContactPerson = ws.Cells(4, 2).Value
        .CompanyPhone = ws.Cells(5, 2).Value
        .CompanyFax = ws.Cells(6, 2).Value
        .Email = ws.Cells(7, 2).Value
        .ComponentDescription = ws.Cells(8, 2).Value
        .ComponentCode = ws.Cells(9, 2).Value
        .MaterialGrade = ws.Cells(10, 2).Value
        .Quantity = ws.Cells(11, 2).Value
        .DateCreated = ws.Cells(12, 2).Value
        .FilePath = FilePath
        .SearchKeywords = .CustomerName & " " & .ComponentDescription & " " & .ComponentCode
    End With

    DataOperations.SafeCloseWorkbook EnquiryWB, False
    LoadEnquiry = EnquiryInfo
    Exit Function

Error_Handler:
    If Not EnquiryWB Is Nothing Then DataOperations.SafeCloseWorkbook EnquiryWB, False
    SystemCore.HandleStandardErrors Err.Number, "LoadEnquiry", "BusinessLogic"
End Function

' **Purpose**: Update existing enquiry with new data
' **Parameters**:
'   - EnquiryInfo (EnquiryData): Updated enquiry information
' **Returns**: Boolean - True if update successful, False if failed
' **Dependencies**: DataOperations.SafeOpenWorkbook, PopulateEnquiryTemplate
' **Side Effects**: Modifies enquiry file, saves changes
' **Errors**: Returns False if file access or update fails
Public Function UpdateEnquiry(ByRef EnquiryInfo As SystemCore.EnquiryData) As Boolean
    Dim EnquiryWB As Workbook

    On Error GoTo Error_Handler

    ' Validate data before updating
    If ValidateEnquiryData(EnquiryInfo) <> "" Then
        UpdateEnquiry = False
        Exit Function
    End If

    Set EnquiryWB = DataOperations.SafeOpenWorkbook(EnquiryInfo.FilePath)
    If EnquiryWB Is Nothing Then
        UpdateEnquiry = False
        Exit Function
    End If

    PopulateEnquiryTemplate EnquiryWB, EnquiryInfo

    EnquiryWB.Save
    DataOperations.SafeCloseWorkbook EnquiryWB

    ' Update search database
    Dim SearchRecord As SystemCore.SearchRecord
    SearchRecord = CreateSearchRecord(SystemCore.rtEnquiry, EnquiryInfo.EnquiryNumber, EnquiryInfo.CustomerName, EnquiryInfo.ComponentDescription, EnquiryInfo.FilePath, EnquiryInfo.SearchKeywords)
    UpdateSearchDatabase SearchRecord

    UpdateEnquiry = True
    Exit Function

Error_Handler:
    If Not EnquiryWB Is Nothing Then DataOperations.SafeCloseWorkbook EnquiryWB, False
    SystemCore.HandleStandardErrors Err.Number, "UpdateEnquiry", "BusinessLogic"
    UpdateEnquiry = False
End Function

' **Purpose**: Validate enquiry data completeness and business rules
' **Parameters**:
'   - EnquiryInfo (EnquiryData): Enquiry data to validate
' **Returns**: String - Validation error messages, empty if valid
' **Dependencies**: None
' **Side Effects**: None
' **Errors**: Returns error descriptions for invalid data
Public Function ValidateEnquiryData(ByRef EnquiryInfo As SystemCore.EnquiryData) As String
    Dim ValidationErrors As String

    If Trim(EnquiryInfo.CustomerName) = "" Then
        ValidationErrors = ValidationErrors & "Customer name is required." & vbCrLf
    End If

    If Trim(EnquiryInfo.ComponentDescription) = "" Then
        ValidationErrors = ValidationErrors & "Component description is required." & vbCrLf
    End If

    If EnquiryInfo.Quantity <= 0 Then
        ValidationErrors = ValidationErrors & "Quantity must be greater than zero." & vbCrLf
    End If

    If Trim(EnquiryInfo.ContactPerson) = "" Then
        ValidationErrors = ValidationErrors & "Contact person is required." & vbCrLf
    End If

    ' Validate email format if provided
    If EnquiryInfo.Email <> "" And InStr(EnquiryInfo.Email, "@") = 0 Then
        ValidationErrors = ValidationErrors & "Invalid email format." & vbCrLf
    End If

    ValidateEnquiryData = ValidationErrors
End Function

' **Purpose**: Create new customer record file
' **Parameters**:
'   - CustomerName (String): Name of customer for new record
' **Returns**: Boolean - True if customer created successfully, False if failed
' **Dependencies**: DataOperations.SafeOpenWorkbook for template access
' **Side Effects**: Creates new customer file in Customers directory
' **Errors**: Returns False if template missing or file creation fails
Public Function CreateNewCustomer(ByVal CustomerName As String) As Boolean
    Dim TemplatePath As String
    Dim NewFilePath As String
    Dim TemplateWB As Workbook
    Dim CleanName As String

    On Error GoTo Error_Handler

    CleanName = SystemCore.CleanFileName(CustomerName)
    TemplatePath = DataOperations.GetRootPath & "\Templates\_client.xls"
    NewFilePath = DataOperations.GetRootPath & "\Customers\" & CleanName & ".xls"

    If DataOperations.FileExists(NewFilePath) Then
        CreateNewCustomer = True ' Already exists
        Exit Function
    End If

    If Not DataOperations.FileExists(TemplatePath) Then
        SystemCore.LogError SystemCore.ERR_FILE_NOT_FOUND, "Customer template not found: " & TemplatePath, "CreateNewCustomer", "BusinessLogic"
        CreateNewCustomer = False
        Exit Function
    End If

    Set TemplateWB = DataOperations.SafeOpenWorkbook(TemplatePath)
    If TemplateWB Is Nothing Then
        CreateNewCustomer = False
        Exit Function
    End If

    TemplateWB.Worksheets(1).Cells(1, 1).Value = CustomerName
    TemplateWB.Worksheets(1).Cells(1, 2).Value = Now ' Creation date

    TemplateWB.SaveAs NewFilePath
    DataOperations.SafeCloseWorkbook TemplateWB

    CreateNewCustomer = True
    Exit Function

Error_Handler:
    If Not TemplateWB Is Nothing Then DataOperations.SafeCloseWorkbook TemplateWB, False
    SystemCore.HandleStandardErrors Err.Number, "CreateNewCustomer", "BusinessLogic"
    CreateNewCustomer = False
End Function

' ===================================================================
' QUOTE MANAGEMENT (CLAUDE.md: Preserve Quote workflow)
' ===================================================================

' **Purpose**: Create quote from existing enquiry
' **Parameters**:
'   - EnquiryInfo (EnquiryData): Source enquiry information
'   - QuoteInfo (QuoteData): Quote information to populate
' **Returns**: Boolean - True if quote created successfully, False if failed
' **Dependencies**: DataOperations.GetNextQuoteNumber, DataOperations.SafeOpenWorkbook
' **Side Effects**: Creates new quote Excel file, updates search database
' **Errors**: Returns False on template missing or file creation failure
' **CLAUDE.md Compliance**: Maintains Enquiry → Quote → Jobs workflow
Public Function CreateQuote(ByRef QuoteInfo As SystemCore.QuoteData) As Boolean
    Dim QuoteNumber As String
    Dim TemplatePath As String
    Dim NewFilePath As String
    Dim TemplateWB As Workbook
    Dim SearchRecord As SystemCore.SearchRecord

    On Error GoTo Error_Handler

    ' Validate quote data before processing
    If ValidateQuoteData(QuoteInfo) <> "" Then
        CreateQuote = False
        Exit Function
    End If

    QuoteNumber = DataOperations.GetNextQuoteNumber()
    If QuoteNumber = "" Then
        CreateQuote = False
        Exit Function
    End If

    QuoteInfo.QuoteNumber = QuoteNumber
    QuoteInfo.DateCreated = Now

    TemplatePath = DataOperations.GetRootPath & "\Templates\_Quote.xls"
    NewFilePath = DataOperations.GetRootPath & "\Quotes\" & QuoteNumber & ".xls"

    If Not DataOperations.FileExists(TemplatePath) Then
        SystemCore.LogError SystemCore.ERR_FILE_NOT_FOUND, "Quote template not found: " & TemplatePath, "CreateQuote", "BusinessLogic"
        CreateQuote = False
        Exit Function
    End If

    Set TemplateWB = DataOperations.SafeOpenWorkbook(TemplatePath)
    If TemplateWB Is Nothing Then
        CreateQuote = False
        Exit Function
    End If

    PopulateQuoteTemplate TemplateWB, QuoteInfo

    TemplateWB.SaveAs NewFilePath
    DataOperations.SafeCloseWorkbook TemplateWB

    QuoteInfo.FilePath = NewFilePath

    ' Update search database
    SearchRecord = CreateSearchRecord(SystemCore.rtQuote, QuoteNumber, QuoteInfo.CustomerName, QuoteInfo.ComponentDescription, NewFilePath, "")
    UpdateSearchDatabase SearchRecord

    CreateQuote = True
    Exit Function

Error_Handler:
    If Not TemplateWB Is Nothing Then DataOperations.SafeCloseWorkbook TemplateWB, False
    SystemCore.HandleStandardErrors Err.Number, "CreateQuote", "BusinessLogic"
    CreateQuote = False
End Function

' **Purpose**: Load quote data from file
' **Parameters**:
'   - FilePath (String): Full path to quote file
' **Returns**: QuoteData - Populated quote structure, empty if failed
' **Dependencies**: DataOperations.SafeOpenWorkbook for file access
' **Side Effects**: Opens and closes quote file
' **Errors**: Returns empty structure if file access fails
Public Function LoadQuote(ByVal FilePath As String) As SystemCore.QuoteData
    Dim QuoteWB As Workbook
    Dim ws As Worksheet
    Dim QuoteInfo As SystemCore.QuoteData

    On Error GoTo Error_Handler

    Set QuoteWB = DataOperations.SafeOpenWorkbook(FilePath)
    If QuoteWB Is Nothing Then
        Exit Function
    End If

    Set ws = QuoteWB.Worksheets(1)

    With QuoteInfo
        .QuoteNumber = ws.Cells(2, 2).Value
        .EnquiryNumber = ws.Cells(3, 2).Value
        .CustomerName = ws.Cells(4, 2).Value
        .ComponentDescription = ws.Cells(5, 2).Value
        .ComponentCode = ws.Cells(6, 2).Value
        .MaterialGrade = ws.Cells(7, 2).Value
        .Quantity = ws.Cells(8, 2).Value
        .UnitPrice = ws.Cells(9, 2).Value
        .TotalPrice = ws.Cells(10, 2).Value
        .LeadTime = ws.Cells(11, 2).Value
        .ValidUntil = ws.Cells(12, 2).Value
        .DateCreated = ws.Cells(13, 2).Value
        .Status = ws.Cells(14, 2).Value
        .FilePath = FilePath
    End With

    DataOperations.SafeCloseWorkbook QuoteWB, False
    LoadQuote = QuoteInfo
    Exit Function

Error_Handler:
    If Not QuoteWB Is Nothing Then DataOperations.SafeCloseWorkbook QuoteWB, False
    SystemCore.HandleStandardErrors Err.Number, "LoadQuote", "BusinessLogic"
End Function

' **Purpose**: Update existing quote with new data
' **Parameters**:
'   - QuoteInfo (QuoteData): Updated quote information
' **Returns**: Boolean - True if update successful, False if failed
' **Dependencies**: DataOperations.SafeOpenWorkbook, PopulateQuoteTemplate
' **Side Effects**: Modifies quote file, saves changes
' **Errors**: Returns False if file access or update fails
Public Function UpdateQuote(ByRef QuoteInfo As SystemCore.QuoteData) As Boolean
    Dim QuoteWB As Workbook

    On Error GoTo Error_Handler

    ' Validate data before updating
    If ValidateQuoteData(QuoteInfo) <> "" Then
        UpdateQuote = False
        Exit Function
    End If

    Set QuoteWB = DataOperations.SafeOpenWorkbook(QuoteInfo.FilePath)
    If QuoteWB Is Nothing Then
        UpdateQuote = False
        Exit Function
    End If

    PopulateQuoteTemplate QuoteWB, QuoteInfo

    QuoteWB.Save
    DataOperations.SafeCloseWorkbook QuoteWB

    ' Update search database
    Dim SearchRecord As SystemCore.SearchRecord
    SearchRecord = CreateSearchRecord(SystemCore.rtQuote, QuoteInfo.QuoteNumber, QuoteInfo.CustomerName, QuoteInfo.ComponentDescription, QuoteInfo.FilePath, "")
    UpdateSearchDatabase SearchRecord

    UpdateQuote = True
    Exit Function

Error_Handler:
    If Not QuoteWB Is Nothing Then DataOperations.SafeCloseWorkbook QuoteWB, False
    SystemCore.HandleStandardErrors Err.Number, "UpdateQuote", "BusinessLogic"
    UpdateQuote = False
End Function

' **Purpose**: Validate quote data completeness and business rules
' **Parameters**:
'   - QuoteInfo (QuoteData): Quote data to validate
' **Returns**: String - Validation error messages, empty if valid
' **Dependencies**: None
' **Side Effects**: None
' **Errors**: Returns error descriptions for invalid data
Public Function ValidateQuoteData(ByRef QuoteInfo As SystemCore.QuoteData) As String
    Dim ValidationErrors As String

    If Trim(QuoteInfo.CustomerName) = "" Then
        ValidationErrors = ValidationErrors & "Customer name is required." & vbCrLf
    End If

    If Trim(QuoteInfo.ComponentDescription) = "" Then
        ValidationErrors = ValidationErrors & "Component description is required." & vbCrLf
    End If

    If QuoteInfo.Quantity <= 0 Then
        ValidationErrors = ValidationErrors & "Quantity must be greater than zero." & vbCrLf
    End If

    If QuoteInfo.UnitPrice <= 0 Then
        ValidationErrors = ValidationErrors & "Unit price must be greater than zero." & vbCrLf
    End If

    ValidateQuoteData = ValidationErrors
End Function

' ===================================================================
' JOB MANAGEMENT (CLAUDE.md: Preserve Jobs workflow)
' ===================================================================

' **Purpose**: Create job from accepted quote
' **Parameters**:
'   - QuoteInfo (QuoteData): Source quote information
'   - JobInfo (JobData): Job information to populate
' **Returns**: Boolean - True if job created successfully, False if failed
' **Dependencies**: DataOperations.GetNextJobNumber, DataOperations.SafeOpenWorkbook
' **Side Effects**: Creates new job Excel file, updates search database, moves quote to archive
' **Errors**: Returns False on template missing or file creation failure
' **CLAUDE.md Compliance**: Maintains Enquiry → Quote → Jobs workflow
Public Function CreateJobFromQuote(ByRef QuoteInfo As SystemCore.QuoteData, ByRef JobInfo As SystemCore.JobData) As Boolean
    Dim JobNumber As String
    Dim TemplatePath As String
    Dim NewFilePath As String
    Dim TemplateWB As Workbook
    Dim SearchRecord As SystemCore.SearchRecord

    On Error GoTo Error_Handler

    JobNumber = DataOperations.GetNextJobNumber()
    If JobNumber = "" Then
        CreateJobFromQuote = False
        Exit Function
    End If

    ' Populate job info from quote
    With JobInfo
        .JobNumber = JobNumber
        .QuoteNumber = QuoteInfo.QuoteNumber
        .CustomerName = QuoteInfo.CustomerName
        .ComponentDescription = QuoteInfo.ComponentDescription
        .ComponentCode = QuoteInfo.ComponentCode
        .MaterialGrade = QuoteInfo.MaterialGrade
        .Quantity = QuoteInfo.Quantity
        .OrderValue = QuoteInfo.TotalPrice
        .DateCreated = Now
        .Status = "Active"
        .DueDate = DateAdd("d", 14, Now) ' Default 14 days lead time
        ' Initialize job-specific fields
        .WorkshopDueDate = .DueDate
        .CustomerDueDate = .DueDate
        .AssignedOperator = ""
        .Operations = ""
        .Pictures = ""
        .Notes = ""
    End With

    TemplatePath = DataOperations.GetRootPath & "\Templates\_Job.xls"
    NewFilePath = DataOperations.GetRootPath & "\WIP\" & JobNumber & ".xls"

    If Not DataOperations.FileExists(TemplatePath) Then
        SystemCore.LogError SystemCore.ERR_FILE_NOT_FOUND, "Job template not found: " & TemplatePath, "CreateJobFromQuote", "BusinessLogic"
        CreateJobFromQuote = False
        Exit Function
    End If

    Set TemplateWB = DataOperations.SafeOpenWorkbook(TemplatePath)
    If TemplateWB Is Nothing Then
        CreateJobFromQuote = False
        Exit Function
    End If

    PopulateJobTemplate TemplateWB, JobInfo

    TemplateWB.SaveAs NewFilePath
    DataOperations.SafeCloseWorkbook TemplateWB

    JobInfo.FilePath = NewFilePath

    ' Update search database
    SearchRecord = CreateSearchRecord(SystemCore.rtJob, JobNumber, JobInfo.CustomerName, JobInfo.ComponentDescription, NewFilePath, "")
    UpdateSearchDatabase SearchRecord

    ' Archive the quote
    ArchiveQuote QuoteInfo

    CreateJobFromQuote = True
    Exit Function

Error_Handler:
    If Not TemplateWB Is Nothing Then DataOperations.SafeCloseWorkbook TemplateWB, False
    SystemCore.HandleStandardErrors Err.Number, "CreateJobFromQuote", "BusinessLogic"
    CreateJobFromQuote = False
End Function

' **Purpose**: Load job data from file
' **Parameters**:
'   - FilePath (String): Full path to job file
' **Returns**: JobData - Populated job structure, empty if failed
' **Dependencies**: DataOperations.SafeOpenWorkbook for file access
' **Side Effects**: Opens and closes job file
' **Errors**: Returns empty structure if file access fails
Public Function LoadJob(ByVal FilePath As String) As SystemCore.JobData
    Dim JobWB As Workbook
    Dim ws As Worksheet
    Dim JobInfo As SystemCore.JobData

    On Error GoTo Error_Handler

    Set JobWB = DataOperations.SafeOpenWorkbook(FilePath)
    If JobWB Is Nothing Then
        Exit Function
    End If

    Set ws = JobWB.Worksheets(1)

    With JobInfo
        .JobNumber = ws.Cells(2, 2).Value
        .QuoteNumber = ws.Cells(3, 2).Value
        .CustomerName = ws.Cells(4, 2).Value
        .ComponentDescription = ws.Cells(5, 2).Value
        .ComponentCode = ws.Cells(6, 2).Value
        .MaterialGrade = ws.Cells(7, 2).Value
        .Quantity = ws.Cells(8, 2).Value
        .OrderValue = ws.Cells(9, 2).Value
        .DueDate = ws.Cells(10, 2).Value
        .WorkshopDueDate = ws.Cells(11, 2).Value
        .CustomerDueDate = ws.Cells(12, 2).Value
        .DateCreated = ws.Cells(13, 2).Value
        .Status = ws.Cells(14, 2).Value
        .AssignedOperator = ws.Cells(15, 2).Value
        .Operations = ws.Cells(16, 2).Value
        .Pictures = ws.Cells(17, 2).Value
        .Notes = ws.Cells(18, 2).Value
        .FilePath = FilePath
    End With

    DataOperations.SafeCloseWorkbook JobWB, False
    LoadJob = JobInfo
    Exit Function

Error_Handler:
    If Not JobWB Is Nothing Then DataOperations.SafeCloseWorkbook JobWB, False
    SystemCore.HandleStandardErrors Err.Number, "LoadJob", "BusinessLogic"
End Function

' **Purpose**: Update existing job with new data
' **Parameters**:
'   - JobInfo (JobData): Updated job information
' **Returns**: Boolean - True if update successful, False if failed
' **Dependencies**: DataOperations.SafeOpenWorkbook, PopulateJobTemplate
' **Side Effects**: Modifies job file, saves changes
' **Errors**: Returns False if file access or update fails
Public Function UpdateJob(ByRef JobInfo As SystemCore.JobData) As Boolean
    Dim JobWB As Workbook

    On Error GoTo Error_Handler

    ' Validate data before updating
    If ValidateJobData(JobInfo) <> "" Then
        UpdateJob = False
        Exit Function
    End If

    Set JobWB = DataOperations.SafeOpenWorkbook(JobInfo.FilePath)
    If JobWB Is Nothing Then
        UpdateJob = False
        Exit Function
    End If

    PopulateJobTemplate JobWB, JobInfo

    JobWB.Save
    DataOperations.SafeCloseWorkbook JobWB

    ' Update search database
    Dim SearchRecord As SystemCore.SearchRecord
    SearchRecord = CreateSearchRecord(SystemCore.rtJob, JobInfo.JobNumber, JobInfo.CustomerName, JobInfo.ComponentDescription, JobInfo.FilePath, "")
    UpdateSearchDatabase SearchRecord

    UpdateJob = True
    Exit Function

Error_Handler:
    If Not JobWB Is Nothing Then DataOperations.SafeCloseWorkbook JobWB, False
    SystemCore.HandleStandardErrors Err.Number, "UpdateJob", "BusinessLogic"
    UpdateJob = False
End Function

' **Purpose**: Validate job data completeness and business rules
' **Parameters**:
'   - JobInfo (JobData): Job data to validate
' **Returns**: String - Validation error messages, empty if valid
' **Dependencies**: None
' **Side Effects**: None
' **Errors**: Returns error descriptions for invalid data
Public Function ValidateJobData(ByRef JobInfo As SystemCore.JobData) As String
    Dim ValidationErrors As String

    If Trim(JobInfo.CustomerName) = "" Then
        ValidationErrors = ValidationErrors & "Customer name is required." & vbCrLf
    End If

    If Trim(JobInfo.ComponentDescription) = "" Then
        ValidationErrors = ValidationErrors & "Component description is required." & vbCrLf
    End If

    If JobInfo.Quantity <= 0 Then
        ValidationErrors = ValidationErrors & "Quantity must be greater than zero." & vbCrLf
    End If

    If JobInfo.OrderValue <= 0 Then
        ValidationErrors = ValidationErrors & "Order value must be greater than zero." & vbCrLf
    End If

    ValidateJobData = ValidationErrors
End Function

' **Purpose**: Close completed job and move to archive
' **Parameters**:
'   - JobNumber (String): Job number to close
' **Returns**: Boolean - True if job closed successfully, False if failed
' **Dependencies**: LoadJob, ArchiveJob
' **Side Effects**: Moves job from WIP to Archive directory, updates status
' **Errors**: Returns False if job not found or archive fails
Public Function CloseJob(ByVal JobNumber As String) As Boolean
    Dim JobPath As String
    Dim JobInfo As SystemCore.JobData

    On Error GoTo Error_Handler

    JobPath = DataOperations.GetRootPath & "\WIP\" & JobNumber & ".xls"

    If Not DataOperations.FileExists(JobPath) Then
        SystemCore.LogError SystemCore.ERR_FILE_NOT_FOUND, "Job file not found: " & JobPath, "CloseJob", "BusinessLogic"
        CloseJob = False
        Exit Function
    End If

    JobInfo = LoadJob(JobPath)
    If JobInfo.JobNumber = "" Then
        CloseJob = False
        Exit Function
    End If

    JobInfo.Status = "Completed"
    UpdateJob JobInfo

    ' Archive the job
    CloseJob = ArchiveJob(JobInfo)
    Exit Function

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "CloseJob", "BusinessLogic"
    CloseJob = False
End Function

' ===================================================================
' SEARCH FUNCTIONALITY (From SearchManager.bas)
' ===================================================================

' **Purpose**: Validate search system can access existing files and directories
' **Parameters**: None
' **Returns**: Boolean - True if all critical directories and files accessible
' **Dependencies**: DataOperations.FileExists, DataOperations.GetRootPath
' **Side Effects**: None
' **Errors**: Returns False if critical files/directories missing
' **CLAUDE.md Compliance**: Ensures compatibility with existing file structure
Public Function ValidateSearchCompatibility() As Boolean
    Dim RootPath As String
    Dim RequiredDirs As Variant
    Dim i As Integer

    On Error GoTo Error_Handler

    RootPath = DataOperations.GetRootPath
    RequiredDirs = Array("Enquiries", "Quotes", "WIP", "Customers", "Templates", "Archive")

    ' Check if main directories exist
    For i = 0 To UBound(RequiredDirs)
        If Not DataOperations.DirectoryExists(RootPath & "\" & RequiredDirs(i)) Then
            SystemCore.LogError SystemCore.ERR_FILE_NOT_FOUND, "Required directory missing: " & RequiredDirs(i), "ValidateSearchCompatibility", "BusinessLogic"
            ValidateSearchCompatibility = False
            Exit Function
        End If
    Next i

    ' Check if search database exists or can be created
    If Not DataOperations.FileExists(RootPath & "\" & SEARCH_FILE) Then
        ' Try to create search database if missing
        If Not CreateSearchDatabase() Then
            ValidateSearchCompatibility = False
            Exit Function
        End If
    End If

    ValidateSearchCompatibility = True
    Exit Function

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "ValidateSearchCompatibility", "BusinessLogic"
    ValidateSearchCompatibility = False
End Function

' **Purpose**: Search all PCS records with basic functionality
' **Parameters**:
'   - SearchTerm (String): Text to search for in records
'   - RecordTypeFilter (RecordType, Optional): Limit search to specific record type
' **Returns**: Variant array of SearchRecord objects, empty array if no matches
' **Dependencies**: SearchRecords_Optimized for actual search implementation
' **Side Effects**: None
' **Errors**: Returns empty array on search failure
' **CLAUDE.md Compliance**: Maintains "finds anything in the system" requirement
Public Function SearchRecords(ByVal SearchTerm As String, Optional ByVal RecordTypeFilter As Long = 0) As Variant
    On Error GoTo Error_Handler

    SearchRecords = SearchRecords_Optimized(SearchTerm, RecordTypeFilter)
    Exit Function

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "SearchRecords", "BusinessLogic"
    SearchRecords = Array()
End Function

' **Purpose**: Search all PCS records with optimization for recent files
' **Parameters**:
'   - SearchTerm (String): Text to search for in records
'   - RecordTypeFilter (RecordType, Optional): Limit search to specific record type
' **Returns**: Variant array of SearchRecord objects, empty array if no matches
' **Dependencies**: DataOperations.SafeOpenWorkbook for database access, LogSearchHistory for tracking
' **Side Effects**: Updates search history database, sorts search database by date
' **Errors**: Returns empty array on database access failure
' **CLAUDE.md Compliance**: Enhanced version maintaining all legacy search functionality
Public Function SearchRecords_Optimized(ByVal SearchTerm As String, Optional ByVal RecordTypeFilter As Long = 0) As Variant
    Dim SearchWB As Workbook
    Dim SearchWS As Worksheet
    Dim LastRow As Long
    Dim i As Long
    Dim Results() As SystemCore.SearchRecord
    Dim ResultCount As Integer
    Dim CurrentRecord As SystemCore.SearchRecord
    Dim RecentCutoff As Date
    Dim RecentResults() As SystemCore.SearchRecord
    Dim OtherResults() As SystemCore.SearchRecord
    Dim RecentCount As Integer
    Dim OtherCount As Integer

    On Error GoTo Error_Handler

    Set SearchWB = DataOperations.SafeOpenWorkbook(DataOperations.GetRootPath & "\" & SEARCH_FILE)
    If SearchWB Is Nothing Then
        SearchRecords_Optimized = Array()
        Exit Function
    End If

    Set SearchWS = SearchWB.Worksheets(1)
    LastRow = SearchWS.Cells(SearchWS.Rows.Count, 1).End(xlUp).Row

    SearchTerm = UCase(SearchTerm)
    ResultCount = 0
    RecentCount = 0
    OtherCount = 0
    RecentCutoff = DateAdd("d", -30, Now)

    ' Quick return if search database is empty
    If LastRow <= 2 Then
        DataOperations.SafeCloseWorkbook SearchWB, False
        SearchRecords_Optimized = Array()
        Exit Function
    End If

    ' Search with recent files prioritized
    For i = 2 To LastRow
        With SearchWS
            If RecordTypeFilter = 0 Or .Cells(i, 1).Value = CStr(RecordTypeFilter) Then
                If InStr(UCase(.Cells(i, 2).Value), SearchTerm) > 0 Or _
                   InStr(UCase(.Cells(i, 3).Value), SearchTerm) > 0 Or _
                   InStr(UCase(.Cells(i, 4).Value), SearchTerm) > 0 Or _
                   InStr(UCase(.Cells(i, 7).Value), SearchTerm) > 0 Then

                    With CurrentRecord
                        .RecordType = SearchWS.Cells(i, 1).Value
                        .RecordNumber = SearchWS.Cells(i, 2).Value
                        .CustomerName = SearchWS.Cells(i, 3).Value
                        .Description = SearchWS.Cells(i, 4).Value
                        .DateCreated = SearchWS.Cells(i, 5).Value
                        .FilePath = SearchWS.Cells(i, 6).Value
                        .Keywords = SearchWS.Cells(i, 7).Value
                    End With

                    ' Separate recent vs older results
                    If CurrentRecord.DateCreated >= RecentCutoff Then
                        ReDim Preserve RecentResults(RecentCount)
                        RecentResults(RecentCount) = CurrentRecord
                        RecentCount = RecentCount + 1
                    Else
                        ReDim Preserve OtherResults(OtherCount)
                        OtherResults(OtherCount) = CurrentRecord
                        OtherCount = OtherCount + 1
                    End If

                    ResultCount = ResultCount + 1
                End If
            End If
        End With
    Next i

    DataOperations.SafeCloseWorkbook SearchWB, False

    ' Combine results: recent files first, then older files
    If ResultCount > 0 Then
        ReDim Results(ResultCount - 1)
        Dim ResultIndex As Integer
        ResultIndex = 0

        ' Add recent results first
        For i = 0 To RecentCount - 1
            Results(ResultIndex) = RecentResults(i)
            ResultIndex = ResultIndex + 1
        Next i

        ' Add other results
        For i = 0 To OtherCount - 1
            Results(ResultIndex) = OtherResults(i)
            ResultIndex = ResultIndex + 1
        Next i

        ' Convert SearchRecord array to 2D array for return
        Dim OutputArray() As String
        ReDim OutputArray(0 To UBound(Results), 0 To 6)

        For i = 0 To UBound(Results)
            OutputArray(i, 0) = Results(i).RecordType
            OutputArray(i, 1) = Results(i).RecordNumber
            OutputArray(i, 2) = Results(i).CustomerName
            OutputArray(i, 3) = Results(i).Description
            OutputArray(i, 4) = CStr(Results(i).DateCreated)
            OutputArray(i, 5) = Results(i).FilePath
            OutputArray(i, 6) = Results(i).Keywords
        Next i

        SearchRecords_Optimized = OutputArray
    Else
        SearchRecords_Optimized = Array()
    End If

    ' Log search for analytics
    LogSearchHistory SearchTerm, ResultCount
    Exit Function

Error_Handler:
    If Not SearchWB Is Nothing Then DataOperations.SafeCloseWorkbook SearchWB, False
    SystemCore.HandleStandardErrors Err.Number, "SearchRecords_Optimized", "BusinessLogic"
    SearchRecords_Optimized = Array()
End Function

' **Purpose**: Create search record from data components
' **Parameters**:
'   - RecordType (RecordType): Type of record (Enquiry, Quote, Job, etc.)
'   - RecordNumber (String): Record number (E00001, Q00001, etc.)
'   - CustomerName (String): Customer name for the record
'   - Description (String): Description/component for the record
'   - FilePath (String): Full path to record file
'   - Keywords (String, Optional): Additional search keywords
' **Returns**: SearchRecord - Populated search record structure
' **Dependencies**: None
' **Side Effects**: None
' **Errors**: None
Public Function CreateSearchRecord(ByVal RecordType As Long, ByVal RecordNumber As String, ByVal CustomerName As String, ByVal Description As String, ByVal FilePath As String, Optional ByVal Keywords As String = "") As SystemCore.SearchRecord
    Dim SearchRecord As SystemCore.SearchRecord

    With SearchRecord
        .RecordType = CStr(RecordType)
        .RecordNumber = RecordNumber
        .CustomerName = CustomerName
        .Description = Description
        .DateCreated = Now
        .FilePath = FilePath
        If Keywords = "" Then
            .Keywords = CustomerName & " " & Description & " " & RecordNumber
        Else
            .Keywords = Keywords
        End If
    End With

    CreateSearchRecord = SearchRecord
End Function

' **Purpose**: Update search database with new or modified record
' **Parameters**:
'   - SearchRecord (SearchRecord): Record to add or update in search database
' **Returns**: Boolean - True if update successful, False if failed
' **Dependencies**: DataOperations.SafeOpenWorkbook for database access
' **Side Effects**: Updates search database file, sorts by date
' **Errors**: Returns False if database access fails
Public Function UpdateSearchDatabase(ByRef SearchRecord As SystemCore.SearchRecord) As Boolean
    Dim SearchWB As Workbook
    Dim SearchWS As Worksheet
    Dim LastRow As Long
    Dim NewRow As Long

    On Error GoTo Error_Handler

    Set SearchWB = DataOperations.SafeOpenWorkbook(DataOperations.GetRootPath & "\" & SEARCH_FILE)
    If SearchWB Is Nothing Then
        UpdateSearchDatabase = False
        Exit Function
    End If

    Set SearchWS = SearchWB.Worksheets(1)
    LastRow = SearchWS.Cells(SearchWS.Rows.Count, 1).End(xlUp).Row
    NewRow = LastRow + 1

    ' Add new record
    With SearchWS
        .Cells(NewRow, 1).Value = SearchRecord.RecordType
        .Cells(NewRow, 2).Value = SearchRecord.RecordNumber
        .Cells(NewRow, 3).Value = SearchRecord.CustomerName
        .Cells(NewRow, 4).Value = SearchRecord.Description
        .Cells(NewRow, 5).Value = SearchRecord.DateCreated
        .Cells(NewRow, 6).Value = SearchRecord.FilePath
        .Cells(NewRow, 7).Value = SearchRecord.Keywords
    End With

    SearchWB.Save
    DataOperations.SafeCloseWorkbook SearchWB

    UpdateSearchDatabase = True
    Exit Function

Error_Handler:
    If Not SearchWB Is Nothing Then DataOperations.SafeCloseWorkbook SearchWB, False
    SystemCore.HandleStandardErrors Err.Number, "UpdateSearchDatabase", "BusinessLogic"
    UpdateSearchDatabase = False
End Function

' **Purpose**: Create search database if missing
' **Returns**: Boolean - True if created successfully, False if failed
' **Dependencies**: DataOperations.CreateNewWorkbook
' **Side Effects**: Creates new search database file
' **Errors**: Returns False if file creation fails
Private Function CreateSearchDatabase() As Boolean
    Dim SearchWB As Workbook
    Dim SearchWS As Worksheet
    Dim SearchPath As String

    On Error GoTo Error_Handler

    SearchPath = DataOperations.GetRootPath & "\" & SEARCH_FILE

    Set SearchWB = DataOperations.CreateNewWorkbook()
    If SearchWB Is Nothing Then
        CreateSearchDatabase = False
        Exit Function
    End If

    Set SearchWS = SearchWB.Worksheets(1)
    SearchWS.Name = "Search"

    ' Create header row
    With SearchWS
        .Cells(1, 1).Value = "RecordType"
        .Cells(1, 2).Value = "RecordNumber"
        .Cells(1, 3).Value = "CustomerName"
        .Cells(1, 4).Value = "Description"
        .Cells(1, 5).Value = "DateCreated"
        .Cells(1, 6).Value = "FilePath"
        .Cells(1, 7).Value = "Keywords"
        .Range("A1:G1").Font.Bold = True
    End With

    SearchWB.SaveAs SearchPath
    SearchWB.Close
    Set SearchWB = Nothing

    CreateSearchDatabase = True
    Exit Function

Error_Handler:
    If Not SearchWB Is Nothing Then
        SearchWB.Close SaveChanges:=False
        Set SearchWB = Nothing
    End If
    SystemCore.HandleStandardErrors Err.Number, "CreateSearchDatabase", "BusinessLogic"
    CreateSearchDatabase = False
End Function

' **Purpose**: Log search history for analytics
' **Parameters**:
'   - SearchTerm (String): Term that was searched
'   - ResultCount (Integer): Number of results found
' **Returns**: None (Subroutine)
' **Dependencies**: DataOperations.FileExists, DataOperations.SafeOpenWorkbook
' **Side Effects**: Updates search history file
' **Errors**: Logs errors if history update fails
Private Sub LogSearchHistory(ByVal SearchTerm As String, ByVal ResultCount As Integer)
    Dim HistoryWB As Workbook
    Dim HistoryWS As Worksheet
    Dim HistoryPath As String
    Dim NewRow As Long

    On Error GoTo Error_Handler

    HistoryPath = DataOperations.GetRootPath & "\" & SEARCH_HISTORY_FILE

    If Not DataOperations.FileExists(HistoryPath) Then
        CreateSearchHistoryFile HistoryPath
    End If

    Set HistoryWB = DataOperations.SafeOpenWorkbook(HistoryPath)
    If HistoryWB Is Nothing Then Exit Sub

    Set HistoryWS = HistoryWB.Worksheets(1)
    NewRow = HistoryWS.Cells(HistoryWS.Rows.Count, 1).End(xlUp).Row + 1

    With HistoryWS
        .Cells(NewRow, 1).Value = Now
        .Cells(NewRow, 2).Value = SearchTerm
        .Cells(NewRow, 3).Value = ResultCount
    End With

    HistoryWB.Save
    DataOperations.SafeCloseWorkbook HistoryWB
    Exit Sub

Error_Handler:
    If Not HistoryWB Is Nothing Then DataOperations.SafeCloseWorkbook HistoryWB, False
    SystemCore.LogError Err.Number, Err.Description, "LogSearchHistory", "BusinessLogic"
End Sub

' **Purpose**: Create search history file if missing
' **Parameters**:
'   - FilePath (String): Path for new history file
' **Returns**: None (Subroutine)
' **Dependencies**: DataOperations.CreateNewWorkbook
' **Side Effects**: Creates new search history file
' **Errors**: Logs errors if file creation fails
Private Sub CreateSearchHistoryFile(ByVal FilePath As String)
    Dim HistoryWB As Workbook
    Dim HistoryWS As Worksheet

    On Error GoTo Error_Handler

    Set HistoryWB = DataOperations.CreateNewWorkbook()
    If HistoryWB Is Nothing Then Exit Sub

    Set HistoryWS = HistoryWB.Worksheets(1)
    HistoryWS.Name = "SearchHistory"

    ' Create header row
    With HistoryWS
        .Cells(1, 1).Value = "DateTime"
        .Cells(1, 2).Value = "SearchTerm"
        .Cells(1, 3).Value = "ResultCount"
        .Range("A1:C1").Font.Bold = True
    End With

    HistoryWB.SaveAs FilePath
    HistoryWB.Close
    Set HistoryWB = Nothing
    Exit Sub

Error_Handler:
    If Not HistoryWB Is Nothing Then
        HistoryWB.Close SaveChanges:=False
        Set HistoryWB = Nothing
    End If
    SystemCore.LogError Err.Number, Err.Description, "CreateSearchHistoryFile", "BusinessLogic"
End Sub

' ===================================================================
' SEARCH OPERATIONS (CLAUDE.md: SearchOperations.bas replacement)
' ===================================================================

' **Purpose**: Update search database with file information from all folders
' **Original**: Interface_VBA/SearchOperations.bas Update_Search()
' **Parameters**: None
' **Returns**: None (Subroutine)
' **Dependencies**: DataOperations.GetValue(), SafeOpenWorkbook()
' **Side Effects**: Opens search.xls, scans Archive/Enquiries/Quotes/WIP folders, updates search database
' **Errors**: May display message boxes and end execution on critical errors
' **CLAUDE.md Compliance**: Exact replacement for SearchOperations.bas Update_Search functionality
Public Sub Update_Search()
    Dim Files(1 To 100000) As String
    Dim FullFilePath As String, MyName As String
    Dim GroupCount As Integer
    Dim FolderName As String
    Dim SearchWB As Workbook
    Dim i As Integer
    Dim ItemType As String
    Dim ItemValue As String
    Dim j As Integer
    Dim fileextension As String

    On Error GoTo Error_Handler

    ' Open search database
    Set SearchWB = DataOperations.SafeOpenWorkbook(DataOperations.GetRootPath & "\search.xls")
    If SearchWB Is Nothing Then
        MsgBox "Cannot open search.xls", vbCritical
        Exit Sub
    End If

    SearchWB.Worksheets(1).Range("A:A").Font.Bold = True
    fileextension = "*.xls"

    ' Skip automatic file listing (legacy behavior) - uncomment next line to enable
    GoTo SkipHERE

    ' Clear existing data (rows 3 to 35000)
    SearchWB.Worksheets(1).Range("3:35000").Clear

    ' Process each folder
    For i = 1 To 4
        Select Case i
            Case 1
                FolderName = DataOperations.GetRootPath & "\Archive"
            Case 2
                FolderName = DataOperations.GetRootPath & "\Enquiries"
            Case 3
                FolderName = DataOperations.GetRootPath & "\Quotes"
            Case 4
                FolderName = DataOperations.GetRootPath & "\WIP"
        End Select

        MyName = Dir(FolderName & "\", vbDirectory)
        If MyName = "" Then
            MsgBox "Folder Not Found: " & FolderName, vbOKOnly, "Error"
            Exit Sub
        End If

        ' Store list of files
        Do Until MyName = ""
            If MyName <> "." And MyName <> ".." Then
                SearchWB.Worksheets(1).Range("A1").Select
                Do
                    SearchWB.ActiveCell.Offset(1, 0).Select
                Loop Until SearchWB.ActiveCell.Value = "" Or SearchWB.ActiveCell.Value = Left(MyName, Len(MyName) - 4)

                SearchWB.ActiveCell.Value = Left(MyName, Len(MyName) - 4)
            End If
            MyName = Dir
        Loop
    Next i

    SearchWB.Worksheets(1).Range("A3").Select

SkipHERE:
    ' Get starting row from user (legacy behavior)
    Dim StartRow As String
    StartRow = InputBox("Please adjust if you wish to move to a specific row", "SKIP TO ROW", SearchWB.ActiveCell.Row)
    If IsNumeric(StartRow) Then
        SearchWB.Worksheets(1).Range("A" & StartRow).Select
    End If

    ' Process each file in search list
    Do
        FolderName = ""
        ' Find the file in one of the folders
        If Dir(DataOperations.GetRootPath & "\Archive\" & SearchWB.ActiveCell.Value & ".xls", vbNormal) <> "" Then
            FolderName = DataOperations.GetRootPath & "\Archive\"
            GoTo CopyInfo
        ElseIf Dir(DataOperations.GetRootPath & "\Enquiries\" & SearchWB.ActiveCell.Value & ".xls", vbNormal) <> "" Then
            FolderName = DataOperations.GetRootPath & "\Enquiries\"
            GoTo CopyInfo
        ElseIf Dir(DataOperations.GetRootPath & "\Quotes\" & SearchWB.ActiveCell.Value & ".xls", vbNormal) <> "" Then
            FolderName = DataOperations.GetRootPath & "\Quotes\"
            GoTo CopyInfo
        ElseIf Dir(DataOperations.GetRootPath & "\WIP\" & SearchWB.ActiveCell.Value & ".xls", vbNormal) <> "" Then
            FolderName = DataOperations.GetRootPath & "\WIP\"
            GoTo CopyInfo
        End If

        MsgBox "CANT FIND THE FILE: " & SearchWB.ActiveCell.Value
        Exit Sub

CopyInfo:
        i = 0
        Do
            i = i + 1
            ItemType = DataOperations.GetValueFromClosedWorkbook(FolderName & SearchWB.ActiveCell.Value & ".xls", "Admin", "A" & i)
            ItemValue = DataOperations.GetValueFromClosedWorkbook(FolderName & SearchWB.ActiveCell.Value & ".xls", "Admin", "B" & i)

            j = 0
            Do
                j = j + 1
                If UCase(SearchWB.Worksheets(1).Range("A1").Offset(0, j).Value) = UCase(ItemType) Then
                    If SearchWB.ActiveCell.Offset(0, j).Value = "" Or UCase(SearchWB.ActiveCell.Offset(0, j).Value) = UCase(ItemValue) Then
                        SearchWB.ActiveCell.Offset(0, j).Value = UCase(ItemValue)
                    Else
                        If InStr(1, ItemType, "DATE", vbTextCompare) > 0 Then
                            If CCur(SearchWB.ActiveCell.Offset(0, j).Value) = CCur(ItemValue) Then
                                SearchWB.ActiveCell.Offset(0, j).Value = UCase(ItemValue)
                            Else
                                If MsgBox("A Difference Exists with regards to - " & ItemType & vbNewLine & "Do you wish to replace : " & SearchWB.ActiveCell.Offset(0, j).Value & " with : " & CDate(ItemValue), vbYesNo) = vbYes Then
                                    SearchWB.ActiveCell.Offset(0, j).Value = UCase(ItemValue)
                                Else
                                    If MsgBox("Do you wish to continue?", vbYesNo) = vbNo Then
                                        Exit Sub
                                    End If
                                End If
                            End If
                        Else
                            If MsgBox("A Difference Exists with regards to - " & ItemType & vbNewLine & "Do you wish to replace : " & SearchWB.ActiveCell.Offset(0, j).Value & " with : " & ItemValue, vbYesNo) = vbYes Then
                                SearchWB.ActiveCell.Offset(0, j).Value = UCase(ItemValue)
                            Else
                                If MsgBox("Do you wish to continue?", vbYesNo) = vbNo Then
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                    SearchWB.ActiveCell.Font.Bold = False
                    GoTo NextType
                End If
            Loop Until SearchWB.Worksheets(1).Range("A1").Offset(0, j + 1).Value = ""
NextType:
        Loop Until ItemType = ""

        SearchWB.ActiveCell.Offset(1, 0).Select
    Loop Until SearchWB.ActiveCell.Value = ""

    SearchWB.Save
    SearchWB.Close
    Set SearchWB = Nothing
    Exit Sub

Error_Handler:
    If Not SearchWB Is Nothing Then
        SearchWB.Close SaveChanges:=False
        Set SearchWB = Nothing
    End If
    SystemCore.HandleStandardErrors Err.Number, "Update_Search", "BusinessLogic"
End Sub

' **Purpose**: Search synchronization with password protection and backup creation
' **Original**: Interface_VBA/SearchOperations.bas SeachSYNC()
' **Parameters**: None
' **Returns**: None (Subroutine)
' **Dependencies**: DataOperations.GetRootPath(), password validation, Calc_Next_Number()
' **Side Effects**: Creates backups, updates Search History.xls, cleans old records
' **Errors**: Ends execution on incorrect password or critical errors
' **CLAUDE.md Compliance**: Exact replacement for SearchOperations.bas SeachSYNC functionality
Public Sub SeachSYNC()
    Dim DCSData(0 To 30) As Variant
    Dim DelDate As Date
    Dim SearchWB As Workbook
    Dim HistoryWB As Workbook
    Dim i As Integer
    Dim JC As Boolean, QN As Boolean, en As Boolean

    On Error GoTo Error_Handler

    ' Password validation (exact legacy behavior)
    If InputBox("PASSWORD") <> SYNC_PASSWORD Then
        MsgBox "ERROR - INCORRECT"
        Exit Sub
    End If

    ' Open Search.xls and create backup
    Set SearchWB = DataOperations.SafeOpenWorkbook(DataOperations.GetRootPath & "\" & SEARCH_FILE)
    If SearchWB Is Nothing Then
        MsgBox "Cannot open Search.xls", vbCritical
        Exit Sub
    End If

    SearchWB.Worksheets(1).Range("A3").Select
    SearchWB.SaveCopyAs DataOperations.GetRootPath & "\Backups\" & Format(Now(), "yyyymmdd") & " - Search.xls"

    ' Open Search History.xls and create backup
    Set HistoryWB = DataOperations.SafeOpenWorkbook(DataOperations.GetRootPath & "\" & SEARCH_HISTORY_FILE)
    If HistoryWB Is Nothing Then
        ' Create history file if missing
        CreateSearchHistoryFile DataOperations.GetRootPath & "\" & SEARCH_HISTORY_FILE
        Set HistoryWB = DataOperations.SafeOpenWorkbook(DataOperations.GetRootPath & "\" & SEARCH_HISTORY_FILE)
    End If

    If Not HistoryWB Is Nothing Then
        HistoryWB.Worksheets(1).Range("A3").Select
        HistoryWB.SaveCopyAs DataOperations.GetRootPath & "\Backups\" & Format(Now(), "yyyymmdd") & " - Search History.xls"
    End If

    ' Synchronize search records to history
    Do While SearchWB.ActiveCell.Value <> ""
        JC = False
        QN = False
        en = False

        ' Determine record type based on filled columns
        If SearchWB.ActiveCell.Offset(0, 3).Value <> "" Then
            JC = True
        ElseIf SearchWB.ActiveCell.Offset(0, 2).Value <> "" Then
            QN = True
        Else
            en = True
        End If

        ' Copy row data from Search
        For i = 0 To 30
            DCSData(i) = SearchWB.ActiveCell.Offset(0, i).Value
        Next i

        ' Find or create matching record in Search History
        If Not HistoryWB Is Nothing Then
            HistoryWB.Worksheets(1).Range("A2").Select
            Do
                HistoryWB.ActiveCell.Offset(1, 0).Select
                If (JC = True And HistoryWB.ActiveCell.Offset(0, 3).Value = DCSData(3)) Or _
                   (QN = True And HistoryWB.ActiveCell.Offset(0, 2).Value = DCSData(2)) Or _
                   (en = True And HistoryWB.ActiveCell.Offset(0, 1).Value = DCSData(1)) Then
                    Exit Do
                End If
            Loop Until HistoryWB.ActiveCell.Value = ""

            ' Fill history record with search data
            For i = 0 To 30
                HistoryWB.ActiveCell.Offset(0, i).Value = DCSData(i)
            Next i
        End If

        SearchWB.Activate
        SearchWB.ActiveCell.Offset(1, 0).Select
    Loop

    ' Save both workbooks
    If Not HistoryWB Is Nothing Then
        HistoryWB.Save
        HistoryWB.Close
        Set HistoryWB = Nothing
    End If
    SearchWB.Save

    ' Clean old records (exact legacy logic)
    SearchWB.Worksheets(1).Range("C3").Select

    Do While SearchWB.Range("A" & SearchWB.ActiveCell.Row).Value <> ""
        If SearchWB.ActiveCell.Value <> "" Then
            If SearchWB.ActiveCell.Offset(0, 1).Value <> "" Then
                ' Job record - keep if within 1000 of current job number
                If CCur(SearchWB.ActiveCell.Offset(0, 2).Value) < DataOperations.Calc_Next_Number("J") - 1000 Then
                    SearchWB.ActiveCell.EntireRow.Delete
                Else
                    SearchWB.ActiveCell.Offset(1, 0).Select
                End If
            Else
                ' Quote record - keep if within 10000 of current quote number
                If CCur(SearchWB.ActiveCell.Offset(0, 2).Value) < DataOperations.Calc_Next_Number("Q") - 10000 Then
                    SearchWB.ActiveCell.EntireRow.Delete
                Else
                    SearchWB.ActiveCell.Offset(1, 0).Select
                End If
            End If
        Else
            SearchWB.ActiveCell.Offset(1, 0).Select
        End If
    Loop

    SearchWB.Save
    SearchWB.Close
    Set SearchWB = Nothing

    MsgBox "COMPLETED"
    Exit Sub

Error_Handler:
    If Not SearchWB Is Nothing Then
        SearchWB.Close SaveChanges:=False
        Set SearchWB = Nothing
    End If
    If Not HistoryWB Is Nothing Then
        HistoryWB.Close SaveChanges:=False
        Set HistoryWB = Nothing
    End If
    SystemCore.HandleStandardErrors Err.Number, "SeachSYNC", "BusinessLogic"
End Sub

' **Purpose**: Update WIP database with job information
' **Parameters**:
'   - JobData (JobData): Job information to save to WIP database
' **Returns**: Boolean - True if successful, False if failed
' **Dependencies**: DataOperations.SaveInfoIntoWIP()
' **Side Effects**: Updates WIP.xls database with job status and information
' **Errors**: Returns False if WIP update fails
' **CLAUDE.md Compliance**: Enhanced WIP database integration for job tracking
Public Function UpdateWIPDatabase(ByRef JobData As SystemCore.JobData) As Boolean
    On Error GoTo Error_Handler

    ' Use DataOperations WIP saving function
    UpdateWIPDatabase = DataOperations.SaveInfoIntoWIP(JobData)
    Exit Function

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "UpdateWIPDatabase", "BusinessLogic"
    UpdateWIPDatabase = False
End Function

' ===================================================================
' PRIVATE HELPER FUNCTIONS
' ===================================================================

' **Purpose**: Populate enquiry template with form data
' **Parameters**:
'   - TemplateWB (Workbook): Template workbook to populate
'   - EnquiryInfo (EnquiryData): Enquiry data to use
' **Returns**: None (Subroutine)
' **Dependencies**: None
' **Side Effects**: Modifies template workbook with enquiry data
' **Errors**: May fail silently if template structure is different
Private Sub PopulateEnquiryTemplate(ByRef TemplateWB As Workbook, ByRef EnquiryInfo As SystemCore.EnquiryData)
    Dim ws As Worksheet

    On Error GoTo Error_Handler

    Set ws = TemplateWB.Worksheets("ADMIN")

    With ws
        .Cells(2, 2).Value = EnquiryInfo.EnquiryNumber
        .Cells(3, 2).Value = EnquiryInfo.CustomerName
        .Cells(4, 2).Value = EnquiryInfo.ContactPerson
        .Cells(5, 2).Value = EnquiryInfo.CompanyPhone
        .Cells(6, 2).Value = EnquiryInfo.CompanyFax
        .Cells(7, 2).Value = EnquiryInfo.Email
        .Cells(8, 2).Value = EnquiryInfo.ComponentDescription
        .Cells(9, 2).Value = EnquiryInfo.ComponentCode
        .Cells(10, 2).Value = EnquiryInfo.MaterialGrade
        .Cells(11, 2).Value = EnquiryInfo.Quantity
        .Cells(12, 2).Value = EnquiryInfo.DateCreated
    End With

    Exit Sub

Error_Handler:
    SystemCore.LogError Err.Number, Err.Description, "PopulateEnquiryTemplate", "BusinessLogic"
End Sub

' **Purpose**: Populate quote template with form data
' **Parameters**:
'   - TemplateWB (Workbook): Template workbook to populate
'   - QuoteInfo (QuoteData): Quote data to use
' **Returns**: None (Subroutine)
' **Dependencies**: None
' **Side Effects**: Modifies template workbook with quote data
' **Errors**: May fail silently if template structure is different
Private Sub PopulateQuoteTemplate(ByRef TemplateWB As Workbook, ByRef QuoteInfo As SystemCore.QuoteData)
    Dim ws As Worksheet

    On Error GoTo Error_Handler

    Set ws = TemplateWB.Worksheets("ADMIN")

    With ws
        .Cells(2, 2).Value = QuoteInfo.QuoteNumber
        .Cells(3, 2).Value = QuoteInfo.EnquiryNumber
        .Cells(4, 2).Value = QuoteInfo.CustomerName
        .Cells(5, 2).Value = QuoteInfo.ComponentDescription
        .Cells(6, 2).Value = QuoteInfo.ComponentCode
        .Cells(7, 2).Value = QuoteInfo.MaterialGrade
        .Cells(8, 2).Value = QuoteInfo.Quantity
        .Cells(9, 2).Value = QuoteInfo.UnitPrice
        .Cells(10, 2).Value = QuoteInfo.TotalPrice
        .Cells(11, 2).Value = QuoteInfo.LeadTime
        .Cells(12, 2).Value = QuoteInfo.ValidUntil
        .Cells(13, 2).Value = QuoteInfo.DateCreated
        .Cells(14, 2).Value = QuoteInfo.Status
    End With

    Exit Sub

Error_Handler:
    SystemCore.LogError Err.Number, Err.Description, "PopulateQuoteTemplate", "BusinessLogic"
End Sub

' **Purpose**: Populate job template with form data
' **Parameters**:
'   - TemplateWB (Workbook): Template workbook to populate
'   - JobInfo (JobData): Job data to use
' **Returns**: None (Subroutine)
' **Dependencies**: None
' **Side Effects**: Modifies template workbook with job data
' **Errors**: May fail silently if template structure is different
Private Sub PopulateJobTemplate(ByRef TemplateWB As Workbook, ByRef JobInfo As SystemCore.JobData)
    Dim ws As Worksheet

    On Error GoTo Error_Handler

    Set ws = TemplateWB.Worksheets("ADMIN")

    With ws
        .Cells(2, 2).Value = JobInfo.JobNumber
        .Cells(3, 2).Value = JobInfo.QuoteNumber
        .Cells(4, 2).Value = JobInfo.CustomerName
        .Cells(5, 2).Value = JobInfo.ComponentDescription
        .Cells(6, 2).Value = JobInfo.ComponentCode
        .Cells(7, 2).Value = JobInfo.MaterialGrade
        .Cells(8, 2).Value = JobInfo.Quantity
        .Cells(9, 2).Value = JobInfo.OrderValue
        .Cells(10, 2).Value = JobInfo.DueDate
        .Cells(11, 2).Value = JobInfo.WorkshopDueDate
        .Cells(12, 2).Value = JobInfo.CustomerDueDate
        .Cells(13, 2).Value = JobInfo.DateCreated
        .Cells(14, 2).Value = JobInfo.Status
        .Cells(15, 2).Value = JobInfo.AssignedOperator
        .Cells(16, 2).Value = JobInfo.Operations
        .Cells(17, 2).Value = JobInfo.Pictures
        .Cells(18, 2).Value = JobInfo.Notes
    End With

    Exit Sub

Error_Handler:
    SystemCore.LogError Err.Number, Err.Description, "PopulateJobTemplate", "BusinessLogic"
End Sub

' **Purpose**: Archive completed quote
' **Parameters**:
'   - QuoteInfo (QuoteData): Quote to archive
' **Returns**: Boolean - True if archiving successful, False if failed
' **Dependencies**: DataOperations.CreateBackup, DataOperations.FileExists
' **Side Effects**: Moves quote file to archive directory
' **Errors**: Returns False if archive operation fails
Private Function ArchiveQuote(ByRef QuoteInfo As SystemCore.QuoteData) As Boolean
    Dim SourcePath As String
    Dim ArchivePath As String

    On Error GoTo Error_Handler

    SourcePath = QuoteInfo.FilePath
    ArchivePath = DataOperations.GetRootPath & "\Archive\" & Dir(SourcePath)

    If Not DataOperations.FileExists(SourcePath) Then
        ArchiveQuote = False
        Exit Function
    End If

    ' Create backup before moving
    DataOperations.CreateBackup SourcePath

    ' Move file to archive
    FileCopy SourcePath, ArchivePath
    Kill SourcePath

    ' Update quote info with new path
    QuoteInfo.FilePath = ArchivePath

    ArchiveQuote = True
    Exit Function

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "ArchiveQuote", "BusinessLogic"
    ArchiveQuote = False
End Function

' **Purpose**: Archive completed job
' **Parameters**:
'   - JobInfo (JobData): Job to archive
' **Returns**: Boolean - True if archiving successful, False if failed
' **Dependencies**: DataOperations.CreateBackup, DataOperations.FileExists
' **Side Effects**: Moves job file from WIP to archive directory
' **Errors**: Returns False if archive operation fails
Private Function ArchiveJob(ByRef JobInfo As SystemCore.JobData) As Boolean
    Dim SourcePath As String
    Dim ArchivePath As String

    On Error GoTo Error_Handler

    SourcePath = JobInfo.FilePath
    ArchivePath = DataOperations.GetRootPath & "\Archive\" & Dir(SourcePath)

    If Not DataOperations.FileExists(SourcePath) Then
        ArchiveJob = False
        Exit Function
    End If

    ' Create backup before moving
    DataOperations.CreateBackup SourcePath

    ' Move file to archive
    FileCopy SourcePath, ArchivePath
    Kill SourcePath

    ' Update job info with new path
    JobInfo.FilePath = ArchivePath

    ArchiveJob = True
    Exit Function

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "ArchiveJob", "BusinessLogic"
    ArchiveJob = False
End Function

' **Purpose**: Save row to search database (legacy compatibility)
' **Parameters**:
'   - FormObject (Object): Form containing data to save to search
' **Returns**: Boolean - True if save successful, False if failed
' **Dependencies**: UpdateSearchDatabase
' **Side Effects**: Updates search database with form data
' **Errors**: Returns False if database update fails
' **CLAUDE.md Compliance**: Maintains legacy form compatibility
Public Function SaveRowIntoSearch(ByRef FormObject As Object) As Boolean
    Dim SearchRecord As SystemCore.SearchRecord
    Dim RecordType As Long
    Dim RecordNumber As String
    Dim CustomerName As String
    Dim Description As String

    On Error GoTo Error_Handler

    ' Extract data from form
    On Error Resume Next
    RecordNumber = FormObject.Enquiry_Number.Value
    If RecordNumber = "" Then RecordNumber = FormObject.Quote_Number.Value
    If RecordNumber = "" Then RecordNumber = FormObject.Job_Number.Value

    CustomerName = FormObject.Customer.Value
    If CustomerName = "" Then CustomerName = FormObject.Customer_Name.Value

    Description = FormObject.Component_Description.Value
    On Error GoTo Error_Handler

    ' Determine record type from number prefix
    If Left(RecordNumber, 1) = "E" Then
        RecordType = SystemCore.rtEnquiry
    ElseIf Left(RecordNumber, 1) = "Q" Then
        RecordType = SystemCore.rtQuote
    ElseIf Left(RecordNumber, 1) = "J" Then
        RecordType = SystemCore.rtJob
    Else
        RecordType = SystemCore.rtEnquiry ' Default
    End If

    SearchRecord = CreateSearchRecord(RecordType, RecordNumber, CustomerName, Description, "", "")
    SaveRowIntoSearch = UpdateSearchDatabase(SearchRecord)
    Exit Function

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "SaveRowIntoSearch", "BusinessLogic"
    SaveRowIntoSearch = False
End Function

' ===================================================================
' HISTORY AND REPORTING FUNCTIONS (Referenced by UserInterface)
' ===================================================================

' **Purpose**: Get job history from search database and archive files
' **Parameters**: None
' **Returns**: Variant - Array of job history records, empty array if none found
' **Dependencies**: SearchRecords_Optimized(), DataOperations.GetRootPath()
' **Side Effects**: None
' **Errors**: Returns empty array if search fails
Public Function GetJobHistory() As Variant
    On Error GoTo Error_Handler

    ' Search for all job records (rtJob = 3)
    GetJobHistory = SearchRecords_Optimized("", 3)
    Exit Function

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "GetJobHistory", "BusinessLogic"
    GetJobHistory = Array()
End Function

' **Purpose**: Get quote history from search database and quote files
' **Parameters**: None
' **Returns**: Variant - Array of quote history records, empty array if none found
' **Dependencies**: SearchRecords_Optimized()
' **Side Effects**: None
' **Errors**: Returns empty array if search fails
Public Function GetQuoteHistory() As Variant
    On Error GoTo Error_Handler

    ' Search for all quote records (rtQuote = 2)
    GetQuoteHistory = SearchRecords_Optimized("", 2)
    Exit Function

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "GetQuoteHistory", "BusinessLogic"
    GetQuoteHistory = Array()
End Function

' **Purpose**: Load search history into search form
' **Parameters**:
'   - SearchForm (Object): Search form to populate with history
' **Returns**: Boolean - True if history loaded successfully, False if failed
' **Dependencies**: DataOperations.SafeOpenWorkbook()
' **Side Effects**: Updates search form with recent search history
' **Errors**: Returns False if history file cannot be accessed
Public Function LoadSearchHistory(SearchForm As Object) As Boolean
    Dim HistoryWB As Workbook
    Dim HistoryWS As Worksheet
    Dim LastRow As Long
    Dim i As Long

    On Error GoTo Error_Handler

    Set HistoryWB = DataOperations.SafeOpenWorkbook(DataOperations.GetRootPath & "\" & SEARCH_HISTORY_FILE)
    If HistoryWB Is Nothing Then
        LoadSearchHistory = False
        Exit Function
    End If

    Set HistoryWS = HistoryWB.Worksheets(1)
    LastRow = HistoryWS.Cells(HistoryWS.Rows.Count, 1).End(xlUp).Row

    ' Clear existing history display
    On Error Resume Next
    SearchForm.HistoryList.Clear
    On Error GoTo Error_Handler

    ' Load recent search terms (last 50 searches)
    Dim StartRow As Long
    StartRow = IIf(LastRow > 52, LastRow - 50, 2) ' Skip header row

    For i = LastRow To StartRow Step -1 ' Reverse order (newest first)
        If HistoryWS.Cells(i, 2).Value <> "" Then
            On Error Resume Next
            SearchForm.HistoryList.AddItem Format(HistoryWS.Cells(i, 1).Value, "dd/mm/yyyy hh:mm") & " - " & HistoryWS.Cells(i, 2).Value & " (" & HistoryWS.Cells(i, 3).Value & " results)"
            On Error GoTo Error_Handler
        End If
    Next i

    DataOperations.SafeCloseWorkbook HistoryWB, False
    LoadSearchHistory = True
    Exit Function

Error_Handler:
    If Not HistoryWB Is Nothing Then DataOperations.SafeCloseWorkbook HistoryWB, False
    SystemCore.HandleStandardErrors Err.Number, "LoadSearchHistory", "BusinessLogic"
    LoadSearchHistory = False
End Function

' **Purpose**: Sort search database by date and record number
' **Parameters**: None
' **Returns**: Boolean - True if sort successful, False if failed
' **Dependencies**: DataOperations.SafeOpenWorkbook()
' **Side Effects**: Reorders search database records
' **Errors**: Returns False if sort operation fails
Public Function SortSearchDatabase() As Boolean
    Dim SearchWB As Workbook
    Dim SearchWS As Worksheet
    Dim LastRow As Long
    Dim SortRange As Range

    On Error GoTo Error_Handler

    Set SearchWB = DataOperations.SafeOpenWorkbook(DataOperations.GetRootPath & "\" & SEARCH_FILE)
    If SearchWB Is Nothing Then
        SortSearchDatabase = False
        Exit Function
    End If

    Set SearchWS = SearchWB.Worksheets(1)
    LastRow = SearchWS.Cells(SearchWS.Rows.Count, 1).End(xlUp).Row

    If LastRow <= 2 Then
        ' No data to sort
        DataOperations.SafeCloseWorkbook SearchWB, False
        SortSearchDatabase = True
        Exit Function
    End If

    ' Define sort range (exclude header row)
    Set SortRange = SearchWS.Range("A2:G" & LastRow)

    ' Sort by date (column 5) descending, then by record number (column 2)
    SortRange.Sort Key1:=SearchWS.Range("E2"), Order1:=xlDescending, _
                   Key2:=SearchWS.Range("B2"), Order2:=xlAscending, _
                   Header:=xlNo

    SearchWB.Save
    DataOperations.SafeCloseWorkbook SearchWB
    SortSearchDatabase = True
    Exit Function

Error_Handler:
    If Not SearchWB Is Nothing Then DataOperations.SafeCloseWorkbook SearchWB, False
    SystemCore.HandleStandardErrors Err.Number, "SortSearchDatabase", "BusinessLogic"
    SortSearchDatabase = False
End Function

' **Purpose**: Mark quote as called through (update status)
' **Parameters**:
'   - QuoteNumber (String): Quote number to mark as called through
' **Returns**: Boolean - True if update successful, False if failed
' **Dependencies**: DataOperations.GetRootPath(), DataOperations.SafeOpenWorkbook()
' **Side Effects**: Updates quote file status field
' **Errors**: Returns False if quote file cannot be updated
Public Function MarkQuoteCalledThrough(QuoteNumber As String) As Boolean
    Dim QuoteFilePath As String
    Dim QuoteWB As Workbook

    On Error GoTo Error_Handler

    ' Build quote file path
    QuoteFilePath = DataOperations.GetRootPath & "\Quotes\" & QuoteNumber & ".xls"

    If Not DataOperations.FileExists(QuoteFilePath) Then
        SystemCore.LogError SystemCore.ERR_FILE_NOT_FOUND, "Quote file not found: " & QuoteFilePath, "MarkQuoteCalledThrough", "BusinessLogic"
        MarkQuoteCalledThrough = False
        Exit Function
    End If

    Set QuoteWB = DataOperations.SafeOpenWorkbook(QuoteFilePath)
    If QuoteWB Is Nothing Then
        MarkQuoteCalledThrough = False
        Exit Function
    End If

    ' Update quote status to "Called Through" (typically in ADMIN sheet, cell B88)
    On Error Resume Next
    QuoteWB.Worksheets("ADMIN").Range("B88").Value = "Called Through"
    On Error GoTo Error_Handler

    QuoteWB.Save
    DataOperations.SafeCloseWorkbook QuoteWB

    ' Update search database
    Dim SearchRecord As SystemCore.SearchRecord
    SearchRecord = CreateSearchRecord(SystemCore.rtQuote, QuoteNumber, "", "Called Through", QuoteFilePath, "")
    UpdateSearchDatabase SearchRecord

    MarkQuoteCalledThrough = True
    Exit Function

Error_Handler:
    If Not QuoteWB Is Nothing Then DataOperations.SafeCloseWorkbook QuoteWB, False
    SystemCore.HandleStandardErrors Err.Number, "MarkQuoteCalledThrough", "BusinessLogic"
    MarkQuoteCalledThrough = False
End Function