Attribute VB_Name = "BusinessController"
' **Purpose**: All business process controllers and workflow management
' **CLAUDE.md Compliance**: Maintains Enquiry → Quote → Jobs workflow, preserves all business logic
Option Explicit

' ===================================================================
' CONSTANTS AND PRIVATE VARIABLES
' ===================================================================

Private Const WIP_FILE As String = "WIP.xls"

' ===================================================================
' ENQUIRY MANAGEMENT (CLAUDE.md: Preserve Enquiry → Quote → Jobs workflow)
' ===================================================================

' **Purpose**: Create new enquiry following PCS business rules
' **Parameters**:
'   - EnquiryInfo (EnquiryData): Complete enquiry information structure
' **Returns**: Boolean - True if enquiry created successfully, False if failed
' **Dependencies**: DataManager.GetNextEnquiryNumber, DataManager.SafeOpenWorkbook, SearchManager.UpdateSearchDatabase
' **Side Effects**: Creates new enquiry Excel file in Enquiries directory, updates search database
' **Errors**: Returns False on template missing, file creation failure, or validation errors
' **CLAUDE.md Compliance**: Maintains Enquiry → Quote → Jobs workflow
Public Function CreateNewEnquiry(ByRef EnquiryInfo As CoreFramework.EnquiryData) As Boolean
    Dim EnquiryNumber As String
    Dim TemplatePath As String
    Dim NewFilePath As String
    Dim TemplateWB As Workbook
    Dim SearchRecord As CoreFramework.SearchRecord

    On Error GoTo Error_Handler

    ' Validate enquiry data before processing
    If ValidateEnquiryData(EnquiryInfo) <> "" Then
        CreateNewEnquiry = False
        Exit Function
    End If

    EnquiryNumber = DataManager.GetNextEnquiryNumber()
    If EnquiryNumber = "" Then
        CreateNewEnquiry = False
        Exit Function
    End If

    EnquiryInfo.EnquiryNumber = EnquiryNumber
    EnquiryInfo.DateCreated = Now

    TemplatePath = DataManager.GetRootPath & "\Templates\_Enq.xls"
    NewFilePath = DataManager.GetRootPath & "\Enquiries\" & EnquiryNumber & ".xls"

    If Not DataManager.FileExists(TemplatePath) Then
        CoreFramework.LogError CoreFramework.ERR_FILE_NOT_FOUND, "Enquiry template not found: " & TemplatePath, "CreateNewEnquiry", "BusinessController"
        CreateNewEnquiry = False
        Exit Function
    End If

    Set TemplateWB = DataManager.SafeOpenWorkbook(TemplatePath)
    If TemplateWB Is Nothing Then
        CreateNewEnquiry = False
        Exit Function
    End If

    PopulateEnquiryTemplate TemplateWB, EnquiryInfo

    TemplateWB.SaveAs NewFilePath
    DataManager.SafeCloseWorkbook TemplateWB

    EnquiryInfo.FilePath = NewFilePath

    ' Update search database
    SearchRecord = SearchManager.CreateSearchRecord(CoreFramework.rtEnquiry, EnquiryNumber, EnquiryInfo.CustomerName, EnquiryInfo.ComponentDescription, NewFilePath, EnquiryInfo.SearchKeywords)
    SearchManager.UpdateSearchDatabase SearchRecord

    ' Create customer record if new
    If Not DataManager.FileExists(DataManager.GetRootPath & "\Customers\" & CoreFramework.CleanFileName(EnquiryInfo.CustomerName) & ".xls") Then
        CreateNewCustomer EnquiryInfo.CustomerName
    End If

    CreateNewEnquiry = True
    Exit Function

Error_Handler:
    If Not TemplateWB Is Nothing Then DataManager.SafeCloseWorkbook TemplateWB, False
    CoreFramework.HandleStandardErrors Err.Number, "CreateNewEnquiry", "BusinessController"
    CreateNewEnquiry = False
End Function

' **Purpose**: Load enquiry data from file
' **Parameters**:
'   - FilePath (String): Full path to enquiry file
' **Returns**: EnquiryData - Populated enquiry structure, empty if failed
' **Dependencies**: DataManager.SafeOpenWorkbook for file access
' **Side Effects**: Opens and closes enquiry file
' **Errors**: Returns empty structure if file access fails
Public Function LoadEnquiry(ByVal FilePath As String) As CoreFramework.EnquiryData
    Dim EnquiryWB As Workbook
    Dim ws As Worksheet
    Dim EnquiryInfo As CoreFramework.EnquiryData

    On Error GoTo Error_Handler

    Set EnquiryWB = DataManager.SafeOpenWorkbook(FilePath)
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

    DataManager.SafeCloseWorkbook EnquiryWB, False
    LoadEnquiry = EnquiryInfo
    Exit Function

Error_Handler:
    If Not EnquiryWB Is Nothing Then DataManager.SafeCloseWorkbook EnquiryWB, False
    CoreFramework.HandleStandardErrors Err.Number, "LoadEnquiry", "BusinessController"
End Function

' **Purpose**: Update existing enquiry with new data
' **Parameters**:
'   - EnquiryInfo (EnquiryData): Updated enquiry information
' **Returns**: Boolean - True if update successful, False if failed
' **Dependencies**: DataManager.SafeOpenWorkbook, PopulateEnquiryTemplate
' **Side Effects**: Modifies enquiry file, saves changes
' **Errors**: Returns False if file access or update fails
Public Function UpdateEnquiry(ByRef EnquiryInfo As CoreFramework.EnquiryData) As Boolean
    Dim EnquiryWB As Workbook

    On Error GoTo Error_Handler

    ' Validate data before updating
    If ValidateEnquiryData(EnquiryInfo) <> "" Then
        UpdateEnquiry = False
        Exit Function
    End If

    Set EnquiryWB = DataManager.SafeOpenWorkbook(EnquiryInfo.FilePath)
    If EnquiryWB Is Nothing Then
        UpdateEnquiry = False
        Exit Function
    End If

    PopulateEnquiryTemplate EnquiryWB, EnquiryInfo

    EnquiryWB.Save
    DataManager.SafeCloseWorkbook EnquiryWB

    ' Update search database
    Dim SearchRecord As CoreFramework.SearchRecord
    SearchRecord = SearchManager.CreateSearchRecord(CoreFramework.rtEnquiry, EnquiryInfo.EnquiryNumber, EnquiryInfo.CustomerName, EnquiryInfo.ComponentDescription, EnquiryInfo.FilePath, EnquiryInfo.SearchKeywords)
    SearchManager.UpdateSearchDatabase SearchRecord

    UpdateEnquiry = True
    Exit Function

Error_Handler:
    If Not EnquiryWB Is Nothing Then DataManager.SafeCloseWorkbook EnquiryWB, False
    CoreFramework.HandleStandardErrors Err.Number, "UpdateEnquiry", "BusinessController"
    UpdateEnquiry = False
End Function

' **Purpose**: Validate enquiry data completeness and business rules
' **Parameters**:
'   - EnquiryInfo (EnquiryData): Enquiry data to validate
' **Returns**: String - Validation error messages, empty if valid
' **Dependencies**: None
' **Side Effects**: None
' **Errors**: Returns error descriptions for invalid data
Public Function ValidateEnquiryData(ByRef EnquiryInfo As CoreFramework.EnquiryData) As String
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
' **Dependencies**: DataManager.SafeOpenWorkbook for template access
' **Side Effects**: Creates new customer file in Customers directory
' **Errors**: Returns False if template missing or file creation fails
Public Function CreateNewCustomer(ByVal CustomerName As String) As Boolean
    Dim TemplatePath As String
    Dim NewFilePath As String
    Dim TemplateWB As Workbook
    Dim CleanName As String

    On Error GoTo Error_Handler

    CleanName = CoreFramework.CleanFileName(CustomerName)
    TemplatePath = DataManager.GetRootPath & "\Templates\_client.xls"
    NewFilePath = DataManager.GetRootPath & "\Customers\" & CleanName & ".xls"

    If DataManager.FileExists(NewFilePath) Then
        CreateNewCustomer = True ' Already exists
        Exit Function
    End If

    If Not DataManager.FileExists(TemplatePath) Then
        CoreFramework.LogError CoreFramework.ERR_FILE_NOT_FOUND, "Customer template not found: " & TemplatePath, "CreateNewCustomer", "BusinessController"
        CreateNewCustomer = False
        Exit Function
    End If

    Set TemplateWB = DataManager.SafeOpenWorkbook(TemplatePath)
    If TemplateWB Is Nothing Then
        CreateNewCustomer = False
        Exit Function
    End If

    TemplateWB.Worksheets(1).Cells(1, 1).Value = CustomerName
    TemplateWB.Worksheets(1).Cells(1, 2).Value = Now ' Creation date

    TemplateWB.SaveAs NewFilePath
    DataManager.SafeCloseWorkbook TemplateWB

    CreateNewCustomer = True
    Exit Function

Error_Handler:
    If Not TemplateWB Is Nothing Then DataManager.SafeCloseWorkbook TemplateWB, False
    CoreFramework.HandleStandardErrors Err.Number, "CreateNewCustomer", "BusinessController"
    CreateNewCustomer = False
End Function

' **Purpose**: Archive completed enquiry
' **Parameters**:
'   - EnquiryInfo (EnquiryData): Enquiry to archive
' **Returns**: Boolean - True if archiving successful, False if failed
' **Dependencies**: DataManager.CreateBackup, DataManager.FileExists
' **Side Effects**: Moves enquiry file to archive directory
' **Errors**: Returns False if archive operation fails
Public Function ArchiveEnquiry(ByRef EnquiryInfo As CoreFramework.EnquiryData) As Boolean
    Dim SourcePath As String
    Dim ArchivePath As String

    On Error GoTo Error_Handler

    SourcePath = EnquiryInfo.FilePath
    ArchivePath = DataManager.GetRootPath & "\Archive\" & Dir(SourcePath)

    If Not DataManager.FileExists(SourcePath) Then
        ArchiveEnquiry = False
        Exit Function
    End If

    ' Create backup before moving
    DataManager.CreateBackup SourcePath

    ' Move file to archive
    FileCopy SourcePath, ArchivePath
    Kill SourcePath

    ' Update enquiry info with new path
    EnquiryInfo.FilePath = ArchivePath

    ArchiveEnquiry = True
    Exit Function

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "ArchiveEnquiry", "BusinessController"
    ArchiveEnquiry = False
End Function

' ===================================================================
' QUOTE MANAGEMENT (CLAUDE.md: Preserve Quote workflow)
' ===================================================================

' **Purpose**: Create quote from existing enquiry
' **Parameters**:
'   - EnquiryInfo (EnquiryData): Source enquiry information
'   - QuoteInfo (QuoteData): Quote information to populate
' **Returns**: Boolean - True if quote created successfully, False if failed
' **Dependencies**: DataManager.GetNextQuoteNumber, DataManager.SafeOpenWorkbook
' **Side Effects**: Creates new quote Excel file, updates search database
' **Errors**: Returns False on template missing or file creation failure
' **CLAUDE.md Compliance**: Maintains Enquiry → Quote → Jobs workflow
Public Function CreateQuoteFromEnquiry(ByRef EnquiryInfo As CoreFramework.EnquiryData, ByRef QuoteInfo As CoreFramework.QuoteData) As Boolean
    Dim QuoteNumber As String
    Dim TemplatePath As String
    Dim NewFilePath As String
    Dim TemplateWB As Workbook
    Dim SearchRecord As CoreFramework.SearchRecord

    On Error GoTo Error_Handler

    QuoteNumber = DataManager.GetNextQuoteNumber()
    If QuoteNumber = "" Then
        CreateQuoteFromEnquiry = False
        Exit Function
    End If

    ' Populate quote info from enquiry (only common fields)
    With QuoteInfo
        .QuoteNumber = QuoteNumber
        .EnquiryNumber = EnquiryInfo.EnquiryNumber
        .CustomerName = EnquiryInfo.CustomerName
        .ComponentDescription = EnquiryInfo.ComponentDescription
        .ComponentCode = EnquiryInfo.ComponentCode
        .MaterialGrade = EnquiryInfo.MaterialGrade
        .Quantity = EnquiryInfo.Quantity
        .DateCreated = Now
        .Status = "New Quote"
        .ValidUntil = DateAdd("d", 30, Now) ' Default 30 days validity
        ' Initialize quote-specific fields
        .UnitPrice = 0 ' To be set by user
        .TotalPrice = 0 ' To be calculated when UnitPrice is set
        .LeadTime = "TBD" ' To be determined by user
    End With

    TemplatePath = DataManager.GetRootPath & "\Templates\_Quote.xls"
    NewFilePath = DataManager.GetRootPath & "\Quotes\" & QuoteNumber & ".xls"

    If Not DataManager.FileExists(TemplatePath) Then
        CoreFramework.LogError CoreFramework.ERR_FILE_NOT_FOUND, "Quote template not found: " & TemplatePath, "CreateQuoteFromEnquiry", "BusinessController"
        CreateQuoteFromEnquiry = False
        Exit Function
    End If

    Set TemplateWB = DataManager.SafeOpenWorkbook(TemplatePath)
    If TemplateWB Is Nothing Then
        CreateQuoteFromEnquiry = False
        Exit Function
    End If

    PopulateQuoteTemplate TemplateWB, QuoteInfo

    TemplateWB.SaveAs NewFilePath
    DataManager.SafeCloseWorkbook TemplateWB

    QuoteInfo.FilePath = NewFilePath

    ' Update search database
    SearchRecord = SearchManager.CreateSearchRecord(CoreFramework.rtQuote, QuoteNumber, QuoteInfo.CustomerName, QuoteInfo.ComponentDescription, NewFilePath)
    SearchManager.UpdateSearchDatabase SearchRecord

    CreateQuoteFromEnquiry = True
    Exit Function

Error_Handler:
    If Not TemplateWB Is Nothing Then DataManager.SafeCloseWorkbook TemplateWB, False
    CoreFramework.HandleStandardErrors Err.Number, "CreateQuoteFromEnquiry", "BusinessController"
    CreateQuoteFromEnquiry = False
End Function

' **Purpose**: Load quote data from file
' **Parameters**:
'   - FilePath (String): Full path to quote file
' **Returns**: QuoteData - Populated quote structure, empty if failed
' **Dependencies**: DataManager.SafeOpenWorkbook for file access
' **Side Effects**: Opens and closes quote file
' **Errors**: Returns empty structure if file access fails
Public Function LoadQuote(ByVal FilePath As String) As CoreFramework.QuoteData
    Dim QuoteWB As Workbook
    Dim ws As Worksheet
    Dim QuoteInfo As CoreFramework.QuoteData

    On Error GoTo Error_Handler

    Set QuoteWB = DataManager.SafeOpenWorkbook(FilePath)
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

    DataManager.SafeCloseWorkbook QuoteWB, False
    LoadQuote = QuoteInfo
    Exit Function

Error_Handler:
    If Not QuoteWB Is Nothing Then DataManager.SafeCloseWorkbook QuoteWB, False
    CoreFramework.HandleStandardErrors Err.Number, "LoadQuote", "BusinessController"
End Function

' **Purpose**: Update existing quote with new data
' **Parameters**:
'   - QuoteInfo (QuoteData): Updated quote information
' **Returns**: Boolean - True if update successful, False if failed
' **Dependencies**: DataManager.SafeOpenWorkbook, PopulateQuoteTemplate
' **Side Effects**: Modifies quote file, saves changes
' **Errors**: Returns False if file access or update fails
Public Function UpdateQuote(ByRef QuoteInfo As CoreFramework.QuoteData) As Boolean
    Dim QuoteWB As Workbook

    On Error GoTo Error_Handler

    ' Validate data before updating
    If ValidateQuoteData(QuoteInfo) <> "" Then
        UpdateQuote = False
        Exit Function
    End If

    Set QuoteWB = DataManager.SafeOpenWorkbook(QuoteInfo.FilePath)
    If QuoteWB Is Nothing Then
        UpdateQuote = False
        Exit Function
    End If

    PopulateQuoteTemplate QuoteWB, QuoteInfo

    QuoteWB.Save
    DataManager.SafeCloseWorkbook QuoteWB

    ' Update search database
    Dim SearchRecord As CoreFramework.SearchRecord
    SearchRecord = SearchManager.CreateSearchRecord(CoreFramework.rtQuote, QuoteInfo.QuoteNumber, QuoteInfo.CustomerName, QuoteInfo.ComponentDescription, QuoteInfo.FilePath)
    SearchManager.UpdateSearchDatabase SearchRecord

    UpdateQuote = True
    Exit Function

Error_Handler:
    If Not QuoteWB Is Nothing Then DataManager.SafeCloseWorkbook QuoteWB, False
    CoreFramework.HandleStandardErrors Err.Number, "UpdateQuote", "BusinessController"
    UpdateQuote = False
End Function

' **Purpose**: Validate quote data completeness and business rules
' **Parameters**:
'   - QuoteInfo (QuoteData): Quote data to validate
' **Returns**: String - Validation error messages, empty if valid
' **Dependencies**: None
' **Side Effects**: None
' **Errors**: Returns error descriptions for invalid data
Public Function ValidateQuoteData(ByRef QuoteInfo As CoreFramework.QuoteData) As String
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

    If QuoteInfo.ValidUntil < Now Then
        ValidationErrors = ValidationErrors & "Quote validity date cannot be in the past." & vbCrLf
    End If

    ValidateQuoteData = ValidationErrors
End Function

' **Purpose**: Accept quote and prepare for job creation
' **Parameters**:
'   - QuoteInfo (QuoteData): Quote to accept
' **Returns**: Boolean - True if acceptance successful, False if failed
' **Dependencies**: UpdateQuote for status change
' **Side Effects**: Changes quote status to "Accepted"
' **Errors**: Returns False if status update fails
Public Function AcceptQuote(ByRef QuoteInfo As CoreFramework.QuoteData) As Boolean
    QuoteInfo.Status = "Quote Accepted"
    AcceptQuote = UpdateQuote(QuoteInfo)
End Function

' **Purpose**: Reject quote with reason
' **Parameters**:
'   - QuoteInfo (QuoteData): Quote to reject
'   - Reason (String, Optional): Reason for rejection
' **Returns**: Boolean - True if rejection successful, False if failed
' **Dependencies**: UpdateQuote for status change
' **Side Effects**: Changes quote status to "Rejected"
' **Errors**: Returns False if status update fails
Public Function RejectQuote(ByRef QuoteInfo As CoreFramework.QuoteData, Optional ByVal Reason As String = "") As Boolean
    QuoteInfo.Status = "Rejected" & IIf(Reason <> "", " - " & Reason, "")
    RejectQuote = UpdateQuote(QuoteInfo)
End Function

' ===================================================================
' JOB MANAGEMENT (CLAUDE.md: Preserve Jobs → Job Cards → WIP workflow)
' ===================================================================

' **Purpose**: Create job from accepted quote
' **Parameters**:
'   - QuoteInfo (QuoteData): Source quote information
'   - JobInfo (JobData): Job information to populate
' **Returns**: Boolean - True if job created successfully, False if failed
' **Dependencies**: DataManager.GetNextJobNumber, DataManager.SafeOpenWorkbook
' **Side Effects**: Creates new job Excel file, updates search database, creates WIP entry
' **Errors**: Returns False on template missing or file creation failure
' **CLAUDE.md Compliance**: Maintains Quote → Jobs → WIP workflow
Public Function CreateJobFromQuote(ByRef QuoteInfo As CoreFramework.QuoteData, ByRef JobInfo As CoreFramework.JobData) As Boolean
    Dim JobNumber As String
    Dim TemplatePath As String
    Dim NewFilePath As String
    Dim TemplateWB As Workbook
    Dim SearchRecord As CoreFramework.SearchRecord
    Dim MissingFields As String

    On Error GoTo Error_Handler

    ' Ensure quote is accepted before creating job
    If QuoteInfo.Status <> "Quote Accepted" Then
        CoreFramework.LogError 0, "Cannot create job from unaccepted quote: " & QuoteInfo.QuoteNumber, "CreateJobFromQuote", "BusinessController"
        CreateJobFromQuote = False
        Exit Function
    End If

    JobNumber = DataManager.GetNextJobNumber()
    If JobNumber = "" Then
        CreateJobFromQuote = False
        Exit Function
    End If

    ' Populate job info from quote (only common fields)
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
        .Status = "New Job"
        ' Calculate dates from quote LeadTime
        If IsNumeric(QuoteInfo.LeadTime) Then
            .DueDate = DateAdd("d", Val(QuoteInfo.LeadTime), Now)
        Else
            .DueDate = DateAdd("d", 14, Now) ' Default 14 days if LeadTime not numeric
        End If
        .WorkshopDueDate = DateAdd("d", -2, .DueDate) ' 2 days before customer due date
        .CustomerDueDate = .DueDate
        .FilePath = NewFilePath
        ' Initialize job-specific fields
        .AssignedOperator = "" ' To be assigned later
        .Operations = "" ' To be defined during job planning
        .Pictures = "" ' To be added during job execution
        .Notes = "" ' To be added as needed
    End With

    ' Check for empty/null transferred fields and notify user
    If Trim(QuoteInfo.CustomerName) = "" Then MissingFields = MissingFields & "Customer Name" & vbCrLf
    If Trim(QuoteInfo.ComponentDescription) = "" Then MissingFields = MissingFields & "Component Description" & vbCrLf
    If Trim(QuoteInfo.ComponentCode) = "" Then MissingFields = MissingFields & "Component Code" & vbCrLf
    If Trim(QuoteInfo.MaterialGrade) = "" Then MissingFields = MissingFields & "Material Grade" & vbCrLf
    If QuoteInfo.Quantity <= 0 Then MissingFields = MissingFields & "Valid Quantity" & vbCrLf
    If QuoteInfo.TotalPrice <= 0 Then MissingFields = MissingFields & "Valid Order Value" & vbCrLf

    ' Inform user about missing fields if any
    If MissingFields <> "" Then
        MsgBox "Job created successfully, but the following fields from the quote are empty or invalid:" & vbCrLf & vbCrLf & _
               MissingFields & vbCrLf & "Please update these fields in the job before proceeding to production.", _
               vbInformation + vbOKOnly, "Job Creation - Missing Fields"
    End If

    TemplatePath = DataManager.GetRootPath & "\Templates\_Job.xls"
    NewFilePath = DataManager.GetRootPath & "\WIP\" & JobNumber & ".xls"

    If Not DataManager.FileExists(TemplatePath) Then
        CoreFramework.LogError CoreFramework.ERR_FILE_NOT_FOUND, "Job template not found: " & TemplatePath, "CreateJobFromQuote", "BusinessController"
        CreateJobFromQuote = False
        Exit Function
    End If

    Set TemplateWB = DataManager.SafeOpenWorkbook(TemplatePath)
    If TemplateWB Is Nothing Then
        CreateJobFromQuote = False
        Exit Function
    End If

    PopulateJobTemplate TemplateWB, JobInfo

    TemplateWB.SaveAs NewFilePath
    DataManager.SafeCloseWorkbook TemplateWB

    JobInfo.FilePath = NewFilePath

    ' Update search database
    SearchRecord = SearchManager.CreateSearchRecord(CoreFramework.rtJob, JobNumber, JobInfo.CustomerName, JobInfo.ComponentDescription, NewFilePath)
    SearchManager.UpdateSearchDatabase SearchRecord

    ' Create WIP entry
    CreateWIPEntry JobInfo

    CreateJobFromQuote = True
    Exit Function

Error_Handler:
    If Not TemplateWB Is Nothing Then DataManager.SafeCloseWorkbook TemplateWB, False
    CoreFramework.HandleStandardErrors Err.Number, "CreateJobFromQuote", "BusinessController"
    CreateJobFromQuote = False
End Function

' **Purpose**: Load job data from file
' **Parameters**:
'   - FilePath (String): Full path to job file
' **Returns**: JobData - Populated job structure, empty if failed
' **Dependencies**: DataManager.SafeOpenWorkbook for file access
' **Side Effects**: Opens and closes job file
' **Errors**: Returns empty structure if file access fails
Public Function LoadJob(ByVal FilePath As String) As CoreFramework.JobData
    Dim JobWB As Workbook
    Dim ws As Worksheet
    Dim JobInfo As CoreFramework.JobData

    On Error GoTo Error_Handler

    Set JobWB = DataManager.SafeOpenWorkbook(FilePath)
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
        .DueDate = ws.Cells(9, 2).Value
        .WorkshopDueDate = ws.Cells(10, 2).Value
        .CustomerDueDate = ws.Cells(11, 2).Value
        .OrderValue = ws.Cells(12, 2).Value
        .DateCreated = ws.Cells(13, 2).Value
        .Status = ws.Cells(14, 2).Value
        .AssignedOperator = ws.Cells(15, 2).Value
        .Operations = ws.Cells(16, 2).Value
        .Pictures = ws.Cells(17, 2).Value
        .Notes = ws.Cells(18, 2).Value
        .FilePath = FilePath
    End With

    DataManager.SafeCloseWorkbook JobWB, False
    LoadJob = JobInfo
    Exit Function

Error_Handler:
    If Not JobWB Is Nothing Then DataManager.SafeCloseWorkbook JobWB, False
    CoreFramework.HandleStandardErrors Err.Number, "LoadJob", "BusinessController"
End Function

' **Purpose**: Update existing job with new data
' **Parameters**:
'   - JobInfo (JobData): Updated job information
' **Returns**: Boolean - True if update successful, False if failed
' **Dependencies**: DataManager.SafeOpenWorkbook, PopulateJobTemplate
' **Side Effects**: Modifies job file, saves changes, updates WIP
' **Errors**: Returns False if file access or update fails
Public Function UpdateJob(ByRef JobInfo As CoreFramework.JobData) As Boolean
    Dim JobWB As Workbook

    On Error GoTo Error_Handler

    ' Validate data before updating
    If ValidateJobData(JobInfo) <> "" Then
        UpdateJob = False
        Exit Function
    End If

    Set JobWB = DataManager.SafeOpenWorkbook(JobInfo.FilePath)
    If JobWB Is Nothing Then
        UpdateJob = False
        Exit Function
    End If

    PopulateJobTemplate JobWB, JobInfo

    JobWB.Save
    DataManager.SafeCloseWorkbook JobWB

    ' Update search database
    Dim SearchRecord As CoreFramework.SearchRecord
    SearchRecord = SearchManager.CreateSearchRecord(CoreFramework.rtJob, JobInfo.JobNumber, JobInfo.CustomerName, JobInfo.ComponentDescription, JobInfo.FilePath)
    SearchManager.UpdateSearchDatabase SearchRecord

    ' Update WIP entry
    UpdateWIPStatus JobInfo

    UpdateJob = True
    Exit Function

Error_Handler:
    If Not JobWB Is Nothing Then DataManager.SafeCloseWorkbook JobWB, False
    CoreFramework.HandleStandardErrors Err.Number, "UpdateJob", "BusinessController"
    UpdateJob = False
End Function

' **Purpose**: Validate job data completeness and business rules
' **Parameters**:
'   - JobInfo (JobData): Job data to validate
' **Returns**: String - Validation error messages, empty if valid
' **Dependencies**: None
' **Side Effects**: None
' **Errors**: Returns error descriptions for invalid data
Public Function ValidateJobData(ByRef JobInfo As CoreFramework.JobData) As String
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

    If JobInfo.DueDate < Now Then
        ValidationErrors = ValidationErrors & "Due date cannot be in the past." & vbCrLf
    End If

    ValidateJobData = ValidationErrors
End Function

' **Purpose**: Assign operator to job
' **Parameters**:
'   - JobInfo (JobData): Job to assign operator to
'   - OperatorName (String): Name of operator to assign
' **Returns**: Boolean - True if assignment successful, False if failed
' **Dependencies**: UpdateJob for data persistence
' **Side Effects**: Changes job assigned operator, updates WIP
' **Errors**: Returns False if update fails
Public Function AssignJobOperator(ByRef JobInfo As CoreFramework.JobData, ByVal OperatorName As String) As Boolean
    JobInfo.AssignedOperator = OperatorName
    JobInfo.Status = "Assigned"
    AssignJobOperator = UpdateJob(JobInfo)
End Function

' **Purpose**: Update job status
' **Parameters**:
'   - JobInfo (JobData): Job to update status for
'   - NewStatus (String): New status to set
' **Returns**: Boolean - True if update successful, False if failed
' **Dependencies**: UpdateJob for data persistence
' **Side Effects**: Changes job status, updates WIP
' **Errors**: Returns False if update fails
Public Function UpdateJobStatus(ByRef JobInfo As CoreFramework.JobData, ByVal NewStatus As String) As Boolean
    JobInfo.Status = NewStatus
    UpdateJobStatus = UpdateJob(JobInfo)
End Function

' **Purpose**: Complete job and move to archive
' **Parameters**:
'   - JobInfo (JobData): Job to complete
' **Returns**: Boolean - True if completion successful, False if failed
' **Dependencies**: UpdateJob, file system operations
' **Side Effects**: Changes job status, moves file to archive, updates WIP
' **Errors**: Returns False if completion process fails
Public Function CompleteJob(ByRef JobInfo As CoreFramework.JobData) As Boolean
    Dim ArchivePath As String

    On Error GoTo Error_Handler

    JobInfo.Status = "Completed"

    ' Update job before archiving
    If Not UpdateJob(JobInfo) Then
        CompleteJob = False
        Exit Function
    End If

    ' Move to archive
    ArchivePath = DataManager.GetRootPath & "\Archive\" & Dir(JobInfo.FilePath)
    FileCopy JobInfo.FilePath, ArchivePath
    Kill JobInfo.FilePath

    JobInfo.FilePath = ArchivePath

    ' Update search database with new path
    Dim SearchRecord As CoreFramework.SearchRecord
    SearchRecord = SearchManager.CreateSearchRecord(CoreFramework.rtJob, JobInfo.JobNumber, JobInfo.CustomerName, JobInfo.ComponentDescription, ArchivePath)
    SearchManager.UpdateSearchDatabase SearchRecord

    CompleteJob = True
    Exit Function

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "CompleteJob", "BusinessController"
    CompleteJob = False
End Function

' ===================================================================
' WIP MANAGEMENT (CLAUDE.md: Preserve WIP Reports workflow)
' ===================================================================

' **Purpose**: Create new WIP entry for job
' **Parameters**:
'   - JobInfo (JobData): Job information for WIP entry
' **Returns**: Boolean - True if WIP entry created successfully, False if failed
' **Dependencies**: DataManager.SafeOpenWorkbook for WIP database access
' **Side Effects**: Adds new row to WIP database
' **Errors**: Returns False if WIP database access fails
Public Function CreateWIPEntry(ByRef JobInfo As CoreFramework.JobData) As Boolean
    Dim WIPWB As Workbook
    Dim WIPWS As Worksheet
    Dim LastRow As Long

    On Error GoTo Error_Handler

    Set WIPWB = DataManager.SafeOpenWorkbook(DataManager.GetRootPath & "\" & WIP_FILE)
    If WIPWB Is Nothing Then
        CreateWIPEntry = False
        Exit Function
    End If

    Set WIPWS = WIPWB.Worksheets(1)
    LastRow = WIPWS.Cells(WIPWS.Rows.Count, 1).End(xlUp).Row + 1

    With WIPWS
        .Cells(LastRow, 1).Value = JobInfo.JobNumber
        .Cells(LastRow, 2).Value = JobInfo.CustomerName
        .Cells(LastRow, 3).Value = JobInfo.ComponentDescription
        .Cells(LastRow, 4).Value = JobInfo.Quantity
        .Cells(LastRow, 5).Value = JobInfo.DueDate
        .Cells(LastRow, 6).Value = JobInfo.AssignedOperator
        .Cells(LastRow, 7).Value = JobInfo.Status
        .Cells(LastRow, 8).Value = Now ' Last updated
        .Cells(LastRow, 9).Value = JobInfo.FilePath
    End With

    WIPWB.Save
    DataManager.SafeCloseWorkbook WIPWB

    CreateWIPEntry = True
    Exit Function

Error_Handler:
    If Not WIPWB Is Nothing Then DataManager.SafeCloseWorkbook WIPWB, False
    CoreFramework.HandleStandardErrors Err.Number, "CreateWIPEntry", "BusinessController"
    CreateWIPEntry = False
End Function

' **Purpose**: Update WIP status for job
' **Parameters**:
'   - JobInfo (JobData): Job with updated status
' **Returns**: Boolean - True if WIP update successful, False if failed
' **Dependencies**: DataManager.SafeOpenWorkbook for WIP database access
' **Side Effects**: Updates existing WIP row
' **Errors**: Returns False if WIP database access fails
Public Function UpdateWIPStatus(ByRef JobInfo As CoreFramework.JobData) As Boolean
    Dim WIPWB As Workbook
    Dim WIPWS As Worksheet
    Dim LastRow As Long
    Dim i As Long

    On Error GoTo Error_Handler

    Set WIPWB = DataManager.SafeOpenWorkbook(DataManager.GetRootPath & "\" & WIP_FILE)
    If WIPWB Is Nothing Then
        UpdateWIPStatus = False
        Exit Function
    End If

    Set WIPWS = WIPWB.Worksheets(1)
    LastRow = WIPWS.Cells(WIPWS.Rows.Count, 1).End(xlUp).Row

    ' Find and update existing WIP entry
    For i = 2 To LastRow
        If WIPWS.Cells(i, 1).Value = JobInfo.JobNumber Then
            With WIPWS
                .Cells(i, 2).Value = JobInfo.CustomerName
                .Cells(i, 3).Value = JobInfo.ComponentDescription
                .Cells(i, 4).Value = JobInfo.Quantity
                .Cells(i, 5).Value = JobInfo.DueDate
                .Cells(i, 6).Value = JobInfo.AssignedOperator
                .Cells(i, 7).Value = JobInfo.Status
                .Cells(i, 8).Value = Now ' Last updated
                .Cells(i, 9).Value = JobInfo.FilePath
            End With
            Exit For
        End If
    Next i

    WIPWB.Save
    DataManager.SafeCloseWorkbook WIPWB

    UpdateWIPStatus = True
    Exit Function

Error_Handler:
    If Not WIPWB Is Nothing Then DataManager.SafeCloseWorkbook WIPWB, False
    CoreFramework.HandleStandardErrors Err.Number, "UpdateWIPStatus", "BusinessController"
    UpdateWIPStatus = False
End Function

' **Purpose**: Save form data to WIP database
' **Parameters**:
'   - FormObject (Object): Form containing WIP data to save
' **Returns**: Boolean - True if save successful, False if failed
' **Dependencies**: DataManager.SafeOpenWorkbook for WIP database access
' **Side Effects**: Updates WIP database with form data
' **Errors**: Returns False if database access or save operation fails
' **CLAUDE.md Compliance**: Replaces legacy SaveWIPCode.bas functionality
Public Function SaveWIPData(ByRef FormObject As Object) As Boolean
    Dim WIPWB As Workbook
    Dim WIPWS As Worksheet
    Dim TargetRow As Long
    Dim ctl As Object
    Dim i As Integer

    On Error GoTo Error_Handler

    ' Open WIP database with retry for read-only
    Do
        Set WIPWB = DataManager.SafeOpenWorkbook(DataManager.GetRootPath & "\" & WIP_FILE)
        If WIPWB Is Nothing Then
            SaveWIPData = False
            Exit Function
        End If

        If WIPWB.ReadOnly = True Then
            DataManager.SafeCloseWorkbook WIPWB, False
            MsgBox "WIP database is read-only. Please ensure no other users have it open.", vbExclamation
            ' Could implement retry logic here
        End If
    Loop Until Not WIPWB.ReadOnly

    Set WIPWS = WIPWB.Worksheets(1)

    ' Find target row (existing record or new row)
    TargetRow = FindWIPRow(WIPWS, FormObject)

    ' Clear existing content for this row
    WIPWS.Rows(TargetRow).ClearContents

    ' Save form controls to WIP database
    For Each ctl In FormObject.Controls
        For i = 0 To 100
            If UCase(WIPWS.Range("A1").Offset(0, i).Value) = UCase(ctl.Name) Then
                Select Case UCase(TypeName(ctl))
                    Case "LABEL"
                        WIPWS.Range("A1").Offset(TargetRow - 1, i).Value = UCase(ctl.Caption)
                    Case "TEXTBOX"
                        WIPWS.Range("A1").Offset(TargetRow - 1, i).Value = UCase(ctl.Value)
                    Case "COMBOBOX"
                        WIPWS.Range("A1").Offset(TargetRow - 1, i).Value = UCase(ctl.Value)
                End Select
                Exit For
            End If
            ' Copy formula from previous row if needed
            If Left(WIPWS.Range("A1").Offset(TargetRow - 2, i).Value, 1) = "=" Then
                WIPWS.Range("A1").Offset(TargetRow - 1, i).Value = WIPWS.Range("A1").Offset(TargetRow - 2, i).Value
            End If
            If WIPWS.Range("A1").Offset(0, i + 1).Value = "" Then Exit For
        Next i
    Next ctl

    WIPWB.Save
    DataManager.SafeCloseWorkbook WIPWB

    SaveWIPData = True
    Exit Function

Error_Handler:
    If Not WIPWB Is Nothing Then DataManager.SafeCloseWorkbook WIPWB, False
    CoreFramework.HandleStandardErrors Err.Number, "SaveWIPData", "BusinessController"
    SaveWIPData = False
End Function

' **Purpose**: Generate WIP report for specified criteria
' **Parameters**:
'   - ReportType (String, Optional): Type of report (All, ByOperator, ByDueDate)
'   - FilterValue (String, Optional): Filter value for report
' **Returns**: Boolean - True if report generation successful, False if failed
' **Dependencies**: DataManager.SafeOpenWorkbook for WIP database access
' **Side Effects**: Creates new report file
' **Errors**: Returns False if report generation fails
Public Function GenerateWIPReport(Optional ByVal ReportType As String = "All", Optional ByVal FilterValue As String = "") As Boolean
    Dim WIPWB As Workbook
    Dim ReportWB As Workbook
    Dim WIPWS As Worksheet
    Dim ReportWS As Worksheet
    Dim ReportPath As String

    On Error GoTo Error_Handler

    Set WIPWB = DataManager.SafeOpenWorkbook(DataManager.GetRootPath & "\" & WIP_FILE)
    If WIPWB Is Nothing Then
        GenerateWIPReport = False
        Exit Function
    End If

    Set ReportWB = DataManager.CreateNewWorkbook()
    If ReportWB Is Nothing Then
        DataManager.SafeCloseWorkbook WIPWB, False
        GenerateWIPReport = False
        Exit Function
    End If

    Set WIPWS = WIPWB.Worksheets(1)
    Set ReportWS = ReportWB.Worksheets(1)

    ' Generate report based on type
    Select Case UCase(ReportType)
        Case "ALL"
            WIPWS.UsedRange.Copy ReportWS.Range("A1")
        Case "BYOPERATOR"
            FilterWIPByOperator WIPWS, ReportWS, FilterValue
        Case "BYDUEDATE"
            FilterWIPByDueDate WIPWS, ReportWS, FilterValue
        Case Else
            WIPWS.UsedRange.Copy ReportWS.Range("A1")
    End Select

    ' Save report
    ReportPath = DataManager.GetRootPath & "\Reports\WIP_Report_" & Format(Now, "yyyymmdd_hhmmss") & ".xls"
    ReportWB.SaveAs ReportPath

    DataManager.SafeCloseWorkbook WIPWB, False
    DataManager.SafeCloseWorkbook ReportWB

    GenerateWIPReport = True
    Exit Function

Error_Handler:
    If Not WIPWB Is Nothing Then DataManager.SafeCloseWorkbook WIPWB, False
    If Not ReportWB Is Nothing Then DataManager.SafeCloseWorkbook ReportWB, False
    CoreFramework.HandleStandardErrors Err.Number, "GenerateWIPReport", "BusinessController"
    GenerateWIPReport = False
End Function

' **Purpose**: Archive completed WIP entries
' **Parameters**: None
' **Returns**: Boolean - True if archiving successful, False if failed
' **Dependencies**: DataManager.SafeOpenWorkbook for WIP database access
' **Side Effects**: Removes completed entries from WIP database
' **Errors**: Returns False if archiving operation fails
Public Function ArchiveCompletedWIP() As Boolean
    Dim WIPWB As Workbook
    Dim WIPWS As Worksheet
    Dim LastRow As Long
    Dim i As Long

    On Error GoTo Error_Handler

    Set WIPWB = DataManager.SafeOpenWorkbook(DataManager.GetRootPath & "\" & WIP_FILE)
    If WIPWB Is Nothing Then
        ArchiveCompletedWIP = False
        Exit Function
    End If

    Set WIPWS = WIPWB.Worksheets(1)
    LastRow = WIPWS.Cells(WIPWS.Rows.Count, 1).End(xlUp).Row

    ' Remove completed jobs (work backwards to avoid index issues)
    For i = LastRow To 2 Step -1
        If UCase(WIPWS.Cells(i, 7).Value) = "COMPLETED" Or UCase(WIPWS.Cells(i, 7).Value) = "CANCELLED" Then
            WIPWS.Rows(i).Delete
        End If
    Next i

    WIPWB.Save
    DataManager.SafeCloseWorkbook WIPWB

    ArchiveCompletedWIP = True
    Exit Function

Error_Handler:
    If Not WIPWB Is Nothing Then DataManager.SafeCloseWorkbook WIPWB, False
    CoreFramework.HandleStandardErrors Err.Number, "ArchiveCompletedWIP", "BusinessController"
    ArchiveCompletedWIP = False
End Function

' **Purpose**: Get WIP entries by operator
' **Parameters**:
'   - OperatorName (String): Name of operator to filter by
' **Returns**: Variant - Array of WIP entries for operator
' **Dependencies**: DataManager.SafeOpenWorkbook for WIP database access
' **Side Effects**: None
' **Errors**: Returns empty array if retrieval fails
Public Function GetWIPByOperator(ByVal OperatorName As String) As Variant
    Dim WIPWB As Workbook
    Dim WIPWS As Worksheet
    Dim LastRow As Long
    Dim i As Long
    Dim Results() As String
    Dim ResultCount As Integer

    On Error GoTo Error_Handler

    Set WIPWB = DataManager.SafeOpenWorkbook(DataManager.GetRootPath & "\" & WIP_FILE)
    If WIPWB Is Nothing Then
        GetWIPByOperator = Array()
        Exit Function
    End If

    Set WIPWS = WIPWB.Worksheets(1)
    LastRow = WIPWS.Cells(WIPWS.Rows.Count, 1).End(xlUp).Row
    ResultCount = 0

    For i = 2 To LastRow
        If UCase(WIPWS.Cells(i, 6).Value) = UCase(OperatorName) Then
            ReDim Preserve Results(ResultCount)
            Results(ResultCount) = WIPWS.Cells(i, 1).Value ' Job Number
            ResultCount = ResultCount + 1
        End If
    Next i

    DataManager.SafeCloseWorkbook WIPWB, False

    If ResultCount > 0 Then
        GetWIPByOperator = Results
    Else
        GetWIPByOperator = Array()
    End If
    Exit Function

Error_Handler:
    If Not WIPWB Is Nothing Then DataManager.SafeCloseWorkbook WIPWB, False
    CoreFramework.HandleStandardErrors Err.Number, "GetWIPByOperator", "BusinessController"
    GetWIPByOperator = Array()
End Function

' **Purpose**: Get WIP entries by due date
' **Parameters**:
'   - DueDate (Date): Due date to filter by
' **Returns**: Variant - Array of WIP entries for due date
' **Dependencies**: DataManager.SafeOpenWorkbook for WIP database access
' **Side Effects**: None
' **Errors**: Returns empty array if retrieval fails
Public Function GetWIPByDueDate(ByVal DueDate As Date) As Variant
    Dim WIPWB As Workbook
    Dim WIPWS As Worksheet
    Dim LastRow As Long
    Dim i As Long
    Dim Results() As String
    Dim ResultCount As Integer

    On Error GoTo Error_Handler

    Set WIPWB = DataManager.SafeOpenWorkbook(DataManager.GetRootPath & "\" & WIP_FILE)
    If WIPWB Is Nothing Then
        GetWIPByDueDate = Array()
        Exit Function
    End If

    Set WIPWS = WIPWB.Worksheets(1)
    LastRow = WIPWS.Cells(WIPWS.Rows.Count, 1).End(xlUp).Row
    ResultCount = 0

    For i = 2 To LastRow
        If DateValue(WIPWS.Cells(i, 5).Value) <= DueDate Then
            ReDim Preserve Results(ResultCount)
            Results(ResultCount) = WIPWS.Cells(i, 1).Value ' Job Number
            ResultCount = ResultCount + 1
        End If
    Next i

    DataManager.SafeCloseWorkbook WIPWB, False

    If ResultCount > 0 Then
        GetWIPByDueDate = Results
    Else
        GetWIPByDueDate = Array()
    End If
    Exit Function

Error_Handler:
    If Not WIPWB Is Nothing Then DataManager.SafeCloseWorkbook WIPWB, False
    CoreFramework.HandleStandardErrors Err.Number, "GetWIPByDueDate", "BusinessController"
    GetWIPByDueDate = Array()
End Function

' ===================================================================
' WORKFLOW ORCHESTRATION
' ===================================================================

' **Purpose**: Process enquiry to quote workflow transition
' **Parameters**:
'   - EnquiryInfo (EnquiryData): Source enquiry
'   - QuoteInfo (QuoteData): Target quote to create
' **Returns**: Boolean - True if transition successful, False if failed
' **Dependencies**: CreateQuoteFromEnquiry
' **Side Effects**: Creates quote file, updates search database
' **Errors**: Returns False if workflow transition fails
Public Function ProcessEnquiryToQuote(ByRef EnquiryInfo As CoreFramework.EnquiryData, ByRef QuoteInfo As CoreFramework.QuoteData) As Boolean
    ProcessEnquiryToQuote = CreateQuoteFromEnquiry(EnquiryInfo, QuoteInfo)
End Function

' **Purpose**: Process quote to job workflow transition
' **Parameters**:
'   - QuoteInfo (QuoteData): Source quote (must be accepted)
'   - JobInfo (JobData): Target job to create
' **Returns**: Boolean - True if transition successful, False if failed
' **Dependencies**: CreateJobFromQuote
' **Side Effects**: Creates job file, updates search database, creates WIP entry
' **Errors**: Returns False if workflow transition fails or quote not accepted
Public Function ProcessQuoteToJob(ByRef QuoteInfo As CoreFramework.QuoteData, ByRef JobInfo As CoreFramework.JobData) As Boolean
    ProcessQuoteToJob = CreateJobFromQuote(QuoteInfo, JobInfo)
End Function

' **Purpose**: Process job to archive workflow transition
' **Parameters**:
'   - JobInfo (JobData): Job to complete and archive
' **Returns**: Boolean - True if transition successful, False if failed
' **Dependencies**: CompleteJob
' **Side Effects**: Archives job file, updates WIP status
' **Errors**: Returns False if workflow transition fails
Public Function ProcessJobToArchive(ByRef JobInfo As CoreFramework.JobData) As Boolean
    ProcessJobToArchive = CompleteJob(JobInfo)
End Function

' **Purpose**: Validate workflow transition is allowed
' **Parameters**:
'   - FromState (String): Current workflow state
'   - ToState (String): Target workflow state
' **Returns**: Boolean - True if transition allowed, False if not
' **Dependencies**: None
' **Side Effects**: None
' **Errors**: Returns False for invalid transitions
Public Function ValidateWorkflowTransition(ByVal FromState As String, ByVal ToState As String) As Boolean
    ' Define valid transitions
    Select Case UCase(FromState)
        Case "ENQUIRY"
            ValidateWorkflowTransition = (UCase(ToState) = "QUOTE")
        Case "QUOTE"
            ValidateWorkflowTransition = (UCase(ToState) = "JOB" Or UCase(ToState) = "REJECTED")
        Case "JOB"
            ValidateWorkflowTransition = (UCase(ToState) = "COMPLETED" Or UCase(ToState) = "CANCELLED")
        Case Else
            ValidateWorkflowTransition = False
    End Select
End Function

' ===================================================================
' CONTRACT MANAGEMENT (CLAUDE.md: Preserve Contract functionality)
' ===================================================================

' **Purpose**: Load contract template data
' **Parameters**:
'   - FilePath (String): Full path to contract file
' **Returns**: ContractData - Populated contract structure, empty if failed
' **Dependencies**: DataManager.SafeOpenWorkbook for file access
' **Side Effects**: Opens and closes contract file
' **Errors**: Returns empty structure if file access fails
Public Function LoadContract(ByVal FilePath As String) As CoreFramework.ContractData
    Dim ContractWB As Workbook
    Dim ws As Worksheet
    Dim ContractInfo As CoreFramework.ContractData

    On Error GoTo Error_Handler

    Set ContractWB = DataManager.SafeOpenWorkbook(FilePath)
    If ContractWB Is Nothing Then
        Exit Function
    End If

    Set ws = ContractWB.Worksheets(1)

    With ContractInfo
        .ContractName = ws.Cells(2, 2).Value
        .CustomerName = ws.Cells(3, 2).Value
        .ComponentDescription = ws.Cells(4, 2).Value
        .StandardOperations = ws.Cells(5, 2).Value
        .LeadTime = ws.Cells(6, 2).Value
        .DateCreated = ws.Cells(7, 2).Value
        .LastUsed = ws.Cells(8, 2).Value
        .FilePath = FilePath
    End With

    DataManager.SafeCloseWorkbook ContractWB, False
    LoadContract = ContractInfo
    Exit Function

Error_Handler:
    If Not ContractWB Is Nothing Then DataManager.SafeCloseWorkbook ContractWB, False
    CoreFramework.HandleStandardErrors Err.Number, "LoadContract", "BusinessController"
End Function

' **Purpose**: Create job from contract template
' **Parameters**:
'   - ContractInfo (ContractData): Source contract template
'   - JobInfo (JobData): Target job to create
' **Returns**: Boolean - True if job created successfully, False if failed
' **Dependencies**: DataManager.GetNextJobNumber, CreateWIPEntry
' **Side Effects**: Creates job file, updates contract usage, creates WIP entry
' **Errors**: Returns False if job creation fails
Public Function CreateJobFromContract(ByRef ContractInfo As CoreFramework.ContractData, ByRef JobInfo As CoreFramework.JobData) As Boolean
    Dim JobNumber As String
    Dim TemplatePath As String
    Dim NewFilePath As String
    Dim TemplateWB As Workbook
    Dim MissingFields As String

    On Error GoTo Error_Handler

    JobNumber = DataManager.GetNextJobNumber()
    If JobNumber = "" Then
        CreateJobFromContract = False
        Exit Function
    End If

    ' Use contract as template
    TemplatePath = ContractInfo.FilePath
    NewFilePath = DataManager.GetRootPath & "\WIP\" & JobNumber & ".xls"

    Set TemplateWB = DataManager.SafeOpenWorkbook(TemplatePath)
    If TemplateWB Is Nothing Then
        CreateJobFromContract = False
        Exit Function
    End If

    ' Populate job info from contract (only common fields)
    With JobInfo
        .JobNumber = JobNumber
        .CustomerName = ContractInfo.CustomerName
        .ComponentDescription = ContractInfo.ComponentDescription
        .Operations = ContractInfo.StandardOperations
        .DateCreated = Now
        .Status = "New Job"
        ' Calculate dates from contract LeadTime
        If IsNumeric(ContractInfo.LeadTime) Then
            .DueDate = DateAdd("d", Val(ContractInfo.LeadTime), Now)
        Else
            .DueDate = DateAdd("d", 14, Now) ' Default 14 days if LeadTime not numeric
            If Trim(ContractInfo.LeadTime) <> "" And Not IsNumeric(ContractInfo.LeadTime) Then
                MissingFields = MissingFields & "Lead Time (not numeric: " & ContractInfo.LeadTime & ")" & vbCrLf
            End If
        End If
        .WorkshopDueDate = DateAdd("d", -2, .DueDate)
        .CustomerDueDate = .DueDate
        .FilePath = NewFilePath
        ' Initialize remaining job fields
        .QuoteNumber = "" ' No source quote for contract jobs
        .ComponentCode = "" ' To be defined
        .MaterialGrade = "" ' To be defined
        .Quantity = 1 ' Default quantity, to be updated
        .OrderValue = 0 ' To be set when pricing is determined
        .AssignedOperator = "" ' To be assigned later
        .Pictures = "" ' To be added during job execution
        .Notes = "Created from contract template: " & ContractInfo.ContractName
    End With

    ' Check for empty/null transferred fields and notify user
    If Trim(ContractInfo.CustomerName) = "" Then MissingFields = MissingFields & "Customer Name" & vbCrLf
    If Trim(ContractInfo.ComponentDescription) = "" Then MissingFields = MissingFields & "Component Description" & vbCrLf
    If Trim(ContractInfo.StandardOperations) = "" Then MissingFields = MissingFields & "Standard Operations" & vbCrLf

    ' Inform user about missing fields if any
    If MissingFields <> "" Then
        MsgBox "Job created successfully from contract, but the following fields are empty or invalid:" & vbCrLf & vbCrLf & _
               MissingFields & vbCrLf & "Please update these fields in the job before proceeding to production.", _
               vbInformation + vbOKOnly, "Job Creation from Contract - Missing Fields"
    End If

    PopulateJobTemplate TemplateWB, JobInfo

    TemplateWB.SaveAs NewFilePath
    DataManager.SafeCloseWorkbook TemplateWB

    ' Update contract usage
    UpdateContractUsage ContractInfo

    ' Create WIP entry
    CreateWIPEntry JobInfo

    CreateJobFromContract = True
    Exit Function

Error_Handler:
    If Not TemplateWB Is Nothing Then DataManager.SafeCloseWorkbook TemplateWB, False
    CoreFramework.HandleStandardErrors Err.Number, "CreateJobFromContract", "BusinessController"
    CreateJobFromContract = False
End Function

' **Purpose**: Update contract last used date
' **Parameters**:
'   - ContractInfo (ContractData): Contract to update usage for
' **Returns**: Boolean - True if update successful, False if failed
' **Dependencies**: DataManager.SafeOpenWorkbook for file access
' **Side Effects**: Updates contract file with last used date
' **Errors**: Returns False if file access or update fails
Public Function UpdateContractUsage(ByRef ContractInfo As CoreFramework.ContractData) As Boolean
    Dim ContractWB As Workbook
    Dim ws As Worksheet

    On Error GoTo Error_Handler

    Set ContractWB = DataManager.SafeOpenWorkbook(ContractInfo.FilePath)
    If ContractWB Is Nothing Then
        UpdateContractUsage = False
        Exit Function
    End If

    Set ws = ContractWB.Worksheets(1)
    ws.Cells(8, 2).Value = Now ' Update last used date

    ContractWB.Save
    DataManager.SafeCloseWorkbook ContractWB

    ContractInfo.LastUsed = Now
    UpdateContractUsage = True
    Exit Function

Error_Handler:
    If Not ContractWB Is Nothing Then DataManager.SafeCloseWorkbook ContractWB, False
    CoreFramework.HandleStandardErrors Err.Number, "UpdateContractUsage", "BusinessController"
    UpdateContractUsage = False
End Function

' ===================================================================
' PRIVATE HELPER FUNCTIONS
' ===================================================================

' **Purpose**: Populate enquiry template with data
' **Parameters**:
'   - wb (Workbook): Template workbook to populate
'   - EnquiryInfo (EnquiryData): Data to populate template with
' **Returns**: None (Subroutine)
' **Dependencies**: None
' **Side Effects**: Modifies worksheet cells with enquiry data
' **Errors**: Logs errors if population fails
Private Sub PopulateEnquiryTemplate(ByRef wb As Workbook, ByRef EnquiryInfo As CoreFramework.EnquiryData)
    Dim ws As Worksheet

    On Error GoTo Error_Handler

    Set ws = wb.Worksheets(1)

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
    CoreFramework.LogError Err.Number, Err.Description, "PopulateEnquiryTemplate", "BusinessController"
End Sub

' **Purpose**: Populate quote template with data
' **Parameters**:
'   - wb (Workbook): Template workbook to populate
'   - QuoteInfo (QuoteData): Data to populate template with
' **Returns**: None (Subroutine)
' **Dependencies**: None
' **Side Effects**: Modifies worksheet cells with quote data
' **Errors**: Logs errors if population fails
Private Sub PopulateQuoteTemplate(ByRef wb As Workbook, ByRef QuoteInfo As CoreFramework.QuoteData)
    Dim ws As Worksheet

    On Error GoTo Error_Handler

    Set ws = wb.Worksheets(1)

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
    CoreFramework.LogError Err.Number, Err.Description, "PopulateQuoteTemplate", "BusinessController"
End Sub

' **Purpose**: Populate job template with data
' **Parameters**:
'   - wb (Workbook): Template workbook to populate
'   - JobInfo (JobData): Data to populate template with
' **Returns**: None (Subroutine)
' **Dependencies**: None
' **Side Effects**: Modifies worksheet cells with job data
' **Errors**: Logs errors if population fails
Private Sub PopulateJobTemplate(ByRef wb As Workbook, ByRef JobInfo As CoreFramework.JobData)
    Dim ws As Worksheet

    On Error GoTo Error_Handler

    Set ws = wb.Worksheets(1)

    With ws
        .Cells(2, 2).Value = JobInfo.JobNumber
        .Cells(3, 2).Value = JobInfo.QuoteNumber
        .Cells(4, 2).Value = JobInfo.CustomerName
        .Cells(5, 2).Value = JobInfo.ComponentDescription
        .Cells(6, 2).Value = JobInfo.ComponentCode
        .Cells(7, 2).Value = JobInfo.MaterialGrade
        .Cells(8, 2).Value = JobInfo.Quantity
        .Cells(9, 2).Value = JobInfo.DueDate
        .Cells(10, 2).Value = JobInfo.WorkshopDueDate
        .Cells(11, 2).Value = JobInfo.CustomerDueDate
        .Cells(12, 2).Value = JobInfo.OrderValue
        .Cells(13, 2).Value = JobInfo.DateCreated
        .Cells(14, 2).Value = JobInfo.Status
        .Cells(15, 2).Value = JobInfo.AssignedOperator
        .Cells(16, 2).Value = JobInfo.Operations
        .Cells(17, 2).Value = JobInfo.Pictures
        .Cells(18, 2).Value = JobInfo.Notes
    End With

    Exit Sub

Error_Handler:
    CoreFramework.LogError Err.Number, Err.Description, "PopulateJobTemplate", "BusinessController"
End Sub

' **Purpose**: Find WIP row for form data or determine where to add new row
' **Parameters**:
'   - WIPWS (Worksheet): WIP worksheet to scan
'   - FormObject (Object): Form containing record identifiers
' **Returns**: Long - Row number for data placement
' **Dependencies**: None
' **Side Effects**: None
' **Errors**: Returns next available row if search fails
Private Function FindWIPRow(ByRef WIPWS As Worksheet, ByRef FormObject As Object) As Long
    Dim i As Long
    Dim LastRow As Long
    Dim QuoteNumber As String
    Dim EnquiryNumber As String
    Dim JobNumber As String
    Dim FileName As String

    On Error GoTo Error_Handler

    ' Get identifiers from form
    On Error Resume Next
    QuoteNumber = FormObject.Controls("Quote_Number").Value
    EnquiryNumber = FormObject.Controls("Enquiry_Number").Value
    JobNumber = FormObject.Controls("Job_Number").Value
    FileName = FormObject.Controls("File_Name").Value
    On Error GoTo Error_Handler

    LastRow = WIPWS.Cells(WIPWS.Rows.Count, 1).End(xlUp).Row

    ' Look for existing record
    For i = 2 To LastRow
        If WIPWS.Cells(i, 3).Value = QuoteNumber Or _
           WIPWS.Cells(i, 3).Value = EnquiryNumber Or _
           WIPWS.Cells(i, 3).Value = JobNumber Or _
           WIPWS.Cells(i, 3).Value = FileName Then
            FindWIPRow = i
            Exit Function
        End If
    Next i

    ' Return next available row
    FindWIPRow = LastRow + 1
    Exit Function

Error_Handler:
    FindWIPRow = WIPWS.Cells(WIPWS.Rows.Count, 1).End(xlUp).Row + 1
End Function

' **Purpose**: Filter WIP data by operator for reporting
' **Parameters**:
'   - SourceWS (Worksheet): Source WIP worksheet
'   - TargetWS (Worksheet): Target report worksheet
'   - OperatorName (String): Operator to filter by
' **Returns**: None (Subroutine)
' **Dependencies**: None
' **Side Effects**: Copies filtered data to target worksheet
' **Errors**: Continues silently if filtering fails
Private Sub FilterWIPByOperator(ByRef SourceWS As Worksheet, ByRef TargetWS As Worksheet, ByVal OperatorName As String)
    Dim i As Long
    Dim LastRow As Long
    Dim TargetRow As Long

    On Error GoTo Error_Handler

    LastRow = SourceWS.Cells(SourceWS.Rows.Count, 1).End(xlUp).Row
    TargetRow = 1

    ' Copy headers
    SourceWS.Rows(1).Copy TargetWS.Rows(1)
    TargetRow = 2

    ' Copy matching rows
    For i = 2 To LastRow
        If UCase(SourceWS.Cells(i, 6).Value) = UCase(OperatorName) Then
            SourceWS.Rows(i).Copy TargetWS.Rows(TargetRow)
            TargetRow = TargetRow + 1
        End If
    Next i

    Exit Sub

Error_Handler:
    ' Filter failed - continue silently
End Sub

' **Purpose**: Filter WIP data by due date for reporting
' **Parameters**:
'   - SourceWS (Worksheet): Source WIP worksheet
'   - TargetWS (Worksheet): Target report worksheet
'   - DueDateFilter (String): Due date filter criteria
' **Returns**: None (Subroutine)
' **Dependencies**: None
' **Side Effects**: Copies filtered data to target worksheet
' **Errors**: Continues silently if filtering fails
Private Sub FilterWIPByDueDate(ByRef SourceWS As Worksheet, ByRef TargetWS As Worksheet, ByVal DueDateFilter As String)
    Dim i As Long
    Dim LastRow As Long
    Dim TargetRow As Long
    Dim FilterDate As Date

    On Error GoTo Error_Handler

    FilterDate = CDate(DueDateFilter)
    LastRow = SourceWS.Cells(SourceWS.Rows.Count, 1).End(xlUp).Row
    TargetRow = 1

    ' Copy headers
    SourceWS.Rows(1).Copy TargetWS.Rows(1)
    TargetRow = 2

    ' Copy matching rows
    For i = 2 To LastRow
        If DateValue(SourceWS.Cells(i, 5).Value) <= FilterDate Then
            SourceWS.Rows(i).Copy TargetWS.Rows(TargetRow)
            TargetRow = TargetRow + 1
        End If
    Next i

    Exit Sub

Error_Handler:
    ' Filter failed - continue silently
End Sub

' ===================================================================
' SYSTEM INITIALIZATION FUNCTIONS
' ===================================================================

' **Purpose**: Initialize empty WIP database with proper structure
' **Parameters**: None
' **Returns**: Boolean - True if initialization successful, False if error
' **Dependencies**: DataManager.GetRootPath, Excel Application object
' **Side Effects**: Creates WIP.xls file with header row
' **Errors**: Returns False on file creation failure, logs error
Public Function InitializeWIPDatabase() As Boolean
    Dim NewWB As Workbook
    Dim WS As Worksheet
    Dim FilePath As String

    On Error GoTo Error_Handler

    FilePath = DataManager.GetRootPath & "\WIP.xls"

    ' Create new workbook
    Set NewWB = Application.Workbooks.Add
    Set WS = NewWB.Worksheets(1)

    ' Set up headers based on CreateWIPEntry structure
    With WS
        .Cells(1, 1).Value = "Job Number"
        .Cells(1, 2).Value = "Customer Name"
        .Cells(1, 3).Value = "Component Description"
        .Cells(1, 4).Value = "Quantity"
        .Cells(1, 5).Value = "Due Date"
        .Cells(1, 6).Value = "Assigned Operator"
        .Cells(1, 7).Value = "Status"
        .Cells(1, 8).Value = "Last Updated"
        .Cells(1, 9).Value = "File Path"

        ' Format headers
        .Range("A1:I1").Font.Bold = True
        .Range("A1:I1").Interior.Color = RGB(200, 200, 200)
        .Columns("A:I").AutoFit
    End With

    ' Save and close
    NewWB.SaveAs FilePath, FileFormat:=xlExcel8
    NewWB.Close SaveChanges:=False
    Set NewWB = Nothing
    Set WS = Nothing

    InitializeWIPDatabase = True
    CoreFramework.LogError 0, "WIP database initialized successfully", "InitializeWIPDatabase", "BusinessController"
    Exit Function

Error_Handler:
    CoreFramework.LogError Err.Number, Err.Description, "InitializeWIPDatabase", "BusinessController"
    If Not NewWB Is Nothing Then
        NewWB.Close SaveChanges:=False
        Set NewWB = Nothing
    End If
    InitializeWIPDatabase = False
End Function