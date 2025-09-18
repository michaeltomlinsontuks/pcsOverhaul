Attribute VB_Name = "JobController"
Option Explicit

Public Function CreateJobFromQuote(ByRef JobInfo As JobData) As Boolean
    Dim JobNumber As String
    Dim NewFilePath As String
    Dim SearchRecord As SearchRecord

    On Error GoTo Error_Handler

    JobNumber = NumberGenerator.GetNextJobNumber()
    If JobNumber = "" Then
        CreateJobFromQuote = False
        Exit Function
    End If

    JobInfo.JobNumber = JobNumber
    JobInfo.DateCreated = Now
    JobInfo.Status = "Active"

    NewFilePath = FileManager.GetRootPath & "\WIP\" & JobNumber & ".xls"

    If CreateJobFile(NewFilePath, JobInfo) Then
        JobInfo.FilePath = NewFilePath

        WIPManager.AddJobToWIP JobInfo

        SearchRecord = SearchService.CreateSearchRecord(rtJob, JobNumber, JobInfo.CustomerName, JobInfo.ComponentDescription, NewFilePath)
        SearchService.UpdateSearchDatabase SearchRecord

        CreateJobFromQuote = True
    Else
        CreateJobFromQuote = False
    End If

    Exit Function

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "CreateJobFromQuote", "JobController"
    CreateJobFromQuote = False
End Function

Public Function CreateDirectJob(ByVal JobInfo As JobData) As Boolean
    Dim JobNumber As String
    Dim NewFilePath As String
    Dim SearchRecord As SearchRecord

    On Error GoTo Error_Handler

    JobNumber = NumberGenerator.GetNextJobNumber()
    If JobNumber = "" Then
        CreateDirectJob = False
        Exit Function
    End If

    JobInfo.JobNumber = JobNumber
    JobInfo.DateCreated = Now

    NewFilePath = FileManager.GetRootPath & "\WIP\" & JobNumber & ".xls"

    If CreateJobFile(NewFilePath, JobInfo) Then
        JobInfo.FilePath = NewFilePath

        WIPManager.AddJobToWIP JobInfo

        SearchRecord = SearchService.CreateSearchRecord(rtJob, JobNumber, JobInfo.CustomerName, JobInfo.ComponentDescription, NewFilePath)
        SearchService.UpdateSearchDatabase SearchRecord

        CreateDirectJob = True
    Else
        CreateDirectJob = False
    End If

    Exit Function

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "CreateDirectJob", "JobController"
    CreateDirectJob = False
End Function

Private Function CreateJobFile(ByVal FilePath As String, ByVal JobInfo As JobData) As Boolean
    Dim TemplateWB As Workbook
    Dim TemplatePath As String

    On Error GoTo Error_Handler

    TemplatePath = FileManager.GetRootPath & "\Templates\_Enq.xls"

    Set TemplateWB = FileManager.SafeOpenWorkbook(TemplatePath)
    If TemplateWB Is Nothing Then
        CreateJobFile = False
        Exit Function
    End If

    PopulateJobTemplate TemplateWB, JobInfo

    TemplateWB.SaveAs FilePath
    FileManager.SafeCloseWorkbook TemplateWB

    CreateJobFile = True
    Exit Function

Error_Handler:
    If Not TemplateWB Is Nothing Then FileManager.SafeCloseWorkbook TemplateWB, False
    ErrorHandler.HandleStandardErrors Err.Number, "CreateJobFile", "JobController"
    CreateJobFile = False
End Function

Private Sub PopulateJobTemplate(ByVal wb As Workbook, ByVal JobInfo As JobData)
    Dim ws As Worksheet

    On Error GoTo Error_Handler

    Set ws = wb.Worksheets(1)

    With ws
        .Cells(2, 2).Value = JobInfo.JobNumber
        .Cells(3, 2).Value = JobInfo.CustomerName
        .Cells(8, 2).Value = JobInfo.ComponentDescription
        .Cells(9, 2).Value = JobInfo.ComponentCode
        .Cells(10, 2).Value = JobInfo.MaterialGrade
        .Cells(11, 2).Value = JobInfo.Quantity
        .Cells(12, 2).Value = JobInfo.DateCreated
        .Cells(13, 2).Value = JobInfo.DueDate
        .Cells(14, 2).Value = JobInfo.WorkshopDueDate
        .Cells(15, 2).Value = JobInfo.CustomerDueDate
        .Cells(16, 2).Value = JobInfo.OrderValue
        .Cells(17, 2).Value = JobInfo.Status
        .Cells(18, 2).Value = JobInfo.AssignedOperator
        .Cells(19, 2).Value = JobInfo.Operations
        .Cells(20, 2).Value = JobInfo.Notes
    End With

    Exit Sub

Error_Handler:
    ErrorHandler.LogError Err.Number, Err.Description, "PopulateJobTemplate", "JobController"
End Sub

Public Function LoadJob(ByVal FilePath As String) As JobData
    Dim JobWB As Workbook
    Dim ws As Worksheet
    Dim JobInfo As JobData

    On Error GoTo Error_Handler

    Set JobWB = FileManager.SafeOpenWorkbook(FilePath)
    If JobWB Is Nothing Then
        Exit Function
    End If

    Set ws = JobWB.Worksheets(1)

    With JobInfo
        .JobNumber = ws.Cells(2, 2).Value
        .QuoteNumber = ws.Cells(2, 2).Value
        .CustomerName = ws.Cells(3, 2).Value
        .ComponentDescription = ws.Cells(8, 2).Value
        .ComponentCode = ws.Cells(9, 2).Value
        .MaterialGrade = ws.Cells(10, 2).Value
        .Quantity = ws.Cells(11, 2).Value
        .DateCreated = ws.Cells(12, 2).Value
        .DueDate = ws.Cells(13, 2).Value
        .WorkshopDueDate = ws.Cells(14, 2).Value
        .CustomerDueDate = ws.Cells(15, 2).Value
        .OrderValue = ws.Cells(16, 2).Value
        .Status = ws.Cells(17, 2).Value
        .AssignedOperator = ws.Cells(18, 2).Value
        .Operations = ws.Cells(19, 2).Value
        .Notes = ws.Cells(20, 2).Value
        .FilePath = FilePath
    End With

    FileManager.SafeCloseWorkbook JobWB, False
    LoadJob = JobInfo
    Exit Function

Error_Handler:
    If Not JobWB Is Nothing Then FileManager.SafeCloseWorkbook JobWB, False
    ErrorHandler.HandleStandardErrors Err.Number, "LoadJob", "JobController"
End Function

Public Function UpdateJob(ByVal JobInfo As JobData) As Boolean
    Dim JobWB As Workbook

    On Error GoTo Error_Handler

    Set JobWB = FileManager.SafeOpenWorkbook(JobInfo.FilePath)
    If JobWB Is Nothing Then
        UpdateJob = False
        Exit Function
    End If

    PopulateJobTemplate JobWB, JobInfo

    JobWB.Save
    FileManager.SafeCloseWorkbook JobWB

    WIPManager.UpdateJobInWIP JobInfo

    UpdateJob = True
    Exit Function

Error_Handler:
    If Not JobWB Is Nothing Then FileManager.SafeCloseWorkbook JobWB, False
    ErrorHandler.HandleStandardErrors Err.Number, "UpdateJob", "JobController"
    UpdateJob = False
End Function

Public Function CloseJob(ByVal JobNumber As String) As Boolean
    Dim JobInfo As JobData
    Dim WIPPath As String
    Dim ArchivePath As String

    On Error GoTo Error_Handler

    WIPPath = FileManager.GetRootPath & "\WIP\" & JobNumber & ".xls"
    ArchivePath = FileManager.GetRootPath & "\Archive\" & JobNumber & ".xls"

    JobInfo = LoadJob(WIPPath)
    If JobInfo.JobNumber = "" Then
        CloseJob = False
        Exit Function
    End If

    JobInfo.Status = "Completed"
    UpdateJob JobInfo

    FileCopy WIPPath, ArchivePath
    Kill WIPPath

    WIPManager.RemoveJobFromWIP JobNumber

    CloseJob = True
    Exit Function

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "CloseJob", "JobController"
    CloseJob = False
End Function

Public Function ValidateJobData(ByVal JobInfo As JobData) As String
    Dim ValidationErrors As String

    If Trim(JobInfo.CustomerName) = "" Then
        ValidationErrors = ValidationErrors & "Customer name is required." & vbCrLf
    End If

    If JobInfo.Quantity <= 0 Then
        ValidationErrors = ValidationErrors & "Quantity must be greater than zero." & vbCrLf
    End If

    If JobInfo.DueDate < Date Then
        ValidationErrors = ValidationErrors & "Due date cannot be in the past." & vbCrLf
    End If

    ValidateJobData = ValidationErrors
End Function