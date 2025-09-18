Attribute VB_Name = "WIPManager"
Option Explicit

Private Const WIP_FILE As String = "WIP.xls"

Public Function AddJobToWIP(ByRef JobInfo As JobData) As Boolean
    Dim WIPWB As Workbook
    Dim WIPWS As Worksheet
    Dim LastRow As Long

    On Error GoTo Error_Handler

    Set WIPWB = FileManager.SafeOpenWorkbook(FileManager.GetRootPath & "\" & WIP_FILE)
    If WIPWB Is Nothing Then
        AddJobToWIP = False
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
        .Cells(LastRow, 6).Value = JobInfo.WorkshopDueDate
        .Cells(LastRow, 7).Value = JobInfo.CustomerDueDate
        .Cells(LastRow, 8).Value = JobInfo.OrderValue
        .Cells(LastRow, 9).Value = JobInfo.Status
        .Cells(LastRow, 10).Value = JobInfo.AssignedOperator
        .Cells(LastRow, 11).Value = JobInfo.DateCreated
        .Cells(LastRow, 12).Value = JobInfo.FilePath
    End With

    WIPWB.Save
    FileManager.SafeCloseWorkbook WIPWB

    AddJobToWIP = True
    Exit Function

Error_Handler:
    If Not WIPWB Is Nothing Then FileManager.SafeCloseWorkbook WIPWB, False
    ErrorHandler.HandleStandardErrors Err.Number, "AddJobToWIP", "WIPManager"
    AddJobToWIP = False
End Function

Public Function UpdateJobInWIP(ByRef JobInfo As JobData) As Boolean
    Dim WIPWB As Workbook
    Dim WIPWS As Worksheet
    Dim LastRow As Long
    Dim i As Long

    On Error GoTo Error_Handler

    Set WIPWB = FileManager.SafeOpenWorkbook(FileManager.GetRootPath & "\" & WIP_FILE)
    If WIPWB Is Nothing Then
        UpdateJobInWIP = False
        Exit Function
    End If

    Set WIPWS = WIPWB.Worksheets(1)
    LastRow = WIPWS.Cells(WIPWS.Rows.Count, 1).End(xlUp).Row

    For i = 2 To LastRow
        If WIPWS.Cells(i, 1).Value = JobInfo.JobNumber Then
            With WIPWS
                .Cells(i, 2).Value = JobInfo.CustomerName
                .Cells(i, 3).Value = JobInfo.ComponentDescription
                .Cells(i, 4).Value = JobInfo.Quantity
                .Cells(i, 5).Value = JobInfo.DueDate
                .Cells(i, 6).Value = JobInfo.WorkshopDueDate
                .Cells(i, 7).Value = JobInfo.CustomerDueDate
                .Cells(i, 8).Value = JobInfo.OrderValue
                .Cells(i, 9).Value = JobInfo.Status
                .Cells(i, 10).Value = JobInfo.AssignedOperator
                .Cells(i, 12).Value = JobInfo.FilePath
            End With

            WIPWB.Save
            FileManager.SafeCloseWorkbook WIPWB
            UpdateJobInWIP = True
            Exit Function
        End If
    Next i

    FileManager.SafeCloseWorkbook WIPWB, False
    UpdateJobInWIP = False
    Exit Function

Error_Handler:
    If Not WIPWB Is Nothing Then FileManager.SafeCloseWorkbook WIPWB, False
    ErrorHandler.HandleStandardErrors Err.Number, "UpdateJobInWIP", "WIPManager"
    UpdateJobInWIP = False
End Function

Public Function RemoveJobFromWIP(ByVal JobNumber As String) As Boolean
    Dim WIPWB As Workbook
    Dim WIPWS As Worksheet
    Dim LastRow As Long
    Dim i As Long

    On Error GoTo Error_Handler

    Set WIPWB = FileManager.SafeOpenWorkbook(FileManager.GetRootPath & "\" & WIP_FILE)
    If WIPWB Is Nothing Then
        RemoveJobFromWIP = False
        Exit Function
    End If

    Set WIPWS = WIPWB.Worksheets(1)
    LastRow = WIPWS.Cells(WIPWS.Rows.Count, 1).End(xlUp).Row

    For i = 2 To LastRow
        If WIPWS.Cells(i, 1).Value = JobNumber Then
            WIPWS.Rows(i).Delete
            WIPWB.Save
            FileManager.SafeCloseWorkbook WIPWB
            RemoveJobFromWIP = True
            Exit Function
        End If
    Next i

    FileManager.SafeCloseWorkbook WIPWB, False
    RemoveJobFromWIP = False
    Exit Function

Error_Handler:
    If Not WIPWB Is Nothing Then FileManager.SafeCloseWorkbook WIPWB, False
    ErrorHandler.HandleStandardErrors Err.Number, "RemoveJobFromWIP", "WIPManager"
    RemoveJobFromWIP = False
End Function

Public Function GetWIPJobs(Optional ByVal CustomerFilter As String = "", Optional ByVal OperatorFilter As String = "") As Variant
    Dim WIPWB As Workbook
    Dim WIPWS As Worksheet
    Dim LastRow As Long
    Dim i As Long
    Dim Jobs() As JobData
    Dim JobCount As Integer
    Dim CurrentJob As JobData

    On Error GoTo Error_Handler

    Set WIPWB = FileManager.SafeOpenWorkbook(FileManager.GetRootPath & "\" & WIP_FILE)
    If WIPWB Is Nothing Then
        GetWIPJobs = Array()
        Exit Function
    End If

    Set WIPWS = WIPWB.Worksheets(1)
    LastRow = WIPWS.Cells(WIPWS.Rows.Count, 1).End(xlUp).Row
    JobCount = 0

    For i = 2 To LastRow
        If (CustomerFilter = "" Or InStr(UCase(WIPWS.Cells(i, 2).Value), UCase(CustomerFilter)) > 0) And _
           (OperatorFilter = "" Or WIPWS.Cells(i, 10).Value = OperatorFilter) Then

            ReDim Preserve Jobs(JobCount)

            With CurrentJob
                .JobNumber = WIPWS.Cells(i, 1).Value
                .CustomerName = WIPWS.Cells(i, 2).Value
                .ComponentDescription = WIPWS.Cells(i, 3).Value
                .Quantity = WIPWS.Cells(i, 4).Value
                .DueDate = WIPWS.Cells(i, 5).Value
                .WorkshopDueDate = WIPWS.Cells(i, 6).Value
                .CustomerDueDate = WIPWS.Cells(i, 7).Value
                .OrderValue = WIPWS.Cells(i, 8).Value
                .Status = WIPWS.Cells(i, 9).Value
                .AssignedOperator = WIPWS.Cells(i, 10).Value
                .DateCreated = WIPWS.Cells(i, 11).Value
                .FilePath = WIPWS.Cells(i, 12).Value
            End With

            Jobs(JobCount) = CurrentJob
            JobCount = JobCount + 1
        End If
    Next i

    FileManager.SafeCloseWorkbook WIPWB, False

    If JobCount > 0 Then
        GetWIPJobs = Jobs
    Else
        GetWIPJobs = Array()
    End If

    Exit Function

Error_Handler:
    If Not WIPWB Is Nothing Then FileManager.SafeCloseWorkbook WIPWB, False
    ErrorHandler.HandleStandardErrors Err.Number, "GetWIPJobs", "WIPManager"
    GetWIPJobs = Array()
End Function

Public Function GenerateWIPReport(ByVal ReportType As String, Optional ByVal FilterValue As String = "") As Boolean
    Dim Jobs As Variant
    Dim ReportWB As Workbook
    Dim ReportWS As Worksheet
    Dim i As Integer
    Dim ReportPath As String

    On Error GoTo Error_Handler

    Jobs = GetWIPJobs()
    If UBound(Jobs) = -1 Then
        GenerateWIPReport = False
        Exit Function
    End If

    Set ReportWB = Workbooks.Add
    Set ReportWS = ReportWB.Worksheets(1)

    CreateReportHeaders ReportWS, ReportType

    For i = 0 To UBound(Jobs)
        If FilterValue = "" Or JobMatchesFilter(Jobs(i), ReportType, FilterValue) Then
            PopulateReportRow ReportWS, Jobs(i), i + 2
        End If
    Next i

    ReportPath = FileManager.GetRootPath & "\Templates\" & ReportType & "_Report_" & Format(Now, "yyyymmdd_hhmmss") & ".xls"
    ReportWB.SaveAs ReportPath
    ReportWB.Close

    GenerateWIPReport = True
    Exit Function

Error_Handler:
    If Not ReportWB Is Nothing Then ReportWB.Close SaveChanges:=False
    ErrorHandler.HandleStandardErrors Err.Number, "GenerateWIPReport", "WIPManager"
    GenerateWIPReport = False
End Function

Private Sub CreateReportHeaders(ByVal ws As Worksheet, ByVal ReportType As String)
    With ws
        .Cells(1, 1).Value = "Job Number"
        .Cells(1, 2).Value = "Customer"
        .Cells(1, 3).Value = "Description"
        .Cells(1, 4).Value = "Quantity"
        .Cells(1, 5).Value = "Due Date"
        .Cells(1, 6).Value = "Status"
        .Cells(1, 7).Value = "Operator"
        .Range("A1:G1").Font.Bold = True
    End With
End Sub

Private Sub PopulateReportRow(ByVal ws As Worksheet, ByRef Job As JobData, ByVal RowNumber As Long)
    With ws
        .Cells(RowNumber, 1).Value = Job.JobNumber
        .Cells(RowNumber, 2).Value = Job.CustomerName
        .Cells(RowNumber, 3).Value = Job.ComponentDescription
        .Cells(RowNumber, 4).Value = Job.Quantity
        .Cells(RowNumber, 5).Value = Job.DueDate
        .Cells(RowNumber, 6).Value = Job.Status
        .Cells(RowNumber, 7).Value = Job.AssignedOperator
    End With
End Sub

Private Function JobMatchesFilter(ByRef Job As JobData, ByVal ReportType As String, ByVal FilterValue As String) As Boolean
    Select Case UCase(ReportType)
        Case "CUSTOMER"
            JobMatchesFilter = (InStr(UCase(Job.CustomerName), UCase(FilterValue)) > 0)
        Case "OPERATOR"
            JobMatchesFilter = (UCase(Job.AssignedOperator) = UCase(FilterValue))
        Case "DUEDATE"
            JobMatchesFilter = (Job.DueDate <= CDate(FilterValue))
        Case Else
            JobMatchesFilter = True
    End Select
End Function