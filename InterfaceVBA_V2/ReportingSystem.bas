Attribute VB_Name = "ReportingSystem"
' **Purpose**: All reporting and output generation functionality
' **CLAUDE.md Compliance**: Maintains WIP reporting requirements, preserves file access patterns
' **Consolidation**: Combines WIPReportManager.bas and any future reporting modules
Option Explicit

' ===================================================================
' TYPES AND CONSTANTS
' ===================================================================

Private Type Jobs
    Dat As Date
    Cust As String
    Job As String
    JobD As Double
    Qty As String
    Cod As String
    Desc As String
    Remarks As String
    DDat As String

    OperatorN(1 To 15) As String
    OperatorType(1 To 15) As String
End Type

Private Const WIP_FILE As String = "WIP.xls"

' ===================================================================
' PUBLIC INTERFACE FUNCTIONS
' ===================================================================

' **Purpose**: Generate WIP reports based on form selections
' **Parameters**:
'   - ReportForm (Object): Form containing report selection options (ROperation, ROperator, etc.)
' **Returns**: Boolean - True if reports generated successfully, False if failed
' **Dependencies**: DataOperations.GetRootPath, DataOperations.SafeOpenWorkbook
' **Side Effects**: Creates report files in Templates directory, shows progress messages
' **Errors**: Returns False if WIP file not found or report generation fails
' **CLAUDE.md Compliance**: Exact replacement for fwip.frm business logic
Public Function GenerateWIPReports(ReportForm As Object) As Boolean
    Dim Job(1 To 5000) As Jobs
    Dim JobCount As Integer
    Dim WIPPath As String
    Dim WIPWB As Workbook

    On Error GoTo Error_Handler

    ' Validate report selection first
    If Not SystemCore.ValidateReportSelection(ReportForm) Then
        GenerateWIPReports = False
        Exit Function
    End If

    ' Update form status
    ReportForm.Label1.Caption = "Please Wait"
    Application.DisplayAlerts = False

    ' Load WIP data
    WIPPath = DataOperations.GetRootPath & "\" & WIP_FILE

    If Not DataOperations.FileExists(WIPPath) Then
        ReportForm.Label1.Caption = "Ready"
        SystemCore.ShowInformation "WIP.xls file not found at: " & WIPPath & vbCrLf & vbCrLf & _
               "The WIP.xls file is created when job cards save their data." & vbCrLf & _
               "Please ensure you have some active jobs that have been saved.", "WIP.xls Not Found"
        GenerateWIPReports = False
        Exit Function
    End If

    ' Open and load WIP data
    Set WIPWB = DataOperations.SafeOpenWorkbook(WIPPath)
    If WIPWB Is Nothing Then
        ReportForm.Label1.Caption = "Ready"
        SystemCore.ShowError "Unable to open WIP.xls at: " & WIPPath, "File Access Error"
        GenerateWIPReports = False
        Exit Function
    End If

    JobCount = LoadWIPDataFromWorkbook(WIPWB, Job)
    DataOperations.SafeCloseWorkbook WIPWB, False

    If JobCount = 0 Then
        ReportForm.Label1.Caption = "Ready"
        SystemCore.ShowInformation "No WIP data found in WIP.xls file." & vbCrLf & vbCrLf & _
               "Please ensure there are active jobs saved in the system.", "No WIP Data"
        GenerateWIPReports = False
        Exit Function
    End If

    ' Hide form during processing
    ReportForm.Hide

    ' Generate requested reports
    If ReportForm.ROperation.Value = True Then
        GenerateOperationReports Job, JobCount
    End If

    If ReportForm.ROperator.Value = True Then
        GenerateOperatorReports Job, JobCount
    End If

    Application.DisplayAlerts = True

    ' Show completion and restore form
    ReportForm.Show
    ReportForm.Label1.Caption = "Complete"
    SystemCore.ShowInformation "WIP reports have been generated successfully!" & vbCrLf & _
           "Reports saved to Templates directory:" & vbCrLf & _
           "- Operation reports (if selected)" & vbCrLf & _
           "- Operator reports (if selected)" & vbCrLf & _
           "Check your Templates folder for the generated files.", "Reports Generated"

    ReportForm.Label1.Caption = "Ready - Select report types and click Go"
    GenerateWIPReports = True
    Exit Function

Error_Handler:
    Application.DisplayAlerts = True
    If Not WIPWB Is Nothing Then DataOperations.SafeCloseWorkbook WIPWB, False
    If Not ReportForm Is Nothing Then
        ReportForm.Show
        ReportForm.Label1.Caption = "Ready"
    End If
    SystemCore.LogError Err.Number, Err.Description, "GenerateWIPReports", "ReportingSystem"
    GenerateWIPReports = False
End Function

' **Purpose**: Export WIP data to Excel file for external analysis
' **Parameters**:
'   - ExportPath (String, Optional): Path for export file, default generates timestamped name
' **Returns**: Boolean - True if export successful, False if failed
' **Dependencies**: DataOperations.GetRootPath, DataOperations.SafeOpenWorkbook
' **Side Effects**: Creates new Excel file with WIP data export
' **Errors**: Returns False if WIP file not accessible or export fails
Public Function ExportWIPData(Optional ExportPath As String = "") As Boolean
    Dim WIPPath As String
    Dim WIPWB As Workbook
    Dim ExportWB As Workbook
    Dim Job(1 To 5000) As Jobs
    Dim JobCount As Integer

    On Error GoTo Error_Handler

    ' Set default export path if not provided
    If ExportPath = "" Then
        ExportPath = DataOperations.GetRootPath & "\WIP_Export_" & Format(Now, "yyyymmdd_hhmmss") & ".xls"
    End If

    ' Load WIP data
    WIPPath = DataOperations.GetRootPath & "\" & WIP_FILE

    If Not DataOperations.FileExists(WIPPath) Then
        SystemCore.ShowError "WIP.xls file not found at: " & WIPPath, "File Not Found"
        ExportWIPData = False
        Exit Function
    End If

    Set WIPWB = DataOperations.SafeOpenWorkbook(WIPPath)
    If WIPWB Is Nothing Then
        ExportWIPData = False
        Exit Function
    End If

    JobCount = LoadWIPDataFromWorkbook(WIPWB, Job)
    DataOperations.SafeCloseWorkbook WIPWB, False

    If JobCount = 0 Then
        SystemCore.ShowWarning "No WIP data to export.", "No Data"
        ExportWIPData = False
        Exit Function
    End If

    ' Create export workbook
    Set ExportWB = DataOperations.CreateNewWorkbook()
    If ExportWB Is Nothing Then
        ExportWIPData = False
        Exit Function
    End If

    ' Export data to new workbook
    If CreateWIPExport(ExportWB, Job, JobCount) Then
        ExportWB.SaveAs ExportPath
        ExportWB.Close
        SystemCore.ShowInformation "WIP data exported successfully to:" & vbCrLf & ExportPath, "Export Complete"
        ExportWIPData = True
    Else
        ExportWB.Close SaveChanges:=False
        ExportWIPData = False
    End If

    Set ExportWB = Nothing
    Exit Function

Error_Handler:
    If Not WIPWB Is Nothing Then DataOperations.SafeCloseWorkbook WIPWB, False
    If Not ExportWB Is Nothing Then ExportWB.Close SaveChanges:=False
    SystemCore.HandleStandardErrors Err.Number, "ExportWIPData", "ReportingSystem"
    ExportWIPData = False
End Function

' **Purpose**: Generate summary statistics for WIP data
' **Parameters**: None
' **Returns**: Variant - Array containing summary statistics, empty array if failed
' **Dependencies**: DataOperations.GetRootPath, DataOperations.SafeOpenWorkbook
' **Side Effects**: None (read-only operation)
' **Errors**: Returns empty array if WIP file not accessible
Public Function GetWIPSummaryStatistics() As Variant
    Dim WIPPath As String
    Dim WIPWB As Workbook
    Dim Job(1 To 5000) As Jobs
    Dim JobCount As Integer
    Dim Summary(0 To 6) As Variant

    On Error GoTo Error_Handler

    WIPPath = DataOperations.GetRootPath & "\" & WIP_FILE

    If Not DataOperations.FileExists(WIPPath) Then
        GetWIPSummaryStatistics = Array()
        Exit Function
    End If

    Set WIPWB = DataOperations.SafeOpenWorkbook(WIPPath)
    If WIPWB Is Nothing Then
        GetWIPSummaryStatistics = Array()
        Exit Function
    End If

    JobCount = LoadWIPDataFromWorkbook(WIPWB, Job)
    DataOperations.SafeCloseWorkbook WIPWB, False

    If JobCount = 0 Then
        GetWIPSummaryStatistics = Array()
        Exit Function
    End If

    ' Calculate summary statistics
    Summary(0) = JobCount ' Total Jobs
    Summary(1) = CountUniqueCustomers(Job, JobCount) ' Unique Customers
    Summary(2) = GetOldestJobDate(Job, JobCount) ' Oldest Job Date
    Summary(3) = GetNewestJobDate(Job, JobCount) ' Newest Job Date
    Summary(4) = CountActiveOperators(Job, JobCount) ' Active Operators
    Summary(5) = GetAverageJobAge(Job, JobCount) ' Average Job Age (days)
    Summary(6) = CountOverdueJobs(Job, JobCount) ' Overdue Jobs

    GetWIPSummaryStatistics = Summary
    Exit Function

Error_Handler:
    If Not WIPWB Is Nothing Then DataOperations.SafeCloseWorkbook WIPWB, False
    SystemCore.HandleStandardErrors Err.Number, "GetWIPSummaryStatistics", "ReportingSystem"
    GetWIPSummaryStatistics = Array()
End Function

' ===================================================================
' PRIVATE HELPER FUNCTIONS
' ===================================================================

' **Purpose**: Load WIP data from consolidated workbook into Jobs array
' **Parameters**:
'   - WIPWB (Workbook): Opened WIP workbook
'   - Job (Jobs array): Array to populate with job data
' **Returns**: Integer - Number of jobs loaded
' **Dependencies**: None
' **Side Effects**: Modifies Job array with loaded data
' **Errors**: Returns 0 if data loading fails
Private Function LoadWIPDataFromWorkbook(WIPWB As Workbook, ByRef Job() As Jobs) As Integer
    Dim i As Integer
    Dim col As Integer
    Dim j As Integer
    Dim x As Integer

    On Error GoTo Error_Handler

    WIPWB.Activate

    ' Find the rightmost column with data (original code used BB1)
    Range("A1").Select
    Selection.End(xlToRight).Select
    col = ActiveCell.Column

    ' Sort the data by date
    Range("A1").Select
    Selection.End(xlDown).Select

    If ActiveCell.Row > 1 Then
        Range("A2", Range("A2").Offset(ActiveCell.Row - 2, col - 1).Address).Select
        Range(Selection, Selection.End(xlDown)).Select
        On Error Resume Next ' In case sorting fails
        Selection.Sort Key1:=Range("A3"), Order1:=xlAscending, Header:=xlYes, _
            OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
        On Error GoTo Error_Handler
    End If

    ' Load data into Jobs array
    Range("A3").Select
    i = 0
    If ActiveCell.Value <> "" Then
        Do While ActiveCell.Value <> "" And i < 5000
            i = i + 1
            With Job(i)
                .Dat = ActiveCell.Offset(0, 0).Value
                .Cust = ActiveCell.Offset(0, 1).Value
                .Job = ActiveCell.Offset(0, 2).Value
                .JobD = ParseJobNumberForSorting(CStr(ActiveCell.Offset(0, 3).Value))
                .Qty = CStr(ActiveCell.Offset(0, 4).Value)
                .Cod = CStr(ActiveCell.Offset(0, 5).Value)
                .Desc = CStr(ActiveCell.Offset(0, 6).Value)
                .Remarks = CStr(ActiveCell.Offset(0, 8).Value)
                .DDat = CStr(ActiveCell.Offset(0, 12).Value)

                ' Load operation data if available
                x = 0
                For j = 1 To 30 Step 2
                    x = x + 1
                    If x <= 15 Then
                        If (14 + j) <= col Then .OperatorType(x) = CStr(ActiveCell.Offset(0, 14 + j).Value)
                        If (15 + j) <= col Then .OperatorN(x) = CStr(ActiveCell.Offset(0, 15 + j).Value)
                    End If
                Next j
            End With
            ActiveCell.Offset(1, 0).Select
        Loop
    End If

    LoadWIPDataFromWorkbook = i
    Exit Function

Error_Handler:
    SystemCore.LogError Err.Number, Err.Description, "LoadWIPDataFromWorkbook", "ReportingSystem"
    LoadWIPDataFromWorkbook = 0
End Function

' **Purpose**: Generate operation-based reports from WIP data
' **Parameters**:
'   - Job (Jobs array): Array of job data
'   - JobCount (Integer): Number of jobs in array
' **Returns**: None (Subroutine)
' **Dependencies**: DataOperations.GetRootPath, DataOperations.CreateNewWorkbook
' **Side Effects**: Creates operation report files in Templates directory
' **Errors**: Logs errors if report generation fails
Private Sub GenerateOperationReports(ByRef Job() As Jobs, ByVal JobCount As Integer)
    Dim OperationTypes(1 To 50) As String
    Dim OperationCount As Integer
    Dim ReportWB As Workbook
    Dim ReportWS As Worksheet
    Dim i As Integer, j As Integer, k As Integer
    Dim CurrentRow As Integer

    On Error GoTo Error_Handler

    ' Extract unique operation types
    OperationCount = 0
    For i = 1 To JobCount
        For j = 1 To 15
            If Job(i).OperatorType(j) <> "" Then
                ' Check if operation type already exists
                Dim Found As Boolean
                Found = False
                For k = 1 To OperationCount
                    If OperationTypes(k) = Job(i).OperatorType(j) Then
                        Found = True
                        Exit For
                    End If
                Next k

                If Not Found Then
                    OperationCount = OperationCount + 1
                    OperationTypes(OperationCount) = Job(i).OperatorType(j)
                End If
            End If
        Next j
    Next i

    ' Generate report for each operation type
    For k = 1 To OperationCount
        Set ReportWB = DataOperations.CreateNewWorkbook()
        If Not ReportWB Is Nothing Then
            Set ReportWS = ReportWB.Worksheets(1)
            ReportWS.Name = Left("Op_" & OperationTypes(k), 31) ' Excel sheet name limit

            ' Create headers
            With ReportWS
                .Cells(1, 1).Value = "Operation Report: " & OperationTypes(k)
                .Cells(1, 1).Font.Bold = True
                .Cells(2, 1).Value = "Generated: " & Format(Now, "dd/mm/yyyy hh:mm")

                .Cells(4, 1).Value = "Date"
                .Cells(4, 2).Value = "Customer"
                .Cells(4, 3).Value = "Job"
                .Cells(4, 4).Value = "Qty"
                .Cells(4, 5).Value = "Code"
                .Cells(4, 6).Value = "Description"
                .Cells(4, 7).Value = "Due Date"
                .Cells(4, 8).Value = "Operator"
                .Range("A4:H4").Font.Bold = True
            End With

            CurrentRow = 5

            ' Add jobs for this operation type
            For i = 1 To JobCount
                For j = 1 To 15
                    If Job(i).OperatorType(j) = OperationTypes(k) Then
                        With ReportWS
                            .Cells(CurrentRow, 1).Value = Job(i).Dat
                            .Cells(CurrentRow, 2).Value = Job(i).Cust
                            .Cells(CurrentRow, 3).Value = Job(i).Job
                            .Cells(CurrentRow, 4).Value = Job(i).Qty
                            .Cells(CurrentRow, 5).Value = Job(i).Cod
                            .Cells(CurrentRow, 6).Value = Job(i).Desc
                            .Cells(CurrentRow, 7).Value = Job(i).DDat
                            .Cells(CurrentRow, 8).Value = Job(i).OperatorN(j)
                        End With
                        CurrentRow = CurrentRow + 1
                        Exit For ' Only add job once per operation type
                    End If
                Next j
            Next i

            ' Auto-fit columns and save
            ReportWS.Columns.AutoFit
            Dim SavePath As String
            SavePath = DataOperations.GetRootPath & "\Templates\WIP_Operation_" & SystemCore.CleanFileName(OperationTypes(k)) & "_" & Format(Now, "yyyymmdd") & ".xls"
            ReportWB.SaveAs SavePath
            ReportWB.Close
            Set ReportWB = Nothing
        End If
    Next k

    Exit Sub

Error_Handler:
    If Not ReportWB Is Nothing Then
        ReportWB.Close SaveChanges:=False
        Set ReportWB = Nothing
    End If
    SystemCore.LogError Err.Number, Err.Description, "GenerateOperationReports", "ReportingSystem"
End Sub

' **Purpose**: Generate operator-based reports from WIP data
' **Parameters**:
'   - Job (Jobs array): Array of job data
'   - JobCount (Integer): Number of jobs in array
' **Returns**: None (Subroutine)
' **Dependencies**: DataOperations.GetRootPath, DataOperations.CreateNewWorkbook
' **Side Effects**: Creates operator report files in Templates directory
' **Errors**: Logs errors if report generation fails
Private Sub GenerateOperatorReports(ByRef Job() As Jobs, ByVal JobCount As Integer)
    Dim Operators(1 To 50) As String
    Dim OperatorCount As Integer
    Dim ReportWB As Workbook
    Dim ReportWS As Worksheet
    Dim i As Integer, j As Integer, k As Integer
    Dim CurrentRow As Integer

    On Error GoTo Error_Handler

    ' Extract unique operators
    OperatorCount = 0
    For i = 1 To JobCount
        For j = 1 To 15
            If Job(i).OperatorN(j) <> "" Then
                ' Check if operator already exists
                Dim Found As Boolean
                Found = False
                For k = 1 To OperatorCount
                    If Operators(k) = Job(i).OperatorN(j) Then
                        Found = True
                        Exit For
                    End If
                Next k

                If Not Found Then
                    OperatorCount = OperatorCount + 1
                    Operators(OperatorCount) = Job(i).OperatorN(j)
                End If
            End If
        Next j
    Next i

    ' Generate report for each operator
    For k = 1 To OperatorCount
        Set ReportWB = DataOperations.CreateNewWorkbook()
        If Not ReportWB Is Nothing Then
            Set ReportWS = ReportWB.Worksheets(1)
            ReportWS.Name = Left("Op_" & Operators(k), 31) ' Excel sheet name limit

            ' Create headers
            With ReportWS
                .Cells(1, 1).Value = "Operator Report: " & Operators(k)
                .Cells(1, 1).Font.Bold = True
                .Cells(2, 1).Value = "Generated: " & Format(Now, "dd/mm/yyyy hh:mm")

                .Cells(4, 1).Value = "Date"
                .Cells(4, 2).Value = "Customer"
                .Cells(4, 3).Value = "Job"
                .Cells(4, 4).Value = "Qty"
                .Cells(4, 5).Value = "Code"
                .Cells(4, 6).Value = "Description"
                .Cells(4, 7).Value = "Due Date"
                .Cells(4, 8).Value = "Operation"
                .Range("A4:H4").Font.Bold = True
            End With

            CurrentRow = 5

            ' Add jobs for this operator
            For i = 1 To JobCount
                For j = 1 To 15
                    If Job(i).OperatorN(j) = Operators(k) Then
                        With ReportWS
                            .Cells(CurrentRow, 1).Value = Job(i).Dat
                            .Cells(CurrentRow, 2).Value = Job(i).Cust
                            .Cells(CurrentRow, 3).Value = Job(i).Job
                            .Cells(CurrentRow, 4).Value = Job(i).Qty
                            .Cells(CurrentRow, 5).Value = Job(i).Cod
                            .Cells(CurrentRow, 6).Value = Job(i).Desc
                            .Cells(CurrentRow, 7).Value = Job(i).DDat
                            .Cells(CurrentRow, 8).Value = Job(i).OperatorType(j)
                        End With
                        CurrentRow = CurrentRow + 1
                    End If
                Next j
            Next i

            ' Auto-fit columns and save
            ReportWS.Columns.AutoFit
            Dim SavePath As String
            SavePath = DataOperations.GetRootPath & "\Templates\WIP_Operator_" & SystemCore.CleanFileName(Operators(k)) & "_" & Format(Now, "yyyymmdd") & ".xls"
            ReportWB.SaveAs SavePath
            ReportWB.Close
            Set ReportWB = Nothing
        End If
    Next k

    Exit Sub

Error_Handler:
    If Not ReportWB Is Nothing Then
        ReportWB.Close SaveChanges:=False
        Set ReportWB = Nothing
    End If
    SystemCore.LogError Err.Number, Err.Description, "GenerateOperatorReports", "ReportingSystem"
End Sub

' **Purpose**: Parse job number for sorting purposes
' **Parameters**:
'   - JobNumber (String): Job number to parse
' **Returns**: Double - Numeric value for sorting
' **Dependencies**: None
' **Side Effects**: None
' **Errors**: Returns 0 if parsing fails
Private Function ParseJobNumberForSorting(ByVal JobNumber As String) As Double
    Dim NumericPart As String

    On Error GoTo Error_Handler

    ' Remove prefix and extract numeric part
    If Len(JobNumber) > 1 Then
        NumericPart = Mid(JobNumber, 2) ' Remove first character (J, E, Q, etc.)
        ' Remove any non-numeric characters except decimal point
        Dim i As Integer
        Dim CleanNumber As String
        For i = 1 To Len(NumericPart)
            Dim Char As String
            Char = Mid(NumericPart, i, 1)
            If IsNumeric(Char) Or Char = "." Then
                CleanNumber = CleanNumber & Char
            End If
        Next i

        If IsNumeric(CleanNumber) Then
            ParseJobNumberForSorting = CDbl(CleanNumber)
        Else
            ParseJobNumberForSorting = 0
        End If
    Else
        ParseJobNumberForSorting = 0
    End If
    Exit Function

Error_Handler:
    ParseJobNumberForSorting = 0
End Function

' **Purpose**: Create WIP data export in workbook
' **Parameters**:
'   - ExportWB (Workbook): Workbook to create export in
'   - Job (Jobs array): Array of job data
'   - JobCount (Integer): Number of jobs in array
' **Returns**: Boolean - True if export created successfully, False if failed
' **Dependencies**: None
' **Side Effects**: Populates export workbook with formatted data
' **Errors**: Returns False if export creation fails
Private Function CreateWIPExport(ExportWB As Workbook, ByRef Job() As Jobs, JobCount As Integer) As Boolean
    Dim ExportWS As Worksheet
    Dim i As Integer, j As Integer
    Dim CurrentRow As Integer

    On Error GoTo Error_Handler

    Set ExportWS = ExportWB.Worksheets(1)
    ExportWS.Name = "WIP_Export_" & Format(Now, "yyyymmdd")

    ' Create headers
    With ExportWS
        .Cells(1, 1).Value = "WIP Data Export"
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Font.Size = 14
        .Cells(2, 1).Value = "Generated: " & Format(Now, "dd/mm/yyyy hh:mm")
        .Cells(3, 1).Value = "Total Jobs: " & JobCount

        .Cells(5, 1).Value = "Date"
        .Cells(5, 2).Value = "Customer"
        .Cells(5, 3).Value = "Job"
        .Cells(5, 4).Value = "Qty"
        .Cells(5, 5).Value = "Code"
        .Cells(5, 6).Value = "Description"
        .Cells(5, 7).Value = "Remarks"
        .Cells(5, 8).Value = "Due Date"

        ' Add operation headers
        For j = 1 To 15
            .Cells(5, 8 + (j * 2 - 1)).Value = "Op" & j & "_Type"
            .Cells(5, 8 + (j * 2)).Value = "Op" & j & "_Name"
        Next j

        .Range("A5:AX5").Font.Bold = True
    End With

    CurrentRow = 6

    ' Add all job data
    For i = 1 To JobCount
        With ExportWS
            .Cells(CurrentRow, 1).Value = Job(i).Dat
            .Cells(CurrentRow, 2).Value = Job(i).Cust
            .Cells(CurrentRow, 3).Value = Job(i).Job
            .Cells(CurrentRow, 4).Value = Job(i).Qty
            .Cells(CurrentRow, 5).Value = Job(i).Cod
            .Cells(CurrentRow, 6).Value = Job(i).Desc
            .Cells(CurrentRow, 7).Value = Job(i).Remarks
            .Cells(CurrentRow, 8).Value = Job(i).DDat

            ' Add operation data
            For j = 1 To 15
                .Cells(CurrentRow, 8 + (j * 2 - 1)).Value = Job(i).OperatorType(j)
                .Cells(CurrentRow, 8 + (j * 2)).Value = Job(i).OperatorN(j)
            Next j
        End With
        CurrentRow = CurrentRow + 1
    Next i

    ' Auto-fit columns
    ExportWS.Columns.AutoFit

    CreateWIPExport = True
    Exit Function

Error_Handler:
    SystemCore.LogError Err.Number, Err.Description, "CreateWIPExport", "ReportingSystem"
    CreateWIPExport = False
End Function

' ===================================================================
' SUMMARY STATISTICS HELPER FUNCTIONS
' ===================================================================

' **Purpose**: Count unique customers in WIP data
' **Parameters**:
'   - Job (Jobs array): Array of job data
'   - JobCount (Integer): Number of jobs in array
' **Returns**: Integer - Number of unique customers
Private Function CountUniqueCustomers(ByRef Job() As Jobs, JobCount As Integer) As Integer
    Dim Customers(1 To 1000) As String
    Dim CustomerCount As Integer
    Dim i As Integer, j As Integer
    Dim Found As Boolean

    CustomerCount = 0
    For i = 1 To JobCount
        Found = False
        For j = 1 To CustomerCount
            If Customers(j) = Job(i).Cust Then
                Found = True
                Exit For
            End If
        Next j

        If Not Found And Job(i).Cust <> "" Then
            CustomerCount = CustomerCount + 1
            Customers(CustomerCount) = Job(i).Cust
        End If
    Next i

    CountUniqueCustomers = CustomerCount
End Function

' **Purpose**: Get oldest job date in WIP data
' **Parameters**:
'   - Job (Jobs array): Array of job data
'   - JobCount (Integer): Number of jobs in array
' **Returns**: Date - Oldest job date
Private Function GetOldestJobDate(ByRef Job() As Jobs, JobCount As Integer) As Date
    Dim OldestDate As Date
    Dim i As Integer

    If JobCount > 0 Then
        OldestDate = Job(1).Dat
        For i = 2 To JobCount
            If Job(i).Dat < OldestDate Then
                OldestDate = Job(i).Dat
            End If
        Next i
    End If

    GetOldestJobDate = OldestDate
End Function

' **Purpose**: Get newest job date in WIP data
' **Parameters**:
'   - Job (Jobs array): Array of job data
'   - JobCount (Integer): Number of jobs in array
' **Returns**: Date - Newest job date
Private Function GetNewestJobDate(ByRef Job() As Jobs, JobCount As Integer) As Date
    Dim NewestDate As Date
    Dim i As Integer

    If JobCount > 0 Then
        NewestDate = Job(1).Dat
        For i = 2 To JobCount
            If Job(i).Dat > NewestDate Then
                NewestDate = Job(i).Dat
            End If
        Next i
    End If

    GetNewestJobDate = NewestDate
End Function

' **Purpose**: Count active operators in WIP data
' **Parameters**:
'   - Job (Jobs array): Array of job data
'   - JobCount (Integer): Number of jobs in array
' **Returns**: Integer - Number of active operators
Private Function CountActiveOperators(ByRef Job() As Jobs, JobCount As Integer) As Integer
    Dim Operators(1 To 100) As String
    Dim OperatorCount As Integer
    Dim i As Integer, j As Integer, k As Integer
    Dim Found As Boolean

    OperatorCount = 0
    For i = 1 To JobCount
        For j = 1 To 15
            If Job(i).OperatorN(j) <> "" Then
                Found = False
                For k = 1 To OperatorCount
                    If Operators(k) = Job(i).OperatorN(j) Then
                        Found = True
                        Exit For
                    End If
                Next k

                If Not Found Then
                    OperatorCount = OperatorCount + 1
                    Operators(OperatorCount) = Job(i).OperatorN(j)
                End If
            End If
        Next j
    Next i

    CountActiveOperators = OperatorCount
End Function

' **Purpose**: Get average job age in days
' **Parameters**:
'   - Job (Jobs array): Array of job data
'   - JobCount (Integer): Number of jobs in array
' **Returns**: Double - Average job age in days
Private Function GetAverageJobAge(ByRef Job() As Jobs, JobCount As Integer) As Double
    Dim TotalAge As Double
    Dim i As Integer

    TotalAge = 0
    For i = 1 To JobCount
        TotalAge = TotalAge + (Now - Job(i).Dat)
    Next i

    If JobCount > 0 Then
        GetAverageJobAge = TotalAge / JobCount
    Else
        GetAverageJobAge = 0
    End If
End Function

' **Purpose**: Count overdue jobs in WIP data
' **Parameters**:
'   - Job (Jobs array): Array of job data
'   - JobCount (Integer): Number of jobs in array
' **Returns**: Integer - Number of overdue jobs
Private Function CountOverdueJobs(ByRef Job() As Jobs, JobCount As Integer) As Integer
    Dim OverdueCount As Integer
    Dim i As Integer
    Dim DueDate As Date

    OverdueCount = 0
    For i = 1 To JobCount
        If IsDate(Job(i).DDat) Then
            DueDate = CDate(Job(i).DDat)
            If DueDate < Now Then
                OverdueCount = OverdueCount + 1
            End If
        End If
    Next i

    CountOverdueJobs = OverdueCount
End Function