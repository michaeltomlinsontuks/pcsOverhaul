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
' DATE FORMAT CONSTANTS (Standardized formatting across all WIP reports)
' ===================================================================
Private Const DATE_FORMAT_DISPLAY As String = "dd/mm/yyyy"
Private Const DATE_FORMAT_DISPLAY_TIME As String = "dd/mm/yyyy hh:mm"
Private Const DATE_FORMAT_EXCEL_COLUMN As String = "dd/mm/yyyy"
Private Const DATE_FORMAT_FILE_TIMESTAMP As String = "yyyymmdd_hhmmss"
Private Const DATE_FORMAT_FILE_DATE As String = "yyyymmdd"

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

    ' Generate additional WIP report types (exact legacy functionality)
    GenerateAdditionalWIPReports ReportForm

    Application.DisplayAlerts = True

    ' Close form and Main interface like original system (fwip.frm lines 531-532)
    ' This leaves the last generated report open for viewing
    ReportForm.Hide
    Application.StatusBar = "WIP reports generated successfully - check Templates folder"
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
        ExportPath = DataOperations.GetRootPath & "\WIP_Export_" & Format(Now, DATE_FORMAT_FILE_TIMESTAMP) & ".xls"
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
                ' Handle DDat (due date) - try to format as date if possible, otherwise keep as string
                Dim DueDateValue As Variant
                DueDateValue = ActiveCell.Offset(0, 12).Value
                If IsDate(DueDateValue) Then
                    .DDat = Format(CDate(DueDateValue), DATE_FORMAT_DISPLAY)
                Else
                    .DDat = CStr(DueDateValue)
                End If

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
                .Cells(2, 1).Value = "Generated: " & Format(Now, DATE_FORMAT_DISPLAY_TIME)

                .Cells(4, 1).Value = "Job Number"
                .Cells(4, 2).Value = "Customer"
                .Cells(4, 3).Value = "Description"
                .Cells(4, 4).Value = "Start Date"
                .Cells(4, 5).Value = "Due Date"
                .Cells(4, 6).Value = "Qty"
                .Cells(4, 7).Value = "Code"
                .Cells(4, 8).Value = "Operator"
                .Range("A4:H4").Font.Bold = True
            End With

            CurrentRow = 5

            ' Add jobs for this operation type
            For i = 1 To JobCount
                For j = 1 To 15
                    If Job(i).OperatorType(j) = OperationTypes(k) Then
                        With ReportWS
                            .Cells(CurrentRow, 1).Value = Job(i).Job
                            .Cells(CurrentRow, 2).Value = Job(i).Cust
                            .Cells(CurrentRow, 3).Value = Job(i).Desc
                            .Cells(CurrentRow, 4).Value = Job(i).Dat
                            .Cells(CurrentRow, 5).Value = Job(i).DDat
                            .Cells(CurrentRow, 6).Value = Job(i).Qty
                            .Cells(CurrentRow, 7).Value = Job(i).Cod
                            .Cells(CurrentRow, 8).Value = Job(i).OperatorN(j)
                        End With
                        CurrentRow = CurrentRow + 1
                        Exit For ' Only add job once per operation type
                    End If
                Next j
            Next i

            ' Apply date formatting to date columns (Column D: Start Date, Column E: Due Date)
            ReportWS.Columns("D:D").NumberFormat = DATE_FORMAT_EXCEL_COLUMN
            ReportWS.Columns("E:E").NumberFormat = DATE_FORMAT_EXCEL_COLUMN

            ' Auto-fit columns and save
            ReportWS.Columns.AutoFit
            Dim SavePath As String
            SavePath = DataOperations.GetRootPath & "\Templates\Operation.xls"
            Application.DisplayAlerts = False
            ReportWB.SaveAs SavePath
            Application.DisplayAlerts = True
            ' Keep last operation report open for viewing (original behavior)
            If k = OperationCount Then
                ' Last report - leave it open
            Else
                ReportWB.Close
            End If
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
                .Cells(2, 1).Value = "Generated: " & Format(Now, DATE_FORMAT_DISPLAY_TIME)

                .Cells(4, 1).Value = "Job Number"
                .Cells(4, 2).Value = "Customer"
                .Cells(4, 3).Value = "Description"
                .Cells(4, 4).Value = "Start Date"
                .Cells(4, 5).Value = "Due Date"
                .Cells(4, 6).Value = "Qty"
                .Cells(4, 7).Value = "Code"
                .Cells(4, 8).Value = "Operation"
                .Range("A4:H4").Font.Bold = True
            End With

            CurrentRow = 5

            ' Add jobs for this operator
            For i = 1 To JobCount
                For j = 1 To 15
                    If Job(i).OperatorN(j) = Operators(k) Then
                        With ReportWS
                            .Cells(CurrentRow, 1).Value = Job(i).Job
                            .Cells(CurrentRow, 2).Value = Job(i).Cust
                            .Cells(CurrentRow, 3).Value = Job(i).Desc
                            .Cells(CurrentRow, 4).Value = Job(i).Dat
                            .Cells(CurrentRow, 5).Value = Job(i).DDat
                            .Cells(CurrentRow, 6).Value = Job(i).Qty
                            .Cells(CurrentRow, 7).Value = Job(i).Cod
                            .Cells(CurrentRow, 8).Value = Job(i).OperatorType(j)
                        End With
                        CurrentRow = CurrentRow + 1
                    End If
                Next j
            Next i

            ' Apply date formatting to date columns (Column D: Start Date, Column E: Due Date)
            ReportWS.Columns("D:D").NumberFormat = DATE_FORMAT_EXCEL_COLUMN
            ReportWS.Columns("E:E").NumberFormat = DATE_FORMAT_EXCEL_COLUMN

            ' Auto-fit columns and save
            ReportWS.Columns.AutoFit
            Dim SavePath As String
            SavePath = DataOperations.GetRootPath & "\Templates\Operator.xls"
            Application.DisplayAlerts = False
            ReportWB.SaveAs SavePath
            Application.DisplayAlerts = True
            ' Keep last operator report open for viewing (original behavior)
            If k = OperatorCount Then
                ' Last report - leave it open
            Else
                ReportWB.Close
            End If
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
    ExportWS.Name = "WIP_Export_" & Format(Now, DATE_FORMAT_FILE_DATE)

    ' Create headers
    With ExportWS
        .Cells(1, 1).Value = "WIP Data Export"
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Font.Size = 14
        .Cells(2, 1).Value = "Generated: " & Format(Now, DATE_FORMAT_DISPLAY_TIME)
        .Cells(3, 1).Value = "Total Jobs: " & JobCount

        .Cells(5, 1).Value = "Job Number"
        .Cells(5, 2).Value = "Customer"
        .Cells(5, 3).Value = "Description"
        .Cells(5, 4).Value = "Remarks"
        .Cells(5, 5).Value = "Start Date"
        .Cells(5, 6).Value = "Due Date"
        .Cells(5, 7).Value = "Qty"
        .Cells(5, 8).Value = "Code"

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
            .Cells(CurrentRow, 1).Value = Job(i).Job
            .Cells(CurrentRow, 2).Value = Job(i).Cust
            .Cells(CurrentRow, 3).Value = Job(i).Desc
            .Cells(CurrentRow, 4).Value = Job(i).Remarks
            .Cells(CurrentRow, 5).Value = Job(i).Dat
            .Cells(CurrentRow, 6).Value = Job(i).DDat
            .Cells(CurrentRow, 7).Value = Job(i).Qty
            .Cells(CurrentRow, 8).Value = Job(i).Cod

            ' Add operation data
            For j = 1 To 15
                .Cells(CurrentRow, 8 + (j * 2 - 1)).Value = Job(i).OperatorType(j)
                .Cells(CurrentRow, 8 + (j * 2)).Value = Job(i).OperatorN(j)
            Next j
        End With
        CurrentRow = CurrentRow + 1
    Next i

    ' Apply date formatting to date columns (Column E: Start Date, Column F: Due Date)
    ExportWS.Columns("E:E").NumberFormat = DATE_FORMAT_EXCEL_COLUMN
    ExportWS.Columns("F:F").NumberFormat = DATE_FORMAT_EXCEL_COLUMN

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

' ===================================================================
' ADDITIONAL WIP REPORT TYPES (CLAUDE.md: Complete legacy functionality)
' ===================================================================

' **Purpose**: Generate all additional WIP report types (exact legacy functionality)
' **Original**: Interface_VBA/fwip.frm lines 302-527
' **Parameters**:
'   - ReportForm (Object): Form containing report selection options
' **Returns**: None (Subroutine)
' **Dependencies**: DataOperations.SafeOpenWorkbook, DataOperations.GetRootPath
' **Side Effects**: Creates multiple Excel files in Templates directory based on selections
' **Errors**: Individual report failures logged but don't stop other reports
' **CLAUDE.md Compliance**: Exact replacement for legacy fwip.frm additional reports functionality
Private Sub GenerateAdditionalWIPReports(ReportForm As Object)
    Dim WIPPath As String
    Dim WIPWB As Workbook
    Dim col As Long
    Dim Sortcol As String, Sortcol1 As String, Sortcol2 As String
    Dim ReportsToGenerate As Collection
    Dim LastReportFile As String

    On Error GoTo Error_Handler

    WIPPath = DataOperations.GetRootPath & "\" & WIP_FILE

    ' Determine which reports will be generated (in order of priority - last one stays open)
    Set ReportsToGenerate = New Collection
    If ReportForm.RDueDate.Value = True Then ReportsToGenerate.Add "RDueDate"
    If ReportForm.RWIP.Value = True Then ReportsToGenerate.Add "RWIP"
    If ReportForm.Job_DueDate.Value = True Then ReportsToGenerate.Add "Job_DueDate"
    If ReportForm.Office_Customer.Value = True Then ReportsToGenerate.Add "Office_Customer"
    If ReportForm.Workshop_Customer.Value = True Then ReportsToGenerate.Add "Workshop_Customer"
    If ReportForm.Office_JobNumber.Value = True Then ReportsToGenerate.Add "Office_JobNumber"
    If ReportForm.Workshop_JobNumber.Value = True Then ReportsToGenerate.Add "Workshop_JobNumber"
    If ReportForm.Job_WorkshopDueDate.Value = True Then ReportsToGenerate.Add "Job_WorkshopDueDate"

    ' Determine the last report that will be generated
    If ReportsToGenerate.Count > 0 Then
        LastReportFile = ReportsToGenerate(ReportsToGenerate.Count)
    End If

    ' Handle RDueDate report
    If ReportForm.RDueDate.Value = True Then
        Set WIPWB = DataOperations.SafeOpenWorkbook(WIPPath, True)
        If Not WIPWB Is Nothing Then
            Application.DisplayAlerts = False
            WIPWB.SaveAs (DataOperations.GetRootPath & "\TEMPLATES\Due Date.xls")
            WIPWB.Worksheets(1).Range("A1").Select
            Application.DisplayAlerts = True
            ' Only close if this is not the last report
            If LastReportFile <> "RDueDate" Then
                DataOperations.SafeCloseWorkbook WIPWB
            End If
        End If
    End If

    ' Handle RWIP report
    If ReportForm.RWIP.Value = True Then
        Set WIPWB = DataOperations.SafeOpenWorkbook(WIPPath, True)
        If Not WIPWB Is Nothing Then
            col = GetLastColumn(WIPWB.Worksheets(1))
            With WIPWB.Worksheets(1)
                .Range("A2", .Range("A2").Offset(100, col - 1).Address).Sort _
                    Key1:=.Range("A3"), Order1:=xlAscending, Header:=xlYes, _
                    OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
                .Range("A1").Select
            End With
            ' Only close if this is not the last report
            If LastReportFile <> "RWIP" Then
                DataOperations.SafeCloseWorkbook WIPWB, False
            End If
        End If
    End If

    ' Handle Job_DueDate report
    If ReportForm.Job_DueDate.Value = True Then
        GenerateJobDueDateReport (LastReportFile = "Job_DueDate")
    End If

    ' Handle Office_Customer report
    If ReportForm.Office_Customer.Value = True Then
        GenerateOfficeCustomerReport (LastReportFile = "Office_Customer")
    End If

    ' Handle Workshop_Customer report
    If ReportForm.Workshop_Customer.Value = True Then
        GenerateWorkshopCustomerReport (LastReportFile = "Workshop_Customer")
    End If

    ' Handle Office_JobNumber report
    If ReportForm.Office_JobNumber.Value = True Then
        GenerateOfficeJobNumberReport (LastReportFile = "Office_JobNumber")
    End If

    ' Handle Workshop_JobNumber report
    If ReportForm.Workshop_JobNumber.Value = True Then
        GenerateWorkshopJobNumberReport (LastReportFile = "Workshop_JobNumber")
    End If

    ' Handle Job_WorkshopDueDate report
    If ReportForm.Job_WorkshopDueDate.Value = True Then
        GenerateWorkshopDueDateReport (LastReportFile = "Job_WorkshopDueDate")
    End If

    Exit Sub

Error_Handler:
    If Not WIPWB Is Nothing Then DataOperations.SafeCloseWorkbook WIPWB, False
    SystemCore.LogError Err.Number, Err.Description, "GenerateAdditionalWIPReports", "ReportingSystem"
End Sub

' **Purpose**: Generate Job Due Date report (exact legacy functionality)
' **Original**: Interface_VBA/fwip.frm lines 323-353
Private Sub GenerateJobDueDateReport(Optional KeepOpen As Boolean = False)
    Dim WIPWB As Workbook
    Dim WIPPath As String
    Dim col As Long
    Dim Sortcol As String

    On Error GoTo Error_Handler

    WIPPath = DataOperations.GetRootPath & "\" & WIP_FILE
    Set WIPWB = DataOperations.SafeOpenWorkbook(WIPPath, True)

    If Not WIPWB Is Nothing Then
        With WIPWB.Worksheets(1)
            ' Find CustomerDelivery_Date column
            Sortcol = FindColumnAddress(.Cells, "CustomerDelivery_Date")

            If Sortcol <> "" Then
                col = GetLastColumn(WIPWB.Worksheets(1))
                .Range("A3", .Range("A3").Offset(1000, col - 1).Address).Sort _
                    Key1:=.Range(Sortcol).Offset(2, 0), Order1:=xlAscending, Header:=xlNo, _
                    OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom

                ShowOfficeCols WIPWB.Worksheets(1)

                ' Apply date formatting to date columns
                .Columns("A:A").NumberFormat = DATE_FORMAT_EXCEL_COLUMN

                .Range("B1").Select

                With .PageSetup
                    .CenterHeader = "OFFICE DUE DATE"
                    .RightHeader = Format(Now, DATE_FORMAT_DISPLAY_TIME)
                End With

                Application.DisplayAlerts = False
                WIPWB.SaveAs (DataOperations.GetRootPath & "\TEMPLATES\CustomerDelivery_Date.xls")
                Application.DisplayAlerts = True
            End If
        End With
        ' Only close if KeepOpen is False
        If Not KeepOpen Then
            DataOperations.SafeCloseWorkbook WIPWB, False
        End If
    End If

    Exit Sub

Error_Handler:
    If Not WIPWB Is Nothing Then DataOperations.SafeCloseWorkbook WIPWB, False
    SystemCore.LogError Err.Number, Err.Description, "GenerateJobDueDateReport", "ReportingSystem"
End Sub

' **Purpose**: Generate Office Customer report (exact legacy functionality)
' **Original**: Interface_VBA/fwip.frm lines 355-388
Private Sub GenerateOfficeCustomerReport(Optional KeepOpen As Boolean = False)
    Dim WIPWB As Workbook
    Dim WIPPath As String
    Dim col As Long
    Dim Sortcol1 As String, Sortcol2 As String

    On Error GoTo Error_Handler

    WIPPath = DataOperations.GetRootPath & "\" & WIP_FILE
    Set WIPWB = DataOperations.SafeOpenWorkbook(WIPPath, True)

    If Not WIPWB Is Nothing Then
        With WIPWB.Worksheets(1)
            ' Find sorting columns
            Sortcol1 = FindColumnAddress(.Cells, "Customer")
            Sortcol2 = FindColumnAddress(.Cells, "Job_Number")

            If Sortcol1 <> "" And Sortcol2 <> "" Then
                col = GetLastColumn(WIPWB.Worksheets(1))
                .Range("A3", .Range("A3").Offset(1000, col - 1).Address).Sort _
                    Key1:=.Range(Sortcol1).Offset(2, 0), Order1:=xlAscending, Header:=xlNo, _
                    Key2:=.Range(Sortcol2).Offset(2, 0), Order2:=xlAscending, _
                    OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom

                ShowOfficeCols WIPWB.Worksheets(1)

                ' Apply date formatting to date columns
                .Columns("A:A").NumberFormat = DATE_FORMAT_EXCEL_COLUMN

                .Range("B1").Select

                With .PageSetup
                    .CenterHeader = "OFFICE CUSTOMER"
                    .RightHeader = Format(Now, DATE_FORMAT_DISPLAY_TIME)
                End With

                Application.DisplayAlerts = False
                WIPWB.SaveAs (DataOperations.GetRootPath & "\TEMPLATES\Office_Customer.xls")
                Application.DisplayAlerts = True
            End If
        End With
        ' Only close if KeepOpen is False
        If Not KeepOpen Then
            DataOperations.SafeCloseWorkbook WIPWB, False
        End If
    End If

    Exit Sub

Error_Handler:
    If Not WIPWB Is Nothing Then DataOperations.SafeCloseWorkbook WIPWB, False
    SystemCore.LogError Err.Number, Err.Description, "GenerateOfficeCustomerReport", "ReportingSystem"
End Sub

' **Purpose**: Generate Workshop Customer report (exact legacy functionality)
' **Original**: Interface_VBA/fwip.frm lines 391-426
Private Sub GenerateWorkshopCustomerReport(Optional KeepOpen As Boolean = False)
    Dim WIPWB As Workbook
    Dim WIPPath As String
    Dim col As Long
    Dim Sortcol1 As String, Sortcol2 As String

    On Error GoTo Error_Handler

    WIPPath = DataOperations.GetRootPath & "\" & WIP_FILE
    Set WIPWB = DataOperations.SafeOpenWorkbook(WIPPath, True)

    If Not WIPWB Is Nothing Then
        With WIPWB.Worksheets(1)
            ' Find sorting columns
            Sortcol1 = FindColumnAddress(.Cells, "Customer")
            Sortcol2 = FindColumnAddress(.Cells, "Job_Number")

            If Sortcol1 <> "" And Sortcol2 <> "" Then
                col = GetLastColumn(WIPWB.Worksheets(1))
                .Range("A3", .Range("A3").Offset(1000, col - 1).Address).Sort _
                    Key1:=.Range(Sortcol1).Offset(2, 0), Order1:=xlAscending, Header:=xlNo, _
                    Key2:=.Range(Sortcol2).Offset(2, 0), Order2:=xlAscending, _
                    OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom

                ShowWorkshopCols WIPWB.Worksheets(1)

                ' Apply date formatting to date columns
                .Columns("A:A").NumberFormat = DATE_FORMAT_EXCEL_COLUMN

                .Range("B1").Select

                With .PageSetup
                    .CenterHeader = "WORKSHOP CUSTOMER"
                    .RightHeader = Format(Now, DATE_FORMAT_DISPLAY_TIME)
                End With

                Application.DisplayAlerts = False
                WIPWB.SaveAs (DataOperations.GetRootPath & "\TEMPLATES\Workshop_Customer.xls")
                Application.DisplayAlerts = True
            End If
        End With
        ' Only close if KeepOpen is False
        If Not KeepOpen Then
            DataOperations.SafeCloseWorkbook WIPWB, False
        End If
    End If

    Exit Sub

Error_Handler:
    If Not WIPWB Is Nothing Then DataOperations.SafeCloseWorkbook WIPWB, False
    SystemCore.LogError Err.Number, Err.Description, "GenerateWorkshopCustomerReport", "ReportingSystem"
End Sub

' **Purpose**: Generate Office Job Number report (exact legacy functionality)
' **Original**: Interface_VBA/fwip.frm lines 429-460
Private Sub GenerateOfficeJobNumberReport(Optional KeepOpen As Boolean = False)
    Dim WIPWB As Workbook
    Dim WIPPath As String
    Dim col As Long
    Dim Sortcol As String

    On Error GoTo Error_Handler

    WIPPath = DataOperations.GetRootPath & "\" & WIP_FILE
    Set WIPWB = DataOperations.SafeOpenWorkbook(WIPPath, True)

    If Not WIPWB Is Nothing Then
        With WIPWB.Worksheets(1)
            ' Find Converted_JN column
            Sortcol = FindColumnAddress(.Cells, "Converted_JN")

            If Sortcol <> "" Then
                col = GetLastColumn(WIPWB.Worksheets(1))
                .Range("A3", .Range("A3").Offset(1000, col - 1).Address).Sort _
                    Key1:=.Range(Sortcol).Offset(2, 0), Order1:=xlAscending, Header:=xlNo, _
                    OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
                    DataOption1:=xlSortTextAsNumbers

                ShowOfficeCols WIPWB.Worksheets(1)

                ' Apply date formatting to date columns
                .Columns("A:A").NumberFormat = DATE_FORMAT_EXCEL_COLUMN

                .Range("B1").Select

                With .PageSetup
                    .CenterHeader = "OFFICE JOB NUMBER"
                    .RightHeader = Format(Now, DATE_FORMAT_DISPLAY_TIME)
                End With

                Application.DisplayAlerts = False
                WIPWB.SaveAs (DataOperations.GetRootPath & "\TEMPLATES\Office_JobNumber.xls")
                Application.DisplayAlerts = True
            End If
        End With
        ' Only close if KeepOpen is False
        If Not KeepOpen Then
            DataOperations.SafeCloseWorkbook WIPWB, False
        End If
    End If

    Exit Sub

Error_Handler:
    If Not WIPWB Is Nothing Then DataOperations.SafeCloseWorkbook WIPWB, False
    SystemCore.LogError Err.Number, Err.Description, "GenerateOfficeJobNumberReport", "ReportingSystem"
End Sub

' **Purpose**: Generate Workshop Job Number report (exact legacy functionality)
' **Original**: Interface_VBA/fwip.frm lines 463-494
Private Sub GenerateWorkshopJobNumberReport(Optional KeepOpen As Boolean = False)
    Dim WIPWB As Workbook
    Dim WIPPath As String
    Dim col As Long
    Dim Sortcol As String

    On Error GoTo Error_Handler

    WIPPath = DataOperations.GetRootPath & "\" & WIP_FILE
    Set WIPWB = DataOperations.SafeOpenWorkbook(WIPPath, True)

    If Not WIPWB Is Nothing Then
        With WIPWB.Worksheets(1)
            ' Find Converted_JN column
            Sortcol = FindColumnAddress(.Cells, "Converted_JN")

            If Sortcol <> "" Then
                col = GetLastColumn(WIPWB.Worksheets(1))
                .Range("A3", .Range("A3").Offset(1000, col - 1).Address).Sort _
                    Key1:=.Range(Sortcol).Offset(2, 0), Order1:=xlAscending, Header:=xlNo, _
                    OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
                    DataOption1:=xlSortTextAsNumbers

                ShowWorkshopCols WIPWB.Worksheets(1)

                ' Apply date formatting to date columns
                .Columns("A:A").NumberFormat = DATE_FORMAT_EXCEL_COLUMN

                .Range("B1").Select

                With .PageSetup
                    .CenterHeader = "WORKSHOP JOB NUMBER"
                    .RightHeader = Format(Now, DATE_FORMAT_DISPLAY_TIME)
                End With

                Application.DisplayAlerts = False
                WIPWB.SaveAs (DataOperations.GetRootPath & "\TEMPLATES\Workshop_JobNumber.xls")
                Application.DisplayAlerts = True
            End If
        End With
        ' Only close if KeepOpen is False
        If Not KeepOpen Then
            DataOperations.SafeCloseWorkbook WIPWB, False
        End If
    End If

    Exit Sub

Error_Handler:
    If Not WIPWB Is Nothing Then DataOperations.SafeCloseWorkbook WIPWB, False
    SystemCore.LogError Err.Number, Err.Description, "GenerateWorkshopJobNumberReport", "ReportingSystem"
End Sub

' **Purpose**: Generate Workshop Due Date report (exact legacy functionality)
' **Original**: Interface_VBA/fwip.frm lines 497-527
Private Sub GenerateWorkshopDueDateReport(Optional KeepOpen As Boolean = False)
    Dim WIPWB As Workbook
    Dim WIPPath As String
    Dim col As Long
    Dim Sortcol As String

    On Error GoTo Error_Handler

    WIPPath = DataOperations.GetRootPath & "\" & WIP_FILE
    Set WIPWB = DataOperations.SafeOpenWorkbook(WIPPath, True)

    If Not WIPWB Is Nothing Then
        With WIPWB.Worksheets(1)
            ' Find Job_WorkshopDueDate column
            Sortcol = FindColumnAddress(.Cells, "Job_WorkshopDueDate")

            If Sortcol <> "" Then
                col = GetLastColumn(WIPWB.Worksheets(1))
                .Range("A3", .Range("A3").Offset(1000, col - 1).Address).Sort _
                    Key1:=.Range(Sortcol).Offset(2, 0), Order1:=xlAscending, Header:=xlNo, _
                    OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom

                ShowWorkshopCols WIPWB.Worksheets(1)

                ' Apply date formatting to date columns
                .Columns("A:A").NumberFormat = DATE_FORMAT_EXCEL_COLUMN

                .Range("B1").Select

                With .PageSetup
                    .CenterHeader = "WORKSHOP DUE DATE"
                    .RightHeader = Format(Now, DATE_FORMAT_DISPLAY_TIME)
                End With

                Application.DisplayAlerts = False
                WIPWB.SaveAs (DataOperations.GetRootPath & "\TEMPLATES\Job_WorkshopDueDate.xls")
                Application.DisplayAlerts = True
            End If
        End With
        ' Only close if KeepOpen is False
        If Not KeepOpen Then
            DataOperations.SafeCloseWorkbook WIPWB, False
        End If
    End If

    Exit Sub

Error_Handler:
    If Not WIPWB Is Nothing Then DataOperations.SafeCloseWorkbook WIPWB, False
    SystemCore.LogError Err.Number, Err.Description, "GenerateWorkshopDueDateReport", "ReportingSystem"
End Sub

' ===================================================================
' HELPER FUNCTIONS FOR REPORT GENERATION
' ===================================================================

' **Purpose**: Find column address for given header value
' **Original**: Interface_VBA/fwip.frm Do...Loop searching logic
Private Function FindColumnAddress(ByRef HeaderCells As Range, ByVal HeaderValue As String) As String
    Dim Cell As Range
    Dim SearchRange As Range

    On Error GoTo Error_Handler

    Set SearchRange = HeaderCells.Rows(1).Cells
    Set Cell = SearchRange.Find(HeaderValue, LookIn:=xlValues, LookAt:=xlWhole)

    If Not Cell Is Nothing Then
        FindColumnAddress = Cell.Address
    Else
        FindColumnAddress = ""
    End If

    Exit Function

Error_Handler:
    FindColumnAddress = ""
End Function

' **Purpose**: Get last column with data in worksheet
' **Original**: Interface_VBA/fwip.frm col calculation logic
Private Function GetLastColumn(ByRef ws As Worksheet) As Long
    On Error GoTo Error_Handler

    ws.Range("BB1").Select
    Selection.End(xlToLeft).Select
    GetLastColumn = ActiveCell.Column
    Exit Function

Error_Handler:
    GetLastColumn = 1
End Function

' **Purpose**: Show only office-relevant columns (exact legacy functionality)
' **Original**: Interface_VBA/fwip.frm ShowOfficeCols() function lines 572-613
Private Sub ShowOfficeCols(ByRef ws As Worksheet)
    Dim Cell As Range

    On Error GoTo Error_Handler

    ws.Range("A1").Select
    Do While ActiveCell.Value <> ""
        Selection.EntireColumn.Hidden = True

        Select Case UCase(ActiveCell.Value)
            Case "JOB_STARTDATE", "JOB_URGENCY", "CUSTOMER", "JOB_NUMBER"
                Selection.EntireColumn.Hidden = False
            Case "COMPONENT_QUANTITY", "COMPONENT_CODE", "COMPONENT_DESCRIPTION"
                Selection.EntireColumn.Hidden = False
            Case "COMPONENT_COMMENTS", "CUSTOMERDELIVERY_DATE", "CUSTOMERORDERNUMBER"
                Selection.EntireColumn.Hidden = False
            Case "COMPONENT_PRICE", "COMPONENT_DRAWINGNUMBER_SAMPLENUMBER"
                Selection.EntireColumn.Hidden = False
        End Select

        ActiveCell.Offset(0, 1).Select
    Loop

    Exit Sub

Error_Handler:
    SystemCore.LogError Err.Number, Err.Description, "ShowOfficeCols", "ReportingSystem"
End Sub

' **Purpose**: Show only workshop-relevant columns (exact legacy functionality)
' **Original**: Interface_VBA/fwip.frm ShowWorkshopCols() function lines 615-716
Private Sub ShowWorkshopCols(ByRef ws As Worksheet)
    Dim Cell As Range
    Dim i As Integer

    On Error GoTo Error_Handler

    ws.Range("A1").Select
    Do While ActiveCell.Value <> ""
        Selection.EntireColumn.Hidden = True

        Select Case UCase(ActiveCell.Value)
            Case "JOB_STARTDATE", "JOB_URGENCY", "CUSTOMER", "JOB_NUMBER"
                Selection.EntireColumn.Hidden = False
            Case "JOB_WORKSHOPDUEDATE", "COMPONENT_QUANTITY", "COMPONENT_CODE"
                Selection.EntireColumn.Hidden = False
            Case "COMPONENT_DESCRIPTION", "COMPONENT_COMMENTS", "COMPONENT_DRAWINGNUMBER_SAMPLENUMBER"
                Selection.EntireColumn.Hidden = False
        End Select

        ' Show all Operation columns (Operation01_Type through Operation15_Operator)
        For i = 1 To 15
            If UCase(ActiveCell.Value) = UCase("Operation" & Format(i, "00") & "_Type") Or _
               UCase(ActiveCell.Value) = UCase("Operation" & Format(i, "00") & "_Operator") Then
                Selection.EntireColumn.Hidden = False
            End If
        Next i

        ActiveCell.Offset(0, 1).Select
    Loop

    Exit Sub

Error_Handler:
    SystemCore.LogError Err.Number, Err.Description, "ShowWorkshopCols", "ReportingSystem"
End Sub