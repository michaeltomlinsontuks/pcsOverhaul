Attribute VB_Name = "WIPReportManager"
' **Purpose**: WIP report generation and management extracted from fwip.frm
' **Original**: Interface_VBA/fwip.frm.Go_Click and related functions
' **CLAUDE.md Compliance**: Extract business logic from forms to modules
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
' **Original**: fwip.frm.Go_Click (lines 32-321)
' **Parameters**:
'   - ReportForm (Object): Form containing report selection options (ROperation, ROperator, etc.)
' **Returns**: Boolean - True if reports generated successfully, False if failed
' **File Dependencies**: WIP.xls from root path
' **Form Usage**: Extracted from fwip.frm to make it a thin wrapper
Public Function GenerateWIPReports(ReportForm As Object) As Boolean
    Dim Job(1 To 5000) As Jobs
    Dim JobCount As Integer
    Dim WIPPath As String
    Dim WIPWB As Workbook

    On Error GoTo Error_Handler

    ' Update form status
    ReportForm.Label1.Caption = "Please Wait"
    Application.DisplayAlerts = False

    ' Load WIP data
    WIPPath = DataManager.GetRootPath & "\" & WIP_FILE

    If Not DataManager.FileExists(WIPPath) Then
        ReportForm.Label1.Caption = "Ready"
        ValidationFramework.ShowInformation "WIP.xls file not found at: " & WIPPath & vbCrLf & vbCrLf & _
               "The WIP.xls file is created when job cards save their data." & vbCrLf & _
               "Please ensure you have some active jobs that have been saved.", "WIP.xls Not Found"
        GenerateWIPReports = False
        Exit Function
    End If

    ' Open and load WIP data
    Set WIPWB = DataManager.SafeOpenWorkbook(WIPPath)
    If WIPWB Is Nothing Then
        ReportForm.Label1.Caption = "Ready"
        ValidationFramework.ShowError "Unable to open WIP.xls at: " & WIPPath, "File Access Error"
        GenerateWIPReports = False
        Exit Function
    End If

    JobCount = LoadWIPDataFromWorkbook(WIPWB, Job)
    DataManager.SafeCloseWorkbook WIPWB, False

    If JobCount = 0 Then
        ReportForm.Label1.Caption = "Ready"
        ValidationFramework.ShowInformation "No WIP data found in WIP.xls file." & vbCrLf & vbCrLf & _
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
    ValidationFramework.ShowInformation "WIP reports have been generated successfully!" & vbCrLf & _
           "Reports saved to Templates directory:" & vbCrLf & _
           "- Operation reports (if selected)" & vbCrLf & _
           "- Operator reports (if selected)" & vbCrLf & _
           "Check your Templates folder for the generated files.", "Reports Generated"

    ReportForm.Label1.Caption = "Ready - Select report types and click Go"
    GenerateWIPReports = True
    Exit Function

Error_Handler:
    Application.DisplayAlerts = True
    If Not WIPWB Is Nothing Then DataManager.SafeCloseWorkbook WIPWB, False
    If Not ReportForm Is Nothing Then
        ReportForm.Show
        ReportForm.Label1.Caption = "Ready"
    End If
    CoreFramework.LogError Err.Number, "GenerateWIPReports", "WIPReportManager", Err.Description
    GenerateWIPReports = False
End Function

' ===================================================================
' PRIVATE HELPER FUNCTIONS
' ===================================================================

' **Purpose**: Load WIP data from consolidated workbook into Jobs array
' **Original**: fwip.frm.Go_Click data loading section (lines 69-128)
' **Parameters**:
'   - WIPWB (Workbook): Opened WIP workbook
'   - Job (Jobs array): Array to populate with job data
' **Returns**: Integer - Number of jobs loaded
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
    CoreFramework.LogError Err.Number, "LoadWIPDataFromWorkbook", "WIPReportManager", Err.Description
    LoadWIPDataFromWorkbook = 0
End Function

' **Purpose**: Generate operation-based reports from WIP data
' **Original**: fwip.frm.Go_Click operation report section (lines 132-204)
' **Parameters**:
'   - Job (Jobs array): Array containing job data
'   - JobCount (Integer): Number of jobs in array
Private Sub GenerateOperationReports(ByRef Job() As Jobs, JobCount As Integer)
    Dim TempSheet As String
    Dim OPP As String
    Dim j As Integer
    Dim k As Integer
    Dim sh As Worksheet

    On Error GoTo Error_Handler

    Workbooks.Add

    OPP = ""
    If MsgBox("Specific Operation?", vbYesNo) = vbYes Then
        OPP = InputBox("Which Operation")
    End If

    For j = 1 To JobCount
        With Job(j)
            For k = 1 To 15
                If OPP <> "" Then
                    If Trim(UCase(.OperatorType(k))) <> Trim(UCase(OPP)) Then GoTo SkipOPP
                End If

                If .OperatorType(k) <> "" Then
                    TempSheet = "OPERATION - " & .OperatorType(k)
                    On Error GoTo AddOperationSheet
                    Sheets(CoreFramework.RemoveInvalidCharacters(Trim(TempSheet))).Select
                    On Error GoTo Error_Handler

                    ActiveCell.FormulaR1C1 = .Dat
                    ActiveCell.Offset(0, 1).FormulaR1C1 = .Cust
                    ActiveCell.Offset(0, 2).FormulaR1C1 = .Job
                    ActiveCell.Offset(0, 3).FormulaR1C1 = .JobD
                    ActiveCell.Offset(0, 4).FormulaR1C1 = .Qty
                    ActiveCell.Offset(0, 5).FormulaR1C1 = .Cod
                    ActiveCell.Offset(0, 6).FormulaR1C1 = .Desc
                    ActiveCell.Offset(0, 7).FormulaR1C1 = .Remarks
                    ActiveCell.Offset(0, 8).FormulaR1C1 = .DDat

                    If k > 1 Then
                        If .OperatorType(k - 1) = "" Then
                            ActiveCell.Offset(0, 9).FormulaR1C1 = "*"
                            Selection.EntireRow.Font.Bold = True
                        End If
                    End If
                    ActiveCell.Offset(1, 0).Select
                End If
                TempSheet = ""
SkipOPP:
            Next k
        End With
    Next j

    ' Format all operation sheets
    For Each sh In Sheets
        FormatWorksheet sh
    Next sh

    DeleteDefaultSheets

    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs (DataManager.GetRootPath & "\TEMPLATES\Operation.xls")
    Exit Sub

AddOperationSheet:
    CreateReportSheet TempSheet
    Resume

Error_Handler:
    CoreFramework.LogError Err.Number, "GenerateOperationReports", "WIPReportManager", Err.Description
End Sub

' **Purpose**: Generate operator-based reports from WIP data
' **Original**: fwip.frm.Go_Click operator report section (lines 206-268)
' **Parameters**:
'   - Job (Jobs array): Array containing job data
'   - JobCount (Integer): Number of jobs in array
Private Sub GenerateOperatorReports(ByRef Job() As Jobs, JobCount As Integer)
    Dim TempSheet As String
    Dim j As Integer
    Dim k As Integer
    Dim sh As Worksheet

    On Error GoTo Error_Handler

    Workbooks.Add

    For j = 1 To JobCount
        With Job(j)
            For k = 1 To 15
                If Trim(.OperatorN(k)) <> "" Then
                    TempSheet = CoreFramework.RemoveInvalidCharacters("OPERATOR - " & Trim(.OperatorN(k)))
                    On Error GoTo AddOperatorSheet
                    Sheets(TempSheet).Select
                    On Error GoTo Error_Handler

                    ActiveCell.FormulaR1C1 = .Dat
                    ActiveCell.Offset(0, 1).FormulaR1C1 = .Cust
                    ActiveCell.Offset(0, 2).FormulaR1C1 = .Job
                    ActiveCell.Offset(0, 3).FormulaR1C1 = .JobD
                    ActiveCell.Offset(0, 4).FormulaR1C1 = .Qty
                    ActiveCell.Offset(0, 5).FormulaR1C1 = .Cod
                    ActiveCell.Offset(0, 6).FormulaR1C1 = .Desc
                    ActiveCell.Offset(0, 7).FormulaR1C1 = .Remarks
                    ActiveCell.Offset(0, 8).FormulaR1C1 = .DDat

                    If k > 1 Then
                        If .OperatorN(k - 1) = "" Then
                            ActiveCell.Offset(0, 9).FormulaR1C1 = "*"
                            Selection.EntireRow.Font.Bold = True
                        End If
                    End If

                    ActiveCell.Offset(1, 0).Select
                End If
                TempSheet = ""
            Next k
        End With
    Next j

    ' Format all operator sheets
    For Each sh In Sheets
        FormatWorksheet sh
    Next sh

    DeleteDefaultSheets

    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs (DataManager.GetRootPath & "\TEMPLATES\Operator.xls")
    Exit Sub

AddOperatorSheet:
    CreateReportSheet TempSheet
    Resume

Error_Handler:
    CoreFramework.LogError Err.Number, "GenerateOperatorReports", "WIPReportManager", Err.Description
End Sub

' **Purpose**: Parse job number string for proper sorting
' **Original**: fwip.frm.ParseJobNumberForSorting (lines 323-340)
' **Parameters**: jobString (String): Job number string to parse
' **Returns**: Double - Numeric value for sorting with sub-parts handled
Private Function ParseJobNumberForSorting(jobString As String) As Double
    Dim parts() As String
    Dim mainPart As Double
    Dim subPart As Double

    If InStr(jobString, "-") > 0 Then
        parts = Split(jobString, "-")
        mainPart = Val(parts(0))
        If UBound(parts) > 0 Then
            subPart = Val(parts(1))
            ParseJobNumberForSorting = mainPart + (subPart / 1000000)
        Else
            ParseJobNumberForSorting = mainPart
        End If
    Else
        ParseJobNumberForSorting = Val(jobString)
    End If
End Function

' **Purpose**: Format worksheet for report presentation
' **Original**: fwip.frm formatting sections in operation and operator reports
' **Parameters**: sh (Worksheet): Worksheet to format
Private Sub FormatWorksheet(sh As Worksheet)
    On Error GoTo Error_Handler

    sh.Select
    Cells.EntireColumn.AutoFit
    Range("A1:i5000").Select
    Selection.Sort Key1:=Range("H2"), Order1:=xlAscending, Key2:=Range("G2") _
        , Order2:=xlAscending, Header:=xlYes, OrderCustom:=1, MatchCase:= _
        False, Orientation:=xlTopToBottom

    With ActiveSheet.PageSetup
        .CenterHeader = ActiveSheet.Name
        .RightHeader = "&D &T"
    End With

    FormatWorksheetBorders
    Range("A1").Select
    Range("a:a").NumberFormat = "DD MMM YYYY"
    Range("i:i").NumberFormat = "DD MMM YYYY"

Exit Sub

Error_Handler:
    CoreFramework.LogError Err.Number, "FormatWorksheet", "WIPReportManager", Err.Description
End Sub

' **Purpose**: Create new report sheet with standard headers
' **Original**: fwip.frm.AddSheet label (lines 288-315)
' **Parameters**: SheetName (String): Name for the new sheet
Private Sub CreateReportSheet(SheetName As String)
    On Error GoTo Error_Handler

    Sheets.Add
    ActiveSheet.Name = CoreFramework.RemoveInvalidCharacters(SheetName)
    ActiveSheet.PageSetup.CenterHeader = SheetName

    ' Set up headers
    ActiveCell.FormulaR1C1 = "DATE"
    ActiveCell.Offset(0, 1).FormulaR1C1 = "CUSTOMER"
    ActiveCell.Offset(0, 2).FormulaR1C1 = "JOB"
    ActiveCell.Offset(0, 3).FormulaR1C1 = "JOB"
    ActiveCell.Offset(0, 4).FormulaR1C1 = "QTY"
    ActiveCell.Offset(0, 5).FormulaR1C1 = "COMPONENT CODE"
    ActiveCell.Offset(0, 6).FormulaR1C1 = "COMPONENT DESCRIPTION"
    ActiveCell.Offset(0, 7).FormulaR1C1 = "REMARKS"
    ActiveCell.Offset(0, 8).FormulaR1C1 = "DUE DATE"

    ' Format columns
    Columns("h:h").NumberFormat = "dd mmm"
    Columns("A:A").NumberFormat = "dd mmm"
    Selection.EntireRow.Font.Bold = True

    ' Set column widths
    Columns("A:A").ColumnWidth = 10
    Columns("b:b").ColumnWidth = 18
    Columns("c:c").ColumnWidth = 10
    Columns("e:e").ColumnWidth = 6
    Columns("g:g").ColumnWidth = 30
    Columns("h:h").ColumnWidth = 20
    Columns("i:i").ColumnWidth = 10
    Cells.RowHeight = 30

    ActiveCell.Offset(1, 0).Select

Error_Handler:
    CoreFramework.LogError Err.Number, "CreateReportSheet", "WIPReportManager", Err.Description
End Sub

' **Purpose**: Apply borders to worksheet for professional appearance
' **Original**: fwip.frm.FormatWorksheetBorders (lines 472-510)
Private Sub FormatWorksheetBorders()
    On Error Resume Next ' In case older Excel doesn't support some border properties

    Cells.Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 1 ' Black color (compatible with older Excel)
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 1
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 1
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 1
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 1
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 1
    End With

    On Error GoTo 0
End Sub

' **Purpose**: Delete default Excel sheets from new workbook
' **Original**: fwip.frm multiple DeleteSheet calls
Private Sub DeleteDefaultSheets()
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("sheet1").Delete
    Worksheets("sheet2").Delete
    Worksheets("sheet3").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub