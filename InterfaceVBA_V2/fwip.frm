VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FWIP
   Caption         =   "WIP Reports"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4470
   OleObjectBlob   =   "fwip.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fwip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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

Private Sub Go_Click()
    Dim TempSheet As String
    Dim Job(1 To 5000) As Jobs
    Dim i As Integer
    Dim col As Integer
    Dim OPP As String
    Dim j As Integer
    Dim k As Integer
    Dim x As Integer
    Dim sh As Worksheet
    Dim Sortcol As String, Sortcol1 As String, Sortcol2 As String

    On Error GoTo Error_Handler

    fwip.Label1.Caption = "Please Wait"
    Application.DisplayAlerts = False

    ' Check for WIP.xls file and create if needed
    Dim WIPPath As String
    WIPPath = DataManager.GetRootPath & "\WIP.xls"

    ' If WIP.xls doesn't exist, create a sample one or show helpful message
    If Not DataManager.FileExists(WIPPath) Then
        If MsgBox("WIP.xls file not found at: " & WIPPath & vbCrLf & vbCrLf & _
                 "Would you like to create a sample WIP.xls file with demo data?", _
                 vbYesNo + vbQuestion, "Create Sample WIP File?") = vbYes Then
            CreateSampleWIPFile WIPPath
        Else
            MsgBox "WIP Reports require a WIP.xls file. Please create one or place your WIP data in: " & vbCrLf & WIPPath, vbInformation
            Exit Sub
        End If
    End If

    ' Open WIP.xls using new module structure
    Dim WIPWB As Workbook
    Set WIPWB = DataManager.SafeOpenWorkbook(WIPPath)
    If WIPWB Is Nothing Then
        MsgBox "Unable to open WIP.xls at: " & WIPPath, vbCritical
        Exit Sub
    End If

    ' Load WIP data using original structure
    WIPWB.Activate
    Range("bb1").Select
    Selection.End(xlToLeft).Select
    col = ActiveCell.Column

    Range("A1").Select
    Selection.End(xlDown).Select

    Range("A2", Range("A2").Offset(ActiveCell.Row, col - 1).Address).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Sort Key1:=Range("h3"), Order1:=xlAscending, Header:=xlYes, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom

    Range("A3").Select
    fwip.Hide
    i = 0
    If ActiveCell.FormulaR1C1 <> "" Then
        Do
            i = i + 1
            With Job(i)
                .Dat = ActiveCell.Offset(0, 0).Value
                .Cust = ActiveCell.Offset(0, 1).Value
                .Job = ActiveCell.Offset(0, 2).Value
                .JobD = ParseJobNumberForSorting(ActiveCell.Offset(0, 3).Value)
                .Qty = ActiveCell.Offset(0, 4).Value
                .Cod = ActiveCell.Offset(0, 5).Value
                .Desc = ActiveCell.Offset(0, 6).Value
                .Remarks = ActiveCell.Offset(0, 8).Value
                .DDat = ActiveCell.Offset(0, 12).Value
                x = 0
                For j = 1 To 30 Step 2
                    x = x + 1
                    .OperatorType(x) = ActiveCell.Offset(0, 14 + j).Value
                Next j
                x = 0
                For j = 1 To 30 Step 2
                    x = x + 1
                    .OperatorN(x) = ActiveCell.Offset(0, 15 + j).Value
                Next j
            End With
            ActiveCell.Offset(1, 0).Select
        Loop Until ActiveCell.FormulaR1C1 = ""
    Else
        DataManager.SafeCloseWorkbook WIPWB, False
        Exit Sub
    End If

    fwip.Hide

    ' Operation Reports
    If ROperation.Value = True Then
        Workbooks.Add

        OPP = ""
        If MsgBox("Specific Operation?", vbYesNo) = vbYes Then
            OPP = InputBox("Which Operation")
        End If

        For j = 1 To i
            With Job(j)
                For k = 1 To 15
                    If OPP <> "" Then
                        If Trim(UCase(.OperatorType(k))) <> Trim(UCase(OPP)) Then GoTo SkipOPP
                    End If

                    If .OperatorType(k) <> "" Then
                        TempSheet = "OPERATION - " & .OperatorType(k)
                        On Error GoTo AddSheet
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
        Next sh

        DeleteSheet ("sheet1")
        DeleteSheet ("sheet2")
        DeleteSheet ("sheet3")

        Application.DisplayAlerts = False
        ActiveWorkbook.SaveAs (DataManager.GetRootPath & "\TEMPLATES\Operation.xls")
    End If

    ' Operator Reports
    If ROperator.Value = True Then
        Workbooks.Add

        For j = 1 To i
            With Job(j)
                For k = 1 To 15
                    If Trim(.OperatorN(k)) <> "" Then
                        TempSheet = CoreFramework.RemoveInvalidCharacters("OPERATOR - " & Trim(.OperatorN(k)))
                        On Error GoTo AddSheet
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
            Range("a:a").NumberFormat = "DD MMM YYYY"
            Range("i:i").NumberFormat = "DD MMM YYYY"
            Range("A1").Select
        Next sh

        DeleteSheet ("sheet1")
        DeleteSheet ("sheet2")
        DeleteSheet ("sheet3")

        Application.DisplayAlerts = False
        ActiveWorkbook.SaveAs (DataManager.GetRootPath & "\TEMPLATES\Operator.xls")
    End If

    ' Close WIP workbook
    Windows("wip.xls").Activate
    If fwip.RDueDate.Value = True Then
        Application.DisplayAlerts = False
        ActiveWorkbook.SaveAs (DataManager.GetRootPath & "\TEMPLATES\Due Date.xls")
        Range("a1").Select
    Else
        ActiveWorkbook.Close False
    End If

    ' WIP Report sorted by date
    If fwip.RWIP.Value = True Then
        Set WIPWB = DataManager.SafeOpenWorkbook(DataManager.GetRootPath & "\WIP.xls")
        WIPWB.Activate

        Range("A2", Range("A2").Offset(ActiveCell.Row, col - 1).Address).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Sort Key1:=Range("A3"), Order1:=xlAscending, Header:=xlYes, _
            OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom

        Range("A1").Select
    End If

    ' Customer Due Date Report
    If fwip.Job_DueDate.Value = True Then
        Set WIPWB = DataManager.SafeOpenWorkbook(DataManager.GetRootPath & "\WIP.xls")
        WIPWB.Activate

        Range("A1").Select
        Do
            ActiveCell.Offset(0, 1).Select
        Loop Until ActiveCell.Value = "CustomerDelivery_Date" Or ActiveCell.FormulaR1C1 = ""

        Sortcol = ActiveCell.Address

        Range("A3", Range("A3").Offset(0, col - 1).Address).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Sort Key1:=Range(Range(Sortcol).Offset(2, 0).Address), Order1:=xlAscending, Header:=xlNo, _
            OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom

        ShowOfficeCols
        Range("b1").Select
        Application.DisplayAlerts = False

        With ActiveSheet.PageSetup
            .CenterHeader = "OFFICE DUE DATE"
            .RightHeader = "&D &T"
        End With

        ActiveWorkbook.SaveAs (DataManager.GetRootPath & "\TEMPLATES\CustomerDelivery_Date.xls")
    End If

    ' Office Customer Report
    If fwip.Office_Customer.Value = True Then
        Set WIPWB = DataManager.SafeOpenWorkbook(DataManager.GetRootPath & "\WIP.xls")
        WIPWB.Activate

        Range("A1").Select
        Do
            ActiveCell.Offset(0, 1).Select
            If UCase(ActiveCell.Value) = UCase("Customer") Then Sortcol1 = ActiveCell.Address
            If UCase(ActiveCell.Value) = UCase("Job_Number") Then Sortcol2 = ActiveCell.Address
        Loop Until ActiveCell.FormulaR1C1 = ""

        Range("A3", Range("A3").Offset(0, col - 1).Address).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Sort Key1:=Range(Range(Sortcol1).Offset(2, 0).Address), Order1:=xlAscending, Header:=xlNo, _
            OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom _
            , Key2:=Range(Range(Sortcol2).Offset(2, 0).Address), Order2:=xlAscending

        ShowOfficeCols
        Range("b1").Select

        With ActiveSheet.PageSetup
            .CenterHeader = "OFFICE CUSTOMER"
            .RightHeader = "&D &T"
        End With

        Application.DisplayAlerts = False
        ActiveWorkbook.SaveAs (DataManager.GetRootPath & "\TEMPLATES\Office_Customer.xls")
    End If

    ' Workshop Customer Report
    If fwip.Workshop_Customer.Value = True Then
        Set WIPWB = DataManager.SafeOpenWorkbook(DataManager.GetRootPath & "\WIP.xls")
        WIPWB.Activate

        Range("A1").Select
        Do
            ActiveCell.Offset(0, 1).Select
            If UCase(ActiveCell.Value) = UCase("Customer") Then Sortcol1 = ActiveCell.Address
            If UCase(ActiveCell.Value) = UCase("Job_Number") Then Sortcol2 = ActiveCell.Address
        Loop Until ActiveCell.FormulaR1C1 = ""

        Range("A3", Range("A3").Offset(0, col - 1).Address).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Sort Key1:=Range(Range(Sortcol1).Offset(2, 0).Address), Order1:=xlAscending, Header:=xlNo, _
            OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom _
            , Key2:=Range(Range(Sortcol2).Offset(2, 0).Address), Order2:=xlAscending

        ShowWorkshopCols
        Range("b1").Select

        With ActiveSheet.PageSetup
            .CenterHeader = "WORKSHOP CUSTOMER"
            .RightHeader = "&D &T"
        End With

        Application.DisplayAlerts = False
        ActiveWorkbook.SaveAs (DataManager.GetRootPath & "\TEMPLATES\Workshop_Customer.xls")
    End If

    ' Office Job Number Report
    If fwip.Office_JobNumber.Value = True Then
        Set WIPWB = DataManager.SafeOpenWorkbook(DataManager.GetRootPath & "\WIP.xls")
        WIPWB.Activate

        Range("A1").Select
        Do
            ActiveCell.Offset(0, 1).Select
        Loop Until ActiveCell.Value = "Converted_JN" Or ActiveCell.FormulaR1C1 = ""

        Sortcol = ActiveCell.Address

        Range("A3", Range("A3").Offset(0, col - 1).Address).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Sort Key1:=Range(Range(Sortcol).Offset(2, 0).Address), Order1:=xlAscending, Header:=xlNo, _
            OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
            DataOption1:=xlSortTextAsNumbers

        ShowOfficeCols
        Range("b1").Select

        With ActiveSheet.PageSetup
            .CenterHeader = "OFFICE JOB NUMBER"
            .RightHeader = "&D &T"
        End With
        Application.DisplayAlerts = False
        ActiveWorkbook.SaveAs (DataManager.GetRootPath & "\TEMPLATES\Office_JobNumber.xls")
    End If

    ' Workshop Job Number Report
    If fwip.Workshop_JobNumber.Value = True Then
        Set WIPWB = DataManager.SafeOpenWorkbook(DataManager.GetRootPath & "\WIP.xls")
        WIPWB.Activate

        Range("A1").Select
        Do
            ActiveCell.Offset(0, 1).Select
        Loop Until ActiveCell.Value = "Converted_JN" Or ActiveCell.FormulaR1C1 = ""

        Sortcol = ActiveCell.Address

        Range("A3", Range("A3").Offset(0, col - 1).Address).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Sort Key1:=Range(Range(Sortcol).Offset(2, 0).Address), Order1:=xlAscending, Header:=xlNo, _
            OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
            DataOption1:=xlSortTextAsNumbers

        ShowWorkshopCols

        Application.DisplayAlerts = False
        Range("b1").Select

        With ActiveSheet.PageSetup
            .CenterHeader = "WORKSHOP JOB NUMBER"
            .RightHeader = "&D &T"
        End With
        ActiveWorkbook.SaveAs (DataManager.GetRootPath & "\TEMPLATES\Workshop_JobNumber.xls")
    End If

    ' Workshop Due Date Report
    If fwip.Job_WorkshopDueDate.Value = True Then
        Set WIPWB = DataManager.SafeOpenWorkbook(DataManager.GetRootPath & "\WIP.xls")
        WIPWB.Activate

        Range("A1").Select
        Do
            ActiveCell.Offset(0, 1).Select
        Loop Until ActiveCell.Value = "Job_WorkshopDueDate" Or ActiveCell.FormulaR1C1 = ""

        Sortcol = ActiveCell.Address

        Range("A3", Range("A3").Offset(0, col - 1).Address).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Sort Key1:=Range(Range(Sortcol).Offset(2, 0).Address), Order1:=xlAscending, Header:=xlNo, _
            OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom

        ShowWorkshopCols

        Application.DisplayAlerts = False
        Range("b1").Select

        With ActiveSheet.PageSetup
            .CenterHeader = "WORKSHOP DUE DATE"
            .RightHeader = "&D &T"
        End With

        ActiveWorkbook.SaveAs (DataManager.GetRootPath & "\TEMPLATES\Job_WorkshopDueDate.xls")
    End If

    Application.DisplayAlerts = True
    Unload fwip
    Exit Sub

AddSheet:
    Sheets.Add
    ActiveSheet.Name = CoreFramework.RemoveInvalidCharacters(TempSheet)
    ActiveSheet.PageSetup.CenterHeader = TempSheet
    ActiveCell.FormulaR1C1 = "DATE"
    ActiveCell.Offset(0, 1).FormulaR1C1 = "CUSTOMER"
    ActiveCell.Offset(0, 2).FormulaR1C1 = "JOB"
    ActiveCell.Offset(0, 3).FormulaR1C1 = "JOB"
    ActiveCell.Offset(0, 4).FormulaR1C1 = "QTY"
    ActiveCell.Offset(0, 5).FormulaR1C1 = "COMPONENT CODE"
    ActiveCell.Offset(0, 6).FormulaR1C1 = "COMPONENT DESCRIPTION"
    ActiveCell.Offset(0, 7).FormulaR1C1 = "REMARKS"
    ActiveCell.Offset(0, 8).FormulaR1C1 = "DUE DATE"
    Columns("h:h").NumberFormat = "dd mmm"
    Columns("A:A").NumberFormat = "dd mmm"
    Selection.EntireRow.Font.Bold = True

    Columns("A:A").ColumnWidth = 10
    Columns("b:b").ColumnWidth = 18
    Columns("c:c").ColumnWidth = 10
    Columns("e:e").ColumnWidth = 6
    Columns("g:g").ColumnWidth = 30
    Columns("h:h").ColumnWidth = 20
    Columns("i:i").ColumnWidth = 10
    Cells.RowHeight = 30

    ActiveCell.Offset(1, 0).Select
    Resume

Error_Handler:
    Application.DisplayAlerts = True
    If Not WIPWB Is Nothing Then DataManager.SafeCloseWorkbook WIPWB, False
    CoreFramework.HandleStandardErrors Err.Number, "Go_Click", "fwip"
End Sub

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

Private Function ShowOfficeCols()
    Range("A1").Select
    Do
        Selection.EntireColumn.Hidden = True

        Select Case ActiveCell.Value
            Case "Job_StartDate"
                Selection.EntireColumn.Hidden = False
            Case "Job_Urgency"
                Selection.EntireColumn.Hidden = False
            Case "CUSTOMER"
                Selection.EntireColumn.Hidden = False
            Case "Job_Number"
                Selection.EntireColumn.Hidden = False
            Case "Component_Quantity"
                Selection.EntireColumn.Hidden = False
            Case "Component_Code"
                Selection.EntireColumn.Hidden = False
            Case "Component_Description"
                Selection.EntireColumn.Hidden = False
            Case "Component_Comments"
                Selection.EntireColumn.Hidden = False
            Case "CustomerDelivery_Date"
                Selection.EntireColumn.Hidden = False
            Case "CustomerOrderNumber"
                Selection.EntireColumn.Hidden = False
            Case "Component_Price"
                Selection.EntireColumn.Hidden = False
            Case "Component_DrawingNumber_SampleNumber"
                Selection.EntireColumn.Hidden = False
        End Select
        ActiveCell.Offset(0, 1).Select

    Loop Until ActiveCell.Value = ""
End Function

Private Function ShowWorkshopCols()
    Range("A1").Select
    Do
        Selection.EntireColumn.Hidden = True

        Select Case ActiveCell.Value
            Case "Job_StartDate"
                Selection.EntireColumn.Hidden = False
            Case "Job_Urgency"
                Selection.EntireColumn.Hidden = False
            Case "CUSTOMER"
                Selection.EntireColumn.Hidden = False
            Case "Job_Number"
                Selection.EntireColumn.Hidden = False
            Case "Job_WorkshopDueDate"
                Selection.EntireColumn.Hidden = False
            Case "Component_Quantity"
                Selection.EntireColumn.Hidden = False
            Case "Component_Code"
                Selection.EntireColumn.Hidden = False
            Case "Component_Description"
                Selection.EntireColumn.Hidden = False
            Case "Component_Comments"
                Selection.EntireColumn.Hidden = False
            Case " "
                Selection.EntireColumn.Hidden = False
            Case "Component_DrawingNumber_SampleNumber"
                Selection.EntireColumn.Hidden = False
            Case "Operation01_Type"
                Selection.EntireColumn.Hidden = False
            Case "Operation01_Operator"
                Selection.EntireColumn.Hidden = False
            Case "Operation02_Type"
                Selection.EntireColumn.Hidden = False
            Case "Operation02_Operator"
                Selection.EntireColumn.Hidden = False
            Case "Operation03_Type"
                Selection.EntireColumn.Hidden = False
            Case "Operation03_Operator"
                Selection.EntireColumn.Hidden = False
            Case "Operation04_Type"
                Selection.EntireColumn.Hidden = False
            Case "Operation04_Operator"
                Selection.EntireColumn.Hidden = False
            Case "Operation05_Type"
                Selection.EntireColumn.Hidden = False
            Case "Operation05_Operator"
                Selection.EntireColumn.Hidden = False
            Case "Operation06_Type"
                Selection.EntireColumn.Hidden = False
            Case "Operation06_Operator"
                Selection.EntireColumn.Hidden = False
            Case "Operation07_Type"
                Selection.EntireColumn.Hidden = False
            Case "Operation07_Operator"
                Selection.EntireColumn.Hidden = False
            Case "Operation08_Type"
                Selection.EntireColumn.Hidden = False
            Case "Operation08_Operator"
                Selection.EntireColumn.Hidden = False
            Case "Operation09_Type"
                Selection.EntireColumn.Hidden = False
            Case "Operation09_Operator"
                Selection.EntireColumn.Hidden = False
            Case "Operation10_Type"
                Selection.EntireColumn.Hidden = False
            Case "Operation10_Operator"
                Selection.EntireColumn.Hidden = False
            Case "Operation11_Type"
                Selection.EntireColumn.Hidden = False
            Case "Operation11_Operator"
                Selection.EntireColumn.Hidden = False
            Case "Operation12_Type"
                Selection.EntireColumn.Hidden = False
            Case "Operation12_Operator"
                Selection.EntireColumn.Hidden = False
            Case "Operation13_Type"
                Selection.EntireColumn.Hidden = False
            Case "Operation13_Operator"
                Selection.EntireColumn.Hidden = False
            Case "Operation14_Type"
                Selection.EntireColumn.Hidden = False
            Case "Operation14_Operator"
                Selection.EntireColumn.Hidden = False
            Case "Operation15_Type"
                Selection.EntireColumn.Hidden = False
            Case "Operation15_Operator"
                Selection.EntireColumn.Hidden = False
        End Select
        ActiveCell.Offset(0, 1).Select

    Loop Until ActiveCell.Value = ""
End Function

Private Sub FormatWorksheetBorders()
    Cells.Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
End Sub

Private Sub DeleteSheet(SheetName As String)
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets(SheetName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub

Private Sub CreateSampleWIPFile(FilePath As String)
    Dim NewWB As Workbook
    Dim NewWS As Worksheet

    On Error GoTo CreateError

    ' Create new workbook
    Set NewWB = DataManager.CreateNewWorkbook()
    Set NewWS = NewWB.Worksheets(1)

    With NewWS
        .Name = "WIP_Data"

        ' Create headers (based on original fwip code column expectations)
        .Cells(1, 1).Value = "Date"           ' Column A (offset 0)
        .Cells(1, 2).Value = "Customer"       ' Column B (offset 1)
        .Cells(1, 3).Value = "Job"           ' Column C (offset 2)
        .Cells(1, 4).Value = "Job_Number"    ' Column D (offset 3)
        .Cells(1, 5).Value = "Quantity"      ' Column E (offset 4)
        .Cells(1, 6).Value = "Component_Code" ' Column F (offset 5)
        .Cells(1, 7).Value = "Component_Description" ' Column G (offset 6)
        .Cells(1, 8).Value = "Notes"         ' Column H (offset 7)
        .Cells(1, 9).Value = "Remarks"       ' Column I (offset 8)
        .Cells(1, 10).Value = "Status"       ' Column J (offset 9)
        .Cells(1, 11).Value = "Priority"     ' Column K (offset 10)
        .Cells(1, 12).Value = "StartDate"    ' Column L (offset 11)
        .Cells(1, 13).Value = "CustomerDelivery_Date" ' Column M (offset 12)
        .Cells(1, 14).Value = "Job_WorkshopDueDate"   ' Column N (offset 13)

        ' Add operation columns (starting at offset 14)
        Dim OpCol As Integer
        Dim i As Integer
        OpCol = 15 ' Column O
        For i = 1 To 15
            .Cells(1, OpCol).Value = "Operation" & Format(i, "00") & "_Type"
            .Cells(1, OpCol + 1).Value = "Operation" & Format(i, "00") & "_Operator"
            OpCol = OpCol + 2
        Next i

        ' Add sample data
        .Cells(3, 1).Value = Now - 5        ' Date
        .Cells(3, 2).Value = "SAMPLE CUSTOMER" ' Customer
        .Cells(3, 3).Value = "Sample Job"   ' Job
        .Cells(3, 4).Value = "J30001"       ' Job Number
        .Cells(3, 5).Value = "10"           ' Quantity
        .Cells(3, 6).Value = "COMP001"      ' Component Code
        .Cells(3, 7).Value = "Sample Component" ' Description
        .Cells(3, 8).Value = "Test notes"   ' Notes
        .Cells(3, 9).Value = "Test remarks" ' Remarks
        .Cells(3, 10).Value = "Active"      ' Status
        .Cells(3, 11).Value = "High"        ' Priority
        .Cells(3, 12).Value = Now - 3       ' Start Date
        .Cells(3, 13).Value = Now + 7       ' Customer Due Date
        .Cells(3, 14).Value = Now + 5       ' Workshop Due Date
        .Cells(3, 15).Value = "Machining"   ' Operation 1 Type
        .Cells(3, 16).Value = "John Doe"    ' Operation 1 Operator

        ' Add second sample row
        .Cells(4, 1).Value = Now - 3
        .Cells(4, 2).Value = "ANOTHER CUSTOMER"
        .Cells(4, 3).Value = "Another Job"
        .Cells(4, 4).Value = "J30002"
        .Cells(4, 5).Value = "5"
        .Cells(4, 6).Value = "COMP002"
        .Cells(4, 7).Value = "Another Component"
        .Cells(4, 8).Value = "More notes"
        .Cells(4, 9).Value = "More remarks"
        .Cells(4, 10).Value = "Active"
        .Cells(4, 11).Value = "Medium"
        .Cells(4, 12).Value = Now - 1
        .Cells(4, 13).Value = Now + 10
        .Cells(4, 14).Value = Now + 8
        .Cells(4, 15).Value = "Welding"
        .Cells(4, 16).Value = "Jane Smith"
        .Cells(4, 17).Value = "Finishing"
        .Cells(4, 18).Value = "Bob Wilson"

        ' Format headers
        .Range("A1:BB1").Font.Bold = True
        .Range("A1:BB1").Interior.Color = RGB(200, 200, 200)

        ' Auto-fit columns
        .Columns("A:BB").AutoFit

        ' Set BB1 value for column detection (this is what the original code looks for)
        .Cells(1, 54).Value = "END_MARKER" ' BB1 = column 54
    End With

    ' Save the file
    NewWB.SaveAs FilePath
    NewWB.Close
    Set NewWB = Nothing

    MsgBox "Sample WIP.xls file created successfully!" & vbCrLf & _
           "Location: " & FilePath & vbCrLf & vbCrLf & _
           "You can now modify this file with your actual WIP data.", vbInformation, "Sample File Created"
    Exit Sub

CreateError:
    If Not NewWB Is Nothing Then
        NewWB.Close SaveChanges:=False
        Set NewWB = Nothing
    End If
    MsgBox "Error creating sample WIP file: " & Err.Description, vbCritical
End Sub