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
    Dim sh As Worksheet

    On Error GoTo Error_Handler

    Application.DisplayAlerts = False

    ' Open WIP.xls using new module structure
    Dim WIPWB As Workbook
    Set WIPWB = DataManager.SafeOpenWorkbook(DataManager.GetRootPath & "\WIP.xls")
    If WIPWB Is Nothing Then
        MsgBox "Unable to open WIP.xls", vbCritical
        Exit Sub
    End If

    ' Load WIP data into Jobs array (simplified version of original)
    WIPWB.Activate
    Range("A1").Select
    Selection.End(xlDown).Select
    i = 0

    Range("A2").Select
    Do While ActiveCell.Value <> ""
        i = i + 1
        With Job(i)
            .Dat = ActiveCell.Offset(0, 7).Value  ' DateCreated
            .Cust = ActiveCell.Offset(0, 1).Value ' CustomerName
            .Job = ActiveCell.Offset(0, 0).Value  ' JobNumber
            .JobD = Val(Replace(ActiveCell.Offset(0, 0).Value, "J", "")) ' Job number as double
            .Qty = ActiveCell.Offset(0, 3).Value  ' Quantity
            .Cod = ""  ' Component code not in simple structure
            .Desc = ActiveCell.Offset(0, 2).Value ' ComponentDescription
            .Remarks = "" ' Not in simple structure
            .DDat = ActiveCell.Offset(0, 4).Value ' DueDate

            ' Simplified operator handling - just use assigned operator
            .OperatorType(1) = "PRODUCTION"
            .OperatorN(1) = ActiveCell.Offset(0, 5).Value ' AssignedOperator
        End With
        ActiveCell.Offset(1, 0).Select
        If i >= 5000 Then Exit Do
    Loop

    DataManager.SafeCloseWorkbook WIPWB, False
    fwip.Hide

    ' Generate Operation Reports (original functionality)
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

                        ActiveCell.Offset(1, 0).Select
                    End If
                    TempSheet = ""
SkipOPP:
                Next k
            End With
        Next j

        ' Format all sheets (original formatting)
        For Each sh In Sheets
            sh.Select
            Cells.EntireColumn.AutoFit
            Range("A1:I5000").Select
            Selection.Sort Key1:=Range("H2"), Order1:=xlAscending, Key2:=Range("G2"), _
                Order2:=xlAscending, Header:=xlYes, OrderCustom:=1, MatchCase:=False, _
                Orientation:=xlTopToBottom

            With ActiveSheet.PageSetup
                .CenterHeader = ActiveSheet.Name
                .RightHeader = "&D &T"
            End With

            ' Add borders (original formatting)
            Cells.Select
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
        Next sh

        ' Save the report using new module structure
        Dim SavePath As String
        SavePath = DataManager.GetRootPath & "\Templates\Operation_Report.xls"
        ActiveWorkbook.SaveAs SavePath
    End If

    Application.DisplayAlerts = True
    Exit Sub

AddSheet:
    Sheets.Add
    ActiveSheet.Name = CoreFramework.RemoveInvalidCharacters(Trim(TempSheet))
    Range("A1").Select
    Resume Next

Error_Handler:
    Application.DisplayAlerts = True
    CoreFramework.HandleStandardErrors Err.Number, "Go_Click", "fwip"
End Sub