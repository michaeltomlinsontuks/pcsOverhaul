Attribute VB_Name = "SaveSearchCode"
Sub SaveRowIntoSearch(frm As Object)

'Save To Search
OpenBook (Main.Main_MasterPath & "Search.xls")
    Do
        If ActiveWorkbook.ReadOnly = True Then
            ActiveWorkbook.Close
            MsgBox ("This workbook is read only, please find the user with this workbook open and close it.")
            OpenBook (Main.Main_MasterPath & "Search.xls")
        End If
    Loop Until ActiveWorkbook.ReadOnly = False

Range("A1").Select

Do
    ActiveCell.Offset(1, 0).Select
Loop Until ActiveCell.FormulaR1C1 = "" Or _
    ActiveCell.FormulaR1C1 = Me.Quote_Number.Value Or _
    ActiveCell.FormulaR1C1 = Me.Enquiry_Number.Value Or _
    ActiveCell.FormulaR1C1 = Me.Job_Number.Value Or _
    ActiveCell.FormulaR1C1 = Me.File_Name.Value

With Sheets("search")
    For Each ctl In Me.Controls
        For i = 0 To 100
            If UCase(.Range("A1").Offset(0, i).FormulaR1C1) = UCase(ctl.Name) Then
                If TypeName(ctl) = "Label" Then .Range("A1").Offset(ActiveCell.Row - 1, i).FormulaR1C1 = UCase(ctl.Caption)
                If UCase(TypeName(ctl)) = "TEXTBOX" Then .Range("A1").Offset(ActiveCell.Row - 1, i).FormulaR1C1 = UCase(ctl.Value)
                If UCase(TypeName(ctl)) = "COMBOBOX" Then .Range("A1").Offset(ActiveCell.Row - 1, i).FormulaR1C1 = UCase(ctl.Value)
                GoTo FormNextSearch
            End If
            If Left(.Range("A1").Offset(ActiveCell.Row - 2, i).FormulaR1C1, 1) = "=" Then .Range("A1").Offset(ActiveCell.Row - 1, i).FormulaR1C1 = .Range("A1").Offset(ActiveCell.Row - 2, i).FormulaR1C1
            If UCase(.Range("a1").Offset(0, 1).FormulaR1C1) = "" Then GoTo FormNextSearch
        Next i
FormNextSearch:
    Next ctl
End With

    Range("A1").Select
        Selection.End(xlToRight).Select
        col = ActiveCell.Column
    Range("A1").Select
    Selection.End(xlDown).Select
    Range("A2", Range("A2").Offset(ActiveCell.Row, col - 1).Address).Select
    Selection.Sort Key1:=Range("e2"), Order1:=xlDescending, Header:=xlYes, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortTextAsNumbers
    Range("b3").Select
        
ActiveWorkbook.Close (True)

End Sub


