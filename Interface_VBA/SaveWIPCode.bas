Attribute VB_Name = "SaveWIPCode"
Sub SaveInfoIntoWIP(frm As Object)

' Save to WIP
OpenBook (Main.Main_MasterPath & "WIP.xls")
    Do
        If ActiveWorkbook.ReadOnly = True Then
            ActiveWorkbook.Close
            MsgBox ("This workbook is read only, please find the user with this workbook open and close it.")
            OpenBook (Main.Main_MasterPath & "WIP.xls")
        End If
    Loop Until ActiveWorkbook.ReadOnly = False

    Range("A1").Select
    
    Do
        ActiveCell.Offset(1, 0).Select
    Loop Until ActiveCell.Offset(0, 2).FormulaR1C1 = "" Or _
        ActiveCell.Offset(0, 2).FormulaR1C1 = Me.Quote_Nmber.Value Or _
        ActiveCell.Offset(0, 2).FormulaR1C1 = Me.Enquiry_Number.Value Or _
        ActiveCell.Offset(0, 2).FormulaR1C1 = Me.Job_Number.Value Or _
        ActiveCell.Offset(0, 2).FormulaR1C1 = Me.File_Name.Value
    
    Selection.EntireRow.ClearContents

    With Sheets(ActiveSheet.Name)
        For Each ctl In Me.Controls
            For i = 0 To 100
                    If UCase(.Range("A1").Offset(0, i).FormulaR1C1) = UCase(ctl.Name) Then
                        If TypeName(ctl) = "Label" Then .Range("A1").Offset(ActiveCell.Row - 1, i).FormulaR1C1 = UCase(ctl.Caption)
                        If UCase(TypeName(ctl)) = "TEXTBOX" Then .Range("A1").Offset(ActiveCell.Row - 1, i).FormulaR1C1 = UCase(ctl.Value)
                        If UCase(TypeName(ctl)) = "COMBOBOX" Then .Range("A1").Offset(ActiveCell.Row - 1, i).FormulaR1C1 = UCase(ctl.Value)
                        GoTo FormNextWIP
                    End If
                    If Left(.Range("A1").Offset(ActiveCell.Row - 2, i).FormulaR1C1, 1) = "=" Then .Range("A1").Offset(ActiveCell.Row - 1, i).FormulaR1C1 = .Range("A1").Offset(ActiveCell.Row - 2, i).FormulaR1C1
                    If UCase(.Range("a1").Offset(0, 1).FormulaR1C1) = "" Then GoTo 6
            Next i
FormNextWIP:
        Next ctl
    End With
    
End Sub



