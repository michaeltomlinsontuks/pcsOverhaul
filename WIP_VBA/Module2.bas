Sub SaveEditJobCard()
'
' Macro1 Macro
' Macro recorded 2007/08/06 by Jason Mogg
'
If ActiveSheet.Name = "Job Card" Then
    MsgBox ("Please click Edit Job Card First")
    End
End If

With Sheets("Job Card")
    Range("a1").Select
    
    For j = 0 To 26
        For k = 0 To 100
            If IsError(.Range(ActiveCell.Offset(k, j).Address).Value) = True Then
                GoTo 5
            Else
                If UCase(ActiveCell.Offset(k, j).Value) <> UCase(.Range(ActiveCell.Offset(k, j).Address).Value) Then
                    For i = 1 To 100
                        If InStr(1, .Range(ActiveCell.Offset(k, j).Address).FormulaR1C1, Sheets("Admin").Range("A1").Offset(i, 0).Value, vbTextCompare) Then
                            Sheets("Admin").Range("A1").Offset(i, 1).Value = UCase(ActiveCell.Offset(k, j).Value)
                            GoTo 5
                        End If
                    Next i
                End If
            End If
5:
            
        Next k
    Next j
End With

Application.DisplayAlerts = False
    ActiveSheet.Delete
Application.DisplayAlerts = True
Range("A1").Select
Sheets("Job Card").Visible = True
Sheets("Job Card").Select

Call SaveToWIPAndSearch

End Sub

Sub EditJobCard()
'
' Macro1 Macro
' Macro recorded 2007/08/06 by Jason Mogg

If ActiveSheet.Name = "Edit JC" Then
    MsgBox ("Please click Edit Job Card First")
    End
End If

    Sheets("Job Card").Select
    Sheets("Job Card").Copy After:=Sheets(3)
    Sheets("Job Card (2)").Select
    Sheets("Job Card (2)").Name = "Edit JC"
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Sheets("Job Card").Visible = False
    Range("A1").Select
End Sub

Sub CancelEditJC()

If ActiveSheet.Name = "Job Card" Then
    MsgBox ("Please click Edit Job Card First")
    End
End If

Application.DisplayAlerts = False
    ActiveSheet.Delete
Application.DisplayAlerts = True
Sheets("Job Card").Visible = True
Sheets("Job Card").Select
Range("A1").Select

End Sub

Sub AddPicture()
If MsgBox("Have you deleted the old picture?", vbYesNo) = vbNo Then
    MsgBox ("Please delete it first before trying again")
    End
End If
    
    Range("Drawing_location").Select
    heit = Selection.RowHeight * 10

    Application.Dialogs(xlDialogInsertPicture).Show

   With Selection
        .PrintObject = True
        .Name = "Drawing"
        .ShapeRange.Height = heit
        .Left = Range("drawing_location").Left + 5
        .Top = Range("drawing_location").Top + 5
    End With
    
    Sheets("ADmin").Range("B22").Value = ""

End Sub

