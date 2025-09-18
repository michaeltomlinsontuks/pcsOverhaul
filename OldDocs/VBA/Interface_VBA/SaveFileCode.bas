Attribute VB_Name = "SaveFileCode"
Public Function SaveToColumns()
' SaveColumnsToFile
j = -1
i = 1
With Worksheets("ADMIN")
    For Each ctl In Me.Controls
        For i = 0 To 100
                If UCase(.Range("A1").Offset(i, 0).FormulaR1C1) = UCase(ctl.Name) Then
                    If UCase(TypeName(ctl)) = "TEXTBOX" Then .Range("A1").Offset(i, 1).FormulaR1C1 = UCase(ctl.Value)
                    If UCase(TypeName(ctl)) = "LABEL" Then .Range("A1").Offset(i, 1).FormulaR1C1 = UCase(ctl.Caption)
                    If UCase(TypeName(ctl)) = "COMBOBOX" Then .Range("A1").Offset(i, 1).FormulaR1C1 = UCase(ctl.Value)
                    GoTo FormFileNext
                End If
                If UCase(.Range("a1").Offset(i, 0).FormulaR1C1) = "" Then GoTo 5
        Next i
FormFileNext:
    Next ctl
End With

If Me.Job_PicturePath.Value <> "" Then
    Range("Drawing_location").Select
    heit = Selection.RowHeight * 10
    ActiveSheet.Pictures.Insert(Main.Main_MasterPath.Value & "images\" & Me.Job_PicturePath.Value).Select
    With Selection
        .PrintObject = True
        .Name = "Drawing"
        .ShapeRange.Height = heit
        .Left = Range("drawing_location").Left + 5
        .Top = Range("drawing_location").Top + 5
    End With
End If

End Function
