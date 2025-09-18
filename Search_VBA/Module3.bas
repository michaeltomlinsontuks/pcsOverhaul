Sub Textify()

Range("N1").Select

Do
    ActiveCell.FormulaR1C1 = CStr(ActiveCell.Value)
    ActiveCell.Offset(1, 0).Select
Loop Until Range("A" & ActiveCell.Row).Value = ""

End Sub
••••ˇˇˇˇ