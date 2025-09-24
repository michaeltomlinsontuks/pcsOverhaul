Sub CreateReferenceNames()

Sheets("Admin").Select
Range("b2").Select

Do
    ActiveWorkbook.Names.Add Name:=ActiveCell.Offset(0, -1).FormulaR1C1, RefersToR1C1:= _
        "=Admin!R" & ActiveCell.Row & "C2"
    ActiveCell.Offset(1, 0).Select
Loop Until ActiveCell.Offset(0, -1).FormulaR1C1 = ""

End Sub
