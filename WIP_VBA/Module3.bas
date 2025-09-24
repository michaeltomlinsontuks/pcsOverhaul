Sub SaveToWIPAndSearch()
Dim InfoCol(1 To 100) As String
Dim InfoInfo(1 To 100) As String
Dim Quote_Number As String
Dim Job_Number As String
Dim Enq_Number As String
Dim File_Name As String

File_Name = ActiveWorkbook.Name

With Sheets("Admin")
    For i = 1 To 100
        InfoCol(i) = .Range("A1").Offset(i - 1, 0).Value
        InfoInfo(i) = .Range("A1").Offset(i - 1, 1).Value
        
        If .Range("A1").Offset(i - 1, 0).Value = "Quote_Number" Then Quote_Number = .Range("A1").Offset(i - 1, 1).Value
        If .Range("A1").Offset(i - 1, 0).Value = "Job_Number" Then Job_Number = .Range("A1").Offset(i - 1, 1).Value
        If .Range("A1").Offset(i - 1, 0).Value = "Enquiry_Number" Then Enq_Number = .Range("A1").Offset(i - 1, 1).Value
        If .Range("A1").Offset(i - 1, 0).Value = "File_Name" Then File_Name = .Range("A1").Offset(i - 1, 1).Value
        
    Next i
End With

If UCase(Right(ActiveWorkbook.Path, 3)) = "WIP" Then MasterPath = Left(ActiveWorkbook.Path, Len(ActiveWorkbook.Path) - 3)
If UCase(Right(ActiveWorkbook.Path, 9)) = "ENQUIRIES" Then MasterPath = Left(ActiveWorkbook.Path, Len(ActiveWorkbook.Path) - 9)
If UCase(Right(ActiveWorkbook.Path, 7)) = "ARCHIVE" Then MasterPath = Left(ActiveWorkbook.Path, Len(ActiveWorkbook.Path) - 7)
If UCase(Right(ActiveWorkbook.Path, 9)) = "CONTRACTS" Then MasterPath = Left(ActiveWorkbook.Path, Len(ActiveWorkbook.Path) - 9)

'    MsgBox (masterpath)
    Workbooks.Open MasterPath & "Search.xls"
    Range("A1").Select
    
    Do
        If ActiveCell.FormulaR1C1 = File_Name Then GoTo 5
        If ActiveCell.FormulaR1C1 = Enq_Number Then GoTo 5
        If ActiveCell.FormulaR1C1 = Quote_Number Then GoTo 5
        If ActiveCell.FormulaR1C1 = Job_Number Then GoTo 5
        ActiveCell.Offset(1, 0).Select
    Loop Until ActiveCell.FormulaR1C1 = ""
    
    MsgBox ("File not found")
    End

5:
    
Do

    If ActiveWorkbook.ReadOnly = True Then
        ActiveWorkbook.Close
        MsgBox ("This workbook is read only, please find the user with this workbook open and close it.")
        Workbooks.Open MasterPath & "Search.xls"
    End If

Loop Until ActiveWorkbook.ReadOnly = False

With Sheets("Search")
For j = 1 To 100
        For i = 0 To 100
            If UCase(.Range("A1").Offset(0, i).FormulaR1C1) = UCase(InfoCol(j)) Then
                ActiveCell.Offset(0, i).Value = UCase(InfoInfo(j))
                GoTo 6
            End If
            If UCase(.Range("a1").Offset(0, 1).FormulaR1C1) = "" Then GoTo 6
        Next i
6:
Next j
End With

ActiveWorkbook.Close True
Workbooks.Open MasterPath & "WIP.xls"

Do

    If ActiveWorkbook.ReadOnly = True Then
        ActiveWorkbook.Close
        MsgBox ("This workbook is read only, please find the user with this workbook open and close it.")
        Workbooks.Open MasterPath & "Wip.xls"
    End If

Loop Until ActiveWorkbook.ReadOnly = False

Range("A1").Select
    
    Do
        If ActiveCell.Offset(0, 2).FormulaR1C1 = InfoInfo(87) Then GoTo 8
        If ActiveCell.Offset(0, 2).FormulaR1C1 = InfoInfo(21) Then GoTo 8
        If ActiveCell.Offset(0, 2).FormulaR1C1 = InfoInfo(15) Then GoTo 8
        ActiveCell.Offset(1, 0).Select
    Loop Until ActiveCell.FormulaR1C1 = ""
    
    If MsgBox("File not found ~ Do you wish to add it to WIP", vbYesNo) = vbNo Then
        End
    End If

8:

With Sheets("Sheet1")
For j = 1 To 100
    For i = 0 To 100
        If UCase(.Range("A1").Offset(0, i).FormulaR1C1) = UCase(InfoCol(j)) Then
            If InStr(1, UCase(InfoCol(j)), "DATE", vbTextCompare) > 0 Then
                ActiveCell.Offset(0, i).Value = CDate(InfoInfo(j))
            Else
                ActiveCell.Offset(0, i).FormulaR1C1 = UCase(InfoInfo(j))
            End If
            GoTo 7
        End If
        If UCase(.Range("a1").Offset(0, 1).FormulaR1C1) = "" Then GoTo 6
    Next i
7:
Next j
End With
ActiveWorkbook.Close True
ActiveWorkbook.Close True
End Sub

Sub Close_fil()
ActiveWorkbook.Close True
End Sub




