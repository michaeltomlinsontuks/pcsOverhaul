# VBA code extracted from 384776.xls
# Extraction date: Mon Jun  2 11:09:20 SAST 2025

# =======================================================
# Module 7
# =======================================================
Attribute VB_Name = "Module1"
Sub CreateReferenceNames()

Sheets("Admin").Select
Range("b2").Select

Do
    ActiveWorkbook.Names.Add Name:=ActiveCell.Offset(0, -1).FormulaR1C1, RefersToR1C1:= _
        "=Admin!R" & ActiveCell.Row & "C2"
    ActiveCell.Offset(1, 0).Select
Loop Until ActiveCell.Offset(0, -1).FormulaR1C1 = ""

End Sub


# =======================================================
# Module 8
# =======================================================
Attribute VB_Name = "Module2"
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



# =======================================================
# Module 9
# =======================================================
Attribute VB_Name = "Module3"
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






# =======================================================
# Module 10
# =======================================================
Attribute VB_Name = "SaveSearch"
Sub SaveToSearchOnly()
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
If UCase(Right(ActiveWorkbook.Path, 6)) = "QUOTES" Then MasterPath = Left(ActiveWorkbook.Path, Len(ActiveWorkbook.Path) - 6)

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
ActiveWorkbook.Close True

End Sub



# =======================================================
# Module 15
# =======================================================
Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Workbook_BeforePrint(Cancel As Boolean)

If Cancel = False Then
    With Worksheets("Admin")
        For i = 1 To 100
        
            If UCase(.Range("a1").Offset(i, 0).FormulaR1C1) = UCase("Job_CardPrinted") Then .Range("a1").Offset(i, 1).FormulaR1C1 = Now()
    
        Next i
    
    End With
    ActiveWorkbook.Save
End If

End Sub



