# VBA code extracted from Due Date.xls
# Extraction date: Mon Jun  2 11:09:13 SAST 2025

# =======================================================
# Module 7
# =======================================================
Attribute VB_Name = "Delete_Sheet"
' Delete_Sheet
' Deletes a sheet without prompting to confirm delete

Option Explicit
Public Function DeleteSheet(SheetName As String)

    Application.DisplayAlerts = False
    Worksheets(SheetName).Delete
    Application.DisplayAlerts = True

End Function


# =======================================================
# Module 8
# =======================================================
Attribute VB_Name = "Module1"
Private Type Jobs
    Dat As Date
    Cust As String
    Job As String
    Qty As String
    Desc As String
    Remarks As String
    DDat As String
    
    OPs(1 To 15) As String
End Type

Sub asd()
Dim TempSheet As String
Dim Job(1 To 5000) As Jobs

Range("A1:z5000").Select
    Selection.Sort Key1:=Range("G2"), Order1:=xlAscending, Header:=xlYes, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal
    
Range("A2").Select
i = 0

Do
    i = i + 1
    With Job(i)
        .Dat = ActiveCell.Value
        .Cust = ActiveCell.Offset(0, 1).FormulaR1C1
        .Job = ActiveCell.Offset(0, 2).FormulaR1C1
        .Qty = ActiveCell.Offset(0, 3).FormulaR1C1
        .Desc = ActiveCell.Offset(0, 4).FormulaR1C1
        .Remarks = ActiveCell.Offset(0, 5).FormulaR1C1
        .DDat = ActiveCell.Offset(0, 6).FormulaR1C1
        
        For j = 1 To 15
            .OPs(j) = ActiveCell.Offset(0, 7 + j).FormulaR1C1
        Next j
    End With
    ActiveCell.Offset(1, 0).Select

Loop Until ActiveCell.FormulaR1C1 = ""

Workbooks.Add
'Operation

For j = 1 To i
    With Job(j)
        For k = 1 To 15
            If .OPs(k) <> "" Then
                TempSheet = "OPERATION - " & Left(.OPs(k), InStr(1, .OPs(k), " : ", vbTextCompare) - 1)
                On Error GoTo AddSheet
                Sheets(TempSheet).Select
            
                ActiveCell.FormulaR1C1 = .Dat
                ActiveCell.Offset(0, 1).FormulaR1C1 = .Cust
                ActiveCell.Offset(0, 2).FormulaR1C1 = .Job
                ActiveCell.Offset(0, 3).FormulaR1C1 = .Qty
                ActiveCell.Offset(0, 4).FormulaR1C1 = .Desc
                ActiveCell.Offset(0, 5).FormulaR1C1 = .Remarks
                ActiveCell.Offset(0, 6).FormulaR1C1 = .DDat
                
                If k > 1 Then
                    If .OPs(k - 1) = "" Then
                        ActiveCell.Offset(0, 7).FormulaR1C1 = "*"
                        Selection.EntireRow.Font.Bold = True
                    End If
                End If
                        
                ActiveCell.Offset(1, 0).Select
            End If
            TempSheet = ""
        Next k
    End With
Next j

For Each sh In Sheets
    sh.Select
    Cells.EntireColumn.AutoFit
    Range("A1:H5000").Select
    Selection.Sort Key1:=Range("H2"), Order1:=xlAscending, Header:=xlYes, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal
        
    With ActiveSheet.PageSetup
    
        .CenterHeader = ActiveSheet.Name
        .RightHeader = "&D &T"
        
    End With

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

    Range("A1").Select
Next sh

DeleteSheet ("sheet1")
DeleteSheet ("sheet2")
DeleteSheet ("sheet3")

Workbooks.Add
'Operator

For j = 1 To i
    With Job(j)
        For k = 1 To 15
            If .OPs(k) <> "" Then
                TempSheet = "OPERATOR - " & Mid(.OPs(k), InStr(1, .OPs(k), " : ", vbTextCompare) + 3, Len(.OPs(k)) - InStr(1, .OPs(k), " : ", vbTextCompare) + 3)
                On Error GoTo AddSheet
                Sheets(TempSheet).Select
            
                ActiveCell.FormulaR1C1 = .Dat
                ActiveCell.Offset(0, 1).FormulaR1C1 = .Cust
                ActiveCell.Offset(0, 2).FormulaR1C1 = .Job
                ActiveCell.Offset(0, 3).FormulaR1C1 = .Qty
                ActiveCell.Offset(0, 4).FormulaR1C1 = .Desc
                ActiveCell.Offset(0, 5).FormulaR1C1 = .Remarks
                ActiveCell.Offset(0, 6).FormulaR1C1 = .DDat
                
                If k > 1 Then
                    If .OPs(k - 1) = "" Then
                        ActiveCell.Offset(0, 7).FormulaR1C1 = "*"
                        Selection.EntireRow.Font.Bold = True
                    End If
                End If
                        
                ActiveCell.Offset(1, 0).Select
            End If
            TempSheet = ""
        Next k
    End With
Next j

For Each sh In Sheets
    sh.Select
    Cells.EntireColumn.AutoFit
    Range("A1:H5000").Select
    Selection.Sort Key1:=Range("H2"), Order1:=xlAscending, Header:=xlYes, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal
    
    With ActiveSheet.PageSetup
    
        .CenterHeader = ActiveSheet.Name
        .RightHeader = "&D &T"
        
    End With

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

    Range("A1").Select
Next sh

DeleteSheet ("sheet1")
DeleteSheet ("sheet2")
DeleteSheet ("sheet3")

Exit Sub
AddSheet:
    Sheets.Add
    ActiveSheet.Name = TempSheet
    ActiveCell.FormulaR1C1 = "DATE"
    ActiveCell.Offset(0, 1).FormulaR1C1 = "CUSTOMER"
    ActiveCell.Offset(0, 2).FormulaR1C1 = "JOB"
    ActiveCell.Offset(0, 3).FormulaR1C1 = "QTY"
    ActiveCell.Offset(0, 4).FormulaR1C1 = "DESCRIPTION"
    ActiveCell.Offset(0, 5).FormulaR1C1 = "REMARKS"
    ActiveCell.Offset(0, 6).FormulaR1C1 = "DUE DATE"
    Columns("G:G").NumberFormat = "dd mmm"
    Columns("A:A").NumberFormat = "dd mmm"
    Selection.EntireRow.Font.Bold = True
    ActiveCell.Offset(1, 0).Select
    
Resume

End Sub
    


# =======================================================
# Module 9
# =======================================================
Attribute VB_Name = "Module2"
Sub WriteCodes()
'
' Macro6 Macro
' Macro recorded 2006/10/11 by Jason
'
Dim MasterPath As String
Dim Typ(1 To 20) As String
Dim Seq(1 To 20) As String
Dim Comments(1 To 20) As String
Dim OP(1 To 20) As String

'MasterPath = Sheets("Enquiry").Range("e1").FormulaR1C1
Range("c2").Select

Do
    MasterPath = ActiveWorkbook.path
        
    If Dir(MasterPath & "\Archive\" & ActiveCell.Value & ".xls", vbNormal) <> "" Then
        MasterPath = ActiveWorkbook.path & "\Archive\"
    End If
    If Dir(MasterPath & "\Enquiries\" & ActiveCell.FormulaR1C1 & ".xls", vbNormal) <> "" Then
        MasterPath = ActiveWorkbook.path & "\Enquiries\"
    End If
    If Dir(MasterPath & "\WIP\" & ActiveCell.FormulaR1C1 & ".xls", vbNormal) <> "" Then
        MasterPath = ActiveWorkbook.path & "\wip\"
    End If
    If Dir(MasterPath & "\Quotes\" & ActiveCell.FormulaR1C1 & ".xls", vbNormal) <> "" Then
        MasterPath = ActiveWorkbook.path & "\Quotes\"
    End If
    
    ActiveCell.Offset(0, 2).FormulaR1C1 = GetValue(MasterPath, ActiveCell.FormulaR1C1 & ".xls", "Enquiry", "b5")
    ActiveCell.Offset(0, 3).FormulaR1C1 = GetValue(MasterPath, ActiveCell.FormulaR1C1 & ".xls", "Enquiry", "b7")
    
    ActiveCell.Offset(1, 0).Select
Loop Until ActiveCell.FormulaR1C1 = ""

ActiveWorkbook.Save

End Sub

Function GetValue(path, File, sheet, ref)
'   Retrieves a value from a closed workbook
    Dim arg As String

'   Make sure the file exists
    If Right(path, 1) <> "\" Then path = path & "\"
    If Dir(path & File) = "" Then
        GetValue = "File Not Found"
        Exit Function
    End If

'   Create the argument
    arg = "'" & path & "[" & File & "]" & sheet & "'!" & _
      Range(ref).Range("A1").Address(, , xlR1C1)

'   Execute an XLM macro
    GetValue = ExecuteExcel4Macro(arg)
    
    If GetValue = 0 Then
        GetValue = ""
    End If
    
End Function







# =======================================================
# Module 11
# =======================================================
Attribute VB_Name = "Sorts"
Sub Sort_JobNumber()

Range("bb1").Select
    Selection.End(xlToLeft).Select
    col = ActiveCell.Column

Range("A1").Select
    Selection.End(xlDown).Select
   
    Range("A2", Range("A2").Offset(0, col - 1).Address).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Sort Key1:=Range("h3"), Order1:=xlAscending, Header:=xlYes, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
        
Range("a3").Select
    Do
        ActiveCell.Value = CDate(ActiveCell.Value)
        ActiveCell.Offset(1, 0).Select
    Loop Until ActiveCell.Value = ""
    
Range("H4").Select
    Do
        ActiveCell.Value = CDate(ActiveCell.Value)
        ActiveCell.Offset(1, 0).Select
    Loop Until ActiveCell.Value = ""
    
Range("J4").Select
    Do
        ActiveCell.Value = CCur(ActiveCell.Value)
        ActiveCell.Offset(1, 0).Select
    Loop Until ActiveCell.Value = ""

Range("k4").Select '
    Do
        ActiveCell.Value = CDate(ActiveCell.Value)
        ActiveCell.Offset(1, 0).Select
    Loop Until ActiveCell.Value = ""

Range("A2").Select

End Sub


