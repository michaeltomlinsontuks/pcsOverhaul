VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FQuote 
   Caption         =   "MEM: Quote"
   ClientHeight    =   8715.001
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7605
   OleObjectBlob   =   "FQuote.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FQuote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub SaveQuote_Click()
On Error GoTo 10

TopCode:

If FQuote.Job_LeadTime.Value = "" Then GoTo 9
If FQuote.Component_Price.Value = "" Then GoTo 9
If Me.Component_Price = "" Then
    If MsgBox("Do you cancel the save in order to enter a Price?", vbYesNo, "MEM") = vbYes Then
        Exit Sub
    End If
End If

With Me
    .Quote_Number.Value = Calc_Next_Number("q")
    Confirm_Next_Number ("q")
    .File_Name.Value = .Quote_Number.Value
    xselect = Me.File_Name.Value
    'MsgBox ("The File Number for this Quote is: " & Me.File_Name.Value)
End With

x = OpenBook(Main.Main_MasterPath.Value & "enquiries\" & Me.Enquiry_Number.Value & ".xls", True)
        
    j = -1
    i = 1
    With Worksheets("ADMIN")
        For Each ctl In Me.Controls
            For i = 0 To 100
                If UCase(.Range("A1").Offset(i, 0).FormulaR1C1) = UCase(ctl.Name) And Left(.Range("A1").Offset(i, 0).Formula, 1) <> "=" Then
                    If UCase(TypeName(ctl)) = "TEXTBOX" Then .Range("A1").Offset(i, 1).FormulaR1C1 = UCase(ctl.Value)
                    If UCase(TypeName(ctl)) = "LABEL" Then .Range("A1").Offset(i, 1).FormulaR1C1 = UCase(ctl.Caption)
                    If UCase(TypeName(ctl)) = "COMBOBOX" Then .Range("A1").Offset(i, 1).FormulaR1C1 = UCase(ctl.Value)
                    GoTo 5
                End If
                If UCase(.Range("a1").Offset(i, 0).FormulaR1C1) = "" Then GoTo 5
            Next i
5:
        Next ctl
    End With
    
ActiveWorkbook.SaveAs (Main.Main_MasterPath.Value & "Quotes\" & Me.File_Name.Value & ".xls")
Kill (Main.Main_MasterPath.Value & "enquiries\" & Me.Enquiry_Number.Value & ".xls")
ActiveWorkbook.Close
               
'Save To Search
x = OpenBook(Main.Main_MasterPath & "Search.xls", False)
    Do
        If ActiveWorkbook.ReadOnly = True Then
            ActiveWorkbook.Close
            MsgBox ("This workbook is read only, please find the user with this workbook open and close it.")
            x = OpenBook(Main.Main_MasterPath & "Search.xls", False)
        End If
    Loop Until ActiveWorkbook.ReadOnly = False

Range("A1").Select

' Check Search Find for Office 2000
    Columns("A:A").Select
    Selection.Find(What:=Me.Enquiry_Number.Value, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate

'Do
'    ActiveCell.Offset(1, 0).Select
'Loop Until ActiveCell.FormulaR1C1 = "" Or _
'    ActiveCell.FormulaR1C1 = Me.Quote_Number.Value Or _
'    ActiveCell.FormulaR1C1 = Me.Enquiry_Number.Value Or _
'    ActiveCell.FormulaR1C1 = Me.File_Name.Value

With Sheets("search")
    For Each ctl In Me.Controls
        For i = 0 To 100
            If UCase(.Range("A1").Offset(0, i).FormulaR1C1) = UCase(ctl.Name) Then
                If TypeName(ctl) = "Label" Then .Range("A1").Offset(ActiveCell.Row - 1, i).FormulaR1C1 = UCase(ctl.Caption)
                If UCase(TypeName(ctl)) = "TEXTBOX" Then .Range("A1").Offset(ActiveCell.Row - 1, i).FormulaR1C1 = UCase(ctl.Value)
                If UCase(TypeName(ctl)) = "COMBOBOX" Then .Range("A1").Offset(ActiveCell.Row - 1, i).FormulaR1C1 = UCase(ctl.Value)
                GoTo 6
            End If
            If Left(.Range("A1").Offset(ActiveCell.Row - 2, i).FormulaR1C1, 1) = "=" Then .Range("A1").Offset(ActiveCell.Row - 1, i).FormulaR1C1 = .Range("A1").Offset(ActiveCell.Row - 2, i).FormulaR1C1
            
            If UCase(.Range("a1").Offset(0, 1).FormulaR1C1) = "" Then GoTo 6
        Next i
6:
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

Unload FQuote

Exit Sub
9:
MsgBox ("Please make sure that Lead Time and Price are filled in and only numbers")
Exit Sub

10:
MsgBox ("ERROR")
Resume

End Sub

Private Sub Search_Component_code_Click()
FSearch.Show
End Sub

Private Sub UserForm_Activate()
    
    If InStr(1, Main.lst.Value, "*") > 1 Then
        xselect = Left(Main.lst.Value, Len(Main.lst.Value) - 2)
    Else
        xselect = Main.lst.Value
    End If
    
    x = OpenBook(Main.Main_MasterPath.Value & "Enquiries\" & xselect & ".xls", True)
    
    With Sheets("Admin")
        For Each ctl In Me.Controls
            i = -1
            Do
                i = i + 1
                    If UCase(.Range("A1").Offset(i, 0).Value) = UCase(ctl.Name) Then
                            If InStr(1, ctl.Name, "Price", vbTextCompare) <> 0 Then
                                If UCase(TypeName(ctl)) = "LABEL" Then ctl.Caption = Insert_Characters(ctl.Name) & " : " & Format(.Range("A1").Offset(i, 1).Value, "R #,##0.00")
                                If UCase(TypeName(ctl)) = "COMBOBOX" Then ctl.Value = Format(.Range("A1").Offset(i, 1).Value, "R #,##0.00")
                                If UCase(TypeName(ctl)) = "TEXTBOX" Then ctl.Value = Format(.Range("A1").Offset(i, 1).Value, "R #,##0.00")
                                
                                GoTo FormLoadNext
                            End If
                            
                            If UCase(TypeName(ctl)) = "LABEL" Then ctl.Caption = Insert_Characters(ctl.Name) & " : " & .Range("A1").Offset(i, 1).Value
                            If UCase(TypeName(ctl)) = "COMBOBOX" Then ctl.Value = .Range("A1").Offset(i, 1).Value
                            If UCase(TypeName(ctl)) = "TEXTBOX" Then ctl.Value = .Range("A1").Offset(i, 1).Value
                        
                        GoTo FormLoadNext
                    End If
            Loop Until .Range("A1").Offset(i, 0).Value = ""
FormLoadNext:
        Next ctl
    End With
ActiveWorkbook.Close

With FQuote
    If .Quote_Date = "" Then
        .Quote_Date.Value = Format(CDate(Now()), "dd-mmm-yyyy")
    End If
    If .Job_LeadTime = "" Then
        .Job_LeadTime.Value = 14
    End If
End With
    Me.System_Status.Value = "New Quote"

End Sub

Public Function GetValue(path, File, sheet, ref)
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
End Function



