VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Main 
   Caption         =   "Main"
   ClientHeight    =   11040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14625
   OleObjectBlob   =   "Main.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub Add_Enquiry_Click()

With FrmEnquiry
    .Enquiry_Date.Caption = Format(Now(), "dd mmm yyyy")
'    .Dat = "Please click here to insert a date"
    .Component_Code = ""
    .Component_Description = ""
    .Customer = ""
    .Component_Grade = ""
    .Component_Quantity = ""
    .Show
End With

    If Main.WIP.Value = True Then
        Main.WIP.Value = False
        Main.WIP.Value = True
    End If
    
    If Main.Enquiries.Value = True Then
        Main.Enquiries.Value = False
        Main.Enquiries.Value = True
    End If
    
    If Main.Archive.Value = True Then
        Main.Archive.Value = False
        Main.Archive.Value = True
    End If
    
    If Main.Quotes.Value = True Then
        Main.Quotes.Value = False
        Main.Quotes.Value = True
    End If
    
End Sub

Private Sub Archive_Click()

If Main.Archive.Value = True Then
    Main.lst.Clear

    x = List_Files("Archive", Main.lst)
    Main.Enquiries.Value = False
    Main.WIP.Value = False
    Main.Thirties.Value = False
    Main.Quotes.Value = False
    Main.JobsInWIP.Value = False
    
'    NextCheck = Now + TimeValue("00:00:05")
'    Application.OnTime NextCheck, "CheckUpdates"

End If

End Sub

Private Sub but_CreateCTItem_Click()

x = OpenBook(Main.Main_MasterPath & "Templates\" & "_Enq.xls", True)
FJG.but_SaveAsCTItem.Visible = True
FJG.butSaveJG.Visible = False
FJG.Show
ActiveWorkbook.Close True

End Sub

Private Sub but_EditCTItem_Click()
Dim FullFilePath As String, MyName As String
Dim GroupCount As Integer
Dim Fil(1 To 10000) As String
'\* Check a Group folder exists
'FullFilePath = "C:\TEMP\Group*"

Dim Typ(1 To 20) As String
Dim Seq(1 To 20) As String
Dim Comments(1 To 20) As String
Dim OP(1 To 20) As String

Main.lst.ListIndex = -1

'fileextension = right(ucase(InputBox("Please enter file extension")),3)
fileextension = "*.*"

FList.lst.Clear

1:

'ChDrive (main.Main_MasterPath & path)

MyName = Dir(Main.Main_MasterPath.Value & "Contracts" & "\", vbDirectory)
    If MyName = "" Then
        MsgBox "Folder Not Found", vbOKOnly, "Test"
            Exit Sub
    End If
'\* Store list of Group folder names

i = 0
Do Until MyName = ""

    If MyName = "." Or MyName = ".." Then GoTo 2

    i = i + 1
    Fil(i) = MyName
    
    FList.lst.AddItem Left(MyName, Len(MyName) - 4)
2:
    
    GroupCount = GroupCount + 1
    
    MyName = Dir
    
Loop

FList.Show

Dim Missed(1 To 100) As Integer

xselect = FList.lst.Value

x = OpenBook(Main.Main_MasterPath.Value & "Contracts\" & xselect & ".xls", False)
Windows(xselect & ".xls").Activate

Unload Me

End Sub

Private Sub But_EditJC_Click()
If InStr(1, Main.lst.Value, "*") > 1 Then
    xselect = Left(Main.lst.Value, Len(Main.lst.Value) - 2)
Else
    xselect = Main.lst.Value
End If

If Dir(Main.Main_MasterPath.Value & "enquiries\" & xselect & ".xls", vbNormal) <> "" Then x = OpenBook(Main.Main_MasterPath & "enquiries\" & xselect & ".xls", False)
If Dir(Main.Main_MasterPath.Value & "archive\" & xselect & ".xls", vbNormal) <> "" Then x = OpenBook(Main.Main_MasterPath & "archive\" & xselect & ".xls", False)
If Dir(Main.Main_MasterPath.Value & "wip\" & xselect & ".xls", vbNormal) <> "" Then x = OpenBook(Main.Main_MasterPath & "WIP\" & xselect & ".xls", False)
If Dir(Main.Main_MasterPath.Value & "QUOTES\" & xselect & ".xls", vbNormal) <> "" Then x = OpenBook(Main.Main_MasterPath & "QUOTES\" & xselect & ".xls", False)

Unload Main


End Sub

Private Sub butEditSearch_Click()

x = OpenBook(Main.Main_MasterPath & "Search.xls", False)

    Range("A1").Select
        Selection.End(xlToRight).Select
        col = ActiveCell.Column
    Range("A1").Select
    Selection.End(xlDown).Select
    Range("A2", Range("A2").Offset(ActiveCell.Row, col - 1).Address).Select
    
    Selection.Sort Key1:=Range("e2"), Order1:=xlDescending, Header:=xlYes, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortTextAsNumbers

Range("C3").Select
Unload Main

End Sub

Private Sub butSearchHistory_Click()
Workbooks.Open Main.Main_MasterPath & "search History.xls", ReadOnly:=True
Range("b1").Select
Main.Hide
Application.Run "'Search History.xls'!Show_Search_Menu"
Unload Main
Exit Sub

FSearch1.Show
End Sub

Private Sub butJobHistory_Click()
    Workbooks.Open Main.Main_MasterPath & "Job History.xls", ReadOnly:=True
    Range("b1").Select
    Main.Hide
    Application.Run "'Job History.xls'!Show_Search_Menu"
    Unload Main
End Sub

Private Sub butQuoteHistory_Click()
    Workbooks.Open Main.Main_MasterPath & "Quote History.xls", ReadOnly:=True
    Range("b1").Select
    Main.Hide
    Application.Run "'Quote History.xls'!Show_Search_Menu"
    Unload Main
End Sub

Private Sub butShowContractsFolder_Click()
Shell "C:\WINDOWS\explorer.exe """ & ActiveWorkbook.path & "\Contracts" & """", vbNormalFocus
End
End Sub

Private Sub butSortSearch_Click()
Workbooks.Open Main.Main_MasterPath & "search.xls", ReadOnly:=False

        Range("A1").Select
        Selection.End(xlToRight).Select
        col = ActiveCell.Column
        Range("A2").Select
        Selection.End(xlDown).Select
        'Main.Hide
        Range("A2", Range("A2").Offset(ActiveCell.Row, col - 1).Address).Select
        Selection.Sort Key1:=Range("e2"), Order1:=xlAscending, Header:=xlYes, _
            OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
            
        Range("b3").Select
ActiveWorkbook.Close True
            
End Sub

Private Sub CalledThrough_Click()

If InStr(1, Main.lst.Value, "*") > 1 Then
    xselect = Left(Main.lst.Value, Len(Main.lst.Value) - 2)
Else
    xselect = Main.lst.Value
End If

Me.File_Name.Value = xselect

If Dir(Main.Main_MasterPath.Value & "quotes\" & xselect & ".xls") <> "" Then
    
    x = OpenBook(Main.Main_MasterPath.Value & "quotes\" & xselect & ".xls", False)
        Sheets("admin").Select
        Me.System_Status.Value = "QUOTE SUBMITTED"
        Range("system_Status").FormulaR1C1 = Me.System_Status.Value
    ActiveWorkbook.SaveAs (Main.Main_MasterPath & "Archive\" & xselect & ".xls")
    ActiveWorkbook.Close
    Kill (Main.Main_MasterPath & "quotes\" & xselect & ".xls")

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
    
    Main.lst.Clear
    
    If Main.Quotes.Value = True Then
        x = List_Files("quotes", Main.lst)
    End If

End If

End Sub

Private Sub CloseJob_Click()
Dim JC As String

If Main.lst.ListIndex < 0 Then
    MsgBox ("Please select a job")
    Exit Sub
End If

If MsgBox("Do you wish to Close this Job (" & Main.lst.Value & ")?", vbYesNo) = vbNo Then Exit Sub

Me.Invoice_Number.Value = InputBox("Please enter an invoice number?")
Me.Invoice_Date.Value = Format(Now(), "dd mmm yyyy")
Me.System_Status.Value = UCase("Job Closed")
    
If Me.Invoice_Number.Value = "" Then
    MsgBox ("You must enter an invoice number to close this job")
    ActiveWorkbook.Close (False)
    Exit Sub
End If
    
    
If Dir(Main.Main_MasterPath & "archive\" & Main.lst.Value & ".xls") <> "" Then
    
    x = OpenBook(Main.Main_MasterPath & "archive\" & Main.lst.Value & ".xls", False)
    
    Sheets("Job Card").Select
    
    If Range("Invoice_Number").FormulaR1C1 <> "" Then
        MsgBox ("ERROR - Invoice Already Exists")
        End
    End If
    
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
                    If UCase(.Range("a1").Offset(i, 0).FormulaR1C1) = "" Then GoTo FormFileNext
            Next i
FormFileNext:
        Next ctl
    End With
    
    ActiveWorkbook.Close True
                
    x = OpenBook(Main.Main_MasterPath & "WIP.xls", False)

        Do
        
            If ActiveWorkbook.ReadOnly = True Then
                ActiveWorkbook.Close
                MsgBox ("This workbook is read only, please find the user with this workbook open and close it.")
                x = OpenBook(Main.Main_MasterPath & "WIP.xls", False)
            End If
        
        Loop Until ActiveWorkbook.ReadOnly = False
    
        Range("A1").Select
        
        Do
            ActiveCell.Offset(1, 0).Select
        Loop Until ActiveCell.Offset(0, 2).FormulaR1C1 = Me.Job_Number.Value Or ActiveCell.FormulaR1C1 = ""
    
        If ActiveCell.FormulaR1C1 = "" Then
            MsgBox ("An error has occured as it does not appear in the WIP File")
            End
        End If
        
        Selection.EntireRow.Delete

    ActiveWorkbook.Close (True)
    
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

    Refresh_Main

Else

    MsgBox ("This Job is not open and therefore can not be closed")
    
End If

If Main.JobsInWIP.Value = True Then
    Main.JobsInWIP.Value = False
    Main.JobsInWIP.Value = True
End If

End Sub

Private Sub ContractWork_Click()
Dim FullFilePath As String, MyName As String
Dim GroupCount As Integer
Dim Fil(1 To 10000) As String
'\* Check a Group folder exists
'FullFilePath = "C:\TEMP\Group*"

Dim Typ(1 To 20) As String
Dim Seq(1 To 20) As String
Dim Comments(1 To 20) As String
Dim OP(1 To 20) As String

Main.lst.ListIndex = -1

'fileextension = right(ucase(InputBox("Please enter file extension")),3)
fileextension = "*.*"

FList.lst.Clear

1:

'ChDrive (main.Main_MasterPath & path)

MyName = Dir(Main.Main_MasterPath.Value & "Contracts" & "\", vbDirectory)
    If MyName = "" Then
        MsgBox "Folder Not Found", vbOKOnly, "Test"
            Exit Sub
    End If
'\* Store list of Group folder names

i = 0
Do Until MyName = ""

    If MyName = "." Or MyName = ".." Then GoTo 2

    i = i + 1
    Fil(i) = MyName
    
    FList.lst.AddItem Left(MyName, Len(MyName) - 4)
2:
    
    GroupCount = GroupCount + 1
    
    MyName = Dir
    
Loop

FList.Show

Dim Missed(1 To 100) As Integer

xselect = FList.lst.Value
x = OpenBook(Main.Main_MasterPath.Value & "Contracts\" & xselect & ".xls", True)
Windows(xselect & ".xls").Activate

FJG.but_SaveAsCTItem.Visible = False
FJG.butSaveJG.Visible = True
FJG.Component_Quantity.SetFocus
FJG.Show
                    
Sheets("Job Card").Select
Range("A1").Select
Range("r3").FormulaR1C1 = ""

ActiveWorkbook.SaveAs Main.Main_MasterPath.Value & "wip\" & FJG.File_Name.Value & ".xls"
Main.File_Name.Value = FJG.File_Name.Value

' Add To Search
ActiveWorkbook.Close True
ActiveWorkbook.Close True

Unload FJG
Unload FAcceptQuote
Unload FList
'Unload Main

End Sub

Private Sub Enquiries_Click()

If Main.Enquiries.Value = True Then
    
    Main.lst.Clear

    x = List_Files("Enquiries", Main.lst)
    Main.Notice_Enquiries.Caption = "Enquiries : " & Check_Files(Main.Main_MasterPath & "enquiries\")
    Main.Quotes.Value = False
    Main.WIP.Value = False
    Main.Archive.Value = False
    Main.JobsInWIP.Value = False
    Main.Thirties.Value = False
    
'    NextCheck = Now + TimeValue("00:00:05")
'    Application.OnTime NextCheck, "CheckUpdates"

End If

End Sub

Private Sub FPrint_Click()

If InStr(1, Main.lst.Value, "*") > 1 Then
    xselect = Left(Main.lst.Value, Len(Main.lst.Value) - 2)
Else
    xselect = Main.lst.Value
End If

If Dir(Main.Main_MasterPath & "enquiries\" & xselect & ".xls", vbNormal) <> "" Then
    MsgBox ("File currently in Enquires")
    Exit Sub
    Workbooks.Open Main.Main_MasterPath & "enquiries\" & xselect & ".xls", ReadOnly:=True
    Sheets("job card").Select
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
End If

If Dir(Main.Main_MasterPath & "archive\" & xselect & ".xls", vbNormal) <> "" Then
    Workbooks.Open Main.Main_MasterPath & "archive\" & xselect & ".xls", ReadOnly:=True
    Sheets("job card").Select
        
    'ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
End If

If Dir(Main.Main_MasterPath & "wip\" & xselect & ".xls", vbNormal) <> "" Then
    Workbooks.Open Main.Main_MasterPath & "wip\" & xselect & ".xls", ReadOnly:=True
    Sheets("job card").Select
    
    'ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
End If

If Dir(Main.Main_MasterPath & "QUOTES\" & xselect & ".xls", vbNormal) <> "" Then
    Workbooks.Open Main.Main_MasterPath & "quotes\" & xselect & ".xls", ReadOnly:=True
    Sheets("job card").Select
    
    'ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
End If

Sheets("Job Card").Select
Application.Dialogs(xlDialogPrint).Show

ActiveWorkbook.Close False

End Sub

Private Sub GotoSearch_Change()
Main.lst.Clear
Main.lst.AddItem Main.GotoSearch.Value
End Sub

Private Sub JobsInWIP_Click()

If Main.JobsInWIP.Value = True Then
    Main.Enquiries.Value = False
    Main.WIP.Value = False
    Main.Archive.Value = False
    Main.Quotes.Value = False
    Main.Thirties.Value = False
Else
    Exit Sub
End If

Main.lst.Clear

Workbooks.Open Main.Main_MasterPath & "WIP.xls", ReadOnly:=True
Range("A1").Select
Selection.End(xlToRight).Select
col = ActiveCell.Column

Range("A1").Select
Selection.End(xlDown).Select

Range("A2", Range("A2").Offset(ActiveCell.Row, col - 1).Address).Select
    Range(Selection, Selection.End(xlDown)).Select
    
    Selection.Sort Key1:=Range("c3"), Order1:=xlDescending, Header:=xlYes, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortTextAsNumbers
Range("c3").Select

Do
    If ActiveCell.FormulaR1C1 <> "" Then Main.lst.AddItem ActiveCell.FormulaR1C1
    ActiveCell.Offset(1, 0).Select
Loop Until ActiveCell.FormulaR1C1 = ""

ActiveWorkbook.Close False

End Sub

Private Sub JumpTheGun_Click()
On Error GoTo 9
Dim Missed(1 To 100) As Integer

xselect = "_Enq"
x = OpenBook(Main.Main_MasterPath.Value & "Templates\" & xselect & ".xls", True)
Windows(xselect & ".xls").Activate

FJG.but_SaveAsCTItem.Visible = False
FJG.butSaveJG.Visible = True

FJG.Show
            
Sheets("Job Card").Select
Range("A1").Select
Range("r3").FormulaR1C1 = ""

ActiveWorkbook.SaveAs Main.Main_MasterPath & "wip\" & FJG.File_Name.Value & ".xls"
Sheets("Job Card").Select
ActiveWorkbook.Close True
 
9:

x = OpenBook(Main.Main_MasterPath & "wip\" & FJG.File_Name.Value & ".xls", False)

ActiveWorkbook.Close True

Unload FAcceptQuote
Unload FList
Unload FJG

Exit Sub
Add_Enquiry_Click
Main.lst.Clear
Main.lst.AddItem Jump_TheGun
Main.lst.Value = Jump_TheGun
Make_Quote_Click
Main.lst.Clear
Main.lst.AddItem Jump_TheGun
Main.lst.Value = Jump_TheGun
CalledThrough_Click

Main.lst.AddItem Jump_TheGun
Main.lst.Value = Jump_TheGun
createjob_Click

End Sub

Private Sub lst_Click()
Dim Pric As Currency
On Error GoTo Err

If InStr(1, Main.lst.Value, "*") > 1 Then
    xselect = Left(Main.lst.Value, Len(Main.lst.Value) - 2)
Else
    xselect = Main.lst.Value
End If

    For Each ctl In Me.Controls
        If TypeName(ctl) = "Textbox" Then ctl.Value = ""
    Next ctl
    
    If Dir(Main.Main_MasterPath.Value & "enquiries\" & xselect & ".xls", vbNormal) <> "" Then
        x = OpenBook(Main.Main_MasterPath.Value & "Enquiries\" & xselect & ".xls", True)
    End If
    If Dir(Main.Main_MasterPath.Value & "quotes\" & xselect & ".xls", vbNormal) <> "" Then
        x = OpenBook(Main.Main_MasterPath.Value & "Quotes\" & xselect & ".xls", True)
    End If
    If Dir(Main.Main_MasterPath.Value & "archive\" & xselect & ".xls", vbNormal) <> "" Then
        x = OpenBook(Main.Main_MasterPath.Value & "Archive\" & xselect & ".xls", True)
    End If
    If Dir(Main.Main_MasterPath.Value & "wip\" & xselect & ".xls", vbNormal) <> "" Then
        x = OpenBook(Main.Main_MasterPath.Value & "WIP\" & xselect & ".xls", True)
    End If

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
                            If UCase(TypeName(ctl)) = "TEXTBOX" Then
                                If InStr(1, ctl.Name, "Date", vbTextCompare) > 0 Then
                                    ctl.Value = Format(.Range("A1").Offset(i, 1).Value, "dd mmm yyyy")
                                Else
                                    ctl.Value = .Range("A1").Offset(i, 1).Value
                                End If
                            End If
                        
                        
                        GoTo FormLoadNext
                    End If
                Loop Until .Range("A1").Offset(i, 0).Value = ""
FormLoadNext:
            Next ctl
        End With
    ActiveWorkbook.Close False
    
Exit Sub
Err:
    MsgBox ("An unexpected error occured - Err# 20090302a")
    'Resume

End Sub

Private Sub Lst_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

If InStr(1, Main.lst.Value, "*") > 1 Then
    xselect = Left(Main.lst.Value, Len(Main.lst.Value) - 2)
Else
    xselect = Main.lst.Value
End If

If Dir(Main.Main_MasterPath.Value & "enquiries\" & xselect & ".xls", vbNormal) <> "" Then x = OpenBook(Main.Main_MasterPath & "enquiries\" & xselect & ".xls", True)
If Dir(Main.Main_MasterPath.Value & "archive\" & xselect & ".xls", vbNormal) <> "" Then x = OpenBook(Main.Main_MasterPath & "archive\" & xselect & ".xls", True)
If Dir(Main.Main_MasterPath.Value & "wip\" & xselect & ".xls", vbNormal) <> "" Then x = OpenBook(Main.Main_MasterPath & "WIP\" & xselect & ".xls", True)
If Dir(Main.Main_MasterPath.Value & "QUOTES\" & xselect & ".xls", vbNormal) <> "" Then x = OpenBook(Main.Main_MasterPath & "QUOTES\" & xselect & ".xls", True)

Unload Main

End Sub

Private Sub createjob_Click()
Dim JC As String
Dim Typ(1 To 20) As String
Dim Seq(1 To 20) As String
Dim Comments(1 To 20) As String
Dim OP(1 To 20) As String

1:

On Error GoTo 9

If Main.lst.ListIndex < 0 Then
    MsgBox ("Please select a job")
    Exit Sub
End If

If MsgBox("Do you wish to make this quote (" & Main.lst.Value & ") a job?", vbYesNo) = vbNo Then Exit Sub

If InStr(1, Main.lst.Value, "*") > 1 Then
    xselect = Left(Main.lst.Value, Len(Main.lst.Value) - 2)
Else
    xselect = Main.lst.Value
End If

If Dir(Main.Main_MasterPath & "Archive\" & xselect & ".xls", vbNormal) <> "" Then
    If UCase(Me.System_Status.Value) = UCase("Quote Submitted") Then
        'x = OpenBook(Main.Main_MasterPath & "Archive\" & xselect & ".xls", False)
        'Sheets("Admin").Select
        
        FAcceptQuote.Show
        
        Unload FAcceptQuote
        
    End If
End If

Refresh_Main
If Main.Visible = False Then Main.Show
Exit Sub

9:
Range("b15").FormulaR1C1 = InputBox("Please enter the date" & vbNewLine & "A calendar should've been displayed (I will set this up on your machines", "MEM", Now())
Resume

OTher:
MsgBox ("An unexpected error has occured - Please quote CreateJob_Click")
End

End Sub

Private Sub Make_Quote_Click()

On Error GoTo 9

If Main.lst.ListIndex < 0 Then
    MsgBox ("Please select a job")
    Exit Sub
End If

If MsgBox("Do you wish to make this enquiry (" & Main.lst.Value & ") a quote?", vbYesNo) = vbNo Then Exit Sub

If Dir(Main.Main_MasterPath.Value & "enquiries\" & Main.lst.Value & ".xls", vbNormal) <> "" Then
    
    FQuote.Show
       
    If Main.WIP.Value = True Then
        Main.WIP.Value = False
        Main.WIP.Value = True
    End If
    
    If Main.Enquiries.Value = True Then
        Main.Enquiries.Value = False
        Main.Enquiries.Value = True
    End If
    
    If Main.Archive.Value = True Then
        Main.Archive.Value = False
        Main.Archive.Value = True
    End If
    
    If Main.JobsInWIP.Value = True Then
        Main.JobsInWIP.Value = False
        Main.JobsInWIP.Value = True
    End If
    
End If

Exit Sub

9:
Range("b15").FormulaR1C1 = InputBox("Please enter the date" & vbNewLine & "A calendar should've been displayed (I will set this up on your machines", "MEM", Now())
Resume


End Sub

Private Sub OpenJob_Click()

If InStr(1, Main.lst.Value, "*") > 1 Then
    xselect = Left(Main.lst.Value, Len(Main.lst.Value) - 2)
Else
    xselect = Main.lst.Value
End If

If Me.System_Status.Value <> "QUOTE ACCEPTED" Then
    MsgBox ("Please accept this quote first")
    Exit Sub
End If
FJobCard.Show

End Sub

Private Sub OpenWIP_Click()

x = OpenBook(Main.Main_MasterPath & "WIP.xls", False)

    Range("A2:as2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Sort Key1:=Range("D2"), Order1:=xlAscending, Key2:=Range("G2") _
        , Order2:=xlAscending, Header:=xlYes, OrderCustom:=1, MatchCase:=False _
        , Orientation:=xlTopToBottom, DataOption1:=xlSortTextAsNumbers, DataOption2:= _
        xlSortNormal

Range("C2").Select
Unload Main

End Sub

Private Sub Quotes_Click()

If Main.Quotes.Value = True Then

    Main.lst.Clear

    x = List_Files("quotes", Main.lst)
    Main.Notice_Quotes.Caption = "Quotes : " & Check_Files(Main.Main_MasterPath & "Quotes\")

    Main.Enquiries.Value = False
    Main.WIP.Value = False
    Main.Archive.Value = False
    Main.JobsInWIP.Value = False
    Main.Thirties.Value = False
'    NextCheck = Now + TimeValue("00:00:05")
'    Application.OnTime NextCheck, "CheckUpdates"

End If

End Sub

Private Sub Search_Click()
Workbooks.Open Main.Main_MasterPath & "search.xls", ReadOnly:=True
Range("b1").Select
Main.Hide
Application.Run "Search.xls!Show_Search_Menu"
Unload Main
Exit Sub

FSearch1.Show
End Sub

Private Sub Thirties_Click()

If Main.Thirties.Value = True Then
    Main.Enquiries.Value = False
    Main.WIP.Value = False
    Main.Archive.Value = False
    Main.Quotes.Value = False
    Main.JobsInWIP.Value = False
Else
    Exit Sub
End If

Main.lst.Clear

Workbooks.Open Main.Main_MasterPath & "search.xls", ReadOnly:=True
    Range("A1").Select
        Selection.End(xlToRight).Select
        col = ActiveCell.Column
    Range("A1").Select
    Selection.End(xlDown).Select
    Range("A2", Range("A2").Offset(ActiveCell.Row, col - 1).Address).Select
    Range(Selection, Selection.End(xlDown)).Select
    
    Selection.Sort Key1:=Range("c3"), Order1:=xlDescending, Header:=xlYes, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortTextAsNumbers
Range("a3").Select

Do
    If Len(ActiveCell.FormulaR1C1) >= 5 Then
        If CCur(Left(ActiveCell.FormulaR1C1, 5)) > 29999 And CCur(Left(ActiveCell.FormulaR1C1, 5)) < 99999 Then Main.lst.AddItem ActiveCell.FormulaR1C1
'    Main.Hide
    End If
    ActiveCell.Offset(1, 0).Select
Loop Until ActiveCell.FormulaR1C1 = ""

    Range("A1").Select
        Selection.End(xlToRight).Select
        col = ActiveCell.Column
    Range("A1").Select
    Selection.End(xlDown).Select
    Range("A2", Range("A2").Offset(ActiveCell.Row, col - 1).Address).Select
    Selection.Sort Key1:=Range("A2"), Order1:=xlDescending, Header:=xlYes, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortTextAsNumbers

ActiveWorkbook.Close False

End Sub

Private Sub UserForm_Activate()

Main.Width = Application.Width
Main.Height = Application.Height
Main.Top = Application.Top
Main.Left = Application.Left

'MenuFrame.Left = (Application.Width / 2) - (MenuFrame.Width / 2)
MenuFrame.Left = Application.Width - MenuFrame.Width

MasterFile = ActiveWorkbook.Name
CheckUpdates

WIP_Click

End Sub

Private Sub UserForm_Terminate()
StopCheck
Unload Main
End Sub

Private Sub WIP_Click()

If Main.WIP.Value = True Then
    
    Main.lst.Clear

    x = List_Files("WIP", Main.lst)
    Main.Notice_WIP.Caption = "WIP : " & Check_Files(Main.Main_MasterPath & "WIP\")
    Main.Quotes.Value = False
    Main.Enquiries.Value = False
    Main.Archive.Value = False
    Main.JobsInWIP.Value = False
    Main.Thirties.Value = False
    
End If

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

    If GetValue = 0 Then GetValue = ""
    
End Function

Private Sub WIPReport_Click()

fwip.Show

End Sub




