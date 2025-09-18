VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FJobCard 
   Caption         =   "MEM: Job Card"
   ClientHeight    =   10200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9615.001
   OleObjectBlob   =   "FJobCard.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FJobCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function RefreshFJobCard()

With FJobCard
    .Operation01_Type.Value = ""
    .Operation02_Type.Value = ""
    .Operation03_Type.Value = ""
    .Operation05_Type.Value = ""
    .Operation06_Type.Value = ""
    .Operation04_Type.Value = ""
    .Operation07_Type.Value = ""
    .Operation08_Type.Value = ""
    .Operation09_Type.Value = ""
    .Operation10_Type.Value = ""
    .Operation11_Type.Value = ""
    .Operation12_Type.Value = ""
    .Operation13_Type.Value = ""
    .Operation14_Type.Value = ""
    .Operation15_Type.Value = ""
    .Operation01_Operator.Value = ""
    .Operation02_Operator.Value = ""
    .Operation03_Operator.Value = ""
    .Operation04_Operator.Value = ""
    .Operation05_Operator.Value = ""
    .Operation06_Operator.Value = ""
    .Operation07_Operator.Value = ""
    .Operation08_Operator.Value = ""
    .Operation09_Operator.Value = ""
    .Operation10_Operator.Value = ""
    .Operation11_Operator.Value = ""
    .Operation12_Operator.Value = ""
    .Operation13_Operator.Value = ""
    .Operation14_Operator.Value = ""
    .Operation15_Operator.Value = ""
    .Operation01_Comment.Value = ""
    .Operation02_Comment.Value = ""
    .Operation03_Comment.Value = ""
    .Operation04_Comment.Value = ""
    .Operation05_Comment.Value = ""
    .Operation06_Comment.Value = ""
    .Operation07_Comment.Value = ""
    .Operation08_Comment.Value = ""
    .Operation09_Comment.Value = ""
    .Operation10_Comment.Value = ""
    .Operation11_Comment.Value = ""
    .Operation12_Comment.Value = ""
    .Operation13_Comment.Value = ""
    .Operation14_Comment.Value = ""
    .Operation15_Comment.Value = ""
End With

End Function

Private Sub CopyFromJobCard_Click()

xselect = InputBox("Please enter the Job Number you wish to copy from")

    For Each ctl In Me.Controls
        If TypeName(ctl) = "Label" Then ctl.Caption = ""
        If TypeName(ctl) = "Textbox" Then ctl.Value = ""
        If TypeName(ctl) = "ComboBox" Then ctl.Value = ""
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
                If UCase(ctl.Name) = "JOB_NUMBER" Then GoTo FormLoadNext
                If UCase(ctl.Name) = "ENQUIRY_NUMBER" Then GoTo FormLoadNext
                If UCase(ctl.Name) = "QUOTE_NUMBER" Then GoTo FormLoadNext
                If UCase(ctl.Name) = "FILE_NAME" Then GoTo FormLoadNext
                'MsgBox (ctl.Name)
                i = -1
                Do
                    
                    i = i + 1
                    'MsgBox (.Range("A1").Offset(i, 0).Value)
                    If UCase(.Range("A1").Offset(i, 0).Value) = UCase(ctl.Name) And UCase(ctl.Name) = "JOB_PICTUREPATH" Then
                        If TypeName(ctl) = "Label" Then ctl.Caption = Insert_Characters(ctl.Name) & " : " & .Range("A1").Offset(i, 1).Value
                        If UCase(TypeName(ctl)) = "COMBOBOX" Then ctl.Value = .Range("A1").Offset(i, 1).Value
                        If UCase(TypeName(ctl)) = "TEXTBOX" Then ctl.Value = .Range("A1").Offset(i, 1).Value
                        GoTo FormLoadNext
                    End If
                    If UCase(.Range("A1").Offset(i, 0).Value) = UCase(ctl.Name) And Left(UCase(ctl.Name), 9) = "OPERATION" Then
        '                MsgBox (TypeName(ctl))
                            If TypeName(ctl) = "Label" Then ctl.Caption = Insert_Characters(ctl.Name) & " : " & .Range("A1").Offset(i, 1).Value
                            If UCase(TypeName(ctl)) = "COMBOBOX" Then ctl.Value = .Range("A1").Offset(i, 1).Value
                            If UCase(TypeName(ctl)) = "TEXTBOX" Then ctl.Value = .Range("A1").Offset(i, 1).Value
                        
                        GoTo FormLoadNext
                    End If
                Loop Until .Range("A1").Offset(i, 0).Value = ""
FormLoadNext:
            Next ctl
        End With
    ActiveWorkbook.Close False

'Me.Job_Number.Value = ""
'Me.Enquiry_Number.Value = ""
'Me.Quote_Number.Value = ""
Me.Job_StartDate.Value = Format(Now(), "dd mmm yyyy")

End Sub

Private Sub Job_PicturePath_Change()
On Error GoTo 9
    Sheets("Job Card").Select
    ActiveSheet.Shapes("Drawing").Delete

9:
If FJobCard.Job_PicturePath.Value <> "" Then FJobCard.Drawing.Picture = LoadPicture(Main.Main_MasterPath.Value & "images\" & FJobCard.Job_PicturePath.Value)
End Sub

Private Sub JobCardTemplates_Click()
On Error GoTo ErrJC

'fileextension = right(ucase(InputBox("Please enter file extension")),3)
fileextension = "*.*"

FList.lst.Clear

1:

'ChDrive (main.Main_MasterPath & path)

MyName = Dir(Main.Main_MasterPath & "Job Templates" & "\", vbDirectory)
    If MyName = "" Then
        MsgBox "Folder Not Found", vbOKOnly, "Test"
            Exit Sub
    End If
'\* Store list of Group folder names

i = 0
Do Until MyName = ""

    If MyName = "." Or MyName = ".." Then GoTo 2

    i = i + 1
    FList.lst.AddItem Left(MyName, Len(MyName) - 4)
2:
    
    GroupCount = GroupCount + 1
    
    MyName = Dir
    
Loop

FList.Show
RefreshFJobCard

With Me

    .Operation01_Type.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "A2")
    .Operation02_Type.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "A3")
    .Operation03_Type.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "A4")
    .Operation04_Type.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "A5")
    .Operation05_Type.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "A6")
    .Operation06_Type.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "A7")
    .Operation07_Type.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "A8")
    .Operation08_Type.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "A9")
    .Operation09_Type.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "A10")
    .Operation10_Type.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "A11")
    .Operation11_Type.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "A12")
    .Operation12_Type.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "A13")
    .Operation13_Type.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "A14")
    .Operation14_Type.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "A15")
    .Operation15_Type.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "A16")
    .Operation01_Operator.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "b2")
    .Operation02_Operator.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "b3")
    .Operation03_Operator.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "b4")
    .Operation04_Operator.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "b5")
    .Operation05_Operator.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "b6")
    .Operation06_Operator.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "b7")
    .Operation07_Operator.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "b8")
    .Operation08_Operator.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "b9")
    .Operation09_Operator.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "b10")
    .Operation10_Operator.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "b11")
    .Operation11_Operator.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "b12")
    .Operation12_Operator.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "b13")
    .Operation13_Operator.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "b14")
    .Operation14_Operator.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "b15")
    .Operation15_Operator.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "b16")
    .Operation01_Comment.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "c2")
    .Operation02_Comment.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "c3")
    .Operation03_Comment.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "c4")
    .Operation04_Comment.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "c5")
    .Operation05_Comment.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "c6")
    .Operation06_Comment.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "c7")
    .Operation07_Comment.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "c8")
    .Operation08_Comment.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "c9")
    .Operation09_Comment.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "c10")
    .Operation10_Comment.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "c11")
    .Operation11_Comment.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "c12")
    .Operation12_Comment.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "c13")
    .Operation13_Comment.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "c14")
    .Operation14_Comment.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "c15")
    .Operation15_Comment.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "c16")

End With

Exit Sub
ErrJC:
    MsgBox ("An Error has occured : 20070625")
    Resume

End Sub

Private Sub SaveJobCard_Click()
On Error GoTo 9
Dim Missed(1 To 100) As Integer
Dim xselect As String
 
If InStr(1, Main.lst.Value, "*") > 1 Then
    xselect = Left(Main.lst.Value, Len(Main.lst.Value) - 2)
Else
    xselect = Main.lst.Value
End If

If Me.Job_Number.Value = "" Then
    Me.Job_Number.Value = Confirm_Next_Number("J")
End If

If FJobCard.Job_StartDate = "" Then
    FJobCard.Job_StartDate = Format(CDate(Now()), "dd-mmm-yyyy")
End If

Me.File_Name.Value = Me.Job_Number.Value

x = OpenBook(Main.Main_MasterPath.Value & "WIP\" & xselect & ".xls", False)
Windows(xselect & ".xls").Activate
    
    Sheets("Admin").Select
    Me.System_Status.Value = UCase("Job Open")

    Sheets("Job Card").Select
    
' SaveColumnsToFile
j = -1
i = 1
With Worksheets("ADMIN")
    For Each ctl In Me.Controls
        For i = 0 To 100
            If UCase(.Range("A1").Offset(i, 0).FormulaR1C1) = UCase(ctl.Name) And Left(.Range("A1").Offset(i, 0).Formula, 1) <> "=" Then
                'MsgBox .Range("A1").Offset(i, 0).FormulaR1C1
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

Sheets("Job Card").Select
Range("A1").Select

Range("r3").FormulaR1C1 = ""

' Save to WIP
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
    Loop Until ActiveCell.Offset(0, 2).FormulaR1C1 = "" Or _
        ActiveCell.Offset(0, 2).FormulaR1C1 = Me.Quote_Number.Value Or _
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
                    If UCase(.Range("a1").Offset(0, 1).FormulaR1C1) = "" Then GoTo FormNextWIP
            Next i
FormNextWIP:
        Next ctl
    End With
  
Range("A1").Select
Selection.End(xlToRight).Select
col = ActiveCell.Column

Range("A1").Select
Selection.End(xlDown).Select

Range("A2", Range("A2").Offset(ActiveCell.Row, col - 1).Address).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Sort Key1:=Range("h3"), Order1:=xlAscending, Header:=xlYes, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
              
ActiveWorkbook.Close (True)

If UCase(ActiveWorkbook.path) = UCase(Main.Main_MasterPath.Value & "Archive") Then
    ActiveWorkbook.Close (True)
Else
    ActiveWorkbook.SaveAs (Main.Main_MasterPath.Value & "Archive\" & Me.Job_Number.Value & ".xls")
    ActiveWorkbook.Close
    Kill (Main.Main_MasterPath & "WIP\" & xselect & ".xls")
End If

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
    
'    ' Check Search Find for Office 2000
    Columns("A:A").Select
    Selection.Find(What:=Me.File_Name.Value, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
        
'    Do
'        ActiveCell.Offset(1, 0).Select
'    Loop Until ActiveCell.FormulaR1C1 = "" Or _
'        ActiveCell.FormulaR1C1 = Me.Quote_Number.Value Or _
'        ActiveCell.FormulaR1C1 = Me.Enquiry_Number.Value Or _
'        ActiveCell.FormulaR1C1 = Me.Job_Number.Value Or _
'       ActiveCell.FormulaR1C1 = Me.File_Name.Value
    
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

Unload FJobCard

x = OpenBook(Main.Main_MasterPath.Value & "Archive\" & Me.Job_Number.Value & ".xls", False)
Unload Main

Exit Sub
9:
MsgBox ("Error - debug")
Resume
Refresh_Main

End Sub

Private Sub UserForm_Initialize()
Dim FullFilePath As String, MyName As String
Dim GroupCount As Integer
'\* Check a Group folder exists
'FullFilePath = "C:\TEMP\Group*"

'fileextension = right(ucase(InputBox("Please enter file extension")),3)
fileextension = "*.*"

1:

'ChDrive (main.Main_MasterPath & path)

MyName = Dir(Main.Main_MasterPath.Value & "images\", vbDirectory)
    If MyName = "" Then
        MsgBox "Folder Not Found", vbOKOnly, "Test"
            Exit Sub
    End If
'\* Store list of Group folder names

Do Until MyName = ""

    If MyName = "." Or MyName = ".." Then GoTo 2

    With FJobCard.Job_PicturePath
        .AddItem MyName 'Left(MyName, Len(MyName) - 4)
    End With
2:
    
    GroupCount = GroupCount + 1
    
    MyName = Dir
    
Loop

'If Dir(main.Main_MasterPath & "enquiries\" & Main.Lst.Value & ".xls", vbNormal) <> "" Then OpenBook (main.Main_MasterPath & "enquiries\" & Main.Lst.Value & ".xls")
'If Dir(main.Main_MasterPath & "archive\" & Main.Lst.Value & ".xls", vbNormal) <> "" Then OpenBook (main.Main_MasterPath & "archive\" & Main.Lst.Value & ".xls")
If InStr(1, Main.lst.Value, "*") > 1 Then
    xselect = Left(Main.lst.Value, Len(Main.lst.Value) - 2)
Else
    xselect = Main.lst.Value
End If

x = OpenBook(Main.Main_MasterPath.Value & "WIP\" & xselect & ".xls", True)
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
            
x = OpenBook(Main.Main_MasterPath & "Operations.xls", True)
Range("A2").Select

Do
    With FJobCard
                   
        Typ = ActiveCell.FormulaR1C1
               
        .Operation01_Type.AddItem Typ
        .Operation02_Type.AddItem Typ
        .Operation03_Type.AddItem Typ
        .Operation05_Type.AddItem Typ
        .Operation06_Type.AddItem Typ
        .Operation04_Type.AddItem Typ
        .Operation07_Type.AddItem Typ
        .Operation08_Type.AddItem Typ
        .Operation09_Type.AddItem Typ
        .Operation10_Type.AddItem Typ
        .Operation11_Type.AddItem Typ
        .Operation12_Type.AddItem Typ
        .Operation13_Type.AddItem Typ
        .Operation14_Type.AddItem Typ
        .Operation15_Type.AddItem Typ
    End With
    
    ActiveCell.Offset(1, 0).Select
Loop Until ActiveCell.FormulaR1C1 = ""
ActiveWorkbook.Close

FJobCard.Operation01_Type.SetFocus

End Sub

Private Sub UserForm_Terminate()

If ActiveWorkbook.Name <> MasterFile And Left(ActiveWorkbook.Name, 5) <> "_Inte" Then
    ActiveWindow.Close (False)
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




