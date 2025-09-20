VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FJG 
   Caption         =   "FJG"
   ClientHeight    =   9720.001
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13650
   OleObjectBlob   =   "FJG.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FJG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub but_SaveAsCTItem_Click()
Dim CTFileName As String

Me.Job_StartDate.Value = ""

    ' SaveColumnsToFile
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

'Application.Dialogs(xlDialogSaveAs).Show
CTFileName = InputBox("Please enter the filename that you wish this file to be saved as")
ActiveWorkbook.SaveAs Main.Main_MasterPath.Value & "Contracts\" & CTFileName & ".xls"

ActiveWorkbook.Close True
Unload FJG

End Sub

Private Sub butSaveJG_Click()
Dim ctl As Object
Dim z As Integer

    Me.Enquiry_Number.Value = BusinessLogic.Calc_Next_Number("E")
    BusinessLogic.Confirm_Next_Number ("E")
    
    If Me.Compilation_TotalNumber.Value > 1 Then
        If Me.Compilation_SequenceNumber.Value = 1 Then
            Me.Quote_Number.Value = BusinessLogic.Calc_Next_Number("Q") & "-1"
            BusinessLogic.Confirm_Next_Number ("q")
            Me.Job_Number.Value = BusinessLogic.Calc_Next_Number("J") & "-1"
            BusinessLogic.Confirm_Next_Number ("J")
        Else
            Me.Quote_Number.Value = Left(Me.Quote_Number.Value, Len(Me.Quote_Number.Value) - 2) & "-" & Me.Compilation_SequenceNumber.Value
            Me.Job_Number.Value = Left(Me.Job_Number.Value, Len(Me.Job_Number.Value) - 2) & "-" & Me.Compilation_SequenceNumber.Value
        End If
    Else
        Me.Job_Number.Value = BusinessLogic.Calc_Next_Number("J")
        BusinessLogic.Confirm_Next_Number ("J")
        Me.Quote_Number.Value = BusinessLogic.Calc_Next_Number("Q")
        BusinessLogic.Confirm_Next_Number ("q")
    End If
    
    Me.File_Name.Value = Me.Job_Number.Value
    Me.System_Status.Value = UCase("Quote Accepted")
    
    ' SaveColumnsToFile
    j = -1
    i = 1
    xselect = "_Enq"
    x = FileOperations.OpenBook(Main.Main_MasterPath.Value & "Templates\" & xselect & ".xls", True)
    Windows(xselect & ".xls").Activate
    
    With Worksheets("ADMIN")
        For Each ctl In Me.Controls
            For i = 0 To 100
                If UCase(.Range("A1").Offset(i, 0).FormulaR1C1) = UCase(ctl.Name) And Left(.Range("A1").Offset(i, 1).Formula, 1) <> "=" Then
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
    
    If Me.Job_PicturePath.Value <> "" Then
        Sheets("jOB cARD").Select
        Range("Drawing_Location").Select
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
    
    'Save To Search
    x = FileOperations.OpenBook(Main.Main_MasterPath & "Search.xls", False)
        Do
            If ActiveWorkbook.ReadOnly = True Then
                ActiveWorkbook.Close
                MsgBox ("This workbook is read only, please find the user with this workbook open and close it.")
                x = FileOperations.OpenBook(Main.Main_MasterPath & "Search.xls", False)
            End If
        Loop Until ActiveWorkbook.ReadOnly = False
    
    Range("A1").Select
    
    Selection.End(xlDown).Select
    
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
                    GoTo 6
                End If
                If Left(.Range("A1").Offset(ActiveCell.Row - 2, i).FormulaR1C1, 1) = "=" Then .Range("A1").Offset(ActiveCell.Row - 1, i).FormulaR1C1 = .Range("A1").Offset(ActiveCell.Row - 2, i).FormulaR1C1
                If UCase(.Range("a1").Offset(0, i + 1).FormulaR1C1) = "" Then GoTo 6
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
        
    If CInt(Me.Compilation_SequenceNumber) < CInt(Me.Compilation_TotalNumber.Value) Then
        Me.Compilation_SequenceNumber.Value = CInt(Me.Compilation_SequenceNumber.Value) + 1
            
            ActiveWorkbook.SaveAs Main.Main_MasterPath & "wip\" & FJG.File_Name.Value & ".xls"
            ActiveWorkbook.Close True
            
            Me.Enquiry_Number.Value = ""
            'Me.Quote_Number.Value = ""
            Me.File_Name.Value = ""
            
            Me.Component_Code.Value = ""
            Me.Component_Grade.Value = ""
            Me.Component_Description.Value = ""
            Me.Component_DrawingNumber_SampleNumber.Value = ""
            Me.Component_Price.Value = ""
            Me.Job_PicturePath.Value = ""
            Me.Operation01_Comment.Value = ""
            Me.Operation01_Operator.Value = ""
            Me.Operation01_Type.Value = ""
            Me.Operation02_Comment.Value = ""
            Me.Operation02_Operator.Value = ""
            Me.Operation02_Type.Value = ""
            Me.Operation03_Comment.Value = ""
            Me.Operation03_Operator.Value = ""
            Me.Operation03_Type.Value = ""
            Me.Operation04_Comment.Value = ""
            Me.Operation04_Operator.Value = ""
            Me.Operation04_Type.Value = ""
            Me.Operation05_Comment.Value = ""
            Me.Operation05_Operator.Value = ""
            Me.Operation05_Type.Value = ""
            Me.Operation06_Comment.Value = ""
            Me.Operation06_Operator.Value = ""
            Me.Operation06_Type.Value = ""
            Me.Operation07_Comment.Value = ""
            Me.Operation07_Operator.Value = ""
            Me.Operation07_Type.Value = ""
            Me.Operation08_Comment.Value = ""
            Me.Operation08_Operator.Value = ""
            Me.Operation08_Type.Value = ""
            Me.Operation09_Comment.Value = ""
            Me.Operation09_Operator.Value = ""
            Me.Operation09_Type.Value = ""
            Me.Operation10_Comment.Value = ""
            Me.Operation10_Operator.Value = ""
            Me.Operation10_Type.Value = ""
            Me.Operation11_Comment.Value = ""
            Me.Operation11_Operator.Value = ""
            Me.Operation11_Type.Value = ""
            Me.Operation12_Comment.Value = ""
            Me.Operation12_Operator.Value = ""
            Me.Operation12_Type.Value = ""
            Me.Operation13_Comment.Value = ""
            Me.Operation13_Operator.Value = ""
            Me.Operation13_Type.Value = ""
            Me.Operation14_Comment.Value = ""
            Me.Operation14_Operator.Value = ""
            Me.Operation14_Type.Value = ""
            Me.Operation15_Comment.Value = ""
            Me.Operation15_Operator.Value = ""
            Me.Operation15_Type.Value = ""
            

        xselect = "_Enq"
        x = FileOperations.OpenBook(Main.Main_MasterPath.Value & "Templates\" & xselect & ".xls", True)
        Windows(xselect & ".xls").Activate
        
        MsgBox ("Please enter the next components details")
        Exit Sub
    End If

Me.Hide

End Sub

Private Sub CopyFromJobCard_Click()

xselect = InputBox("Please enter the Job Number you wish to copy from")

    For Each ctl In Me.Controls
        If InStr(1, Left(UCase(ctl.Name), 6), "OPERAT", vbTextCompare) > 0 Then
            If TypeName(ctl) = "Textbox" Then ctl.Value = ""
            If TypeName(ctl) = "ComboBox" Then ctl.Value = ""
        End If
    Next ctl

    If Dir(Main.Main_MasterPath.Value & "enquiries\" & xselect & ".xls", vbNormal) <> "" Then
        x = FileOperations.OpenBook(Main.Main_MasterPath.Value & "Enquiries\" & xselect & ".xls", True)
        GoTo FileFound
    End If
    If Dir(Main.Main_MasterPath.Value & "quotes\" & xselect & ".xls", vbNormal) <> "" Then
        x = FileOperations.OpenBook(Main.Main_MasterPath.Value & "Quotes\" & xselect & ".xls", True)
        GoTo FileFound
    End If
    If Dir(Main.Main_MasterPath.Value & "archive\" & xselect & ".xls", vbNormal) <> "" Then
        x = FileOperations.OpenBook(Main.Main_MasterPath.Value & "Archive\" & xselect & ".xls", True)
        GoTo FileFound
    End If
    If Dir(Main.Main_MasterPath.Value & "wip\" & xselect & ".xls", vbNormal) <> "" Then
        x = FileOperations.OpenBook(Main.Main_MasterPath.Value & "WIP\" & xselect & ".xls", True)
        GoTo FileFound
    End If
    
    MsgBox ("File Not Found")
    Exit Sub
FileFound:

        With Sheets("Admin")
            For Each ctl In Me.Controls
                If InStr(1, Left(UCase(ctl.Name), 6), "OPERAT", vbTextCompare) > 0 Then
                    i = -1
                    Do
                        i = i + 1
                        If UCase(.Range("A1").Offset(i, 0).Value) = UCase(ctl.Name) Then
            '                MsgBox (TypeName(ctl))
                                If TypeName(ctl) = "Label" Then ctl.Caption = Insert_Characters(ctl.Name) & " : " & .Range("A1").Offset(i, 1).Value
                                If UCase(TypeName(ctl)) = "COMBOBOX" Then ctl.Value = .Range("A1").Offset(i, 1).Value
                                If UCase(TypeName(ctl)) = "TEXTBOX" Then ctl.Value = .Range("A1").Offset(i, 1).Value
                            
                            GoTo FormLoadNext
                        End If
                    Loop Until .Range("A1").Offset(i, 0).Value = ""
                End If
FormLoadNext:
            Next ctl
        End With
    ActiveWorkbook.Close False

'Me.Job_Number.Value = ""
'Me.Enquiry_Number.Value = ""
'Me.Quote_Number.Value = ""
'Me.Job_StartDate.Value = Format(Now(), "dd mmm yyyy")

End Sub

Private Sub Job_PicturePath_Change()
On Error GoTo 9
    Sheets("Job Card").Select
    ActiveSheet.Shapes("Drawing").Delete

9:
If Dir(Main.Main_MasterPath.Value & "images\" & Me.Job_PicturePath.Value, vbNormal) <> "" And Me.Job_PicturePath.Value <> "" Then
    Me.Drawing.Picture = LoadPicture(Main.Main_MasterPath.Value & "images\" & Me.Job_PicturePath.Value)
End If
End Sub

Private Sub Job_Urgency_Change()
If UCase(Me.Job_Urgency.Value) = "NORMAL" Then Me.Job_LeadTime.Value = "14"
If UCase(Me.Job_Urgency.Value) = "BREAK DOWN" Then Me.Job_LeadTime.Value = "7"
If UCase(Me.Job_Urgency.Value) = "URGENT" Then Me.Job_LeadTime.Value = "10"
End Sub

Private Sub JobCardTemplates_Click()

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

'For j = 1 To i
'    If GetValue(main.Main_MasterPath & "WIP", Fil(j), "Job card", "R3") = "New" Then
'        FList.Lst.AddItem Left(Fil(j), Len(Fil(j)) - 4) & "  *"
'    End If
'Next j

FList.Show

With Me

    .Operation01_Type.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "A2")
    .Operation02_Type.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "A3")
    .Operation03_Type.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "A4")
    .Operation04_Type.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "A5")
    .Operation05_Type.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "A6")
    .Operation06_Type.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "A7")
    .Operation07_Type.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "A8")
    .Operation08_Type.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "A9")
    .Operation09_Type.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "A10")
    .Operation10_Type.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "A11")
    .Operation11_Type.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "A12")
    .Operation12_Type.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "A13")
    .Operation13_Type.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "A14")
    .Operation14_Type.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "A15")
    .Operation15_Type.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "A16")
    .Operation01_Operator.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "b2")
    .Operation02_Operator.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "b3")
    .Operation03_Operator.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "b4")
    .Operation04_Operator.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "b5")
    .Operation05_Operator.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "b6")
    .Operation06_Operator.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "b7")
    .Operation07_Operator.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "b8")
    .Operation08_Operator.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "b9")
    .Operation09_Operator.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "b10")
    .Operation10_Operator.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "b11")
    .Operation11_Operator.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "b12")
    .Operation12_Operator.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "b13")
    .Operation13_Operator.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "b14")
    .Operation14_Operator.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "b15")
    .Operation15_Operator.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "b16")
    .Operation01_Comment.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "c2")
    .Operation02_Comment.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "c3")
    .Operation03_Comment.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "c4")
    .Operation04_Comment.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "c5")
    .Operation05_Comment.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "c6")
    .Operation06_Comment.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "c7")
    .Operation07_Comment.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "c8")
    .Operation08_Comment.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "c9")
    .Operation09_Comment.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "c10")
    .Operation10_Comment.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "c11")
    .Operation11_Comment.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "c12")
    .Operation12_Comment.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "c13")
    .Operation13_Comment.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "c14")
    .Operation14_Comment.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "c15")
    .Operation15_Comment.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.lst.Value & ".xls", "JC Seq", "c16")

End With

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

    With Me.Job_PicturePath
        .AddItem MyName 'Left(MyName, Len(MyName) - 4)
    End With
2:
    
    GroupCount = GroupCount + 1
    
    MyName = Dir
    
Loop

x = FileOperations.OpenBook(Main.Main_MasterPath.Value & "Operations.xls", True)
    Range("A2").Select
    Do
        With Me
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
ActiveWorkbook.Close False

With Me
    .Customer.Clear
    x = DirectoryHelpers.List_Files("Customers", .Customer)
    .Job_Urgency.AddItem "NORMAL"
    .Job_Urgency.AddItem "BREAK DOWN"
    .Job_Urgency.AddItem "URGENT"
End With

x = FileOperations.OpenBook(Main.Main_MasterPath.Value & "templates\price list.xls", True)
    Sheets("Component_Descriptions").Select
    Range("a2").Select
    Do
        If ActiveCell.FormulaR1C1 <> ActiveCell.Offset(-1, 0).FormulaR1C1 Then Me.Component_Code.AddItem ActiveCell.FormulaR1C1
        ActiveCell.Offset(1, 0).Select
    Loop Until ActiveCell.FormulaR1C1 = ""
ActiveWorkbook.Close False

Workbooks.Open Main.Main_MasterPath.Value & "templates\Component_Grades.xls", ReadOnly:=True
    Range("A2").Select
    Do
        With Me.Component_Grade
            .AddItem ActiveCell.FormulaR1C1
        End With
        ActiveCell.Offset(1, 0).Select
    Loop Until ActiveCell.FormulaR1C1 = ""
ActiveWorkbook.Close (False)

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

Me.Job_StartDate.Value = Format(Now(), "dd mmm yyyy")
Me.Job_LeadTime.Value = "14"

If Me.Enquiry_Number.Value = "" Then
    Me.but_SaveAsCTItem.Visible = True
    Me.butSaveJG.Visible = False
Else
    Me.but_SaveAsCTItem.Visible = False
    Me.butSaveJG.Visible = True
End If

End Sub

Private Sub UserForm_Terminate()
End
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


Public Function Insert_Characters(Str As String)

j = Len(Str)
i = 0

For i = 2 To j
    If Mid(Str, i, 1) = "_" Then
        Str = Mid(Str, 1, i - 1) & " " & Mid(Str, i + 1, Len(Str) - i)
        i = i + 1
    Else
        If UCase(Mid(Str, i, 1)) = Mid(Str, i, 1) Then
            Str = Mid(Str, 1, i - 1) & " " & Mid(Str, i, Len(Str) - i + 1)
            j = j + 1
            i = i + 1
        End If
    End If
Next i

If InStr(1, Str, "Component ", vbTextCompare) > 0 Then
    Str = Right(Str, Len(Str) - Len("Component "))
End If

Insert_Characters = Str

End Function




