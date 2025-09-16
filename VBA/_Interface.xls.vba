# VBA code extracted from _Interface.xls
# Extraction date: Mon Jun  2 11:09:07 SAST 2025

# =======================================================
# Module 58
# =======================================================
Attribute VB_Name = "Calc_Numbers"
Public Function Calc_Next_Number(Typ As String)
Dim FullFilePath As String, MyName As String
Dim GroupCount As Integer
'\* Check a Group folder exists
'FullFilePath = "C:\TEMP\Group*"

'fileextension = right(ucase(InputBox("Please enter file extension")),3)
fileextension = "*.*"

'ChDrive (main.Main_MasterPath & path)

    path = "templates"
    
    MyName = Dir(Main.Main_MasterPath.Value & path & "\", vbDirectory)
        If MyName = "" Then
            MsgBox "Folder Not Found", vbOKOnly, "Test"
                Exit Function
        End If
    '\* Store list of Group folder names
     
    Do Until MyName = ""
    
        If MyName = "." Or MyName = ".." Then GoTo 2
                
            If Left(UCase(Typ), 1) = "E" And Left(MyName, 4) = "E - " Then
                Calc_Next_Number = Mid(MyName, InStr(1, MyName, "-", vbTextCompare) + 2, Len(MyName) - 8) + 1
                
                GoTo 8
                
            End If
            
            If Left(UCase(Typ), 1) = "J" And Left(MyName, 4) = "J - " Then
                Calc_Next_Number = Mid(MyName, InStr(1, MyName, "-", vbTextCompare) + 2, Len(MyName) - 8) + 1
                
                GoTo 8
                
            End If
                 
            If Left(UCase(Typ), 1) = "Q" And Left(MyName, 4) = "Q - " Then
                Calc_Next_Number = Mid(MyName, InStr(1, MyName, "-", vbTextCompare) + 2, Len(MyName) - 8) + 1
                
                GoTo 8
                
            End If
                
            GroupCount = GroupCount + 1
            
2:
        
        MyName = Dir
        
    Loop

8:

End Function

Public Function Confirm_Next_Number(Typ As String)
Dim FullFilePath As String, MyName As String
Dim GroupCount As Integer
'\* Check a Group folder exists
'FullFilePath = "C:\TEMP\Group*"

'fileextension = right(ucase(InputBox("Please enter file extension")),3)
fileextension = "*.*"

'ChDrive (main.Main_MasterPath & path)

    path = "templates"
    
    MyName = Dir(Main.Main_MasterPath.Value & path & "\", vbDirectory)
        If MyName = "" Then
            MsgBox "Folder Not Found", vbOKOnly, "Test"
                Exit Function
        End If
    '\* Store list of Group folder names
     
    Do Until MyName = ""
    
        If MyName = "." Or MyName = ".." Then GoTo 2
                
            If Left(UCase(Typ), 1) = "E" And Left(MyName, 4) = "E - " Then
                Confirm_Next_Number = Mid(MyName, InStr(1, MyName, "-", vbTextCompare) + 2, Len(MyName) - 8) + 1
                
                FileCopy Main.Main_MasterPath & path & "\" & MyName, Main.Main_MasterPath & path & "\" & "E - " & Confirm_Next_Number & ".TXT"
                Kill Main.Main_MasterPath & path & "\" & MyName

                GoTo 8
                
            End If
            
            If Left(UCase(Typ), 1) = "J" And Left(MyName, 4) = "J - " Then
                Confirm_Next_Number = Mid(MyName, InStr(1, MyName, "-", vbTextCompare) + 2, Len(MyName) - 8) + 1
                
                FileCopy Main.Main_MasterPath & path & "\" & MyName, Main.Main_MasterPath & path & "\" & "J - " & Confirm_Next_Number & ".TXT"
                Kill Main.Main_MasterPath & path & "\" & MyName

                GoTo 8
                
            End If
            
            If Left(UCase(Typ), 1) = "Q" And Left(MyName, 4) = "Q - " Then
                Confirm_Next_Number = Mid(MyName, InStr(1, MyName, "-", vbTextCompare) + 2, Len(MyName) - 8) + 1
                
                FileCopy Main.Main_MasterPath & path & "\" & MyName, Main.Main_MasterPath & path & "\" & "Q - " & Confirm_Next_Number & ".TXT"
                Kill Main.Main_MasterPath & path & "\" & MyName

                GoTo 8
                
            End If
                
            GroupCount = GroupCount + 1
            
2:
        
        MyName = Dir
        
    Loop

8:

End Function




# =======================================================
# Module 59
# =======================================================
Attribute VB_Name = "Check_Dir"

' FUNCTION TO CHANGE DIRECTORY / CREATE DIRECTORY

Public Function CheckDir(Direc As String)

    If Dir(Direc, vbDirectory) = "" Then
        MkDir (Direc)
        ChDir (Direc)
    Else
        ChDir (Direc)
    End If
    
End Function



# =======================================================
# Module 60
# =======================================================
Attribute VB_Name = "Check_Updates"
Public NextCheck As Date

Public Function CheckUpdates()

If Main.Visible = False Or NextCheck > Now() Then
    If NextCheck = "12:00:00 AM" Then GoTo 5
    StopCheck
    Exit Function
End If

5:

On Error GoTo 8

If "Enquiries : " & Check_Files(Main.Main_MasterPath & "enquiries\") <> Main.Notice_Enquiries.Caption Then
    Main.Notice_Enquiries.Caption = "Enquiries : " & Check_Files(Main.Main_MasterPath & "enquiries\") & "*"
End If

If "Quotes : " & Check_Files(Main.Main_MasterPath & "Quotes\") <> Main.Notice_Quotes.Caption Then
    Main.Notice_Quotes.Caption = "Quotes : " & Check_Files(Main.Main_MasterPath & "Quotes\") & "*"
End If

If "WIP : " & Check_Files(Main.Main_MasterPath & "WIP\") <> Main.Notice_WIP.Caption Then
    Main.Notice_WIP.Caption = "WIP : " & Check_Files(Main.Main_MasterPath & "WIP\") & "*"
End If

NextCheck = Now + TimeValue("00:05:00")
Application.OnTime NextCheck, "CheckUpdates", NextCheck + TimeValue("00:01:00")

8:

End Function

Public Function StopCheck()

On Error Resume Next
Application.OnTime NextCheck, "CheckUpdates", , Schedule:=False


End Function
Public Function Check_Files(path As String)

Dim FullFilePath As String, MyName As String
Dim GroupCount As Integer
'\* Check a Group folder exists
'FullFilePath = "C:\TEMP\Group*"

GroupCount = 0
'fileextension = right(ucase(InputBox("Please enter file extension")),3)
fileextension = "*.*"

1:

'ChDrive (main.Main_MasterPath & path)

MyName = Dir(path, vbDirectory)
    If MyName = "" Then
        StopCheck
'        MsgBox "Folder Not Found", vbOKOnly, "Test"
            Exit Function
    End If
'\* Store list of Group folder names

Do Until MyName = ""

    If MyName = "." Or MyName = ".." Or MyName = "_Users.xls" Then GoTo 2
    
    GroupCount = GroupCount + 1

2:
    
    MyName = Dir
    
Loop
    
Check_Files = GroupCount

End Function






# =======================================================
# Module 61
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
# Module 62
# =======================================================
Attribute VB_Name = "FAcceptQuote"
Attribute VB_Base = "0{5874FC2D-FA8F-4094-ACE2-620BEF806D9D}{39DEA3E6-D1FE-4F54-84FC-956ADC2D6EDE}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub butSAVE_Click()

If FAcceptQuote.CustomerOrderNumber.Value = "" Then
    MsgBox ("Please enter a Customer Order Number")
    Exit Sub
End If

If CInt(Me.Compilation_SequenceNumber.Value) = "1" Then
    Me.Job_Number.Value = Confirm_Next_Number("J")
    If Me.Compilation_TotalNumber <> "1" Then
        Me.Job_Number.Value = Me.Job_Number.Value & "-1"
    End If
End If

Me.File_Name.Value = Me.Job_Number.Value
Me.System_Status.Value = UCase("Quote Accepted")

' SaveColumnsToFile
x = OpenBook(Main.Main_MasterPath.Value & "Archive\" & Me.Quote_Number.Value & ".xls", False)
j = -1
i = 1
With Worksheets("ADMIN")
    For Each ctl In Me.Controls
        For i = 0 To 100
            If UCase(.Range("A1").Offset(i, 0).FormulaR1C1) = UCase(ctl.Name) And Left(.Range("A1").Offset(i, 1).Formula, 1) <> "=" Then
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

' Next Compilation
If Me.Compilation_TotalNumber.Value <> "1" And Me.Compilation_SequenceNumber.Value <> Me.Compilation_TotalNumber.Value Then
    ActiveWorkbook.SaveAs (Main.Main_MasterPath.Value & "WIP\" & Me.File_Name.Value & ".xls")
    Me.Compilation_SequenceNumber.Value = CInt(Me.Compilation_SequenceNumber.Value) + 1
    Me.Job_Number.Value = Left(Me.Job_Number.Value, Len(Me.Job_Number.Value) - 2) & "-" & Me.Compilation_SequenceNumber.Value
    
    Me.File_Name.Value = Left(Me.Job_Number.Value, Len(Me.Job_Number.Value) - 2) & "-" & Me.Compilation_SequenceNumber.Value
    
    Me.Component_Code.Value = ""
    Me.Component_Description.Value = ""
    Me.Component_Price.Value = ""
    Me.Component_Grade.Enabled = True
    
    MsgBox ("Please enter the next Parts details")
    Exit Sub
End If

ContFAcceptQuote.Visible = False

    ShowSheet ("Job Card")
    Sheets("Job Card").Select
                
    ActiveWorkbook.SaveAs (Main.Main_MasterPath.Value & "WIP\" & Me.File_Name.Value & ".xls")
    ActiveWorkbook.Close
    Kill (Main.Main_MasterPath.Value & "Archive\" & Me.Quote_Number.Value & ".xls")
    Unload Me
    
End Sub

Private Sub Job_Urgency_Change()

If UCase(Me.Job_Urgency.Value) = "NORMAL" Then Me.Job_LeadTime.Value = "14"
If UCase(Me.Job_Urgency.Value) = "BREAK DOWN" Then Me.Job_LeadTime.Value = "7"
If UCase(Me.Job_Urgency.Value) = "URGENT" Then Me.Job_LeadTime.Value = "10"

End Sub

Private Sub UserForm_Activate()
ContFAcceptQuote.Visible = False
 
If InStr(1, Main.Lst.Value, "*") > 1 Then
    xselect = Left(Main.Lst.Value, Len(Main.Lst.Value) - 2)
Else
    xselect = Main.Lst.Value
End If

If FList.Lst.Value <> "" Then
    xselect = FList.Lst.Value
End If

FList.Lst.Value = ""

With FAcceptQuote.Job_Urgency
    .AddItem "NORMAL"
    .AddItem "BREAK DOWN"
    .AddItem "URGENT"
End With

x = OpenBook(Main.Main_MasterPath.Value & "Archive\" & xselect & ".xls", True)
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
                    If InStr(1, UCase(ctl.Name), UCase("Date"), vbTextCompare) > 0 Then
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
ActiveWorkbook.Close

Me.System_Status.Value = UCase("Quote Accepted")
Me.Job_StartDate.Value = Format(Now(), "dd mmm yyyy")
Me.CustomerOrderNumber.SetFocus

End Sub

Public Function GetValue(path, File, sheet, ref)
'   Retrieves a value from a closed workbook\
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




# =======================================================
# Module 63
# =======================================================
Attribute VB_Name = "FEnquiry"
Attribute VB_Base = "0{E1B1BCDF-D6BE-40CF-A868-D1BF7183761D}{7F41ABB9-12C4-4FF8-B07B-A237017A7C34}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub AddMore_Click()
Dim Eq As Integer

With Me
    .Enquiry_Number.Value = Calc_Next_Number("E")
    Confirm_Next_Number ("E")
    .File_Name.Value = .Enquiry_Number.Value
    MsgBox ("The File Number for this Enquiry is: " & Me.File_Name.Value)
End With

If FrmEnquiry.Enquiry_Date.Caption = "Please click here to insert a date" Then
    If MsgBox("Do you cancel the save in order to enter a Date?", vbYesNo, "MEM") = vbYes Then
        Exit Sub
    End If
End If
    
If FrmEnquiry.Component_Quantity = "" Then
    If MsgBox("Do you wish to cancel the save in order to enter a Component_Quantity?", vbYesNo, "MEM") = vbYes Then
        Exit Sub
    End If
End If
    
If FrmEnquiry.Customer = "" Then
    If MsgBox("Do you wish to cancel the save in order to enter a Customer?", vbYesNo, "MEM") = vbYes Then
        Exit Sub
    End If
End If
    
If FrmEnquiry.Component_Description = "" Then
    If MsgBox("Do you wish to cancel the save in order to select a product?", vbYesNo, "MEM") = vbYes Then
        Exit Sub
    End If
End If

'Windows("Price List.xls").Activate
'ActiveWorkbook.Close (False)

x = OpenBook(Main.Main_MasterPath & "Templates\" & "_Enq.xls", True)

FrmEnquiry.System_Status.Value = "To Quote"
Dim ctl As Object

j = -1
i = 1

With Worksheets("ADMIN")

    For Each ctl In Me.Controls
        For i = 0 To 100
                If UCase(.Range("A1").Offset(i, 0).FormulaR1C1) = UCase(ctl.Name) Then
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

With FrmEnquiry
    
    Sheets("admin").Select

    ActiveWorkbook.SaveAs (Main.Main_MasterPath & "enquiries\" & .Enquiry_Number & ".xls")
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

Do
    ActiveCell.Offset(1, 0).Select
Loop Until ActiveCell.FormulaR1C1 = "" Or _
    ActiveCell.FormulaR1C1 = Me.Enquiry_Number.Value Or _
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
    
End With

' New Enq

With FrmEnquiry
    .Enquiry_Number.Value = Calc_Next_Number("E")
    
    If .Customer.Value = "" Then .Customer.Clear

    x = List_Files("Customers", FrmEnquiry.Customer)

End With

x = OpenBook(Main.Main_MasterPath & "templates\price list.xls", True)

'FrmEnquiry.Enquiry_Date.Caption = Now()

Sheets("Component_Descriptions").Select
Range("a2").Select

Do
    If ActiveCell.FormulaR1C1 <> ActiveCell.Offset(-1, 0).FormulaR1C1 Then FrmEnquiry.Component_Code.AddItem ActiveCell.FormulaR1C1
    ActiveCell.Offset(1, 0).Select
Loop Until ActiveCell.FormulaR1C1 = ""

' To add more Component_Grades to the enquiry form, add them below using ".AddItem " & What you want to add

Sheets("Component_Descriptions").Select


With Me
    .Enquiry_Number.Value = ""
    .File_Name.Value = ""
    .Component_Description = ""
    .Component_Code = ""
    .Component_Grade = ""
    .Component_DrawingNumber_SampleNumber = ""
    .Component_Quantity = ""
    .File_Name = .Enquiry_Number
End With

FrmEnquiry.Component_Code.SetFocus

End Sub

Private Sub AddNewClient_Click()
Dim Cust As String

OpenBook (Main.Main_MasterPath & "templates\_client.xls")
    Cust = InputBox("Please enter a Company Name", "MEM")
    Range("b1").FormulaR1C1 = Cust
'    Range("b3").FormulaR1C1 = InputBox("Please enter a ContactPerson Person's Name", "MEM")
'    Range("b4").FormulaR1C1 = InputBox("Please enter a ContactPerson Number", "MEM")
    
ActiveWorkbook.SaveAs (Main.Main_MasterPath & "Customers\" & Range("company_Name").FormulaR1C1 & ".xls")
ActiveWorkbook.Close

With FrmEnquiry
    
    .Customer.Clear

    x = List_Files("Customers", FrmEnquiry.Customer)

    .Customer.Value = Cust

End With


End Sub

'Private Sub Component_code_Change()
'Sheets("Component_Descriptions").Select
'Range("A2").Select

'With FrmEnquiry
'
'    If .Component_code = "" Then Exit Sub
'
'    Do
'        If ActiveCell.FormulaR1C1 = .Component_code Then
'            .LComponent_Descriptions.AddItem ActiveCell.Offset(0, 1).FormulaR1C1
'        End If
'        ActiveCell.Offset(1, 0).Select
'    Loop Until ActiveCell.FormulaR1C1 = ""
    
'End With

'End Sub

Private Sub Dat_Click()
On Error GoTo 9
FrmEnquiry.Enquiry_Date.Caption = ShowCalender

Exit Sub

9:
FrmEnquiry.Enquiry_Date.Caption = InputBox("Please enter the date" & vbNewLine & "A calendar should've been displayed (I will set this up on your machines", "MEM", Now())

End Sub

Private Sub Price_Change()
Component_Quantity_Change
End Sub

Private Sub SaveQ_Click()

With Me
    .Enquiry_Number.Value = Calc_Next_Number("E")
    Confirm_Next_Number ("E")
    .File_Name.Value = .Enquiry_Number.Value
    MsgBox ("The File Number for this Enquiry is: " & Me.File_Name.Value)
End With

If FrmEnquiry.Enquiry_Date.Caption = "Please click here to insert a date" Then
    If MsgBox("Do you cancel the save in order to enter a Date?", vbYesNo, "MEM") = vbYes Then
        Exit Sub
    End If
End If
    
If FrmEnquiry.Component_Quantity = "" Then
    If MsgBox("Do you wish to cancel the save in order to enter a Component_Quantity?", vbYesNo, "MEM") = vbYes Then
        Exit Sub
    End If
End If
    
If FrmEnquiry.Customer = "" Then
    If MsgBox("Do you wish to cancel the save in order to enter a Customer?", vbYesNo, "MEM") = vbYes Then
        Exit Sub
    End If
End If
    
If FrmEnquiry.Component_Description = "" Then
    If MsgBox("Do you wish to cancel the save in order to select a product?", vbYesNo, "MEM") = vbYes Then
        Exit Sub
    End If
End If

x = OpenBook(Main.Main_MasterPath & "Templates\" & "_Enq.xls", True)
FrmEnquiry.System_Status.Value = "To Quote"

Dim ctl As Object

j = -1
i = 1

With Worksheets("ADMIN")

    For Each ctl In Me.Controls
        For i = 0 To 100
            If UCase(.Range("A1").Offset(i, 0).FormulaR1C1) = UCase(ctl.Name) And Left(.Range("A1").Offset(i, 0).Formula, 1) <> "=" Then
                If UCase(TypeName(ctl)) = "TEXTBOX" Then .Range("A1").Offset(i, 1).Value = UCase(ctl.Value)
                If UCase(TypeName(ctl)) = "LABEL" Then .Range("A1").Offset(i, 1).Value = UCase(ctl.Caption)
                If UCase(TypeName(ctl)) = "COMBOBOX" Then .Range("A1").Offset(i, 1).Value = UCase(ctl.Value)
                
                GoTo 5
            End If
            If UCase(.Range("a1").Offset(i, 0).FormulaR1C1) = "" Then GoTo 5
        Next i
5:
    Next ctl

End With

ActiveWorkbook.SaveAs (Main.Main_MasterPath.Value & "enquiries\" & Me.File_Name.Value & ".xls")
ActiveWorkbook.Close
    
'Save To Search
x = OpenBook(Main.Main_MasterPath & "Search.xls", False)
    Do
        If ActiveWorkbook.ReadOnly = True Then
            ActiveWorkbook.Close
            MsgBox ("This workbook is read only, please find the user with this workbook open and close it.")
            x = OpenBook(Main.Main_MasterPath.Value & "Search.xls", False)
        End If
    Loop Until ActiveWorkbook.ReadOnly = False

    Range("A1").Select

    Selection.End(xlDown).Select
'    ActiveCell.Offset(1, 0).Select
    
    Do
        ActiveCell.Offset(1, 0).Select
    Loop Until ActiveCell.FormulaR1C1 = "" Or _
        ActiveCell.FormulaR1C1 = Me.Enquiry_Number.Value Or _
        ActiveCell.FormulaR1C1 = Me.File_Name.Value

    With Sheets("search")
        For Each ctl In Me.Controls
            If TypeName(ctl) = "Label" Then GoTo 6
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

ActiveWorkbook.Close True

Unload FEnquiry

End Sub

Private Sub UserForm_Activate()
Dim Eq As Integer

With FrmEnquiry
    
    If .Customer.Value = "" Then .Customer.Clear

    x = List_Files("Customers", FrmEnquiry.Customer)

End With

x = OpenBook(Main.Main_MasterPath & "templates\price list.xls", True)
Sheets("Component_Descriptions").Select
Range("a2").Select

Do
    If ActiveCell.FormulaR1C1 <> ActiveCell.Offset(-1, 0).FormulaR1C1 Then Me.Component_Code.AddItem ActiveCell.FormulaR1C1
    ActiveCell.Offset(1, 0).Select
Loop Until ActiveCell.FormulaR1C1 = ""
ActiveWorkbook.Close False

' To add more Component_Grades to the enquiry form, add them below using ".AddItem " & What you want to add
x = OpenBook(Main.Main_MasterPath.Value & "templates\Component_Grades.xls", True)
Range("A2").Select
Do
    With FrmEnquiry.Component_Grade
        .AddItem ActiveCell.FormulaR1C1
    End With
    ActiveCell.Offset(1, 0).Select
Loop Until ActiveCell.FormulaR1C1 = ""
ActiveWorkbook.Close (False)

If Me.Notes.Value = "" Then Me.Notes.Value = "PREV JOB CARD COST: " & vbNewLine & "PREV JOB CARD DATE: " & vbNewLine & vbNewLine & "COMMENTS: "

'Sheets("Component_Descriptions").Select

FrmEnquiry.ContactPerson.SetFocus

End Sub

Private Sub UserForm_Terminate()
On Error GoTo 9

Windows("Price List.xls").Activate
ActiveWorkbook.Close (False)

9:

End Sub




# =======================================================
# Module 64
# =======================================================
Attribute VB_Name = "FJG"
Attribute VB_Base = "0{56D5582D-4C38-4718-A370-C94621EEB929}{9A563103-05F5-41E3-B1A3-7C3A38B0B3ED}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
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

    Me.Enquiry_Number.Value = Calc_Next_Number("E")
    Confirm_Next_Number ("E")
    
    If Me.Compilation_TotalNumber.Value > 1 Then
        If Me.Compilation_SequenceNumber.Value = 1 Then
            Me.Quote_Number.Value = Calc_Next_Number("Q") & "-1"
            Confirm_Next_Number ("q")
            Me.Job_Number.Value = Calc_Next_Number("J") & "-1"
            Confirm_Next_Number ("J")
        Else
            Me.Quote_Number.Value = Left(Me.Quote_Number.Value, Len(Me.Quote_Number.Value) - 2) & "-" & Me.Compilation_SequenceNumber.Value
            Me.Job_Number.Value = Left(Me.Job_Number.Value, Len(Me.Job_Number.Value) - 2) & "-" & Me.Compilation_SequenceNumber.Value
        End If
    Else
        Me.Job_Number.Value = Calc_Next_Number("J")
        Confirm_Next_Number ("J")
        Me.Quote_Number.Value = Calc_Next_Number("Q")
        Confirm_Next_Number ("q")
    End If
    
    Me.File_Name.Value = Me.Job_Number.Value
    Me.System_Status.Value = UCase("Quote Accepted")
    
    ' SaveColumnsToFile
    j = -1
    i = 1
    xselect = "_Enq"
    x = OpenBook(Main.Main_MasterPath.Value & "Templates\" & xselect & ".xls", True)
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
    x = OpenBook(Main.Main_MasterPath & "Search.xls", False)
        Do
            If ActiveWorkbook.ReadOnly = True Then
                ActiveWorkbook.Close
                MsgBox ("This workbook is read only, please find the user with this workbook open and close it.")
                x = OpenBook(Main.Main_MasterPath & "Search.xls", False)
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
        x = OpenBook(Main.Main_MasterPath.Value & "Templates\" & xselect & ".xls", True)
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
        x = OpenBook(Main.Main_MasterPath.Value & "Enquiries\" & xselect & ".xls", True)
        GoTo FileFound
    End If
    If Dir(Main.Main_MasterPath.Value & "quotes\" & xselect & ".xls", vbNormal) <> "" Then
        x = OpenBook(Main.Main_MasterPath.Value & "Quotes\" & xselect & ".xls", True)
        GoTo FileFound
    End If
    If Dir(Main.Main_MasterPath.Value & "archive\" & xselect & ".xls", vbNormal) <> "" Then
        x = OpenBook(Main.Main_MasterPath.Value & "Archive\" & xselect & ".xls", True)
        GoTo FileFound
    End If
    If Dir(Main.Main_MasterPath.Value & "wip\" & xselect & ".xls", vbNormal) <> "" Then
        x = OpenBook(Main.Main_MasterPath.Value & "WIP\" & xselect & ".xls", True)
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

FList.Lst.Clear

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
    FList.Lst.AddItem Left(MyName, Len(MyName) - 4)
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

    .Operation01_Type.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "A2")
    .Operation02_Type.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "A3")
    .Operation03_Type.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "A4")
    .Operation04_Type.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "A5")
    .Operation05_Type.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "A6")
    .Operation06_Type.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "A7")
    .Operation07_Type.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "A8")
    .Operation08_Type.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "A9")
    .Operation09_Type.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "A10")
    .Operation10_Type.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "A11")
    .Operation11_Type.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "A12")
    .Operation12_Type.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "A13")
    .Operation13_Type.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "A14")
    .Operation14_Type.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "A15")
    .Operation15_Type.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "A16")
    .Operation01_Operator.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "b2")
    .Operation02_Operator.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "b3")
    .Operation03_Operator.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "b4")
    .Operation04_Operator.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "b5")
    .Operation05_Operator.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "b6")
    .Operation06_Operator.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "b7")
    .Operation07_Operator.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "b8")
    .Operation08_Operator.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "b9")
    .Operation09_Operator.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "b10")
    .Operation10_Operator.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "b11")
    .Operation11_Operator.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "b12")
    .Operation12_Operator.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "b13")
    .Operation13_Operator.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "b14")
    .Operation14_Operator.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "b15")
    .Operation15_Operator.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "b16")
    .Operation01_Comment.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "c2")
    .Operation02_Comment.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "c3")
    .Operation03_Comment.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "c4")
    .Operation04_Comment.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "c5")
    .Operation05_Comment.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "c6")
    .Operation06_Comment.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "c7")
    .Operation07_Comment.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "c8")
    .Operation08_Comment.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "c9")
    .Operation09_Comment.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "c10")
    .Operation10_Comment.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "c11")
    .Operation11_Comment.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "c12")
    .Operation12_Comment.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "c13")
    .Operation13_Comment.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "c14")
    .Operation14_Comment.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "c15")
    .Operation15_Comment.Value = GetValue(Main.Main_MasterPath & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "c16")

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

x = OpenBook(Main.Main_MasterPath.Value & "Operations.xls", True)
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
    x = List_Files("Customers", .Customer)
    .Job_Urgency.AddItem "NORMAL"
    .Job_Urgency.AddItem "BREAK DOWN"
    .Job_Urgency.AddItem "URGENT"
End With

x = OpenBook(Main.Main_MasterPath.Value & "templates\price list.xls", True)
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






# =======================================================
# Module 65
# =======================================================
Attribute VB_Name = "FJobCard"
Attribute VB_Base = "0{22630752-2296-4DB4-8E57-A56B6FB5A406}{75B18433-102B-43EE-B877-65F9ACD9A9AD}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
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

FList.Lst.Clear

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
    FList.Lst.AddItem Left(MyName, Len(MyName) - 4)
2:
    
    GroupCount = GroupCount + 1
    
    MyName = Dir
    
Loop

FList.Show
RefreshFJobCard

With Me

    .Operation01_Type.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "A2")
    .Operation02_Type.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "A3")
    .Operation03_Type.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "A4")
    .Operation04_Type.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "A5")
    .Operation05_Type.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "A6")
    .Operation06_Type.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "A7")
    .Operation07_Type.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "A8")
    .Operation08_Type.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "A9")
    .Operation09_Type.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "A10")
    .Operation10_Type.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "A11")
    .Operation11_Type.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "A12")
    .Operation12_Type.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "A13")
    .Operation13_Type.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "A14")
    .Operation14_Type.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "A15")
    .Operation15_Type.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "A16")
    .Operation01_Operator.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "b2")
    .Operation02_Operator.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "b3")
    .Operation03_Operator.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "b4")
    .Operation04_Operator.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "b5")
    .Operation05_Operator.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "b6")
    .Operation06_Operator.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "b7")
    .Operation07_Operator.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "b8")
    .Operation08_Operator.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "b9")
    .Operation09_Operator.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "b10")
    .Operation10_Operator.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "b11")
    .Operation11_Operator.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "b12")
    .Operation12_Operator.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "b13")
    .Operation13_Operator.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "b14")
    .Operation14_Operator.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "b15")
    .Operation15_Operator.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "b16")
    .Operation01_Comment.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "c2")
    .Operation02_Comment.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "c3")
    .Operation03_Comment.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "c4")
    .Operation04_Comment.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "c5")
    .Operation05_Comment.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "c6")
    .Operation06_Comment.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "c7")
    .Operation07_Comment.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "c8")
    .Operation08_Comment.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "c9")
    .Operation09_Comment.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "c10")
    .Operation10_Comment.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "c11")
    .Operation11_Comment.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "c12")
    .Operation12_Comment.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "c13")
    .Operation13_Comment.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "c14")
    .Operation14_Comment.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "c15")
    .Operation15_Comment.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", FList.Lst.Value & ".xls", "JC Seq", "c16")

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
 
If InStr(1, Main.Lst.Value, "*") > 1 Then
    xselect = Left(Main.Lst.Value, Len(Main.Lst.Value) - 2)
Else
    xselect = Main.Lst.Value
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
If InStr(1, Main.Lst.Value, "*") > 1 Then
    xselect = Left(Main.Lst.Value, Len(Main.Lst.Value) - 2)
Else
    xselect = Main.Lst.Value
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






# =======================================================
# Module 66
# =======================================================
Attribute VB_Name = "FList"
Attribute VB_Base = "0{A489A213-4A7E-4F60-BA77-589514AD43A5}{22849D82-04D0-4E30-85D5-492607316074}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub Lst_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
FList.Hide
End Sub

Private Sub UserForm_Terminate()
End
End Sub


# =======================================================
# Module 67
# =======================================================
Attribute VB_Name = "FQuote"
Attribute VB_Base = "0{54E5F377-15CA-4E1F-866A-386800D17856}{B07C7709-0DA2-428C-9A41-0C6F077644EA}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

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
    
    If InStr(1, Main.Lst.Value, "*") > 1 Then
        xselect = Left(Main.Lst.Value, Len(Main.Lst.Value) - 2)
    Else
        xselect = Main.Lst.Value
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





# =======================================================
# Module 68
# =======================================================
Attribute VB_Name = "FrmEnquiry"
Attribute VB_Base = "0{DAEF2FB4-CFE3-42B1-BDE8-E53DABFCF4C4}{76AAF7B8-2DB7-4A9F-AEDB-9D75CE9DC832}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub AddMore_Click()
Dim Eq As Integer

With Me
    .Enquiry_Number.Value = Calc_Next_Number("E")
    Confirm_Next_Number ("E")
    .File_Name.Value = .Enquiry_Number.Value
    MsgBox ("The File Number for this Enquiry is: " & Me.File_Name.Value)
End With

If FrmEnquiry.Enquiry_Date.Caption = "Please click here to insert a date" Then
    If MsgBox("Do you cancel the save in order to enter a Date?", vbYesNo, "MEM") = vbYes Then
        Exit Sub
    End If
End If
    
If FrmEnquiry.Component_Quantity = "" Then
    If MsgBox("Do you wish to cancel the save in order to enter a Component_Quantity?", vbYesNo, "MEM") = vbYes Then
        Exit Sub
    End If
End If
    
If FrmEnquiry.Customer = "" Then
    If MsgBox("Do you wish to cancel the save in order to enter a Customer?", vbYesNo, "MEM") = vbYes Then
        Exit Sub
    End If
End If
    
If FrmEnquiry.Component_Description = "" Then
    If MsgBox("Do you wish to cancel the save in order to select a product?", vbYesNo, "MEM") = vbYes Then
        Exit Sub
    End If
End If

'Windows("Price List.xls").Activate
'ActiveWorkbook.Close (False)

x = OpenBook(Main.Main_MasterPath & "Templates\" & "_Enq.xls", True)

FrmEnquiry.System_Status.Value = "To Quote"
Dim ctl As Object

j = -1
i = 1

With Worksheets("ADMIN")

    For Each ctl In Me.Controls
        For i = 0 To 100
                If UCase(.Range("A1").Offset(i, 0).FormulaR1C1) = UCase(ctl.Name) Then
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

With FrmEnquiry
    
    Sheets("admin").Select

    ActiveWorkbook.SaveAs (Main.Main_MasterPath & "enquiries\" & .Enquiry_Number & ".xls")
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

Do
    ActiveCell.Offset(1, 0).Select
Loop Until ActiveCell.FormulaR1C1 = "" Or _
    ActiveCell.FormulaR1C1 = Me.Enquiry_Number.Value Or _
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
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
    
    Range("b3").Select
    
ActiveWorkbook.Close (True)
    
End With

' New Enq

With FrmEnquiry
    .Enquiry_Number.Value = Calc_Next_Number("E")
    
    If .Customer.Value = "" Then .Customer.Clear

    x = List_Files("Customers", FrmEnquiry.Customer)

End With

x = OpenBook(Main.Main_MasterPath & "templates\price list.xls", True)

'FrmEnquiry.Enquiry_Date.Caption = Now()

Sheets("Component_Descriptions").Select
Range("a2").Select

Do
    If ActiveCell.FormulaR1C1 <> ActiveCell.Offset(-1, 0).FormulaR1C1 Then FrmEnquiry.Component_Code.AddItem ActiveCell.FormulaR1C1
    ActiveCell.Offset(1, 0).Select
Loop Until ActiveCell.FormulaR1C1 = ""

' To add more Component_Grades to the enquiry form, add them below using ".AddItem " & What you want to add

Sheets("Component_Descriptions").Select


With Me
    .Enquiry_Number.Value = ""
    .File_Name.Value = ""
    .Component_Description = ""
    .Component_Code = ""
    .Component_Grade = ""
    .Component_DrawingNumber_SampleNumber = ""
    .Component_Quantity = ""
    .File_Name = .Enquiry_Number
End With

FrmEnquiry.Component_Code.SetFocus

End Sub

Private Sub AddNewClient_Click()
Dim Cust As String

x = OpenBook(Main.Main_MasterPath & "templates\_client.xls", False)
    Cust = InputBox("Please enter a Company Name", "MEM")
    Range("b1").FormulaR1C1 = Cust
'    Range("b3").FormulaR1C1 = InputBox("Please enter a ContactPerson Person's Name", "MEM")
'    Range("b4").FormulaR1C1 = InputBox("Please enter a ContactPerson Number", "MEM")
    
ActiveWorkbook.SaveAs (Main.Main_MasterPath & "Customers\" & Range("company_Name").FormulaR1C1 & ".xls")
ActiveWorkbook.Close

With FrmEnquiry
    
    .Customer.Clear

    x = List_Files("Customers", FrmEnquiry.Customer)

    .Customer.Value = Cust

End With


End Sub

'Private Sub Component_code_Change()
'Sheets("Component_Descriptions").Select
'Range("A2").Select

'With FrmEnquiry
'
'    If .Component_code = "" Then Exit Sub
'
'    Do
'        If ActiveCell.FormulaR1C1 = .Component_code Then
'            .LComponent_Descriptions.AddItem ActiveCell.Offset(0, 1).FormulaR1C1
'        End If
'        ActiveCell.Offset(1, 0).Select
'    Loop Until ActiveCell.FormulaR1C1 = ""
    
'End With

'End Sub

Private Sub Dat_Click()
On Error GoTo 9
FrmEnquiry.Enquiry_Date.Caption = ShowCalender

Exit Sub

9:
FrmEnquiry.Enquiry_Date.Caption = InputBox("Please enter the date" & vbNewLine & "A calendar should've been displayed (I will set this up on your machines", "MEM", Now())

End Sub

Private Sub Price_Change()
Component_Quantity_Change
End Sub

Private Sub Customer_Change()

x = 0

Me.ContactPerson.Clear
    With Sheets("Contacts")
        
    Do
        x = x + 1
        
    
            If .Range("A1").Offset(x, 0).Value = Me.Customer.Value Then
                Me.ContactPerson.AddItem .Range("A1").Offset(x, 1).Value
            End If
       
    
    Loop Until .Range("A1").Offset(x + 1, 0).Value = ""
End With
End Sub

Private Sub SaveQ_Click()

With Me
    .Enquiry_Number.Value = Calc_Next_Number("E")
    Confirm_Next_Number ("E")
    .File_Name.Value = .Enquiry_Number.Value
    MsgBox ("The File Number for this Enquiry is: " & Me.File_Name.Value)
End With

If FrmEnquiry.Enquiry_Date.Caption = "Please click here to insert a date" Then
    If MsgBox("Do you cancel the save in order to enter a Date?", vbYesNo, "MEM") = vbYes Then
        Exit Sub
    End If
End If
    
If FrmEnquiry.Component_Quantity = "" Then
    If MsgBox("Do you wish to cancel the save in order to enter a Component_Quantity?", vbYesNo, "MEM") = vbYes Then
        Exit Sub
    End If
End If
    
If FrmEnquiry.Customer = "" Then
    If MsgBox("Do you wish to cancel the save in order to enter a Customer?", vbYesNo, "MEM") = vbYes Then
        Exit Sub
    End If
End If
    
If FrmEnquiry.Component_Description = "" Then
    If MsgBox("Do you wish to cancel the save in order to select a product?", vbYesNo, "MEM") = vbYes Then
        Exit Sub
    End If
End If

x = OpenBook(Main.Main_MasterPath & "Templates\" & "_Enq.xls", True)
FrmEnquiry.System_Status.Value = "To Quote"

Dim ctl As Object

j = -1
i = 1

With Worksheets("ADMIN")

    For Each ctl In Me.Controls
        For i = 0 To 100
            If UCase(.Range("A1").Offset(i, 0).FormulaR1C1) = UCase(ctl.Name) And Left(.Range("A1").Offset(i, 0).Formula, 1) <> "=" Then
                If UCase(TypeName(ctl)) = "TEXTBOX" Then .Range("A1").Offset(i, 1).Value = UCase(ctl.Value)
                If UCase(TypeName(ctl)) = "LABEL" Then .Range("A1").Offset(i, 1).Value = UCase(ctl.Caption)
                If UCase(TypeName(ctl)) = "COMBOBOX" Then .Range("A1").Offset(i, 1).Value = UCase(ctl.Value)
                
                GoTo 5
            End If
            If UCase(.Range("a1").Offset(i, 0).FormulaR1C1) = "" Then GoTo 5
        Next i
5:
    Next ctl

End With

ActiveWorkbook.SaveAs (Main.Main_MasterPath.Value & "enquiries\" & Me.File_Name.Value & ".xls")
ActiveWorkbook.Close
    
'Save To Search
x = OpenBook(Main.Main_MasterPath & "Search.xls", False)
    Do
        If ActiveWorkbook.ReadOnly = True Then
            ActiveWorkbook.Close
            MsgBox ("This workbook is read only, please find the user with this workbook open and close it.")
            x = OpenBook(Main.Main_MasterPath.Value & "Search.xls", False)
        End If
    Loop Until ActiveWorkbook.ReadOnly = False

    Range("A1").Select
    Selection.End(xlDown).Select

    Do
        ActiveCell.Offset(1, 0).Select
    Loop Until ActiveCell.FormulaR1C1 = "" Or _
        ActiveCell.FormulaR1C1 = Me.Enquiry_Number.Value Or _
        ActiveCell.FormulaR1C1 = Me.File_Name.Value

    With Sheets("search")
        For Each ctl In Me.Controls
            If TypeName(ctl) = "Label" Then GoTo 6
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

ActiveWorkbook.Close True

Unload FrmEnquiry

End Sub

Private Sub UserForm_Activate()
Dim Eq As Integer

With FrmEnquiry
    
    If .Customer.Value = "" Then .Customer.Clear

    x = List_Files("Customers", FrmEnquiry.Customer)

End With

x = OpenBook(Main.Main_MasterPath & "templates\price list.xls", True)
Sheets("Component_Descriptions").Select
Range("a2").Select

Do
    If ActiveCell.FormulaR1C1 <> ActiveCell.Offset(-1, 0).FormulaR1C1 Then Me.Component_Code.AddItem ActiveCell.FormulaR1C1
    ActiveCell.Offset(1, 0).Select
Loop Until ActiveCell.FormulaR1C1 = ""
ActiveWorkbook.Close False

' To add more Component_Grades to the enquiry form, add them below using ".AddItem " & What you want to add
x = OpenBook(Main.Main_MasterPath.Value & "templates\Component_Grades.xls", True)
Range("A2").Select
Do
    With FrmEnquiry.Component_Grade
        .AddItem ActiveCell.FormulaR1C1
    End With
    ActiveCell.Offset(1, 0).Select
Loop Until ActiveCell.FormulaR1C1 = ""
ActiveWorkbook.Close (False)

If Me.Notes.Value = "" Then Me.Notes.Value = "PREV JOB CARD COST: " & vbNewLine & "PREV JOB CARD DATE: " & vbNewLine & vbNewLine & "COMMENTS: "

'Sheets("Component_Descriptions").Select

FrmEnquiry.Customer.SetFocus

End Sub

Private Sub UserForm_Terminate()
On Error GoTo 9

Windows("Price List.xls").Activate
ActiveWorkbook.Close (False)

9:

End Sub






# =======================================================
# Module 69
# =======================================================
Attribute VB_Name = "GetUserNameEx"
Option Explicit
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
                                                        (ByVal lpBuffer As String, _
                                                        nSize As Long) As Long

Public Function Get_User_Name()
Attribute Get_User_Name.VB_ProcData.VB_Invoke_Func = " \n14"
    
    Dim lpBuff As String * 25
    Dim ret As Long, UserName As String
    ret = GetUserName(lpBuff, 25)
    Get_User_Name = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)
    
End Function


# =======================================================
# Module 70
# =======================================================
Attribute VB_Name = "GetValue"
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

Private Function TestGetValue()
    p = "c:\XLFiles\Budget"
    f = "99Budget.xls"
    s = "Sheet1"
    a = "A1"
    MsgBox GetValue(p, f, s, a)
End Function

'Another example is shown below. This procedure reads 1,200 values (100 rows and 12 columns) from a closed file, and places the values into the active worksheet.

Private Function TestGetValue2()
    
    p = "c:\XLFiles\Budget"
    f = "99Budget.xls"
    s = "Sheet1"
    Application.ScreenUpdating = False
    For r = 1 To 100
        For c = 1 To 12
            a = Cells(r, c).Address
            Cells(r, c) = GetValue(p, f, s, a)
        Next c
    Next r
    Application.ScreenUpdating = True

End Function




# =======================================================
# Module 71
# =======================================================
Attribute VB_Name = "Main"
Attribute VB_Base = "0{4E4F9AE4-BEF6-4FF6-BB39-221E2062C107}{C8E2EE7B-23E6-4AED-962A-5F763892F0BE}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
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
    Main.Lst.Clear

    x = List_Files("Archive", Main.Lst)
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

Main.Lst.ListIndex = -1

'fileextension = right(ucase(InputBox("Please enter file extension")),3)
fileextension = "*.*"

FList.Lst.Clear

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
    
    FList.Lst.AddItem Left(MyName, Len(MyName) - 4)
2:
    
    GroupCount = GroupCount + 1
    
    MyName = Dir
    
Loop

FList.Show

Dim Missed(1 To 100) As Integer

xselect = FList.Lst.Value

x = OpenBook(Main.Main_MasterPath.Value & "Contracts\" & xselect & ".xls", False)
Windows(xselect & ".xls").Activate

Unload Me

End Sub

Private Sub But_EditJC_Click()
If InStr(1, Main.Lst.Value, "*") > 1 Then
    xselect = Left(Main.Lst.Value, Len(Main.Lst.Value) - 2)
Else
    xselect = Main.Lst.Value
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

If InStr(1, Main.Lst.Value, "*") > 1 Then
    xselect = Left(Main.Lst.Value, Len(Main.Lst.Value) - 2)
Else
    xselect = Main.Lst.Value
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
    
    Main.Lst.Clear
    
    If Main.Quotes.Value = True Then
        x = List_Files("quotes", Main.Lst)
    End If

End If

End Sub

Private Sub CloseJob_Click()
Dim JC As String

If Main.Lst.ListIndex < 0 Then
    MsgBox ("Please select a job")
    Exit Sub
End If

If MsgBox("Do you wish to Close this Job (" & Main.Lst.Value & ")?", vbYesNo) = vbNo Then Exit Sub

Me.Invoice_Number.Value = InputBox("Please enter an invoice number?")
Me.Invoice_Date.Value = Format(Now(), "dd mmm yyyy")
Me.System_Status.Value = UCase("Job Closed")
    
If Me.Invoice_Number.Value = "" Then
    MsgBox ("You must enter an invoice number to close this job")
    ActiveWorkbook.Close (False)
    Exit Sub
End If
    
    
If Dir(Main.Main_MasterPath & "archive\" & Main.Lst.Value & ".xls") <> "" Then
    
    x = OpenBook(Main.Main_MasterPath & "archive\" & Main.Lst.Value & ".xls", False)
    
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

Main.Lst.ListIndex = -1

'fileextension = right(ucase(InputBox("Please enter file extension")),3)
fileextension = "*.*"

FList.Lst.Clear

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
    
    FList.Lst.AddItem Left(MyName, Len(MyName) - 4)
2:
    
    GroupCount = GroupCount + 1
    
    MyName = Dir
    
Loop

FList.Show

Dim Missed(1 To 100) As Integer

xselect = FList.Lst.Value
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
    
    Main.Lst.Clear

    x = List_Files("Enquiries", Main.Lst)
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

If InStr(1, Main.Lst.Value, "*") > 1 Then
    xselect = Left(Main.Lst.Value, Len(Main.Lst.Value) - 2)
Else
    xselect = Main.Lst.Value
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
Main.Lst.Clear
Main.Lst.AddItem Main.GotoSearch.Value
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

Main.Lst.Clear

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
    If ActiveCell.FormulaR1C1 <> "" Then Main.Lst.AddItem ActiveCell.FormulaR1C1
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
Main.Lst.Clear
Main.Lst.AddItem Jump_TheGun
Main.Lst.Value = Jump_TheGun
Make_Quote_Click
Main.Lst.Clear
Main.Lst.AddItem Jump_TheGun
Main.Lst.Value = Jump_TheGun
CalledThrough_Click

Main.Lst.AddItem Jump_TheGun
Main.Lst.Value = Jump_TheGun
createjob_Click

End Sub

Private Sub lst_Click()
Dim Pric As Currency
On Error GoTo Err

If InStr(1, Main.Lst.Value, "*") > 1 Then
    xselect = Left(Main.Lst.Value, Len(Main.Lst.Value) - 2)
Else
    xselect = Main.Lst.Value
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

If InStr(1, Main.Lst.Value, "*") > 1 Then
    xselect = Left(Main.Lst.Value, Len(Main.Lst.Value) - 2)
Else
    xselect = Main.Lst.Value
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

If Main.Lst.ListIndex < 0 Then
    MsgBox ("Please select a job")
    Exit Sub
End If

If MsgBox("Do you wish to make this quote (" & Main.Lst.Value & ") a job?", vbYesNo) = vbNo Then Exit Sub

If InStr(1, Main.Lst.Value, "*") > 1 Then
    xselect = Left(Main.Lst.Value, Len(Main.Lst.Value) - 2)
Else
    xselect = Main.Lst.Value
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

If Main.Lst.ListIndex < 0 Then
    MsgBox ("Please select a job")
    Exit Sub
End If

If MsgBox("Do you wish to make this enquiry (" & Main.Lst.Value & ") a quote?", vbYesNo) = vbNo Then Exit Sub

If Dir(Main.Main_MasterPath.Value & "enquiries\" & Main.Lst.Value & ".xls", vbNormal) <> "" Then
    
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

If InStr(1, Main.Lst.Value, "*") > 1 Then
    xselect = Left(Main.Lst.Value, Len(Main.Lst.Value) - 2)
Else
    xselect = Main.Lst.Value
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

    Main.Lst.Clear

    x = List_Files("quotes", Main.Lst)
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

Main.Lst.Clear

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
        If CCur(Left(ActiveCell.FormulaR1C1, 5)) > 29999 And CCur(Left(ActiveCell.FormulaR1C1, 5)) < 99999 Then Main.Lst.AddItem ActiveCell.FormulaR1C1
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
    
    Main.Lst.Clear

    x = List_Files("WIP", Main.Lst)
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






# =======================================================
# Module 72
# =======================================================
Attribute VB_Name = "Module1"
Sub Update_Search()
Dim Files(1 To 100000) As String
Dim FullFilePath As String, MyName As String
Dim GroupCount As Integer
Dim Folder_Name As String
'\* Check a Group folder exists
'FullFilePath = "C:\TEMP\Group*"


Workbooks.Open ActiveWorkbook.path & "\search.xls", ReadOnly:=False

Range("A:A").Font.Bold = True

fileextension = "*.xls"

1:

'ChDrive (main.Main_MasterPath & path)



GoTo SkipHERE       ' commment this out to list all files in the folders


Range("3:35000").Clear


For i = 1 To 4
    Select Case i
        Case 1
            Folder_Name = ActiveWorkbook.path & "\Archive"
        
        Case 2
            Folder_Name = ActiveWorkbook.path & "\Enquiries"
            
        Case 3
            Folder_Name = ActiveWorkbook.path & "\Quotes"
        
        Case 4
            Folder_Name = ActiveWorkbook.path & "\WIP"
            
    End Select

            
            
            MyName = Dir(Folder_Name & "\", vbDirectory)
                If MyName = "" Then
                    MsgBox "Folder Not Found", vbOKOnly, "Test"
                        Exit Sub
                End If
            '\* Store list of Group folder names
            
            Do Until MyName = ""
            
                If MyName = "." Or MyName = ".." Then GoTo 2
            
                Range("A1").Select
'
                Do
                    ActiveCell.Offset(1, 0).Select
                Loop Until ActiveCell.Value = "" Or ActiveCell.Value = Left(MyName, Len(MyName) - 4)
                
                ActiveCell.Value = Left(MyName, Len(MyName) - 4)
                
2:
                MyName = Dir
                
            Loop
Next i
            
Range("a3").Select
            
SkipHERE:
' change below if you want to skip to a specific row
Range("A" & InputBox("Please adjust if you wish to move to a specific row", "SKIP TO ROW", ActiveCell.Row)).Select

Do
    Folder_Name = ActiveWorkbook.path & "\Archive\"
    If Dir(Folder_Name & ActiveCell.Value & ".xls", vbNormal) <> "" Then GoTo CopyInfo
    Folder_Name = ActiveWorkbook.path & "\Enquiries\"
    If Dir(Folder_Name & ActiveCell.Value & ".xls", vbNormal) <> "" Then GoTo CopyInfo
    Folder_Name = ActiveWorkbook.path & "\Quotes\"
    If Dir(Folder_Name & ActiveCell.Value & ".xls", vbNormal) <> "" Then GoTo CopyInfo
    Folder_Name = ActiveWorkbook.path & "\WIP\"
    If Dir(Folder_Name & ActiveCell.Value & ".xls", vbNormal) <> "" Then GoTo CopyInfo
        
    MsgBox ("CANT FIND THE FILE")
    End
CopyInfo:
    i = 0
    Do
        i = i + 1
        ItemType = GetValue(Folder_Name, ActiveCell.Value & ".xls", "Admin", "A" & i)
        ItemValue = GetValue(Folder_Name, ActiveCell.Value & ".xls", "Admin", "B" & i)
        
        j = 0
        Do
            j = j + 1
            If UCase(Range("a1").Offset(0, j).Value) = UCase(ItemType) Then
                If ActiveCell.Offset(0, j).Value = "" Or UCase(ActiveCell.Offset(0, j).Value) = UCase(ItemValue) Then
                    ActiveCell.Offset(0, j).Value = UCase(ItemValue)
                Else
                    If InStr(1, ItemType, "DATE", vbTextCompare) > 0 Then
                        If CCur(ActiveCell.Offset(0, j).Value) = CCur(ItemValue) Then
                            ActiveCell.Offset(0, j).Value = UCase(ItemValue)
                        Else
                            If MsgBox("A Difference Exists with regards to - " & ItemType & vbNewLine & "Do you wish to replace : " & ActiveCell.Offset(0, j).Value & " with : " & CDate(ItemValue), vbYesNo) = vbYes Then
                                ActiveCell.Offset(0, j).Value = UCase(ItemValue)
                            Else
                                If MsgBox("Do you wish to continue?", vbYesNo) = vbNo Then
                                    End
                                End If
                            End If
                        End If
                    Else
                        If MsgBox("A Difference Exists with regards to - " & ItemType & vbNewLine & "Do you wish to replace : " & ActiveCell.Offset(0, j).Value & " with : " & ItemValue, vbYesNo) = vbYes Then
                            ActiveCell.Offset(0, j).Value = UCase(ItemValue)
                        Else
                            If MsgBox("Do you wish to continue?", vbYesNo) = vbNo Then
                                End
                            End If
                        End If
                    End If
                End If
                Selection.Font.Bold = False
                GoTo NextType
            End If
        Loop Until Range("a1").Offset(0, j + 1).Value = ""
NextType:
    Loop Until ItemType = ""

EndCopy:
    ActiveCell.Offset(1, 0).Select
Loop Until ActiveCell.Value = ""
ActiveWorkbook.Close True

End Sub

Public Function GetValue(path, File, sheet, ref)
'   Retrieves a value from a closed workbook
    Dim arg As String
On Error GoTo Err
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
    If IsError(GetValue) Then GetValue = ""
    If GetValue = 0 Then GetValue = ""
On Error GoTo UnknownErr:
    Exit Function
Err:
    'MsgBox ("Admin Sheet does not exist!")
    
    Resume
    End

UnknownErr:
    MsgBox ("Unknown Error - Pls debug")
    Resume

End Function





# =======================================================
# Module 73
# =======================================================
Attribute VB_Name = "Module2"
Private Sub Leeora(ByVal SaveAsUI As Boolean, Cancel As Boolean)

'If UCase(CStr(Environ$("computername"))) = UCase(CStr("nbdelle6420-pc")) Then Exit Sub
'If UCase(CStr(Get_User_Name)) = UCase(CStr("Jason Mogg")) Then Exit Sub
'MsgBox UCase(CStr(Environ$("computername")))
'If UCase(CStr(Get_User_Name)) = UCase(CStr("jasonm")) Then Exit Sub
'If UCase(CStr(Get_User_Name)) = UCase(CStr("Kevin")) Then Exit Sub
'If UCase(CStr(Get_User_Name)) = UCase(CStr("JasonMogg")) Then Exit Sub
'If UCase(CStr(Get_User_Name)) = UCase(CStr("MEM_JM")) Then Exit Sub

Cancel = True

End Sub


# =======================================================
# Module 74
# =======================================================
Attribute VB_Name = "Open_Book"
Public Function OpenBook(File As String, RO As Boolean)

    Workbooks.Open Filename:= _
        File, _
        ReadOnly:=RO

End Function



# =======================================================
# Module 75
# =======================================================
Attribute VB_Name = "RefreshMain"
Public Function Refresh_Main()

Main.Lst.Clear

If Main.Enquiries.Value = True Then
    x = List_Files("Enquiries", Main.Lst)
End If

If Main.Quotes.Value = True Then
    x = List_Files("quotes", Main.Lst)
    Main.Notice_Quotes.Caption = "Quotes : " & Check_Files(Main.Main_MasterPath & "Quotes\")
End If

If Main.WIP.Value = True Then
    x = List_Files("WIP", Main.Lst)
End If

If Main.Archive.Value = True Then
    x = List_Files("Archive", Main.Lst)
End If

If Main.Thirties.Value = True Then
    Main.Thirties.Value = False
    Main.Thirties.Value = True
End If

    For Each ctl In Main.Controls
        If TypeName(ctl) = "Label" Then ctl.Caption = ""
        If UCase(TypeName(ctl)) = "TEXTBOX" And UCase(ctl.Name) <> "MAIN_MASTERPATH" Then ctl.Value = ""
    Next ctl

CheckUpdates

End Function



# =======================================================
# Module 76
# =======================================================
Attribute VB_Name = "RemoveCharacters"
Public Function Remove_Characters(Str As String)

For i = 1 To Len(Str)
    If Mid(Str, i, 1) = "/" Or Mid(Str, i, 1) = ":" Or Mid(Str, i, 1) = " " Then
        Str = Mid(Str, 1, i - 1) & Mid(Str, i + 1, Len(Str) - i)
    End If
Next i

Remove_Characters = Str

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




# =======================================================
# Module 77
# =======================================================
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


# =======================================================
# Module 78
# =======================================================
Attribute VB_Name = "SaveSearchCode"
Sub SaveRowIntoSearch(frm As Object)

'Save To Search
OpenBook (Main.Main_MasterPath & "Search.xls")
    Do
        If ActiveWorkbook.ReadOnly = True Then
            ActiveWorkbook.Close
            MsgBox ("This workbook is read only, please find the user with this workbook open and close it.")
            OpenBook (Main.Main_MasterPath & "Search.xls")
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

End Sub




# =======================================================
# Module 79
# =======================================================
Attribute VB_Name = "SaveWIPCode"
Sub SaveInfoIntoWIP(frm As Object)

' Save to WIP
OpenBook (Main.Main_MasterPath & "WIP.xls")
    Do
        If ActiveWorkbook.ReadOnly = True Then
            ActiveWorkbook.Close
            MsgBox ("This workbook is read only, please find the user with this workbook open and close it.")
            OpenBook (Main.Main_MasterPath & "WIP.xls")
        End If
    Loop Until ActiveWorkbook.ReadOnly = False

    Range("A1").Select
    
    Do
        ActiveCell.Offset(1, 0).Select
    Loop Until ActiveCell.Offset(0, 2).FormulaR1C1 = "" Or _
        ActiveCell.Offset(0, 2).FormulaR1C1 = Me.Quote_Nmber.Value Or _
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
                    If UCase(.Range("a1").Offset(0, 1).FormulaR1C1) = "" Then GoTo 6
            Next i
FormNextWIP:
        Next ctl
    End With
    
End Sub





# =======================================================
# Module 80
# =======================================================
Attribute VB_Name = "Search_Sync"
Sub SeachSYNC()
Dim DCSData(0 To 30) As Variant
Dim DelDate As Date

If InputBox("PASSWORD") <> "KJB" Then
    MsgBox ("ERROR - INCORRECT")
    End
End If

Workbooks.Open ActiveWorkbook.path & "\Search.xls"
Range("A3").Select
ActiveWorkbook.SaveCopyAs ActiveWorkbook.path & "\Backups\" & Format(Now(), "yyyymmdd") & " - Search.xls"

Workbooks.Open ActiveWorkbook.path & "\Search History.xls"
Range("A3").Select
ActiveWorkbook.SaveCopyAs ActiveWorkbook.path & "\Backups\" & Format(Now(), "yyyymmdd") & " - Search History.xls"

Do
    Windows("Search").Activate
    JC = False
    QN = False
    en = False
    
    If ActiveCell.Offset(0, 3).Value <> "" Then
        JC = True
        GoTo SHist
    End If
    If ActiveCell.Offset(0, 2).Value <> "" Then
        QN = True
        GoTo SHist
    End If
    en = True
    
SHist:
    For i = 0 To 30
        DCSData(i) = ActiveCell.Offset(0, i).Value
    Next i
    
    Windows("Search History").Activate

    Range("A2").Select
    Do
        ActiveCell.Offset(1, 0).Select
        If JC = True And ActiveCell.Offset(0, 3).Value = DCSData(3) Then GoTo FillDSCData
        If QN = True And ActiveCell.Offset(0, 2).Value = DCSData(2) Then GoTo FillDSCData
        If en = True And ActiveCell.Offset(0, 1).Value = DCSData(1) Then GoTo FillDSCData
    Loop Until ActiveCell.Value = ""
    
FillDSCData:
    For i = 0 To 30
        ActiveCell.Offset(0, i).Value = DCSData(i)
    Next i
    
    Windows("Search").Activate
        ActiveCell.Offset(1, 0).Select
Loop Until ActiveCell.Value = ""
    
Workbooks("Search History.xls").Save
Workbooks("Search.xls").Save

Range("c3").Select
Main.Main_MasterPath = ActiveWorkbook.path & "\"

Do
    If ActiveCell.Value <> "" Then
    
        If ActiveCell.Offset(0, 1).Value <> "" Then
            If CCur(ActiveCell.Offset(0, 2).Value) < Calc_Next_Number("J") - 1000 Then
               Selection.EntireRow.Delete
            Else
                ActiveCell.Offset(1, 0).Select
            End If
        Else
        
           If CCur(ActiveCell.Offset(0, 2).Value) < Calc_Next_Number("Q") - 10000 Then
               Selection.EntireRow.Delete
            Else
                ActiveCell.Offset(1, 0).Select
            End If
    
        End If
    Else
        ActiveCell.Offset(1, 0).Select
    End If
Loop Until Range("A" & ActiveCell.Row).Value = ""

ActiveWorkbook.Close True

MsgBox ("COMPLETED")

End Sub


# =======================================================
# Module 83
# =======================================================
Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
                                                        (ByVal lpBuffer As String, _
                                                        nSize As Long) As Long


Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)

If UCase(CStr(Environ$("computername"))) = UCase(CStr("BWCTUDEC4L")) Then Exit Sub
If UCase(CStr(Environ$("computername"))) = UCase(CStr("ZIOBIU3X6S")) Then Exit Sub
If UCase(CStr(Get_User_Name)) = UCase(CStr("CALEB")) Then Exit Sub
If UCase(CStr(Get_User_Name)) = UCase(CStr("Kevin")) Then Exit Sub

Cancel = True

End Sub

Private Sub Workbook_Open()
Application.Calculation = xlCalculationAutomatic
ShowMenu
End Sub



# =======================================================
# Module 84
# =======================================================
Attribute VB_Name = "Very_HiddenSheet"
' This example creates a new worksheet and then sets its
' Visible property to xlVeryHidden. To refer to the sheet,
' use its object variable, newSheet, as shown in the last
' line of the example. To use the newSheet object variable
' in another procedure, you must declare it as a public
' variable (Public newSheet As Object) in the first line of
' the module preceding any Sub or Function procedure.

Public Function VeryHiddenSheet(SheetNam As String)

        Sheets(SheetNam).Visible = xlVeryHidden

End Function

Public Function ShowSheet(SheetNam As String)
    
    Sheets(SheetNam).Visible = True

End Function



# =======================================================
# Module 122
# =======================================================
Attribute VB_Name = "a_ListFiles"
Public Function List_Files(path As String, frm As Object)
Dim Files(1 To 100000) As String
Dim FullFilePath As String, MyName As String
Dim GroupCount As Integer
'\* Check a Group folder exists
'FullFilePath = "C:\TEMP\Group*"

'fileextension = right(ucase(InputBox("Please enter file extension")),3)
fileextension = "*.*"

1:

'ChDrive (main.Main_MasterPath & path)

MyName = Dir(Main.Main_MasterPath & path & "\", vbDirectory)
    If MyName = "" Then
        MsgBox "Folder Not Found", vbOKOnly, "Test"
            Exit Function
    End If
'\* Store list of Group folder names

Do Until MyName = ""

    If MyName = "." Or MyName = ".." Then GoTo 2

    GroupCount = GroupCount + 1
    Files(GroupCount) = MyName
    
2:
    MyName = Dir
    
Loop

For i = 1 To GroupCount
    With frm
        x = Files(i)
        If path = "WIP" Then
            If GetValue(Main.Main_MasterPath.Value & path & "\", x, "ADMIN", "b88") = UCase("Quote Accepted") Then
                .AddItem Left(x, Len(x) - 4) & " *"
            Else
                .AddItem Left(x, Len(x) - 4)
            End If
        Else
            If path = "quotes" Then
                If GetValue(Main.Main_MasterPath & path & "\", x, "Admin", "b88") = "New Quote" Then
                    .AddItem Left(x, Len(x) - 4) & " *"
                Else
                    .AddItem Left(x, Len(x) - 4)
                End If
            Else
                .AddItem Left(x, Len(x) - 4)
            End If
        End If
    
    End With
Next i

End Function


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




# =======================================================
# Module 123
# =======================================================
Attribute VB_Name = "a_Main"
Sub ShowMenu()

Main.Main_MasterPath.Value = ActiveWorkbook.path & "\"
Main.Show

End Sub

Sub sadf()

Do
    ActiveCell.Value = ActiveCell.Offset(-1, 0).Value - 1
    ActiveCell.Offset(1, 0).Select
Loop Until ActiveCell.Offset(-1, 0).Value = 1011

End Sub


# =======================================================
# Module 125
# =======================================================
Attribute VB_Name = "fwip"
Attribute VB_Base = "0{0613A0D3-B4AB-496A-AC33-522966CBFD73}{72C6DCB3-085C-4A64-875D-3C695AC02840}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Type Jobs
    Dat As Date
    Cust As String
    Job As String
    JobD As Currency
    Qty As String
    Cod As String
    Desc As String
    Remarks As String
    DDat As String
    
    OperatorN(1 To 15) As String
    OperatorType(1 To 15) As String
    
End Type
Private Sub Go_Click()
Dim TempSheet As String
Dim Job(1 To 5000) As Jobs

fwip.Label1.Caption = "Please Wait"

Application.DisplayAlerts = False

'OpenBook (main.Main_MasterPath & "Wip.xls")

Workbooks.Open Main.Main_MasterPath & "WIP.xls", ReadOnly:=True
Range("bb1").Select
Selection.End(xlToLeft).Select
col = ActiveCell.Column

Range("A1").Select
Selection.End(xlDown).Select

Range("A2", Range("A2").Offset(ActiveCell.Row, col - 1).Address).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Sort Key1:=Range("h3"), Order1:=xlAscending, Header:=xlYes, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
            
Range("A3").Select
fwip.Hide
Main.Hide
i = 0
If ActiveCell.FormulaR1C1 <> "" Then
    Do
        i = i + 1
        With Job(i)
            .Dat = ActiveCell.Offset(0, 0).FormulaR1C1
            .Cust = ActiveCell.Offset(0, 1).FormulaR1C1
            .Job = ActiveCell.Offset(0, 2).FormulaR1C1
            .JobD = CCur(ActiveCell.Offset(0, 3).Value)
            .Qty = ActiveCell.Offset(0, 4).FormulaR1C1
            .Cod = ActiveCell.Offset(0, 5).FormulaR1C1
            .Desc = ActiveCell.Offset(0, 6).FormulaR1C1
            .Remarks = ActiveCell.Offset(0, 8).FormulaR1C1
            .DDat = ActiveCell.Offset(0, 12).FormulaR1C1
            x = 0
            For j = 1 To 30 Step 2
                x = x + 1
                .OperatorType(x) = ActiveCell.Offset(0, 14 + j).FormulaR1C1
            Next j
            x = 0
            For j = 1 To 30 Step 2
                x = x + 1
                .OperatorN(x) = ActiveCell.Offset(0, 15 + j).FormulaR1C1
            Next j
        End With
        ActiveCell.Offset(1, 0).Select
    Loop Until ActiveCell.FormulaR1C1 = ""
Else
    End
End If

fwip.Hide
Main.Hide
'Operation
If ROperation.Value = True Then
    Workbooks.Add
    
    OPP = ""
    If MsgBox("Specific Operation?", vbYesNo) = vbYes Then
        OPP = InputBox("Which Operation")
    End If
    
    For j = 1 To i
        With Job(j)
            For k = 1 To 15
                If OPP <> "" Then
                    If Trim(UCase(.OperatorType(k))) <> Trim(UCase(OPP)) Then GoTo SkipOPP
                End If
                    
                If .OperatorType(k) <> "" Then
                    TempSheet = "OPERATION - " & .OperatorType(k)
                    On Error GoTo AddSheet
                        Sheets(Remove_Characters(Trim(TempSheet))).Select
                    On Error GoTo Err
                    
                    ActiveCell.FormulaR1C1 = .Dat
                    ActiveCell.Offset(0, 1).FormulaR1C1 = .Cust
                    ActiveCell.Offset(0, 2).FormulaR1C1 = .Job
                    ActiveCell.Offset(0, 3).FormulaR1C1 = .JobD
                    ActiveCell.Offset(0, 4).FormulaR1C1 = .Qty
                    ActiveCell.Offset(0, 5).FormulaR1C1 = .Cod
                    ActiveCell.Offset(0, 6).FormulaR1C1 = .Desc
                    ActiveCell.Offset(0, 7).FormulaR1C1 = .Remarks
                    ActiveCell.Offset(0, 8).FormulaR1C1 = .DDat
                    
                    If k > 1 Then
                        If .OperatorType(k - 1) = "" Then
                            ActiveCell.Offset(0, 9).FormulaR1C1 = "*"
                            Selection.EntireRow.Font.Bold = True
                        End If
                    End If
                    ActiveCell.Offset(1, 0).Select
                End If
                TempSheet = ""
SkipOPP:
            Next k
        End With
    Next j
    
    For Each sh In Sheets
        sh.Select
        Cells.EntireColumn.AutoFit
        Range("A1:i5000").Select
        Selection.Sort Key1:=Range("H2"), Order1:=xlAscending, Key2:=Range("G2") _
            , Order2:=xlAscending, Header:=xlYes, OrderCustom:=1, MatchCase:= _
            False, Orientation:=xlTopToBottom
            
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
        
        Range("a:a").NumberFormat = "DD MMM YYYY"
        Range("i:i").NumberFormat = "DD MMM YYYY"
    Next sh
    
    DeleteSheet ("sheet1")
    DeleteSheet ("sheet2")
    DeleteSheet ("sheet3")

    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs (Main.Main_MasterPath & "TEMPLATES\Operation.xls")
End If

'Operator
If ROperator.Value = True Then
    Workbooks.Add
    
    For j = 1 To i
        With Job(j)
            For k = 1 To 15
                If Trim(.OperatorN(k)) <> "" Then
                    TempSheet = Remove_Characters("OPERATOR - " & Trim(.OperatorN(k)))
                    On Error GoTo AddSheet
                        Sheets(TempSheet).Select
                    On Error GoTo Err
                    
                    ActiveCell.FormulaR1C1 = .Dat
                    ActiveCell.Offset(0, 1).FormulaR1C1 = .Cust
                    ActiveCell.Offset(0, 2).FormulaR1C1 = .Job
                    ActiveCell.Offset(0, 3).FormulaR1C1 = .JobD
                    ActiveCell.Offset(0, 4).FormulaR1C1 = .Qty
                    ActiveCell.Offset(0, 5).FormulaR1C1 = .Cod
                    ActiveCell.Offset(0, 6).FormulaR1C1 = .Desc
                    ActiveCell.Offset(0, 7).FormulaR1C1 = .Remarks
                    ActiveCell.Offset(0, 8).FormulaR1C1 = .DDat
                    
                    If k > 1 Then
                        If .OperatorN(k - 1) = "" Then
                            ActiveCell.Offset(0, 9).FormulaR1C1 = "*"
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
        Range("A1:i5000").Select
        Selection.Sort Key1:=Range("H2"), Order1:=xlAscending, Key2:=Range("G2") _
            , Order2:=xlAscending, Header:=xlYes, OrderCustom:=1, MatchCase:= _
            False, Orientation:=xlTopToBottom
            
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
        
        Range("a:a").NumberFormat = "DD MMM YYYY"
        Range("i:i").NumberFormat = "DD MMM YYYY"
        Range("A1").Select
    Next sh
    
    DeleteSheet ("sheet1")
    DeleteSheet ("sheet2")
    DeleteSheet ("sheet3")
    
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs (Main.Main_MasterPath & "TEMPLATES\Operator.xls")
    
End If

'End With

Windows("wip.xls").Activate
If fwip.RDueDate.Value = True Then
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs (Main.Main_MasterPath & "TEMPLATES\Due Date.xls")
    Range("a1").Select
Else
    ActiveWorkbook.Close False
End If

If fwip.RWIP.Value = True Then
    Workbooks.Open Main.Main_MasterPath & "WIP.xls", ReadOnly:=True
    
    Range("A2", Range("A2").Offset(ActiveCell.Row, col - 1).Address).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Sort Key1:=Range("A3"), Order1:=xlAscending, Header:=xlYes, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
            
    Range("A1").Select
End If

'New

If fwip.Job_DueDate.Value = True Then
    Workbooks.Open Main.Main_MasterPath & "WIP.xls", ReadOnly:=True
    
    Windows("wip.xls").Activate
    Range("A1").Select
    Do
        ActiveCell.Offset(0, 1).Select
    Loop Until ActiveCell.Value = "CustomerDelivery_Date" Or ActiveCell.FormulaR1C1 = ""
    
    Sortcol = ActiveCell.Address
    
    Range("A3", Range("A3").Offset(0, col - 1).Address).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Sort Key1:=Range(Range(Sortcol).Offset(2, 0).Address), Order1:=xlAscending, Header:=xlNo, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
    
    ShowOfficeCols
        Range("b1").Select
    Application.DisplayAlerts = False
                    
    With ActiveSheet.PageSetup
    
        .CenterHeader = "OFFICE DUE DATE"
        .RightHeader = "&D &T"
        
    End With
    
    ActiveWorkbook.SaveAs (Main.Main_MasterPath & "TEMPLATES\CustomerDelivery_Date.xls")

   
End If

If fwip.Office_Customer.Value = True Then
    Workbooks.Open Main.Main_MasterPath & "WIP.xls", ReadOnly:=True
    
    Windows("wip.xls").Activate
    Range("A1").Select
    Do
        ActiveCell.Offset(0, 1).Select
        If UCase(ActiveCell.Value) = UCase("Customer") Then Sortcol1 = ActiveCell.Address
        If UCase(ActiveCell.Value) = UCase("Job_Number") Then Sortcol2 = ActiveCell.Address
        'If ActiveCell.Value = "" Then Sortcol2 = ActiveCell.Address
        
    Loop Until ActiveCell.FormulaR1C1 = ""

    
    Range("A3", Range("A3").Offset(0, col - 1).Address).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Sort Key1:=Range(Range(Sortcol1).Offset(2, 0).Address), Order1:=xlAscending, Header:=xlNo, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom _
        , Key2:=Range(Range(Sortcol2).Offset(2, 0).Address), Order2:=xlAscending '_
        ', Key3:=Range(Range(Sortcol3).Offset(2, 0).Address), Order3:=xlAscending
    
    ShowOfficeCols
        Range("b1").Select
                            
    With ActiveSheet.PageSetup
    
        .CenterHeader = "OFFICE CUSTOMER"
        .RightHeader = "&D &T"
        
    End With
    
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs (Main.Main_MasterPath & "TEMPLATES\Office_Customer.xls")
End If

'CUSTOMER
If fwip.Workshop_Customer.Value = True Then
    Workbooks.Open Main.Main_MasterPath & "WIP.xls", ReadOnly:=True
    
    Windows("wip.xls").Activate
    Range("A1").Select
    Do
        ActiveCell.Offset(0, 1).Select
         
        If UCase(ActiveCell.Value) = UCase("Customer") Then Sortcol1 = ActiveCell.Address
        If UCase(ActiveCell.Value) = UCase("Job_Number") Then Sortcol2 = ActiveCell.Address
        'If ActiveCell.Value = "" Then Sortcol2 = ActiveCell.Address
        
    Loop Until ActiveCell.FormulaR1C1 = ""
    
'    Sortcol = ActiveCell.Address
    
    Range("A3", Range("A3").Offset(0, col - 1).Address).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Sort Key1:=Range(Range(Sortcol1).Offset(2, 0).Address), Order1:=xlAscending, Header:=xlNo, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom _
        , Key2:=Range(Range(Sortcol2).Offset(2, 0).Address), Order2:=xlAscending '_
        ', Key3:=Range(Range(Sortcol3).Offset(2, 0).Address), Order3:=xlAscending
    
    ShowWorkshopCols
        Range("b1").Select
                            
    With ActiveSheet.PageSetup
    
        .CenterHeader = "WORKSHOP CUSTOMER"
        .RightHeader = "&D &T"
        
    End With

    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs (Main.Main_MasterPath & "TEMPLATES\Workshop_Customer.xls")
End If

'OFFICE JOB NUMBER
If fwip.Office_JobNumber.Value = True Then
    Workbooks.Open Main.Main_MasterPath & "WIP.xls", ReadOnly:=True
    
    Windows("wip.xls").Activate
    Range("A1").Select
    Do
        ActiveCell.Offset(0, 1).Select
    Loop Until ActiveCell.Value = "Converted_JN" Or ActiveCell.FormulaR1C1 = ""
    
    Sortcol = ActiveCell.Address
    
    Range("A3", Range("A3").Offset(0, col - 1).Address).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Sort Key1:=Range(Range(Sortcol).Offset(2, 0).Address), Order1:=xlAscending, Header:=xlNo, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortTextAsNumbers
        
    
    ShowOfficeCols
            Range("b1").Select
                                    
    With ActiveSheet.PageSetup
    
        .CenterHeader = "OFFICE JOB NUMBER"
        .RightHeader = "&D &T"
        
    End With
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs (Main.Main_MasterPath & "TEMPLATES\Office_JobNumber.xls")

   
End If

'WORKSHOP JOB NUMBER
If fwip.Workshop_JobNumber.Value = True Then
    Workbooks.Open Main.Main_MasterPath & "WIP.xls", ReadOnly:=True
    
    Windows("wip.xls").Activate
    Range("A1").Select
    Do
        ActiveCell.Offset(0, 1).Select
    Loop Until ActiveCell.Value = "Converted_JN" Or ActiveCell.FormulaR1C1 = ""
    
    Sortcol = ActiveCell.Address
    
    Range("A3", Range("A3").Offset(0, col - 1).Address).Select
    Range(Selection, Selection.End(xlDown)).Select

    Selection.Sort Key1:=Range(Range(Sortcol).Offset(2, 0).Address), Order1:=xlAscending, Header:=xlNo, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortTextAsNumbers
        
    ShowWorkshopCols
        
    Application.DisplayAlerts = False
        Range("b1").Select
                        
    With ActiveSheet.PageSetup
    
        .CenterHeader = "WORKSHOP JOB NUMBER"
        .RightHeader = "&D &T"
        
    End With
    ActiveWorkbook.SaveAs (Main.Main_MasterPath & "TEMPLATES\Workshop_JobNumber.xls")

End If

'DUEDATE
If fwip.Job_WorkshopDueDate.Value = True Then
    Workbooks.Open Main.Main_MasterPath & "WIP.xls", ReadOnly:=True
    
    Windows("wip.xls").Activate
    Range("A1").Select
    Do
        ActiveCell.Offset(0, 1).Select
    Loop Until ActiveCell.Value = "Job_WorkshopDueDate" Or ActiveCell.FormulaR1C1 = ""
    
    Sortcol = ActiveCell.Address
    
    Range("A3", Range("A3").Offset(0, col - 1).Address).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Sort Key1:=Range(Range(Sortcol).Offset(2, 0).Address), Order1:=xlAscending, Header:=xlNo, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
    
    ShowWorkshopCols
    
    Application.DisplayAlerts = False
        Range("b1").Select
                            
    With ActiveSheet.PageSetup
    
        .CenterHeader = "WORKSHOP DUE DATE"
        .RightHeader = "&D &T"
        
    End With
    
    ActiveWorkbook.SaveAs (Main.Main_MasterPath & "TEMPLATES\Job_WorkshopDueDate.xls")
   
End If

Application.DisplayAlerts = True

Unload fwip
Unload Main

Exit Sub
AddSheet:
    Sheets.Add
    ActiveSheet.Name = Remove_Characters(TempSheet)
    ActiveSheet.PageSetup.CenterHeader = TempSheet
    ActiveCell.FormulaR1C1 = "DATE"
    ActiveCell.Offset(0, 1).FormulaR1C1 = "CUSTOMER"
    ActiveCell.Offset(0, 2).FormulaR1C1 = "JOB"
    ActiveCell.Offset(0, 3).FormulaR1C1 = "JOB"
    ActiveCell.Offset(0, 4).FormulaR1C1 = "QTY"
    ActiveCell.Offset(0, 5).FormulaR1C1 = "COMPONENT CODE"
    ActiveCell.Offset(0, 6).FormulaR1C1 = "COMPONENT DESCRIPTION"
    ActiveCell.Offset(0, 7).FormulaR1C1 = "REMARKS"
    ActiveCell.Offset(0, 8).FormulaR1C1 = "DUE DATE"
    Columns("h:h").NumberFormat = "dd mmm"
    Columns("A:A").NumberFormat = "dd mmm"
    Selection.EntireRow.Font.Bold = True
    
    Columns("A:A").ColumnWidth = 10
    Columns("b:b").ColumnWidth = 18
    Columns("c:c").ColumnWidth = 10
    Columns("e:e").ColumnWidth = 6
    Columns("g:g").ColumnWidth = 30
    Columns("h:h").ColumnWidth = 20
    Columns("i:i").ColumnWidth = 10
    Cells.RowHeight = 30
    
    ActiveCell.Offset(1, 0).Select
    
Resume

Err:
    MsgBox ("CRITCAL ERROR")
Resume

End Sub


Private Function ShowOfficeCols()
    Range("A1").Select
    Do
        Selection.EntireColumn.Hidden = True
        
        Select Case ActiveCell.Value
            Case "Job_StartDate"
                Selection.EntireColumn.Hidden = False
            Case "Job_Urgency"
                Selection.EntireColumn.Hidden = False
            Case "CUSTOMER"
                Selection.EntireColumn.Hidden = False
            Case "Job_Number"
                Selection.EntireColumn.Hidden = False
            Case "Component_Quantity"
                Selection.EntireColumn.Hidden = False
            Case "Component_Code"
                    Selection.EntireColumn.Hidden = False
            Case "Component_Description"
                Selection.EntireColumn.Hidden = False
            Case "Component_Comments"
                Selection.EntireColumn.Hidden = False
            Case "CustomerDelivery_Date"
                Selection.EntireColumn.Hidden = False
            Case "CustomerOrderNumber"
                Selection.EntireColumn.Hidden = False
            Case "Component_Price"
                Selection.EntireColumn.Hidden = False
            Case "Component_Comments"
                Selection.EntireColumn.Hidden = False
            Case "Component_Comments"
                Selection.EntireColumn.Hidden = False
            Case "Component_Comments"
                Selection.EntireColumn.Hidden = False
            Case "Component_DrawingNumber_SampleNumber"
                Selection.EntireColumn.Hidden = False
                
         End Select
         ActiveCell.Offset(0, 1).Select
         
       Loop Until ActiveCell.Value = ""
End Function

Private Function ShowWorkshopCols()
    Range("A1").Select
    Do
        Selection.EntireColumn.Hidden = True

        Select Case ActiveCell.Value
            Case "Job_StartDate"
                Selection.EntireColumn.Hidden = False
            Case "Job_Urgency"
                Selection.EntireColumn.Hidden = False
            Case "CUSTOMER"
                Selection.EntireColumn.Hidden = False
            Case "Job_Number"
                Selection.EntireColumn.Hidden = False
            Case "Job_WorkshopDueDate"
                Selection.EntireColumn.Hidden = False
            Case "Job_WorkshopDueDate"
                Selection.EntireColumn.Hidden = False
            Case "Component_Quantity"
                Selection.EntireColumn.Hidden = False
            Case "Component_Code"
                    Selection.EntireColumn.Hidden = False
            Case "Component_Description"
                Selection.EntireColumn.Hidden = False
            Case "Component_Comments"
                Selection.EntireColumn.Hidden = False
              Case "Component_Comments"
                Selection.EntireColumn.Hidden = False
            Case "Component_Comments"
                Selection.EntireColumn.Hidden = False
            Case "Component_Comments"
                Selection.EntireColumn.Hidden = False
            Case " "
                Selection.EntireColumn.Hidden = False
            Case "Component_DrawingNumber_SampleNumber"
                Selection.EntireColumn.Hidden = False
            Case "Operation01_Type"
            Selection.EntireColumn.Hidden = False
            Case "Operation01_Operator"
            Selection.EntireColumn.Hidden = False
            Case "Operation02_Type"
            Selection.EntireColumn.Hidden = False
            Case "Operation02_Operator"
            Selection.EntireColumn.Hidden = False
            Case "Operation03_Type"
            Selection.EntireColumn.Hidden = False
            Case "Operation03_Operator"
            Selection.EntireColumn.Hidden = False
            Case "Operation04_Type"
            Selection.EntireColumn.Hidden = False
            Case "Operation04_Operator"
            Selection.EntireColumn.Hidden = False
            Case "Operation05_Type"
            Selection.EntireColumn.Hidden = False
            Case "Operation05_Operator"
            Selection.EntireColumn.Hidden = False
            Case "Operation06_Type"
            Selection.EntireColumn.Hidden = False
            Case "Operation06_Operator"
            Selection.EntireColumn.Hidden = False
            Case "Operation07_Type"
            Selection.EntireColumn.Hidden = False
            Case "Operation07_Operator"
            Selection.EntireColumn.Hidden = False
            Case "Operation08_Type"
            Selection.EntireColumn.Hidden = False
            Case "Operation08_Operator"
            Selection.EntireColumn.Hidden = False
            Case "Operation09_Type"
            Selection.EntireColumn.Hidden = False
            Case "Operation09_Operator"
            Selection.EntireColumn.Hidden = False
            Case "Operation10_Type"
            Selection.EntireColumn.Hidden = False
            Case "Operation10_Operator"
            Selection.EntireColumn.Hidden = False
            Case "Operation11_Type"
            Selection.EntireColumn.Hidden = False
            Case "Operation11_Operator"
            Selection.EntireColumn.Hidden = False
            Case "Operation12_Type"
            Selection.EntireColumn.Hidden = False
            Case "Operation12_Operator"
            Selection.EntireColumn.Hidden = False
            Case "Operation13_Type"
            Selection.EntireColumn.Hidden = False
            Case "Operation13_Type"
            Selection.EntireColumn.Hidden = False
            Case "Operation14_Type"
            Selection.EntireColumn.Hidden = False
            Case "Operation14_Operator"
            Selection.EntireColumn.Hidden = False
            Case "Operation15_Type"
            Selection.EntireColumn.Hidden = False
            Case "Operation15_Operator"
            Selection.EntireColumn.Hidden = False

         End Select
         ActiveCell.Offset(0, 1).Select
         
       Loop Until ActiveCell.Value = ""
End Function



