VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmEnquiry 
   Caption         =   "MEM: Enquiry"
   ClientHeight    =   8865.001
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7680
   OleObjectBlob   =   "FrmEnquiry.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmEnquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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




