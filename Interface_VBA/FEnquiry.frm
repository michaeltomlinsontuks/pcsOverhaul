
' **Purpose**: Validates enquiry form data using standardized popup validation
' **Parameters**: None
' **Returns**: Boolean - True if all validations pass, False if any fail
' **Dependencies**: ValidationFramework
' **Side Effects**: Shows validation popup messages, sets focus to invalid fields
' **Errors**: Returns False on validation failure
Private Function ValidateEnquiryForm() As Boolean
    ValidateEnquiryForm = True

    ' Validate Customer
    If Not ValidationFramework.ValidateRequired(FrmEnquiry.Customer.Value, "Customer", FrmEnquiry.Customer) Then
        ValidateEnquiryForm = False
        Exit Function
    End If

    ' Validate Component Description
    If Not ValidationFramework.ValidateRequired(FrmEnquiry.Component_Description.Value, "Component Description", FrmEnquiry.Component_Description) Then
        ValidateEnquiryForm = False
        Exit Function
    End If

    ' Validate Component Quantity
    If Not ValidationFramework.ValidatePositiveNumber(FrmEnquiry.Component_Quantity.Value, "Component Quantity", FrmEnquiry.Component_Quantity) Then
        ValidateEnquiryForm = False
        Exit Function
    End If

    ' Validate Date (special handling for date caption)
    If Not ValidationFramework.ValidateSpecialDateCaption(FrmEnquiry.Enquiry_Date.Caption, "Enquiry Date") Then
        ValidateEnquiryForm = False
        Exit Function
    End If

    ' Additional business logic validation
    If Not ValidateCustomerExists() Then
        ValidateEnquiryForm = False
        Exit Function
    End If
End Function

' **Purpose**: Validates that customer exists in system
' **Parameters**: None
' **Returns**: Boolean - True if customer exists or user chooses to create, False otherwise
' **Dependencies**: ValidationFramework.ShowConfirmation, AddNewClient_Click
' **Side Effects**: May create new customer file
' **Errors**: Returns False if customer validation fails
Private Function ValidateCustomerExists() As Boolean
    ValidateCustomerExists = True

    ' Check if customer file exists
    If Dir(Main.Main_MasterPath & "Customers\" & FrmEnquiry.Customer.Value & ".xls") = "" Then
        If ValidationFramework.ShowConfirmation("Customer '" & FrmEnquiry.Customer.Value & "' does not exist. Create new customer?", "Customer Not Found") Then
            AddNewClient_Click
        Else
            FrmEnquiry.Customer.SetFocus
            ValidateCustomerExists = False
        End If
    End If
End Function
Private Sub AddMore_Click()
Dim Eq As Integer

' Validate form before processing
If Not ValidateEnquiryForm() Then Exit Sub

With Me
    .Enquiry_Number.Value = BusinessLogic.Calc_Next_Number("E")
    BusinessLogic.Confirm_Next_Number ("E")
    .File_Name.Value = .Enquiry_Number.Value
    ValidationFramework.ShowInformation "The File Number for this Enquiry is: " & Me.File_Name.Value, "Enquiry Saved"
End With

'Windows("Price List.xls").Activate
'ActiveWorkbook.Close (False)

x = FileOperations.OpenBook(Main.Main_MasterPath & "Templates\" & "_Enq.xls", True)

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
x = FileOperations.OpenBook(Main.Main_MasterPath & "Search.xls", False)
    Do
        If ActiveWorkbook.ReadOnly = True Then
            ActiveWorkbook.Close
            MsgBox ("This workbook is read only, please find the user with this workbook open and close it.")
            x = FileOperations.OpenBook(Main.Main_MasterPath & "Search.xls", False)
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
    .Enquiry_Number.Value = BusinessLogic.Calc_Next_Number("E")
    
    If .Customer.Value = "" Then .Customer.Clear

    x = DirectoryHelpers.List_Files("Customers", FrmEnquiry.Customer)

End With

x = FileOperations.OpenBook(Main.Main_MasterPath & "templates\price list.xls", True)

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

FileOperations.OpenBook (Main.Main_MasterPath & "templates\_client.xls")
    Cust = InputBox("Please enter a Company Name", "MEM")
    Range("b1").FormulaR1C1 = Cust
'    Range("b3").FormulaR1C1 = InputBox("Please enter a ContactPerson Person's Name", "MEM")
'    Range("b4").FormulaR1C1 = InputBox("Please enter a ContactPerson Number", "MEM")
    
ActiveWorkbook.SaveAs (Main.Main_MasterPath & "Customers\" & Range("company_Name").FormulaR1C1 & ".xls")
ActiveWorkbook.Close

With FrmEnquiry
    
    .Customer.Clear

    x = DirectoryHelpers.List_Files("Customers", FrmEnquiry.Customer)

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

' Validate form before processing
If Not ValidateEnquiryForm() Then Exit Sub

With Me
    .Enquiry_Number.Value = BusinessLogic.Calc_Next_Number("E")
    BusinessLogic.Confirm_Next_Number ("E")
    .File_Name.Value = .Enquiry_Number.Value
    ValidationFramework.ShowInformation "The File Number for this Enquiry is: " & Me.File_Name.Value, "Enquiry Saved"
End With

x = FileOperations.OpenBook(Main.Main_MasterPath & "Templates\" & "_Enq.xls", True)
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
x = FileOperations.OpenBook(Main.Main_MasterPath & "Search.xls", False)
    Do
        If ActiveWorkbook.ReadOnly = True Then
            ActiveWorkbook.Close
            MsgBox ("This workbook is read only, please find the user with this workbook open and close it.")
            x = FileOperations.OpenBook(Main.Main_MasterPath.Value & "Search.xls", False)
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

    x = DirectoryHelpers.List_Files("Customers", FrmEnquiry.Customer)

End With

x = FileOperations.OpenBook(Main.Main_MasterPath & "templates\price list.xls", True)
Sheets("Component_Descriptions").Select
Range("a2").Select

Do
    If ActiveCell.FormulaR1C1 <> ActiveCell.Offset(-1, 0).FormulaR1C1 Then Me.Component_Code.AddItem ActiveCell.FormulaR1C1
    ActiveCell.Offset(1, 0).Select
Loop Until ActiveCell.FormulaR1C1 = ""
ActiveWorkbook.Close False

' To add more Component_Grades to the enquiry form, add them below using ".AddItem " & What you want to add
x = FileOperations.OpenBook(Main.Main_MasterPath.Value & "templates\Component_Grades.xls", True)
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


