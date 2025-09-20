Private Sub AddMore_Click()
    On Error GoTo Error_Handler

    If SaveCurrentEnquiry() Then
        ClearForm
        Me.Enquiry_Date.Caption = Format(Now(), "dd mmm yyyy")
        MsgBox "Enquiry saved successfully. Ready for next enquiry.", vbInformation
    End If
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "AddMore_Click", "FEnquiry"
End Sub

Private Sub SaveQ_Click()
    On Error GoTo Error_Handler

    If SaveCurrentEnquiry() Then
        MsgBox "Enquiry saved successfully.", vbInformation
        Unload Me
    End If
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "SaveQ_Click", "FEnquiry"
End Sub

Private Sub AddNewClient_Click()
    Dim CustomerName As String

    On Error GoTo Error_Handler

    CustomerName = Trim(Me.Customer.Value)
    If CustomerName = "" Then
        MsgBox "Please enter a customer name first.", vbInformation
        Exit Sub
    End If

    If BusinessController.CreateNewCustomer(CustomerName) Then
        MsgBox "Customer '" & CustomerName & "' created successfully.", vbInformation
    Else
        MsgBox "Failed to create customer '" & CustomerName & "'.", vbCritical
    End If
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "AddNewClient_Click", "FEnquiry"
End Sub

Private Sub Dat_Click()
    On Error GoTo Error_Handler

    Dim SelectedDate As Date
    SelectedDate = ShowCalendar()

    If SelectedDate <> 0 Then
        Me.Enquiry_Date.Caption = Format(SelectedDate, "dd mmm yyyy")
    End If
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "Dat_Click", "FEnquiry"
End Sub

Private Sub Cancel_Click()
    Unload Me
End Sub

Private Function SaveCurrentEnquiry() As Boolean
    Dim EnquiryInfo As EnquiryData
    Dim ValidationErrors As String

    On Error GoTo Error_Handler

    With EnquiryInfo
        .CustomerName = Trim(Me.Customer.Value)
        .ContactPerson = Trim(Me.Contact_Person.Value)
        .CompanyPhone = Trim(Me.Company_Phone.Value)
        .CompanyFax = Trim(Me.Company_Fax.Value)
        .Email = Trim(Me.Email.Value)
        .ComponentDescription = Trim(Me.Component_Description.Value)
        .ComponentCode = Trim(Me.Component_Code.Value)
        .MaterialGrade = Trim(Me.Component_Grade.Value)

        If IsNumeric(Me.Component_Quantity.Value) Then
            .Quantity = CLng(Me.Component_Quantity.Value)
        Else
            .Quantity = 0
        End If

        .SearchKeywords = .CustomerName & " " & .ComponentDescription & " " & .ComponentCode
    End With

    ValidationErrors = BusinessController.ValidateEnquiryData(EnquiryInfo)
    If ValidationErrors <> "" Then
        MsgBox "Please correct the following errors:" & vbCrLf & vbCrLf & ValidationErrors, vbExclamation
        SaveCurrentEnquiry = False
        Exit Function
    End If

    If Me.Enquiry_Date.Caption = "Please click here to insert a date" Then
        If MsgBox("Do you want to cancel the save to enter a date?", vbYesNo, "MEM") = vbYes Then
            SaveCurrentEnquiry = False
            Exit Function
        End If
    End If

    SaveCurrentEnquiry = BusinessController.CreateNewEnquiry(EnquiryInfo)

    If SaveCurrentEnquiry Then
        Me.File_Name.Value = EnquiryInfo.EnquiryNumber
        Me.Enquiry_Number.Value = EnquiryInfo.EnquiryNumber
        MsgBox "The File Number for this Enquiry is: " & EnquiryInfo.EnquiryNumber, vbInformation
    End If
    Exit Function

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "SaveCurrentEnquiry", "FEnquiry"
    SaveCurrentEnquiry = False
End Function

Private Sub ClearForm()
    On Error GoTo Error_Handler

    Me.Customer.Value = ""
    Me.Contact_Person.Value = ""
    Me.Company_Phone.Value = ""
    Me.Company_Fax.Value = ""
    Me.Email.Value = ""
    Me.Component_Description.Value = ""
    Me.Component_Code.Value = ""
    Me.Component_Grade.Value = ""
    Me.Component_Quantity.Value = ""
    Me.File_Name.Value = ""
    Me.Enquiry_Number.Value = ""
    Me.System_Status.Value = "To Quote"
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "ClearForm", "FEnquiry"
End Sub

Public Sub LoadEnquiry(ByVal FilePath As String)
    Dim EnquiryInfo As EnquiryData

    On Error GoTo Error_Handler

    EnquiryInfo = BusinessController.LoadEnquiry(FilePath)

    If EnquiryInfo.EnquiryNumber <> "" Then
        With Me
            .Enquiry_Number.Value = EnquiryInfo.EnquiryNumber
            .Customer.Value = EnquiryInfo.CustomerName
            .Contact_Person.Value = EnquiryInfo.ContactPerson
            .Company_Phone.Value = EnquiryInfo.CompanyPhone
            .Company_Fax.Value = EnquiryInfo.CompanyFax
            .Email.Value = EnquiryInfo.Email
            .Component_Description.Value = EnquiryInfo.ComponentDescription
            .Component_Code.Value = EnquiryInfo.ComponentCode
            .Component_Grade.Value = EnquiryInfo.MaterialGrade
            .Component_Quantity.Value = EnquiryInfo.Quantity
            .File_Name.Value = EnquiryInfo.EnquiryNumber
            .Enquiry_Date.Caption = Format(EnquiryInfo.DateCreated, "dd mmm yyyy")
        End With
    End If
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "LoadEnquiry", "FEnquiry"
End Sub

Private Function ShowCalendar() As Date
    On Error GoTo Error_Handler

    ShowCalendar = CDate(InputBox("Enter date (dd/mm/yyyy):", "Date Selection", Format(Now, "dd/mm/yyyy")))
    Exit Function

Error_Handler:
    ShowCalendar = 0
End Function

Private Sub Component_Description_Change()
    On Error GoTo Error_Handler

    If Len(Me.Component_Description.Value) > 0 Then
        LoadComponentCodes
    End If
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "Component_Description_Change", "FEnquiry"
End Sub

Private Sub LoadComponentCodes()
    On Error GoTo Error_Handler

    Dim PriceListPath As String
    PriceListPath = DataManager.GetRootPath & "\Templates\Price List.xls"

    If DataManager.FileExists(PriceListPath) Then
        Dim ComponentCode As String
        ComponentCode = DataUtilities.FindComponentCode(PriceListPath, Me.Component_Description.Value)

        If ComponentCode <> "" Then
            Me.Component_Code.Value = ComponentCode
        End If
    End If
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "LoadComponentCodes", "FEnquiry"
End Sub

Private Sub Component_Code_Change()
    On Error GoTo Error_Handler

    If Len(Me.Component_Code.Value) > 0 Then
        LoadGrades
    End If
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "Component_Code_Change", "FEnquiry"
End Sub

Private Sub LoadGrades()
    On Error GoTo Error_Handler

    Dim GradesPath As String
    GradesPath = DataManager.GetRootPath & "\Templates\Component_Grades.xls"

    If DataManager.FileExists(GradesPath) Then
        Dim Grades As Variant
        Grades = DataUtilities.GetComponentGrades(GradesPath, Me.Component_Code.Value)

        If UBound(Grades) >= 0 Then
            Me.Component_Grade.Value = Grades(0)
        End If
    End If
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "LoadGrades", "FEnquiry"
End Sub

Private Sub Price_Change()
    On Error GoTo Error_Handler

    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "Price_Change", "FEnquiry"
End Sub