Private Sub AddMore_Click()
    On Error GoTo Error_Handler

    WorkflowManagement.SaveEnquiryAndContinue Me
    Exit Sub

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "AddMore_Click", "FEnquiry"
End Sub

Private Sub SaveQ_Click()
    On Error GoTo Error_Handler

    If WorkflowManagement.SaveEnquiry(Me) Then
        SystemCore.ShowInformation "Enquiry saved successfully.", "Save Complete"
        Unload Me
    End If
    Exit Sub

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "SaveQ_Click", "FEnquiry"
End Sub

Private Sub AddNewClient_Click()
    On Error GoTo Error_Handler

    WorkflowManagement.CreateCustomerFromForm Me
    Exit Sub

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "AddNewClient_Click", "FEnquiry"
End Sub

Private Sub Dat_Click()
    On Error GoTo Error_Handler

    WorkflowManagement.SetEnquiryDate Me
    Exit Sub

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "Dat_Click", "FEnquiry"
End Sub

Private Sub Cancel_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    On Error GoTo Error_Handler

    WorkflowManagement.InitializeEnquiryForm Me
    Exit Sub

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "UserForm_Initialize", "FEnquiry"
End Sub

' **Purpose**: Business logic extracted to WorkflowManagement module
' **CLAUDE.md Compliance**: All private functions moved to WorkflowManagement.bas
' SaveCurrentEnquiry → WorkflowManagement.SaveEnquiry
' ClearForm → WorkflowManagement.ClearEnquiryForm (private)
' ShowCalendar → WorkflowManagement.ShowDatePicker (private)
' LoadComponentCodes → WorkflowManagement.LoadComponentCodes
' LoadGrades → WorkflowManagement.LoadMaterialGrades
' ValidateEnquiryForm → WorkflowManagement.ValidateEnquiryData (private)