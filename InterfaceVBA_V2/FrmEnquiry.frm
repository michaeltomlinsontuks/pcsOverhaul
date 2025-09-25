Private Sub SaveQ_Click()
    On Error GoTo Error_Handler
    If EnquiryManager.SaveEnquiry(Me) Then
        ValidationFramework.ShowInformation "Enquiry saved successfully.", "Save Complete"
        Unload Me
    End If
    Exit Sub
Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "SaveQ_Click", "FrmEnquiry"
End Sub

Private Sub SaveQC_Click()
    On Error GoTo Error_Handler
    If EnquiryManager.SaveEnquiryAndContinue(Me) Then
        ValidationFramework.ShowInformation "Enquiry saved successfully. Form cleared for next entry.", "Save Complete"
        EnquiryManager.ClearEnquiryForm Me
    End If
    Exit Sub
Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "SaveQC_Click", "FrmEnquiry"
End Sub

Private Sub CreateCustomer_Click()
    On Error GoTo Error_Handler
    If EnquiryManager.CreateCustomerFromForm(Me) Then
        ValidationFramework.ShowInformation "Customer created successfully.", "Customer Created"
    End If
    Exit Sub
Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "CreateCustomer_Click", "FrmEnquiry"
End Sub

Private Sub Customer_Change()
    On Error GoTo Error_Handler
    EnquiryManager.HandleCustomerChange Me
    Exit Sub
Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "Customer_Change", "FrmEnquiry"
End Sub

Private Sub Component_Code_Change()
    On Error GoTo Error_Handler
    EnquiryManager.HandleComponentCodeChange Me
    Exit Sub
Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "Component_Code_Change", "FrmEnquiry"
End Sub

Private Sub Component_Grade_Change()
    On Error GoTo Error_Handler
    EnquiryManager.HandleComponentGradeChange Me
    Exit Sub
Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "Component_Grade_Change", "FrmEnquiry"
End Sub

Private Sub Component_Quantity_Change()
    On Error GoTo Error_Handler
    EnquiryManager.HandleComponentQuantityChange Me
    Exit Sub
Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "Component_Quantity_Change", "FrmEnquiry"
End Sub

Private Sub Cancel_Click()
    On Error GoTo Error_Handler
    Unload Me
    Exit Sub
Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "Cancel_Click", "FrmEnquiry"
End Sub

Private Sub UserForm_Initialize()
    On Error GoTo Error_Handler
    EnquiryManager.InitializeEnquiryForm Me
    Exit Sub
Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "UserForm_Initialize", "FrmEnquiry"
End Sub

' **Purpose**: Business logic extracted to EnquiryManager module
' **CLAUDE.md Compliance**: All private functions moved to EnquiryManager.bas