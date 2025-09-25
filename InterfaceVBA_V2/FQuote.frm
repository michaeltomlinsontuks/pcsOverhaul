Private CurrentEnquiryPath As String

Private Sub SaveQuote_Click()
    On Error GoTo Error_Handler

    If WorkflowManagement.SaveQuote(Me) Then
        SystemCore.ShowInformation "Quote saved successfully.", "Save Complete"
        Unload Me
    End If
    Exit Sub

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "SaveQuote_Click", "FQuote"
End Sub

Private Sub Cancel_Click()
    Unload Me
End Sub

Private Sub UnitPrice_Change()
    On Error GoTo Error_Handler

    WorkflowManagement.CalculateQuoteTotalPrice Me
    Exit Sub

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "UnitPrice_Change", "FQuote"
End Sub

Private Sub Quantity_Change()
    On Error GoTo Error_Handler

    WorkflowManagement.CalculateQuoteTotalPrice Me
    Exit Sub

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "Quantity_Change", "FQuote"
End Sub

Private Sub Component_Code_Change()
    On Error GoTo Error_Handler

    WorkflowManagement.LoadComponentPricing Me
    Exit Sub

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "Component_Code_Change", "FQuote"
End Sub

Private Sub ValidUntil_Click()
    On Error GoTo Error_Handler

    WorkflowManagement.SetQuoteValidUntilDate Me
    Exit Sub

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "ValidUntil_Click", "FQuote"
End Sub

Private Sub Search_Component_code_Click()
    On Error GoTo Error_Handler

    WorkflowManagement.SearchComponentCode Me
    Exit Sub

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "Search_Component_code_Click", "FQuote"
End Sub

Private Sub UserForm_Initialize()
    On Error GoTo Error_Handler

    WorkflowManagement.InitializeQuoteForm Me
    Exit Sub

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "UserForm_Initialize", "FQuote"
End Sub

Public Sub LoadFromEnquiry(EnquiryPath As String)
    On Error GoTo Error_Handler

    CurrentEnquiryPath = EnquiryPath
    WorkflowManagement.LoadQuoteFromEnquiry Me, EnquiryPath
    Exit Sub

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "LoadFromEnquiry", "FQuote"
End Sub

' **Purpose**: Business logic extracted to WorkflowManagement module
' **CLAUDE.md Compliance**: All private functions moved to WorkflowManagement.bas
' SaveCurrentQuote → WorkflowManagement.SaveQuote
' CalculateTotalPrice → WorkflowManagement.CalculateQuoteTotalPrice
' LoadPricing → WorkflowManagement.LoadComponentPricing (private)
' ShowCalendar → WorkflowManagement.ShowDatePicker (private)
' ClearForm → WorkflowManagement.ClearQuoteForm (private)