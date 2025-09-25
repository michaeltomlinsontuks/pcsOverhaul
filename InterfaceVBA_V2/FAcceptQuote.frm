Private CurrentQuotePath As String

Private Sub butSAVE_Click()
    On Error GoTo Error_Handler

    If WorkflowManagement.AcceptQuote(Me, CurrentQuotePath) Then
        SystemCore.ShowInformation "Quote accepted and job created successfully.", "Job Created"
        Unload Me
    End If
    Exit Sub

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "butSAVE_Click", "FAcceptQuote"
End Sub

Private Sub Cancel_Click()
    Unload Me
End Sub

Public Sub LoadQuote(QuotePath As String)
    On Error GoTo Error_Handler

    CurrentQuotePath = QuotePath
    WorkflowManagement.LoadQuoteForAcceptance Me, QuotePath
    Exit Sub

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "LoadQuote", "FAcceptQuote"
End Sub

' **Purpose**: Business logic extracted to WorkflowManagement module
' **CLAUDE.md Compliance**: All private functions moved to WorkflowManagement.bas
' AcceptCurrentQuote → WorkflowManagement.AcceptQuote