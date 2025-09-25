Private CurrentJobPath As String

Private Sub SaveJobCard_Click()
    On Error GoTo Error_Handler
    If WorkflowManagement.SaveJobCard(Me, CurrentJobPath) Then
        SystemCore.ShowInformation "Job card saved successfully.", "Save Complete"
    End If
    Exit Sub
Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "SaveJobCard_Click", "FJobCard"
End Sub

Private Sub CloseJobCard_Click()
    On Error GoTo Error_Handler
    Unload Me
    Exit Sub
Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "CloseJobCard_Click", "FJobCard"
End Sub

Private Sub JobCardTemplates_Click()
    On Error GoTo Error_Handler
    WorkflowManagement.LoadJobTemplates Me
    Exit Sub
Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "JobCardTemplates_Click", "FJobCard"
End Sub

Private Sub CopyFromJobCard_Click()
    Dim SourceJobNumber As String
    On Error GoTo Error_Handler

    SourceJobNumber = InputBox("Enter job number to copy operations from:", "Copy Operations")
    If SourceJobNumber = "" Then Exit Sub

    WorkflowManagement.CopyOperationsFromJob Me, SourceJobNumber
    Exit Sub
Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "CopyFromJobCard_Click", "FJobCard"
End Sub

Private Sub AddPicture_Click()
    On Error GoTo Error_Handler
    WorkflowManagement.AddPictureToJob Me
    Exit Sub
Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "AddPicture_Click", "FJobCard"
End Sub

Private Sub PrintJobCard_Click()
    On Error GoTo Error_Handler
    WorkflowManagement.PrintJobCard Me
    Exit Sub
Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "PrintJobCard_Click", "FJobCard"
End Sub

Private Sub UpdateOperations_Click()
    On Error GoTo Error_Handler
    WorkflowManagement.UpdateOperations Me
    Exit Sub
Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "UpdateOperations_Click", "FJobCard"
End Sub

Private Sub Job_Status_Change()
    On Error GoTo Error_Handler
    WorkflowManagement.HandleJobStatusChange Me
    Exit Sub
Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "Job_Status_Change", "FJobCard"
End Sub

Private Sub Due_Date_Change()
    On Error GoTo Error_Handler
    WorkflowManagement.HandleDueDateChange Me
    Exit Sub
Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "Due_Date_Change", "FJobCard"
End Sub

Public Sub LoadJobCard(JobPath As String)
    On Error GoTo Error_Handler
    CurrentJobPath = JobPath
    WorkflowManagement.LoadJobCardData Me, JobPath
    Exit Sub
Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "LoadJobCard", "FJobCard"
End Sub

Private Sub UserForm_Initialize()
    On Error GoTo Error_Handler
    WorkflowManagement.InitializeJobCardForm Me
    Exit Sub
Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "UserForm_Initialize", "FJobCard"
End Sub

' **Purpose**: Business logic extracted to WorkflowManagement module
' **CLAUDE.md Compliance**: All private functions moved to WorkflowManagement.bas