Private CurrentMode As String

Private Sub butSaveJG_Click()
    On Error GoTo Error_Handler

    If WorkflowManagement.SaveDirectJob(Me) Then
        SystemCore.ShowInformation "Job created successfully.", "Job Created"
        Unload Me
    End If
    Exit Sub

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "butSaveJG_Click", "FJG"
End Sub

Private Sub but_SaveAsCTItem_Click()
    On Error GoTo Error_Handler

    If WorkflowManagement.SaveAsContract(Me) Then
        SystemCore.ShowInformation "Contract template saved successfully.", "Contract Saved"
        Unload Me
    End If
    Exit Sub

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "but_SaveAsCTItem_Click", "FJG"
End Sub

Private Sub Cancel_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    On Error GoTo Error_Handler

    WorkflowManagement.InitializeJobGenerationForm Me
    Exit Sub

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "UserForm_Initialize", "FJG"
End Sub

' **Purpose**: Business logic extracted to WorkflowManagement module
' **CLAUDE.md Compliance**: All private functions moved to WorkflowManagement.bas
' SaveDirectJob → WorkflowManagement.SaveDirectJob
' SaveAsContract → WorkflowManagement.SaveAsContract