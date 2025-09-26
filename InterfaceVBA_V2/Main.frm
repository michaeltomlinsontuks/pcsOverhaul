Private Sub Add_Enquiry_Click()
    On Error GoTo Error_Handler
    UserInterface.AddEnquiry Me
    Exit Sub
Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "Add_Enquiry_Click", "Main"
End Sub

Private Sub Archive_Click()
    On Error GoTo Error_Handler
    UserInterface.ShowArchiveFiles Me
    Exit Sub
Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "Archive_Click", "Main"
End Sub

Private Sub but_CreateCTItem_Click()
    On Error GoTo Error_Handler
    UserInterface.CreateContractTemplateItem
    Exit Sub
Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "but_CreateCTItem_Click", "Main"
End Sub

Private Sub but_EditCTItem_Click()
    On Error GoTo Error_Handler
    UserInterface.EditContractTemplateItem Me
    Exit Sub
Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "but_EditCTItem_Click", "Main"
End Sub

Private Sub But_EditJC_Click()
    On Error GoTo Error_Handler
    UserInterface.EditJobCard Me
    Exit Sub
Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "But_EditJC_Click", "Main"
End Sub

Private Sub butEditSearch_Click()
    On Error GoTo Error_Handler
    UserInterface.EditSearchDatabase Me
    Exit Sub
Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "butEditSearch_Click", "Main"
End Sub

Private Sub butSearchHistory_Click()
    On Error GoTo Error_Handler
    UserInterface.ShowSearchHistory Me
    Exit Sub
Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "butSearchHistory_Click", "Main"
End Sub

Private Sub butJobHistory_Click()
    On Error GoTo Error_Handler
    UserInterface.ShowJobHistory Me
    Exit Sub
Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "butJobHistory_Click", "Main"
End Sub

Private Sub butQuoteHistory_Click()
    On Error GoTo Error_Handler
    UserInterface.ShowQuoteHistory Me
    Exit Sub
Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "butQuoteHistory_Click", "Main"
End Sub

Private Sub butShowContractsFolder_Click()
    On Error GoTo Error_Handler
    UserInterface.ShowContractsFolder
    Exit Sub
Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "butShowContractsFolder_Click", "Main"
End Sub

Private Sub butSortSearch_Click()
    On Error GoTo Error_Handler
    UserInterface.SortSearchDatabase
    Exit Sub
Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "butSortSearch_Click", "Main"
End Sub

Private Sub CalledThrough_Click()
    On Error GoTo Error_Handler
    UserInterface.MarkQuoteCalledThrough Me
    Exit Sub
Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "CalledThrough_Click", "Main"
End Sub

Private Sub CloseJob_Click()
    On Error GoTo Error_Handler
    If UserInterface.CloseJob(Me) Then
        SystemCore.ShowInformation "Job closed successfully.", "Job Closed"
    End If
    Exit Sub
Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "CloseJob_Click", "Main"
End Sub

Private Sub Enquiries_Click()
    On Error GoTo Error_Handler
    UserInterface.ShowEnquiries Me
    Exit Sub
Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "Enquiries_Click", "Main"
End Sub

Private Sub Quotes_Click()
    On Error GoTo Error_Handler
    UserInterface.ShowQuotes Me
    Exit Sub
Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "Quotes_Click", "Main"
End Sub

Private Sub WIP_Click()
    On Error GoTo Error_Handler
    UserInterface.ShowWIPFiles Me
    Exit Sub
Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "WIP_Click", "Main"
End Sub

Private Sub AcceptQuote_Click()
    On Error GoTo Error_Handler
    UserInterface.AcceptQuote Me
    Exit Sub
Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "AcceptQuote_Click", "Main"
End Sub

Private Sub UserForm_Initialize()
    On Error GoTo Error_Handler
    UserInterface.InitializeMainInterface Me
    Exit Sub
Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "UserForm_Initialize", "Main"
End Sub

Private Sub lst_Click()
    On Error GoTo Error_Handler
    UserInterface.HandleMainListChange
    Exit Sub
Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "lst_Click", "Main"
End Sub

Private Sub Lst_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo Error_Handler
    UserInterface.OpenSelectedFile
    Exit Sub
Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "Lst_DblClick", "Main"
End Sub

Private Sub WIPReport_Click()
    On Error GoTo Error_Handler
    fwip.Show
    Exit Sub
Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "WIPReport_Click", "Main"
End Sub

Private Sub JumpTheGun_Click()
    On Error GoTo Error_Handler
    UserInterface.JumpTheGun
    Exit Sub
Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "JumpTheGun_Click", "Main"
End Sub

Private Sub ContractWork_Click()
    On Error GoTo Error_Handler
    UserInterface.ContractWork
    Exit Sub
Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "ContractWork_Click", "Main"
End Sub

Private Sub OpenJob_Click()
    On Error GoTo Error_Handler
    UserInterface.OpenJob Me
    Exit Sub
Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "OpenJob_Click", "Main"
End Sub

' **Purpose**: Business logic extracted to UserInterface module
' **CLAUDE.md Compliance**: All private functions moved to UserInterface.bas