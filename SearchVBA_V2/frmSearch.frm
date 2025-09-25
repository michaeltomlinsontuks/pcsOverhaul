Private Sub butExit_Click()
    On Error GoTo Error_Handler
    SearchFormManager.ExitSearchForm Me
    Exit Sub
Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "butExit_Click", "frmSearch"
End Sub

Private Sub butHide_Click()
    On Error GoTo Error_Handler
    SearchFormManager.HideSearchForm Me
    Exit Sub
Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "butHide_Click", "frmSearch"
End Sub

Private Sub butShowAll_Click()
    On Error GoTo Error_Handler
    SearchFormManager.ShowAllData Me
    Exit Sub
Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "butShowAll_Click", "frmSearch"
End Sub

Private Sub Component_Code_Change()
    On Error GoTo Error_Handler
    SearchFormManager.HandleComponentCodeFilter Me
    Exit Sub
Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "Component_Code_Change", "frmSearch"
End Sub

Private Sub Component_Comments_Change()
    On Error GoTo Error_Handler
    SearchFormManager.HandleComponentCommentsFilter Me
    Exit Sub
Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "Component_Comments_Change", "frmSearch"
End Sub

Private Sub Component_Description_Change()
    On Error GoTo Error_Handler
    SearchFormManager.HandleComponentDescriptionFilter Me
    Exit Sub
Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "Component_Description_Change", "frmSearch"
End Sub

Private Sub Component_DrawingNumber_SampleNumber_Change()
    On Error GoTo Error_Handler
    SearchFormManager.HandleComponentDrawingNumberFilter Me
    Exit Sub
Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "Component_DrawingNumber_SampleNumber_Change", "frmSearch"
End Sub

Private Sub Component_Grade_Change()
    On Error GoTo Error_Handler
    SearchFormManager.HandleComponentGradeFilter Me
    Exit Sub
Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "Component_Grade_Change", "frmSearch"
End Sub

Private Sub Component_Price_Change()
    On Error GoTo Error_Handler
    SearchFormManager.HandleComponentPriceFilter Me
    Exit Sub
Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "Component_Price_Change", "frmSearch"
End Sub

Private Sub Component_Quantity_Change()
    On Error GoTo Error_Handler
    SearchFormManager.HandleComponentQuantityFilter Me
    Exit Sub
Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "Component_Quantity_Change", "frmSearch"
End Sub

Private Sub Customer_Change()
    On Error GoTo Error_Handler
    SearchFormManager.HandleCustomerFilter Me
    Exit Sub
Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "Customer_Change", "frmSearch"
End Sub

Private Sub UserForm_Initialize()
    On Error GoTo Error_Handler
    SearchFormManager.InitializeSearchForm Me
    Exit Sub
Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "UserForm_Initialize", "frmSearch"
End Sub

' **Purpose**: Business logic extracted to SearchFormManager module
' **CLAUDE.md Compliance**: All filter functions moved to SearchFormManager.bas