Private Sub butExit_Click()
    On Error GoTo Error_Handler
    SearchFormManager.ExitSearchForm Me
    Exit Sub
Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "butExit_Click", "frmSearchNew"
End Sub

Private Sub butHide_Click()
    On Error GoTo Error_Handler
    SearchFormManager.HideSearchForm Me
    Exit Sub
Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "butHide_Click", "frmSearchNew"
End Sub

Private Sub butShowAll_Click()
    On Error GoTo Error_Handler
    SearchFormManager.ShowAllData Me
    Exit Sub
Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "butShowAll_Click", "frmSearchNew"
End Sub

Private Sub butAdvancedSearch_Click()
    On Error GoTo Error_Handler
    Dim SearchCriteria As Collection
    Set SearchCriteria = New Collection

    ' Add search criteria based on form inputs
    If Me.txtSearchTerm.Value <> "" Then
        SearchCriteria.Add SearchFormManager.CreateSearchCriterion("Description", "Contains", Me.txtSearchTerm.Value)
    End If

    Dim Results As Variant
    Results = SearchFormManager.PerformAdvancedSearch(Me, SearchCriteria)

    If IsArray(Results) Then
        ValidationFramework.ShowInformation "Found " & UBound(Results, 1) + 1 & " results", "Search Complete"
    Else
        ValidationFramework.ShowInformation "No results found", "Search Complete"
    End If
    Exit Sub
Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "butAdvancedSearch_Click", "frmSearchNew"
End Sub

Private Sub butClearAll_Click()
    On Error GoTo Error_Handler
    SearchFormManager.ClearAllFilters Me
    Exit Sub
Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "butClearAll_Click", "frmSearchNew"
End Sub

Private Sub butExportResults_Click()
    On Error GoTo Error_Handler
    Dim Results As Variant
    Results = SearchManager.SearchRecords(Me.txtSearchTerm.Value, 0)

    If IsArray(Results) Then
        Dim ExportPath As String
        ExportPath = DataManager.GetRootPath & "\SearchResults_" & Format(Now, "yyyymmdd_hhmmss") & ".xls"

        If SearchFormManager.ExportSearchResults(Results, ExportPath) Then
            ValidationFramework.ShowInformation "Results exported to: " & ExportPath, "Export Complete"
        Else
            ValidationFramework.ShowError "Export failed", "Export Error"
        End If
    Else
        ValidationFramework.ShowWarning "No results to export", "Export Warning"
    End If
    Exit Sub
Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "butExportResults_Click", "frmSearchNew"
End Sub

Private Sub UserForm_Initialize()
    On Error GoTo Error_Handler
    SearchFormManager.InitializeSearchForm Me
    Exit Sub
Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "UserForm_Initialize", "frmSearchNew"
End Sub

' **Purpose**: Business logic extracted to SearchFormManager module
' **CLAUDE.md Compliance**: Advanced search functions moved to SearchFormManager.bas