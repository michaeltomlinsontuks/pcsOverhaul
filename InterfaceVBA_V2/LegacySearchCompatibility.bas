Attribute VB_Name = "LegacySearchCompatibility"
' **Purpose**: Legacy search compatibility module maintaining exact function signatures
' **CLAUDE.md Compliance**: Preserves all existing function signatures for form button compatibility
Option Explicit

' **Purpose**: Legacy wrapper - Update search database from file system
' **Parameters**: None
' **Returns**: None (Subroutine)
' **Dependencies**: SearchManager.Update_Search
' **Side Effects**: Calls new SearchManager implementation
' **Errors**: Handled by SearchManager
' **CLAUDE.md Compliance**: Maintains exact signature for Module1.bas replacement
Public Sub Update_Search()
    SearchManager.Update_Search
End Sub

' **Purpose**: Legacy wrapper - Get value from closed workbook
' **Parameters**:
'   - Path (String): Directory path to file
'   - File (String): Filename
'   - Sheet (String): Sheet name
'   - Ref (String): Cell reference
' **Returns**: Variant - Cell value or error message
' **Dependencies**: SearchManager.GetValue
' **Side Effects**: None
' **Errors**: Handled by SearchManager
' **CLAUDE.md Compliance**: Maintains exact signature for Module1.bas replacement
Public Function GetValue(ByVal Path As String, ByVal File As String, ByVal Sheet As String, ByVal Ref As String) As Variant
    GetValue = SearchManager.GetValue(Path, File, Sheet, Ref)
End Function

' **Purpose**: Legacy wrapper - Save form data to search database
' **Parameters**:
'   - frm (Object): Form object containing data to save
' **Returns**: None (Subroutine)
' **Dependencies**: SearchManager.SaveRowIntoSearch
' **Side Effects**: Updates search database
' **Errors**: Handled by SearchManager
' **CLAUDE.md Compliance**: Maintains exact signature for SaveSearchCode.bas replacement
Public Sub SaveRowIntoSearch(ByRef frm As Object)
    SearchManager.SaveRowIntoSearch frm
End Sub