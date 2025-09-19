Attribute VB_Name = "SearchModules"
' **Purpose**: Search supporting modules maintaining exact legacy signatures
' **CLAUDE.md Compliance**: Preserves all existing module procedures for seamless compatibility
Option Explicit

' **Purpose**: Show search menu - exact signature match for legacy compatibility
' **Parameters**: None
' **Returns**: None (Subroutine)
' **Dependencies**: frmSearchNew form
' **Side Effects**: Shows search form
' **Errors**: None
' **CLAUDE.md Compliance**: Maintains exact signature for Module1.bas replacement
Public Sub Show_Search_Menu()
    frmSearchNew.Show
End Sub

' **Purpose**: Legacy macro for sorting search data (ascending)
' **Parameters**: None
' **Returns**: None (Subroutine)
' **Dependencies**: None
' **Side Effects**: Sorts selected range in ascending order
' **Errors**: Handles selection errors
' **CLAUDE.md Compliance**: Maintains exact signature for Module2.bas replacement
Public Sub Macro1()
    On Error GoTo Error_Handler

    ' Enhanced version with error handling
    Range("A8891").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select

    ' Sort descending first (legacy behavior)
    Selection.Sort Key1:=Range("A3"), Order1:=xlDescending, Header:=xlYes, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal

    ' Then sort ascending (legacy behavior)
    Selection.Sort Key1:=Range("A3"), Order1:=xlAscending, Header:=xlYes, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal

    Range("A6").Select
    Selection.End(xlUp).Select
    Range("A5").Select
    Exit Sub

Error_Handler:
    CoreFramework.LogError Err.Number, "Error in search sort macro: " & Err.Description, "Macro1", "SearchModules"
End Sub

' **Purpose**: Legacy macro for sorting search data (text as numbers)
' **Parameters**: None
' **Returns**: None (Subroutine)
' **Dependencies**: None
' **Side Effects**: Sorts selected range treating text as numbers
' **Errors**: Handles selection errors
' **CLAUDE.md Compliance**: Maintains exact signature for Module2.bas replacement
Public Sub Macro2()
    On Error GoTo Error_Handler

    Selection.Sort Key1:=Range("A2"), Order1:=xlAscending, Header:=xlYes, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortTextAsNumbers
    Exit Sub

Error_Handler:
    CoreFramework.LogError Err.Number, "Error in search sort macro 2: " & Err.Description, "Macro2", "SearchModules"
End Sub

' **Purpose**: Legacy macro to convert numbers to text in column N
' **Parameters**: None
' **Returns**: None (Subroutine)
' **Dependencies**: None
' **Side Effects**: Converts column N values to text strings
' **Errors**: Handles conversion errors
' **CLAUDE.md Compliance**: Maintains exact signature for Module3.bas replacement
Public Sub Textify()
    On Error GoTo Error_Handler

    Range("N1").Select

    Do
        ActiveCell.FormulaR1C1 = CStr(ActiveCell.Value)
        ActiveCell.Offset(1, 0).Select
    Loop Until Range("A" & ActiveCell.Row).Value = ""
    Exit Sub

Error_Handler:
    CoreFramework.LogError Err.Number, "Error in textify macro: " & Err.Description, "Textify", "SearchModules"
End Sub

' **Purpose**: Enhanced search with performance optimizations
' **Parameters**:
'   - SearchTerm (String): Term to search for
'   - MaxResults (Long, Optional): Maximum results to return (default 100)
' **Returns**: Variant - Array of search results
' **Dependencies**: SearchManager.SearchRecords_Optimized
' **Side Effects**: May update search database if needed
' **Errors**: Returns empty array on error
' **CLAUDE.md Compliance**: Enhanced functionality while maintaining compatibility
Public Function PerformEnhancedSearch(ByVal SearchTerm As String, Optional ByVal MaxResults As Long = 100) As Variant
    On Error GoTo Error_Handler

    ' Use optimized search from SearchManager
    Dim Results As Variant
    Results = SearchManager.SearchRecords_Optimized(SearchTerm)

    ' Limit results if requested
    If IsArray(Results) And UBound(Results) >= 0 Then
        If UBound(Results) + 1 > MaxResults Then
            Dim LimitedResults() As CoreFramework.SearchRecord
            ReDim LimitedResults(MaxResults - 1)

            Dim i As Long
            For i = 0 To MaxResults - 1
                LimitedResults(i) = Results(i)
            Next i

            PerformEnhancedSearch = LimitedResults
        Else
            PerformEnhancedSearch = Results
        End If
    Else
        PerformEnhancedSearch = Array()
    End If
    Exit Function

Error_Handler:
    CoreFramework.LogError Err.Number, "Error in enhanced search: " & Err.Description, "PerformEnhancedSearch", "SearchModules"
    PerformEnhancedSearch = Array()
End Function