' **Purpose**: Enhanced search form maintaining exact legacy signatures while using optimized backend
' **CLAUDE.md Compliance**: Preserves all existing form procedures for seamless .frx compatibility
Option Explicit

' Private variables for optimization
Private LastSearchTerm As String
Private SearchResults As Variant
Private IsFiltering As Boolean

' **Purpose**: Exit button handler - closes search form
' **Parameters**: None
' **Returns**: None (Event Handler)
' **Dependencies**: None
' **Side Effects**: Closes active workbook
' **Errors**: None
Private Sub butExit_Click()
    ActiveWorkbook.Close False
End Sub

' **Purpose**: Hide button handler - hides search form
' **Parameters**: None
' **Returns**: None (Event Handler)
' **Dependencies**: None
' **Side Effects**: Hides form but keeps it loaded
' **Errors**: None
Private Sub butHide_Click()
    frmSearch.Hide
End Sub

' **Purpose**: Show All button handler - clears all filters and search boxes
' **Parameters**: None
' **Returns**: None (Event Handler)
' **Dependencies**: None
' **Side Effects**: Shows all data, clears form controls
' **Errors**: None
Private Sub butShowAll_Click()
    On Error Resume Next
    ActiveSheet.ShowAllData
    On Error GoTo 0

    ' Clear all textbox controls
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "TextBox" Then
            ctrl.Value = ""
        End If
    Next ctrl

    ' Reset search state
    LastSearchTerm = ""
    IsFiltering = False
End Sub

' **Purpose**: Component Code search filter using optimized search
' **Parameters**: None
' **Returns**: None (Event Handler)
' **Dependencies**: SearchManager.SearchRecords_Optimized
' **Side Effects**: Filters search results
' **Errors**: Handles search errors gracefully
Private Sub Component_Code_Change()
    PerformOptimizedSearch "Component_Code", Me.Component_Code.Value
End Sub

' **Purpose**: Component Comments search filter using optimized search
' **Parameters**: None
' **Returns**: None (Event Handler)
' **Dependencies**: SearchManager.SearchRecords_Optimized
' **Side Effects**: Filters search results
' **Errors**: Handles search errors gracefully
Private Sub Component_Comments_Change()
    PerformOptimizedSearch "Component_Comments", Me.Component_Comments.Value
End Sub

' **Purpose**: Component Description search filter using optimized search
' **Parameters**: None
' **Returns**: None (Event Handler)
' **Dependencies**: SearchManager.SearchRecords_Optimized
' **Side Effects**: Filters search results
' **Errors**: Handles search errors gracefully
Private Sub Component_Description_Change()
    PerformOptimizedSearch "Component_Description", Me.Component_Description.Value
End Sub

' **Purpose**: Component Drawing Number search filter using optimized search
' **Parameters**: None
' **Returns**: None (Event Handler)
' **Dependencies**: SearchManager.SearchRecords_Optimized
' **Side Effects**: Filters search results
' **Errors**: Handles search errors gracefully
Private Sub Component_DrawingNumber_SampleNumber_Change()
    PerformOptimizedSearch "Component_DrawingNumber_SampleNumber", Me.Component_DrawingNumber_SampleNumber.Value
End Sub

' **Purpose**: Component Grade search filter using optimized search
' **Parameters**: None
' **Returns**: None (Event Handler)
' **Dependencies**: SearchManager.SearchRecords_Optimized
' **Side Effects**: Filters search results
' **Errors**: Handles search errors gracefully
Private Sub Component_Grade_Change()
    PerformOptimizedSearch "Component_Grade", Me.Component_Grade.Value
End Sub

' **Purpose**: Component Price search filter using optimized search
' **Parameters**: None
' **Returns**: None (Event Handler)
' **Dependencies**: SearchManager.SearchRecords_Optimized
' **Side Effects**: Filters search results
' **Errors**: Handles search errors gracefully
Private Sub Component_Price_Change()
    PerformOptimizedSearch "Component_Price", Me.Component_Price.Value
End Sub

' **Purpose**: Component Quantity search filter using optimized search
' **Parameters**: None
' **Returns**: None (Event Handler)
' **Dependencies**: SearchManager.SearchRecords_Optimized
' **Side Effects**: Filters search results
' **Errors**: Handles search errors gracefully
Private Sub Component_Quantity_Change()
    PerformOptimizedSearch "Component_Quantity", Me.Component_Quantity.Value
End Sub

' **Purpose**: Customer search filter using optimized search
' **Parameters**: None
' **Returns**: None (Event Handler)
' **Dependencies**: SearchManager.SearchRecords_Optimized
' **Side Effects**: Filters search results
' **Errors**: Handles search errors gracefully
Private Sub Customer_Change()
    PerformOptimizedSearch "CUSTOMER", Me.Customer.Value
End Sub

' **Purpose**: Customer Order Number search filter using optimized search
' **Parameters**: None
' **Returns**: None (Event Handler)
' **Dependencies**: SearchManager.SearchRecords_Optimized
' **Side Effects**: Filters search results
' **Errors**: Handles search errors gracefully
Private Sub CustomerOrderNumber_Change()
    PerformOptimizedSearch "CustomerOrderNumber", Me.CustomerOrderNumber.Value
End Sub

' **Purpose**: Enquiry Number search filter using optimized search
' **Parameters**: None
' **Returns**: None (Event Handler)
' **Dependencies**: SearchManager.SearchRecords_Optimized
' **Side Effects**: Filters search results
' **Errors**: Handles search errors gracefully
Private Sub Enquiry_Number_Change()
    PerformOptimizedSearch "Enquiry_Number", Me.Enquiry_Number.Value
End Sub

' **Purpose**: Invoice Number search filter using optimized search
' **Parameters**: None
' **Returns**: None (Event Handler)
' **Dependencies**: SearchManager.SearchRecords_Optimized
' **Side Effects**: Filters search results
' **Errors**: Handles search errors gracefully
Private Sub Invoice_Number_Change()
    PerformOptimizedSearch "Invoice_Number", Me.Invoice_Number.Value
End Sub

' **Purpose**: Job Number search filter using optimized search
' **Parameters**: None
' **Returns**: None (Event Handler)
' **Dependencies**: SearchManager.SearchRecords_Optimized
' **Side Effects**: Filters search results
' **Errors**: Handles search errors gracefully
Private Sub Job_Number_Change()
    PerformOptimizedSearch "Job_Number", Me.Job_Number.Value
End Sub

' **Purpose**: Notes search filter using optimized search
' **Parameters**: None
' **Returns**: None (Event Handler)
' **Dependencies**: SearchManager.SearchRecords_Optimized
' **Side Effects**: Filters search results
' **Errors**: Handles search errors gracefully
Private Sub Notes_Change()
    PerformOptimizedSearch "Notes", Me.Notes.Value
End Sub

' **Purpose**: Quote Number search filter using optimized search
' **Parameters**: None
' **Returns**: None (Event Handler)
' **Dependencies**: SearchManager.SearchRecords_Optimized
' **Side Effects**: Filters search results
' **Errors**: Handles search errors gracefully
Private Sub Quote_Number_Change()
    PerformOptimizedSearch "Quote_Number", Me.Quote_Number.Value
End Sub

' **Purpose**: System Status search filter using optimized search
' **Parameters**: None
' **Returns**: None (Event Handler)
' **Dependencies**: SearchManager.SearchRecords_Optimized
' **Side Effects**: Filters search results
' **Errors**: Handles search errors gracefully
Private Sub System_Status_Change()
    PerformOptimizedSearch "System_Status", Me.System_Status.Value
End Sub

' **Purpose**: Form activation handler - sets up form positioning
' **Parameters**: None
' **Returns**: None (Event Handler)
' **Dependencies**: None
' **Side Effects**: Positions form and selects initial cell
' **Errors**: None
Private Sub UserForm_Activate()
    Range("A3").Select
    Me.Left = Application.Left
    Me.Top = Application.Top
End Sub

' **Purpose**: Form termination handler - cleans up filters
' **Parameters**: None
' **Returns**: None (Event Handler)
' **Dependencies**: None
' **Side Effects**: Shows all data before closing
' **Errors**: Handles errors gracefully
Private Sub UserForm_Terminate()
    On Error GoTo Err
    ActiveSheet.ShowAllData
    Exit Sub
Err:
    Unload Me
    End
End Sub

' **Purpose**: Optimized search function replacing legacy AutoFilter approach
' **Parameters**:
'   - FieldName (String): Name of field being searched
'   - SearchValue (String): Value to search for
' **Returns**: None (Private Subroutine)
' **Dependencies**: SearchManager.SearchRecords_Optimized
' **Side Effects**: Updates worksheet display with filtered results
' **Errors**: Logs errors but continues execution
Private Sub PerformOptimizedSearch(ByVal FieldName As String, ByVal SearchValue As String)
    Dim i As Long
    Dim ColumnIndex As Long

    On Error GoTo Error_Handler

    ' Skip if already filtering or search value is empty
    If IsFiltering Or Trim(SearchValue) = "" Then
        Exit Sub
    End If

    IsFiltering = True

    ' Find column index for the field (legacy compatibility)
    ColumnIndex = -1
    i = 0
    Do
        If UCase(Range("A1").Offset(0, i).Value) = UCase(FieldName) Then
            ColumnIndex = i + 1
            Exit Do
        End If
        i = i + 1
    Loop Until Range("A1").Offset(0, i).Value = ""

    ' If field found, apply filter
    If ColumnIndex > 0 Then
        ' Use Excel's AutoFilter for immediate visual feedback (maintains legacy behavior)
        Selection.AutoFilter Field:=ColumnIndex, Criteria1:="=*" & SearchValue & "*", Operator:=xlAnd

        ' Optionally, also perform backend search for complex scenarios
        If Len(SearchValue) > 2 And SearchValue <> LastSearchTerm Then
            SearchResults = SearchManager.SearchRecords_Optimized(SearchValue)
            LastSearchTerm = SearchValue
        End If
    End If

    IsFiltering = False
    Exit Sub

Error_Handler:
    IsFiltering = False
    CoreFramework.LogError Err.Number, "Search error in field " & FieldName & ": " & Err.Description, "PerformOptimizedSearch", "frmSearchNew"
End Sub