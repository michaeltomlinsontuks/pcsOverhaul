Attribute VB_Name = "FList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' **Purpose**: Simple list selection dialog form
' **CLAUDE.md Compliance**: Uses existing form functionality, no new forms created
' **Dependencies**: ValidationFramework for user messages

Private Sub Lst_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ' Double-click to select and close
    FList.Hide
End Sub

Private Sub UserForm_Terminate()
    ' Clean termination
End Sub

' **Purpose**: Initialize list with items
' **Parameters**:
'   - items (Variant): Array of items to display
' **Returns**: None (Subroutine)
' **Dependencies**: None
' **Side Effects**: Populates list control with items
' **Errors**: Handles invalid item arrays gracefully
Public Sub LoadList(ByVal items As Variant)
    Dim i As Integer

    On Error GoTo Error_Handler

    Lst.Clear

    If IsArray(items) Then
        For i = 0 To UBound(items)
            Lst.AddItem items(i)
        Next i
    End If

    Exit Sub

Error_Handler:
    CoreFramework.LogError Err.Number, Err.Description, "LoadList", "FList"
End Sub

' **Purpose**: Get selected item from list
' **Parameters**: None
' **Returns**: String - Selected item text, empty if none selected
' **Dependencies**: None
' **Side Effects**: None
' **Errors**: Returns empty string if no selection
Public Function GetSelectedItem() As String
    On Error GoTo Error_Handler

    If Lst.ListIndex >= 0 Then
        GetSelectedItem = Lst.List(Lst.ListIndex)
    Else
        GetSelectedItem = ""
    End If

    Exit Function

Error_Handler:
    CoreFramework.LogError Err.Number, Err.Description, "GetSelectedItem", "FList"
    GetSelectedItem = ""
End Function

' **Purpose**: Show list dialog and return selected item
' **Parameters**:
'   - items (Variant): Array of items to display
'   - title (String, Optional): Dialog title
' **Returns**: String - Selected item, empty if cancelled
' **Dependencies**: LoadList, ValidationFramework
' **Side Effects**: Shows modal dialog
' **Errors**: Returns empty string on error
Public Function ShowListDialog(ByVal items As Variant, Optional ByVal title As String = "Select Item") As String
    On Error GoTo Error_Handler

    Me.Caption = title
    LoadList items

    Me.Show vbModal

    ShowListDialog = GetSelectedItem()

    Exit Function

Error_Handler:
    CoreFramework.LogError Err.Number, Err.Description, "ShowListDialog", "FList"
    ShowListDialog = ""
End Function