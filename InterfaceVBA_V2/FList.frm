Private Sub OK_Click()
    On Error GoTo Error_Handler

    If ListManager.ValidateSelection(Me) Then
        Me.Tag = ListManager.GetSelectedValue(Me)
        Unload Me
    End If
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "OK_Click", "FList"
End Sub

Private Sub Cancel_Click()
    Me.Tag = ""
    Unload Me
End Sub

Public Sub ShowListDialog(ListItems As Variant, Title As String)
    On Error GoTo Error_Handler

    ListManager.PopulateList Me, ListItems, Title
    Me.Show
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "ShowListDialog", "FList"
End Sub

' **Purpose**: Business logic extracted to ListManager module
' **CLAUDE.md Compliance**: Generic list functionality moved to ListManager.bas