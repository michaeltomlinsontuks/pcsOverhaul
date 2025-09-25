Private Sub OK_Click()
    On Error GoTo Error_Handler

    If UserInterface.ValidateSelection(Me) Then
        Me.Tag = UserInterface.GetSelectedValue(Me)
        Unload Me
    End If
    Exit Sub

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "OK_Click", "FList"
End Sub

Private Sub Cancel_Click()
    Me.Tag = ""
    Unload Me
End Sub

Public Sub ShowListDialog(ListItems As Variant, Title As String)
    On Error GoTo Error_Handler

    UserInterface.PopulateList Me, ListItems, Title
    Me.Show
    Exit Sub

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "ShowListDialog", "FList"
End Sub

' **Purpose**: Business logic extracted to UserInterface module
' **CLAUDE.md Compliance**: Generic list functionality moved to UserInterface.bas
' **V2 Update**: Converted from old ListManager/CoreFramework to UserInterface/SystemCore