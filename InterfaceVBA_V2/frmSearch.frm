Private Sub butExit_Click()
    On Error GoTo Error_Handler
    ActiveWorkbook.Close False
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "butExit_Click", "frmSearch"
End Sub

Private Sub butHide_Click()
    On Error GoTo Error_Handler
    frmSearch.Hide
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "butHide_Click", "frmSearch"
End Sub

Private Sub butShowAll_Click()
    On Error GoTo Error_Handler

    ActiveSheet.ShowAllData

    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "TextBox" Then
            ctrl.Value = ""
        End If
    Next ctrl
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "butShowAll_Click", "frmSearch"
End Sub

Private Sub Component_Code_Change()
    On Error GoTo Error_Handler
    ApplyAutoFilter "Component_Code"
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "Component_Code_Change", "frmSearch"
End Sub

Private Sub Component_Comments_Change()
    On Error GoTo Error_Handler
    ApplyAutoFilter "Component_Comments"
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "Component_Comments_Change", "frmSearch"
End Sub

Private Sub Component_Description_Change()
    On Error GoTo Error_Handler
    ApplyAutoFilter "Component_Description"
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "Component_Description_Change", "frmSearch"
End Sub

Private Sub Component_DrawingNumber_SampleNumber_Change()
    On Error GoTo Error_Handler
    ApplyAutoFilter "Component_DrawingNumber_SampleNumber"
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "Component_DrawingNumber_SampleNumber_Change", "frmSearch"
End Sub

Private Sub Component_Grade_Change()
    On Error GoTo Error_Handler
    ApplyAutoFilter "Component_Grade"
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "Component_Grade_Change", "frmSearch"
End Sub

Private Sub Component_Price_Change()
    On Error GoTo Error_Handler
    ApplyAutoFilter "Component_Price"
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "Component_Price_Change", "frmSearch"
End Sub

Private Sub Component_Quantity_Change()
    On Error GoTo Error_Handler
    ApplyAutoFilter "Component_Quantity"
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "Component_Quantity_Change", "frmSearch"
End Sub

Private Sub Customer_Change()
    On Error GoTo Error_Handler
    ApplyAutoFilter "CUSTOMER"
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "Customer_Change", "frmSearch"
End Sub

Private Sub CustomerOrderNumber_Change()
    On Error GoTo Error_Handler
    ApplyAutoFilter "CustomerOrderNumber"
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "CustomerOrderNumber_Change", "frmSearch"
End Sub

Private Sub Enquiry_Number_Change()
    On Error GoTo Error_Handler
    ApplyAutoFilter "Enquiry_Number"
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "Enquiry_Number_Change", "frmSearch"
End Sub

Private Sub Invoice_Number_Change()
    On Error GoTo Error_Handler
    ApplyAutoFilter "Invoice_Number"
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "Invoice_Number_Change", "frmSearch"
End Sub

Private Sub Job_Number_Change()
    On Error GoTo Error_Handler
    ApplyAutoFilter "Job_Number"
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "Job_Number_Change", "frmSearch"
End Sub

Private Sub Notes_Change()
    On Error GoTo Error_Handler
    ApplyAutoFilter "Notes"
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "Notes_Change", "frmSearch"
End Sub

Private Sub Quote_Number_Change()
    On Error GoTo Error_Handler
    ApplyAutoFilter "Quote_Number"
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "Quote_Number_Change", "frmSearch"
End Sub

Private Sub System_Status_Change()
    On Error GoTo Error_Handler
    ApplyAutoFilter "System_Status"
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "System_Status_Change", "frmSearch"
End Sub

Private Sub UserForm_Activate()
    On Error GoTo Error_Handler

    Range("A3").Select
    Me.Left = Application.Left
    Me.Top = Application.Top
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "UserForm_Activate", "frmSearch"
End Sub

Private Sub UserForm_Terminate()
    On Error GoTo Error_Handler

    ActiveSheet.ShowAllData
    Exit Sub

Error_Handler:
    Unload Me
    End
End Sub

' Helper function to consolidate the repetitive AutoFilter logic
Private Sub ApplyAutoFilter(ByVal FieldName As String)
    Dim varib As String
    Dim i As Integer

    On Error GoTo Error_Handler

    varib = UCase(FieldName)
    i = -1

    Do
        i = i + 1
        If UCase(Range("a1").Offset(0, i).Value) = varib Then
            Selection.AutoFilter Field:=i + 1, Criteria1:="=*" & Me.Controls(varib).Value & "*", Operator:=xlAnd
            Exit Sub
        End If
    Loop Until Range("a1").Offset(0, i + 1).Value = ""
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "ApplyAutoFilter", "frmSearch"
End Sub