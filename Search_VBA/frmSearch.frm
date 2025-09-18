Private Sub butExit_Click()
ActiveWorkbook.Close False
End Sub

Private Sub butHide_Click()
frmSearch.Hide
End Sub

Private Sub butShowAll_Click()
ActiveSheet.ShowAllData
For Each ctrl In Me.Controls
    If TypeName(ctrl) = "TextBox" Then
        ctrl.Value = ""
    End If
Next ctrl

End Sub

Private Sub Component_Code_Change()
varib = UCase("Component_Code")

i = -1
Do
    i = i + 1
    If UCase(Range("a1").Offset(0, i).Value) = varib Then
        Selection.AutoFilter Field:=i + 1, Criteria1:="=*" & Me.Controls(varib).Value & "*", Operator:=xlAnd
        Exit Sub
    End If
Loop Until Range("a1").Offset(0, i + 1).Value = ""

End Sub

Private Sub Component_Comments_Change()
varib = UCase("Component_Comments")

i = -1
Do
    i = i + 1
    If UCase(Range("a1").Offset(0, i).Value) = varib Then
        Selection.AutoFilter Field:=i + 1, Criteria1:="=*" & Me.Controls(varib).Value & "*", Operator:=xlAnd
        Exit Sub
    End If
Loop Until Range("a1").Offset(0, i + 1).Value = ""

End Sub

Private Sub Component_Description_Change()
varib = UCase("Component_Description")

i = -1
Do
    i = i + 1
    If UCase(Range("a1").Offset(0, i).Value) = varib Then
        Selection.AutoFilter Field:=i + 1, Criteria1:="=*" & Me.Controls(varib).Value & "*", Operator:=xlAnd
        Exit Sub
    End If
Loop Until Range("a1").Offset(0, i + 1).Value = ""

End Sub

Private Sub Component_DrawingNumber_SampleNumber_Change()
varib = UCase("Component_DrawingNumber_SampleNumber")

i = -1
Do
    i = i + 1
    If UCase(Range("a1").Offset(0, i).Value) = varib Then
        Selection.AutoFilter Field:=i + 1, Criteria1:="=*" & Me.Controls(varib).Value & "*", Operator:=xlAnd
        Exit Sub
    End If
Loop Until Range("a1").Offset(0, i + 1).Value = ""

End Sub

Private Sub Component_Grade_Change()
varib = UCase("Component_Grade")

i = -1
Do
    i = i + 1
    If UCase(Range("a1").Offset(0, i).Value) = varib Then
        Selection.AutoFilter Field:=i + 1, Criteria1:="=*" & Me.Controls(varib).Value & "*", Operator:=xlAnd
        Exit Sub
    End If
Loop Until Range("a1").Offset(0, i + 1).Value = ""

End Sub

Private Sub Component_Price_Change()
varib = UCase("Component_Price")

i = -1
Do
    i = i + 1
    If UCase(Range("a1").Offset(0, i).Value) = varib Then
        Selection.AutoFilter Field:=i + 1, Criteria1:="=*" & Me.Controls(varib).Value & "*", Operator:=xlAnd
        Exit Sub
    End If
Loop Until Range("a1").Offset(0, i + 1).Value = ""

End Sub

Private Sub Component_Quantity_Change()
varib = UCase("Component_Quantity")

i = -1
Do
    i = i + 1
    If UCase(Range("a1").Offset(0, i).Value) = varib Then
        Selection.AutoFilter Field:=i + 1, Criteria1:="=*" & Me.Controls(varib).Value & "*", Operator:=xlAnd
        Exit Sub
    End If
Loop Until Range("a1").Offset(0, i + 1).Value = ""

End Sub

Private Sub Customer_Change()
varib = UCase("CUSTOMER")

i = -1
Do
    i = i + 1
    If UCase(Range("a1").Offset(0, i).Value) = varib Then
        Selection.AutoFilter Field:=i + 1, Criteria1:="=*" & Me.Controls(varib).Value & "*", Operator:=xlAnd
        Exit Sub
    End If
Loop Until Range("a1").Offset(0, i + 1).Value = ""

End Sub

Private Sub CustomerOrderNumber_Change()
varib = UCase("CustomerOrderNumber")

i = -1
Do
    i = i + 1
    If UCase(Range("a1").Offset(0, i).Value) = varib Then
        Selection.AutoFilter Field:=i + 1, Criteria1:="=*" & Me.Controls(varib).Value & "*", Operator:=xlAnd
        Exit Sub
    End If
Loop Until Range("a1").Offset(0, i + 1).Value = ""

End Sub

Private Sub Enquiry_Number_Change()
varib = UCase("Enquiry_Number")

i = -1
Do
    i = i + 1
    If UCase(Range("a1").Offset(0, i).Value) = varib Then
        Selection.AutoFilter Field:=i + 1, Criteria1:="=*" & Me.Controls(varib).Value & "*", Operator:=xlAnd
        Exit Sub
    End If
Loop Until Range("a1").Offset(0, i + 1).Value = ""

End Sub

Private Sub Invoice_Number_Change()
varib = UCase("Invoice_Number")

i = -1
Do
    i = i + 1
    If UCase(Range("a1").Offset(0, i).Value) = varib Then
        Selection.AutoFilter Field:=i + 1, Criteria1:="=*" & Me.Controls(varib).Value & "*", Operator:=xlAnd
        Exit Sub
    End If
Loop Until Range("a1").Offset(0, i + 1).Value = ""

End Sub

Private Sub Job_Number_Change()
varib = UCase("Job_Number")

i = -1
Do
    i = i + 1
    If UCase(Range("a1").Offset(0, i).Value) = varib Then
        Selection.AutoFilter Field:=i + 1, Criteria1:="=*" & Me.Controls(varib).Value & "*", Operator:=xlAnd
        Exit Sub
    End If
Loop Until Range("a1").Offset(0, i + 1).Value = ""

End Sub

Private Sub Notes_Change()
varib = UCase("Notes")

i = -1
Do
    i = i + 1
    If UCase(Range("a1").Offset(0, i).Value) = varib Then
        Selection.AutoFilter Field:=i + 1, Criteria1:="=*" & Me.Controls(varib).Value & "*", Operator:=xlAnd
        Exit Sub
    End If
Loop Until Range("a1").Offset(0, i + 1).Value = ""

End Sub

Private Sub Quote_Number_Change()
varib = UCase("Quote_Number")

i = -1
Do
    i = i + 1
    If UCase(Range("a1").Offset(0, i).Value) = varib Then
        Selection.AutoFilter Field:=i + 1, Criteria1:="=*" & Me.Controls(varib).Value & "*", Operator:=xlAnd
        Exit Sub
    End If
Loop Until Range("a1").Offset(0, i + 1).Value = ""

End Sub

Private Sub System_Status_Change()
varib = UCase("System_Status")

i = -1
Do
    i = i + 1
    If UCase(Range("a1").Offset(0, i).Value) = varib Then
        Selection.AutoFilter Field:=i + 1, Criteria1:="=*" & Me.Controls(varib).Value & "*", Operator:=xlAnd
        Exit Sub
    End If
Loop Until Range("a1").Offset(0, i + 1).Value = ""

End Sub

Private Sub UserForm_Activate()
Range("A3").Select
Me.Left = Application.Left
Me.Top = Application.Top

End Sub

Private Sub UserForm_Terminate()
On Error GoTo Err
ActiveSheet.ShowAllData
Exit Sub
Err:
    Unload Me
    End

End Sub
