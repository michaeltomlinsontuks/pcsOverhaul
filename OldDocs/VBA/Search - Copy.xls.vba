# VBA code extracted from Search - Copy.xls
# Extraction date: Mon Jun  2 11:09:15 SAST 2025

# =======================================================
# Module 7
# =======================================================
Attribute VB_Name = "Module1"
Sub Show_Search_Menu()
frmSearch.Show
End Sub


# =======================================================
# Module 8
# =======================================================
Attribute VB_Name = "Module2"
Sub Macro1()
Attribute Macro1.VB_Description = "Macro recorded 02/03/2009 by K J Bigham"
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
' Macro recorded 02/03/2009 by K J Bigham
'

'
    Range("A8891").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Sort Key1:=Range("A3"), Order1:=xlDescending, Header:=xlYes, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal
    
    
    Selection.Sort Key1:=Range("A3"), Order1:=xlAscending, Header:=xlYes, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal
        
    Range("A6").Select
    Selection.End(xlUp).Select
    Range("A5").Select
End Sub
Sub Macro2()
Attribute Macro2.VB_Description = "Macro recorded 02/03/2009 by K J Bigham"
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
' Macro recorded 02/03/2009 by K J Bigham
'

'
    Selection.Sort Key1:=Range("A2"), Order1:=xlAscending, Header:=xlYes, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortTextAsNumbers
End Sub


# =======================================================
# Module 9
# =======================================================
Attribute VB_Name = "Module3"
Sub Textify()
    
Range("N1").Select

Do
    ActiveCell.FormulaR1C1 = CStr(ActiveCell.Value)
    ActiveCell.Offset(1, 0).Select
Loop Until Range("A" & ActiveCell.Row).Value = ""
    
End Sub


# =======================================================
# Module 20
# =======================================================
Attribute VB_Name = "frmSearch"
Attribute VB_Base = "0{85CABEE4-1727-4C12-81FF-8014BBB774C3}{A202464A-523B-46A3-86B2-49E4F3AAA278}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
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


