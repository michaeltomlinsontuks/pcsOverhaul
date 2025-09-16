# VBA code extracted from Search.xls
# Extraction date: Mon Jun  2 11:09:06 SAST 2025

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
Attribute VB_Base = "0{29105004-67C0-4A12-8183-22B84CE7A890}{A9B75E24-879D-482D-A416-D12782D84151}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private colMap As Object

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
    If Not colMap Is Nothing Then
        If colMap.Exists(UCase("Component_Code")) Then
            Selection.AutoFilter Field:=colMap.Item(UCase("Component_Code")), Criteria1:="=*" & Me.Controls("Component_Code").Value & "*", Operator:=xlAnd
        End If
    End If
End Sub

Private Sub Component_Comments_Change()
    If Not colMap Is Nothing Then
        If colMap.Exists(UCase("Component_Comments")) Then
            Selection.AutoFilter Field:=colMap.Item(UCase("Component_Comments")), Criteria1:="=*" & Me.Controls("Component_Comments").Value & "*", Operator:=xlAnd
        End If
    End If
End Sub

Private Sub Component_Description_Change()
    If Not colMap Is Nothing Then
        If colMap.Exists(UCase("Component_Description")) Then
            Selection.AutoFilter Field:=colMap.Item(UCase("Component_Description")), Criteria1:="=*" & Me.Controls("Component_Description").Value & "*", Operator:=xlAnd
        End If
    End If
End Sub

Private Sub Component_DrawingNumber_SampleNumber_Change()
    If Not colMap Is Nothing Then
        If colMap.Exists(UCase("Component_DrawingNumber_SampleNumber")) Then
            Selection.AutoFilter Field:=colMap.Item(UCase("Component_DrawingNumber_SampleNumber")), Criteria1:="=*" & Me.Controls("Component_DrawingNumber_SampleNumber").Value & "*", Operator:=xlAnd
        End If
    End If
End Sub

Private Sub Component_Grade_Change()
    If Not colMap Is Nothing Then
        If colMap.Exists(UCase("Component_Grade")) Then
            Selection.AutoFilter Field:=colMap.Item(UCase("Component_Grade")), Criteria1:="=*" & Me.Controls("Component_Grade").Value & "*", Operator:=xlAnd
        End If
    End If
End Sub

Private Sub Component_Price_Change()
    If Not colMap Is Nothing Then
        If colMap.Exists(UCase("Component_Price")) Then
            Selection.AutoFilter Field:=colMap.Item(UCase("Component_Price")), Criteria1:="=*" & Me.Controls("Component_Price").Value & "*", Operator:=xlAnd
        End If
    End If
End Sub

Private Sub Component_Quantity_Change()
    If Not colMap Is Nothing Then
        If colMap.Exists(UCase("Component_Quantity")) Then
            Selection.AutoFilter Field:=colMap.Item(UCase("Component_Quantity")), Criteria1:="=*" & Me.Controls("Component_Quantity").Value & "*", Operator:=xlAnd
        End If
    End If
End Sub

Private Sub Customer_Change()
    If Not colMap Is Nothing Then
        If colMap.Exists(UCase("CUSTOMER")) Then
            Selection.AutoFilter Field:=colMap.Item(UCase("CUSTOMER")), Criteria1:="=*" & Me.Controls("CUSTOMER").Value & "*", Operator:=xlAnd
        End If
    End If
End Sub

Private Sub CustomerOrderNumber_Change()
    If Not colMap Is Nothing Then
        If colMap.Exists(UCase("CustomerOrderNumber")) Then
            Selection.AutoFilter Field:=colMap.Item(UCase("CustomerOrderNumber")), Criteria1:="=*" & Me.Controls("CustomerOrderNumber").Value & "*", Operator:=xlAnd
        End If
    End If
End Sub

Private Sub Enquiry_Number_Change()
    If Not colMap Is Nothing Then
        If colMap.Exists(UCase("Enquiry_Number")) Then
            Selection.AutoFilter Field:=colMap.Item(UCase("Enquiry_Number")), Criteria1:="=*" & Me.Controls("Enquiry_Number").Value & "*", Operator:=xlAnd
        End If
    End If
End Sub

Private Sub Invoice_Number_Change()
    If Not colMap Is Nothing Then
        If colMap.Exists(UCase("Invoice_Number")) Then
            Selection.AutoFilter Field:=colMap.Item(UCase("Invoice_Number")), Criteria1:="=*" & Me.Controls("Invoice_Number").Value & "*", Operator:=xlAnd
        End If
    End If
End Sub

Private Sub Job_Number_Change()
    If Not colMap Is Nothing Then
        If colMap.Exists(UCase("Job_Number")) Then
            Selection.AutoFilter Field:=colMap.Item(UCase("Job_Number")), Criteria1:="=*" & Me.Controls("Job_Number").Value & "*", Operator:=xlAnd
        End If
    End If
End Sub

Private Sub Notes_Change()
    If Not colMap Is Nothing Then
        If colMap.Exists(UCase("Notes")) Then
            Selection.AutoFilter Field:=colMap.Item(UCase("Notes")), Criteria1:="=*" & Me.Controls("Notes").Value & "*", Operator:=xlAnd
        End If
    End If
End Sub

Private Sub Quote_Number_Change()
    If Not colMap Is Nothing Then
        If colMap.Exists(UCase("Quote_Number")) Then
            Selection.AutoFilter Field:=colMap.Item(UCase("Quote_Number")), Criteria1:="=*" & Me.Controls("Quote_Number").Value & "*", Operator:=xlAnd
        End If
    End If
End Sub

Private Sub System_Status_Change()
    If Not colMap Is Nothing Then
        If colMap.Exists(UCase("System_Status")) Then
            Selection.AutoFilter Field:=colMap.Item(UCase("System_Status")), Criteria1:="=*" & Me.Controls("System_Status").Value & "*", Operator:=xlAnd
        End If
    End If
End Sub

Private Sub UserForm_Activate()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Set colMap = CreateObject("Scripting.Dictionary")
    Dim i As Long
    Dim headerCell As Range
    
    ' Assuming headers are in Row 1
    Set headerCell = ActiveSheet.Range("A1")
    i = 1
    Do Until IsEmpty(headerCell.Offset(0, i - 1).Value)
        colMap.Add UCase(headerCell.Offset(0, i - 1).Value), i
        i = i + 1
    Loop
    
    Range("A3").Select
    Me.Left = Application.Left
    Me.Top = Application.Top
End Sub

Private Sub UserForm_Terminate()
    On Error GoTo Err
    ActiveSheet.ShowAllData
    Set colMap = Nothing ' Release memory
    Application.ScreenUpdating = True
    Application.EnableEvents = True
Exit Sub
Err:
    Unload Me
    End
End Sub
