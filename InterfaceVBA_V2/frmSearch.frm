Private SearchWB As Workbook
Private SearchWS As Worksheet

Private Sub butExit_Click()
    On Error Resume Next
    If Not SearchWB Is Nothing Then
        SearchWB.Close False
    End If
    Me.Hide
End Sub

Private Sub butHide_Click()
    Me.Hide
End Sub

Private Sub butShowAll_Click()
    On Error Resume Next
    If Not SearchWS Is Nothing Then
        SearchWS.ShowAllData
    End If
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "TextBox" Then
            ctrl.Value = ""
        End If
    Next ctrl
End Sub

Private Sub Component_Code_Change()
    If SearchWS Is Nothing Then Exit Sub

    varib = UCase("Component_Code")

    i = -1
    Do
        i = i + 1
        If UCase(SearchWS.Range("a1").Offset(0, i).Value) = varib Then
            SearchWS.Range("A1").CurrentRegion.AutoFilter Field:=i + 1, Criteria1:="=*" & Me.Controls(varib).Value & "*", Operator:=xlAnd
            Exit Sub
        End If
    Loop Until SearchWS.Range("a1").Offset(0, i + 1).Value = ""
End Sub

Private Sub Component_Comments_Change()
    If SearchWS Is Nothing Then Exit Sub

    varib = UCase("Component_Comments")

    i = -1
    Do
        i = i + 1
        If UCase(SearchWS.Range("a1").Offset(0, i).Value) = varib Then
            SearchWS.Range("A1").CurrentRegion.AutoFilter Field:=i + 1, Criteria1:="=*" & Me.Controls(varib).Value & "*", Operator:=xlAnd
            Exit Sub
        End If
    Loop Until SearchWS.Range("a1").Offset(0, i + 1).Value = ""
End Sub

Private Sub Component_Description_Change()
    If SearchWS Is Nothing Then Exit Sub

    varib = UCase("Component_Description")

    i = -1
    Do
        i = i + 1
        If UCase(SearchWS.Range("a1").Offset(0, i).Value) = varib Then
            SearchWS.Range("A1").CurrentRegion.AutoFilter Field:=i + 1, Criteria1:="=*" & Me.Controls(varib).Value & "*", Operator:=xlAnd
            Exit Sub
        End If
    Loop Until SearchWS.Range("a1").Offset(0, i + 1).Value = ""
End Sub

Private Sub Component_DrawingNumber_SampleNumber_Change()
    If SearchWS Is Nothing Then Exit Sub

    varib = UCase("Component_DrawingNumber_SampleNumber")

    i = -1
    Do
        i = i + 1
        If UCase(SearchWS.Range("a1").Offset(0, i).Value) = varib Then
            SearchWS.Range("A1").CurrentRegion.AutoFilter Field:=i + 1, Criteria1:="=*" & Me.Controls(varib).Value & "*", Operator:=xlAnd
            Exit Sub
        End If
    Loop Until SearchWS.Range("a1").Offset(0, i + 1).Value = ""
End Sub

Private Sub Component_Grade_Change()
    If SearchWS Is Nothing Then Exit Sub

    varib = UCase("Component_Grade")

    i = -1
    Do
        i = i + 1
        If UCase(SearchWS.Range("a1").Offset(0, i).Value) = varib Then
            SearchWS.Range("A1").CurrentRegion.AutoFilter Field:=i + 1, Criteria1:="=*" & Me.Controls(varib).Value & "*", Operator:=xlAnd
            Exit Sub
        End If
    Loop Until SearchWS.Range("a1").Offset(0, i + 1).Value = ""
End Sub

Private Sub Component_Price_Change()
    If SearchWS Is Nothing Then Exit Sub

    varib = UCase("Component_Price")

    i = -1
    Do
        i = i + 1
        If UCase(SearchWS.Range("a1").Offset(0, i).Value) = varib Then
            SearchWS.Range("A1").CurrentRegion.AutoFilter Field:=i + 1, Criteria1:="=*" & Me.Controls(varib).Value & "*", Operator:=xlAnd
            Exit Sub
        End If
    Loop Until SearchWS.Range("a1").Offset(0, i + 1).Value = ""
End Sub

Private Sub Component_Quantity_Change()
    If SearchWS Is Nothing Then Exit Sub

    varib = UCase("Component_Quantity")

    i = -1
    Do
        i = i + 1
        If UCase(SearchWS.Range("a1").Offset(0, i).Value) = varib Then
            SearchWS.Range("A1").CurrentRegion.AutoFilter Field:=i + 1, Criteria1:="=*" & Me.Controls(varib).Value & "*", Operator:=xlAnd
            Exit Sub
        End If
    Loop Until SearchWS.Range("a1").Offset(0, i + 1).Value = ""
End Sub

Private Sub Customer_Change()
    If SearchWS Is Nothing Then Exit Sub

    varib = UCase("CUSTOMER")

    i = -1
    Do
        i = i + 1
        If UCase(SearchWS.Range("a1").Offset(0, i).Value) = varib Then
            SearchWS.Range("A1").CurrentRegion.AutoFilter Field:=i + 1, Criteria1:="=*" & Me.Controls(varib).Value & "*", Operator:=xlAnd
            Exit Sub
        End If
    Loop Until SearchWS.Range("a1").Offset(0, i + 1).Value = ""
End Sub

Private Sub CustomerOrderNumber_Change()
    If SearchWS Is Nothing Then Exit Sub

    varib = UCase("CustomerOrderNumber")

    i = -1
    Do
        i = i + 1
        If UCase(SearchWS.Range("a1").Offset(0, i).Value) = varib Then
            SearchWS.Range("A1").CurrentRegion.AutoFilter Field:=i + 1, Criteria1:="=*" & Me.Controls(varib).Value & "*", Operator:=xlAnd
            Exit Sub
        End If
    Loop Until SearchWS.Range("a1").Offset(0, i + 1).Value = ""
End Sub

Private Sub Enquiry_Number_Change()
    If SearchWS Is Nothing Then Exit Sub

    varib = UCase("Enquiry_Number")

    i = -1
    Do
        i = i + 1
        If UCase(SearchWS.Range("a1").Offset(0, i).Value) = varib Then
            SearchWS.Range("A1").CurrentRegion.AutoFilter Field:=i + 1, Criteria1:="=*" & Me.Controls(varib).Value & "*", Operator:=xlAnd
            Exit Sub
        End If
    Loop Until SearchWS.Range("a1").Offset(0, i + 1).Value = ""
End Sub

Private Sub Invoice_Number_Change()
    If SearchWS Is Nothing Then Exit Sub

    varib = UCase("Invoice_Number")

    i = -1
    Do
        i = i + 1
        If UCase(SearchWS.Range("a1").Offset(0, i).Value) = varib Then
            SearchWS.Range("A1").CurrentRegion.AutoFilter Field:=i + 1, Criteria1:="=*" & Me.Controls(varib).Value & "*", Operator:=xlAnd
            Exit Sub
        End If
    Loop Until SearchWS.Range("a1").Offset(0, i + 1).Value = ""
End Sub

Private Sub Job_Number_Change()
    If SearchWS Is Nothing Then Exit Sub

    varib = UCase("Job_Number")

    i = -1
    Do
        i = i + 1
        If UCase(SearchWS.Range("a1").Offset(0, i).Value) = varib Then
            SearchWS.Range("A1").CurrentRegion.AutoFilter Field:=i + 1, Criteria1:="=*" & Me.Controls(varib).Value & "*", Operator:=xlAnd
            Exit Sub
        End If
    Loop Until SearchWS.Range("a1").Offset(0, i + 1).Value = ""
End Sub

Private Sub Notes_Change()
    If SearchWS Is Nothing Then Exit Sub

    varib = UCase("Notes")

    i = -1
    Do
        i = i + 1
        If UCase(SearchWS.Range("a1").Offset(0, i).Value) = varib Then
            SearchWS.Range("A1").CurrentRegion.AutoFilter Field:=i + 1, Criteria1:="=*" & Me.Controls(varib).Value & "*", Operator:=xlAnd
            Exit Sub
        End If
    Loop Until SearchWS.Range("a1").Offset(0, i + 1).Value = ""
End Sub

Private Sub Quote_Number_Change()
    If SearchWS Is Nothing Then Exit Sub

    varib = UCase("Quote_Number")

    i = -1
    Do
        i = i + 1
        If UCase(SearchWS.Range("a1").Offset(0, i).Value) = varib Then
            SearchWS.Range("A1").CurrentRegion.AutoFilter Field:=i + 1, Criteria1:="=*" & Me.Controls(varib).Value & "*", Operator:=xlAnd
            Exit Sub
        End If
    Loop Until SearchWS.Range("a1").Offset(0, i + 1).Value = ""
End Sub

Private Sub System_Status_Change()
    If SearchWS Is Nothing Then Exit Sub

    varib = UCase("System_Status")

    i = -1
    Do
        i = i + 1
        If UCase(SearchWS.Range("a1").Offset(0, i).Value) = varib Then
            SearchWS.Range("A1").CurrentRegion.AutoFilter Field:=i + 1, Criteria1:="=*" & Me.Controls(varib).Value & "*", Operator:=xlAnd
            Exit Sub
        End If
    Loop Until SearchWS.Range("a1").Offset(0, i + 1).Value = ""
End Sub

Private Sub UserForm_Activate()
    On Error GoTo Error_Handler

    ' Open the search database using V2 FileManager
    Set SearchWB = DataManager.SafeOpenWorkbook(DataManager.GetRootPath & "\Search.xls")
    If SearchWB Is Nothing Then
        MsgBox "Could not open search database.", vbCritical
        Me.Hide
        Exit Sub
    End If

    Set SearchWS = SearchWB.Worksheets(1)

    ' Sort by date (recent files first) when opening
    SearchManager.SortSearchDatabase

    ' Set form position
    Me.Left = Application.Left
    Me.Top = Application.Top

    ' Select starting cell
    If Not SearchWS Is Nothing Then
        SearchWS.Range("A3").Select
    End If
    Exit Sub

Error_Handler:
    MsgBox "Error initializing search form: " & Err.Description, vbCritical
    Me.Hide
End Sub

Private Sub UserForm_Terminate()
    On Error GoTo Err
    If Not SearchWS Is Nothing Then
        SearchWS.ShowAllData
    End If
    If Not SearchWB Is Nothing Then
        DataManager.SafeCloseWorkbook SearchWB, False
    End If
    Exit Sub
Err:
    On Error Resume Next
    If Not SearchWB Is Nothing Then
        SearchWB.Close False
    End If
    Set SearchWS = Nothing
    Set SearchWB = Nothing
End Sub
