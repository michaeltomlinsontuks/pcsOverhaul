
' **Purpose**: Validates accept quote form data using standardized popup validation
' **Parameters**: None
' **Returns**: Boolean - True if all validations pass, False if any fail
' **Dependencies**: ValidationFramework
' **Side Effects**: Shows validation popup messages, sets focus to invalid fields
' **Errors**: Returns False on validation failure
Private Function ValidateAcceptQuoteForm() As Boolean
    ValidateAcceptQuoteForm = True

    ' Validate Customer Order Number
    If Not ValidationFramework.ValidateRequired(FAcceptQuote.CustomerOrderNumber.Value, "Customer Order Number", FAcceptQuote.CustomerOrderNumber) Then
        ValidateAcceptQuoteForm = False
        Exit Function
    End If

    ' Validate Compilation Sequence Number
    If Not ValidationFramework.ValidatePositiveNumber(FAcceptQuote.Compilation_SequenceNumber.Value, "Compilation Sequence Number", FAcceptQuote.Compilation_SequenceNumber) Then
        ValidateAcceptQuoteForm = False
        Exit Function
    End If

    ' Validate Job Lead Time if present
    If Trim(FAcceptQuote.Job_LeadTime.Value) <> "" Then
        If Not ValidationFramework.ValidatePositiveNumber(FAcceptQuote.Job_LeadTime.Value, "Job Lead Time", FAcceptQuote.Job_LeadTime) Then
            ValidateAcceptQuoteForm = False
            Exit Function
        End If
    End If
End Function
Private Sub butSAVE_Click()

' Validate form before processing
If Not ValidateAcceptQuoteForm() Then Exit Sub

If CInt(Me.Compilation_SequenceNumber.Value) = "1" Then
    Me.Job_Number.Value = Confirm_Next_Number("J")
    If Me.Compilation_TotalNumber <> "1" Then
        Me.Job_Number.Value = Me.Job_Number.Value & "-1"
    End If
End If

Me.File_Name.Value = Me.Job_Number.Value
Me.System_Status.Value = UCase("Quote Accepted")

' SaveColumnsToFile
x = FileOperations.OpenBook(Main.Main_MasterPath.Value & "Archive\" & Me.Quote_Number.Value & ".xls", False)
j = -1
i = 1
With Worksheets("ADMIN")
    For Each ctl In Me.Controls
        For i = 0 To 100
            If UCase(.Range("A1").Offset(i, 0).FormulaR1C1) = UCase(ctl.Name) And Left(.Range("A1").Offset(i, 1).Formula, 1) <> "=" Then
                If UCase(TypeName(ctl)) = "TEXTBOX" Then .Range("A1").Offset(i, 1).FormulaR1C1 = UCase(ctl.Value)
                If UCase(TypeName(ctl)) = "LABEL" Then .Range("A1").Offset(i, 1).FormulaR1C1 = UCase(ctl.Caption)
                If UCase(TypeName(ctl)) = "COMBOBOX" Then .Range("A1").Offset(i, 1).FormulaR1C1 = UCase(ctl.Value)
                GoTo FormFileNext
            End If
            If UCase(.Range("a1").Offset(i, 0).FormulaR1C1) = "" Then GoTo FormFileNext
        Next i
FormFileNext:
    Next ctl
End With

'Save To Search
x = FileOperations.OpenBook(Main.Main_MasterPath & "Search.xls", False)
    Do
        If ActiveWorkbook.ReadOnly = True Then
            ActiveWorkbook.Close
            MsgBox ("This workbook is read only, please find the user with this workbook open and close it.")
            x = FileOperations.OpenBook(Main.Main_MasterPath & "Search.xls", False)
        End If
    Loop Until ActiveWorkbook.ReadOnly = False

    Range("A1").Select
    
    Do
        ActiveCell.Offset(1, 0).Select
    Loop Until ActiveCell.FormulaR1C1 = "" Or _
        ActiveCell.FormulaR1C1 = Me.Quote_Number.Value Or _
        ActiveCell.FormulaR1C1 = Me.Enquiry_Number.Value Or _
        ActiveCell.FormulaR1C1 = Me.Job_Number.Value Or _
        ActiveCell.FormulaR1C1 = Me.File_Name.Value
    
    With Sheets("search")
        For Each ctl In Me.Controls
            For i = 0 To 100
                If UCase(.Range("A1").Offset(0, i).FormulaR1C1) = UCase(ctl.Name) Then
                    If TypeName(ctl) = "Label" Then .Range("A1").Offset(ActiveCell.Row - 1, i).FormulaR1C1 = UCase(ctl.Caption)
                    If UCase(TypeName(ctl)) = "TEXTBOX" Then .Range("A1").Offset(ActiveCell.Row - 1, i).FormulaR1C1 = UCase(ctl.Value)
                    If UCase(TypeName(ctl)) = "COMBOBOX" Then .Range("A1").Offset(ActiveCell.Row - 1, i).FormulaR1C1 = UCase(ctl.Value)
                    GoTo FormNextSearch
                End If
                If Left(.Range("A1").Offset(ActiveCell.Row - 2, i).FormulaR1C1, 1) = "=" Then .Range("A1").Offset(ActiveCell.Row - 1, i).FormulaR1C1 = .Range("A1").Offset(ActiveCell.Row - 2, i).FormulaR1C1
                
                If UCase(.Range("a1").Offset(0, 1).FormulaR1C1) = "" Then GoTo FormNextSearch
            Next i
FormNextSearch:
        Next ctl
    End With
    
    Range("A1").Select
        Selection.End(xlToRight).Select
        col = ActiveCell.Column
    Range("A1").Select
    Selection.End(xlDown).Select
    Range("A2", Range("A2").Offset(ActiveCell.Row, col - 1).Address).Select
    Selection.Sort Key1:=Range("e2"), Order1:=xlDescending, Header:=xlYes, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortTextAsNumbers
    Range("b3").Select
        
ActiveWorkbook.Close (True)

' Next Compilation
If Me.Compilation_TotalNumber.Value <> "1" And Me.Compilation_SequenceNumber.Value <> Me.Compilation_TotalNumber.Value Then
    ActiveWorkbook.SaveAs (Main.Main_MasterPath.Value & "WIP\" & Me.File_Name.Value & ".xls")
    Me.Compilation_SequenceNumber.Value = CInt(Me.Compilation_SequenceNumber.Value) + 1
    Me.Job_Number.Value = Left(Me.Job_Number.Value, Len(Me.Job_Number.Value) - 2) & "-" & Me.Compilation_SequenceNumber.Value
    
    Me.File_Name.Value = Left(Me.Job_Number.Value, Len(Me.Job_Number.Value) - 2) & "-" & Me.Compilation_SequenceNumber.Value
    
    Me.Component_Code.Value = ""
    Me.Component_Description.Value = ""
    Me.Component_Price.Value = ""
    Me.Component_Grade.Enabled = True
    
    MsgBox ("Please enter the next Parts details")
    Exit Sub
End If

ContFAcceptQuote.Visible = False

    ShowSheet ("Job Card")
    Sheets("Job Card").Select
                
    ActiveWorkbook.SaveAs (Main.Main_MasterPath.Value & "WIP\" & Me.File_Name.Value & ".xls")
    ActiveWorkbook.Close
    Kill (Main.Main_MasterPath.Value & "Archive\" & Me.Quote_Number.Value & ".xls")
    Unload Me
    
End Sub

Private Sub Job_Urgency_Change()

If UCase(Me.Job_Urgency.Value) = "NORMAL" Then Me.Job_LeadTime.Value = "14"
If UCase(Me.Job_Urgency.Value) = "BREAK DOWN" Then Me.Job_LeadTime.Value = "7"
If UCase(Me.Job_Urgency.Value) = "URGENT" Then Me.Job_LeadTime.Value = "10"

End Sub

Private Sub UserForm_Activate()
ContFAcceptQuote.Visible = False
 
If InStr(1, Main.lst.Value, "*") > 1 Then
    xselect = Left(Main.lst.Value, Len(Main.lst.Value) - 2)
Else
    xselect = Main.lst.Value
End If

If FList.lst.Value <> "" Then
    xselect = FList.lst.Value
End If

FList.lst.Value = ""

With FAcceptQuote.Job_Urgency
    .AddItem "NORMAL"
    .AddItem "BREAK DOWN"
    .AddItem "URGENT"
End With

x = FileOperations.OpenBook(Main.Main_MasterPath.Value & "Archive\" & xselect & ".xls", True)
With Sheets("Admin")
    For Each ctl In Me.Controls
        i = -1
        Do
            i = i + 1
            If UCase(.Range("A1").Offset(i, 0).Value) = UCase(ctl.Name) Then
                If InStr(1, ctl.Name, "Price", vbTextCompare) <> 0 Then
                    If UCase(TypeName(ctl)) = "LABEL" Then ctl.Caption = CoreUtilities.Insert_Characters(ctl.Name) & " : " & Format(.Range("A1").Offset(i, 1).Value, "R #,##0.00")
                    If UCase(TypeName(ctl)) = "COMBOBOX" Then ctl.Value = Format(.Range("A1").Offset(i, 1).Value, "R #,##0.00")
                    If UCase(TypeName(ctl)) = "TEXTBOX" Then ctl.Value = Format(.Range("A1").Offset(i, 1).Value, "R #,##0.00")
                    
                    GoTo FormLoadNext
                End If
                
                If UCase(TypeName(ctl)) = "LABEL" Then ctl.Caption = CoreUtilities.Insert_Characters(ctl.Name) & " : " & .Range("A1").Offset(i, 1).Value
                If UCase(TypeName(ctl)) = "COMBOBOX" Then ctl.Value = .Range("A1").Offset(i, 1).Value
                If UCase(TypeName(ctl)) = "TEXTBOX" Then
                    If InStr(1, UCase(ctl.Name), UCase("Date"), vbTextCompare) > 0 Then
                        ctl.Value = Format(.Range("A1").Offset(i, 1).Value, "dd mmm yyyy")
                    Else
                        ctl.Value = .Range("A1").Offset(i, 1).Value
                    End If
                End If
                GoTo FormLoadNext
            End If
        Loop Until .Range("A1").Offset(i, 0).Value = ""
FormLoadNext:
    Next ctl
End With
ActiveWorkbook.Close

Me.System_Status.Value = UCase("Quote Accepted")
Me.Job_StartDate.Value = Format(Now(), "dd mmm yyyy")
Me.CustomerOrderNumber.SetFocus

End Sub

Public Function GetValue(path, File, sheet, ref)
'   Retrieves a value from a closed workbook\
    Dim arg As String

'   Make sure the file exists
    If Right(path, 1) <> "\" Then path = path & "\"
    If Dir(path & File) = "" Then
        GetValue = "File Not Found"
        Exit Function
    End If

'   Create the argument
    arg = "'" & path & "[" & File & "]" & sheet & "'!" & _
      Range(ref).Range("A1").Address(, , xlR1C1)

'   Execute an XLM macro
    GetValue = ExecuteExcel4Macro(arg)
End Function


