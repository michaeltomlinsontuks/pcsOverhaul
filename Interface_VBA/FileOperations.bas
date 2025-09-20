Attribute VB_Name = "FileOperations"
' **Purpose**: All file and workbook operations for the PCS system
' **Consolidates**: Open_Book.bas, SaveFileCode.bas, SaveSearchCode.bas, SaveWIPCode.bas
' **Original Functionality**: All functions preserved exactly as original modules

Option Explicit

' **Purpose**: Open workbook with error handling (original functionality)
' **Parameters**:
'   - File (String): Full path to file to open
'   - RO (Boolean): Open as read-only if True
' **Returns**: None
' **Dependencies**: Excel Workbooks.Open method
' **Original Module**: Open_Book.bas
Public Function OpenBook(File As String, RO As Boolean)
    Workbooks.Open Filename:= _
        File, _
        ReadOnly:=RO
End Function

' **Purpose**: Save form controls to worksheet columns (original functionality)
' **Parameters**: None (uses Me object - must be called from form)
' **Returns**: None
' **Dependencies**: ADMIN worksheet, form controls
' **Side Effects**: Updates ADMIN worksheet, inserts pictures
' **Original Module**: SaveFileCode.bas
Public Function SaveToColumns()
    ' SaveColumnsToFile
    j = -1
    i = 1
    With Worksheets("ADMIN")
        For Each ctl In Me.Controls
            For i = 0 To 100
                    If UCase(.Range("A1").Offset(i, 0).FormulaR1C1) = UCase(ctl.Name) Then
                        If UCase(TypeName(ctl)) = "TEXTBOX" Then .Range("A1").Offset(i, 1).FormulaR1C1 = UCase(ctl.Value)
                        If UCase(TypeName(ctl)) = "LABEL" Then .Range("A1").Offset(i, 1).FormulaR1C1 = UCase(ctl.Caption)
                        If UCase(TypeName(ctl)) = "COMBOBOX" Then .Range("A1").Offset(i, 1).FormulaR1C1 = UCase(ctl.Value)
                        GoTo FormFileNext
                    End If
                    If UCase(.Range("a1").Offset(i, 0).FormulaR1C1) = "" Then GoTo 5
            Next i
FormFileNext:
        Next ctl
    End With

    If Me.Job_PicturePath.Value <> "" Then
        Range("Drawing_location").Select
        heit = Selection.RowHeight * 10
        ActiveSheet.Pictures.Insert(Main.Main_MasterPath.Value & "images\" & Me.Job_PicturePath.Value).Select
        With Selection
            .PrintObject = True
            .Name = "Drawing"
            .ShapeRange.Height = heit
            .Left = Range("drawing_location").Left + 5
            .Top = Range("drawing_location").Top + 5
        End With
    End If

End Function

' **Purpose**: Save form data into search database (original functionality)
' **Parameters**:
'   - frm (Object): Form object containing data to save
' **Returns**: None
' **Dependencies**: OpenBook function, Main.Main_MasterPath, Search.xls
' **Side Effects**: Updates Search.xls file, sorts data
' **Original Module**: SaveSearchCode.bas
Sub SaveRowIntoSearch(frm As Object)
    'Save To Search
    OpenBook (Main.Main_MasterPath & "Search.xls")
        Do
            If ActiveWorkbook.ReadOnly = True Then
                ActiveWorkbook.Close
                MsgBox ("This workbook is read only, please find the user with this workbook open and close it.")
                OpenBook (Main.Main_MasterPath & "Search.xls")
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
End Sub

' **Purpose**: Save form data into WIP database (original functionality)
' **Parameters**:
'   - frm (Object): Form object containing data to save
' **Returns**: None
' **Dependencies**: OpenBook function, Main.Main_MasterPath, WIP.xls
' **Side Effects**: Updates WIP.xls file, clears existing row content
' **Original Module**: SaveWIPCode.bas
Sub SaveInfoIntoWIP(frm As Object)
    ' Save to WIP
    OpenBook (Main.Main_MasterPath & "WIP.xls")
        Do
            If ActiveWorkbook.ReadOnly = True Then
                ActiveWorkbook.Close
                MsgBox ("This workbook is read only, please find the user with this workbook open and close it.")
                OpenBook (Main.Main_MasterPath & "WIP.xls")
            End If
        Loop Until ActiveWorkbook.ReadOnly = False

        Range("A1").Select

        Do
            ActiveCell.Offset(1, 0).Select
        Loop Until ActiveCell.Offset(0, 2).FormulaR1C1 = "" Or _
            ActiveCell.Offset(0, 2).FormulaR1C1 = Me.Quote_Nmber.Value Or _
            ActiveCell.Offset(0, 2).FormulaR1C1 = Me.Enquiry_Number.Value Or _
            ActiveCell.Offset(0, 2).FormulaR1C1 = Me.Job_Number.Value Or _
            ActiveCell.Offset(0, 2).FormulaR1C1 = Me.File_Name.Value

        Selection.EntireRow.ClearContents

        With Sheets(ActiveSheet.Name)
            For Each ctl In Me.Controls
                For i = 0 To 100
                        If UCase(.Range("A1").Offset(0, i).FormulaR1C1) = UCase(ctl.Name) Then
                            If TypeName(ctl) = "Label" Then .Range("A1").Offset(ActiveCell.Row - 1, i).FormulaR1C1 = UCase(ctl.Caption)
                            If UCase(TypeName(ctl)) = "TEXTBOX" Then .Range("A1").Offset(ActiveCell.Row - 1, i).FormulaR1C1 = UCase(ctl.Value)
                            If UCase(TypeName(ctl)) = "COMBOBOX" Then .Range("A1").Offset(ActiveCell.Row - 1, i).FormulaR1C1 = UCase(ctl.Value)
                            GoTo FormNextWIP
                        End If
                        If Left(.Range("A1").Offset(ActiveCell.Row - 2, i).FormulaR1C1, 1) = "=" Then .Range("A1").Offset(ActiveCell.Row - 1, i).FormulaR1C1 = .Range("A1").Offset(ActiveCell.Row - 2, i).FormulaR1C1
                        If UCase(.Range("a1").Offset(0, 1).FormulaR1C1) = "" Then GoTo 6
                Next i
FormNextWIP:
            Next ctl
        End With
End Sub