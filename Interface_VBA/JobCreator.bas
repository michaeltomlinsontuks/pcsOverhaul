Attribute VB_Name = "JobCreator"
' **Purpose**: Job creation and management functions extracted from FJG.frm
' **Original**: FJG.frm.butSaveJG_Click() and related functions
' **Dependencies**: Main module for path, file operations, existing helper functions
' **CLAUDE.md Compliance**: Preserves exact workflow while extracting business logic to module

Option Explicit

' **Purpose**: Create new job from form data
' **Original**: FJG.frm.butSaveJG_Click()
' **Parameters**: JobForm (Object) - The FJG form containing job data
' **Returns**: Boolean - True if job created successfully, False if failed
' **Dependencies**: Main.Main_MasterPath, OpenBook, Calc_Next_Number, Confirm_Next_Number
' **Side Effects**: Creates job files, updates search database, manages compilation sequences
Public Function CreateJob(JobForm As Object) As Boolean
    On Error GoTo Error_Handler

    Dim ctl As Object
    Dim z As Integer
    Dim i As Integer
    Dim j As Integer
    Dim x As Boolean
    Dim xselect As String

    CreateJob = False

    ' Generate job numbers and set status
    JobForm.Enquiry_Number.Value = Calc_Next_Number("E")
    Confirm_Next_Number ("E")

    If JobForm.Compilation_TotalNumber.Value > 1 Then
        If JobForm.Compilation_SequenceNumber.Value = 1 Then
            JobForm.Quote_Number.Value = Calc_Next_Number("Q") & "-1"
            Confirm_Next_Number ("q")
            JobForm.Job_Number.Value = Calc_Next_Number("J") & "-1"
            Confirm_Next_Number ("J")
        Else
            JobForm.Quote_Number.Value = Left(JobForm.Quote_Number.Value, Len(JobForm.Quote_Number.Value) - 2) & "-" & JobForm.Compilation_SequenceNumber.Value
            JobForm.Job_Number.Value = Left(JobForm.Job_Number.Value, Len(JobForm.Job_Number.Value) - 2) & "-" & JobForm.Compilation_SequenceNumber.Value
        End If
    Else
        JobForm.Job_Number.Value = Calc_Next_Number("J")
        Confirm_Next_Number ("J")
        JobForm.Quote_Number.Value = Calc_Next_Number("Q")
        Confirm_Next_Number ("q")
    End If

    JobForm.File_Name.Value = JobForm.Job_Number.Value
    JobForm.System_Status.Value = UCase("Quote Accepted")

    ' Open template and save job data
    xselect = "_Enq"
    x = OpenBook(Main.Main_MasterPath.Value & "Templates\" & xselect & ".xls", True)
    Windows(xselect & ".xls").Activate

    ' Save form data to Admin sheet
    SaveFormDataToWorkbook JobForm, "ADMIN"

    ' Handle job picture if specified
    If JobForm.Job_PicturePath.Value <> "" Then
        InsertJobPicture JobForm
    End If

    ' Save to Search database
    If Not SaveToSearchDatabase(JobForm) Then
        GoTo Error_Handler
    End If

    ' Handle compilation sequences
    If CInt(JobForm.Compilation_SequenceNumber) < CInt(JobForm.Compilation_TotalNumber.Value) Then
        ' Save current job and prepare for next component
        ActiveWorkbook.SaveAs Main.Main_MasterPath & "wip\" & JobForm.File_Name.Value & ".xls"
        ActiveWorkbook.Close True

        ' Clear component-specific fields for next iteration
        ClearComponentFields JobForm

        ' Increment sequence number
        JobForm.Compilation_SequenceNumber.Value = CInt(JobForm.Compilation_SequenceNumber.Value) + 1

        ' Reopen template for next component
        xselect = "_Enq"
        x = OpenBook(Main.Main_MasterPath.Value & "Templates\" & xselect & ".xls", True)
        Windows(xselect & ".xls").Activate

        MsgBox ("Please enter the next components details")
        CreateJob = True
        Exit Function
    End If

    ' Hide form on successful completion
    JobForm.Hide
    CreateJob = True
    Exit Function

Error_Handler:
    MsgBox ("Error creating job: " & Err.Description)
    CreateJob = False
End Function

' **Purpose**: Save form data to workbook Admin sheet
' **Original**: FJG.frm.butSaveJG_Click() inline code
' **Parameters**: JobForm (Object), SheetName (String)
' **Returns**: Nothing
' **Dependencies**: Worksheets object
' **Side Effects**: Updates Admin sheet with form control values
Private Sub SaveFormDataToWorkbook(JobForm As Object, SheetName As String)
    Dim ctl As Object
    Dim i As Integer

    With Worksheets(SheetName)
        For Each ctl In JobForm.Controls
            For i = 0 To 100
                If UCase(.Range("A1").Offset(i, 0).FormulaR1C1) = UCase(ctl.Name) And Left(.Range("A1").Offset(i, 1).Formula, 1) <> "=" Then
                    If UCase(TypeName(ctl)) = "TEXTBOX" Then .Range("A1").Offset(i, 1).FormulaR1C1 = UCase(ctl.Value)
                    If UCase(TypeName(ctl)) = "LABEL" Then .Range("A1").Offset(i, 1).FormulaR1C1 = UCase(ctl.Caption)
                    If UCase(TypeName(ctl)) = "COMBOBOX" Then .Range("A1").Offset(i, 1).FormulaR1C1 = UCase(ctl.Value)
                    GoTo NextControl
                End If
                If UCase(.Range("a1").Offset(i, 0).FormulaR1C1) = "" Then GoTo NextControl
            Next i
NextControl:
        Next ctl
    End With
End Sub

' **Purpose**: Insert job picture into worksheet
' **Original**: FJG.frm.butSaveJG_Click() picture handling code
' **Parameters**: JobForm (Object)
' **Returns**: Nothing
' **Dependencies**: Main.Main_MasterPath, ActiveSheet
' **Side Effects**: Inserts and positions picture in Drawing_Location range
Private Sub InsertJobPicture(JobForm As Object)
    Dim heit As Double

    Sheets("jOB cARD").Select
    Range("Drawing_Location").Select
    heit = Selection.RowHeight * 10
    ActiveSheet.Pictures.Insert(Main.Main_MasterPath.Value & "images\" & JobForm.Job_PicturePath.Value).Select
    With Selection
        .PrintObject = True
        .Name = "Drawing"
        .ShapeRange.Height = heit
        .Left = Range("drawing_location").Left + 5
        .Top = Range("drawing_location").Top + 5
    End With
End Sub

' **Purpose**: Save job data to Search database
' **Original**: FJG.frm.butSaveJG_Click() search database code
' **Parameters**: JobForm (Object)
' **Returns**: Boolean - True if saved successfully
' **Dependencies**: Main.Main_MasterPath, OpenBook
' **Side Effects**: Updates Search.xls with job information
Private Function SaveToSearchDatabase(JobForm As Object) As Boolean
    On Error GoTo Error_Handler

    Dim ctl As Object
    Dim i As Integer
    Dim x As Boolean
    Dim col As Integer

    SaveToSearchDatabase = False

    ' Open Search database
    x = OpenBook(Main.Main_MasterPath & "Search.xls", False)
    Do
        If ActiveWorkbook.ReadOnly = True Then
            ActiveWorkbook.Close
            MsgBox ("This workbook is read only, please find the user with this workbook open and close it.")
            x = OpenBook(Main.Main_MasterPath & "Search.xls", False)
        End If
    Loop Until ActiveWorkbook.ReadOnly = False

    Range("A1").Select
    Selection.End(xlDown).Select

    ' Find next empty row or existing record
    Do
        ActiveCell.Offset(1, 0).Select
    Loop Until ActiveCell.FormulaR1C1 = "" Or _
        ActiveCell.FormulaR1C1 = JobForm.Quote_Number.Value Or _
        ActiveCell.FormulaR1C1 = JobForm.Enquiry_Number.Value Or _
        ActiveCell.FormulaR1C1 = JobForm.Job_Number.Value Or _
        ActiveCell.FormulaR1C1 = JobForm.File_Name.Value

    ' Save form data to search sheet
    With Sheets("search")
        For Each ctl In JobForm.Controls
            For i = 0 To 100
                If UCase(.Range("A1").Offset(0, i).FormulaR1C1) = UCase(ctl.Name) Then
                    If TypeName(ctl) = "Label" Then .Range("A1").Offset(ActiveCell.Row - 1, i).FormulaR1C1 = UCase(ctl.Caption)
                    If UCase(TypeName(ctl)) = "TEXTBOX" Then .Range("A1").Offset(ActiveCell.Row - 1, i).FormulaR1C1 = UCase(ctl.Value)
                    If UCase(TypeName(ctl)) = "COMBOBOX" Then .Range("A1").Offset(ActiveCell.Row - 1, i).FormulaR1C1 = UCase(ctl.Value)
                    GoTo NextSearchControl
                End If
                If Left(.Range("A1").Offset(ActiveCell.Row - 2, i).FormulaR1C1, 1) = "=" Then .Range("A1").Offset(ActiveCell.Row - 1, i).FormulaR1C1 = .Range("A1").Offset(ActiveCell.Row - 2, i).FormulaR1C1
                If UCase(.Range("a1").Offset(0, i + 1).FormulaR1C1) = "" Then GoTo NextSearchControl
            Next i
NextSearchControl:
        Next ctl
    End With

    ' Sort and format search data
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

    SaveToSearchDatabase = True
    Exit Function

Error_Handler:
    SaveToSearchDatabase = False
End Function

' **Purpose**: Clear component-specific fields for next compilation sequence
' **Original**: FJG.frm.butSaveJG_Click() field clearing code
' **Parameters**: JobForm (Object)
' **Returns**: Nothing
' **Dependencies**: None
' **Side Effects**: Clears specific form fields while preserving job-level data
Private Sub ClearComponentFields(JobForm As Object)
    JobForm.Enquiry_Number.Value = ""
    JobForm.File_Name.Value = ""

    JobForm.Component_Code.Value = ""
    JobForm.Component_Grade.Value = ""
    JobForm.Component_Description.Value = ""
    JobForm.Component_DrawingNumber_SampleNumber.Value = ""
    JobForm.Component_Price.Value = ""
    JobForm.Job_PicturePath.Value = ""

    ' Clear all operation fields
    Dim i As Integer
    For i = 1 To 15
        JobForm.Controls("Operation" & Format(i, "00") & "_Comment").Value = ""
        JobForm.Controls("Operation" & Format(i, "00") & "_Operator").Value = ""
        JobForm.Controls("Operation" & Format(i, "00") & "_Type").Value = ""
    Next i
End Sub

' **Purpose**: Copy operation data from existing job
' **Original**: FJG.frm.CopyFromJobCard_Click()
' **Parameters**: JobForm (Object), JobNumber (String)
' **Returns**: Boolean - True if copy successful
' **Dependencies**: Main.Main_MasterPath, OpenBook, Dir function
' **Side Effects**: Populates operation fields from existing job file
Public Function CopyOperationsFromJob(JobForm As Object, JobNumber As String) As Boolean
    On Error GoTo Error_Handler

    Dim ctl As Object
    Dim i As Integer
    Dim x As Boolean

    CopyOperationsFromJob = False

    ' Clear existing operation fields
    For Each ctl In JobForm.Controls
        If InStr(1, Left(UCase(ctl.Name), 6), "OPERAT", vbTextCompare) > 0 Then
            If TypeName(ctl) = "Textbox" Then ctl.Value = ""
            If TypeName(ctl) = "ComboBox" Then ctl.Value = ""
        End If
    Next ctl

    ' Try to find the job file in different folders
    If Dir(Main.Main_MasterPath.Value & "enquiries\" & JobNumber & ".xls", vbNormal) <> "" Then
        x = OpenBook(Main.Main_MasterPath.Value & "Enquiries\" & JobNumber & ".xls", True)
    ElseIf Dir(Main.Main_MasterPath.Value & "quotes\" & JobNumber & ".xls", vbNormal) <> "" Then
        x = OpenBook(Main.Main_MasterPath.Value & "Quotes\" & JobNumber & ".xls", True)
    ElseIf Dir(Main.Main_MasterPath.Value & "archive\" & JobNumber & ".xls", vbNormal) <> "" Then
        x = OpenBook(Main.Main_MasterPath.Value & "Archive\" & JobNumber & ".xls", True)
    ElseIf Dir(Main.Main_MasterPath.Value & "wip\" & JobNumber & ".xls", vbNormal) <> "" Then
        x = OpenBook(Main.Main_MasterPath.Value & "WIP\" & JobNumber & ".xls", True)
    Else
        MsgBox ("File Not Found")
        Exit Function
    End If

    ' Copy operation data from Admin sheet
    With Sheets("Admin")
        For Each ctl In JobForm.Controls
            If InStr(1, Left(UCase(ctl.Name), 6), "OPERAT", vbTextCompare) > 0 Then
                i = -1
                Do
                    i = i + 1
                    If UCase(.Range("A1").Offset(i, 0).Value) = UCase(ctl.Name) Then
                        If TypeName(ctl) = "Label" Then ctl.Caption = Insert_Characters(ctl.Name) & " : " & .Range("A1").Offset(i, 1).Value
                        If UCase(TypeName(ctl)) = "COMBOBOX" Then ctl.Value = .Range("A1").Offset(i, 1).Value
                        If UCase(TypeName(ctl)) = "TEXTBOX" Then ctl.Value = .Range("A1").Offset(i, 1).Value
                        GoTo NextOperationControl
                    End If
                Loop Until .Range("A1").Offset(i, 0).Value = ""
            End If
NextOperationControl:
        Next ctl
    End With

    ActiveWorkbook.Close False
    CopyOperationsFromJob = True
    Exit Function

Error_Handler:
    CopyOperationsFromJob = False
    MsgBox ("Error copying operations: " & Err.Description)
End Function

' **Purpose**: Save job card template as contract item
' **Original**: FJG.frm.but_SaveAsCTItem_Click()
' **Parameters**: JobForm (Object), ContractFileName (String)
' **Returns**: Boolean - True if saved successfully
' **Dependencies**: Main.Main_MasterPath, ActiveWorkbook
' **Side Effects**: Saves current job as contract template file
Public Function SaveAsContractItem(JobForm As Object, ContractFileName As String) As Boolean
    On Error GoTo Error_Handler

    SaveAsContractItem = False

    ' Clear start date for template
    JobForm.Job_StartDate.Value = ""

    ' Save form data to current workbook
    SaveFormDataToWorkbook JobForm, "ADMIN"

    ' Save as contract template
    ActiveWorkbook.SaveAs Main.Main_MasterPath.Value & "Contracts\" & ContractFileName & ".xls"
    ActiveWorkbook.Close True

    SaveAsContractItem = True
    Exit Function

Error_Handler:
    SaveAsContractItem = False
    MsgBox ("Error saving contract item: " & Err.Description)
End Function