Attribute VB_Name = "JobCardManager"
' **Purpose**: Job card management functions extracted from FJobCard.frm
' **Original**: FJobCard.frm.SaveJobCard_Click() and related functions
' **Dependencies**: Main module for path, ValidationFramework, file operations
' **CLAUDE.md Compliance**: Preserves exact workflow while extracting business logic to module

Option Explicit

' **Purpose**: Save job card and move to archive
' **Original**: FJobCard.frm.SaveJobCard_Click()
' **Parameters**: JobCardForm (Object) - The FJobCard form containing job card data
' **Returns**: Boolean - True if job card saved successfully, False if failed
' **Dependencies**: Main.Main_MasterPath, OpenBook, ValidationFramework
' **Side Effects**: Creates archive file, updates WIP and Search databases, removes from WIP folder
Public Function SaveJobCard(JobCardForm As Object) As Boolean
    On Error GoTo Error_Handler

    Dim Missed(1 To 100) As Integer
    Dim xselect As String
    Dim ctl As Object
    Dim i As Integer
    Dim j As Integer
    Dim x As Boolean
    Dim col As Integer

    SaveJobCard = False

    ' Validate form before processing
    If Not ValidateJobCardForm(JobCardForm) Then Exit Function

    ' Determine selected job file
    If InStr(1, Main.lst.Value, "*") > 1 Then
        xselect = Left(Main.lst.Value, Len(Main.lst.Value) - 2)
    Else
        xselect = Main.lst.Value
    End If

    ' Generate job number if not specified
    If JobCardForm.Job_Number.Value = "" Then
        JobCardForm.Job_Number.Value = Confirm_Next_Number("J")
    End If

    ' Set default start date if not specified
    If JobCardForm.Job_StartDate = "" Then
        JobCardForm.Job_StartDate = Format(CDate(Now()), "dd-mmm-yyyy")
    End If

    JobCardForm.File_Name.Value = JobCardForm.Job_Number.Value

    ' Open and update the WIP file
    x = OpenBook(Main.Main_MasterPath.Value & "WIP\" & xselect & ".xls", False)
    Windows(xselect & ".xls").Activate

    Sheets("Admin").Select
    JobCardForm.System_Status.Value = UCase("Job Open")

    Sheets("Job Card").Select

    ' Save form data to Admin sheet
    SaveJobCardDataToWorkbook JobCardForm, "ADMIN"

    ' Insert job picture if specified
    If JobCardForm.Job_PicturePath.Value <> "" Then
        InsertJobCardPicture JobCardForm
    End If

    Sheets("Job Card").Select
    Range("A1").Select
    Range("r3").FormulaR1C1 = ""

    ' Update WIP database
    If Not UpdateWIPDatabase(JobCardForm) Then
        GoTo Error_Handler
    End If

    ' Move to Archive and update Search
    If Not MoveJobToArchive(JobCardForm, xselect) Then
        GoTo Error_Handler
    End If

    ' Update Search database
    If Not UpdateSearchDatabase(JobCardForm) Then
        GoTo Error_Handler
    End If

    ' Unload form and open archived job
    Unload JobCardForm
    x = OpenBook(Main.Main_MasterPath.Value & "Archive\" & JobCardForm.Job_Number.Value & ".xls", False)
    Unload Main

    SaveJobCard = True
    Exit Function

Error_Handler:
    MsgBox ("Error saving job card: " & Err.Description)
    SaveJobCard = False
End Function

' **Purpose**: Validate job card form data
' **Original**: FJobCard.frm.ValidateJobCardForm()
' **Parameters**: JobCardForm (Object)
' **Returns**: Boolean - True if validation passes
' **Dependencies**: ValidationFramework
' **Side Effects**: Shows validation messages, sets focus to invalid fields
Private Function ValidateJobCardForm(JobCardForm As Object) As Boolean
    ValidateJobCardForm = True

    ' Validate Job Number
    If Not ValidationFramework.ValidateRequired(JobCardForm.Job_Number.Value, "Job Number", JobCardForm.Job_Number) Then
        ValidateJobCardForm = False
        Exit Function
    End If

    ' Validate Job Start Date
    If Not ValidationFramework.ValidateRequired(JobCardForm.Job_StartDate.Value, "Job Start Date", JobCardForm.Job_StartDate) Then
        ValidateJobCardForm = False
        Exit Function
    End If

    ' Validate at least one operation is specified
    If Not ValidateOperationsExist(JobCardForm) Then
        ValidateJobCardForm = False
        Exit Function
    End If
End Function

' **Purpose**: Validate that at least one operation is specified
' **Original**: FJobCard.frm.ValidateOperationsExist()
' **Parameters**: JobCardForm (Object)
' **Returns**: Boolean - True if operations exist
' **Dependencies**: ValidationFramework
' **Side Effects**: Shows validation popup if no operations found
Private Function ValidateOperationsExist(JobCardForm As Object) As Boolean
    Dim i As Integer
    Dim hasOperations As Boolean

    ValidateOperationsExist = True
    hasOperations = False

    ' Check first 15 operations for any content
    For i = 1 To 15
        Dim operationType As String
        operationType = JobCardForm.Controls("Operation" & Format(i, "00") & "_Type").Value
        If Trim(operationType) <> "" Then
            hasOperations = True
            Exit For
        End If
    Next i

    If Not hasOperations Then
        ValidationFramework.ShowWarning "At least one operation must be specified for this job.", "Operations Required"
        ValidateOperationsExist = False
    End If
End Function

' **Purpose**: Save job card form data to workbook Admin sheet
' **Original**: FJobCard.frm.SaveJobCard_Click() inline save code
' **Parameters**: JobCardForm (Object), SheetName (String)
' **Returns**: Nothing
' **Dependencies**: Worksheets object
' **Side Effects**: Updates Admin sheet with job card data
Private Sub SaveJobCardDataToWorkbook(JobCardForm As Object, SheetName As String)
    Dim ctl As Object
    Dim i As Integer

    With Worksheets(SheetName)
        For Each ctl In JobCardForm.Controls
            For i = 0 To 100
                If UCase(.Range("A1").Offset(i, 0).FormulaR1C1) = UCase(ctl.Name) And Left(.Range("A1").Offset(i, 0).Formula, 1) <> "=" Then
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

' **Purpose**: Insert job picture into job card
' **Original**: FJobCard.frm.SaveJobCard_Click() picture handling
' **Parameters**: JobCardForm (Object)
' **Returns**: Nothing
' **Dependencies**: Main.Main_MasterPath, ActiveSheet
' **Side Effects**: Inserts and positions picture in Drawing_location range
Private Sub InsertJobCardPicture(JobCardForm As Object)
    Dim heit As Double

    Range("Drawing_location").Select
    heit = Selection.RowHeight * 10
    ActiveSheet.Pictures.Insert(Main.Main_MasterPath.Value & "images\" & JobCardForm.Job_PicturePath.Value).Select
    With Selection
        .PrintObject = True
        .Name = "Drawing"
        .ShapeRange.Height = heit
        .Left = Range("drawing_location").Left + 5
        .Top = Range("drawing_location").Top + 5
    End With
End Sub

' **Purpose**: Update WIP database with job card data
' **Original**: FJobCard.frm.SaveJobCard_Click() WIP database update code
' **Parameters**: JobCardForm (Object)
' **Returns**: Boolean - True if update successful
' **Dependencies**: Main.Main_MasterPath, OpenBook
' **Side Effects**: Updates WIP.xls with job information and sorts data
Private Function UpdateWIPDatabase(JobCardForm As Object) As Boolean
    On Error GoTo Error_Handler

    Dim ctl As Object
    Dim i As Integer
    Dim x As Boolean
    Dim col As Integer

    UpdateWIPDatabase = False

    ' Open WIP database
    x = OpenBook(Main.Main_MasterPath & "WIP.xls", False)
    Do
        If ActiveWorkbook.ReadOnly = True Then
            ActiveWorkbook.Close
            MsgBox ("This workbook is read only, please find the user with this workbook open and close it.")
            x = OpenBook(Main.Main_MasterPath & "WIP.xls", False)
        End If
    Loop Until ActiveWorkbook.ReadOnly = False

    Range("A1").Select

    ' Find existing record or create new one
    Do
        ActiveCell.Offset(1, 0).Select
    Loop Until ActiveCell.Offset(0, 2).FormulaR1C1 = "" Or _
        ActiveCell.Offset(0, 2).FormulaR1C1 = JobCardForm.Quote_Number.Value Or _
        ActiveCell.Offset(0, 2).FormulaR1C1 = JobCardForm.Enquiry_Number.Value Or _
        ActiveCell.Offset(0, 2).FormulaR1C1 = JobCardForm.Job_Number.Value Or _
        ActiveCell.Offset(0, 2).FormulaR1C1 = JobCardForm.File_Name.Value

    Selection.EntireRow.ClearContents

    ' Save job card data to WIP sheet
    With Sheets(ActiveSheet.Name)
        For Each ctl In JobCardForm.Controls
            For i = 0 To 100
                If UCase(.Range("A1").Offset(0, i).FormulaR1C1) = UCase(ctl.Name) Then
                    If TypeName(ctl) = "Label" Then .Range("A1").Offset(ActiveCell.Row - 1, i).FormulaR1C1 = UCase(ctl.Caption)
                    If UCase(TypeName(ctl)) = "TEXTBOX" Then .Range("A1").Offset(ActiveCell.Row - 1, i).FormulaR1C1 = UCase(ctl.Value)
                    If UCase(TypeName(ctl)) = "COMBOBOX" Then .Range("A1").Offset(ActiveCell.Row - 1, i).FormulaR1C1 = UCase(ctl.Value)
                    GoTo NextWIPControl
                End If
                If Left(.Range("A1").Offset(ActiveCell.Row - 2, i).FormulaR1C1, 1) = "=" Then .Range("A1").Offset(ActiveCell.Row - 1, i).FormulaR1C1 = .Range("A1").Offset(ActiveCell.Row - 2, i).FormulaR1C1
                If UCase(.Range("a1").Offset(0, 1).FormulaR1C1) = "" Then GoTo NextWIPControl
            Next i
NextWIPControl:
        Next ctl
    End With

    ' Sort WIP data
    Range("A1").Select
    Selection.End(xlToRight).Select
    col = ActiveCell.Column

    Range("A1").Select
    Selection.End(xlDown).Select

    Range("A2", Range("A2").Offset(ActiveCell.Row, col - 1).Address).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Sort Key1:=Range("h3"), Order1:=xlAscending, Header:=xlYes, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom

    ActiveWorkbook.Close (True)
    UpdateWIPDatabase = True
    Exit Function

Error_Handler:
    UpdateWIPDatabase = False
End Function

' **Purpose**: Move completed job to Archive folder
' **Original**: FJobCard.frm.SaveJobCard_Click() archive code
' **Parameters**: JobCardForm (Object), SelectedFile (String)
' **Returns**: Boolean - True if move successful
' **Dependencies**: Main.Main_MasterPath, ActiveWorkbook, Kill function
' **Side Effects**: Saves job to Archive folder, removes from WIP folder
Private Function MoveJobToArchive(JobCardForm As Object, SelectedFile As String) As Boolean
    On Error GoTo Error_Handler

    MoveJobToArchive = False

    ' Check if already in Archive
    If UCase(ActiveWorkbook.path) = UCase(Main.Main_MasterPath.Value & "Archive") Then
        ActiveWorkbook.Close (True)
    Else
        ' Save to Archive and remove from WIP
        ActiveWorkbook.SaveAs (Main.Main_MasterPath.Value & "Archive\" & JobCardForm.Job_Number.Value & ".xls")
        ActiveWorkbook.Close
        Kill (Main.Main_MasterPath & "WIP\" & SelectedFile & ".xls")
    End If

    MoveJobToArchive = True
    Exit Function

Error_Handler:
    MoveJobToArchive = False
End Function

' **Purpose**: Update Search database with completed job
' **Original**: FJobCard.frm.SaveJobCard_Click() search update code
' **Parameters**: JobCardForm (Object)
' **Returns**: Boolean - True if update successful
' **Dependencies**: Main.Main_MasterPath, OpenBook
' **Side Effects**: Updates Search.xls with job completion data
Private Function UpdateSearchDatabase(JobCardForm As Object) As Boolean
    On Error GoTo Error_Handler

    Dim ctl As Object
    Dim i As Integer
    Dim x As Boolean
    Dim col As Integer

    UpdateSearchDatabase = False

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

    ' Find existing record using search
    Columns("A:A").Select
    Selection.Find(What:=JobCardForm.File_Name.Value, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate

    ' Update search record
    With Sheets("search")
        For Each ctl In JobCardForm.Controls
            For i = 0 To 100
                If UCase(.Range("A1").Offset(0, i).FormulaR1C1) = UCase(ctl.Name) Then
                    If TypeName(ctl) = "Label" Then .Range("A1").Offset(ActiveCell.Row - 1, i).FormulaR1C1 = UCase(ctl.Caption)
                    If UCase(TypeName(ctl)) = "TEXTBOX" Then .Range("A1").Offset(ActiveCell.Row - 1, i).FormulaR1C1 = UCase(ctl.Value)
                    If UCase(TypeName(ctl)) = "COMBOBOX" Then .Range("A1").Offset(ActiveCell.Row - 1, i).FormulaR1C1 = UCase(ctl.Value)
                    GoTo NextSearchControl
                End If
                If Left(.Range("A1").Offset(ActiveCell.Row - 2, i).FormulaR1C1, 1) = "=" Then .Range("A1").Offset(ActiveCell.Row - 1, i).FormulaR1C1 = .Range("A1").Offset(ActiveCell.Row - 2, i).FormulaR1C1
                If UCase(.Range("a1").Offset(0, 1).FormulaR1C1) = "" Then GoTo NextSearchControl
            Next i
NextSearchControl:
        Next ctl
    End With

    ' Sort search data
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

    UpdateSearchDatabase = True
    Exit Function

Error_Handler:
    UpdateSearchDatabase = False
End Function

' **Purpose**: Copy job card data from existing job
' **Original**: FJobCard.frm.CopyFromJobCard_Click()
' **Parameters**: JobCardForm (Object), JobNumber (String)
' **Returns**: Boolean - True if copy successful
' **Dependencies**: Main.Main_MasterPath, OpenBook, Dir function
' **Side Effects**: Populates form fields from existing job file, clears form first
Public Function CopyFromJobCard(JobCardForm As Object, JobNumber As String) As Boolean
    On Error GoTo Error_Handler

    Dim ctl As Object
    Dim i As Integer
    Dim x As Boolean

    CopyFromJobCard = False

    ' Clear all form fields first
    For Each ctl In JobCardForm.Controls
        If TypeName(ctl) = "Label" Then ctl.Caption = ""
        If TypeName(ctl) = "Textbox" Then ctl.Value = ""
        If TypeName(ctl) = "ComboBox" Then ctl.Value = ""
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

    ' Copy data from Admin sheet
    With Sheets("Admin")
        For Each ctl In JobCardForm.Controls
            If UCase(ctl.Name) = "JOB_NUMBER" Then GoTo NextControl
            If UCase(ctl.Name) = "ENQUIRY_NUMBER" Then GoTo NextControl
            If UCase(ctl.Name) = "QUOTE_NUMBER" Then GoTo NextControl
            If UCase(ctl.Name) = "FILE_NAME" Then GoTo NextControl

            i = -1
            Do
                i = i + 1
                If UCase(.Range("A1").Offset(i, 0).Value) = UCase(ctl.Name) And UCase(ctl.Name) = "JOB_PICTUREPATH" Then
                    If TypeName(ctl) = "Label" Then ctl.Caption = Insert_Characters(ctl.Name) & " : " & .Range("A1").Offset(i, 1).Value
                    If UCase(TypeName(ctl)) = "COMBOBOX" Then ctl.Value = .Range("A1").Offset(i, 1).Value
                    If UCase(TypeName(ctl)) = "TEXTBOX" Then ctl.Value = .Range("A1").Offset(i, 1).Value
                    GoTo NextControl
                End If
                If UCase(.Range("A1").Offset(i, 0).Value) = UCase(ctl.Name) And Left(UCase(ctl.Name), 9) = "OPERATION" Then
                    If TypeName(ctl) = "Label" Then ctl.Caption = Insert_Characters(ctl.Name) & " : " & .Range("A1").Offset(i, 1).Value
                    If UCase(TypeName(ctl)) = "COMBOBOX" Then ctl.Value = .Range("A1").Offset(i, 1).Value
                    If UCase(TypeName(ctl)) = "TEXTBOX" Then ctl.Value = .Range("A1").Offset(i, 1).Value
                    GoTo NextControl
                End If
            Loop Until .Range("A1").Offset(i, 0).Value = ""
NextControl:
        Next ctl
    End With

    ActiveWorkbook.Close False

    ' Set current date and clear certain fields
    JobCardForm.Job_StartDate.Value = Format(Now(), "dd mmm yyyy")

    CopyFromJobCard = True
    Exit Function

Error_Handler:
    CopyFromJobCard = False
    MsgBox ("Error copying job card: " & Err.Description)
End Function

' **Purpose**: Load job card template operations
' **Original**: FJobCard.frm.JobCardTemplates_Click()
' **Parameters**: JobCardForm (Object), TemplateFileName (String)
' **Returns**: Boolean - True if template loaded successfully
' **Dependencies**: Main.Main_MasterPath, GetValue function, FList form
' **Side Effects**: Populates operation fields from job template file
Public Function LoadJobCardTemplate(JobCardForm As Object, TemplateFileName As String) As Boolean
    On Error GoTo Error_Handler

    LoadJobCardTemplate = False

    ' Clear existing operations
    RefreshJobCardOperations JobCardForm

    ' Load template operations
    With JobCardForm
        .Operation01_Type.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", TemplateFileName & ".xls", "JC Seq", "A2")
        .Operation02_Type.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", TemplateFileName & ".xls", "JC Seq", "A3")
        .Operation03_Type.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", TemplateFileName & ".xls", "JC Seq", "A4")
        .Operation04_Type.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", TemplateFileName & ".xls", "JC Seq", "A5")
        .Operation05_Type.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", TemplateFileName & ".xls", "JC Seq", "A6")
        .Operation06_Type.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", TemplateFileName & ".xls", "JC Seq", "A7")
        .Operation07_Type.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", TemplateFileName & ".xls", "JC Seq", "A8")
        .Operation08_Type.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", TemplateFileName & ".xls", "JC Seq", "A9")
        .Operation09_Type.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", TemplateFileName & ".xls", "JC Seq", "A10")
        .Operation10_Type.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", TemplateFileName & ".xls", "JC Seq", "A11")
        .Operation11_Type.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", TemplateFileName & ".xls", "JC Seq", "A12")
        .Operation12_Type.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", TemplateFileName & ".xls", "JC Seq", "A13")
        .Operation13_Type.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", TemplateFileName & ".xls", "JC Seq", "A14")
        .Operation14_Type.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", TemplateFileName & ".xls", "JC Seq", "A15")
        .Operation15_Type.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", TemplateFileName & ".xls", "JC Seq", "A16")

        ' Load operators
        .Operation01_Operator.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", TemplateFileName & ".xls", "JC Seq", "b2")
        .Operation02_Operator.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", TemplateFileName & ".xls", "JC Seq", "b3")
        .Operation03_Operator.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", TemplateFileName & ".xls", "JC Seq", "b4")
        .Operation04_Operator.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", TemplateFileName & ".xls", "JC Seq", "b5")
        .Operation05_Operator.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", TemplateFileName & ".xls", "JC Seq", "b6")
        .Operation06_Operator.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", TemplateFileName & ".xls", "JC Seq", "b7")
        .Operation07_Operator.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", TemplateFileName & ".xls", "JC Seq", "b8")
        .Operation08_Operator.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", TemplateFileName & ".xls", "JC Seq", "b9")
        .Operation09_Operator.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", TemplateFileName & ".xls", "JC Seq", "b10")
        .Operation10_Operator.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", TemplateFileName & ".xls", "JC Seq", "b11")
        .Operation11_Operator.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", TemplateFileName & ".xls", "JC Seq", "b12")
        .Operation12_Operator.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", TemplateFileName & ".xls", "JC Seq", "b13")
        .Operation13_Operator.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", TemplateFileName & ".xls", "JC Seq", "b14")
        .Operation14_Operator.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", TemplateFileName & ".xls", "JC Seq", "b15")
        .Operation15_Operator.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", TemplateFileName & ".xls", "JC Seq", "b16")

        ' Load comments
        .Operation01_Comment.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", TemplateFileName & ".xls", "JC Seq", "c2")
        .Operation02_Comment.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", TemplateFileName & ".xls", "JC Seq", "c3")
        .Operation03_Comment.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", TemplateFileName & ".xls", "JC Seq", "c4")
        .Operation04_Comment.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", TemplateFileName & ".xls", "JC Seq", "c5")
        .Operation05_Comment.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", TemplateFileName & ".xls", "JC Seq", "c6")
        .Operation06_Comment.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", TemplateFileName & ".xls", "JC Seq", "c7")
        .Operation07_Comment.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", TemplateFileName & ".xls", "JC Seq", "c8")
        .Operation08_Comment.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", TemplateFileName & ".xls", "JC Seq", "c9")
        .Operation09_Comment.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", TemplateFileName & ".xls", "JC Seq", "c10")
        .Operation10_Comment.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", TemplateFileName & ".xls", "JC Seq", "c11")
        .Operation11_Comment.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", TemplateFileName & ".xls", "JC Seq", "c12")
        .Operation12_Comment.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", TemplateFileName & ".xls", "JC Seq", "c13")
        .Operation13_Comment.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", TemplateFileName & ".xls", "JC Seq", "c14")
        .Operation14_Comment.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", TemplateFileName & ".xls", "JC Seq", "c15")
        .Operation15_Comment.Value = GetValue(Main.Main_MasterPath.Value & "Job Templates", TemplateFileName & ".xls", "JC Seq", "c16")
    End With

    LoadJobCardTemplate = True
    Exit Function

Error_Handler:
    LoadJobCardTemplate = False
    MsgBox ("Error loading template: " & Err.Description)
End Function

' **Purpose**: Clear all operation fields
' **Original**: FJobCard.frm.RefreshFJobCard()
' **Parameters**: JobCardForm (Object)
' **Returns**: Nothing
' **Dependencies**: None
' **Side Effects**: Clears all operation Type, Operator, and Comment fields
Private Sub RefreshJobCardOperations(JobCardForm As Object)
    With JobCardForm
        ' Clear Operation Types
        .Operation01_Type.Value = ""
        .Operation02_Type.Value = ""
        .Operation03_Type.Value = ""
        .Operation04_Type.Value = ""
        .Operation05_Type.Value = ""
        .Operation06_Type.Value = ""
        .Operation07_Type.Value = ""
        .Operation08_Type.Value = ""
        .Operation09_Type.Value = ""
        .Operation10_Type.Value = ""
        .Operation11_Type.Value = ""
        .Operation12_Type.Value = ""
        .Operation13_Type.Value = ""
        .Operation14_Type.Value = ""
        .Operation15_Type.Value = ""

        ' Clear Operation Operators
        .Operation01_Operator.Value = ""
        .Operation02_Operator.Value = ""
        .Operation03_Operator.Value = ""
        .Operation04_Operator.Value = ""
        .Operation05_Operator.Value = ""
        .Operation06_Operator.Value = ""
        .Operation07_Operator.Value = ""
        .Operation08_Operator.Value = ""
        .Operation09_Operator.Value = ""
        .Operation10_Operator.Value = ""
        .Operation11_Operator.Value = ""
        .Operation12_Operator.Value = ""
        .Operation13_Operator.Value = ""
        .Operation14_Operator.Value = ""
        .Operation15_Operator.Value = ""

        ' Clear Operation Comments
        .Operation01_Comment.Value = ""
        .Operation02_Comment.Value = ""
        .Operation03_Comment.Value = ""
        .Operation04_Comment.Value = ""
        .Operation05_Comment.Value = ""
        .Operation06_Comment.Value = ""
        .Operation07_Comment.Value = ""
        .Operation08_Comment.Value = ""
        .Operation09_Comment.Value = ""
        .Operation10_Comment.Value = ""
        .Operation11_Comment.Value = ""
        .Operation12_Comment.Value = ""
        .Operation13_Comment.Value = ""
        .Operation14_Comment.Value = ""
        .Operation15_Comment.Value = ""
    End With
End Sub