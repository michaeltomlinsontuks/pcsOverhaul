Attribute VB_Name = "UserInterface"
' **Purpose**: UI management, form coordination, and application lifecycle management
' **CLAUDE.md Compliance**: No new forms created, maintains all existing form integrations
' **Consolidation**: Combines InterfaceManager.bas, MainInterfaceManager.bas, SearchFormManager.bas, ListManager.bas
Option Explicit

' ===================================================================
' CONSTANTS AND PRIVATE VARIABLES
' ===================================================================

' Legacy compatibility variables for file monitoring
Public NextCheck As Date

' **Purpose**: Flag to prevent recursive checkbox events causing automation errors
' **CLAUDE.md Compliance**: Added to fix checkbox circular event triggers
Private UpdatingCheckboxes As Boolean

' ===================================================================
' FILE LISTING FUNCTIONS (CLAUDE.md: Exact original compatibility)
' ===================================================================

' **Purpose**: List files in directory with status indicators (exact legacy compatibility)
' **Original**: Interface_VBA/a_ListFiles.bas.List_Files function
' **Parameters**:
'   - path (String): Directory name under Main_MasterPath (e.g. "Archive", "WIP")
'   - frm (Object): ListBox control to populate
' **Returns**: None (matches original function)
' **File Dependencies**: Uses Main.Main_MasterPath control value
' **Side Effects**: Populates list control with files and status indicators
' **Errors**: Shows message box if folder not found
' **CLAUDE.md Compliance**: Exact replacement for a_ListFiles.bas List_Files function
Public Function ListFiles(path As String, frm As Object)
    Dim Files(1 To 100000) As String
    Dim FullFilePath As String, MyName As String
    Dim GroupCount As Integer
    Dim x As String
    Dim i As Long
    Dim TestPath As String

    On Error GoTo Error_Handler

    ' Build the full path - ensure proper path separator
    Dim MasterPath As String
    MasterPath = Main.Main_MasterPath.Value
    If Right(MasterPath, 1) <> "\" Then MasterPath = MasterPath & "\"
    TestPath = MasterPath & path & "\"

    ' Check if directory exists first
    MyName = Dir(TestPath, vbDirectory)
    If MyName = "" Then
        MsgBox "Folder Not Found: " & TestPath, vbOKOnly, "Directory Missing"
        Exit Function
    End If

    ' Store list of file names
    Do Until MyName = ""
        If MyName = "." Or MyName = ".." Then GoTo NextFile

        GroupCount = GroupCount + 1
        Files(GroupCount) = MyName

NextFile:
        MyName = Dir
    Loop

    ' Add files to form with status indicators
    For i = 1 To GroupCount
        x = Files(i)

        ' Add status indicators based on file type and content
        If path = "WIP" Then
            ' Check if file has "Quote Accepted" status (support multiple formats)
            Dim WIPStatus As String
            WIPStatus = CStr(GetValueFromFile(MasterPath & path & "\", x, "ADMIN", "b88"))
            Dim NormalizedWIPStatus As String
            NormalizedWIPStatus = LCase(Trim(Replace(WIPStatus, " ", "")))

            If NormalizedWIPStatus = "quoteaccepted" Or NormalizedWIPStatus = "accepted" Then
                frm.AddItem Left(x, Len(x) - 4) & " *"
            Else
                frm.AddItem Left(x, Len(x) - 4)
            End If
        ElseIf path = "Quotes" Then
            ' Check if quote is new
            If GetValueFromFile(MasterPath & path & "\", x, "Admin", "b88") = "New Quote" Then
                frm.AddItem Left(x, Len(x) - 4) & " *"
            Else
                frm.AddItem Left(x, Len(x) - 4)
            End If
        Else
            ' No status indicators for Archive and Enquiries
            frm.AddItem Left(x, Len(x) - 4)
        End If
    Next i

    Exit Function

Error_Handler:
    SystemCore.LogError Err.Number, Err.Description, "ListFiles", "UserInterface"
End Function

' **Purpose**: Get value from closed workbook without opening it (safer version)
' **Original**: Interface_VBA/a_ListFiles.bas.GetValue function
' **Parameters**:
'   - path (String): Directory path with trailing backslash
'   - File (String): Excel filename
'   - sheet (String): Sheet name
'   - ref (String): Cell reference
' **Returns**: Variant - Cell value or "File Not Found"
' **Dependencies**: None (avoids ExecuteExcel4Macro to prevent save prompts)
' **Side Effects**: None
' **Errors**: Returns "File Not Found" if file access fails
' **CLAUDE.md Compliance**: Safer replacement that doesn't trigger save prompts
Private Function GetValueFromFile(path As String, File As String, sheet As String, ref As String) As Variant
    On Error GoTo Error_Handler

    ' Make sure the file exists
    If Right(path, 1) <> "\" Then path = path & "\"
    If Dir(path & File) = "" Then
        GetValueFromFile = "File Not Found"
        Exit Function
    End If

    ' Use DataOperations.GetValue instead of ExecuteExcel4Macro to avoid save prompts
    GetValueFromFile = DataOperations.GetValue(path & File, sheet, ref)

    If GetValueFromFile = 0 Then GetValueFromFile = ""

    Exit Function

Error_Handler:
    GetValueFromFile = "File Not Found"
End Function

' ===================================================================
' APPLICATION LIFECYCLE MANAGEMENT
' ===================================================================

' **Purpose**: Main system entry point - shows PCS main menu
' **Original**: Interface_VBA/a_Main.bas ShowMenu()
' **Parameters**: None
' **Returns**: None (Subroutine)
' **File Dependencies**: ActiveWorkbook.Path for setting Main.MasterPath
' **Form Usage**: Primary entry point for PCS system
' **CLAUDE.md Compliance**: Exact replacement for a_Main.bas ShowMenu functionality
Public Sub ShowMenu()
    On Error GoTo Error_Handler

    ' Set the master path from active workbook to 20081222 subdirectory (exact legacy behavior)
    Dim WorkbookPath As String
    WorkbookPath = ActiveWorkbook.Path

    ' Check if path already ends with 20081222 to avoid double path
    If Right(WorkbookPath, 8) = "20081222" Then
        Main.Main_MasterPath.Value = WorkbookPath & "\"
    Else
        Main.Main_MasterPath.Value = WorkbookPath & "\20081222\"
    End If

    ' Show the main form
    Main.Show

    ' Initialize the application after showing main form
    If Not InitializeApplication() Then
        SystemCore.LogError 0, "Application initialization failed after ShowMenu", "ShowMenu", "UserInterface"
    End If

    ' Initialize main interface AFTER path is set and application is initialized
    InitializeMainInterface Main

    Exit Sub

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "ShowMenu", "UserInterface"
End Sub

' **Purpose**: Initialize PCS application and validate system readiness
' **Parameters**: None
' **Returns**: Boolean - True if initialization successful, False if critical failure
' **Dependencies**: SystemCore.ValidateSystemRequirements, DataOperations.ValidateDirectoryStructure
' **Side Effects**: Validates all system components, logs system status, may display user messages
' **Errors**: Returns False on system validation failure, logs all issues
' **CLAUDE.md Compliance**: Preserves all existing system integrations
Public Function InitializeApplication() As Boolean
    Dim SystemConfig As SystemCore.SystemConfig
    Dim InitErrors As String

    On Error GoTo Error_Handler

    ' Log application startup
    SystemCore.LogError 0, "PCS Application initialization started", "InitializeApplication", "UserInterface"

    ' Get and validate system configuration
    SystemConfig = SystemCore.GetSystemConfig()

    If SystemConfig.RootPath = "" Then
        InitErrors = InitErrors & "Unable to determine system root path." & vbCrLf
    End If

    If SystemConfig.CurrentUser = "" Then
        InitErrors = InitErrors & "Unable to determine current user." & vbCrLf
    End If

    ' Validate system requirements
    If Not SystemCore.ValidateSystemRequirements() Then
        InitErrors = InitErrors & "System requirements validation failed." & vbCrLf
    End If

    ' Validate directory structure
    If Not DataOperations.ValidateDirectoryStructure() Then
        InitErrors = InitErrors & "Directory structure validation failed." & vbCrLf
    End If

    ' Create missing directories if needed
    If Not DataOperations.CreateDirectoryStructure() Then
        InitErrors = InitErrors & "Unable to create required directories." & vbCrLf
    End If

    ' Validate all business controllers
    If Not ValidateBusinessControllers() Then
        InitErrors = InitErrors & "Business controller validation failed." & vbCrLf
    End If

    ' Check search database integrity
    If Not ValidateSearchSystem() Then
        InitErrors = InitErrors & "Search system validation failed." & vbCrLf
    End If

    ' Display errors if any
    If InitErrors <> "" Then
        MsgBox "PCS Application initialization completed with warnings:" & vbCrLf & vbCrLf & InitErrors, vbExclamation, "Initialization Warnings"
        SystemCore.LogError 0, "Initialization warnings: " & InitErrors, "InitializeApplication", "UserInterface"
        InitializeApplication = False
        Exit Function
    End If

    ' Log successful initialization
    SystemCore.LogError 0, "PCS Application initialization completed successfully", "InitializeApplication", "UserInterface"

    InitializeApplication = True
    Exit Function

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "InitializeApplication", "UserInterface"
    InitializeApplication = False
End Function

' **Purpose**: Safely shutdown PCS application
' **Parameters**: None
' **Returns**: Boolean - True if shutdown successful, False if issues occurred
' **Dependencies**: None
' **Side Effects**: Closes all forms, saves pending data, logs shutdown
' **Errors**: Returns False if shutdown process encounters errors
Public Function ShutdownApplication() As Boolean
    On Error GoTo Error_Handler

    ' Log application shutdown
    SystemCore.LogError 0, "PCS Application shutdown initiated", "ShutdownApplication", "UserInterface"

    ' Close all user forms
    If Not CloseAllForms() Then
        SystemCore.LogError 0, "Warning: Some forms could not be closed properly", "ShutdownApplication", "UserInterface"
    End If

    ' Perform final data validation
    If Not PerformFinalDataValidation() Then
        SystemCore.LogError 0, "Warning: Final data validation found issues", "ShutdownApplication", "UserInterface"
    End If

    ' Clear any temporary data
    ClearTemporaryData

    ' Log successful shutdown
    SystemCore.LogError 0, "PCS Application shutdown completed successfully", "ShutdownApplication", "UserInterface"

    ShutdownApplication = True
    Exit Function

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "ShutdownApplication", "UserInterface"
    ShutdownApplication = False
End Function

' **Purpose**: Check and update Main form file count displays (exact legacy compatibility)
' **Parameters**: None
' **Returns**: Boolean - True if update successful, False if failed
' **Dependencies**: Main form, Check_Files function, DataOperations.GetRootPath
' **Side Effects**: Updates Main form Notice labels, schedules next update check
' **Errors**: Returns False if update check fails
' **CLAUDE.md Compliance**: Exact replacement for legacy Check_Updates.bas functionality
Public Function CheckForUpdates() As Boolean
    On Error GoTo Error_Handler

    ' Exit if Main form not visible or not ready for next check
    If Main.Visible = False Or NextCheck > Now() Then
        If NextCheck = "12:00:00 AM" Then GoTo ContinueCheck
        StopCheck
        CheckForUpdates = True
        Exit Function
    End If

ContinueCheck:
    Dim RootPath As String
    RootPath = DataOperations.GetRootPath()

    ' Update Enquiries count display
    Dim EnquiriesCount As String
    EnquiriesCount = "Enquiries : " & Check_Files(RootPath & "\Enquiries\")
    If EnquiriesCount <> Main.Notice_Enquiries.Caption Then
        Main.Notice_Enquiries.Caption = EnquiriesCount & "*"
    End If

    ' Update Quotes count display
    Dim QuotesCount As String
    QuotesCount = "Quotes : " & Check_Files(RootPath & "\Quotes\")
    If QuotesCount <> Main.Notice_Quotes.Caption Then
        Main.Notice_Quotes.Caption = QuotesCount & "*"
    End If

    ' Update WIP count display
    Dim WIPCount As String
    WIPCount = "WIP : " & Check_Files(RootPath & "\WIP\")
    If WIPCount <> Main.Notice_WIP.Caption Then
        Main.Notice_WIP.Caption = WIPCount & "*"
    End If

    ' Schedule next update check (every 5 minutes)
    NextCheck = Now + TimeValue("00:05:00")
    Application.OnTime NextCheck, "CheckForUpdates", NextCheck + TimeValue("00:01:00")

    CheckForUpdates = True
    Exit Function

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "CheckForUpdates", "UserInterface"
    CheckForUpdates = False
End Function

' **Purpose**: Refresh Main form UI with current data (exact legacy compatibility)
' **Parameters**: None
' **Returns**: Boolean - True if refresh successful, False if failed
' **Dependencies**: Main form, List_Files function, CheckForUpdates
' **Side Effects**: Clears and repopulates Main.lst based on checkbox selections, clears form controls
' **Errors**: Returns False if refresh operation fails
' **CLAUDE.md Compliance**: Exact replacement for legacy RefreshMain.bas functionality
Public Function RefreshMainInterface() As Boolean
    Dim ctl As Control

    On Error GoTo Error_Handler

    ' Clear the main list box (exact legacy behavior)
    Main.lst.Clear

    ' Populate list based on checkbox selections (exact legacy logic)
    If Main.Enquiries.Value = True Then
        Call List_Files("Enquiries", Main.lst)
    End If

    If Main.Quotes.Value = True Then
        Call List_Files("quotes", Main.lst)
    End If

    If Main.WIP.Value = True Then
        Call List_Files("WIP", Main.lst)
    End If

    If Main.Archive.Value = True Then
        Call List_Files("Archive", Main.lst)
    End If


    ' Clear all form controls
    For Each ctl In Main.Controls
        On Error Resume Next
        Select Case TypeName(ctl)
            Case "TextBox"
                ctl.Value = ""
            Case "ComboBox"
                ctl.ListIndex = -1
            Case "Label"
                If Left(ctl.Name, 6) <> "Notice" And ctl.Name <> "Label1" Then
                    ctl.Caption = ""
                End If
        End Select
        On Error GoTo Error_Handler
    Next ctl

    ' Update file counts
    CheckForUpdates

    RefreshMainInterface = True
    Exit Function

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "RefreshMainInterface", "UserInterface"
    RefreshMainInterface = False
End Function

' **Purpose**: Handle main form list selection change
' **Parameters**: None
' **Returns**: None (Subroutine)
' **Dependencies**: Main form, DataOperations.GetValue
' **Side Effects**: Updates form labels with selected file information
' **Errors**: May fail silently if file access problems
' **CLAUDE.md Compliance**: Maintains exact legacy list selection behavior
Public Sub HandleMainListChange()
    Dim SelectedFile As String
    Dim FilePath As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim ctl As Object
    Dim i As Integer

    On Error GoTo Error_Handler

    If Main.lst.ListIndex < 0 Then Exit Sub

    SelectedFile = Main.lst.Value
    If InStr(SelectedFile, "*") > 0 Then
        SelectedFile = Left(SelectedFile, Len(SelectedFile) - 2)
    End If

    ' Clear all textbox values first (exact legacy behavior)
    For Each ctl In Main.Controls
        If TypeName(ctl) = "Textbox" Then ctl.Value = ""
    Next ctl

    ' Determine file location based on checkbox selections
    Dim RootPath As String
    RootPath = DataOperations.GetRootPath()
    If Right(RootPath, 1) <> "\" Then RootPath = RootPath & "\"

    If Main.Enquiries.Value = True Then
        FilePath = RootPath & "Enquiries\" & SelectedFile & ".xls"
    ElseIf Main.Quotes.Value = True Then
        FilePath = RootPath & "Quotes\" & SelectedFile & ".xls"
    ElseIf Main.WIP.Value = True Then
        FilePath = RootPath & "WIP\" & SelectedFile & ".xls"
    ElseIf Main.Archive.Value = True Then
        FilePath = RootPath & "Archive\" & SelectedFile & ".xls"
    End If

    ' Load file data if it exists
    If DataOperations.FileExists(FilePath) Then
        Set wb = DataOperations.SafeOpenWorkbook(FilePath, True)
        If wb Is Nothing Then Exit Sub

        Set ws = wb.Worksheets("Admin")

        ' Load all form controls from Admin sheet (exact legacy algorithm)
        For Each ctl In Main.Controls
            i = -1
            Do
                i = i + 1
                If UCase(ws.Range("A1").Offset(i, 0).Value) = UCase(ctl.Name) Then
                    ' Handle price fields with currency formatting
                    If InStr(1, ctl.Name, "Price", vbTextCompare) <> 0 Then
                        If UCase(TypeName(ctl)) = "LABEL" Then
                            ctl.Caption = SystemCore.Insert_Characters(ctl.Name) & " : " & Format(ws.Range("A1").Offset(i, 1).Value, "R #,##0.00")
                        ElseIf UCase(TypeName(ctl)) = "COMBOBOX" Then
                            ctl.Value = Format(ws.Range("A1").Offset(i, 1).Value, "R #,##0.00")
                        ElseIf UCase(TypeName(ctl)) = "TEXTBOX" Then
                            ctl.Value = Format(ws.Range("A1").Offset(i, 1).Value, "R #,##0.00")
                        End If
                        GoTo FormLoadNext
                    End If

                    ' Handle regular fields
                    If UCase(TypeName(ctl)) = "LABEL" Then
                        ctl.Caption = SystemCore.Insert_Characters(ctl.Name) & " : " & ws.Range("A1").Offset(i, 1).Value
                    ElseIf UCase(TypeName(ctl)) = "COMBOBOX" Then
                        ctl.Value = ws.Range("A1").Offset(i, 1).Value
                    ElseIf UCase(TypeName(ctl)) = "TEXTBOX" Then
                        ' Handle date fields with date formatting
                        If InStr(1, ctl.Name, "Date", vbTextCompare) > 0 Then
                            ctl.Value = Format(ws.Range("A1").Offset(i, 1).Value, "dd mmm yyyy")
                        Else
                            ctl.Value = ws.Range("A1").Offset(i, 1).Value
                        End If
                    End If
                    GoTo FormLoadNext
                End If
                If ws.Range("A1").Offset(i, 0).Value = "" Then Exit Do
            Loop
FormLoadNext:
        Next ctl

        DataOperations.SafeCloseWorkbook wb, False
    End If

    Exit Sub

Error_Handler:
    If Not wb Is Nothing Then DataOperations.SafeCloseWorkbook wb, False
    SystemCore.LogError Err.Number, Err.Description, "HandleMainListChange", "UserInterface"
End Sub

' **Purpose**: Open selected file from main list
' **Parameters**: None
' **Returns**: None (Subroutine)
' **Dependencies**: Main form, DataOperations.SafeOpenWorkbook
' **Side Effects**: Opens Excel file for editing
' **Errors**: May display error if file cannot be opened
Public Sub OpenSelectedFile()
    Dim SelectedFile As String
    Dim FilePath As String

    On Error GoTo Error_Handler

    If Main.lst.ListIndex < 0 Then
        SystemCore.ShowWarning "Please select a file to open.", "No Selection"
        Exit Sub
    End If

    SelectedFile = Main.lst.Value
    If InStr(SelectedFile, "*") > 0 Then
        SelectedFile = Left(SelectedFile, Len(SelectedFile) - 2)
    End If

    ' Determine file location based on checkbox selections
    Dim RootPath As String
    RootPath = DataOperations.GetRootPath()
    If Right(RootPath, 1) <> "\" Then RootPath = RootPath & "\"

    If Main.Enquiries.Value = True Then
        FilePath = RootPath & "Enquiries\" & SelectedFile & ".xls"
    ElseIf Main.Quotes.Value = True Then
        FilePath = RootPath & "Quotes\" & SelectedFile & ".xls"
    ElseIf Main.WIP.Value = True Then
        FilePath = RootPath & "WIP\" & SelectedFile & ".xls"
    ElseIf Main.Archive.Value = True Then
        FilePath = RootPath & "Archive\" & SelectedFile & ".xls"
    End If

    If DataOperations.FileExists(FilePath) Then
        Dim wb As Workbook
        Set wb = DataOperations.SafeOpenWorkbook(FilePath)
        If wb Is Nothing Then
            SystemCore.ShowError "Unable to open file: " & FilePath, "File Open Error"
        End If
    Else
        SystemCore.ShowError "File not found: " & FilePath, "File Not Found"
    End If

    Exit Sub

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "OpenSelectedFile", "UserInterface"
End Sub

' ===================================================================
' SEARCH FORM MANAGEMENT
' ===================================================================

' **Purpose**: Initialize search form with default values
' **Parameters**:
'   - SearchForm (Object): Search form to initialize
' **Returns**: None (Subroutine)
' **Dependencies**: None
' **Side Effects**: Clears form controls, sets default values
' **Errors**: None
Public Sub InitializeSearchForm(SearchForm As Object)
    On Error Resume Next
    SearchForm.SearchTerm.Value = ""
    SearchForm.RecordTypeFilter.ListIndex = 0 ' All types
    SearchForm.ResultsList.Clear
    SearchForm.SearchTerm.SetFocus
    On Error GoTo 0
End Sub

' **Purpose**: Perform search based on form criteria
' **Parameters**:
'   - SearchForm (Object): Search form containing criteria
' **Returns**: Boolean - True if search successful, False if failed
' **Dependencies**: BusinessLogic.SearchRecords, PopulateSearchResults
' **Side Effects**: Updates search results list in form
' **Errors**: Returns False if search fails
Public Function PerformSearch(SearchForm As Object) As Boolean
    Dim SearchTerm As String
    Dim RecordTypeFilter As Long
    Dim SearchResults As Variant

    On Error GoTo Error_Handler

    SearchTerm = Trim(SearchForm.SearchTerm.Value)
    If SearchTerm = "" Then
        SystemCore.ShowWarning "Please enter a search term.", "Search Term Required"
        SearchForm.SearchTerm.SetFocus
        PerformSearch = False
        Exit Function
    End If

    RecordTypeFilter = GetRecordTypeFromForm(SearchForm)

    SearchResults = BusinessLogic.SearchRecords(SearchTerm, RecordTypeFilter)

    If IsArray(SearchResults) And UBound(SearchResults) >= 0 Then
        PopulateSearchResults SearchForm, SearchResults
        SystemCore.ShowInformation "Found " & (UBound(SearchResults) + 1) & " result(s).", "Search Complete"
        PerformSearch = True
    Else
        SearchForm.ResultsList.Clear
        SystemCore.ShowInformation "No results found for: " & SearchTerm, "No Results"
        PerformSearch = True ' Not an error, just no results
    End If

    Exit Function

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "PerformSearch", "UserInterface"
    PerformSearch = False
End Function

' **Purpose**: Export search results to Excel file
' **Parameters**:
'   - SearchForm (Object): Search form containing results
' **Returns**: Boolean - True if export successful, False if failed
' **Dependencies**: DataOperations.CreateNewWorkbook
' **Side Effects**: Creates new Excel file with search results
' **Errors**: Returns False if export fails
Public Function ExportSearchResults(SearchForm As Object) As Boolean
    Dim ExportWB As Workbook
    Dim ExportWS As Worksheet
    Dim i As Integer
    Dim ExportPath As String

    On Error GoTo Error_Handler

    If SearchForm.ResultsList.ListCount = 0 Then
        SystemCore.ShowWarning "No search results to export.", "No Results"
        ExportSearchResults = False
        Exit Function
    End If

    Set ExportWB = DataOperations.CreateNewWorkbook()
    If ExportWB Is Nothing Then
        ExportSearchResults = False
        Exit Function
    End If

    Set ExportWS = ExportWB.Worksheets(1)
    ExportWS.Name = "Search_Results"

    ' Create headers
    With ExportWS
        .Cells(1, 1).Value = "Search Results Export"
        .Cells(1, 1).Font.Bold = True
        .Cells(2, 1).Value = "Generated: " & DataOperations.FormatDateTime(Now)

        .Cells(4, 1).Value = "Record Type"
        .Cells(4, 2).Value = "Record Number"
        .Cells(4, 3).Value = "Customer"
        .Cells(4, 4).Value = "Description"
        .Cells(4, 5).Value = "Date Created"
        .Cells(4, 6).Value = "File Path"
        .Range("A4:F4").Font.Bold = True
    End With

    ' Export results (simplified - would need actual result data structure)
    For i = 0 To SearchForm.ResultsList.ListCount - 1
        ExportWS.Cells(5 + i, 1).Value = SearchForm.ResultsList.List(i)
    Next i

    ExportWS.Columns.AutoFit

    Dim RootPath As String
    RootPath = DataOperations.GetRootPath()
    If Right(RootPath, 1) <> "\" Then RootPath = RootPath & "\"
    ExportPath = RootPath & "SearchResults_" & Format(Now, "yyyymmdd_hhmmss") & ".xls"
    ExportWB.SaveAs ExportPath
    ExportWB.Close

    SystemCore.ShowInformation "Search results exported to:" & vbCrLf & ExportPath, "Export Complete"
    ExportSearchResults = True
    Exit Function

Error_Handler:
    If Not ExportWB Is Nothing Then ExportWB.Close SaveChanges:=False
    SystemCore.HandleStandardErrors Err.Number, "ExportSearchResults", "UserInterface"
    ExportSearchResults = False
End Function

' ===================================================================
' GENERIC LIST MANAGEMENT
' ===================================================================

' **Purpose**: Populate list with items
' **Parameters**:
'   - ListForm (Object): Form containing list control
'   - ListItems (Variant): Array of items to display
'   - Title (String): Form title to set
' **Returns**: None (Subroutine)
' **Dependencies**: None
' **Side Effects**: Updates form title and populates list
' **Errors**: None
Public Sub PopulateList(ListForm As Object, ListItems As Variant, Title As String)
    Dim i As Integer

    On Error Resume Next
    ListForm.Caption = Title
    ListForm.lst.Clear

    If IsArray(ListItems) Then
        For i = 0 To UBound(ListItems)
            ListForm.lst.AddItem ListItems(i)
        Next i
    End If
    On Error GoTo 0
End Sub

' **Purpose**: Validate list selection
' **Parameters**:
'   - ListForm (Object): Form containing list control
' **Returns**: Boolean - True if item selected, False if no selection
' **Dependencies**: SystemCore.ShowWarning
' **Side Effects**: Shows warning if no selection
' **Errors**: Returns False if no selection
Public Function ValidateListSelection(ListForm As Object) As Boolean
    ValidateListSelection = (ListForm.lst.ListIndex >= 0)
    If Not ValidateListSelection Then
        SystemCore.ShowWarning "Please select an item from the list.", "Selection Required"
    End If
End Function

' **Purpose**: Get selected value from list
' **Parameters**:
'   - ListForm (Object): Form containing list control
' **Returns**: String - Selected value, empty if no selection
' **Dependencies**: None
' **Side Effects**: None
' **Errors**: Returns empty string if no selection
Public Function GetSelectedValue(ListForm As Object) As String
    If ListForm.lst.ListIndex >= 0 Then
        GetSelectedValue = ListForm.lst.List(ListForm.lst.ListIndex)
    Else
        GetSelectedValue = ""
    End If
End Function

' ===================================================================
' FORM LIFECYCLE MANAGEMENT
' ===================================================================

' **Purpose**: Show form with standard initialization
' **Parameters**:
'   - FormName (String): Name of form to show
'   - InitializeData (Boolean, Optional): Whether to initialize form data
' **Returns**: Boolean - True if form shown successfully, False if failed
' **Dependencies**: Various form initialization functions
' **Side Effects**: Shows specified form, may initialize data
' **Errors**: Returns False if form cannot be shown
Public Function ShowForm(FormName As String, Optional InitializeData As Boolean = True) As Boolean
    On Error GoTo Error_Handler

    Select Case UCase(FormName)
        Case "ENQUIRY", "FENQUIRY"
            FEnquiry.Show
            If InitializeData Then WorkflowManagement.InitializeEnquiryForm(FEnquiry)

        Case "QUOTE", "FQUOTE"
            FQuote.Show
            If InitializeData Then WorkflowManagement.InitializeQuoteForm(FQuote)

        Case "JOBCARD", "FJOBCARD"
            FJobCard.Show
            If InitializeData Then WorkflowManagement.LoadJobTemplates(FJobCard)

        Case "ACCEPTQUOTE", "FACCEPTQUOTE"
            FAcceptQuote.Show

        Case "SEARCH", "FRMSEARCH"
            ' **Purpose**: Open Search.xls directly as per original Interface_VBA/Main.frm Search_Click() method
            ' **CLAUDE.md Compliance**: Uses existing Search.xls file, no new forms created
            OpenSearchDatabase
            ' Remove the InitializeData call since we're not using a form anymore

        Case "WIP", "FWIP"
            fwip.Show

        Case "JOBGENERATION", "FJG"
            FJG.Show
            If InitializeData Then WorkflowManagement.InitializeJobGenerationForm(FJG)

        Case Else
            SystemCore.ShowWarning "Unknown form: " & FormName, "Form Not Found"
            ShowForm = False
            Exit Function
    End Select

    ShowForm = True
    Exit Function

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "ShowForm", "UserInterface"
    ShowForm = False
End Function

' **Purpose**: Close all open user forms
' **Parameters**: None
' **Returns**: Boolean - True if all forms closed successfully, False if issues
' **Dependencies**: None
' **Side Effects**: Closes all user forms
' **Errors**: Returns False if any forms cannot be closed
Private Function CloseAllForms() As Boolean
    On Error Resume Next

    ' Close all known forms
    FEnquiry.Hide
    FQuote.Hide
    FJobCard.Hide
    FAcceptQuote.Hide
    frmSearch.Hide
    fwip.Hide
    FJG.Hide
    FList.Hide

    ' Return True if no errors occurred
    CloseAllForms = (Err.Number = 0)
    On Error GoTo 0
End Function

' ===================================================================
' PRIVATE HELPER FUNCTIONS
' ===================================================================

' **Purpose**: Validate business controller functionality
' **Parameters**: None
' **Returns**: Boolean - True if all business controllers functional
' **Dependencies**: BusinessLogic.ValidateSearchCompatibility
' **Side Effects**: None
' **Errors**: Returns False if any controller validation fails
Private Function ValidateBusinessControllers() As Boolean
    On Error GoTo Error_Handler

    ' Test core business functions
    If Not BusinessLogic.ValidateSearchCompatibility() Then
        ValidateBusinessControllers = False
        Exit Function
    End If

    ' Could add more controller validations here
    ValidateBusinessControllers = True
    Exit Function

Error_Handler:
    ValidateBusinessControllers = False
End Function

' **Purpose**: Validate search system functionality
' **Parameters**: None
' **Returns**: Boolean - True if search system functional
' **Dependencies**: BusinessLogic.ValidateSearchCompatibility
' **Side Effects**: None
' **Errors**: Returns False if search validation fails
Private Function ValidateSearchSystem() As Boolean
    ValidateSearchSystem = BusinessLogic.ValidateSearchCompatibility()
End Function

' **Purpose**: Perform final data validation before shutdown
' **Parameters**: None
' **Returns**: Boolean - True if all data valid
' **Dependencies**: None
' **Side Effects**: None
' **Errors**: Returns False if data validation issues found
Private Function PerformFinalDataValidation() As Boolean
    ' Placeholder for final validation logic
    PerformFinalDataValidation = True
End Function

' **Purpose**: Clear temporary data and cache
' **Parameters**: None
' **Returns**: None (Subroutine)
' **Dependencies**: None
' **Side Effects**: Clears temporary data
' **Errors**: None
Private Sub ClearTemporaryData()
    ' Clear any temporary data structures
    On Error Resume Next
    ' Implementation would depend on specific temporary data used
    On Error GoTo 0
End Sub

' **Purpose**: Stop scheduled update checks
' **Parameters**: None
' **Returns**: None (Subroutine)
' **Dependencies**: Application.OnTime
' **Side Effects**: Cancels scheduled update checks
' **Errors**: None
Public Sub StopCheck()
    On Error Resume Next
    Application.OnTime NextCheck, "CheckForUpdates", , False
    On Error GoTo 0
End Sub

' **Purpose**: Count files in directory (legacy compatibility)
' **Parameters**:
'   - DirectoryPath (String): Directory to count files in
' **Returns**: Integer - Number of files found
' **Dependencies**: DataOperations.CountFilesInFolder
' **Side Effects**: None
' **Errors**: Returns 0 if directory not accessible
Private Function Check_Files(DirectoryPath As String) As Integer
    Check_Files = DataOperations.CountFilesInFolder(DirectoryPath, "*.xls")
End Function

' **Purpose**: Populate files in list control (legacy compatibility)
' **Parameters**:
'   - DirectoryName (String): Name of directory under root path
'   - ListControl (Object): List control to populate
' **Returns**: None (Subroutine)
' **Dependencies**: DataOperations.GetFileListWithStatus
' **Side Effects**: Populates list control with files and status
' **Errors**: None
Private Sub List_Files(DirectoryName As String, ListControl As Object)
    DataOperations.GetFileListWithStatus DirectoryName, ListControl
End Sub

' **Purpose**: Get record type filter from search form
' **Parameters**:
'   - SearchForm (Object): Search form containing filter controls
' **Returns**: Long - Record type filter value
' **Dependencies**: None
' **Side Effects**: None
' **Errors**: Returns 0 (all types) if error
Private Function GetRecordTypeFromForm(SearchForm As Object) As Long
    On Error Resume Next
    Select Case SearchForm.RecordTypeFilter.ListIndex
        Case 0
            GetRecordTypeFromForm = 0 ' All types
        Case 1
            GetRecordTypeFromForm = SystemCore.rtEnquiry
        Case 2
            GetRecordTypeFromForm = SystemCore.rtQuote
        Case 3
            GetRecordTypeFromForm = SystemCore.rtJob
        Case 4
            GetRecordTypeFromForm = SystemCore.rtContract
        Case Else
            GetRecordTypeFromForm = 0
    End Select
    On Error GoTo 0
End Function

' **Purpose**: Populate search results in form
' **Parameters**:
'   - SearchForm (Object): Search form containing results list
'   - SearchResults (Variant): Array of search results
' **Returns**: None (Subroutine)
' **Dependencies**: None
' **Side Effects**: Populates results list in search form
' **Errors**: None
Private Sub PopulateSearchResults(SearchForm As Object, SearchResults As Variant)
    Dim i As Integer

    On Error Resume Next
    SearchForm.ResultsList.Clear

    If IsArray(SearchResults) Then
        For i = 0 To UBound(SearchResults, 1)
            ' Format result display (customize based on actual result structure)
            Dim ResultText As String
            If UBound(SearchResults, 2) >= 3 Then
                ResultText = SearchResults(i, 1) & " - " & SearchResults(i, 2) & " - " & SearchResults(i, 3)
            Else
                ResultText = CStr(SearchResults(i, 0))
            End If
            SearchForm.ResultsList.AddItem ResultText
        Next i
    End If
    On Error GoTo 0
End Sub

' ===================================================================
' MAIN INTERFACE MANAGEMENT - METHODS CALLED FROM Main.frm
' ===================================================================

' **Purpose**: Initialize Main form interface with default settings
' **Original**: MainInterfaceManager.bas (various initialization functions)
' **Parameters**:
'   - MainForm (Object): Main form to initialize
' **Returns**: None (Subroutine)
' **File Dependencies**: Main form controls
' **Form Usage**: Called from Main.UserForm_Initialize
Public Sub InitializeMainInterface(MainForm As Object)
    On Error GoTo Error_Handler

    ' Ensure MasterPath is set before any checkbox operations
    If MainForm.Main_MasterPath.Value = "" Then
        Dim WorkbookPath As String
        WorkbookPath = ActiveWorkbook.Path
        If Right(WorkbookPath, 8) = "20081222" Then
            MainForm.Main_MasterPath.Value = WorkbookPath & "\"
        Else
            MainForm.Main_MasterPath.Value = WorkbookPath & "\20081222\"
        End If
    End If

    ' Set WIP as default selection (matches original UserForm_Activate)
    ' This will trigger ShowWIPFiles automatically due to checkbox click event
    MainForm.WIP.Value = True

    DisplayStatusMessage "PCS Interface Ready", "Info"

    ' Start update check timer
    CheckForUpdates

    Exit Sub

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "InitializeMainInterface", "UserInterface"
End Sub

' **Purpose**: Open enquiry form to add new enquiry
' **Original**: Interface_VBA/Main.frm.Add_Enquiry_Click business logic
' **Parameters**:
'   - MainForm (Object): Main form reference
' **Returns**: None (Subroutine)
' **File Dependencies**: FEnquiry form
' **Form Usage**: Called from Main.Add_Enquiry_Click
Public Sub AddEnquiry(MainForm As Object)
    On Error GoTo Error_Handler

    If ShowForm("FEnquiry", True) Then
        DisplayStatusMessage "New enquiry form opened", "Info"
    End If

    Exit Sub

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "AddEnquiry", "UserInterface"
End Sub

' **Purpose**: Show archive files in main list
' **Original**: Interface_VBA/Main.frm.Archive_Click business logic
' **Parameters**:
'   - MainForm (Object): Main form reference
' **Returns**: None (Subroutine)
' **File Dependencies**: Archive directory, Main.lst control
' **Form Usage**: Called from Main.Archive_Click
Public Sub ShowArchiveFiles(MainForm As Object)
    On Error GoTo Error_Handler

    ' Check if Archive checkbox is actually selected (matches original logic)
    If MainForm.Archive.Value = True Then
        ' Prevent recursive checkbox events
        If UpdatingCheckboxes Then Exit Sub
        UpdatingCheckboxes = True

        ' Clear the list first (exact original behavior)
        MainForm.lst.Clear

        ' Populate list with archive files (matches original List_Files call)
        ListFiles "Archive", MainForm.lst

        ' Clear other checkboxes (exact original behavior)
        MainForm.Enquiries.Value = False
        MainForm.Quotes.Value = False
        MainForm.WIP.Value = False
        MainForm.JobsInWIP.Value = False

        ' Reset flag
        UpdatingCheckboxes = False

        DisplayStatusMessage "Archive files displayed", "Info"
    End If

    Exit Sub

Error_Handler:
    UpdatingCheckboxes = False
    SystemCore.HandleStandardErrors Err.Number, "ShowArchiveFiles", "UserInterface"
End Sub

' **Purpose**: Show enquiry files in main list
' **Original**: Interface_VBA/Main.frm.Enquiries_Click business logic
' **Parameters**:
'   - MainForm (Object): Main form reference
' **Returns**: None (Subroutine)
' **File Dependencies**: Enquiries directory, Main.lst control
' **Form Usage**: Called from Main.Enquiries_Click
Public Sub ShowEnquiries(MainForm As Object)
    On Error GoTo Error_Handler

    ' Check if Enquiries checkbox is actually selected (matches original logic)
    If MainForm.Enquiries.Value = True Then
        ' Prevent recursive checkbox events
        If UpdatingCheckboxes Then Exit Sub
        UpdatingCheckboxes = True

        ' Clear the list first (exact original behavior)
        MainForm.lst.Clear

        ' Populate list with enquiry files (matches original List_Files call)
        ListFiles "Enquiries", MainForm.lst

        ' Clear other checkboxes (exact original behavior)
        MainForm.Quotes.Value = False
        MainForm.WIP.Value = False
        MainForm.Archive.Value = False
        MainForm.JobsInWIP.Value = False

        ' Reset flag
        UpdatingCheckboxes = False

        DisplayStatusMessage "Enquiry files displayed", "Info"
    End If

    Exit Sub

Error_Handler:
    UpdatingCheckboxes = False
    SystemCore.HandleStandardErrors Err.Number, "ShowEnquiries", "UserInterface"
End Sub

' **Purpose**: Show quote files in main list
' **Original**: Interface_VBA/Main.frm.Quotes_Click business logic
' **Parameters**:
'   - MainForm (Object): Main form reference
' **Returns**: None (Subroutine)
' **File Dependencies**: Quotes directory, Main.lst control
' **Form Usage**: Called from Main.Quotes_Click
Public Sub ShowQuotes(MainForm As Object)
    On Error GoTo Error_Handler

    ' Check if Quotes checkbox is actually selected (matches original logic)
    If MainForm.Quotes.Value = True Then
        ' Prevent recursive checkbox events
        If UpdatingCheckboxes Then Exit Sub
        UpdatingCheckboxes = True

        ' Clear the list first (exact original behavior)
        MainForm.lst.Clear

        ' Populate list with quote files (matches original List_Files call)
        ListFiles "Quotes", MainForm.lst

        ' Clear other checkboxes (exact original behavior)
        MainForm.Enquiries.Value = False
        MainForm.WIP.Value = False
        MainForm.Archive.Value = False
        MainForm.JobsInWIP.Value = False

        ' Reset flag
        UpdatingCheckboxes = False

        DisplayStatusMessage "Quote files displayed", "Info"
    End If

    Exit Sub

Error_Handler:
    UpdatingCheckboxes = False
    SystemCore.HandleStandardErrors Err.Number, "ShowQuotes", "UserInterface"
End Sub

' **Purpose**: Show WIP files in main list
' **Original**: Interface_VBA/Main.frm.WIP_Click business logic
' **Parameters**:
'   - MainForm (Object): Main form reference
' **Returns**: None (Subroutine)
' **File Dependencies**: WIP directory, Main.lst control
' **Form Usage**: Called from Main.WIP_Click
Public Sub ShowWIPFiles(MainForm As Object)
    On Error GoTo Error_Handler

    ' Check if WIP checkbox is actually selected (matches original logic)
    If MainForm.WIP.Value = True Then
        ' Prevent recursive checkbox events
        If UpdatingCheckboxes Then Exit Sub
        UpdatingCheckboxes = True

        ' Clear the list first (exact original behavior)
        MainForm.lst.Clear

        ' Populate list with WIP files (matches original List_Files call)
        ListFiles "WIP", MainForm.lst

        ' Clear other checkboxes (exact original behavior)
        MainForm.Enquiries.Value = False
        MainForm.Quotes.Value = False
        MainForm.Archive.Value = False
        MainForm.JobsInWIP.Value = False

        ' Reset flag
        UpdatingCheckboxes = False

        DisplayStatusMessage "WIP files displayed", "Info"
    End If

    Exit Sub

Error_Handler:
    UpdatingCheckboxes = False
    SystemCore.HandleStandardErrors Err.Number, "ShowWIPFiles", "UserInterface"
End Sub

' **Purpose**: Show jobs from WIP database in main list
' **Original**: Interface_VBA/Main.frm.JobsInWIP_Click business logic
' **Parameters**:
'   - MainForm (Object): Main form reference
' **Returns**: None (Subroutine)
' **File Dependencies**: WIP.xls database, Main.lst control
' **Form Usage**: Called from Main.JobsInWIP_Click
Public Sub ShowJobsInWIP(MainForm As Object)
    On Error GoTo Error_Handler

    ' Check if JobsInWIP checkbox is actually selected (matches original logic)
    If MainForm.JobsInWIP.Value = True Then
        ' Prevent recursive checkbox events
        If UpdatingCheckboxes Then Exit Sub
        UpdatingCheckboxes = True

        ' Clear the list first (exact original behavior)
        MainForm.lst.Clear

        ' Clear other checkboxes (mutual exclusivity - original behavior)
        MainForm.Enquiries.Value = False
        MainForm.WIP.Value = False
        MainForm.Archive.Value = False
        MainForm.Quotes.Value = False
        MainForm.Thirties.Value = False

        ' Get WIP database jobs using V2 infrastructure
        Dim WIPJobs As Variant
        WIPJobs = DataOperations.GetWIPDatabaseJobs()

        ' Populate main list with job numbers
        If IsArray(WIPJobs) Then
            Dim i As Integer
            For i = 0 To UBound(WIPJobs)
                If Trim(WIPJobs(i)) <> "" Then
                    MainForm.lst.AddItem WIPJobs(i)
                End If
            Next i
        End If

        ' Reset flag
        UpdatingCheckboxes = False
        DisplayStatusMessage "Jobs in WIP displayed", "Info"
    End If
    Exit Sub

Error_Handler:
    UpdatingCheckboxes = False
    SystemCore.HandleStandardErrors Err.Number, "ShowJobsInWIP", "UserInterface"
End Sub

' **Purpose**: Accept quote and convert to job with validation
' **Original**: Interface_VBA/Main.frm.AcceptQuote_Click business logic
' **Parameters**:
'   - MainForm (Object): Main form reference
' **Returns**: None (Subroutine)
' **File Dependencies**: FAcceptQuote form
' **Form Usage**: Called from Main.AcceptQuote_Click
Public Sub AcceptQuote(MainForm As Object)
    On Error GoTo Error_Handler

    ' Validate workflow prerequisites before proceeding
    If Not SystemCore.ValidateWorkflowPrerequisites("AcceptQuote", MainForm) Then
        Exit Sub
    End If

    If ShowForm("FAcceptQuote", True) Then
        DisplayStatusMessage "Accept quote form opened", "Info"
    End If

    Exit Sub

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "AcceptQuote", "UserInterface"
End Sub

' **Purpose**: Close selected job with validation
' **Original**: Interface_VBA/MainInterfaceManager.bas.CloseJob
' **Parameters**:
'   - MainForm (Object): Main form reference
' **Returns**: Boolean - True if job closed successfully
' **File Dependencies**: Selected job file, Archive directory
' **Form Usage**: Called from Main.CloseJob_Click
Public Function CloseJob(MainForm As Object) As Boolean
    Dim SelectedFile As String

    On Error GoTo Error_Handler

    ' Validate workflow prerequisites before proceeding
    If Not SystemCore.ValidateWorkflowPrerequisites("CloseJob", MainForm) Then
        CloseJob = False
        Exit Function
    End If

    SelectedFile = MainForm.lst.Value
    If InStr(SelectedFile, "*") > 0 Then
        SelectedFile = Left(SelectedFile, Len(SelectedFile) - 2)
    End If

    ' Confirm job closure
    If SystemCore.ShowQuestion("Are you sure you want to close job: " & SelectedFile & "?", "Confirm Job Closure") = vbYes Then
        ' Move job from WIP to Archive
        If WorkflowManagement.MoveJobToArchive(SelectedFile) Then
            ' Refresh main form display
            RefreshMainForm MainForm
            SystemCore.ShowInformation "Job " & SelectedFile & " closed and moved to Archive successfully", "Job Closed"
            CloseJob = True
        Else
            SystemCore.ShowError "Failed to close job: " & SelectedFile, "Job Closure Error"
            CloseJob = False
        End If
    Else
        CloseJob = False
    End If

    Exit Function

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "CloseJob", "UserInterface"
    CloseJob = False
End Function

' ===================================================================
' CONTRACT TEMPLATE MANAGEMENT
' ===================================================================

' **Purpose**: Create new contract template item
' **Original**: Interface_VBA/Main.frm.but_CreateCTItem_Click business logic
' **Parameters**: None
' **Returns**: None (Subroutine)
' **File Dependencies**: Contract templates directory
' **Form Usage**: Called from Main.but_CreateCTItem_Click
Public Sub CreateContractTemplateItem()
    On Error GoTo Error_Handler

    If WorkflowManagement.CreateContractTemplate() Then
        DisplayStatusMessage "Contract template created successfully", "Info"
    End If

    Exit Sub

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "CreateContractTemplateItem", "UserInterface"
End Sub

' **Purpose**: Edit existing contract template item
' **Original**: Interface_VBA/Main.frm.but_EditCTItem_Click business logic
' **Parameters**:
'   - MainForm (Object): Main form reference
' **Returns**: None (Subroutine)
' **File Dependencies**: Contract templates, selected file
' **Form Usage**: Called from Main.but_EditCTItem_Click
Public Sub EditContractTemplateItem(MainForm As Object)
    Dim ContractsPath As String
    Dim MyName As String
    Dim ContractFiles() As String
    Dim FileCount As Integer
    Dim i As Integer
    Dim SelectedFile As String

    On Error GoTo Error_Handler

    Dim RootPath As String
    RootPath = DataOperations.GetRootPath()
    If Right(RootPath, 1) <> "\" Then RootPath = RootPath & "\"
    ContractsPath = RootPath & "Contracts\"

    ' Check if Contracts folder exists
    If Dir(ContractsPath, vbDirectory) = "" Then
        SystemCore.ShowError "Contracts folder not found: " & ContractsPath, "Folder Not Found"
        Exit Sub
    End If

    ' Clear FList and populate with contract template files
    FList.lst.Clear

    ' Scan Contracts directory for files
    MyName = Dir(ContractsPath, vbDirectory)
    If MyName = "" Then
        SystemCore.ShowError "No contract templates found in: " & ContractsPath, "No Templates"
        Exit Sub
    End If

    FileCount = 0
    ReDim ContractFiles(1 To 10000)

    Do Until MyName = ""
        ' Skip current and parent directory entries
        If MyName <> "." And MyName <> ".." Then
            FileCount = FileCount + 1
            ContractFiles(FileCount) = MyName
            ' Add to FList without .xls extension (per original implementation)
            FList.lst.AddItem Left(MyName, Len(MyName) - 4)
        End If
        MyName = Dir
    Loop

    ' Show FList form for user selection
    FList.Show

    ' Get selected file from FList
    SelectedFile = FList.lst.Value

    ' Open selected contract template for editing
    If WorkflowManagement.EditContractTemplate(SelectedFile) Then
        DisplayStatusMessage "Contract template " & SelectedFile & " opened for editing", "Info"
        ' Unload main form as per original implementation
        Unload MainForm
    End If

    Exit Sub

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "EditContractTemplateItem", "UserInterface"
End Sub

' **Purpose**: Show contracts folder
' **Original**: Interface_VBA/Main.frm.butShowContractsFolder_Click business logic
' **Parameters**: None
' **Returns**: None (Subroutine)
' **File Dependencies**: Contracts directory
' **Form Usage**: Called from Main.butShowContractsFolder_Click
Public Sub ShowContractsFolder()
    On Error GoTo Error_Handler

    ' Use exact original logic from Interface_VBA/Main.frm.butShowContractsFolder_Click
    Shell "C:\WINDOWS\explorer.exe """ & ActiveWorkbook.Path & "\Contracts" & """", vbNormalFocus

    Exit Sub

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "ShowContractsFolder", "UserInterface"
End Sub

' **Purpose**: Jump the gun workflow automation - create job directly from template
' **Original**: Interface_VBA/Main.frm.JumpTheGun_Click business logic
' **Parameters**: None
' **Returns**: None (Subroutine)
' **File Dependencies**: Templates/_Enq.xls
' **Form Usage**: Called from Main.JumpTheGun_Click
Public Sub JumpTheGun()
    On Error GoTo Error_Handler

    If WorkflowManagement.JumpTheGun() Then
        DisplayStatusMessage "Jump the gun workflow started successfully", "Info"
    End If

    Exit Sub

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "JumpTheGun", "UserInterface"
End Sub

' **Purpose**: Create WIP job from contract template
' **Original**: Interface_VBA/Main.frm.ContractWork_Click business logic
' **Parameters**: None
' **Returns**: None (Subroutine)
' **File Dependencies**: Contracts directory, selected contract template
' **Form Usage**: Called from Main.ContractWork_Click
Public Sub ContractWork()
    On Error GoTo Error_Handler

    If WorkflowManagement.ContractWork() Then
        DisplayStatusMessage "Contract work workflow started successfully", "Info"
    End If

    Exit Sub

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "ContractWork", "UserInterface"
End Sub

' ===================================================================
' JOB CARD MANAGEMENT
' ===================================================================

' **Purpose**: Edit job card with validation
' **Original**: Interface_VBA/Main.frm.But_EditJC_Click business logic - Edit selected WIP job
' **Parameters**:
'   - MainForm (Object): Main form reference
' **Returns**: None (Subroutine)
' **File Dependencies**: FJobCard form, WIP job files
' **Form Usage**: Called from Main.But_EditJC_Click
Public Sub EditJobCard(MainForm As Object)
    Dim SelectedJob As String
    Dim JobPath As String
    Dim RootPath As String
    Dim JobWB As Workbook

    On Error GoTo Error_Handler

    ' Validate workflow prerequisites before proceeding
    If Not SystemCore.ValidateWorkflowPrerequisites("EditJobCard", MainForm) Then
        Exit Sub
    End If

    ' Get selected job name and strip status indicator if present
    SelectedJob = MainForm.lst.Value
    If InStr(1, SelectedJob, "*") > 1 Then
        SelectedJob = Left(SelectedJob, Len(SelectedJob) - 2)
    End If

    ' Build job file path based on context (Edit Job Card should work on WIP files)
    RootPath = MainForm.Main_MasterPath.Value
    If Right(RootPath, 1) <> "\" Then RootPath = RootPath & "\"

    If MainForm.JobsInWIP.Value = True Then
        ' JobsInWIP context - job number from database, ensure .xls extension
        If Right(SelectedJob, 4) <> ".xls" Then
            JobPath = RootPath & "WIP\" & SelectedJob & ".xls"
        Else
            JobPath = RootPath & "WIP\" & SelectedJob
        End If
    Else
        ' Regular WIP context - file name from directory listing
        JobPath = RootPath & "WIP\" & SelectedJob & ".xls"
    End If

    ' Validate job file exists in WIP
    If Not DataOperations.FileExists(JobPath) Then
        SystemCore.ShowError "WIP job file not found: " & JobPath, "File Not Found"
        Exit Sub
    End If

    ' Open job file for editing (not read-only like OpenJob)
    Set JobWB = DataOperations.SafeOpenWorkbook(JobPath, False)
    If JobWB Is Nothing Then
        SystemCore.ShowError "Unable to open job file for editing: " & JobPath, "File Access Error"
        Exit Sub
    End If

    ' Activate job file and Job Card sheet as per original workflow
    JobWB.Activate
    If JobWB.Worksheets.Count > 1 Then
        ' Activate Job Card sheet if it exists
        On Error Resume Next
        JobWB.Worksheets("Job Card").Activate
        On Error GoTo Error_Handler
    End If

    ' Load job data into form for editing
    WorkflowManagement.LoadJobCardData FJobCard, JobPath
    FJobCard.Show

    Exit Sub

Error_Handler:
    If Not JobWB Is Nothing Then DataOperations.SafeCloseWorkbook JobWB, False
    SystemCore.HandleStandardErrors Err.Number, "EditJobCard", "UserInterface"
End Sub

' **Purpose**: Open job card form with validation
' **Original**: Interface_VBA/Main.frm.OpenJob_Click business logic
' **Parameters**:
'   - MainForm (Object): Main form reference
' **Returns**: None (Subroutine)
' **File Dependencies**: FJobCard form, System_Status field
' **Form Usage**: Called from Main.OpenJob_Click
Public Sub OpenJob(MainForm As Object)
    Dim SelectedJob As String
    Dim JobPath As String
    Dim RootPath As String
    Dim JobWB As Workbook

    On Error GoTo Error_Handler

    ' Validate workflow prerequisites before proceeding
    If Not SystemCore.ValidateWorkflowPrerequisites("OpenJob", MainForm) Then
        Exit Sub
    End If

    ' Get selected job name and strip status indicator if present
    SelectedJob = MainForm.lst.Value
    If InStr(1, SelectedJob, "*") > 1 Then
        SelectedJob = Left(SelectedJob, Len(SelectedJob) - 2)
    End If

    ' Build job file path based on context
    RootPath = MainForm.Main_MasterPath.Value
    If Right(RootPath, 1) <> "\" Then RootPath = RootPath & "\"

    If MainForm.JobsInWIP.Value = True Then
        ' JobsInWIP context - job number from database, ensure .xls extension
        If Right(SelectedJob, 4) <> ".xls" Then
            JobPath = RootPath & "WIP\" & SelectedJob & ".xls"
        Else
            JobPath = RootPath & "WIP\" & SelectedJob
        End If
    Else
        ' Regular WIP context - file name from directory listing
        JobPath = RootPath & "WIP\" & SelectedJob & ".xls"
    End If

    ' Validate job file exists
    If Not DataOperations.FileExists(JobPath) Then
        SystemCore.ShowError "Job file not found: " & JobPath, "File Not Found"
        Exit Sub
    End If

    ' Check job status from file before opening (as per original system validation)
    Set JobWB = DataOperations.SafeOpenWorkbook(JobPath, True)
    If JobWB Is Nothing Then
        SystemCore.ShowError "Unable to open job file: " & JobPath, "File Access Error"
        Exit Sub
    End If

    ' Check if quote has been accepted (support multiple valid status formats)
    Dim JobStatus As String
    On Error Resume Next
    JobStatus = JobWB.Worksheets("Admin").Range("B88").Value ' System_Status field
    On Error GoTo Error_Handler

    ' Normalize status for comparison (handle multiple valid formats)
    Dim NormalizedStatus As String
    NormalizedStatus = LCase(Trim(Replace(JobStatus, " ", "")))

    ' Accept multiple valid status formats: "Quote Accepted", "quote accepted", "QuoteAccepted", etc.
    If NormalizedStatus <> "quoteaccepted" And NormalizedStatus <> "accepted" Then
        DataOperations.SafeCloseWorkbook JobWB, True
        SystemCore.ShowWarning "Please accept this quote first before opening the job card." & vbCrLf & _
                              "Current status: '" & JobStatus & "'" & vbCrLf & _
                              "Expected: 'Quote Accepted'", "Quote Not Accepted"
        Exit Sub
    End If

    ' Keep job file open and activate it as per original Interface_VBA/Main.frm.OpenJob_Click
    JobWB.Activate
    If JobWB.Worksheets.Count > 1 Then
        ' Activate Job Card sheet if it exists
        On Error Resume Next
        JobWB.Worksheets("Job Card").Activate
        On Error GoTo Error_Handler
    End If

    ' Show job card form and load job data as per original workflow
    WorkflowManagement.LoadJobCardData FJobCard, JobPath
    FJobCard.Show

    Exit Sub

Error_Handler:
    If Not JobWB Is Nothing Then DataOperations.SafeCloseWorkbook JobWB, True
    SystemCore.HandleStandardErrors Err.Number, "OpenJob", "UserInterface"
End Sub

' ===================================================================
' SEARCH AND HISTORY MANAGEMENT
' ===================================================================

' **Purpose**: Edit search database
' **Original**: Interface_VBA/Main.frm.butEditSearch_Click business logic
' **Parameters**:
'   - MainForm (Object): Main form reference
' **Returns**: None (Subroutine)
' **File Dependencies**: Search.xls database
' **Form Usage**: Called from Main.butEditSearch_Click
Public Sub EditSearchDatabase(MainForm As Object)
    Dim SearchPath As String
    Dim SearchWB As Workbook
    Dim SearchWS As Worksheet
    Dim LastCol As Long
    Dim LastRow As Long
    Dim SortRange As Range

    On Error GoTo Error_Handler

    Dim RootPath As String
    RootPath = DataOperations.GetRootPath()
    If Right(RootPath, 1) <> "\" Then RootPath = RootPath & "\"
    SearchPath = RootPath & "Search.xls"

    If Not DataOperations.FileExists(SearchPath) Then
        SystemCore.ShowError "Search database not found: " & SearchPath, "File Not Found"
        Exit Sub
    End If

    ' Open Search.xls for editing (not readonly)
    Set SearchWB = DataOperations.SafeOpenWorkbook(SearchPath, False)
    If SearchWB Is Nothing Then
        SystemCore.ShowError "Unable to open search database", "File Open Error"
        Exit Sub
    End If

    Set SearchWS = SearchWB.Worksheets(1)

    ' Find last column and row as per original implementation
    SearchWS.Range("A1").Select
    Selection.End(xlToRight).Select
    LastCol = ActiveCell.Column

    SearchWS.Range("A1").Select
    Selection.End(xlDown).Select
    LastRow = ActiveCell.Row

    ' Select data range excluding headers and sort by column E descending
    If LastRow > 1 And LastCol > 1 Then
        Set SortRange = SearchWS.Range("A2", SearchWS.Cells(LastRow, LastCol))
        SortRange.Sort Key1:=SearchWS.Range("E2"), Order1:=xlDescending, _
                      Header:=xlYes, OrderCustom:=1, MatchCase:=False, _
                      Orientation:=xlTopToBottom, DataOption1:=xlSortTextAsNumbers
    End If

    ' Select C3 as per original implementation
    SearchWS.Range("C3").Select

    ' Hide main form as per original
    MainForm.Hide

    DisplayStatusMessage "Search database opened for editing - sorted by date descending", "Info"
    Exit Sub

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "EditSearchDatabase", "UserInterface"
End Sub

' **Purpose**: Show search history
' **Original**: Interface_VBA/Main.frm.butSearchHistory_Click business logic
' **Parameters**:
'   - MainForm (Object): Main form reference
' **Returns**: None (Subroutine)
' **File Dependencies**: Search form, search history data
' **Form Usage**: Called from Main.butSearchHistory_Click
Public Sub ShowSearchHistory(MainForm As Object)
    On Error GoTo Error_Handler

    If ShowForm("frmSearch", True) Then
        ' Load search history into the form
        BusinessLogic.LoadSearchHistory frmSearch
        DisplayStatusMessage "Search history displayed", "Info"
    End If

    Exit Sub

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "ShowSearchHistory", "UserInterface"
End Sub

' **Purpose**: Show job history
' **Original**: Interface_VBA/Main.frm.butJobHistory_Click business logic
' **Parameters**:
'   - MainForm (Object): Main form reference
' **Returns**: None (Subroutine)
' **File Dependencies**: Archive files, search database
' **Form Usage**: Called from Main.butJobHistory_Click
Public Sub ShowJobHistory(MainForm As Object)
    On Error GoTo Error_Handler

    Dim JobHistory As Variant
    JobHistory = BusinessLogic.GetJobHistory()

    If IsArray(JobHistory) Then
        PopulateList FList, JobHistory, "Job History"
        FList.Show
        DisplayStatusMessage "Job history displayed", "Info"
    Else
        SystemCore.ShowInformation "No job history found.", "Job History"
    End If

    Exit Sub

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "ShowJobHistory", "UserInterface"
End Sub

' **Purpose**: Show quote history
' **Original**: Interface_VBA/Main.frm.butQuoteHistory_Click business logic
' **Parameters**:
'   - MainForm (Object): Main form reference
' **Returns**: None (Subroutine)
' **File Dependencies**: Quote files, search database
' **Form Usage**: Called from Main.butQuoteHistory_Click
Public Sub ShowQuoteHistory(MainForm As Object)
    On Error GoTo Error_Handler

    Dim QuoteHistory As Variant
    QuoteHistory = BusinessLogic.GetQuoteHistory()

    If IsArray(QuoteHistory) Then
        PopulateList FList, QuoteHistory, "Quote History"
        FList.Show
        DisplayStatusMessage "Quote history displayed", "Info"
    Else
        SystemCore.ShowInformation "No quote history found.", "Quote History"
    End If

    Exit Sub

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "ShowQuoteHistory", "UserInterface"
End Sub

' **Purpose**: Sort search database
' **Original**: Interface_VBA/Main.frm.butSortSearch_Click business logic
' **Parameters**: None
' **Returns**: None (Subroutine)
' **File Dependencies**: Search.xls database
' **Form Usage**: Called from Main.butSortSearch_Click
Public Sub SortSearchDatabase()
    On Error GoTo Error_Handler

    If BusinessLogic.SortSearchDatabase() Then
        DisplayStatusMessage "Search database sorted successfully", "Info"
    Else
        SystemCore.ShowError "Failed to sort search database", "Sort Error"
    End If

    Exit Sub

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "SortSearchDatabase", "UserInterface"
End Sub

' **Purpose**: Mark quote as called through with validation
' **Original**: Interface_VBA/Main.frm.CalledThrough_Click business logic
' **Parameters**:
'   - MainForm (Object): Main form reference
' **Returns**: None (Subroutine)
' **File Dependencies**: Selected quote file
' **Form Usage**: Called from Main.CalledThrough_Click
Public Sub MarkQuoteCalledThrough(MainForm As Object)
    Dim SelectedFile As String

    On Error GoTo Error_Handler

    ' Validate workflow prerequisites before proceeding
    If Not SystemCore.ValidateWorkflowPrerequisites("SubmitQuote", MainForm) Then
        Exit Sub
    End If

    SelectedFile = MainForm.lst.Value
    If InStr(SelectedFile, "*") > 0 Then
        SelectedFile = Left(SelectedFile, Len(SelectedFile) - 2)
    End If

    If BusinessLogic.MarkQuoteCalledThrough(SelectedFile) Then
        RefreshMainInterface
        DisplayStatusMessage "Quote " & SelectedFile & " marked as called through", "Info"
    Else
        SystemCore.ShowError "Failed to mark quote as called through", "Update Error"
    End If

    Exit Sub

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "MarkQuoteCalledThrough", "UserInterface"
End Sub

' ===================================================================
' PUBLIC UTILITY FUNCTIONS FOR FORMS
' ===================================================================

' **Purpose**: Display status message in main form
' **Parameters**:
'   - Message (String): Message to display
'   - MessageType (String, Optional): Type of message (Info, Warning, Error)
' **Returns**: None (Subroutine)
' **Dependencies**: Main form
' **Side Effects**: Updates main form status label
' **Errors**: None
Public Sub DisplayStatusMessage(Message As String, Optional MessageType As String = "Info")
    On Error Resume Next
    Main.Label1.Caption = Format(Now, "hh:mm:ss") & " - " & Message

    ' Could add color coding based on message type
    Select Case UCase(MessageType)
        Case "WARNING"
            Main.Label1.ForeColor = RGB(255, 165, 0) ' Orange
        Case "ERROR"
            Main.Label1.ForeColor = RGB(255, 0, 0) ' Red
        Case Else
            Main.Label1.ForeColor = RGB(0, 0, 0) ' Black
    End Select
    On Error GoTo 0
End Sub

' **Purpose**: Update main form progress indicator
' **Parameters**:
'   - Progress (Integer): Progress percentage (0-100)
'   - Operation (String, Optional): Description of current operation
' **Returns**: None (Subroutine)
' **Dependencies**: Main form
' **Side Effects**: Updates progress display in main form
' **Errors**: None
Public Sub UpdateProgress(Progress As Integer, Optional Operation As String = "")
    On Error Resume Next
    If Operation <> "" Then
        Main.Label1.Caption = Operation & " (" & Progress & "%)"
    Else
        Main.Label1.Caption = "Progress: " & Progress & "%"
    End If

    ' Force immediate screen update
    DoEvents
    On Error GoTo 0
End Sub

' **Purpose**: Reset main form to ready state
' **Parameters**: None
' **Returns**: None (Subroutine)
' **Dependencies**: Main form
' **Side Effects**: Resets form display to ready state
' **Errors**: None
Public Sub ResetToReady()
    On Error Resume Next
    Main.Label1.Caption = "Ready"
    Main.Label1.ForeColor = RGB(0, 0, 0)
    On Error GoTo 0
End Sub

' **Purpose**: Validate that a selection has been made in list dialog
' **Original**: Interface_VBA/FList.frm ValidateSelection logic
' **Parameters**:
'   - ListForm (Object): List form containing selection
' **Returns**: Boolean - True if selection is valid, False otherwise
' **Dependencies**: List form controls
' **Side Effects**: May show validation messages
' **Errors**: Returns False on validation failure
Public Function ValidateSelection(ListForm As Object) As Boolean
    On Error GoTo Error_Handler

    ValidateSelection = False

    ' Check if a selection has been made in the list
    If ListForm.lst.ListIndex >= 0 Then
        ValidateSelection = True
    Else
        SystemCore.ShowWarning "Please select an item from the list.", "Selection Required"
        ValidateSelection = False
    End If

    Exit Function

Error_Handler:
    ValidateSelection = False
    SystemCore.HandleStandardErrors Err.Number, "ValidateSelection", "UserInterface"
End Function

' **Purpose**: Open Search.xls database directly (replaces frmSearch.Show)
' **Original**: Interface_VBA/Main.frm Search_Click() method
' **Parameters**: None
' **Returns**: Boolean - True if search database opened successfully
' **CLAUDE.md Compliance**: Uses existing Search.xls file, no new forms created
Public Function OpenSearchDatabase() As Boolean
    Dim SearchPath As String
    Dim SearchWB As Workbook

    On Error GoTo Error_Handler

    Dim RootPath As String
    RootPath = DataOperations.GetRootPath()
    If Right(RootPath, 1) <> "\" Then RootPath = RootPath & "\"
    SearchPath = RootPath & "Search.xls"

    ' Check if search file exists
    If Dir(SearchPath) = "" Then
        SystemCore.ShowError "Search database not found: " & SearchPath, "Search Database Missing"
        OpenSearchDatabase = False
        Exit Function
    End If

    ' Open search.xls in read-only mode as per original implementation
    Set SearchWB = DataOperations.SafeOpenWorkbook(SearchPath, True)
    If SearchWB Is Nothing Then
        SystemCore.ShowError "Unable to open search database", "Search Database Error"
        OpenSearchDatabase = False
        Exit Function
    End If

    ' Select cell B1 and hide main form as per original
    SearchWB.Sheets("search").Range("B1").Select
    Main.Hide

    ' Run the search menu macro as per original
    On Error Resume Next
    Dim SearchFileName As String
    SearchFileName = SearchWB.Name
    Application.Run SearchFileName & "!Show_Search_Menu"

    ' If that fails, try without the macro call - just leave the file open
    If Err.Number <> 0 Then
        Err.Clear
        ' Just leave the search file open and selected - user can navigate manually
    End If
    On Error GoTo Error_Handler

    OpenSearchDatabase = True
    Exit Function

Error_Handler:
    SystemCore.LogError Err.Number, Err.Description, "OpenSearchDatabase", "UserInterface"
    If Not SearchWB Is Nothing Then SearchWB.Close False
    OpenSearchDatabase = False
End Function

