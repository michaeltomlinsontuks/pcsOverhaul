Attribute VB_Name = "DataManager"
' **Purpose**: All file operations, Excel data access, and directory management
' **CLAUDE.md Compliance**: Maintains all directory structure requirements, 32/64-bit compatibility
Option Explicit

' ===================================================================
' CONSTANTS AND PRIVATE VARIABLES
' ===================================================================

Private Const NUMBERS_FILE As String = "Templates\number_tracking.xls"
Private Const ROOT_PATH As String = ""

' ===================================================================
' FILE SYSTEM OPERATIONS (CLAUDE.md: Preserve directory structure)
' ===================================================================

' **Purpose**: Get the root path for PCS system operations
' **Parameters**: None
' **Returns**: String - Root directory path for PCS system
' **Dependencies**: ThisWorkbook object
' **Side Effects**: None
' **Errors**: Returns empty string if workbook path unavailable
' **CLAUDE.md Compliance**: Preserves existing directory structure access
Public Function GetRootPath() As String
    On Error GoTo Error_Handler

    If ROOT_PATH = "" Then
        GetRootPath = ThisWorkbook.Path
    Else
        GetRootPath = ROOT_PATH
    End If
    Exit Function

Error_Handler:
    CoreFramework.LogError Err.Number, Err.Description, "GetRootPath", "DataManager"
    GetRootPath = ""
End Function

' **Purpose**: Validate all required PCS directory structure exists
' **Parameters**: None
' **Returns**: Boolean - True if all directories exist, False if any missing
' **Dependencies**: DirExists() for individual directory checking
' **Side Effects**: Logs missing directories to error log
' **Errors**: Logs each missing directory, does not create directories
' **CLAUDE.md Compliance**: Preserves existing directory structure, no changes made
Public Function ValidateDirectoryStructure() As Boolean
    Dim RequiredDirs As Variant
    Dim i As Integer

    On Error GoTo Error_Handler

    RequiredDirs = Array("Enquiries", "Quotes", "WIP", "Archive", "Contracts", _
                        "Customers", "Templates", "Job Templates", "images", "Backups")

    For i = 0 To UBound(RequiredDirs)
        If Not DirExists(GetRootPath & "\" & RequiredDirs(i)) Then
            ValidateDirectoryStructure = False
            CoreFramework.LogError 0, "Missing directory: " & RequiredDirs(i), "ValidateDirectoryStructure", "DataManager"
            Exit Function
        End If
    Next i

    ValidateDirectoryStructure = True
    Exit Function

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "ValidateDirectoryStructure", "DataManager"
    ValidateDirectoryStructure = False
End Function

' **Purpose**: Create missing PCS directory structure
' **Parameters**: None
' **Returns**: Boolean - True if all directories created successfully, False if failed
' **Dependencies**: DirExists() for checking, MkDir for creation
' **Side Effects**: Creates missing directories in file system
' **Errors**: Returns False if any directory creation fails
' **CLAUDE.md Compliance**: Only creates missing directories, preserves existing structure
Public Function CreateDirectoryStructure() As Boolean
    Dim RequiredDirs As Variant
    Dim i As Integer
    Dim DirPath As String

    On Error GoTo Error_Handler

    RequiredDirs = Array("Enquiries", "Quotes", "WIP", "Archive", "Contracts", _
                        "Customers", "Templates", "Job Templates", "images", "Backups")

    For i = 0 To UBound(RequiredDirs)
        DirPath = GetRootPath & "\" & RequiredDirs(i)
        If Not DirExists(DirPath) Then
            MkDir DirPath
            CoreFramework.LogError 0, "Created missing directory: " & RequiredDirs(i), "CreateDirectoryStructure", "DataManager"
        End If
    Next i

    CreateDirectoryStructure = True
    Exit Function

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "CreateDirectoryStructure", "DataManager"
    CreateDirectoryStructure = False
End Function

' **Purpose**: Check if directory exists
' **Parameters**:
'   - DirPath (String): Full path to directory to check
' **Returns**: Boolean - True if directory exists, False if not
' **Dependencies**: VBA Dir function
' **Side Effects**: None
' **Errors**: Returns False if error occurs during check
Public Function DirExists(ByVal DirPath As String) As Boolean
    On Error GoTo Error_Handler
    DirExists = (Dir(DirPath, vbDirectory) <> "")
    Exit Function

Error_Handler:
    DirExists = False
End Function

' **Purpose**: Check if file exists
' **Parameters**:
'   - FilePath (String): Full path to file to check
' **Returns**: Boolean - True if file exists, False if not
' **Dependencies**: VBA Dir function
' **Side Effects**: None
' **Errors**: Returns False if error occurs during check
Public Function FileExists(ByVal FilePath As String) As Boolean
    On Error GoTo Error_Handler
    FileExists = (Dir(FilePath) <> "")
    Exit Function

Error_Handler:
    FileExists = False
End Function

' **Purpose**: Get list of files in specified directory
' **Parameters**:
'   - DirectoryName (String): Name of subdirectory under root path
' **Returns**: Variant - Array of filenames, empty array if no files or error
' **Dependencies**: GetRootPath(), DirExists(), VBA Dir function
' **Side Effects**: None
' **Errors**: Returns empty array if directory not found or access error
Public Function GetFileList(ByVal DirectoryName As String) As Variant
    Dim DirPath As String
    Dim FileName As String
    Dim FileList() As String
    Dim FileCount As Integer

    On Error GoTo Error_Handler

    DirPath = GetRootPath & "\" & DirectoryName & "\"

    If Not DirExists(DirPath) Then
        CoreFramework.LogError CoreFramework.ERR_PATH_NOT_FOUND, "Directory not found: " & DirPath, "GetFileList", "DataManager"
        GetFileList = Array()
        Exit Function
    End If

    FileName = Dir(DirPath & "*.xls*")
    FileCount = 0

    Do While FileName <> ""
        ReDim Preserve FileList(FileCount)
        FileList(FileCount) = FileName
        FileCount = FileCount + 1
        FileName = Dir
    Loop

    If FileCount > 0 Then
        GetFileList = FileList
    Else
        GetFileList = Array()
    End If
    Exit Function

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "GetFileList", "DataManager"
    GetFileList = Array()
End Function

' **Purpose**: Get file list with status indicators for display
' **Parameters**:
'   - DirectoryName (String): Name of subdirectory under root path
'   - FormObject (Object): Form object to populate with file list
' **Returns**: None (Subroutine)
' **Dependencies**: GetRootPath(), DirExists(), GetValueFromClosedWorkbook()
' **Side Effects**: Populates form object with file list and status indicators
' **Errors**: Exits function if directory not found, logs errors
' **CLAUDE.md Compliance**: Replaces legacy a_ListFiles.bas functionality
Public Sub GetFileListWithStatus(ByVal DirectoryName As String, ByRef FormObject As Object)
    Dim Files(1 To 100000) As String
    Dim FullFilePath As String, MyName As String
    Dim GroupCount As Integer
    Dim i As Integer
    Dim x As String
    Dim StatusValue As String

    On Error GoTo Error_Handler

    FullFilePath = GetRootPath & "\" & DirectoryName & "\"

    MyName = Dir(FullFilePath, vbDirectory)
    If MyName = "" Then
        MsgBox "Folder Not Found: " & DirectoryName, vbOKOnly, "Error"
        Exit Sub
    End If

    ' Store list of files
    Do Until MyName = ""
        If MyName <> "." And MyName <> ".." And Right(UCase(MyName), 4) = ".XLS" Then
            GroupCount = GroupCount + 1
            Files(GroupCount) = MyName
        End If
        MyName = Dir
    Loop

    ' Populate form with files and status indicators
    For i = 1 To GroupCount
        x = Files(i)

        ' Check status based on directory type
        Select Case UCase(DirectoryName)
            Case "WIP"
                StatusValue = GetValueFromClosedWorkbook(FullFilePath & x, "ADMIN", "B88")
                If UCase(StatusValue) = "QUOTE ACCEPTED" Then
                    FormObject.AddItem Left(x, Len(x) - 4) & " *"
                Else
                    FormObject.AddItem Left(x, Len(x) - 4)
                End If

            Case "QUOTES"
                StatusValue = GetValueFromClosedWorkbook(FullFilePath & x, "Admin", "B88")
                If UCase(StatusValue) = "NEW QUOTE" Then
                    FormObject.AddItem Left(x, Len(x) - 4) & " *"
                Else
                    FormObject.AddItem Left(x, Len(x) - 4)
                End If

            Case Else
                FormObject.AddItem Left(x, Len(x) - 4)
        End Select
    Next i

    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "GetFileListWithStatus", "DataManager"
End Sub

' **Purpose**: Create backup copy of file with timestamp
' **Parameters**:
'   - FilePath (String): Full path to file to backup
' **Returns**: Boolean - True if backup created successfully, False if failed
' **Dependencies**: GetRootPath(), DirExists(), MkDir, FileCopy
' **Side Effects**: Creates backup file in Backups directory
' **Errors**: Returns False if backup creation fails
Public Function CreateBackup(ByVal FilePath As String) As Boolean
    Dim BackupPath As String
    Dim BackupDir As String

    On Error GoTo Error_Handler

    BackupDir = GetRootPath & "\Backups\"
    If Not DirExists(BackupDir) Then
        MkDir BackupDir
    End If

    BackupPath = BackupDir & Format(Now, "yyyymmdd_hhmmss_") & Dir(FilePath)

    FileCopy FilePath, BackupPath
    CreateBackup = True
    Exit Function

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "CreateBackup", "DataManager"
    CreateBackup = False
End Function

' **Purpose**: Count files in folder matching pattern
' **Parameters**:
'   - FolderPath (String): Full path to folder
'   - FilePattern (String): File pattern to match (e.g., "*.xls")
' **Returns**: Long - Number of matching files
' **Dependencies**: DirExists(), VBA Dir function
' **Side Effects**: None
' **Errors**: Returns 0 if folder not found or error
Public Function CountFilesInFolder(ByVal FolderPath As String, ByVal FilePattern As String) As Long
    Dim FileName As String
    Dim FileCount As Long

    On Error GoTo Error_Handler

    If Not DirExists(FolderPath) Then
        CountFilesInFolder = 0
        Exit Function
    End If

    FileName = Dir(FolderPath & "\" & FilePattern)
    FileCount = 0

    Do While FileName <> ""
        FileCount = FileCount + 1
        FileName = Dir
    Loop

    CountFilesInFolder = FileCount
    Exit Function

Error_Handler:
    CountFilesInFolder = 0
End Function

' ===================================================================
' WORKBOOK OPERATIONS (CLAUDE.md: 32/64-bit Excel compatibility)
' ===================================================================

' **Purpose**: Safely open Excel workbook with error handling and validation
' **Parameters**:
'   - FilePath (String): Full path to Excel file to open
' **Returns**: Workbook object if successful, Nothing if failed
' **Dependencies**: FileExists(), CoreFramework.ErrorHandler for error logging
' **Side Effects**: Opens workbook in Excel application, logs errors if failed
' **Errors**: Returns Nothing on file not found, permission denied, or corruption
' **CLAUDE.md Compliance**: Maintains 32/64-bit Excel compatibility
Public Function SafeOpenWorkbook(ByVal FilePath As String) As Workbook
    Dim wb As Workbook

    On Error GoTo Error_Handler

    If Not FileExists(FilePath) Then
        CoreFramework.LogError CoreFramework.ERR_FILE_NOT_FOUND, "File not found: " & FilePath, "SafeOpenWorkbook", "DataManager"
        Set SafeOpenWorkbook = Nothing
        Exit Function
    End If

    Set wb = Workbooks.Open(FilePath, ReadOnly:=False, UpdateLinks:=False)
    Set SafeOpenWorkbook = wb
    Exit Function

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "SafeOpenWorkbook", "DataManager"
    Set SafeOpenWorkbook = Nothing
End Function

' **Purpose**: Safely close workbook with optional save
' **Parameters**:
'   - wb (Workbook): Workbook object to close
'   - SaveChanges (Boolean, Optional): Whether to save changes (default True)
' **Returns**: Boolean - True if closed successfully, False if failed
' **Dependencies**: None
' **Side Effects**: Closes workbook in Excel application
' **Errors**: Returns False if close operation fails
Public Function SafeCloseWorkbook(ByRef wb As Workbook, Optional ByVal SaveChanges As Boolean = True) As Boolean
    On Error GoTo Error_Handler

    If Not wb Is Nothing Then
        wb.Close SaveChanges:=SaveChanges
        Set wb = Nothing
        SafeCloseWorkbook = True
    End If
    Exit Function

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "SafeCloseWorkbook", "DataManager"
    SafeCloseWorkbook = False
End Function

' **Purpose**: Create new workbook from template or blank
' **Parameters**: None
' **Returns**: Workbook - New workbook object, Nothing if failed
' **Dependencies**: Excel Workbooks collection
' **Side Effects**: Creates new workbook in Excel application
' **Errors**: Returns Nothing if creation fails
Public Function CreateNewWorkbook() As Workbook
    On Error GoTo Error_Handler

    Set CreateNewWorkbook = Workbooks.Add
    Exit Function

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "CreateNewWorkbook", "DataManager"
    Set CreateNewWorkbook = Nothing
End Function

' **Purpose**: Open workbook with enhanced security and validation
' **Parameters**:
'   - FilePath (String): Full path to Excel file to open
' **Returns**: Workbook object if successful, Nothing if failed
' **Dependencies**: FileExists(), SafeOpenWorkbook()
' **Side Effects**: Opens workbook with security settings
' **Errors**: Returns Nothing on validation failure
' **CLAUDE.md Compliance**: Replaces legacy Open_Book.bas functionality
Public Function OpenWorkbookSecure(ByVal FilePath As String) As Workbook
    Dim wb As Workbook

    On Error GoTo Error_Handler

    ' Additional validation for secure opening
    If Not FileExists(FilePath) Then
        CoreFramework.LogError CoreFramework.ERR_FILE_NOT_FOUND, "Secure open failed - file not found: " & FilePath, "OpenWorkbookSecure", "DataManager"
        Set OpenWorkbookSecure = Nothing
        Exit Function
    End If

    ' Check if file is already open
    On Error Resume Next
    Set wb = Workbooks(Dir(FilePath))
    On Error GoTo Error_Handler

    If Not wb Is Nothing Then
        ' File already open, return existing workbook
        Set OpenWorkbookSecure = wb
        Exit Function
    End If

    ' Open with security settings
    Set wb = Workbooks.Open(FilePath, _
                            UpdateLinks:=False, _
                            ReadOnly:=False, _
                            Format:=xlNormal, _
                            Password:="", _
                            WriteResPassword:="", _
                            IgnoreReadOnlyRecommended:=True, _
                            Origin:=xlWindows, _
                            Delimiter:="", _
                            Editable:=True, _
                            Notify:=False, _
                            Converter:=0, _
                            AddToMru:=False, _
                            Local:=False, _
                            CorruptLoad:=xlNormalLoad)

    Set OpenWorkbookSecure = wb
    Exit Function

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "OpenWorkbookSecure", "DataManager"
    Set OpenWorkbookSecure = Nothing
End Function

' **Purpose**: Delete worksheet from workbook safely
' **Parameters**:
'   - wb (Workbook): Workbook containing worksheet to delete
'   - SheetName (String): Name of worksheet to delete
' **Returns**: Boolean - True if deleted successfully, False if failed
' **Dependencies**: None
' **Side Effects**: Removes worksheet from workbook
' **Errors**: Returns False if worksheet not found or deletion fails
' **CLAUDE.md Compliance**: Replaces legacy Delete_Sheet.bas functionality
Public Function DeleteWorksheet(ByRef wb As Workbook, ByVal SheetName As String) As Boolean
    Dim ws As Worksheet

    On Error GoTo Error_Handler

    If wb Is Nothing Then
        DeleteWorksheet = False
        Exit Function
    End If

    ' Check if worksheet exists
    Set ws = Nothing
    On Error Resume Next
    Set ws = wb.Worksheets(SheetName)
    On Error GoTo Error_Handler

    If ws Is Nothing Then
        CoreFramework.LogError 0, "Worksheet not found: " & SheetName, "DeleteWorksheet", "DataManager"
        DeleteWorksheet = False
        Exit Function
    End If

    ' Disable alerts during deletion
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True

    DeleteWorksheet = True
    Exit Function

Error_Handler:
    Application.DisplayAlerts = True
    CoreFramework.HandleStandardErrors Err.Number, "DeleteWorksheet", "DataManager"
    DeleteWorksheet = False
End Function

' ===================================================================
' DATA ACCESS OPERATIONS
' ===================================================================

' **Purpose**: Get single cell value from Excel file
' **Parameters**:
'   - FilePath (String): Full path to Excel file
'   - SheetName (String): Name of worksheet
'   - CellAddress (String): Cell address (e.g., "A1")
' **Returns**: Variant - Cell value, empty string if error
' **Dependencies**: FileExists(), SafeOpenWorkbook(), SafeCloseWorkbook()
' **Side Effects**: Opens and closes workbook
' **Errors**: Returns empty string if file not found or cell access fails
Public Function GetValue(ByVal FilePath As String, ByVal SheetName As String, ByVal CellAddress As String) As Variant
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim CellValue As Variant

    On Error GoTo Error_Handler

    If Not FileExists(FilePath) Then
        GetValue = ""
        Exit Function
    End If

    Set wb = SafeOpenWorkbook(FilePath)
    If wb Is Nothing Then
        GetValue = ""
        Exit Function
    End If

    Set ws = wb.Worksheets(SheetName)
    CellValue = ws.Range(CellAddress).Value

    SafeCloseWorkbook wb, False

    GetValue = CellValue
    Exit Function

Error_Handler:
    If Not wb Is Nothing Then SafeCloseWorkbook wb, False
    CoreFramework.HandleStandardErrors Err.Number, "GetValue", "DataManager"
    GetValue = ""
End Function

' **Purpose**: Get cell value from closed workbook using Excel 4.0 macro
' **Parameters**:
'   - FilePath (String): Full path to Excel file
'   - SheetName (String): Name of worksheet
'   - CellAddress (String): Cell address (e.g., "A1")
' **Returns**: Variant - Cell value, empty string if error
' **Dependencies**: ExecuteExcel4Macro function
' **Side Effects**: None (does not open workbook)
' **Errors**: Returns empty string if file not found or macro execution fails
' **CLAUDE.md Compliance**: Enhanced version of legacy GetValue.bas functionality
Public Function GetValueFromClosedWorkbook(ByVal FilePath As String, ByVal SheetName As String, ByVal CellAddress As String) As Variant
    Dim Formula As String
    Dim TempCell As Range
    Dim arg As String

    On Error GoTo Error_Handler

    ' Check if file exists
    If Dir(FilePath) = "" Then
        GetValueFromClosedWorkbook = "File Not Found"
        Exit Function
    End If

    ' Create the Excel 4.0 macro argument
    arg = "'" & FilePath & "[" & Dir(FilePath) & "]" & SheetName & "'!" & _
          Range(CellAddress).Range("A1").Address(, , xlR1C1)

    ' Execute the macro
    GetValueFromClosedWorkbook = ExecuteExcel4Macro(arg)

    Exit Function

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "GetValueFromClosedWorkbook", "DataManager"
    GetValueFromClosedWorkbook = ""
End Function

' **Purpose**: Set single cell value in Excel file
' **Parameters**:
'   - FilePath (String): Full path to Excel file
'   - SheetName (String): Name of worksheet
'   - CellAddress (String): Cell address (e.g., "A1")
'   - Value (Variant): Value to set in cell
' **Returns**: Boolean - True if successful, False if failed
' **Dependencies**: SafeOpenWorkbook(), SafeCloseWorkbook()
' **Side Effects**: Opens workbook, modifies cell, saves and closes workbook
' **Errors**: Returns False if file access or cell update fails
Public Function SetValue(ByVal FilePath As String, ByVal SheetName As String, ByVal CellAddress As String, ByVal Value As Variant) As Boolean
    Dim wb As Workbook
    Dim ws As Worksheet

    On Error GoTo Error_Handler

    Set wb = SafeOpenWorkbook(FilePath)
    If wb Is Nothing Then
        SetValue = False
        Exit Function
    End If

    Set ws = wb.Worksheets(SheetName)
    ws.Range(CellAddress).Value = Value

    wb.Save
    SafeCloseWorkbook wb

    SetValue = True
    Exit Function

Error_Handler:
    If Not wb Is Nothing Then SafeCloseWorkbook wb, False
    CoreFramework.HandleStandardErrors Err.Number, "SetValue", "DataManager"
    SetValue = False
End Function

' **Purpose**: Get entire row data from Excel file
' **Parameters**:
'   - FilePath (String): Full path to Excel file
'   - SheetName (String): Name of worksheet
'   - RowNumber (Long): Row number to retrieve
' **Returns**: Variant - Array of row values, empty array if error
' **Dependencies**: SafeOpenWorkbook(), SafeCloseWorkbook()
' **Side Effects**: Opens and closes workbook
' **Errors**: Returns empty array if file access or row retrieval fails
Public Function GetRowData(ByVal FilePath As String, ByVal SheetName As String, ByVal RowNumber As Long) As Variant
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim LastCol As Long
    Dim RowData As Variant

    On Error GoTo Error_Handler

    Set wb = SafeOpenWorkbook(FilePath)
    If wb Is Nothing Then
        GetRowData = Array()
        Exit Function
    End If

    Set ws = wb.Worksheets(SheetName)
    LastCol = ws.Cells(RowNumber, ws.Columns.Count).End(xlToLeft).Column

    If LastCol > 0 Then
        RowData = ws.Range(ws.Cells(RowNumber, 1), ws.Cells(RowNumber, LastCol)).Value
    Else
        RowData = Array()
    End If

    SafeCloseWorkbook wb, False

    GetRowData = RowData
    Exit Function

Error_Handler:
    If Not wb Is Nothing Then SafeCloseWorkbook wb, False
    CoreFramework.HandleStandardErrors Err.Number, "GetRowData", "DataManager"
    GetRowData = Array()
End Function

' **Purpose**: Get entire column data from Excel file
' **Parameters**:
'   - FilePath (String): Full path to Excel file
'   - SheetName (String): Name of worksheet
'   - ColumnNumber (Long): Column number to retrieve
' **Returns**: Variant - Array of column values, empty array if error
' **Dependencies**: SafeOpenWorkbook(), SafeCloseWorkbook()
' **Side Effects**: Opens and closes workbook
' **Errors**: Returns empty array if file access or column retrieval fails
Public Function GetColumnData(ByVal FilePath As String, ByVal SheetName As String, ByVal ColumnNumber As Long) As Variant
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim ColumnData As Variant

    On Error GoTo Error_Handler

    Set wb = SafeOpenWorkbook(FilePath)
    If wb Is Nothing Then
        GetColumnData = Array()
        Exit Function
    End If

    Set ws = wb.Worksheets(SheetName)
    LastRow = ws.Cells(ws.Rows.Count, ColumnNumber).End(xlUp).Row

    If LastRow > 0 Then
        ColumnData = ws.Range(ws.Cells(1, ColumnNumber), ws.Cells(LastRow, ColumnNumber)).Value
    Else
        ColumnData = Array()
    End If

    SafeCloseWorkbook wb, False

    GetColumnData = ColumnData
    Exit Function

Error_Handler:
    If Not wb Is Nothing Then SafeCloseWorkbook wb, False
    CoreFramework.HandleStandardErrors Err.Number, "GetColumnData", "DataManager"
    GetColumnData = Array()
End Function

' **Purpose**: Get range data from Excel file
' **Parameters**:
'   - FilePath (String): Full path to Excel file
'   - SheetName (String): Name of worksheet
'   - RangeAddress (String): Range address (e.g., "A1:C10")
' **Returns**: Variant - Array of range values, empty array if error
' **Dependencies**: SafeOpenWorkbook(), SafeCloseWorkbook()
' **Side Effects**: Opens and closes workbook
' **Errors**: Returns empty array if file access or range retrieval fails
Public Function GetRangeData(ByVal FilePath As String, ByVal SheetName As String, ByVal RangeAddress As String) As Variant
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim RangeData As Variant

    On Error GoTo Error_Handler

    Set wb = SafeOpenWorkbook(FilePath)
    If wb Is Nothing Then
        GetRangeData = Array()
        Exit Function
    End If

    Set ws = wb.Worksheets(SheetName)
    RangeData = ws.Range(RangeAddress).Value

    SafeCloseWorkbook wb, False

    GetRangeData = RangeData
    Exit Function

Error_Handler:
    If Not wb Is Nothing Then SafeCloseWorkbook wb, False
    CoreFramework.HandleStandardErrors Err.Number, "GetRangeData", "DataManager"
    GetRangeData = Array()
End Function

' **Purpose**: Find value in worksheet and return row number
' **Parameters**:
'   - FilePath (String): Full path to Excel file
'   - SheetName (String): Name of worksheet
'   - SearchValue (Variant): Value to search for
'   - SearchColumn (Long, Optional): Column to search in (default 1)
' **Returns**: Long - Row number if found, 0 if not found
' **Dependencies**: SafeOpenWorkbook(), SafeCloseWorkbook()
' **Side Effects**: Opens and closes workbook
' **Errors**: Returns 0 if file access fails or value not found
Public Function FindValue(ByVal FilePath As String, ByVal SheetName As String, ByVal SearchValue As Variant, Optional ByVal SearchColumn As Long = 1) As Long
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim FoundCell As Range

    On Error GoTo Error_Handler

    Set wb = SafeOpenWorkbook(FilePath)
    If wb Is Nothing Then
        FindValue = 0
        Exit Function
    End If

    Set ws = wb.Worksheets(SheetName)
    Set FoundCell = ws.Columns(SearchColumn).Find(SearchValue, LookIn:=xlValues, LookAt:=xlWhole)

    If Not FoundCell Is Nothing Then
        FindValue = FoundCell.Row
    Else
        FindValue = 0
    End If

    SafeCloseWorkbook wb, False

    Exit Function

Error_Handler:
    If Not wb Is Nothing Then SafeCloseWorkbook wb, False
    CoreFramework.HandleStandardErrors Err.Number, "FindValue", "DataManager"
    FindValue = 0
End Function

' **Purpose**: Update Excel data with error handling and validation
' **Parameters**:
'   - FilePath (String): Full path to Excel file
'   - SheetName (String): Name of worksheet
'   - Updates (Variant): Array of updates to apply
' **Returns**: Boolean - True if all updates successful, False if any failed
' **Dependencies**: SafeOpenWorkbook(), SafeCloseWorkbook()
' **Side Effects**: Opens workbook, applies updates, saves and closes workbook
' **Errors**: Returns False if file access or update operations fail
Public Function UpdateExcelData(ByVal FilePath As String, ByVal SheetName As String, ByVal Updates As Variant) As Boolean
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim i As Integer

    On Error GoTo Error_Handler

    Set wb = SafeOpenWorkbook(FilePath)
    If wb Is Nothing Then
        UpdateExcelData = False
        Exit Function
    End If

    Set ws = wb.Worksheets(SheetName)

    ' Apply updates (assumes Updates is array of arrays: [CellAddress, Value])
    For i = LBound(Updates) To UBound(Updates)
        ws.Range(Updates(i)(0)).Value = Updates(i)(1)
    Next i

    wb.Save
    SafeCloseWorkbook wb

    UpdateExcelData = True
    Exit Function

Error_Handler:
    If Not wb Is Nothing Then SafeCloseWorkbook wb, False
    CoreFramework.HandleStandardErrors Err.Number, "UpdateExcelData", "DataManager"
    UpdateExcelData = False
End Function

' ===================================================================
' NUMBER GENERATION OPERATIONS
' ===================================================================

' **Purpose**: Get next enquiry number in sequence
' **Parameters**: None
' **Returns**: String - Next enquiry number (E00001 format), empty if error
' **Dependencies**: GetNextNumber()
' **Side Effects**: Updates number tracking file
' **Errors**: Returns empty string if number generation fails
Public Function GetNextEnquiryNumber() As String
    GetNextEnquiryNumber = GetNextNumber("E")
End Function

' **Purpose**: Get next quote number in sequence
' **Parameters**: None
' **Returns**: String - Next quote number (Q00001 format), empty if error
' **Dependencies**: GetNextNumber()
' **Side Effects**: Updates number tracking file
' **Errors**: Returns empty string if number generation fails
Public Function GetNextQuoteNumber() As String
    GetNextQuoteNumber = GetNextNumber("Q")
End Function

' **Purpose**: Get next job number in sequence
' **Parameters**: None
' **Returns**: String - Next job number (J00001 format), empty if error
' **Dependencies**: GetNextNumber()
' **Side Effects**: Updates number tracking file
' **Errors**: Returns empty string if number generation fails
Public Function GetNextJobNumber() As String
    GetNextJobNumber = GetNextNumber("J")
End Function

' **Purpose**: Get next number in sequence for specified prefix
' **Parameters**:
'   - Prefix (String): Number prefix (E, Q, J)
' **Returns**: String - Next number with prefix, empty if error
' **Dependencies**: SafeOpenWorkbook(), GetLastNumberFromSheet(), UpdateNumberInSheet()
' **Side Effects**: Creates number tracking file if missing, updates number sequence
' **Errors**: Returns empty string if file access or number generation fails
Private Function GetNextNumber(ByVal Prefix As String) As String
    Dim NumbersWB As Workbook
    Dim NumbersWS As Worksheet
    Dim LastNumber As Long
    Dim NextNumber As Long
    Dim NumbersFile As String

    On Error GoTo Error_Handler

    NumbersFile = GetRootPath & "\" & NUMBERS_FILE

    If Not FileExists(NumbersFile) Then
        CreateNumbersFile NumbersFile
    End If

    Set NumbersWB = SafeOpenWorkbook(NumbersFile)
    If NumbersWB Is Nothing Then
        GetNextNumber = ""
        Exit Function
    End If

    Set NumbersWS = NumbersWB.Worksheets(1)

    LastNumber = GetLastNumberFromSheet(NumbersWS, Prefix)
    NextNumber = LastNumber + 1

    UpdateNumberInSheet NumbersWS, Prefix, NextNumber

    NumbersWB.Save
    SafeCloseWorkbook NumbersWB

    GetNextNumber = Prefix & Format(NextNumber, "00000")
    Exit Function

Error_Handler:
    If Not NumbersWB Is Nothing Then SafeCloseWorkbook NumbersWB, False
    CoreFramework.HandleStandardErrors Err.Number, "GetNextNumber", "DataManager"
    GetNextNumber = ""
End Function

' **Purpose**: Get last used number for prefix from worksheet
' **Parameters**:
'   - ws (Worksheet): Number tracking worksheet
'   - Prefix (String): Number prefix to find
' **Returns**: Long - Last number used, 0 if not found
' **Dependencies**: None
' **Side Effects**: None
' **Errors**: Returns 0 if error or prefix not found
Private Function GetLastNumberFromSheet(ByVal ws As Worksheet, ByVal Prefix As String) As Long
    Dim i As Long

    On Error GoTo Error_Handler

    For i = 1 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If ws.Cells(i, 1).Value = Prefix Then
            GetLastNumberFromSheet = ws.Cells(i, 2).Value
            Exit Function
        End If
    Next i

    GetLastNumberFromSheet = 0
    Exit Function

Error_Handler:
    GetLastNumberFromSheet = 0
End Function

' **Purpose**: Update number tracking worksheet with new number
' **Parameters**:
'   - ws (Worksheet): Number tracking worksheet
'   - Prefix (String): Number prefix to update
'   - Number (Long): New number value
' **Returns**: None (Subroutine)
' **Dependencies**: None
' **Side Effects**: Updates worksheet cells with new number and timestamp
' **Errors**: Logs errors if update fails
Private Sub UpdateNumberInSheet(ByVal ws As Worksheet, ByVal Prefix As String, ByVal Number As Long)
    Dim i As Long
    Dim Found As Boolean

    On Error GoTo Error_Handler

    For i = 1 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If ws.Cells(i, 1).Value = Prefix Then
            ws.Cells(i, 2).Value = Number
            ws.Cells(i, 3).Value = Now
            Found = True
            Exit For
        End If
    Next i

    If Not Found Then
        i = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        ws.Cells(i, 1).Value = Prefix
        ws.Cells(i, 2).Value = Number
        ws.Cells(i, 3).Value = Now
    End If

    Exit Sub

Error_Handler:
    CoreFramework.LogError Err.Number, Err.Description, "UpdateNumberInSheet", "DataManager"
End Sub

' **Purpose**: Create number tracking file with initial structure
' **Parameters**:
'   - FilePath (String): Full path for new number tracking file
' **Returns**: None (Subroutine)
' **Dependencies**: Excel Workbooks.Add
' **Side Effects**: Creates new Excel file with number tracking structure
' **Errors**: Logs errors if file creation fails
Private Sub CreateNumbersFile(ByVal FilePath As String)
    Dim NewWB As Workbook
    Dim NewWS As Worksheet

    On Error GoTo Error_Handler

    Set NewWB = Workbooks.Add
    Set NewWS = NewWB.Worksheets(1)

    With NewWS
        .Name = "NumberTracking"
        .Cells(1, 1).Value = "Prefix"
        .Cells(1, 2).Value = "Last Number"
        .Cells(1, 3).Value = "Last Updated"

        .Cells(2, 1).Value = "E"
        .Cells(2, 2).Value = 0
        .Cells(2, 3).Value = Now

        .Cells(3, 1).Value = "Q"
        .Cells(3, 2).Value = 0
        .Cells(3, 3).Value = Now

        .Cells(4, 1).Value = "J"
        .Cells(4, 2).Value = 0
        .Cells(4, 3).Value = Now

        .Range("A1:C1").Font.Bold = True
        .Columns("A:C").AutoFit
    End With

    NewWB.SaveAs FilePath
    NewWB.Close
    Set NewWB = Nothing

    Exit Sub

Error_Handler:
    If Not NewWB Is Nothing Then
        NewWB.Close SaveChanges:=False
        Set NewWB = Nothing
    End If
    CoreFramework.HandleStandardErrors Err.Number, "CreateNumbersFile", "DataManager"
End Sub

' **Purpose**: Validate number format and prefix
' **Parameters**:
'   - Number (String): Number to validate
'   - ExpectedPrefix (String): Expected prefix character
' **Returns**: Boolean - True if valid format, False if invalid
' **Dependencies**: None
' **Side Effects**: None
' **Errors**: Returns False if validation fails
Public Function ValidateNumber(ByVal Number As String, ByVal ExpectedPrefix As String) As Boolean
    If Len(Number) < 6 Then
        ValidateNumber = False
        Exit Function
    End If

    If Left(Number, 1) <> ExpectedPrefix Then
        ValidateNumber = False
        Exit Function
    End If

    If Not IsNumeric(Mid(Number, 2)) Then
        ValidateNumber = False
        Exit Function
    End If

    ValidateNumber = True
End Function

' **Purpose**: Reserve next number without committing to use
' **Parameters**:
'   - Prefix (String): Number prefix (E, Q, J)
' **Returns**: String - Reserved number, empty if error
' **Dependencies**: GetNextNumber()
' **Side Effects**: Increments number sequence
' **Errors**: Returns empty string if reservation fails
Public Function ReserveNumber(ByVal Prefix As String) As String
    ReserveNumber = GetNextNumber(Prefix)
End Function

' **Purpose**: Confirm usage of previously reserved number
' **Parameters**:
'   - Number (String): Number to confirm usage
' **Returns**: Boolean - Always True (placeholder for future implementation)
' **Dependencies**: None
' **Side Effects**: None
' **Errors**: None (always succeeds)
Public Function ConfirmNumberUsage(ByVal Number As String) As Boolean
    ConfirmNumberUsage = True
End Function

' ===================================================================
' FORM DATA PERSISTENCE (CLAUDE.md: Replaces SaveFileCode.bas)
' ===================================================================

' **Purpose**: Save form data to worksheet by matching control names to cells
' **Parameters**:
'   - FormObject (Object): Form containing controls to save
'   - wb (Workbook): Target workbook for saving
'   - SheetName (String): Target worksheet name
' **Returns**: Boolean - True if save successful, False if failed
' **Dependencies**: None
' **Side Effects**: Updates worksheet cells with form control values
' **Errors**: Returns False if save operation fails
' **CLAUDE.md Compliance**: Replaces legacy SaveFileCode.bas SaveToColumns functionality
Public Function SaveFormToWorksheet(ByRef FormObject As Object, ByRef wb As Workbook, ByVal SheetName As String) As Boolean
    Dim ws As Worksheet
    Dim ctl As Object
    Dim i As Integer

    On Error GoTo Error_Handler

    Set ws = wb.Worksheets(SheetName)

    ' Iterate through form controls and save values
    For Each ctl In FormObject.Controls
        For i = 0 To 100
            If UCase(ws.Range("A1").Offset(i, 0).Value) = UCase(ctl.Name) Then
                Select Case UCase(TypeName(ctl))
                    Case "TEXTBOX"
                        ws.Range("A1").Offset(i, 1).Value = ctl.Value
                    Case "LABEL"
                        ws.Range("A1").Offset(i, 1).Value = ctl.Caption
                    Case "COMBOBOX"
                        ws.Range("A1").Offset(i, 1).Value = ctl.Value
                End Select
                Exit For
            End If
            If ws.Range("A1").Offset(i, 0).Value = "" Then Exit For
        Next i
    Next ctl

    SaveFormToWorksheet = True
    Exit Function

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "SaveFormToWorksheet", "DataManager"
    SaveFormToWorksheet = False
End Function

' **Purpose**: Load form data from worksheet by matching cell names to controls
' **Parameters**:
'   - FormObject (Object): Form containing controls to load
'   - wb (Workbook): Source workbook for loading
'   - SheetName (String): Source worksheet name
' **Returns**: Boolean - True if load successful, False if failed
' **Dependencies**: None
' **Side Effects**: Updates form control values with worksheet data
' **Errors**: Returns False if load operation fails
Public Function LoadFormFromWorksheet(ByRef FormObject As Object, ByRef wb As Workbook, ByVal SheetName As String) As Boolean
    Dim ws As Worksheet
    Dim ctl As Object
    Dim i As Integer
    Dim ControlName As String

    On Error GoTo Error_Handler

    Set ws = wb.Worksheets(SheetName)

    ' Iterate through form controls and load values
    For Each ctl In FormObject.Controls
        ControlName = UCase(ctl.Name)

        For i = 0 To 100
            If UCase(ws.Range("A1").Offset(i, 0).Value) = ControlName Then
                Select Case UCase(TypeName(ctl))
                    Case "TEXTBOX"
                        ctl.Value = ws.Range("A1").Offset(i, 1).Value
                    Case "LABEL"
                        ctl.Caption = ws.Range("A1").Offset(i, 1).Value
                    Case "COMBOBOX"
                        ctl.Value = ws.Range("A1").Offset(i, 1).Value
                End Select
                Exit For
            End If
            If ws.Range("A1").Offset(i, 0).Value = "" Then Exit For
        Next i
    Next ctl

    LoadFormFromWorksheet = True
    Exit Function

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "LoadFormFromWorksheet", "DataManager"
    LoadFormFromWorksheet = False
End Function

' **Purpose**: Save form data to Admin worksheet with standardized structure
' **Parameters**:
'   - FormObject (Object): Form containing controls to save
'   - wb (Workbook): Target workbook for saving
' **Returns**: Boolean - True if save successful, False if failed
' **Dependencies**: SaveFormToWorksheet()
' **Side Effects**: Updates ADMIN worksheet with form data
' **Errors**: Returns False if save operation fails
Public Function SaveFormToAdmin(ByRef FormObject As Object, ByRef wb As Workbook) As Boolean
    SaveFormToAdmin = SaveFormToWorksheet(FormObject, wb, "ADMIN")
End Function

' **Purpose**: Update picture in worksheet from form control
' **Parameters**:
'   - FormObject (Object): Form containing picture path control
'   - wb (Workbook): Target workbook for picture update
'   - SheetName (String): Target worksheet name
'   - PictureControlName (String): Name of control containing picture path
' **Returns**: Boolean - True if update successful, False if failed
' **Dependencies**: GetRootPath()
' **Side Effects**: Inserts or updates picture in worksheet
' **Errors**: Returns False if picture insertion fails
' **CLAUDE.md Compliance**: Enhanced version of legacy picture handling
Public Function UpdatePictureInWorksheet(ByRef FormObject As Object, ByRef wb As Workbook, ByVal SheetName As String, ByVal PictureControlName As String) As Boolean
    Dim ws As Worksheet
    Dim PictureControl As Object
    Dim PicturePath As String
    Dim DrawingRange As Range

    On Error GoTo Error_Handler

    Set ws = wb.Worksheets(SheetName)
    Set PictureControl = FormObject.Controls(PictureControlName)

    If PictureControl.Value <> "" Then
        PicturePath = GetRootPath & "\images\" & PictureControl.Value

        If FileExists(PicturePath) Then
            ' Find drawing location range
            Set DrawingRange = ws.Range("Drawing_location")

            ' Remove existing picture if present
            On Error Resume Next
            ws.Shapes("Drawing").Delete
            On Error GoTo Error_Handler

            ' Insert new picture
            With ws.Pictures.Insert(PicturePath)
                .Name = "Drawing"
                .PrintObject = True
                .Height = DrawingRange.RowHeight * 10
                .Left = DrawingRange.Left + 5
                .Top = DrawingRange.Top + 5
            End With
        End If
    End If

    UpdatePictureInWorksheet = True
    Exit Function

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "UpdatePictureInWorksheet", "DataManager"
    UpdatePictureInWorksheet = False
End Function

' ===================================================================
' UTILITY FUNCTIONS
' ===================================================================

' **Purpose**: Generate next filename with counter in specified directory
' **Parameters**:
'   - DirectoryName (String): Name of subdirectory under root path
'   - Prefix (String): Filename prefix
'   - Extension (String): File extension (including dot)
' **Returns**: String - Next available filename
' **Dependencies**: GetRootPath(), FileExists()
' **Side Effects**: None
' **Errors**: Returns generic filename if error occurs
Public Function GetNextFileName(ByVal DirectoryName As String, ByVal Prefix As String, ByVal Extension As String) As String
    Dim DirPath As String
    Dim Counter As Integer
    Dim FileName As String

    On Error GoTo Error_Handler

    DirPath = GetRootPath & "\" & DirectoryName & "\"
    Counter = 1

    Do
        FileName = Prefix & Format(Counter, "0000") & Extension
        Counter = Counter + 1
    Loop While FileExists(DirPath & FileName)

    GetNextFileName = FileName
    Exit Function

Error_Handler:
    GetNextFileName = Prefix & "0001" & Extension
End Function

' **Purpose**: Format currency value for display
' **Parameters**:
'   - Amount (Currency): Currency amount to format
' **Returns**: String - Formatted currency string
' **Dependencies**: None
' **Side Effects**: None
' **Errors**: Returns "£0.00" if formatting fails
Public Function FormatCurrency(ByVal Amount As Currency) As String
    On Error GoTo Error_Handler
    FormatCurrency = Format(Amount, "£#,##0.00")
    Exit Function

Error_Handler:
    FormatCurrency = "£0.00"
End Function

' **Purpose**: Format date value for display
' **Parameters**:
'   - DateValue (Date): Date to format
' **Returns**: String - Formatted date string
' **Dependencies**: None
' **Side Effects**: None
' **Errors**: Returns current date if formatting fails
Public Function FormatDate(ByVal DateValue As Date) As String
    On Error GoTo Error_Handler
    FormatDate = Format(DateValue, "dd/mm/yyyy")
    Exit Function

Error_Handler:
    FormatDate = Format(Now, "dd/mm/yyyy")
End Function

' ===================================================================
' SYSTEM INITIALIZATION FUNCTIONS
' ===================================================================

' **Purpose**: Initialize number tracking database with proper structure
' **Parameters**: None
' **Returns**: Boolean - True if initialization successful, False if error
' **Dependencies**: CreateNumbersFile, GetRootPath
' **Side Effects**: Creates Templates\number_tracking.xls file
' **Errors**: Returns False on file creation failure, logs error
Public Function InitializeNumberTracking() As Boolean
    Dim FilePath As String

    On Error GoTo Error_Handler

    FilePath = GetRootPath & "\" & NUMBERS_FILE

    ' Use existing CreateNumbersFile function
    CreateNumbersFile FilePath

    ' Verify file was created
    If FileExists(FilePath) Then
        InitializeNumberTracking = True
        CoreFramework.LogError 0, "Number tracking database initialized successfully", "InitializeNumberTracking", "DataManager"
    Else
        InitializeNumberTracking = False
        CoreFramework.LogError 0, "Failed to create number tracking database", "InitializeNumberTracking", "DataManager"
    End If

    Exit Function

Error_Handler:
    CoreFramework.LogError Err.Number, Err.Description, "InitializeNumberTracking", "DataManager"
    InitializeNumberTracking = False
End Function