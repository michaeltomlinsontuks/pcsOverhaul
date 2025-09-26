Attribute VB_Name = "DataOperations"
' **Purpose**: All file operations, Excel data access, and directory management
' **CLAUDE.md Compliance**: Maintains all directory structure requirements, 32/64-bit compatibility
' **Consolidation**: Combines DataManager.bas and DataUtilities.bas
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
    Dim BasePath As String

    On Error GoTo Error_Handler

    ' First try to get path from Main form if available
    On Error Resume Next
    If Main.Main_MasterPath.Value <> "" Then
        BasePath = Main.Main_MasterPath.Value
        If Right(BasePath, 1) = "\" Then
            BasePath = Left(BasePath, Len(BasePath) - 1)
        End If
        GetRootPath = BasePath
        Exit Function
    End If
    On Error GoTo Error_Handler

    ' Fall back to ROOT_PATH constant or ThisWorkbook.Path
    If ROOT_PATH = "" Then
        BasePath = ThisWorkbook.Path
    Else
        BasePath = ROOT_PATH
    End If

    ' Remove trailing backslash for consistency
    If Right(BasePath, 1) = "\" Then
        BasePath = Left(BasePath, Len(BasePath) - 1)
    End If

    GetRootPath = BasePath

    Exit Function

Error_Handler:
    SystemCore.LogError Err.Number, Err.Description, "GetRootPath", "DataOperations"
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
            SystemCore.LogError 0, "Missing directory: " & RequiredDirs(i), "ValidateDirectoryStructure", "DataOperations"
            Exit Function
        End If
    Next i

    ValidateDirectoryStructure = True
    Exit Function

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "ValidateDirectoryStructure", "DataOperations"
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
            SystemCore.LogError 0, "Created missing directory: " & RequiredDirs(i), "CreateDirectoryStructure", "DataOperations"
        End If
    Next i

    CreateDirectoryStructure = True
    Exit Function

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "CreateDirectoryStructure", "DataOperations"
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

' **Purpose**: Check if directory exists (alias for DirExists for compatibility)
' **Parameters**:
'   - DirPath (String): Full path to directory to check
' **Returns**: Boolean - True if directory exists, False if not
' **Dependencies**: DirExists function
' **Side Effects**: None
' **Errors**: Returns False if error occurs during check
Public Function DirectoryExists(ByVal DirPath As String) As Boolean
    DirectoryExists = DirExists(DirPath)
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
        SystemCore.LogError SystemCore.ERR_PATH_NOT_FOUND, "Directory not found: " & DirPath, "GetFileList", "DataOperations"
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
    SystemCore.HandleStandardErrors Err.Number, "GetFileList", "DataOperations"
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
    SystemCore.HandleStandardErrors Err.Number, "GetFileListWithStatus", "DataOperations"
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
    SystemCore.HandleStandardErrors Err.Number, "CreateBackup", "DataOperations"
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
' **Dependencies**: FileExists(), SystemCore.ErrorHandler for error logging
' **Side Effects**: Opens workbook in Excel application, logs errors if failed
' **Errors**: Returns Nothing on file not found, permission denied, or corruption
' **CLAUDE.md Compliance**: Maintains 32/64-bit Excel compatibility
Public Function SafeOpenWorkbook(ByVal FilePath As String, Optional ByVal ReadOnlyFlag As Boolean = False) As Workbook
    Dim wb As Workbook

    On Error GoTo Error_Handler

    If Not FileExists(FilePath) Then
        SystemCore.LogError SystemCore.ERR_FILE_NOT_FOUND, "File not found: " & FilePath, "SafeOpenWorkbook", "DataOperations"
        Set SafeOpenWorkbook = Nothing
        Exit Function
    End If

    ' Suppress Excel prompts, alerts, and screen updating during file opening
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    Application.ScreenUpdating = False

    Set wb = Workbooks.Open(FilePath, ReadOnly:=ReadOnlyFlag, UpdateLinks:=0)

    ' Restore alerts and screen updating
    Application.DisplayAlerts = True
    Application.AskToUpdateLinks = True
    Application.ScreenUpdating = True

    Set SafeOpenWorkbook = wb
    Exit Function

Error_Handler:
    ' Restore alerts and screen updating even on error
    Application.DisplayAlerts = True
    Application.AskToUpdateLinks = True
    Application.ScreenUpdating = True

    SystemCore.HandleStandardErrors Err.Number, "SafeOpenWorkbook", "DataOperations"
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
        ' Suppress screen updating and alerts during close
        Application.DisplayAlerts = False
        Application.ScreenUpdating = False

        wb.Close SaveChanges:=SaveChanges
        Set wb = Nothing

        ' Restore alerts and screen updating
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True

        SafeCloseWorkbook = True
    End If
    Exit Function

Error_Handler:
    ' Restore alerts and screen updating even on error
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    SystemCore.HandleStandardErrors Err.Number, "SafeCloseWorkbook", "DataOperations"
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
    SystemCore.HandleStandardErrors Err.Number, "CreateNewWorkbook", "DataOperations"
    Set CreateNewWorkbook = Nothing
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

    Set wb = SafeOpenWorkbook(FilePath, True)
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
    SystemCore.HandleStandardErrors Err.Number, "GetValue", "DataOperations"
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
    SystemCore.HandleStandardErrors Err.Number, "GetValueFromClosedWorkbook", "DataOperations"
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
    SystemCore.HandleStandardErrors Err.Number, "SetValue", "DataOperations"
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

    Set wb = SafeOpenWorkbook(FilePath, True)
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
    SystemCore.HandleStandardErrors Err.Number, "GetRowData", "DataOperations"
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

    Set wb = SafeOpenWorkbook(FilePath, True)
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
    SystemCore.HandleStandardErrors Err.Number, "GetColumnData", "DataOperations"
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

    Set wb = SafeOpenWorkbook(FilePath, True)
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
    SystemCore.HandleStandardErrors Err.Number, "GetRangeData", "DataOperations"
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

    Set wb = SafeOpenWorkbook(FilePath, True)
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
    SystemCore.HandleStandardErrors Err.Number, "FindValue", "DataOperations"
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
    SystemCore.HandleStandardErrors Err.Number, "UpdateExcelData", "DataOperations"
    UpdateExcelData = False
End Function

' ===================================================================
' TEMPLATE-BASED NUMBER GENERATION (CLAUDE.md: Calc_Numbers.bas replacement)
' ===================================================================

' **Purpose**: Calculate next number based on template files in Templates directory
' **Original**: Interface_VBA/Calc_Numbers.bas Calc_Next_Number()
' **Parameters**:
'   - Typ (String): Type of number to calculate ("E", "Q", "J")
' **Returns**: Variant - Next number in sequence, 0 if error
' **File Dependencies**: Templates directory with "E - ###.TXT", "Q - ###.TXT", "J - ###.TXT" files
' **Form Usage**: Used by enquiry, quote, and job forms for number generation
' **CLAUDE.md Compliance**: Exact replacement for Calc_Numbers.bas functionality
Public Function Calc_Next_Number(Typ As String) As Variant
    Dim TemplatesPath As String
    Dim MyName As String
    Dim MaxNumber As Long
    Dim CurrentNumber As Long
    Dim PrefixPattern As String
    Dim NumberStart As Integer
    Dim NumberEnd As Integer

    On Error GoTo Error_Handler

    TemplatesPath = GetRootPath & "\Templates\"

    If Not DirExists(TemplatesPath) Then
        SystemCore.ShowError "Templates folder not found at: " & TemplatesPath, "Folder Not Found"
        Calc_Next_Number = 0
        Exit Function
    End If

    ' Set prefix pattern based on type
    Select Case UCase(Left(Typ, 1))
        Case "E"
            PrefixPattern = "E - "
        Case "Q"
            PrefixPattern = "Q - "
        Case "J"
            PrefixPattern = "J - "
        Case Else
            SystemCore.LogError 0, "Invalid type parameter: " & Typ, "Calc_Next_Number", "DataOperations"
            Calc_Next_Number = 0
            Exit Function
    End Select

    MaxNumber = 0
    MyName = Dir(TemplatesPath & "*", vbNormal)

    Do Until MyName = ""
        If Left(UCase(MyName), 4) = UCase(PrefixPattern) And Right(UCase(MyName), 4) = ".TXT" Then
            ' Extract number from filename "E - 123.TXT"
            NumberStart = InStr(1, MyName, "-", vbTextCompare) + 2
            NumberEnd = Len(MyName) - 4 ' Remove .TXT extension
            If NumberEnd > NumberStart Then
                CurrentNumber = CLng(Trim(Mid(MyName, NumberStart, NumberEnd - NumberStart + 1)))
                If CurrentNumber > MaxNumber Then
                    MaxNumber = CurrentNumber
                End If
            End If
        End If
        MyName = Dir
    Loop

    Calc_Next_Number = MaxNumber + 1
    Exit Function

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "Calc_Next_Number", "DataOperations"
    Calc_Next_Number = 0
End Function

' **Purpose**: Confirm and update template file with next number
' **Original**: Interface_VBA/Calc_Numbers.bas Confirm_Next_Number()
' **Parameters**:
'   - Typ (String): Type of number to confirm ("E", "Q", "J")
' **Returns**: Variant - Confirmed next number, 0 if error
' **File Dependencies**: Templates directory with template files for number tracking
' **Form Usage**: Used after Calc_Next_Number to commit the number and update template file
' **CLAUDE.md Compliance**: Exact replacement for Calc_Numbers.bas functionality with FileCopy+Kill logic
Public Function Confirm_Next_Number(Typ As String) As Variant
    Dim TemplatesPath As String
    Dim MyName As String
    Dim MaxNumber As Long
    Dim CurrentNumber As Long
    Dim PrefixPattern As String
    Dim OldFilePath As String
    Dim NewFilePath As String
    Dim NextNumber As Long

    On Error GoTo Error_Handler

    TemplatesPath = GetRootPath & "\Templates\"

    If Not DirExists(TemplatesPath) Then
        SystemCore.ShowError "Templates folder not found at: " & TemplatesPath, "Folder Not Found"
        Confirm_Next_Number = 0
        Exit Function
    End If

    ' Set prefix pattern based on type
    Select Case UCase(Left(Typ, 1))
        Case "E"
            PrefixPattern = "E - "
        Case "Q"
            PrefixPattern = "Q - "
        Case "J"
            PrefixPattern = "J - "
        Case Else
            SystemCore.LogError 0, "Invalid type parameter: " & Typ, "Confirm_Next_Number", "DataOperations"
            Confirm_Next_Number = 0
            Exit Function
    End Select

    MaxNumber = 0
    MyName = Dir(TemplatesPath, vbDirectory)

    ' Find the current highest number and the file to update
    Do Until MyName = ""
        If MyName <> "." And MyName <> ".." Then
            If Left(MyName, 4) = PrefixPattern Then
                ' Extract number from filename "E - 123.TXT"
                CurrentNumber = CLng(Mid(MyName, InStr(1, MyName, "-", vbTextCompare) + 2, Len(MyName) - 8))
                If CurrentNumber > MaxNumber Then
                    MaxNumber = CurrentNumber
                    OldFilePath = TemplatesPath & MyName
                End If
            End If
        End If
        MyName = Dir
    Loop

    NextNumber = MaxNumber + 1
    NewFilePath = TemplatesPath & PrefixPattern & NextNumber & ".TXT"

    ' Update the template file using FileCopy + Kill (exact legacy behavior)
    If OldFilePath <> "" And FileExists(OldFilePath) Then
        FileCopy OldFilePath, NewFilePath
        Kill OldFilePath
    Else
        ' Create new template file if none exists
        Dim FileNum As Integer
        FileNum = FreeFile
        Open NewFilePath For Output As FileNum
        Print #FileNum, "Template file for " & PrefixPattern & "number tracking"
        Close FileNum
    End If

    Confirm_Next_Number = NextNumber
    Exit Function

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "Confirm_Next_Number", "DataOperations"
    Confirm_Next_Number = 0
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
    SystemCore.HandleStandardErrors Err.Number, "GetNextNumber", "DataOperations"
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
    SystemCore.LogError Err.Number, Err.Description, "UpdateNumberInSheet", "DataOperations"
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
    SystemCore.HandleStandardErrors Err.Number, "CreateNumbersFile", "DataOperations"
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
    SystemCore.HandleStandardErrors Err.Number, "SaveFormToWorksheet", "DataOperations"
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
    SystemCore.HandleStandardErrors Err.Number, "LoadFormFromWorksheet", "DataOperations"
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
    SystemCore.HandleStandardErrors Err.Number, "UpdatePictureInWorksheet", "DataOperations"
    UpdatePictureInWorksheet = False
End Function

' ===================================================================
' WIP DATABASE INTEGRATION (CLAUDE.md: SaveWIPCode.bas replacement)
' ===================================================================

' **Purpose**: Save form information into WIP database
' **Original**: Interface_VBA/SaveWIPCode.bas SaveInfoIntoWIP()
' **Parameters**:
'   - FormObject (Object): Form containing data to save
' **Returns**: Boolean - True if successful, False if failed
' **Dependencies**: GetRootPath(), SafeOpenWorkbook(), OpenBook() for compatibility
' **Side Effects**: Opens WIP.xls, saves form data, closes workbook
' **Errors**: Returns False if WIP.xls cannot be opened or saved
' **CLAUDE.md Compliance**: Exact replacement for SaveWIPCode.bas functionality
Public Function SaveInfoIntoWIP(ByRef FormObject As Object) As Boolean
    Dim WipPath As String
    Dim WipWB As Workbook
    Dim WipWS As Worksheet
    Dim LastRow As Long
    Dim TargetRow As Long
    Dim ctl As Object
    Dim i As Integer
    Dim MatchFound As Boolean

    On Error GoTo Error_Handler

    WipPath = GetRootPath & "\WIP.xls"

    ' Check if WIP.xls exists
    If Not FileExists(WipPath) Then
        CreateWIPDatabase WipPath
    End If

    ' Open WIP file with read-only check loop (exact legacy behavior)
    Set WipWB = SafeOpenWorkbook(WipPath)
    If WipWB Is Nothing Then
        SaveInfoIntoWIP = False
        Exit Function
    End If

    Do While WipWB.ReadOnly
        WipWB.Close
        MsgBox "This workbook is read only, please find the user with this workbook open and close it."
        Set WipWB = SafeOpenWorkbook(WipPath)
        If WipWB Is Nothing Then
            SaveInfoIntoWIP = False
            Exit Function
        End If
    Loop

    Set WipWS = WipWB.Worksheets(1)

    ' Find the row to update (exact legacy logic)
    TargetRow = 2 ' Start from row 2 (row 1 is headers)
    MatchFound = False

    Do While WipWS.Cells(TargetRow, 1).Value <> ""
        ' Check for matching Quote_Number, Enquiry_Number, Job_Number, or File_Name
        On Error Resume Next
        If WipWS.Cells(TargetRow, 3).Value = FormObject.Quote_Number.Value Or _
           WipWS.Cells(TargetRow, 3).Value = FormObject.Enquiry_Number.Value Or _
           WipWS.Cells(TargetRow, 3).Value = FormObject.Job_Number.Value Or _
           WipWS.Cells(TargetRow, 3).Value = FormObject.File_Name.Value Then
            MatchFound = True
            Exit Do
        End If
        On Error GoTo Error_Handler
        TargetRow = TargetRow + 1
    Loop

    ' Clear the target row if match found, or use next empty row
    If MatchFound Then
        WipWS.Rows(TargetRow).ClearContents
    End If

    ' Save form controls to WIP (exact legacy algorithm)
    For Each ctl In FormObject.Controls
        For i = 0 To 100
            If UCase(WipWS.Cells(1, i + 1).Value) = UCase(ctl.Name) Then
                Select Case UCase(TypeName(ctl))
                    Case "LABEL"
                        WipWS.Cells(TargetRow, i + 1).Value = UCase(ctl.Caption)
                    Case "TEXTBOX"
                        WipWS.Cells(TargetRow, i + 1).Value = UCase(ctl.Value)
                    Case "COMBOBOX"
                        WipWS.Cells(TargetRow, i + 1).Value = UCase(ctl.Value)
                End Select
                Exit For
            End If
            ' Copy formula from row above if it starts with "=" (exact legacy logic)
            If Left(WipWS.Cells(TargetRow - 1, i + 1).Formula, 1) = "=" Then
                WipWS.Cells(TargetRow, i + 1).Formula = WipWS.Cells(TargetRow - 1, i + 1).Formula
            End If
            ' Break if we hit an empty header
            If UCase(WipWS.Cells(1, i + 2).Value) = "" Then Exit For
        Next i
    Next ctl

    WipWB.Save
    WipWB.Close
    Set WipWB = Nothing

    SaveInfoIntoWIP = True
    Exit Function

Error_Handler:
    If Not WipWB Is Nothing Then
        WipWB.Close SaveChanges:=False
        Set WipWB = Nothing
    End If
    SystemCore.HandleStandardErrors Err.Number, "SaveInfoIntoWIP", "DataOperations"
    SaveInfoIntoWIP = False
End Function

' **Purpose**: Create WIP database file if missing
' **Parameters**:
'   - FilePath (String): Path for new WIP database file
' **Returns**: Boolean - True if created successfully, False if failed
' **Dependencies**: CreateNewWorkbook()
' **Side Effects**: Creates new WIP.xls file with proper structure
' **Errors**: Returns False if file creation fails
Private Function CreateWIPDatabase(ByVal FilePath As String) As Boolean
    Dim WipWB As Workbook
    Dim WipWS As Worksheet

    On Error GoTo Error_Handler

    Set WipWB = CreateNewWorkbook()
    If WipWB Is Nothing Then
        CreateWIPDatabase = False
        Exit Function
    End If

    Set WipWS = WipWB.Worksheets(1)
    WipWS.Name = "WIP"

    ' Create basic header structure
    With WipWS
        .Cells(1, 1).Value = "ID"
        .Cells(1, 2).Value = "DATE"
        .Cells(1, 3).Value = "NUMBER"
        .Cells(1, 4).Value = "CUSTOMER"
        .Cells(1, 5).Value = "DESCRIPTION"
        .Cells(1, 6).Value = "STATUS"
        .Range("A1:F1").Font.Bold = True
    End With

    WipWB.SaveAs FilePath
    WipWB.Close
    Set WipWB = Nothing

    CreateWIPDatabase = True
    Exit Function

Error_Handler:
    If Not WipWB Is Nothing Then
        WipWB.Close SaveChanges:=False
        Set WipWB = Nothing
    End If
    SystemCore.HandleStandardErrors Err.Number, "CreateWIPDatabase", "DataOperations"
    CreateWIPDatabase = False
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

' **Purpose**: Format date with time for display consistency
' **Parameters**:
'   - DateValue (Date): Date to format
' **Returns**: String - Formatted date string with time
' **Dependencies**: None
' **Side Effects**: None
' **Errors**: Returns current date/time if formatting fails
Public Function FormatDateTime(ByVal DateValue As Date) As String
    On Error GoTo Error_Handler
    FormatDateTime = Format(DateValue, "dd/mm/yyyy hh:mm")
    Exit Function

Error_Handler:
    FormatDateTime = Format(Now, "dd/mm/yyyy hh:mm")
End Function

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
        SystemCore.LogError 0, "Number tracking database initialized successfully", "InitializeNumberTracking", "DataOperations"
    Else
        InitializeNumberTracking = False
        SystemCore.LogError 0, "Failed to create number tracking database", "InitializeNumberTracking", "DataOperations"
    End If

    Exit Function

Error_Handler:
    SystemCore.LogError Err.Number, Err.Description, "InitializeNumberTracking", "DataOperations"
    InitializeNumberTracking = False
End Function

' **Purpose**: Get values from a range in a closed workbook (for forms compatibility)
' **Parameters**:
'   - FilePath (String): Path to the workbook file
'   - SheetName (String): Name of the worksheet
'   - RangeAddress (String): Range address (e.g., "A:A", "A1:A10")
' **Returns**: Variant array of values or empty array if failed
' **Dependencies**: SafeOpenWorkbook, SafeCloseWorkbook
' **Side Effects**: None
' **Errors**: Returns empty array if file access fails
' **CLAUDE.md Compliance**: Provides compatibility function for forms
Public Function GetRangeValues(ByVal FilePath As String, ByVal SheetName As String, ByVal RangeAddress As String) As Variant
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim RangeData As Variant
    Dim ResultArray() As String
    Dim i As Long
    Dim ValidCount As Long

    On Error GoTo Error_Handler

    Set wb = SafeOpenWorkbook(FilePath, True)
    If wb Is Nothing Then
        GetRangeValues = Array()
        Exit Function
    End If

    Set ws = wb.Worksheets(SheetName)

    ' Get range data - handle both single column and specific ranges
    If RangeAddress = "A:A" Then
        ' Get column A data up to last used row
        Dim LastRow As Long
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If LastRow > 1 Then
            RangeData = ws.Range("A1:A" & LastRow).Value
        Else
            RangeData = Array()
        End If
    Else
        RangeData = ws.Range(RangeAddress).Value
    End If

    SafeCloseWorkbook wb, False

    ' Process the data into a clean array
    If IsArray(RangeData) Then
        ValidCount = 0
        ReDim ResultArray(0 To UBound(RangeData, 1) - 1)

        For i = LBound(RangeData, 1) To UBound(RangeData, 1)
            Dim CellValue As String
            If IsArray(RangeData) And UBound(RangeData, 2) >= 1 Then
                CellValue = Trim(CStr(RangeData(i, 1)))
            Else
                CellValue = Trim(CStr(RangeData(i)))
            End If

            If CellValue <> "" Then
                ResultArray(ValidCount) = CellValue
                ValidCount = ValidCount + 1
            End If
        Next i

        If ValidCount > 0 Then
            ReDim Preserve ResultArray(0 To ValidCount - 1)
            GetRangeValues = ResultArray
        Else
            GetRangeValues = Array()
        End If
    Else
        ' Single value case
        If Trim(CStr(RangeData)) <> "" Then
            GetRangeValues = Array(Trim(CStr(RangeData)))
        Else
            GetRangeValues = Array()
        End If
    End If
    Exit Function

Error_Handler:
    If Not wb Is Nothing Then SafeCloseWorkbook wb, False
    SystemCore.HandleStandardErrors Err.Number, "GetRangeValues", "DataOperations"
    GetRangeValues = Array()
End Function

' ===================================================================
' DATA UTILITIES (From DataUtilities.bas)
' ===================================================================

' **Purpose**: Get component codes from template file
' **Returns**: Variant - Array of component codes, empty array if failed
' **Dependencies**: GetRootPath(), GetRangeValues()
' **Side Effects**: None
' **Errors**: Returns empty array if template file not found
Public Function GetComponentCodes() As Variant
    Dim TemplatePath As String

    On Error GoTo Error_Handler

    TemplatePath = GetRootPath & "\Templates\Component_Grades.xls"
    If FileExists(TemplatePath) Then
        GetComponentCodes = GetRangeValues(TemplatePath, "Sheet1", "A:A")
    Else
        GetComponentCodes = Array()
    End If
    Exit Function

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "GetComponentCodes", "DataOperations"
    GetComponentCodes = Array()
End Function

' **Purpose**: Get material grades from template file
' **Returns**: Variant - Array of material grades, empty array if failed
' **Dependencies**: GetRootPath(), GetRangeValues()
' **Side Effects**: None
' **Errors**: Returns empty array if template file not found
Public Function GetMaterialGrades() As Variant
    Dim TemplatePath As String

    On Error GoTo Error_Handler

    TemplatePath = GetRootPath & "\Templates\Component_Grades.xls"
    If FileExists(TemplatePath) Then
        GetMaterialGrades = GetRangeValues(TemplatePath, "Sheet1", "A:A")
    Else
        GetMaterialGrades = Array()
    End If
    Exit Function

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "GetMaterialGrades", "DataOperations"
    GetMaterialGrades = Array()
End Function

' **Purpose**: Get customer list from customer files
' **Returns**: Variant - Array of customer names, empty array if failed
' **Dependencies**: GetRootPath(), GetRangeValues()
' **Side Effects**: None
' **Errors**: Returns empty array if customer file not found
Public Function GetCustomerList() As Variant
    Dim CustomerPath As String

    On Error GoTo Error_Handler

    ' Try multiple customer files that may exist
    CustomerPath = GetRootPath & "\Templates\Office_Customer.xls"
    If Not FileExists(CustomerPath) Then
        CustomerPath = GetRootPath & "\Templates\Workshop_Customer.xls"
    End If
    If Not FileExists(CustomerPath) Then
        CustomerPath = GetRootPath & "\Templates\_Client.xls"
    End If

    If FileExists(CustomerPath) Then
        GetCustomerList = GetRangeValues(CustomerPath, "Sheet1", "A:A")
    Else
        GetCustomerList = Array()
    End If
    Exit Function

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "GetCustomerList", "DataOperations"
    GetCustomerList = Array()
End Function

' **Purpose**: Get component price from price list
' **Parameters**:
'   - ComponentCode (String): Component code to look up
' **Returns**: Variant - Price value, 0 if not found
' **Dependencies**: GetRootPath(), LookupValue()
' **Side Effects**: None
' **Errors**: Returns 0 if price list not found or component not found
Public Function GetComponentPrice(ByVal ComponentCode As String) As Variant
    Dim PriceListPath As String

    On Error GoTo Error_Handler

    PriceListPath = GetRootPath & "\Templates\Price_List.xls"
    If FileExists(PriceListPath) Then
        GetComponentPrice = LookupValue(PriceListPath, ComponentCode, 1, 2)
    Else
        GetComponentPrice = 0
    End If
    Exit Function

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "GetComponentPrice", "DataOperations"
    GetComponentPrice = 0
End Function

' **Purpose**: Lookup value in any Excel table
' **Parameters**:
'   - TablePath (String): Full path to lookup table file
'   - SearchValue (Variant): Value to search for
'   - SearchColumn (Long, Optional): Column to search in (default 1)
'   - ReturnColumn (Long, Optional): Column to return value from (default 2)
' **Returns**: Variant - Found value, empty string if not found
' **Dependencies**: SafeOpenWorkbook for file access
' **Side Effects**: Opens and closes lookup table file
' **Errors**: Returns empty string if file access fails or value not found
Public Function LookupValue(ByVal TablePath As String, ByVal SearchValue As Variant, Optional ByVal SearchColumn As Long = 1, Optional ByVal ReturnColumn As Long = 2) As Variant
    Dim TableWB As Workbook
    Dim TableWS As Worksheet
    Dim LastRow As Long
    Dim i As Long
    Dim FoundValue As Variant

    On Error GoTo Error_Handler

    LookupValue = ""

    ' Validate inputs
    If Not FileExists(TablePath) Then Exit Function
    If SearchColumn < 1 Or ReturnColumn < 1 Then Exit Function

    Set TableWB = SafeOpenWorkbook(TablePath, True)
    If TableWB Is Nothing Then Exit Function

    Set TableWS = TableWB.Worksheets(1)
    LastRow = TableWS.Cells(TableWS.Rows.Count, SearchColumn).End(xlUp).Row

    ' Search for value
    For i = 2 To LastRow ' Skip header row
        If TableWS.Cells(i, SearchColumn).Value = SearchValue Then
            FoundValue = TableWS.Cells(i, ReturnColumn).Value
            Exit For
        End If
    Next i

    SafeCloseWorkbook TableWB, False
    LookupValue = FoundValue
    Exit Function

Error_Handler:
    If Not TableWB Is Nothing Then SafeCloseWorkbook TableWB, False
    SystemCore.HandleStandardErrors Err.Number, "LookupValue", "DataOperations"
    LookupValue = ""
End Function

' ===================================================================
' LEGACY COMPATIBILITY FUNCTIONS (CLAUDE.md: Exact legacy function signatures)
' ===================================================================

' **Purpose**: Open workbook with exact legacy compatibility (exact signature match)
' **Parameters**:
'   - File (String): Full path to Excel file to open
'   - RO (Boolean): ReadOnly flag - True for read-only, False for write access
' **Returns**: Nothing (matches legacy behavior)
' **Dependencies**: Excel Workbooks.Open
' **Side Effects**: Opens workbook in Excel application
' **Errors**: May raise Excel errors if file access fails
' **CLAUDE.md Compliance**: Exact replacement for legacy Open_Book.bas OpenBook functionality
Public Function OpenBook(File As String, RO As Boolean)
    On Error GoTo Error_Handler

    ' Suppress Excel prompts, alerts, and screen updating during file opening
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    Application.ScreenUpdating = False

    Workbooks.Open Filename:=File, ReadOnly:=RO, UpdateLinks:=0

    ' Restore alerts and screen updating
    Application.DisplayAlerts = True
    Application.AskToUpdateLinks = True
    Application.ScreenUpdating = True

    Exit Function

Error_Handler:
    ' Restore alerts and screen updating even on error
    Application.DisplayAlerts = True
    Application.AskToUpdateLinks = True
    Application.ScreenUpdating = True

    SystemCore.LogError Err.Number, Err.Description, "OpenBook", "DataOperations"
    ' Re-raise the error to maintain legacy behavior
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

' **Purpose**: Delete worksheet from active workbook (exact legacy compatibility)
' **Parameters**:
'   - SheetName (String): Name of worksheet to delete from active workbook
' **Returns**: Nothing (matches legacy behavior)
' **Dependencies**: Application.DisplayAlerts, ActiveWorkbook.Worksheets
' **Side Effects**: Deletes worksheet from active workbook without confirmation
' **Errors**: May raise Excel errors if worksheet not found
' **CLAUDE.md Compliance**: Exact replacement for legacy Delete_Sheet.bas DeleteSheet functionality
Public Function DeleteSheet(SheetName As String)
    On Error GoTo Error_Handler

    Application.DisplayAlerts = False
    Worksheets(SheetName).Delete
    Application.DisplayAlerts = True
    Exit Function

Error_Handler:
    Application.DisplayAlerts = True
    SystemCore.LogError Err.Number, Err.Description, "DeleteSheet", "DataOperations"
    ' Re-raise the error to maintain legacy behavior
    Err.Raise Err.Number, Err.Source, Err.Description
End Function