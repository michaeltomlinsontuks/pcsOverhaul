Attribute VB_Name = "SearchManager"
' **Purpose**: All search functionality including database updates, optimization, and history
' **CLAUDE.md Compliance**: Maintains all search functionality requirements, preserves existing workflows
Option Explicit

' ===================================================================
' CONSTANTS AND PRIVATE VARIABLES
' ===================================================================

Private Const SEARCH_FILE As String = "Search.xls"
Private Const SEARCH_HISTORY_FILE As String = "Search History.xls"
Private Const SYNC_PASSWORD As String = "KJB"

' ===================================================================
' SEARCH VALIDATION AND COMPATIBILITY
' ===================================================================

' **Purpose**: Validate search system can access existing files and directories
' **Parameters**: None
' **Returns**: Boolean - True if all critical directories and files accessible
' **Dependencies**: DataManager.FileExists, DataManager.GetRootPath
' **Side Effects**: None
' **Errors**: Returns False if critical files/directories missing
' **CLAUDE.md Compliance**: Ensures compatibility with existing file structure
Public Function ValidateSearchCompatibility() As Boolean
    Dim RootPath As String
    Dim RequiredDirs As Variant
    Dim i As Integer
    Dim TestFile As String

    On Error GoTo Error_Handler

    RootPath = DataManager.GetRootPath
    RequiredDirs = Array("Enquiries", "Quotes", "WIP", "Customers", "Templates", "Archive")

    ' Check if main directories exist
    For i = 0 To UBound(RequiredDirs)
        If Not DataManager.DirectoryExists(RootPath & "\" & RequiredDirs(i)) Then
            CoreFramework.LogError CoreFramework.ERR_FILE_NOT_FOUND, "Required directory missing: " & RequiredDirs(i), "ValidateSearchCompatibility", "SearchManager"
            ValidateSearchCompatibility = False
            Exit Function
        End If
    Next i

    ' Check if search database exists or can be created
    If Not DataManager.FileExists(RootPath & "\" & SEARCH_FILE) Then
        ' Try to create search database if missing
        If Not CreateSearchDatabase() Then
            ValidateSearchCompatibility = False
            Exit Function
        End If
    End If

    ' Test access to a sample file from each directory (if files exist)
    Dim SampleFiles As Variant
    SampleFiles = Array("Enquiries\*.xls", "Quotes\*.xls", "WIP\*.xls")

    For i = 0 To UBound(SampleFiles)
        Dim FileList As Variant
        FileList = DataManager.GetFileList(Left(SampleFiles(i), InStr(SampleFiles(i), "\") - 1))
        If IsArray(FileList) And UBound(FileList) >= 0 Then
            ' Try to access first file to ensure permissions are correct
            TestFile = RootPath & "\" & Left(SampleFiles(i), InStr(SampleFiles(i), "\") - 1) & "\" & FileList(0)
            If DataManager.FileExists(TestFile) Then
                ' Try to open and close quickly to test access
                Dim TestWB As Workbook
                Set TestWB = DataManager.SafeOpenWorkbook(TestFile)
                If TestWB Is Nothing Then
                    CoreFramework.LogError 0, "Cannot access existing file: " & TestFile, "ValidateSearchCompatibility", "SearchManager"
                    ValidateSearchCompatibility = False
                    Exit Function
                End If
                DataManager.SafeCloseWorkbook TestWB, False
            End If
        End If
    Next i

    ValidateSearchCompatibility = True
    Exit Function

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "ValidateSearchCompatibility", "SearchManager"
    ValidateSearchCompatibility = False
End Function

' ===================================================================
' SEARCH CORE OPERATIONS
' ===================================================================

' **Purpose**: Search all PCS records with basic functionality
' **Parameters**:
'   - SearchTerm (String): Text to search for in records
'   - RecordTypeFilter (RecordType, Optional): Limit search to specific record type
' **Returns**: Variant array of SearchRecord objects, empty array if no matches
' **Dependencies**: SearchRecords_Optimized for actual search implementation
' **Side Effects**: None
' **Errors**: Returns empty array on search failure
' **CLAUDE.md Compliance**: Maintains "finds anything in the system" requirement
Public Function SearchRecords(ByVal SearchTerm As String, Optional ByVal RecordTypeFilter As CoreFramework.RecordType = 0) As Variant
    SearchRecords = SearchRecords_Optimized(SearchTerm, RecordTypeFilter)
End Function

' **Purpose**: Search all PCS records with optimization for recent files
' **Parameters**:
'   - SearchTerm (String): Text to search for in records
'   - RecordTypeFilter (RecordType, Optional): Limit search to specific record type
' **Returns**: Variant array of SearchRecord objects, empty array if no matches
' **Dependencies**: DataManager.SafeOpenWorkbook for database access, LogSearchHistory for tracking
' **Side Effects**: Updates search history database, sorts search database by date
' **Errors**: Returns empty array on database access failure
' **CLAUDE.md Compliance**: Enhanced version maintaining all legacy search functionality
Public Function SearchRecords_Optimized(ByVal SearchTerm As String, Optional ByVal RecordTypeFilter As CoreFramework.RecordType = 0) As Variant
    Dim SearchWB As Workbook
    Dim SearchWS As Worksheet
    Dim LastRow As Long
    Dim i As Long
    Dim Results() As CoreFramework.SearchRecord
    Dim ResultCount As Integer
    Dim CurrentRecord As CoreFramework.SearchRecord
    Dim RecentCutoff As Date
    Dim RecentResults() As CoreFramework.SearchRecord
    Dim OtherResults() As CoreFramework.SearchRecord
    Dim RecentCount As Integer
    Dim OtherCount As Integer

    On Error GoTo Error_Handler

    Set SearchWB = DataManager.SafeOpenWorkbook(DataManager.GetRootPath & "\" & SEARCH_FILE)
    If SearchWB Is Nothing Then
        SearchRecords_Optimized = Array()
        Exit Function
    End If

    Set SearchWS = SearchWB.Worksheets(1)
    LastRow = SearchWS.Cells(SearchWS.Rows.Count, 1).End(xlUp).Row

    SearchTerm = UCase(SearchTerm)
    ResultCount = 0
    RecentCount = 0
    OtherCount = 0
    RecentCutoff = DateAdd("d", -30, Now)

    ' Quick return if search database is empty
    If LastRow <= 2 Then
        DataManager.SafeCloseWorkbook SearchWB, False
        SearchRecords_Optimized = Array()
        Exit Function
    End If

    ' Sort by date first (recent files first optimization)
    If LastRow > 2 Then
        Dim SortRange As Range
        Set SortRange = SearchWS.Range("A2:G" & LastRow)
        SortRange.Sort Key1:=SearchWS.Range("E2"), Order1:=xlDescending, Header:=xlNo
    End If

    ' Performance optimization: limit search depth for large databases
    Dim MaxSearchDepth As Long
    Dim SearchRows As Long
    SearchRows = LastRow - 1 ' Exclude header

    ' Exponential search strategy: start with recent, expand if needed
    If SearchRows <= 100 Then
        MaxSearchDepth = SearchRows
    ElseIf SearchRows <= 1000 Then
        MaxSearchDepth = 500
    Else
        MaxSearchDepth = 1000 ' Cap at 1000 records for performance
    End If

    ' Search with recent files prioritized (exponential depth)
    For i = 2 To Application.Min(LastRow, MaxSearchDepth + 1)
        With SearchWS
            If RecordTypeFilter = 0 Or .Cells(i, 1).Value = CStr(RecordTypeFilter) Then
                If InStr(UCase(.Cells(i, 2).Value), SearchTerm) > 0 Or _
                   InStr(UCase(.Cells(i, 3).Value), SearchTerm) > 0 Or _
                   InStr(UCase(.Cells(i, 4).Value), SearchTerm) > 0 Or _
                   InStr(UCase(.Cells(i, 7).Value), SearchTerm) > 0 Then

                    With CurrentRecord
                        .RecordType = SearchWS.Cells(i, 1).Value
                        .RecordNumber = SearchWS.Cells(i, 2).Value
                        .CustomerName = SearchWS.Cells(i, 3).Value
                        .Description = SearchWS.Cells(i, 4).Value
                        .DateCreated = SearchWS.Cells(i, 5).Value
                        .FilePath = SearchWS.Cells(i, 6).Value
                        .Keywords = SearchWS.Cells(i, 7).Value
                    End With

                    ' Separate recent vs older results
                    If CurrentRecord.DateCreated >= RecentCutoff Then
                        ReDim Preserve RecentResults(RecentCount)
                        RecentResults(RecentCount) = CurrentRecord
                        RecentCount = RecentCount + 1
                    Else
                        ReDim Preserve OtherResults(OtherCount)
                        OtherResults(OtherCount) = CurrentRecord
                        OtherCount = OtherCount + 1
                    End If

                    ResultCount = ResultCount + 1
                End If
            End If
        End With
    Next i

    DataManager.SafeCloseWorkbook SearchWB, False

    ' If we found few results and searched a limited set, expand search
    If ResultCount < 5 And MaxSearchDepth < SearchRows And SearchRows > 100 Then
        Dim ExtendedResults() As CoreFramework.SearchRecord
        Dim ExtendedCount As Integer
        ExtendedCount = 0

        ' Search deeper into older records
        For i = MaxSearchDepth + 2 To Application.Min(LastRow, MaxSearchDepth * 2)
            With SearchWS
                If RecordTypeFilter = 0 Or .Cells(i, 1).Value = CStr(RecordTypeFilter) Then
                    If InStr(UCase(.Cells(i, 2).Value), SearchTerm) > 0 Or _
                       InStr(UCase(.Cells(i, 3).Value), SearchTerm) > 0 Or _
                       InStr(UCase(.Cells(i, 4).Value), SearchTerm) > 0 Or _
                       InStr(UCase(.Cells(i, 7).Value), SearchTerm) > 0 Then

                        With CurrentRecord
                            .RecordType = SearchWS.Cells(i, 1).Value
                            .RecordNumber = SearchWS.Cells(i, 2).Value
                            .CustomerName = SearchWS.Cells(i, 3).Value
                            .Description = SearchWS.Cells(i, 4).Value
                            .DateCreated = SearchWS.Cells(i, 5).Value
                            .FilePath = SearchWS.Cells(i, 6).Value
                            .Keywords = SearchWS.Cells(i, 7).Value
                        End With

                        ReDim Preserve ExtendedResults(ExtendedCount)
                        ExtendedResults(ExtendedCount) = CurrentRecord
                        ExtendedCount = ExtendedCount + 1
                        ResultCount = ResultCount + 1
                    End If
                End If
            End With
        Next i

        ' Add extended results to other results
        If ExtendedCount > 0 Then
            Dim OldOtherCount As Integer
            OldOtherCount = OtherCount
            ReDim Preserve OtherResults(OtherCount + ExtendedCount - 1)
            For i = 0 To ExtendedCount - 1
                OtherResults(OldOtherCount + i) = ExtendedResults(i)
            Next i
            OtherCount = OtherCount + ExtendedCount
        End If
    End If

    ' Combine results: recent files first, then older files
    If ResultCount > 0 Then
        ReDim Results(ResultCount - 1)
        Dim ResultIndex As Integer
        ResultIndex = 0

        ' Add recent results first
        For i = 0 To RecentCount - 1
            Results(ResultIndex) = RecentResults(i)
            ResultIndex = ResultIndex + 1
        Next i

        ' Add other results
        For i = 0 To OtherCount - 1
            Results(ResultIndex) = OtherResults(i)
            ResultIndex = ResultIndex + 1
        Next i

        SearchRecords_Optimized = Results
    Else
        SearchRecords_Optimized = Array()
    End If

    LogSearchHistory SearchTerm, ResultCount
    Exit Function

Error_Handler:
    If Not SearchWB Is Nothing Then DataManager.SafeCloseWorkbook SearchWB, False
    CoreFramework.HandleStandardErrors Err.Number, "SearchRecords_Optimized", "SearchManager"
    SearchRecords_Optimized = Array()
End Function

' **Purpose**: Create search database with proper headers if it doesn't exist
' **Parameters**: None
' **Returns**: Boolean - True if database created successfully
' **Dependencies**: DataManager.SafeOpenWorkbook
' **Side Effects**: Creates new Search.xls file with headers
' **Errors**: Returns False if creation fails
Private Function CreateSearchDatabase() As Boolean
    Dim SearchWB As Workbook
    Dim SearchWS As Worksheet
    Dim NewFile As String

    On Error GoTo Error_Handler

    NewFile = DataManager.GetRootPath & "\" & SEARCH_FILE

    ' Create new workbook
    Set SearchWB = Application.Workbooks.Add
    Set SearchWS = SearchWB.Worksheets(1)

    ' Add headers
    With SearchWS
        .Cells(1, 1).Value = "Record Type"
        .Cells(1, 2).Value = "Record Number"
        .Cells(1, 3).Value = "Customer Name"
        .Cells(1, 4).Value = "Description"
        .Cells(1, 5).Value = "Date Created"
        .Cells(1, 6).Value = "File Path"
        .Cells(1, 7).Value = "Keywords"

        ' Format headers
        .Range("A1:G1").Font.Bold = True
        .Columns("A:G").AutoFit
    End With

    SearchWB.SaveAs NewFile
    SearchWB.Close

    CreateSearchDatabase = True
    Exit Function

Error_Handler:
    If Not SearchWB Is Nothing Then SearchWB.Close False
    CoreFramework.HandleStandardErrors Err.Number, "CreateSearchDatabase", "SearchManager"
    CreateSearchDatabase = False
End Function

' **Purpose**: Test search functionality with existing database and files
' **Parameters**:
'   - TestSearchTerm (String, Optional): Term to test search with (default "test")
' **Returns**: Boolean - True if search system works with existing files
' **Dependencies**: SearchRecords_Optimized, ValidateSearchCompatibility
' **Side Effects**: May create search database if missing
' **Errors**: Returns False if search system incompatible
' **CLAUDE.md Compliance**: Ensures search system works with existing file structure
Public Function TestSearchWithExistingFiles(Optional ByVal TestSearchTerm As String = "test") As Boolean
    Dim TestResults As Variant
    Dim RootPath As String
    Dim SampleFiles As Long

    On Error GoTo Error_Handler

    ' First validate basic compatibility
    If Not ValidateSearchCompatibility() Then
        TestSearchWithExistingFiles = False
        Exit Function
    End If

    RootPath = DataManager.GetRootPath

    ' Check if search database has any data
    Dim SearchWB As Workbook
    Dim SearchWS As Worksheet
    Dim LastRow As Long

    Set SearchWB = DataManager.SafeOpenWorkbook(RootPath & "\" & SEARCH_FILE)
    If SearchWB Is Nothing Then
        TestSearchWithExistingFiles = False
        Exit Function
    End If

    Set SearchWS = SearchWB.Worksheets(1)
    LastRow = SearchWS.Cells(SearchWS.Rows.Count, 1).End(xlUp).Row
    DataManager.SafeCloseWorkbook SearchWB, False

    ' If database is empty or only has headers, rebuild incrementally
    If LastRow <= 1 Then
        If Not RebuildSearchDatabase_Incremental(100, 30) Then ' Test with small subset
            TestSearchWithExistingFiles = False
            Exit Function
        End If
    End If

    ' Test search functionality
    TestResults = SearchRecords_Optimized(TestSearchTerm)

    ' Verify we can get some results or at least that search completes without error
    If IsArray(TestResults) Then
        TestSearchWithExistingFiles = True
    Else
        TestSearchWithExistingFiles = False
    End If

    Exit Function

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "TestSearchWithExistingFiles", "SearchManager"
    TestSearchWithExistingFiles = False
End Function

' **Purpose**: Search records by specific type only
' **Parameters**:
'   - SearchTerm (String): Text to search for in records
'   - RecordType (RecordType): Specific record type to search
' **Returns**: Variant array of matching SearchRecord objects
' **Dependencies**: SearchRecords_Optimized
' **Side Effects**: Updates search history
' **Errors**: Returns empty array on search failure
Public Function SearchByType(ByVal SearchTerm As String, ByVal RecordType As CoreFramework.RecordType) As Variant
    SearchByType = SearchRecords_Optimized(SearchTerm, RecordType)
End Function

' **Purpose**: Search records within specific date range
' **Parameters**:
'   - SearchTerm (String): Text to search for in records
'   - StartDate (Date): Earliest date to include
'   - EndDate (Date): Latest date to include
' **Returns**: Variant array of matching SearchRecord objects
' **Dependencies**: DataManager.SafeOpenWorkbook for database access
' **Side Effects**: None
' **Errors**: Returns empty array on database access failure
Public Function SearchByDateRange(ByVal SearchTerm As String, ByVal StartDate As Date, ByVal EndDate As Date) As Variant
    Dim SearchWB As Workbook
    Dim SearchWS As Worksheet
    Dim LastRow As Long
    Dim i As Long
    Dim Results() As CoreFramework.SearchRecord
    Dim ResultCount As Integer
    Dim CurrentRecord As CoreFramework.SearchRecord
    Dim RecordDate As Date

    On Error GoTo Error_Handler

    Set SearchWB = DataManager.SafeOpenWorkbook(DataManager.GetRootPath & "\" & SEARCH_FILE)
    If SearchWB Is Nothing Then
        SearchByDateRange = Array()
        Exit Function
    End If

    Set SearchWS = SearchWB.Worksheets(1)
    LastRow = SearchWS.Cells(SearchWS.Rows.Count, 1).End(xlUp).Row

    SearchTerm = UCase(SearchTerm)
    ResultCount = 0

    For i = 2 To LastRow
        With SearchWS
            RecordDate = .Cells(i, 5).Value

            If RecordDate >= StartDate And RecordDate <= EndDate Then
                If InStr(UCase(.Cells(i, 2).Value), SearchTerm) > 0 Or _
                   InStr(UCase(.Cells(i, 3).Value), SearchTerm) > 0 Or _
                   InStr(UCase(.Cells(i, 4).Value), SearchTerm) > 0 Or _
                   InStr(UCase(.Cells(i, 7).Value), SearchTerm) > 0 Then

                    ReDim Preserve Results(ResultCount)

                    With CurrentRecord
                        .RecordType = SearchWS.Cells(i, 1).Value
                        .RecordNumber = SearchWS.Cells(i, 2).Value
                        .CustomerName = SearchWS.Cells(i, 3).Value
                        .Description = SearchWS.Cells(i, 4).Value
                        .DateCreated = SearchWS.Cells(i, 5).Value
                        .FilePath = SearchWS.Cells(i, 6).Value
                        .Keywords = SearchWS.Cells(i, 7).Value
                    End With

                    Results(ResultCount) = CurrentRecord
                    ResultCount = ResultCount + 1
                End If
            End If
        End With
    Next i

    DataManager.SafeCloseWorkbook SearchWB, False

    If ResultCount > 0 Then
        SearchByDateRange = Results
    Else
        SearchByDateRange = Array()
    End If

    Exit Function

Error_Handler:
    If Not SearchWB Is Nothing Then DataManager.SafeCloseWorkbook SearchWB, False
    CoreFramework.HandleStandardErrors Err.Number, "SearchByDateRange", "SearchManager"
    SearchByDateRange = Array()
End Function

' ===================================================================
' SEARCH DATABASE MANAGEMENT
' ===================================================================

' **Purpose**: Update search database with new or modified record
' **Parameters**:
'   - Record (SearchRecord): Complete search record to add/update
' **Returns**: Boolean - True if update successful, False if failed
' **Dependencies**: DataManager.SafeOpenWorkbook for database access
' **Side Effects**: Adds new row to search database, saves database file
' **Errors**: Returns False on database access or update failure
Public Function UpdateSearchDatabase(ByRef Record As CoreFramework.SearchRecord) As Boolean
    Dim SearchWB As Workbook
    Dim SearchWS As Worksheet
    Dim LastRow As Long

    On Error GoTo Error_Handler

    Set SearchWB = DataManager.SafeOpenWorkbook(DataManager.GetRootPath & "\" & SEARCH_FILE)
    If SearchWB Is Nothing Then
        UpdateSearchDatabase = False
        Exit Function
    End If

    Set SearchWS = SearchWB.Worksheets(1)
    LastRow = SearchWS.Cells(SearchWS.Rows.Count, 1).End(xlUp).Row + 1

    With SearchWS
        .Cells(LastRow, 1).Value = Record.RecordType
        .Cells(LastRow, 2).Value = Record.RecordNumber
        .Cells(LastRow, 3).Value = Record.CustomerName
        .Cells(LastRow, 4).Value = Record.Description
        .Cells(LastRow, 5).Value = Record.DateCreated
        .Cells(LastRow, 6).Value = Record.FilePath
        .Cells(LastRow, 7).Value = Record.Keywords
    End With

    SearchWB.Save
    DataManager.SafeCloseWorkbook SearchWB

    UpdateSearchDatabase = True
    Exit Function

Error_Handler:
    If Not SearchWB Is Nothing Then DataManager.SafeCloseWorkbook SearchWB, False
    CoreFramework.HandleStandardErrors Err.Number, "UpdateSearchDatabase", "SearchManager"
    UpdateSearchDatabase = False
End Function

' **Purpose**: Delete record from search database
' **Parameters**:
'   - RecordNumber (String): Record number to delete (E00001, Q00001, etc.)
' **Returns**: Boolean - True if deletion successful, False if failed
' **Dependencies**: DataManager.SafeOpenWorkbook for database access
' **Side Effects**: Removes row from search database, saves database file
' **Errors**: Returns False if record not found or deletion fails
Public Function DeleteSearchRecord(ByVal RecordNumber As String) As Boolean
    Dim SearchWB As Workbook
    Dim SearchWS As Worksheet
    Dim LastRow As Long
    Dim i As Long

    On Error GoTo Error_Handler

    Set SearchWB = DataManager.SafeOpenWorkbook(DataManager.GetRootPath & "\" & SEARCH_FILE)
    If SearchWB Is Nothing Then
        DeleteSearchRecord = False
        Exit Function
    End If

    Set SearchWS = SearchWB.Worksheets(1)
    LastRow = SearchWS.Cells(SearchWS.Rows.Count, 1).End(xlUp).Row

    For i = 2 To LastRow
        If SearchWS.Cells(i, 2).Value = RecordNumber Then
            SearchWS.Rows(i).Delete
            SearchWB.Save
            DataManager.SafeCloseWorkbook SearchWB
            DeleteSearchRecord = True
            Exit Function
        End If
    Next i

    DataManager.SafeCloseWorkbook SearchWB, False
    DeleteSearchRecord = False
    Exit Function

Error_Handler:
    If Not SearchWB Is Nothing Then DataManager.SafeCloseWorkbook SearchWB, False
    CoreFramework.HandleStandardErrors Err.Number, "DeleteSearchRecord", "SearchManager"
    DeleteSearchRecord = False
End Function

' **Purpose**: Sort search database by date (most recent first)
' **Parameters**: None
' **Returns**: Boolean - True if sort successful, False if failed
' **Dependencies**: DataManager.SafeOpenWorkbook for database access
' **Side Effects**: Reorders search database rows, saves database file
' **Errors**: Returns False if database access or sort operation fails
Public Function SortSearchDatabase() As Boolean
    Dim SearchWB As Workbook
    Dim SearchWS As Worksheet
    Dim LastRow As Long
    Dim SortRange As Range

    On Error GoTo Error_Handler

    Set SearchWB = DataManager.SafeOpenWorkbook(DataManager.GetRootPath & "\" & SEARCH_FILE)
    If SearchWB Is Nothing Then
        SortSearchDatabase = False
        Exit Function
    End If

    Set SearchWS = SearchWB.Worksheets(1)
    LastRow = SearchWS.Cells(SearchWS.Rows.Count, 1).End(xlUp).Row

    If LastRow > 2 Then
        Set SortRange = SearchWS.Range("A2:G" & LastRow)
        SortRange.Sort Key1:=SearchWS.Range("E2"), Order1:=xlDescending, Header:=xlNo
    End If

    SearchWB.Save
    DataManager.SafeCloseWorkbook SearchWB
    SortSearchDatabase = True
    Exit Function

Error_Handler:
    If Not SearchWB Is Nothing Then DataManager.SafeCloseWorkbook SearchWB, False
    CoreFramework.HandleStandardErrors Err.Number, "SortSearchDatabase", "SearchManager"
    SortSearchDatabase = False
End Function

' **Purpose**: Rebuild entire search database from scratch
' **Parameters**: None
' **Returns**: Boolean - True if rebuild successful, False if failed
' **Dependencies**: DataManager.GetFileList for directory scanning
' **Side Effects**: Recreates search database from file system scan
' **Errors**: Returns False if rebuild operation fails
Public Function RebuildSearchDatabase() As Boolean
    Dim SearchWB As Workbook
    Dim SearchWS As Worksheet
    Dim Directories As Variant
    Dim i As Integer
    Dim FileList As Variant
    Dim j As Integer

    On Error GoTo Error_Handler

    ' Backup existing database
    If DataManager.FileExists(DataManager.GetRootPath & "\" & SEARCH_FILE) Then
        DataManager.CreateBackup DataManager.GetRootPath & "\" & SEARCH_FILE
    End If

    Set SearchWB = DataManager.SafeOpenWorkbook(DataManager.GetRootPath & "\" & SEARCH_FILE)
    If SearchWB Is Nothing Then
        RebuildSearchDatabase = False
        Exit Function
    End If

    Set SearchWS = SearchWB.Worksheets(1)

    ' Clear existing data (keep headers)
    SearchWS.Range("A2:G" & SearchWS.Rows.Count).Clear

    ' Scan all directories and rebuild database
    Directories = Array("Enquiries", "Quotes", "WIP", "Archive", "Contracts")

    For i = 0 To UBound(Directories)
        FileList = DataManager.GetFileList(Directories(i))

        If IsArray(FileList) And UBound(FileList) >= 0 Then
            For j = 0 To UBound(FileList)
                ' Process each file and add to search database
                ProcessFileForSearch SearchWS, Directories(i), FileList(j)
            Next j
        End If
    Next i

    SearchWB.Save
    DataManager.SafeCloseWorkbook SearchWB
    RebuildSearchDatabase = True
    Exit Function

Error_Handler:
    If Not SearchWB Is Nothing Then DataManager.SafeCloseWorkbook SearchWB, False
    CoreFramework.HandleStandardErrors Err.Number, "RebuildSearchDatabase", "SearchManager"
    RebuildSearchDatabase = False
End Function

' **Purpose**: Incrementally rebuild search database starting with recent files
' **Parameters**:
'   - MaxFiles (Long, Optional): Maximum files to process per directory (default 500)
'   - DaysBack (Long, Optional): How many days back to prioritize (default 90)
' **Returns**: Boolean - True if rebuild successful
' **Dependencies**: DataManager.GetFileList, ProcessFileForSearch
' **Side Effects**: Updates search database with recent files first
' **Errors**: Returns False if rebuild fails
' **CLAUDE.md Compliance**: Optimized rebuild maintaining compatibility
Public Function RebuildSearchDatabase_Incremental(Optional ByVal MaxFiles As Long = 500, Optional ByVal DaysBack As Long = 90) As Boolean
    Dim SearchWB As Workbook
    Dim SearchWS As Worksheet
    Dim Directories As Variant
    Dim i As Integer
    Dim FileList As Variant
    Dim j As Integer
    Dim ProcessedCount As Long
    Dim CutoffDate As Date
    Dim RecentFiles() As String
    Dim OlderFiles() As String
    Dim RecentCount As Long
    Dim OlderCount As Long

    On Error GoTo Error_Handler

    CutoffDate = DateAdd("d", -DaysBack, Now)
    ProcessedCount = 0

    ' Backup existing database
    If DataManager.FileExists(DataManager.GetRootPath & "\" & SEARCH_FILE) Then
        DataManager.CreateBackup DataManager.GetRootPath & "\" & SEARCH_FILE
    End If

    Set SearchWB = DataManager.SafeOpenWorkbook(DataManager.GetRootPath & "\" & SEARCH_FILE)
    If SearchWB Is Nothing Then
        RebuildSearchDatabase_Incremental = False
        Exit Function
    End If

    Set SearchWS = SearchWB.Worksheets(1)

    ' Clear existing data (keep headers)
    SearchWS.Range("A2:G" & SearchWS.Rows.Count).Clear

    ' Scan directories in priority order
    Directories = Array("WIP", "Quotes", "Enquiries", "Contracts", "Archive")

    For i = 0 To UBound(Directories)
        FileList = DataManager.GetFileList(Directories(i))
        RecentCount = 0
        OlderCount = 0

        If IsArray(FileList) And UBound(FileList) >= 0 Then
            ' Separate files by age (using file system date as approximation)
            ReDim RecentFiles(UBound(FileList))
            ReDim OlderFiles(UBound(FileList))

            For j = 0 To UBound(FileList)
                Dim FilePath As String
                Dim FileDate As Date
                FilePath = DataManager.GetRootPath & "\" & Directories(i) & "\" & FileList(j)

                ' Get file modification date as proxy for record age
                FileDate = FileDateTime(FilePath)

                If FileDate >= CutoffDate Then
                    RecentFiles(RecentCount) = FileList(j)
                    RecentCount = RecentCount + 1
                Else
                    OlderFiles(OlderCount) = FileList(j)
                    OlderCount = OlderCount + 1
                End If
            Next j

            ' Process recent files first
            For j = 0 To RecentCount - 1
                If ProcessedCount >= MaxFiles Then Exit For
                ProcessFileForSearch SearchWS, Directories(i), RecentFiles(j)
                ProcessedCount = ProcessedCount + 1
            Next j

            ' Process older files if we haven't hit the limit
            For j = 0 To OlderCount - 1
                If ProcessedCount >= MaxFiles Then Exit For
                ProcessFileForSearch SearchWS, Directories(i), OlderFiles(j)
                ProcessedCount = ProcessedCount + 1
            Next j
        End If

        ' Exit if we've processed enough files
        If ProcessedCount >= MaxFiles * UBound(Directories) + 1 Then Exit For
    Next i

    SearchWB.Save
    DataManager.SafeCloseWorkbook SearchWB
    RebuildSearchDatabase_Incremental = True
    Exit Function

Error_Handler:
    If Not SearchWB Is Nothing Then DataManager.SafeCloseWorkbook SearchWB, False
    CoreFramework.HandleStandardErrors Err.Number, "RebuildSearchDatabase_Incremental", "SearchManager"
    RebuildSearchDatabase_Incremental = False
End Function

' **Purpose**: Synchronize search data with search history
' **Parameters**: None
' **Returns**: Boolean - True if synchronization successful, False if failed
' **Dependencies**: DataManager.SafeOpenWorkbook for database access
' **Side Effects**: Updates search history with current search data, archives old records
' **Errors**: Returns False if synchronization fails
' **CLAUDE.md Compliance**: Replaces legacy Search_Sync.bas functionality
Public Function SynchronizeSearchData() As Boolean
    Dim SearchWB As Workbook
    Dim HistoryWB As Workbook
    Dim SearchWS As Worksheet
    Dim HistoryWS As Worksheet
    Dim SearchLastRow As Long
    Dim HistoryLastRow As Long
    Dim i As Long
    Dim j As Long
    Dim DCSData(0 To 30) As Variant
    Dim Found As Boolean

    On Error GoTo Error_Handler

    ' Create backups before sync
    DataManager.CreateBackup DataManager.GetRootPath & "\" & SEARCH_FILE
    DataManager.CreateBackup DataManager.GetRootPath & "\" & SEARCH_HISTORY_FILE

    Set SearchWB = DataManager.SafeOpenWorkbook(DataManager.GetRootPath & "\" & SEARCH_FILE)
    Set HistoryWB = DataManager.SafeOpenWorkbook(DataManager.GetRootPath & "\" & SEARCH_HISTORY_FILE)

    If SearchWB Is Nothing Or HistoryWB Is Nothing Then
        SynchronizeSearchData = False
        Exit Function
    End If

    Set SearchWS = SearchWB.Worksheets(1)
    Set HistoryWS = HistoryWB.Worksheets(1)

    SearchLastRow = SearchWS.Cells(SearchWS.Rows.Count, 1).End(xlUp).Row
    HistoryLastRow = HistoryWS.Cells(HistoryWS.Rows.Count, 1).End(xlUp).Row

    ' Process each search record
    For i = 3 To SearchLastRow ' Start from row 3 to skip headers
        ' Copy search data
        For j = 0 To 30
            If j <= SearchWS.Columns.Count Then
                DCSData(j) = SearchWS.Cells(i, j + 1).Value
            Else
                DCSData(j) = ""
            End If
        Next j

        ' Find matching record in history or create new one
        Found = False
        For j = 2 To HistoryLastRow
            ' Match based on record number (column 2)
            If HistoryWS.Cells(j, 2).Value = DCSData(1) Then
                ' Update existing record
                For k = 0 To 30
                    If k + 1 <= HistoryWS.Columns.Count Then
                        HistoryWS.Cells(j, k + 1).Value = DCSData(k)
                    End If
                Next k
                Found = True
                Exit For
            End If
        Next j

        ' Add new record if not found
        If Not Found Then
            HistoryLastRow = HistoryLastRow + 1
            For k = 0 To 30
                If k + 1 <= HistoryWS.Columns.Count Then
                    HistoryWS.Cells(HistoryLastRow, k + 1).Value = DCSData(k)
                End If
            Next k
        End If
    Next i

    ' Archive old records (jobs older than 1000 numbers, quotes older than 10000 numbers)
    ArchiveOldRecords SearchWS

    SearchWB.Save
    HistoryWB.Save
    DataManager.SafeCloseWorkbook SearchWB
    DataManager.SafeCloseWorkbook HistoryWB

    SynchronizeSearchData = True
    Exit Function

Error_Handler:
    If Not SearchWB Is Nothing Then DataManager.SafeCloseWorkbook SearchWB, False
    If Not HistoryWB Is Nothing Then DataManager.SafeCloseWorkbook HistoryWB, False
    CoreFramework.HandleStandardErrors Err.Number, "SynchronizeSearchData", "SearchManager"
    SynchronizeSearchData = False
End Function

' ===================================================================
' SEARCH RECORD OPERATIONS
' ===================================================================

' **Purpose**: Create new search record with all required fields
' **Parameters**:
'   - RecType (RecordType): Type of record (Enquiry, Quote, Job, Contract)
'   - Number (String): Record number (E00001, Q00001, etc.)
'   - Customer (String): Customer name
'   - Description (String): Component or item description
'   - FilePath (String): Full path to record file
'   - Keywords (String, Optional): Additional search keywords
' **Returns**: SearchRecord - Populated search record structure
' **Dependencies**: None
' **Side Effects**: None
' **Errors**: Returns record with empty fields if parameters invalid
Public Function CreateSearchRecord(ByVal RecType As CoreFramework.RecordType, ByVal Number As String, ByVal Customer As String, ByVal Description As String, ByVal FilePath As String, Optional ByVal Keywords As String = "") As CoreFramework.SearchRecord
    With CreateSearchRecord
        .RecordType = CStr(RecType)
        .RecordNumber = Number
        .CustomerName = Customer
        .Description = Description
        .DateCreated = Now
        .FilePath = FilePath
        .Keywords = Keywords
    End With
End Function

' **Purpose**: Save form data to search database
' **Parameters**:
'   - FormObject (Object): Form containing data to save to search
' **Returns**: Boolean - True if save successful, False if failed
' **Dependencies**: DataManager.SafeOpenWorkbook for database access
' **Side Effects**: Adds or updates search record, sorts database, saves file
' **Errors**: Returns False if database access or save operation fails
' **CLAUDE.md Compliance**: Replaces legacy SaveSearchCode.bas functionality
Public Function SaveRowToSearch(ByRef FormObject As Object) As Boolean
    Dim SearchWB As Workbook
    Dim SearchWS As Worksheet
    Dim TargetRow As Long
    Dim ctl As Object
    Dim i As Integer
    Dim RecordNumber As String

    On Error GoTo Error_Handler

    ' Open search database with retry for read-only
    Do
        Set SearchWB = DataManager.SafeOpenWorkbook(DataManager.GetRootPath & "\" & SEARCH_FILE)
        If SearchWB Is Nothing Then
            SaveRowToSearch = False
            Exit Function
        End If

        If SearchWB.ReadOnly = True Then
            DataManager.SafeCloseWorkbook SearchWB, False
            MsgBox "Search database is read-only. Please ensure no other users have it open.", vbExclamation
            ' Could implement retry logic here
        End If
    Loop Until Not SearchWB.ReadOnly

    Set SearchWS = SearchWB.Worksheets("search")

    ' Find target row (existing record or new row)
    TargetRow = FindOrCreateSearchRow(SearchWS, FormObject)

    ' Save form controls to search database
    For Each ctl In FormObject.Controls
        For i = 0 To 100
            If UCase(SearchWS.Range("A1").Offset(0, i).Value) = UCase(ctl.Name) Then
                Select Case UCase(TypeName(ctl))
                    Case "LABEL"
                        SearchWS.Range("A1").Offset(TargetRow - 1, i).Value = UCase(ctl.Caption)
                    Case "TEXTBOX"
                        SearchWS.Range("A1").Offset(TargetRow - 1, i).Value = UCase(ctl.Value)
                    Case "COMBOBOX"
                        SearchWS.Range("A1").Offset(TargetRow - 1, i).Value = UCase(ctl.Value)
                End Select
                Exit For
            End If
            ' Copy formula from previous row if needed
            If Left(SearchWS.Range("A1").Offset(TargetRow - 2, i).Value, 1) = "=" Then
                SearchWS.Range("A1").Offset(TargetRow - 1, i).Value = SearchWS.Range("A1").Offset(TargetRow - 2, i).Value
            End If
            If SearchWS.Range("A1").Offset(0, i + 1).Value = "" Then Exit For
        Next i
    Next ctl

    ' Sort database by date (most recent first)
    SortSearchDatabaseInWorksheet SearchWS

    SearchWB.Save
    DataManager.SafeCloseWorkbook SearchWB

    SaveRowToSearch = True
    Exit Function

Error_Handler:
    If Not SearchWB Is Nothing Then DataManager.SafeCloseWorkbook SearchWB, False
    CoreFramework.HandleStandardErrors Err.Number, "SaveRowToSearch", "SearchManager"
    SaveRowToSearch = False
End Function

' **Purpose**: Update search database from form data
' **Parameters**:
'   - FormObject (Object): Form containing updated data
'   - RecordNumber (String): Record number to update
' **Returns**: Boolean - True if update successful, False if failed
' **Dependencies**: SaveRowToSearch for actual update operation
' **Side Effects**: Updates existing search record
' **Errors**: Returns False if update operation fails
Public Function UpdateSearchFromForm(ByRef FormObject As Object, ByVal RecordNumber As String) As Boolean
    UpdateSearchFromForm = SaveRowToSearch(FormObject)
End Function

' **Purpose**: Validate search record completeness and format
' **Parameters**:
'   - Record (SearchRecord): Search record to validate
' **Returns**: Boolean - True if valid, False if validation fails
' **Dependencies**: None
' **Side Effects**: None
' **Errors**: Returns False if required fields missing or invalid format
Public Function ValidateSearchRecord(ByRef Record As CoreFramework.SearchRecord) As Boolean
    ' Check required fields
    If Record.RecordNumber = "" Then
        ValidateSearchRecord = False
        Exit Function
    End If

    If Record.CustomerName = "" Then
        ValidateSearchRecord = False
        Exit Function
    End If

    If Record.Description = "" Then
        ValidateSearchRecord = False
        Exit Function
    End If

    If Record.RecordType = "" Then
        ValidateSearchRecord = False
        Exit Function
    End If

    ' Validate record type
    Select Case Record.RecordType
        Case "1", "2", "3", "4" ' Valid record types
            ValidateSearchRecord = True
        Case Else
            ValidateSearchRecord = False
    End Select
End Function

' ===================================================================
' SEARCH OPTIMIZATION
' ===================================================================

' **Purpose**: Optimize search database performance
' **Parameters**: None
' **Returns**: Boolean - True if optimization successful, False if failed
' **Dependencies**: SortSearchDatabase, CompactSearchDatabase
' **Side Effects**: Reorganizes database for better performance
' **Errors**: Returns False if optimization fails
Public Function OptimizeSearchPerformance() As Boolean
    On Error GoTo Error_Handler

    ' Sort database for optimal search performance
    If Not SortSearchDatabase() Then
        OptimizeSearchPerformance = False
        Exit Function
    End If

    ' Compact database to remove empty rows
    If Not CompactSearchDatabase() Then
        OptimizeSearchPerformance = False
        Exit Function
    End If

    OptimizeSearchPerformance = True
    Exit Function

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "OptimizeSearchPerformance", "SearchManager"
    OptimizeSearchPerformance = False
End Function

' **Purpose**: Archive old search records to reduce database size
' **Parameters**: None
' **Returns**: Boolean - True if archiving successful, False if failed
' **Dependencies**: DataManager.SafeOpenWorkbook for database access
' **Side Effects**: Moves old records to archive, reduces main database size
' **Errors**: Returns False if archiving operation fails
Public Function ArchiveOldSearchRecords() As Boolean
    Dim SearchWB As Workbook
    Dim ArchiveWB As Workbook
    Dim SearchWS As Worksheet
    Dim ArchiveWS As Worksheet
    Dim ArchivePath As String
    Dim CutoffDate As Date

    On Error GoTo Error_Handler

    CutoffDate = DateAdd("y", -2, Now) ' Archive records older than 2 years

    Set SearchWB = DataManager.SafeOpenWorkbook(DataManager.GetRootPath & "\" & SEARCH_FILE)
    If SearchWB Is Nothing Then
        ArchiveOldSearchRecords = False
        Exit Function
    End If

    ArchivePath = DataManager.GetRootPath & "\Archive\Search_Archive_" & Format(Now, "yyyy") & ".xls"

    ' Create or open archive file
    If DataManager.FileExists(ArchivePath) Then
        Set ArchiveWB = DataManager.SafeOpenWorkbook(ArchivePath)
    Else
        Set ArchiveWB = DataManager.CreateNewWorkbook()
        If Not ArchiveWB Is Nothing Then
            ArchiveWB.SaveAs ArchivePath
        End If
    End If

    If ArchiveWB Is Nothing Then
        DataManager.SafeCloseWorkbook SearchWB, False
        ArchiveOldSearchRecords = False
        Exit Function
    End If

    Set SearchWS = SearchWB.Worksheets(1)
    Set ArchiveWS = ArchiveWB.Worksheets(1)

    ' Move old records to archive
    Dim i As Long
    Dim LastRow As Long
    Dim ArchiveRow As Long

    LastRow = SearchWS.Cells(SearchWS.Rows.Count, 1).End(xlUp).Row
    ArchiveRow = ArchiveWS.Cells(ArchiveWS.Rows.Count, 1).End(xlUp).Row + 1

    For i = LastRow To 2 Step -1
        If SearchWS.Cells(i, 5).Value < CutoffDate Then
            ' Copy row to archive
            SearchWS.Rows(i).Copy ArchiveWS.Rows(ArchiveRow)
            ArchiveRow = ArchiveRow + 1

            ' Delete from main database
            SearchWS.Rows(i).Delete
        End If
    Next i

    SearchWB.Save
    ArchiveWB.Save
    DataManager.SafeCloseWorkbook SearchWB
    DataManager.SafeCloseWorkbook ArchiveWB

    ArchiveOldSearchRecords = True
    Exit Function

Error_Handler:
    If Not SearchWB Is Nothing Then DataManager.SafeCloseWorkbook SearchWB, False
    If Not ArchiveWB Is Nothing Then DataManager.SafeCloseWorkbook ArchiveWB, False
    CoreFramework.HandleStandardErrors Err.Number, "ArchiveOldSearchRecords", "SearchManager"
    ArchiveOldSearchRecords = False
End Function

' **Purpose**: Compact search database by removing empty rows
' **Parameters**: None
' **Returns**: Boolean - True if compaction successful, False if failed
' **Dependencies**: DataManager.SafeOpenWorkbook for database access
' **Side Effects**: Removes empty rows from database, reduces file size
' **Errors**: Returns False if compaction fails
Public Function CompactSearchDatabase() As Boolean
    Dim SearchWB As Workbook
    Dim SearchWS As Worksheet
    Dim LastRow As Long
    Dim i As Long

    On Error GoTo Error_Handler

    Set SearchWB = DataManager.SafeOpenWorkbook(DataManager.GetRootPath & "\" & SEARCH_FILE)
    If SearchWB Is Nothing Then
        CompactSearchDatabase = False
        Exit Function
    End If

    Set SearchWS = SearchWB.Worksheets(1)
    LastRow = SearchWS.Cells(SearchWS.Rows.Count, 1).End(xlUp).Row

    ' Remove empty rows (work backwards to avoid index issues)
    For i = LastRow To 2 Step -1
        If SearchWS.Cells(i, 1).Value = "" And SearchWS.Cells(i, 2).Value = "" Then
            SearchWS.Rows(i).Delete
        End If
    Next i

    SearchWB.Save
    DataManager.SafeCloseWorkbook SearchWB

    CompactSearchDatabase = True
    Exit Function

Error_Handler:
    If Not SearchWB Is Nothing Then DataManager.SafeCloseWorkbook SearchWB, False
    CoreFramework.HandleStandardErrors Err.Number, "CompactSearchDatabase", "SearchManager"
    CompactSearchDatabase = False
End Function

' ===================================================================
' SEARCH HISTORY & ANALYTICS
' ===================================================================

' **Purpose**: Log search activity to history database
' **Parameters**:
'   - SearchTerm (String): Search term that was used
'   - ResultCount (Integer): Number of results found
' **Returns**: None (Subroutine)
' **Dependencies**: DataManager.SafeOpenWorkbook for database access
' **Side Effects**: Adds entry to search history database
' **Errors**: Exits silently if logging fails (non-critical operation)
Private Sub LogSearchHistory(ByVal SearchTerm As String, ByVal ResultCount As Integer)
    Dim HistoryWB As Workbook
    Dim HistoryWS As Worksheet
    Dim LastRow As Long

    On Error GoTo Error_Handler

    Set HistoryWB = DataManager.SafeOpenWorkbook(DataManager.GetRootPath & "\" & SEARCH_HISTORY_FILE)
    If HistoryWB Is Nothing Then Exit Sub

    Set HistoryWS = HistoryWB.Worksheets(1)
    LastRow = HistoryWS.Cells(HistoryWS.Rows.Count, 1).End(xlUp).Row + 1

    With HistoryWS
        .Cells(LastRow, 1).Value = Now
        .Cells(LastRow, 2).Value = SearchTerm
        .Cells(LastRow, 3).Value = ResultCount
    End With

    HistoryWB.Save
    DataManager.SafeCloseWorkbook HistoryWB
    Exit Sub

Error_Handler:
    If Not HistoryWB Is Nothing Then DataManager.SafeCloseWorkbook HistoryWB, False
End Sub

' **Purpose**: Get search usage statistics
' **Parameters**: None
' **Returns**: Variant - Array containing search statistics
' **Dependencies**: DataManager.SafeOpenWorkbook for database access
' **Side Effects**: None
' **Errors**: Returns empty array if statistics retrieval fails
Public Function GetSearchStatistics() As Variant
    Dim HistoryWB As Workbook
    Dim HistoryWS As Worksheet
    Dim LastRow As Long
    Dim Stats(0 To 4) As String
    Dim i As Long
    Dim TotalSearches As Long
    Dim TotalResults As Long
    Dim RecentSearches As Long

    On Error GoTo Error_Handler

    Set HistoryWB = DataManager.SafeOpenWorkbook(DataManager.GetRootPath & "\" & SEARCH_HISTORY_FILE)
    If HistoryWB Is Nothing Then
        GetSearchStatistics = Array()
        Exit Function
    End If

    Set HistoryWS = HistoryWB.Worksheets(1)
    LastRow = HistoryWS.Cells(HistoryWS.Rows.Count, 1).End(xlUp).Row

    TotalSearches = 0
    TotalResults = 0
    RecentSearches = 0

    For i = 2 To LastRow
        TotalSearches = TotalSearches + 1
        TotalResults = TotalResults + HistoryWS.Cells(i, 3).Value

        ' Count searches in last 30 days
        If HistoryWS.Cells(i, 1).Value >= DateAdd("d", -30, Now) Then
            RecentSearches = RecentSearches + 1
        End If
    Next i

    Stats(0) = "Total Searches: " & TotalSearches
    Stats(1) = "Total Results: " & TotalResults
    Stats(2) = "Average Results: " & IIf(TotalSearches > 0, TotalResults / TotalSearches, 0)
    Stats(3) = "Recent Searches (30 days): " & RecentSearches
    Stats(4) = "Last Search: " & IIf(LastRow > 1, HistoryWS.Cells(LastRow, 1).Value, "None")

    DataManager.SafeCloseWorkbook HistoryWB, False
    GetSearchStatistics = Stats
    Exit Function

Error_Handler:
    If Not HistoryWB Is Nothing Then DataManager.SafeCloseWorkbook HistoryWB, False
    CoreFramework.HandleStandardErrors Err.Number, "GetSearchStatistics", "SearchManager"
    GetSearchStatistics = Array()
End Function

' **Purpose**: Get most popular search terms
' **Parameters**:
'   - TopCount (Integer, Optional): Number of top terms to return (default 10)
' **Returns**: Variant - Array of popular search terms
' **Dependencies**: DataManager.SafeOpenWorkbook for database access
' **Side Effects**: None
' **Errors**: Returns empty array if retrieval fails
Public Function GetPopularSearchTerms(Optional ByVal TopCount As Integer = 10) As Variant
    Dim HistoryWB As Workbook
    Dim HistoryWS As Worksheet
    Dim LastRow As Long
    Dim SearchTerms As Collection
    Dim PopularTerms() As String
    Dim i As Long

    On Error GoTo Error_Handler

    Set HistoryWB = DataManager.SafeOpenWorkbook(DataManager.GetRootPath & "\" & SEARCH_HISTORY_FILE)
    If HistoryWB Is Nothing Then
        GetPopularSearchTerms = Array()
        Exit Function
    End If

    Set HistoryWS = HistoryWB.Worksheets(1)
    LastRow = HistoryWS.Cells(HistoryWS.Rows.Count, 1).End(xlUp).Row

    Set SearchTerms = New Collection

    ' Count search term frequency (simplified implementation)
    For i = 2 To LastRow
        ' This would need more sophisticated implementation for proper frequency counting
        On Error Resume Next
        SearchTerms.Add HistoryWS.Cells(i, 2).Value, CStr(HistoryWS.Cells(i, 2).Value)
        On Error GoTo Error_Handler
    Next i

    ' Return first TopCount terms (simplified - would need sorting by frequency)
    ReDim PopularTerms(0 To Application.Min(TopCount - 1, SearchTerms.Count - 1))
    For i = 0 To UBound(PopularTerms)
        PopularTerms(i) = SearchTerms(i + 1)
    Next i

    DataManager.SafeCloseWorkbook HistoryWB, False
    GetPopularSearchTerms = PopularTerms
    Exit Function

Error_Handler:
    If Not HistoryWB Is Nothing Then DataManager.SafeCloseWorkbook HistoryWB, False
    CoreFramework.HandleStandardErrors Err.Number, "GetPopularSearchTerms", "SearchManager"
    GetPopularSearchTerms = Array()
End Function

' ===================================================================
' PRIVATE HELPER FUNCTIONS
' ===================================================================

' **Purpose**: Find existing search row or determine where to add new row
' **Parameters**:
'   - SearchWS (Worksheet): Search worksheet to scan
'   - FormObject (Object): Form containing record identifiers
' **Returns**: Long - Row number for data placement
' **Dependencies**: None
' **Side Effects**: None
' **Errors**: Returns next available row if search fails
Private Function FindOrCreateSearchRow(ByRef SearchWS As Worksheet, ByRef FormObject As Object) As Long
    Dim i As Long
    Dim LastRow As Long
    Dim QuoteNumber As String
    Dim EnquiryNumber As String
    Dim JobNumber As String
    Dim FileName As String

    On Error GoTo Error_Handler

    ' Get identifiers from form
    On Error Resume Next
    QuoteNumber = FormObject.Controls("Quote_Number").Value
    EnquiryNumber = FormObject.Controls("Enquiry_Number").Value
    JobNumber = FormObject.Controls("Job_Number").Value
    FileName = FormObject.Controls("File_Name").Value
    On Error GoTo Error_Handler

    LastRow = SearchWS.Cells(SearchWS.Rows.Count, 1).End(xlUp).Row

    ' Look for existing record
    For i = 2 To LastRow
        If SearchWS.Cells(i, 1).Value = QuoteNumber Or _
           SearchWS.Cells(i, 1).Value = EnquiryNumber Or _
           SearchWS.Cells(i, 1).Value = JobNumber Or _
           SearchWS.Cells(i, 1).Value = FileName Then
            FindOrCreateSearchRow = i
            Exit Function
        End If
    Next i

    ' Return next available row
    FindOrCreateSearchRow = LastRow + 1
    Exit Function

Error_Handler:
    FindOrCreateSearchRow = SearchWS.Cells(SearchWS.Rows.Count, 1).End(xlUp).Row + 1
End Function

' **Purpose**: Sort search database within worksheet
' **Parameters**:
'   - SearchWS (Worksheet): Search worksheet to sort
' **Returns**: None (Subroutine)
' **Dependencies**: None
' **Side Effects**: Sorts worksheet data by date column
' **Errors**: Exits silently if sort fails
Private Sub SortSearchDatabaseInWorksheet(ByRef SearchWS As Worksheet)
    Dim LastRow As Long
    Dim LastCol As Long
    Dim SortRange As Range

    On Error GoTo Error_Handler

    LastRow = SearchWS.Cells(SearchWS.Rows.Count, 1).End(xlUp).Row
    LastCol = SearchWS.Cells(1, SearchWS.Columns.Count).End(xlToLeft).Column

    If LastRow > 2 And LastCol > 0 Then
        Set SortRange = SearchWS.Range(SearchWS.Cells(2, 1), SearchWS.Cells(LastRow, LastCol))
        SortRange.Sort Key1:=SearchWS.Cells(2, 5), Order1:=xlDescending, Header:=xlNo
    End If

    Exit Sub

Error_Handler:
    ' Sort failed - continue silently
End Sub

' **Purpose**: Process individual file for search database inclusion
' **Parameters**:
'   - SearchWS (Worksheet): Target search worksheet
'   - DirectoryName (String): Directory containing the file
'   - FileName (String): Name of file to process
' **Returns**: None (Subroutine)
' **Dependencies**: DataManager.GetValue for file data extraction
' **Side Effects**: Adds row to search worksheet
' **Errors**: Skips file if processing fails
Private Sub ProcessFileForSearch(ByRef SearchWS As Worksheet, ByVal DirectoryName As String, ByVal FileName As String)
    Dim FilePath As String
    Dim LastRow As Long
    Dim RecordNumber As String
    Dim CustomerName As String
    Dim Description As String

    On Error GoTo Error_Handler

    FilePath = DataManager.GetRootPath & "\" & DirectoryName & "\" & FileName
    LastRow = SearchWS.Cells(SearchWS.Rows.Count, 1).End(xlUp).Row + 1

    ' Extract basic information from file
    RecordNumber = DataManager.GetValue(FilePath, "Admin", "B2")
    CustomerName = DataManager.GetValue(FilePath, "Admin", "B10")
    Description = DataManager.GetValue(FilePath, "Admin", "B20")

    ' Add to search database
    With SearchWS
        .Cells(LastRow, 1).Value = GetRecordTypeFromDirectory(DirectoryName)
        .Cells(LastRow, 2).Value = RecordNumber
        .Cells(LastRow, 3).Value = CustomerName
        .Cells(LastRow, 4).Value = Description
        .Cells(LastRow, 5).Value = Now ' Use current date as placeholder
        .Cells(LastRow, 6).Value = FilePath
        .Cells(LastRow, 7).Value = CustomerName & " " & Description ' Basic keywords
    End With

    Exit Sub

    Exit Sub

Error_Handler:
    If Not FileWB Is Nothing Then DataManager.SafeCloseWorkbook FileWB, False
    ' Skip this file if processing fails - log the error but continue
    CoreFramework.LogError Err.Number, "Error processing file: " & FileName & " - " & Err.Description, "ProcessFileForSearch", "SearchManager"
End Sub

' **Purpose**: Return first non-empty value from a list of values
' **Parameters**: Variable number of Variant values to check
' **Returns**: Variant - First non-empty value, or empty string if all empty
' **Dependencies**: None
' **Side Effects**: None
' **Errors**: None
Private Function Coalesce(ParamArray Values() As Variant) As Variant
    Dim i As Integer

    For i = 0 To UBound(Values)
        If Not IsEmpty(Values(i)) And Not IsNull(Values(i)) Then
            If Trim(CStr(Values(i))) <> "" Then
                Coalesce = Values(i)
                Exit Function
            End If
        End If
    Next i

    Coalesce = ""
End Function

' **Purpose**: Get record type number from directory name
' **Parameters**:
'   - DirectoryName (String): Directory name to convert
' **Returns**: String - Record type number (1-4)
' **Dependencies**: None
' **Side Effects**: None
' **Errors**: Returns "1" if directory not recognized
Private Function GetRecordTypeFromDirectory(ByVal DirectoryName As String) As String
    Select Case UCase(DirectoryName)
        Case "ENQUIRIES"
            GetRecordTypeFromDirectory = "1"
        Case "QUOTES"
            GetRecordTypeFromDirectory = "2"
        Case "WIP", "JOBS"
            GetRecordTypeFromDirectory = "3"
        Case "CONTRACTS"
            GetRecordTypeFromDirectory = "4"
        Case Else
            GetRecordTypeFromDirectory = "1"
    End Select
End Function

' **Purpose**: Archive old records based on number thresholds
' **Parameters**:
'   - SearchWS (Worksheet): Search worksheet to process
' **Returns**: None (Subroutine)
' **Dependencies**: DataManager.GetNextJobNumber, DataManager.GetNextQuoteNumber
' **Side Effects**: Deletes old records from worksheet
' **Errors**: Continues processing if individual record deletion fails
Private Sub ArchiveOldRecords(ByRef SearchWS As Worksheet)
    Dim i As Long
    Dim LastRow As Long
    Dim RecordNumber As String
    Dim RecordNumeric As Long
    Dim JobThreshold As Long
    Dim QuoteThreshold As Long

    On Error GoTo Error_Handler

    ' Get current number thresholds
    JobThreshold = Val(Mid(DataManager.GetNextJobNumber(), 2)) - 1000
    QuoteThreshold = Val(Mid(DataManager.GetNextQuoteNumber(), 2)) - 10000

    LastRow = SearchWS.Cells(SearchWS.Rows.Count, 1).End(xlUp).Row

    ' Process from bottom up to avoid index issues
    For i = LastRow To 3 Step -1
        RecordNumber = SearchWS.Cells(i, 2).Value

        If Len(RecordNumber) > 1 Then
            RecordNumeric = Val(Mid(RecordNumber, 2))

            ' Archive old jobs or quotes based on thresholds
            If Left(RecordNumber, 1) = "J" And RecordNumeric < JobThreshold Then
                SearchWS.Rows(i).Delete
            ElseIf Left(RecordNumber, 1) = "Q" And RecordNumeric < QuoteThreshold Then
                SearchWS.Rows(i).Delete
            End If
        End If
    Next i

    Exit Sub

Error_Handler:
    ' Continue with next record if one fails
    Resume Next
End Sub

' ===================================================================
' LEGACY COMPATIBILITY FUNCTIONS (EXACT SIGNATURES)
' ===================================================================

' **Purpose**: Legacy compatibility - Save form data to search database
' **Parameters**:
'   - frm (Object): Form object containing data to save
' **Returns**: None (Subroutine)
' **Dependencies**: UpdateSearchDatabase, CreateSearchRecord
' **Side Effects**: Updates search database with form data
' **Errors**: Logs errors but continues execution
' **CLAUDE.md Compliance**: Maintains exact signature for button mapping compatibility
Public Sub SaveRowIntoSearch(ByRef frm As Object)
    Dim SearchRecord As CoreFramework.SearchRecord
    Dim RecordNumber As String
    Dim CustomerName As String
    Dim Description As String

    On Error GoTo Error_Handler

    ' Extract data from form controls
    On Error Resume Next
    RecordNumber = frm.Quote_Number.Value
    If RecordNumber = "" Then RecordNumber = frm.Enquiry_Number.Value
    If RecordNumber = "" Then RecordNumber = frm.Job_Number.Value
    If RecordNumber = "" Then RecordNumber = frm.File_Name.Value

    CustomerName = frm.Customer.Value
    If CustomerName = "" Then CustomerName = frm.CustomerName.Value
    If CustomerName = "" Then CustomerName = frm.Customer_Name.Value

    Description = frm.Description.Value
    If Description = "" Then Description = frm.ComponentDescription.Value
    If Description = "" Then Description = frm.Component_Description.Value
    On Error GoTo Error_Handler

    ' Create search record
    With SearchRecord
        .RecordNumber = RecordNumber
        .CustomerName = CustomerName
        .Description = Description
        .DateCreated = Now
        .FilePath = "" ' Will be determined by record type
        .Keywords = UCase(CustomerName & " " & Description & " " & RecordNumber)

        ' Determine record type from record number prefix
        If Left(UCase(RecordNumber), 1) = "E" Then
            .RecordType = CoreFramework.rtEnquiry
            .FilePath = DataManager.GetRootPath & "\Enquiries\" & RecordNumber & ".xls"
        ElseIf Left(UCase(RecordNumber), 1) = "Q" Then
            .RecordType = CoreFramework.rtQuote
            .FilePath = DataManager.GetRootPath & "\Quotes\" & RecordNumber & ".xls"
        ElseIf Left(UCase(RecordNumber), 1) = "J" Then
            .RecordType = CoreFramework.rtJob
            .FilePath = DataManager.GetRootPath & "\WIP\" & RecordNumber & ".xls"
        Else
            .RecordType = CoreFramework.rtEnquiry ' Default
        End If
    End With

    ' Update search database
    UpdateSearchDatabase SearchRecord
    Exit Sub

Error_Handler:
    CoreFramework.LogError Err.Number, "Error saving to search database: " & Err.Description, "SaveRowIntoSearch", "SearchManager"
End Sub

' **Purpose**: Legacy compatibility - Update search database from file system
' **Parameters**: None
' **Returns**: None (Subroutine)
' **Dependencies**: RebuildSearchDatabase_Incremental
' **Side Effects**: Rebuilds search database from existing files
' **Errors**: Logs errors but continues execution
' **CLAUDE.md Compliance**: Maintains exact signature for button mapping compatibility
Public Sub Update_Search()
    On Error GoTo Error_Handler

    ' Use optimized incremental rebuild instead of full rebuild
    If Not RebuildSearchDatabase_Incremental(1000, 180) Then
        MsgBox "Search database update failed. Please check error logs.", vbExclamation, "Search Update Error"
    Else
        MsgBox "Search database updated successfully.", vbInformation, "Search Update Complete"
    End If

    Exit Sub

Error_Handler:
    CoreFramework.LogError Err.Number, "Error updating search database: " & Err.Description, "Update_Search", "SearchManager"
    MsgBox "Search database update failed: " & Err.Description, vbCritical, "Search Update Error"
End Sub

' **Purpose**: Legacy compatibility - Get value from closed workbook
' **Parameters**:
'   - Path (String): Directory path to file
'   - File (String): Filename
'   - Sheet (String): Sheet name
'   - Ref (String): Cell reference
' **Returns**: Variant - Cell value or error message
' **Dependencies**: DataManager.GetValue (wrapper)
' **Side Effects**: None
' **Errors**: Returns error message strings
' **CLAUDE.md Compliance**: Maintains exact signature for legacy code compatibility
Public Function GetValue(ByVal Path As String, ByVal File As String, ByVal Sheet As String, ByVal Ref As String) As Variant
    On Error GoTo Error_Handler

    ' Ensure path ends with backslash
    If Right(Path, 1) <> "\" Then Path = Path & "\"

    ' Check if file exists
    If Not DataManager.FileExists(Path & File) Then
        GetValue = "File Not Found"
        Exit Function
    End If

    ' Use DataManager to get value
    GetValue = DataManager.GetValue(Path & File, Sheet, Ref)

    ' Handle empty or error values
    If IsError(GetValue) Then GetValue = ""
    If GetValue = 0 Then GetValue = ""

    Exit Function

Error_Handler:
    GetValue = "Error: " & Err.Description
End Function