Attribute VB_Name = "SearchService"
Option Explicit

Private Const SEARCH_FILE As String = "Search.xls"
Private Const SEARCH_HISTORY_FILE As String = "Search History.xls"

Public Function UpdateSearchDatabase(ByRef Record As SearchRecord) As Boolean
    Dim SearchWB As Workbook
    Dim SearchWS As Worksheet
    Dim LastRow As Long

    On Error GoTo Error_Handler

    Set SearchWB = FileManager.SafeOpenWorkbook(FileManager.GetRootPath & "\" & SEARCH_FILE)
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
    FileManager.SafeCloseWorkbook SearchWB

    UpdateSearchDatabase = True
    Exit Function

Error_Handler:
    If Not SearchWB Is Nothing Then FileManager.SafeCloseWorkbook SearchWB, False
    ErrorHandler.HandleStandardErrors Err.Number, "UpdateSearchDatabase", "SearchService"
    UpdateSearchDatabase = False
End Function

Public Function SearchRecords(ByVal SearchTerm As String, Optional ByVal RecordTypeFilter As RecordType = 0) As Variant
    Dim SearchWB As Workbook
    Dim SearchWS As Worksheet
    Dim LastRow As Long
    Dim i As Long
    Dim Results() As SearchRecord
    Dim ResultCount As Integer
    Dim CurrentRecord As SearchRecord

    On Error GoTo Error_Handler

    Set SearchWB = FileManager.SafeOpenWorkbook(FileManager.GetRootPath & "\" & SEARCH_FILE)
    If SearchWB Is Nothing Then
        SearchRecords = Array()
        Exit Function
    End If

    Set SearchWS = SearchWB.Worksheets(1)
    LastRow = SearchWS.Cells(SearchWS.Rows.Count, 1).End(xlUp).Row

    SearchTerm = UCase(SearchTerm)
    ResultCount = 0

    For i = 2 To LastRow
        With SearchWS
            If RecordTypeFilter = 0 Or .Cells(i, 1).Value = RecordTypeFilter Then
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

    FileManager.SafeCloseWorkbook SearchWB, False

    If ResultCount > 0 Then
        SearchRecords = Results
    Else
        SearchRecords = Array()
    End If

    LogSearchHistory SearchTerm, ResultCount
    Exit Function

Error_Handler:
    If Not SearchWB Is Nothing Then FileManager.SafeCloseWorkbook SearchWB, False
    ErrorHandler.HandleStandardErrors Err.Number, "SearchRecords", "SearchService"
    SearchRecords = Array()
End Function

Public Function DeleteSearchRecord(ByVal RecordNumber As String) As Boolean
    Dim SearchWB As Workbook
    Dim SearchWS As Worksheet
    Dim LastRow As Long
    Dim i As Long

    On Error GoTo Error_Handler

    Set SearchWB = FileManager.SafeOpenWorkbook(FileManager.GetRootPath & "\" & SEARCH_FILE)
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
            FileManager.SafeCloseWorkbook SearchWB
            DeleteSearchRecord = True
            Exit Function
        End If
    Next i

    FileManager.SafeCloseWorkbook SearchWB, False
    DeleteSearchRecord = False
    Exit Function

Error_Handler:
    If Not SearchWB Is Nothing Then FileManager.SafeCloseWorkbook SearchWB, False
    ErrorHandler.HandleStandardErrors Err.Number, "DeleteSearchRecord", "SearchService"
    DeleteSearchRecord = False
End Function

Public Function SortSearchDatabase() As Boolean
    Dim SearchWB As Workbook
    Dim SearchWS As Worksheet
    Dim LastRow As Long
    Dim SortRange As Range

    On Error GoTo Error_Handler

    Set SearchWB = FileManager.SafeOpenWorkbook(FileManager.GetRootPath & "\" & SEARCH_FILE)
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
    FileManager.SafeCloseWorkbook SearchWB
    SortSearchDatabase = True
    Exit Function

Error_Handler:
    If Not SearchWB Is Nothing Then FileManager.SafeCloseWorkbook SearchWB, False
    ErrorHandler.HandleStandardErrors Err.Number, "SortSearchDatabase", "SearchService"
    SortSearchDatabase = False
End Function

Private Sub LogSearchHistory(ByVal SearchTerm As String, ByVal ResultCount As Integer)
    Dim HistoryWB As Workbook
    Dim HistoryWS As Worksheet
    Dim LastRow As Long

    On Error GoTo Error_Handler

    Set HistoryWB = FileManager.SafeOpenWorkbook(FileManager.GetRootPath & "\" & SEARCH_HISTORY_FILE)
    If HistoryWB Is Nothing Then Exit Sub

    Set HistoryWS = HistoryWB.Worksheets(1)
    LastRow = HistoryWS.Cells(HistoryWS.Rows.Count, 1).End(xlUp).Row + 1

    With HistoryWS
        .Cells(LastRow, 1).Value = Now
        .Cells(LastRow, 2).Value = SearchTerm
        .Cells(LastRow, 3).Value = ResultCount
    End With

    HistoryWB.Save
    FileManager.SafeCloseWorkbook HistoryWB
    Exit Sub

Error_Handler:
    If Not HistoryWB Is Nothing Then FileManager.SafeCloseWorkbook HistoryWB, False
End Sub

Public Function CreateSearchRecord(ByVal RecType As RecordType, ByVal Number As String, ByVal Customer As String, ByVal Description As String, ByVal FilePath As String, Optional ByVal Keywords As String = "") As SearchRecord
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

Public Function SaveSearchHistory(ByVal SearchTerm As String) As Boolean
    Dim HistoryWB As Workbook
    Dim HistoryWS As Worksheet
    Dim LastRow As Long
    Dim HistoryPath As String

    On Error GoTo Error_Handler

    HistoryPath = FileManager.GetRootPath & "\" & SEARCH_HISTORY_FILE

    Set HistoryWB = FileManager.SafeOpenWorkbook(HistoryPath)
    If HistoryWB Is Nothing Then
        SaveSearchHistory = False
        Exit Function
    End If

    Set HistoryWS = HistoryWB.Worksheets(1)
    LastRow = HistoryWS.Cells(HistoryWS.Rows.Count, 1).End(xlUp).Row

    If LastRow < 1 Then LastRow = 1

    With HistoryWS
        .Cells(LastRow + 1, 1).Value = SearchTerm
        .Cells(LastRow + 1, 2).Value = Now
    End With

    HistoryWB.Save
    FileManager.SafeCloseWorkbook HistoryWB
    SaveSearchHistory = True
    Exit Function

Error_Handler:
    If Not HistoryWB Is Nothing Then FileManager.SafeCloseWorkbook HistoryWB, False
    ErrorHandler.HandleStandardErrors Err.Number, "SaveSearchHistory", "SearchService"
    SaveSearchHistory = False
End Function

Public Function GetSearchHistory() As Variant
    Dim HistoryWB As Workbook
    Dim HistoryWS As Worksheet
    Dim LastRow As Long
    Dim i As Long
    Dim History() As String
    Dim HistoryCount As Integer
    Dim HistoryPath As String

    On Error GoTo Error_Handler

    HistoryPath = FileManager.GetRootPath & "\" & SEARCH_HISTORY_FILE

    Set HistoryWB = FileManager.SafeOpenWorkbook(HistoryPath)
    If HistoryWB Is Nothing Then
        GetSearchHistory = Array()
        Exit Function
    End If

    Set HistoryWS = HistoryWB.Worksheets(1)
    LastRow = HistoryWS.Cells(HistoryWS.Rows.Count, 1).End(xlUp).Row

    If LastRow < 2 Then
        FileManager.SafeCloseWorkbook HistoryWB, False
        GetSearchHistory = Array()
        Exit Function
    End If

    HistoryCount = 0
    ReDim History(0 To LastRow - 2)

    For i = LastRow To 2 Step -1
        If i - 2 <= UBound(History) And HistoryCount < 10 Then
            History(HistoryCount) = HistoryWS.Cells(i, 1).Value
            HistoryCount = HistoryCount + 1
        End If
        If HistoryCount >= 10 Then Exit For
    Next i

    If HistoryCount > 0 Then
        ReDim Preserve History(0 To HistoryCount - 1)
        GetSearchHistory = History
    Else
        GetSearchHistory = Array()
    End If

    FileManager.SafeCloseWorkbook HistoryWB, False
    Exit Function

Error_Handler:
    If Not HistoryWB Is Nothing Then FileManager.SafeCloseWorkbook HistoryWB, False
    ErrorHandler.HandleStandardErrors Err.Number, "GetSearchHistory", "SearchService"
    GetSearchHistory = Array()
End Function