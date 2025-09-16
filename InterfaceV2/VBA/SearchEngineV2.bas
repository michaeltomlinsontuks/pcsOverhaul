Attribute VB_Name = "SearchEngineV2"
Option Explicit

Private Type SearchResult
    FilePath As String
    CustomerName As String
    ComponentCode As String
    ComponentDesc As String
    Status As String
    MatchScore As Integer
    FileType As String
    ModDate As Date
End Type

Private Const MAX_RESULTS = 100

Public Function ExecuteSmartSearch(searchTerm As String) As SearchResult()
    Dim results() As SearchResult
    Dim resultCount As Long
    Dim allFiles() As String
    Dim i As Long
    Dim tempResult As SearchResult

    If Len(Trim(searchTerm)) = 0 Then
        ExecuteSmartSearch = results
        Exit Function
    End If

    ReDim results(1 To MAX_RESULTS)
    resultCount = 0

    allFiles = BuildFileList()

    For i = LBound(allFiles) To UBound(allFiles)
        If allFiles(i) <> "" Then
            tempResult = SearchFile(allFiles(i), searchTerm)
            If tempResult.MatchScore > 0 Then
                resultCount = resultCount + 1
                If resultCount <= MAX_RESULTS Then
                    results(resultCount) = tempResult
                End If
            End If
        End If

        If i Mod 50 = 0 Then
            DoEvents
        End If
    Next i

    If resultCount > 0 Then
        ReDim Preserve results(1 To resultCount)
        results = RankResults(results)
    Else
        ReDim results(1 To 0)
    End If

    ExecuteSmartSearch = results
End Function

Private Function SearchFile(filePath As String, searchTerm As String) As SearchResult
    Dim result As SearchResult
    Dim fileName As String
    Dim cachedValue As String
    Dim score As Integer

    fileName = GetFileNameFromPath(filePath)
    result.FilePath = filePath
    result.FileType = GetFileTypeFromPath(filePath)
    result.ModDate = GetFileModDate(filePath)

    score = 0

    If InStr(1, fileName, searchTerm, vbTextCompare) > 0 Then
        score = score + 50
    End If

    cachedValue = CacheManager.GetCachedValue(filePath, "CustomerName")
    If cachedValue <> "" Then
        result.CustomerName = cachedValue
        If InStr(1, cachedValue, searchTerm, vbTextCompare) > 0 Then
            score = score + 40
        End If
    End If

    cachedValue = CacheManager.GetCachedValue(filePath, "ComponentCode")
    If cachedValue <> "" Then
        result.ComponentCode = cachedValue
        If InStr(1, cachedValue, searchTerm, vbTextCompare) > 0 Then
            score = score + 45
        End If
    End If

    cachedValue = CacheManager.GetCachedValue(filePath, "ComponentDesc")
    If cachedValue <> "" Then
        result.ComponentDesc = cachedValue
        If InStr(1, cachedValue, searchTerm, vbTextCompare) > 0 Then
            score = score + 35
        End If
    End If

    cachedValue = CacheManager.GetCachedValue(filePath, "Status")
    If cachedValue <> "" Then
        result.Status = cachedValue
        If InStr(1, cachedValue, searchTerm, vbTextCompare) > 0 Then
            score = score + 20
        End If
    End If

    If score = 0 And result.CustomerName = "" Then
        score = SearchFileContent(filePath, searchTerm, result)
    End If

    If result.FileType = "WIP" Then score = score + 10
    If result.FileType = "Quote" Then score = score + 8
    If result.FileType = "Enquiry" Then score = score + 5

    If DateDiff("d", result.ModDate, Now) < 30 Then
        score = score + 5
    End If

    result.MatchScore = score
    SearchFile = result
End Function

Private Function SearchFileContent(filePath As String, searchTerm As String, ByRef result As SearchResult) As Integer
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim score As Integer
    Dim customerCell As Range
    Dim codeCell As Range
    Dim descCell As Range

    On Error GoTo ErrorHandler

    score = 0
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set wb = Application.Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Set ws = wb.Worksheets(1)

    Set customerCell = ws.Range("C4")
    If Not customerCell Is Nothing Then
        result.CustomerName = CStr(customerCell.Value)
        If InStr(1, result.CustomerName, searchTerm, vbTextCompare) > 0 Then
            score = score + 40
        End If
        CacheManager.CacheFileMetadata filePath, result.CustomerName, "", "", ""
    End If

    Set codeCell = ws.Range("C6")
    If Not codeCell Is Nothing Then
        result.ComponentCode = CStr(codeCell.Value)
        If InStr(1, result.ComponentCode, searchTerm, vbTextCompare) > 0 Then
            score = score + 45
        End If
    End If

    Set descCell = ws.Range("C7")
    If Not descCell Is Nothing Then
        result.ComponentDesc = CStr(descCell.Value)
        If InStr(1, result.ComponentDesc, searchTerm, vbTextCompare) > 0 Then
            score = score + 35
        End If
    End If

    CacheManager.CacheFileMetadata filePath, result.CustomerName, result.ComponentCode, result.ComponentDesc, result.Status

    wb.Close False

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    SearchFileContent = score
    Exit Function

ErrorHandler:
    If Not wb Is Nothing Then wb.Close False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    SearchFileContent = 0
End Function

Private Function RankResults(results() As SearchResult) As SearchResult()
    Dim i As Long, j As Long
    Dim temp As SearchResult

    For i = LBound(results) To UBound(results) - 1
        For j = i + 1 To UBound(results)
            If results(i).MatchScore < results(j).MatchScore Then
                temp = results(i)
                results(i) = results(j)
                results(j) = temp
            ElseIf results(i).MatchScore = results(j).MatchScore Then
                If results(i).ModDate < results(j).ModDate Then
                    temp = results(i)
                    results(i) = results(j)
                    results(j) = temp
                End If
            End If
        Next j
    Next i

    RankResults = results
End Function

Private Function BuildFileList() As String()
    Dim fileList() As String
    Dim fileCount As Long
    Dim tempArray() As String

    fileCount = 0
    ReDim fileList(1 To 1000)

    tempArray = GetFilesFromDirectory(Application.ActiveWorkbook.Path & "\Enquiries\", "*.xls")
    fileCount = AppendFiles(fileList, tempArray, fileCount)

    tempArray = GetFilesFromDirectory(Application.ActiveWorkbook.Path & "\Quotes\", "*.xls")
    fileCount = AppendFiles(fileList, tempArray, fileCount)

    tempArray = GetFilesFromDirectory(Application.ActiveWorkbook.Path & "\WIP\", "*.xls")
    fileCount = AppendFiles(fileList, tempArray, fileCount)

    tempArray = GetFilesFromDirectory(Application.ActiveWorkbook.Path & "\Archive\", "*.xls")
    fileCount = AppendFiles(fileList, tempArray, fileCount)

    If fileCount > 0 Then
        ReDim Preserve fileList(1 To fileCount)
    Else
        ReDim fileList(1 To 0)
    End If

    BuildFileList = fileList
End Function

Private Function AppendFiles(ByRef fileList() As String, newFiles() As String, startIndex As Long) As Long
    Dim i As Long
    Dim currentIndex As Long

    currentIndex = startIndex

    For i = LBound(newFiles) To UBound(newFiles)
        If newFiles(i) <> "" Then
            currentIndex = currentIndex + 1
            If currentIndex <= UBound(fileList) Then
                fileList(currentIndex) = newFiles(i)
            End If
        End If
    Next i

    AppendFiles = currentIndex
End Function

Private Function GetFilesFromDirectory(dirPath As String, pattern As String) As String()
    Dim files() As String
    Dim fileName As String
    Dim fileCount As Long

    fileCount = 0
    ReDim files(1 To 100)

    fileName = Dir(dirPath & pattern)
    Do While fileName <> ""
        fileCount = fileCount + 1
        If fileCount <= UBound(files) Then
            files(fileCount) = dirPath & fileName
        End If
        fileName = Dir
    Loop

    If fileCount > 0 Then
        ReDim Preserve files(1 To fileCount)
    Else
        ReDim files(1 To 0)
    End If

    GetFilesFromDirectory = files
End Function

Private Function GetFileNameFromPath(filePath As String) As String
    Dim lastSlash As Long
    lastSlash = InStrRev(filePath, "\")
    If lastSlash > 0 Then
        GetFileNameFromPath = Mid(filePath, lastSlash + 1)
    Else
        GetFileNameFromPath = filePath
    End If
End Function

Private Function GetFileTypeFromPath(filePath As String) As String
    If InStr(filePath, "\WIP\") > 0 Then
        GetFileTypeFromPath = "WIP"
    ElseIf InStr(filePath, "\Quotes\") > 0 Then
        GetFileTypeFromPath = "Quote"
    ElseIf InStr(filePath, "\Enquiries\") > 0 Then
        GetFileTypeFromPath = "Enquiry"
    ElseIf InStr(filePath, "\Archive\") > 0 Then
        GetFileTypeFromPath = "Archive"
    Else
        GetFileTypeFromPath = "Other"
    End If
End Function

Private Function GetFileModDate(filePath As String) As Date
    On Error GoTo ErrorHandler
    GetFileModDate = FileDateTime(filePath)
    Exit Function

ErrorHandler:
    GetFileModDate = Now - 365
End Function