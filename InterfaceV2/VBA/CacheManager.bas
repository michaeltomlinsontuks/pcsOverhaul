Attribute VB_Name = "CacheManager"
Option Explicit

Private metadataCache As Object
Private Const MAX_CACHE_ENTRIES = 500
Private Const CACHE_FILE_PATH = "SearchCache.txt"
Private cacheInitialized As Boolean

Public Sub InitializeCache()
    If metadataCache Is Nothing Then
        Set metadataCache = CreateObject("Scripting.Dictionary")
        cacheInitialized = True
        LoadCacheFromFile
    End If
End Sub

Public Function GetCachedValue(filePath As String, fieldName As String) As String
    If Not cacheInitialized Then InitializeCache

    Dim cacheKey As String
    Dim cacheValue As String
    Dim fields() As String

    cacheKey = LCase(filePath)

    If metadataCache.Exists(cacheKey) Then
        cacheValue = metadataCache(cacheKey)
        fields = Split(cacheValue, "|")

        If UBound(fields) >= 4 Then
            Select Case LCase(fieldName)
                Case "customername"
                    GetCachedValue = fields(0)
                Case "componentcode"
                    GetCachedValue = fields(1)
                Case "componentdesc"
                    GetCachedValue = fields(2)
                Case "status"
                    GetCachedValue = fields(3)
                Case "moddate"
                    GetCachedValue = fields(4)
                Case Else
                    GetCachedValue = ""
            End Select
        End If
    Else
        GetCachedValue = ""
    End If
End Function

Public Sub CacheFileMetadata(filePath As String, customer As String, component As String, description As String, status As String)
    If Not cacheInitialized Then InitializeCache

    Dim cacheKey As String
    Dim cacheValue As String
    Dim modDate As String

    cacheKey = LCase(filePath)
    modDate = CStr(GetFileModificationDate(filePath))

    cacheValue = customer & "|" & component & "|" & description & "|" & status & "|" & modDate

    If metadataCache.Count >= MAX_CACHE_ENTRIES And Not metadataCache.Exists(cacheKey) Then
        EvictOldestEntry
    End If

    metadataCache(cacheKey) = cacheValue
End Sub

Public Sub LoadCacheFromFile()
    If Not cacheInitialized Then InitializeCache

    Dim filePath As String
    Dim fileNum As Integer
    Dim lineText As String
    Dim parts() As String
    Dim cacheKey As String
    Dim cacheValue As String

    filePath = Application.ActiveWorkbook.Path & "\" & CACHE_FILE_PATH

    If Dir(filePath) = "" Then Exit Sub

    On Error GoTo ErrorHandler

    fileNum = FreeFile
    Open filePath For Input As #fileNum

    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText
        If Len(Trim(lineText)) > 0 And InStr(lineText, "=") > 0 Then
            parts = Split(lineText, "=", 2)
            If UBound(parts) = 1 Then
                cacheKey = parts(0)
                cacheValue = parts(1)

                If IsValidCacheEntry(cacheKey, cacheValue) Then
                    metadataCache(cacheKey) = cacheValue
                End If
            End If
        End If
    Loop

    Close #fileNum
    Exit Sub

ErrorHandler:
    If fileNum > 0 Then Close #fileNum
End Sub

Public Sub SaveCacheToFile()
    If Not cacheInitialized Then Exit Sub

    Dim filePath As String
    Dim fileNum As Integer
    Dim key As Variant

    filePath = Application.ActiveWorkbook.Path & "\" & CACHE_FILE_PATH

    On Error GoTo ErrorHandler

    fileNum = FreeFile
    Open filePath For Output As #fileNum

    Print #fileNum, "# PCS Interface V2 Search Cache"
    Print #fileNum, "# Generated: " & Format(Now, "yyyy-mm-dd hh:mm:ss")
    Print #fileNum, "# Format: filepath=customer|component|description|status|moddate"
    Print #fileNum, ""

    For Each key In metadataCache.Keys
        Print #fileNum, key & "=" & metadataCache(key)
    Next key

    Close #fileNum
    Exit Sub

ErrorHandler:
    If fileNum > 0 Then Close #fileNum
End Sub

Public Sub ClearCache()
    If Not cacheInitialized Then InitializeCache
    metadataCache.RemoveAll
End Sub

Public Function GetCacheStats() As String
    If Not cacheInitialized Then InitializeCache

    Dim stats As String
    stats = "Cache Entries: " & metadataCache.Count & "/" & MAX_CACHE_ENTRIES
    stats = stats & vbCrLf & "Cache File: " & CACHE_FILE_PATH

    Dim filePath As String
    filePath = Application.ActiveWorkbook.Path & "\" & CACHE_FILE_PATH
    If Dir(filePath) <> "" Then
        stats = stats & vbCrLf & "File Size: " & FileLen(filePath) & " bytes"
        stats = stats & vbCrLf & "File Modified: " & Format(FileDateTime(filePath), "yyyy-mm-dd hh:mm:ss")
    Else
        stats = stats & vbCrLf & "Cache file not found"
    End If

    GetCacheStats = stats
End Function

Public Sub BuildCacheInBackground()
    If Not cacheInitialized Then InitializeCache

    Dim directories() As String
    Dim i As Long
    Dim fileList() As String
    Dim j As Long

    ReDim directories(1 To 4)
    directories(1) = Application.ActiveWorkbook.Path & "\Enquiries\"
    directories(2) = Application.ActiveWorkbook.Path & "\Quotes\"
    directories(3) = Application.ActiveWorkbook.Path & "\WIP\"
    directories(4) = Application.ActiveWorkbook.Path & "\Archive\"

    For i = 1 To UBound(directories)
        If Dir(directories(i), vbDirectory) <> "" Then
            fileList = GetFilesFromDirectory(directories(i))

            For j = LBound(fileList) To UBound(fileList)
                If fileList(j) <> "" And Not IsCached(fileList(j)) Then
                    CacheFileFromDisk fileList(j)

                    If j Mod 10 = 0 Then
                        DoEvents
                    End If
                End If
            Next j
        End If
    Next i

    SaveCacheToFile
End Sub

Private Sub EvictOldestEntry()
    Dim key As Variant
    Dim oldestKey As String
    Dim oldestDate As Date
    Dim currentDate As Date
    Dim cacheValue As String
    Dim fields() As String

    oldestDate = Now

    For Each key In metadataCache.Keys
        cacheValue = metadataCache(key)
        fields = Split(cacheValue, "|")

        If UBound(fields) >= 4 Then
            On Error Resume Next
            currentDate = CDate(fields(4))
            If Err.Number = 0 And currentDate < oldestDate Then
                oldestDate = currentDate
                oldestKey = key
            End If
            On Error GoTo 0
        End If
    Next key

    If oldestKey <> "" Then
        metadataCache.Remove oldestKey
    End If
End Sub

Private Function IsValidCacheEntry(cacheKey As String, cacheValue As String) As Boolean
    Dim fields() As String
    Dim filePath As String
    Dim cachedModDate As Date
    Dim actualModDate As Date

    IsValidCacheEntry = False

    fields = Split(cacheValue, "|")
    If UBound(fields) < 4 Then Exit Function

    filePath = cacheKey

    If Dir(filePath) = "" Then Exit Function

    On Error Resume Next
    cachedModDate = CDate(fields(4))
    actualModDate = FileDateTime(filePath)

    If Err.Number = 0 Then
        IsValidCacheEntry = (cachedModDate = actualModDate)
    End If
    On Error GoTo 0
End Function

Private Function IsCached(filePath As String) As Boolean
    If Not cacheInitialized Then InitializeCache
    IsCached = metadataCache.Exists(LCase(filePath))
End Function

Private Sub CacheFileFromDisk(filePath As String)
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim customer As String
    Dim component As String
    Dim description As String
    Dim status As String

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set wb = Application.Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Set ws = wb.Worksheets(1)

    customer = CStr(ws.Range("C4").Value)
    component = CStr(ws.Range("C6").Value)
    description = CStr(ws.Range("C7").Value)

    If InStr(filePath, "\WIP\") > 0 Then
        status = "WIP"
    ElseIf InStr(filePath, "\Quotes\") > 0 Then
        status = "Quote"
    ElseIf InStr(filePath, "\Enquiries\") > 0 Then
        status = "Enquiry"
    ElseIf InStr(filePath, "\Archive\") > 0 Then
        status = "Archive"
    Else
        status = "Unknown"
    End If

    CacheFileMetadata filePath, customer, component, description, status

    wb.Close False

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub

ErrorHandler:
    If Not wb Is Nothing Then wb.Close False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Private Function GetFilesFromDirectory(dirPath As String) As String()
    Dim files() As String
    Dim fileName As String
    Dim fileCount As Long

    fileCount = 0
    ReDim files(1 To 100)

    fileName = Dir(dirPath & "*.xls")
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

Private Function GetFileModificationDate(filePath As String) As Date
    On Error GoTo ErrorHandler
    GetFileModificationDate = FileDateTime(filePath)
    Exit Function

ErrorHandler:
    GetFileModificationDate = Now
End Function