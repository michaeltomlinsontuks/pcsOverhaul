Attribute VB_Name = "FileUtilities"
Option Explicit

Private Type FileInfo
    FullPath As String
    ModDate As Date
    Size As Long
    IsValid As Boolean
End Type

Public Function GetValueFast(filePath As String, sheetName As String, cellRef As String) As Variant
    Dim cachedValue As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim result As Variant

    cachedValue = CacheManager.GetCachedValue(filePath, "cell_" & sheetName & "_" & cellRef)
    If cachedValue <> "" Then
        GetValueFast = cachedValue
        Exit Function
    End If

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    Set wb = Application.Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False, IgnoreReadOnlyRecommended:=True)

    If sheetName = "" Then
        Set ws = wb.Worksheets(1)
    Else
        Set ws = wb.Worksheets(sheetName)
    End If

    result = ws.Range(cellRef).Value

    CacheManager.CacheFileMetadata filePath & "_cell_" & sheetName & "_" & cellRef, CStr(result), "", "", ""

    wb.Close False

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic

    GetValueFast = result
    Exit Function

ErrorHandler:
    If Not wb Is Nothing Then wb.Close False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    GetValueFast = ""
End Function

Public Function GetMultipleValuesFast(filePath As String, sheetName As String, cellRefs() As String) As Variant()
    Dim values() As Variant
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim i As Long

    ReDim values(LBound(cellRefs) To UBound(cellRefs))

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    Set wb = Application.Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False, IgnoreReadOnlyRecommended:=True)

    If sheetName = "" Then
        Set ws = wb.Worksheets(1)
    Else
        Set ws = wb.Worksheets(sheetName)
    End If

    For i = LBound(cellRefs) To UBound(cellRefs)
        values(i) = ws.Range(cellRefs(i)).Value
    Next i

    wb.Close False

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic

    GetMultipleValuesFast = values
    Exit Function

ErrorHandler:
    If Not wb Is Nothing Then wb.Close False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic

    For i = LBound(values) To UBound(values)
        values(i) = ""
    Next i
    GetMultipleValuesFast = values
End Function

Public Function BuildFileList() As String()
    Static lastBuildTime As Date
    Static cachedFileList() As String
    Static cacheValid As Boolean

    Dim currentTime As Date
    Dim directories() As String
    Dim allFiles() As String
    Dim fileCount As Long
    Dim i As Long

    currentTime = Now

    If cacheValid And DateDiff("n", lastBuildTime, currentTime) < 5 Then
        BuildFileList = cachedFileList
        Exit Function
    End If

    ReDim directories(1 To 4)
    directories(1) = Application.ActiveWorkbook.Path & "\Enquiries\"
    directories(2) = Application.ActiveWorkbook.Path & "\Quotes\"
    directories(3) = Application.ActiveWorkbook.Path & "\WIP\"
    directories(4) = Application.ActiveWorkbook.Path & "\Archive\"

    ReDim allFiles(1 To 1000)
    fileCount = 0

    For i = 1 To UBound(directories)
        If Dir(directories(i), vbDirectory) <> "" Then
            fileCount = AddDirectoryFiles(allFiles, directories(i), fileCount)
        End If
    Next i

    If fileCount > 0 Then
        ReDim Preserve allFiles(1 To fileCount)
        allFiles = SortFilesByDate(allFiles)
    Else
        ReDim allFiles(1 To 0)
    End If

    cachedFileList = allFiles
    lastBuildTime = currentTime
    cacheValid = True

    BuildFileList = allFiles
End Function

Public Function CheckFileExists(filePath As String) As Boolean
    CheckFileExists = (Dir(filePath) <> "")
End Function

Public Function GetFileInfo(filePath As String) As FileInfo
    Dim info As FileInfo

    info.FullPath = filePath
    info.IsValid = CheckFileExists(filePath)

    If info.IsValid Then
        On Error Resume Next
        info.ModDate = FileDateTime(filePath)
        info.Size = FileLen(filePath)
        If Err.Number <> 0 Then
            info.IsValid = False
        End If
        On Error GoTo 0
    End If

    GetFileInfo = info
End Function

Public Function CreateBackupFile(originalPath As String) As String
    Dim backupPath As String
    Dim fileName As String
    Dim extension As String
    Dim baseName As String
    Dim timestamp As String
    Dim dotPos As Long

    fileName = GetFileNameFromPath(originalPath)
    dotPos = InStrRev(fileName, ".")

    If dotPos > 0 Then
        baseName = Left(fileName, dotPos - 1)
        extension = Mid(fileName, dotPos)
    Else
        baseName = fileName
        extension = ""
    End If

    timestamp = Format(Now, "yyyymmdd_hhmmss")
    backupPath = GetDirectoryFromPath(originalPath) & baseName & "_backup_" & timestamp & extension

    On Error GoTo ErrorHandler

    FileCopy originalPath, backupPath
    CreateBackupFile = backupPath
    Exit Function

ErrorHandler:
    CreateBackupFile = ""
End Function

Public Function ValidateFileIntegrity(filePath As String) As Boolean
    Dim wb As Workbook
    Dim isValid As Boolean

    isValid = False

    On Error GoTo ErrorHandler

    Application.DisplayAlerts = False
    Set wb = Application.Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)

    If Not wb Is Nothing Then
        If wb.Worksheets.Count > 0 Then
            isValid = True
        End If
        wb.Close False
    End If

    Application.DisplayAlerts = True
    ValidateFileIntegrity = isValid
    Exit Function

ErrorHandler:
    If Not wb Is Nothing Then wb.Close False
    Application.DisplayAlerts = True
    ValidateFileIntegrity = False
End Function

Public Sub OptimizeFileAccess()
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
End Sub

Public Sub RestoreFileAccess()
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With
End Sub

Private Function AddDirectoryFiles(ByRef allFiles() As String, dirPath As String, startIndex As Long) As Long
    Dim fileName As String
    Dim currentIndex As Long
    Dim fullPath As String

    currentIndex = startIndex
    fileName = Dir(dirPath & "*.xls")

    Do While fileName <> ""
        currentIndex = currentIndex + 1
        If currentIndex <= UBound(allFiles) Then
            fullPath = dirPath & fileName
            allFiles(currentIndex) = fullPath
        End If
        fileName = Dir
    Loop

    AddDirectoryFiles = currentIndex
End Function

Private Function SortFilesByDate(files() As String) As String()
    Dim i As Long, j As Long
    Dim temp As String
    Dim date1 As Date, date2 As Date

    For i = LBound(files) To UBound(files) - 1
        For j = i + 1 To UBound(files)
            On Error Resume Next
            date1 = FileDateTime(files(i))
            date2 = FileDateTime(files(j))
            On Error GoTo 0

            If date1 < date2 Then
                temp = files(i)
                files(i) = files(j)
                files(j) = temp
            End If
        Next j
    Next i

    SortFilesByDate = files
End Function

Private Function GetFileNameFromPath(fullPath As String) As String
    Dim lastSlash As Long
    lastSlash = InStrRev(fullPath, "\")
    If lastSlash > 0 Then
        GetFileNameFromPath = Mid(fullPath, lastSlash + 1)
    Else
        GetFileNameFromPath = fullPath
    End If
End Function

Private Function GetDirectoryFromPath(fullPath As String) As String
    Dim lastSlash As Long
    lastSlash = InStrRev(fullPath, "\")
    If lastSlash > 0 Then
        GetDirectoryFromPath = Left(fullPath, lastSlash)
    Else
        GetDirectoryFromPath = ""
    End If
End Function

Public Function GetFileTypeFromPath(filePath As String) As String
    If InStr(filePath, "\WIP\") > 0 Then
        GetFileTypeFromPath = "WIP"
    ElseIf InStr(filePath, "\Quotes\") > 0 Then
        GetFileTypeFromPath = "Quote"
    ElseIf InStr(filePath, "\Enquiries\") > 0 Then
        GetFileTypeFromPath = "Enquiry"
    ElseIf InStr(filePath, "\Archive\") > 0 Then
        GetFileTypeFromPath = "Archive"
    ElseIf InStr(filePath, "\Contracts\") > 0 Then
        GetFileTypeFromPath = "Contract"
    ElseIf InStr(filePath, "\Customers\") > 0 Then
        GetFileTypeFromPath = "Customer"
    Else
        GetFileTypeFromPath = "Other"
    End If
End Function

Public Function CleanFileName(fileName As String) As String
    Dim cleanName As String
    Dim invalidChars As String
    Dim i As Long

    invalidChars = "\/:*?""<>|"
    cleanName = fileName

    For i = 1 To Len(invalidChars)
        cleanName = Replace(cleanName, Mid(invalidChars, i, 1), "_")
    Next i

    CleanFileName = cleanName
End Function