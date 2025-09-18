Attribute VB_Name = "PerformanceMonitor"
Option Explicit

Private Type PerformanceMetrics
    SearchCount As Long
    TotalSearchTime As Double
    AverageSearchTime As Double
    CacheHitRate As Double
    FileOpenCount As Long
    TotalFileOpenTime As Double
    AverageFileOpenTime As Double
    LastResetTime As Date
End Type

Private metrics As PerformanceMetrics
Private logFile As String

Public Sub InitializeMonitoring()
    logFile = Application.ActiveWorkbook.Path & "\PerformanceLog.txt"
    ResetMetrics
End Sub

Public Sub ResetMetrics()
    With metrics
        .SearchCount = 0
        .TotalSearchTime = 0
        .AverageSearchTime = 0
        .CacheHitRate = 0
        .FileOpenCount = 0
        .TotalFileOpenTime = 0
        .AverageFileOpenTime = 0
        .LastResetTime = Now
    End With
End Sub

Public Sub LogSearchOperation(searchTerm As String, searchTime As Double, resultCount As Long)
    With metrics
        .SearchCount = .SearchCount + 1
        .TotalSearchTime = .TotalSearchTime + searchTime
        .AverageSearchTime = .TotalSearchTime / .SearchCount
    End With

    LogToFile "SEARCH", "Term: " & searchTerm & ", Time: " & Format(searchTime, "0.00") & "s, Results: " & resultCount
End Sub

Public Sub LogFileOperation(filePath As String, operation As String, operationTime As Double)
    With metrics
        .FileOpenCount = .FileOpenCount + 1
        .TotalFileOpenTime = .TotalFileOpenTime + operationTime
        .AverageFileOpenTime = .TotalFileOpenTime / .FileOpenCount
    End With

    LogToFile "FILE", operation & ": " & GetFileNameFromPath(filePath) & ", Time: " & Format(operationTime, "0.00") & "s"
End Sub

Public Sub LogCacheOperation(operation As String, details As String)
    LogToFile "CACHE", operation & ": " & details
End Sub

Public Sub LogError(errorSource As String, errorNumber As Long, errorDescription As String)
    LogToFile "ERROR", errorSource & " - #" & errorNumber & ": " & errorDescription
End Sub

Public Function GetPerformanceReport() As String
    Dim report As String
    Dim uptime As Double

    uptime = DateDiff("s", metrics.LastResetTime, Now)

    report = "=== PCS Interface V2 Performance Report ===" & vbCrLf & vbCrLf

    report = report & "Monitoring Period: " & Format(metrics.LastResetTime, "yyyy-mm-dd hh:mm:ss") & _
             " to " & Format(Now, "yyyy-mm-dd hh:mm:ss") & vbCrLf
    report = report & "Uptime: " & Format(uptime / 3600, "0.0") & " hours" & vbCrLf & vbCrLf

    ' Search Performance
    report = report & "SEARCH PERFORMANCE:" & vbCrLf
    report = report & "  Total Searches: " & metrics.SearchCount & vbCrLf
    If metrics.SearchCount > 0 Then
        report = report & "  Average Search Time: " & Format(metrics.AverageSearchTime, "0.00") & " seconds" & vbCrLf
        report = report & "  Searches per Hour: " & Format((metrics.SearchCount / uptime) * 3600, "0.0") & vbCrLf
    End If
    report = report & vbCrLf

    ' File Operation Performance
    report = report & "FILE OPERATIONS:" & vbCrLf
    report = report & "  Total File Opens: " & metrics.FileOpenCount & vbCrLf
    If metrics.FileOpenCount > 0 Then
        report = report & "  Average File Open Time: " & Format(metrics.AverageFileOpenTime, "0.00") & " seconds" & vbCrLf
        report = report & "  File Opens per Hour: " & Format((metrics.FileOpenCount / uptime) * 3600, "0.0") & vbCrLf
    End If
    report = report & vbCrLf

    ' Cache Performance
    report = report & "CACHE PERFORMANCE:" & vbCrLf
    report = report & CacheManager.GetCacheStats() & vbCrLf & vbCrLf

    ' Performance Targets vs Actual
    report = report & "PERFORMANCE TARGETS:" & vbCrLf
    report = report & "  Search Time Target: <2.0 seconds" & vbCrLf
    If metrics.AverageSearchTime > 0 Then
        If metrics.AverageSearchTime <= 2 Then
            report = report & "  Search Time Status: ✓ MEETING TARGET (" & Format(metrics.AverageSearchTime, "0.00") & "s)" & vbCrLf
        Else
            report = report & "  Search Time Status: ✗ ABOVE TARGET (" & Format(metrics.AverageSearchTime, "0.00") & "s)" & vbCrLf
        End If
    End If

    report = report & "  File Open Target: <1.0 seconds" & vbCrLf
    If metrics.AverageFileOpenTime > 0 Then
        If metrics.AverageFileOpenTime <= 1 Then
            report = report & "  File Open Status: ✓ MEETING TARGET (" & Format(metrics.AverageFileOpenTime, "0.00") & "s)" & vbCrLf
        Else
            report = report & "  File Open Status: ✗ ABOVE TARGET (" & Format(metrics.AverageFileOpenTime, "0.00") & "s)" & vbCrLf
        End If
    End If

    GetPerformanceReport = report
End Function

Public Sub ShowPerformanceDialog()
    Dim report As String
    report = GetPerformanceReport()

    ' Create a custom dialog or use a simple message box
    MsgBox report, vbInformation, "Performance Monitor"
End Sub

Public Sub ExportPerformanceReport()
    Dim filePath As String
    Dim fileNum As Integer
    Dim report As String

    filePath = Application.ActiveWorkbook.Path & "\PerformanceReport_" & Format(Now, "yyyymmdd_hhmmss") & ".txt"
    report = GetPerformanceReport()

    On Error GoTo ErrorHandler

    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, report
    Close #fileNum

    MsgBox "Performance report exported to: " & filePath, vbInformation, "Export Complete"
    Exit Sub

ErrorHandler:
    If fileNum > 0 Then Close #fileNum
    MsgBox "Error exporting performance report: " & Err.Description, vbExclamation, "Export Error"
End Sub

Private Sub LogToFile(category As String, message As String)
    Dim fileNum As Integer
    Dim logEntry As String

    On Error GoTo ErrorHandler

    logEntry = Format(Now, "yyyy-mm-dd hh:mm:ss") & " [" & category & "] " & message

    fileNum = FreeFile
    Open logFile For Append As #fileNum
    Print #fileNum, logEntry
    Close #fileNum
    Exit Sub

ErrorHandler:
    If fileNum > 0 Then Close #fileNum
    ' Don't show error messages for logging failures to avoid infinite loops
End Sub

Private Function GetFileNameFromPath(fullPath As String) As String
    Dim lastSlash As Long
    lastSlash = InStrRev(fullPath, "\")
    If lastSlash > 0 Then
        GetFileNameFromPath = Mid(fullPath, lastSlash + 1)
    Else
        GetFileNameFromPath = fullPath
    End If
End Function

Public Function StartTimer() As Double
    StartTimer = Timer
End Function

Public Function StopTimer(startTime As Double) As Double
    StopTimer = Timer - startTime
End Function

Public Sub MonitorMemoryUsage()
    ' This would require Windows API calls to get actual memory usage
    ' For now, we'll track basic VBA object counts and cache size
    Dim memoryReport As String

    memoryReport = "Memory Usage Estimate:" & vbCrLf
    memoryReport = memoryReport & "Cache Entries: " & CacheManager.GetCacheStats() & vbCrLf

    LogToFile "MEMORY", memoryReport
End Sub

Public Sub SchedulePerformanceCheck()
    ' Performance monitoring scheduling disabled for VBA compatibility
    ' Call PeriodicCheck manually as needed
End Sub

Public Sub PeriodicCheck()
    MonitorMemoryUsage

    ' Log current performance state
    LogToFile "PERIODIC", "Searches: " & metrics.SearchCount & ", Avg Time: " & Format(metrics.AverageSearchTime, "0.00") & "s"
End Sub

Public Sub StopMonitoring()
    ' Log final statistics
    LogToFile "SHUTDOWN", "Final Stats - " & GetPerformanceReport()
End Sub