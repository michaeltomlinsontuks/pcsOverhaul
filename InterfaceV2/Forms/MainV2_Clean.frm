VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainV2
   Caption         =   "PCS Interface V2"
   ClientHeight    =   9000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16500
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainV2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private currentFilters As DataTypes.FilterState
Private lastRefreshTime As Date
Private refreshInProgress As Boolean
Private performanceMetrics As String

Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    InitializeInterface
    CacheManager.InitializeCache
    LoadUserPreferences
    RefreshListSmart
    Exit Sub

ErrorHandler:
    MsgBox "Error initializing interface: " & Err.Description, vbCritical, "Initialization Error"
End Sub

Private Sub InitializeInterface()
    With Me
        .Width = 16500
        .Height = 9000
        .Caption = "PCS Interface V2 - Enhanced Performance Dashboard"
    End With

    SetupPerformanceCounters

    currentFilters.NewEnquiries = True
    currentFilters.QuotesToSubmit = True
    currentFilters.WIPToSequence = True
    currentFilters.JobsInWIP = True
    currentFilters.ShowArchived = False
    currentFilters.DateRangeStart = DateAdd("m", -3, Now)
    currentFilters.DateRangeEnd = Now

    refreshInProgress = False
End Sub

Private Sub SetupPerformanceCounters()
    ' Initialize performance counters
    On Error Resume Next
    ' Labels will be created manually or can be added programmatically
    On Error GoTo 0
End Sub

Public Function RefreshListSmart() As Boolean
    Dim startTime As Double
    Dim needsRefresh As Boolean
    Dim timeSinceLastRefresh As Double

    If refreshInProgress Then
        RefreshListSmart = False
        Exit Function
    End If

    On Error GoTo ErrorHandler

    startTime = Timer
    refreshInProgress = True

    ShowProgress "Checking for updates...", 10

    timeSinceLastRefresh = DateDiff("s", lastRefreshTime, Now)
    needsRefresh = (timeSinceLastRefresh > 60) Or FiltersChanged() Or (lastRefreshTime = 0)

    If Not needsRefresh Then
        RefreshListSmart = True
        refreshInProgress = False
        HideProgress
        Exit Function
    End If

    ShowProgress "Building file list...", 25
    DoEvents

    Dim fileList() As String
    Dim filteredFiles() As String

    fileList = FileUtilities.BuildFileList()
    filteredFiles = ApplyFilters(fileList)

    ShowProgress "Populating list...", 60
    DoEvents

    PopulateMainList filteredFiles

    ShowProgress "Updating counters...", 80
    DoEvents

    UpdateCounters filteredFiles
    UpdateCacheStats

    ShowProgress "Complete", 100

    lastRefreshTime = Now
    performanceMetrics = "Last refresh: " & Format(Timer - startTime, "0.00") & "s"

    refreshInProgress = False
    HideProgress
    RefreshListSmart = True
    Exit Function

ErrorHandler:
    refreshInProgress = False
    HideProgress
    MsgBox "Error refreshing list: " & Err.Description, vbExclamation, "Refresh Error"
    RefreshListSmart = False
End Function

Private Function ApplyFilters(fileList() As String) As String()
    Dim filteredList() As String
    Dim filteredCount As Long
    Dim i As Long
    Dim filePath As String
    Dim fileType As String
    Dim includeFile As Boolean

    If UBound(fileList) < LBound(fileList) Then
        ReDim filteredList(1 To 0)
        ApplyFilters = filteredList
        Exit Function
    End If

    ReDim filteredList(1 To UBound(fileList) + 1)
    filteredCount = 0

    For i = LBound(fileList) To UBound(fileList)
        filePath = fileList(i)
        If filePath <> "" Then
            fileType = FileUtilities.GetFileTypeFromPath(filePath)
            includeFile = False

            Select Case fileType
                Case "Enquiry"
                    includeFile = currentFilters.NewEnquiries
                Case "Quote"
                    includeFile = currentFilters.QuotesToSubmit
                Case "WIP"
                    includeFile = currentFilters.WIPToSequence Or currentFilters.JobsInWIP
                Case "Archive"
                    includeFile = currentFilters.ShowArchived
            End Select

            If includeFile And WithinDateRange(filePath) Then
                filteredCount = filteredCount + 1
                filteredList(filteredCount) = filePath
            End If
        End If

        If i Mod 20 = 0 Then DoEvents
    Next i

    If filteredCount > 0 Then
        ReDim Preserve filteredList(1 To filteredCount)
    Else
        ReDim filteredList(1 To 0)
    End If

    ApplyFilters = filteredList
End Function

Private Sub PopulateMainList(fileList() As String)
    Dim i As Long
    Dim displayText As String
    Dim filePath As String
    Dim fileName As String
    Dim fileType As String
    Dim customer As String
    Dim component As String

    ' This would populate a listbox control when available
    Debug.Print "Populating list with " & UBound(fileList) & " files"

    If UBound(fileList) < LBound(fileList) Then Exit Sub

    For i = LBound(fileList) To UBound(fileList)
        filePath = fileList(i)
        If filePath <> "" Then
            fileName = GetFileNameOnly(filePath)
            fileType = FileUtilities.GetFileTypeFromPath(filePath)
            customer = CacheManager.GetCachedValue(filePath, "CustomerName")
            component = CacheManager.GetCachedValue(filePath, "ComponentCode")

            displayText = fileName & " | " & fileType & " | " & customer & " | " & component
            Debug.Print displayText

            If i Mod 50 = 0 Then DoEvents
        End If
    Next i
End Sub

Private Sub UpdateCounters(fileList() As String)
    Dim enquiryCount As Long, quoteCount As Long, wipCount As Long, jobCount As Long
    Dim i As Long
    Dim fileType As String

    If UBound(fileList) >= LBound(fileList) Then
        For i = LBound(fileList) To UBound(fileList)
            If fileList(i) <> "" Then
                fileType = FileUtilities.GetFileTypeFromPath(fileList(i))

                Select Case fileType
                    Case "Enquiry": enquiryCount = enquiryCount + 1
                    Case "Quote": quoteCount = quoteCount + 1
                    Case "WIP": wipCount = wipCount + 1
                    Case "Archive": jobCount = jobCount + 1
                End Select
            End If
        Next i
    End If

    Debug.Print "Counters - Enquiries: " & enquiryCount & ", Quotes: " & quoteCount & ", WIP: " & wipCount & ", Jobs: " & jobCount
End Sub

Private Sub UpdateCacheStats()
    Debug.Print "Cache Stats: " & CacheManager.GetCacheStats()
End Sub

Private Function FiltersChanged() As Boolean
    Static lastFilters As DataTypes.FilterState

    FiltersChanged = (lastFilters.NewEnquiries <> currentFilters.NewEnquiries) Or _
                    (lastFilters.QuotesToSubmit <> currentFilters.QuotesToSubmit) Or _
                    (lastFilters.WIPToSequence <> currentFilters.WIPToSequence) Or _
                    (lastFilters.JobsInWIP <> currentFilters.JobsInWIP) Or _
                    (lastFilters.ShowArchived <> currentFilters.ShowArchived) Or _
                    (lastFilters.DateRangeStart <> currentFilters.DateRangeStart) Or _
                    (lastFilters.DateRangeEnd <> currentFilters.DateRangeEnd)

    lastFilters = currentFilters
End Function

Private Function WithinDateRange(filePath As String) As Boolean
    Dim fileDate As Date

    On Error Resume Next
    fileDate = FileDateTime(filePath)
    If Err.Number <> 0 Then
        WithinDateRange = True
        Exit Function
    End If
    On Error GoTo 0

    WithinDateRange = (fileDate >= currentFilters.DateRangeStart) And _
                     (fileDate <= currentFilters.DateRangeEnd)
End Function

Private Sub ShowProgress(message As String, percentage As Integer)
    Debug.Print message & " (" & percentage & "%)"
    DoEvents
End Sub

Private Sub HideProgress()
    Debug.Print "Ready"
End Sub

Private Function GetFileNameOnly(fullPath As String) As String
    Dim lastSlash As Long
    lastSlash = InStrRev(fullPath, "\")
    If lastSlash > 0 Then
        GetFileNameOnly = Mid(fullPath, lastSlash + 1)
    Else
        GetFileNameOnly = fullPath
    End If
End Function

Private Sub LoadUserPreferences()
    ' Load user preferences
End Sub

Private Sub SaveUserPreferences()
    ' Save current filter settings
End Sub

Private Sub UserForm_Terminate()
    SaveUserPreferences
    CacheManager.SaveCacheToFile
End Sub