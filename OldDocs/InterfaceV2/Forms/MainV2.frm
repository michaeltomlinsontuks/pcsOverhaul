Attribute VB_Name = "MainV2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Type FilterState
    NewEnquiries As Boolean
    QuotesToSubmit As Boolean
    WIPToSequence As Boolean
    JobsInWIP As Boolean
    ShowArchived As Boolean
    DateRangeStart As Date
    DateRangeEnd As Date
End Type

Private currentFilters As FilterState
Private lastRefreshTime As Date
Private refreshInProgress As Boolean
Private performanceMetrics As String

Private Sub UserForm_Initialize()
    InitializeInterface
    CacheManager.InitializeCache
    LoadUserPreferences
    RefreshListSmart
End Sub

Private Sub InitializeInterface()
    With Me
        .Width = 16500
        .Height = 9000
        .Caption = "PCS Interface V2 - Enhanced Performance Dashboard"
    End With

    InitializeControls
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

Private Sub InitializeControls()
    With lstMain
        .MultiSelect = fmMultiSelectSingle
        .ListStyle = fmListStylePlain
        .BackColor = RGB(255, 255, 255)
        .ForeColor = RGB(0, 0, 0)
    End With

    With txtPreview
        .MultiLine = True
        .ScrollBars = fmScrollBarsBoth
        .BackColor = RGB(248, 248, 248)
        .Locked = True
    End With

    With lblPerformance
        .Caption = "Performance: Ready"
        .ForeColor = RGB(0, 128, 0)
    End With

    With prgProgress
        .Visible = False
        .Min = 0
        .Max = 100
    End With
End Sub

Private Sub SetupPerformanceCounters()
    lblEnquiryCount.Caption = "Enquiries: 0"
    lblQuoteCount.Caption = "Quotes: 0"
    lblWIPCount.Caption = "WIP: 0"
    lblJobCount.Caption = "Jobs: 0"
    lblCacheStats.Caption = "Cache: Initializing..."
End Sub

Public Function RefreshListSmart() As Boolean
    Dim startTime As Double
    Dim needsRefresh As Boolean
    Dim timeSinceLastRefresh As Double

    If refreshInProgress Then
        RefreshListSmart = False
        Exit Function
    End If

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
    Dim i As Long
    Dim fileCount As Long

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
    lblPerformance.Caption = "Performance: " & performanceMetrics

    refreshInProgress = False
    HideProgress
    RefreshListSmart = True
End Function

Private Function ApplyFilters(fileList() As String) As String()
    Dim filteredList() As String
    Dim filteredCount As Long
    Dim i As Long
    Dim filePath As String
    Dim fileType As String
    Dim includeFile As Boolean

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

    lstMain.Clear

    For i = LBound(fileList) To UBound(fileList)
        filePath = fileList(i)
        If filePath <> "" Then
            fileName = GetFileNameOnly(filePath)
            fileType = FileUtilities.GetFileTypeFromPath(filePath)
            customer = CacheManager.GetCachedValue(filePath, "CustomerName")
            component = CacheManager.GetCachedValue(filePath, "ComponentCode")

            displayText = fileName & " | " & fileType & " | " & customer & " | " & component
            lstMain.AddItem displayText

            If i Mod 50 = 0 Then DoEvents
        End If
    Next i
End Sub

Private Sub UpdateCounters(fileList() As String)
    Dim enquiryCount As Long, quoteCount As Long, wipCount As Long, jobCount As Long
    Dim i As Long
    Dim fileType As String

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

    lblEnquiryCount.Caption = "Enquiries: " & enquiryCount
    lblQuoteCount.Caption = "Quotes: " & quoteCount
    lblWIPCount.Caption = "WIP: " & wipCount
    lblJobCount.Caption = "Jobs: " & jobCount
End Sub

Private Sub UpdateCacheStats()
    lblCacheStats.Caption = "Cache: " & CacheManager.GetCacheStats()
End Sub

Private Function FiltersChanged() As Boolean
    Static lastFilters As FilterState

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
    lblStatus.Caption = message
    prgProgress.Value = percentage
    prgProgress.Visible = True
    DoEvents
End Sub

Private Sub HideProgress()
    prgProgress.Visible = False
    lblStatus.Caption = "Ready"
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

' Event Handlers
Private Sub chkNewEnquiries_Click()
    currentFilters.NewEnquiries = chkNewEnquiries.Value
    RefreshListSmart
End Sub

Private Sub chkQuotesToSubmit_Click()
    currentFilters.QuotesToSubmit = chkQuotesToSubmit.Value
    RefreshListSmart
End Sub

Private Sub chkWIPToSequence_Click()
    currentFilters.WIPToSequence = chkWIPToSequence.Value
    RefreshListSmart
End Sub

Private Sub chkJobsInWIP_Click()
    currentFilters.JobsInWIP = chkJobsInWIP.Value
    RefreshListSmart
End Sub

Private Sub chkShowArchived_Click()
    currentFilters.ShowArchived = chkShowArchived.Value
    RefreshListSmart
End Sub

Private Sub lstMain_Click()
    ShowPreview
End Sub

Private Sub ShowPreview()
    If lstMain.ListIndex >= 0 Then
        Dim selectedText As String
        Dim filePath As String
        Dim parts() As String

        selectedText = lstMain.List(lstMain.ListIndex)
        parts = Split(selectedText, " | ")

        If UBound(parts) >= 0 Then
            filePath = FindFullPath(parts(0))
            LoadFilePreview filePath
        End If
    End If
End Sub

Private Function FindFullPath(fileName As String) As String
    Dim searchPaths() As String
    Dim i As Long
    Dim fullPath As String

    ReDim searchPaths(1 To 4)
    searchPaths(1) = Application.ActiveWorkbook.Path & "\Enquiries\" & fileName
    searchPaths(2) = Application.ActiveWorkbook.Path & "\Quotes\" & fileName
    searchPaths(3) = Application.ActiveWorkbook.Path & "\WIP\" & fileName
    searchPaths(4) = Application.ActiveWorkbook.Path & "\Archive\" & fileName

    For i = 1 To UBound(searchPaths)
        If Dir(searchPaths(i)) <> "" Then
            FindFullPath = searchPaths(i)
            Exit Function
        End If
    Next i

    FindFullPath = ""
End Function

Private Sub LoadFilePreview(filePath As String)
    If filePath = "" Then
        txtPreview.Text = "File not found"
        Exit Sub
    End If

    Dim previewText As String
    Dim customer As String
    Dim component As String
    Dim description As String
    Dim modDate As String

    customer = CacheManager.GetCachedValue(filePath, "CustomerName")
    component = CacheManager.GetCachedValue(filePath, "ComponentCode")
    description = CacheManager.GetCachedValue(filePath, "ComponentDesc")

    On Error Resume Next
    modDate = Format(FileDateTime(filePath), "yyyy-mm-dd hh:mm:ss")
    On Error GoTo 0

    previewText = "File: " & filePath & vbCrLf
    previewText = previewText & "Modified: " & modDate & vbCrLf
    previewText = previewText & "Customer: " & customer & vbCrLf
    previewText = previewText & "Component: " & component & vbCrLf
    previewText = previewText & "Description: " & description & vbCrLf

    txtPreview.Text = previewText
End Sub

Private Sub btnRefresh_Click()
    RefreshListSmart
End Sub

Private Sub btnSearch_Click()
    frmSearchV2.Show
End Sub

Private Sub btnCacheStats_Click()
    MsgBox CacheManager.GetCacheStats(), vbInformation, "Cache Statistics"
End Sub

Private Sub btnRebuildCache_Click()
    Dim response As VbMsgBoxResult
    response = MsgBox("This will rebuild the search cache. This may take several minutes. Continue?", _
                     vbYesNo + vbQuestion, "Rebuild Cache")

    If response = vbYes Then
        CacheManager.ClearCache
        CacheManager.BuildCacheInBackground
        MsgBox "Cache rebuild completed.", vbInformation, "Cache Rebuild"
        RefreshListSmart
    End If
End Sub

Private Sub LoadUserPreferences()
    ' Load user preferences from file if available
    ' This would typically read from SystemConfig.txt
End Sub

Private Sub SaveUserPreferences()
    ' Save current filter settings to SystemConfig.txt
End Sub

Private Sub UserForm_Terminate()
    SaveUserPreferences
    CacheManager.SaveCacheToFile
End Sub