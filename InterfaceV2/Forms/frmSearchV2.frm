VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSearchV2
   Caption         =   "PCS Search V2 - Enhanced Search Interface"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12000
   OleObjectBlob   =   "frmSearchV2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSearchV2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private searchTimer As Long
Private lastSearchTerm As String
Private searchInProgress As Boolean
Private searchResults() As Object

Private Sub UserForm_Initialize()
    InitializeSearchInterface
    CacheManager.InitializeCache
End Sub

Private Sub InitializeSearchInterface()
    With Me
        .Width = 12000
        .Height = 7200
        .Caption = "PCS Search V2 - Enhanced Search Interface"
    End With

    InitializeControls
    SetupResultsList

    searchInProgress = False
    lastSearchTerm = ""
End Sub

Private Sub InitializeControls()
    With txtSearch
        .Text = ""
        .SetFocus
    End With

    With lstResults
        .MultiSelect = fmMultiSelectSingle
        .ListStyle = fmListStylePlain
        .BackColor = RGB(255, 255, 255)
        .ColumnCount = 5
        .ColumnWidths = "150;100;200;200;100"
    End With

    With txtResultPreview
        .MultiLine = True
        .ScrollBars = fmScrollBarsBoth
        .BackColor = RGB(248, 248, 248)
        .Locked = True
    End With

    With lblSearchStats
        .Caption = "Enter search term to begin"
        .ForeColor = RGB(100, 100, 100)
    End With

    With prgSearch
        .Visible = False
        .Min = 0
        .Max = 100
    End With
End Sub

Private Sub SetupResultsList()
    lstResults.Clear

    ' Add header row
    lstResults.AddItem "File Name" & vbTab & "Type" & vbTab & "Customer" & vbTab & "Component" & vbTab & "Score"

    ' Style the header row
    With lstResults
        .List(0, 0) = "File Name"
        .List(0, 1) = "Type"
        .List(0, 2) = "Customer"
        .List(0, 3) = "Component"
        .List(0, 4) = "Score"
    End With
End Sub

Private Sub txtSearch_Change()
    Dim currentTerm As String
    currentTerm = Trim(txtSearch.Text)

    ' Reset timer for debounced search
    searchTimer = Timer

    ' Clear results if search term is empty
    If Len(currentTerm) = 0 Then
        ClearResults
        lblSearchStats.Caption = "Enter search term to begin"
        Exit Sub
    End If

    ' Update status
    lblSearchStats.Caption = "Searching as you type..."

    ' Start debounced search after 500ms delay
    Application.OnTime Now + TimeSerial(0, 0, 1), "DelayedSearch"
End Sub

Public Sub DelayedSearch()
    ' Check if enough time has passed since last keystroke
    If Timer - searchTimer >= 0.5 Then
        ExecuteSearch
    Else
        ' Reschedule if user is still typing
        Application.OnTime Now + TimeSerial(0, 0, 1), "DelayedSearch"
    End If
End Sub

Private Sub ExecuteSearch()
    Dim searchTerm As String
    Dim startTime As Double

    searchTerm = Trim(txtSearch.Text)

    ' Don't search if term hasn't changed or is too short
    If searchTerm = lastSearchTerm Or Len(searchTerm) < 2 Then
        Exit Sub
    End If

    If searchInProgress Then Exit Sub

    searchInProgress = True
    startTime = Timer
    lastSearchTerm = searchTerm

    ShowSearchProgress "Searching...", 10

    ' Execute the smart search
    Dim results() As Object
    results = SearchEngineV2.ExecuteSmartSearch(searchTerm)

    ShowSearchProgress "Processing results...", 70

    ' Display results
    DisplaySearchResults results

    ShowSearchProgress "Complete", 100

    ' Update statistics
    Dim searchTime As Double
    searchTime = Timer - startTime

    lblSearchStats.Caption = "Found " & UBound(results) + 1 & " results in " & _
                           Format(searchTime, "0.00") & " seconds"

    HideSearchProgress
    searchInProgress = False
End Sub

Private Sub DisplaySearchResults(results() As Object)
    Dim i As Long
    Dim result As Object
    Dim fileName As String
    Dim displayRow As String

    ClearResults

    If UBound(results) = -1 Then
        lstResults.AddItem "No results found"
        Exit Sub
    End If

    ' Store results for later use
    ReDim searchResults(LBound(results) To UBound(results))

    For i = LBound(results) To UBound(results)
        Set result = results(i)
        Set searchResults(i) = result

        fileName = GetFileNameFromPath(result.FilePath)

        displayRow = fileName & vbTab & _
                    result.FileType & vbTab & _
                    result.CustomerName & vbTab & _
                    result.ComponentCode & vbTab & _
                    result.MatchScore

        lstResults.AddItem displayRow

        ' Color-code result types
        ColorCodeResult lstResults.ListCount - 1, result.FileType

        If i Mod 10 = 0 Then DoEvents
    Next i

    ' Highlight matching text in results
    HighlightSearchMatches lastSearchTerm
End Sub

Private Sub ColorCodeResult(rowIndex As Long, fileType As String)
    ' This would set row colors based on file type
    ' VBA ListBox has limited formatting options
    ' Implementation would depend on specific ListBox control used

    Select Case fileType
        Case "WIP"
            ' Red for WIP items (high priority)
        Case "Quote"
            ' Orange for quotes (medium priority)
        Case "Enquiry"
            ' Blue for enquiries (normal priority)
        Case "Archive"
            ' Gray for archived items (low priority)
    End Select
End Sub

Private Sub HighlightSearchMatches(searchTerm As String)
    ' Enhanced search result highlighting would go here
    ' This is limited by VBA ListBox capabilities
    ' Could be implemented with custom drawing or rich text controls
End Sub

Private Sub lstResults_Click()
    If lstResults.ListIndex > 0 Then ' Skip header row
        ShowResultPreview lstResults.ListIndex - 1 ' Adjust for header
    End If
End Sub

Private Sub ShowResultPreview(resultIndex As Long)
    If resultIndex >= 0 And resultIndex <= UBound(searchResults) Then
        Dim result As Object
        Set result = searchResults(resultIndex)

        Dim previewText As String
        previewText = "File: " & result.FilePath & vbCrLf
        previewText = previewText & "Type: " & result.FileType & vbCrLf
        previewText = previewText & "Customer: " & result.CustomerName & vbCrLf
        previewText = previewText & "Component Code: " & result.ComponentCode & vbCrLf
        previewText = previewText & "Description: " & result.ComponentDesc & vbCrLf
        previewText = previewText & "Status: " & result.Status & vbCrLf
        previewText = previewText & "Match Score: " & result.MatchScore & vbCrLf
        previewText = previewText & "Modified: " & Format(result.ModDate, "yyyy-mm-dd hh:mm:ss") & vbCrLf

        txtResultPreview.Text = previewText
    End If
End Sub

Private Sub lstResults_DblClick()
    If lstResults.ListIndex > 0 Then
        OpenSelectedFile lstResults.ListIndex - 1
    End If
End Sub

Private Sub OpenSelectedFile(resultIndex As Long)
    If resultIndex >= 0 And resultIndex <= UBound(searchResults) Then
        Dim result As Object
        Set result = searchResults(resultIndex)

        On Error GoTo ErrorHandler

        ' Open the file
        Application.Workbooks.Open result.FilePath

        ' Close search form
        Unload Me
        Exit Sub

ErrorHandler:
        MsgBox "Unable to open file: " & result.FilePath & vbCrLf & _
               "Error: " & Err.Description, vbExclamation, "File Open Error"
    End If
End Sub

Private Sub btnOpenFile_Click()
    If lstResults.ListIndex > 0 Then
        OpenSelectedFile lstResults.ListIndex - 1
    End If
End Sub

Private Sub btnCopyPath_Click()
    If lstResults.ListIndex > 0 Then
        Dim result As Object
        Set result = searchResults(lstResults.ListIndex - 1)

        ' Copy file path to clipboard (Windows API would be needed for full implementation)
        ' For now, show it in a message box
        MsgBox result.FilePath, vbInformation, "File Path"
    End If
End Sub

Private Sub btnShowInExplorer_Click()
    If lstResults.ListIndex > 0 Then
        Dim result As Object
        Set result = searchResults(lstResults.ListIndex - 1)

        ' Open Windows Explorer to file location
        On Error Resume Next
        Shell "explorer.exe /select," & Chr(34) & result.FilePath & Chr(34), vbNormalFocus
        On Error GoTo 0
    End If
End Sub

Private Sub btnAdvancedSearch_Click()
    ' Show advanced search options
    frmAdvancedSearch.Show vbModal
End Sub

Private Sub ShowSearchProgress(message As String, percentage As Integer)
    lblSearchStatus.Caption = message
    prgSearch.Value = percentage
    prgSearch.Visible = True
    DoEvents
End Sub

Private Sub HideSearchProgress()
    prgSearch.Visible = False
    lblSearchStatus.Caption = ""
End Sub

Private Sub ClearResults()
    lstResults.Clear
    SetupResultsList
    txtResultPreview.Text = ""
    ReDim searchResults(1 To 0)
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

' Quick action buttons
Private Sub btnNewEnquiry_Click()
    ' Create new enquiry from selected search result
    If lstResults.ListIndex > 0 Then
        Dim result As Object
        Set result = searchResults(lstResults.ListIndex - 1)

        ' Pre-populate enquiry form with customer data
        ' This would interface with the existing enquiry creation system
        MsgBox "Would create new enquiry for customer: " & result.CustomerName, vbInformation
    End If
End Sub

Private Sub btnConvertToQuote_Click()
    ' Convert selected enquiry to quote
    If lstResults.ListIndex > 0 Then
        Dim result As Object
        Set result = searchResults(lstResults.ListIndex - 1)

        If result.FileType = "Enquiry" Then
            ' Interface with existing quote conversion system
            MsgBox "Would convert enquiry to quote: " & result.FilePath, vbInformation
        Else
            MsgBox "Selected item is not an enquiry", vbExclamation
        End If
    End If
End Sub

Private Sub btnCreateJob_Click()
    ' Create job from selected quote
    If lstResults.ListIndex > 0 Then
        Dim result As Object
        Set result = searchResults(lstResults.ListIndex - 1)

        If result.FileType = "Quote" Then
            ' Interface with existing job creation system
            MsgBox "Would create job from quote: " & result.FilePath, vbInformation
        Else
            MsgBox "Selected item is not a quote", vbExclamation
        End If
    End If
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Cancel any pending search timers
    On Error Resume Next
    Application.OnTime Now + TimeSerial(0, 0, 1), "DelayedSearch", , False
    On Error GoTo 0
End Sub