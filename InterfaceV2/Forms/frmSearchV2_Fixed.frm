VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSearchV2
   Caption         =   "PCS Search V2 - Enhanced Search Interface"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSearchV2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private searchTimer As Double
Private lastSearchTerm As String
Private searchInProgress As Boolean
Private searchResults() As SearchEngineV2.SearchResult

' Control declarations to prevent compile errors
Private WithEvents txtSearch As MSForms.TextBox
Private WithEvents lstResults As MSForms.ListBox
Private WithEvents txtResultPreview As MSForms.TextBox
Private WithEvents lblSearchStats As MSForms.Label
Private WithEvents lblSearchStatus As MSForms.Label
Private WithEvents prgSearch As MSForms.Label ' Using Label as placeholder for ProgressBar
Private WithEvents btnOpenFile As MSForms.CommandButton
Private WithEvents btnCopyPath As MSForms.CommandButton
Private WithEvents btnShowInExplorer As MSForms.CommandButton
Private WithEvents btnAdvancedSearch As MSForms.CommandButton
Private WithEvents btnNewEnquiry As MSForms.CommandButton
Private WithEvents btnConvertToQuote As MSForms.CommandButton
Private WithEvents btnCreateJob As MSForms.CommandButton
Private WithEvents btnClose As MSForms.CommandButton

Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    InitializeSearchInterface
    CacheManager.InitializeCache
    Exit Sub

ErrorHandler:
    MsgBox "Error initializing search interface: " & Err.Description, vbCritical, "Initialization Error"
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
    On Error Resume Next

    If Not txtSearch Is Nothing Then
        With txtSearch
            .Text = ""
            .SetFocus
        End With
    End If

    If Not lstResults Is Nothing Then
        With lstResults
            .MultiSelect = fmMultiSelectSingle
            .ListStyle = fmListStylePlain
            .BackColor = RGB(255, 255, 255)
            .ColumnCount = 5
            .ColumnWidths = "150;100;200;200;100"
        End With
    End If

    If Not txtResultPreview Is Nothing Then
        With txtResultPreview
            .MultiLine = True
            .ScrollBars = fmScrollBarsBoth
            .BackColor = RGB(248, 248, 248)
            .Locked = True
        End With
    End If

    If Not lblSearchStats Is Nothing Then
        With lblSearchStats
            .Caption = "Enter search term to begin"
            .ForeColor = RGB(100, 100, 100)
        End With
    End If

    If Not prgSearch Is Nothing Then
        With prgSearch
            .Visible = False
            .Caption = "Search Progress"
        End With
    End If

    On Error GoTo 0
End Sub

Private Sub SetupResultsList()
    On Error Resume Next
    If Not lstResults Is Nothing Then
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
    End If
    On Error GoTo 0
End Sub

Private Sub txtSearch_Change()
    Dim currentTerm As String

    On Error Resume Next
    If Not txtSearch Is Nothing Then
        currentTerm = Trim(txtSearch.Text)

        ' Reset timer for debounced search
        searchTimer = Timer

        ' Clear results if search term is empty
        If Len(currentTerm) = 0 Then
            ClearResults
            If Not lblSearchStats Is Nothing Then lblSearchStats.Caption = "Enter search term to begin"
            Exit Sub
        End If

        ' Update status
        If Not lblSearchStats Is Nothing Then lblSearchStats.Caption = "Searching as you type..."

        ' Execute search directly (simplified approach)
        If Len(currentTerm) >= 2 And currentTerm <> lastSearchTerm Then
            ExecuteSearch
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub ExecuteSearch()
    Dim searchTerm As String
    Dim startTime As Double

    On Error GoTo ErrorHandler

    If Not txtSearch Is Nothing Then searchTerm = Trim(txtSearch.Text)

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
    Dim results() As SearchEngineV2.SearchResult
    results = SearchEngineV2.ExecuteSmartSearch(searchTerm)

    ShowSearchProgress "Processing results...", 70

    ' Display results
    DisplaySearchResults results

    ShowSearchProgress "Complete", 100

    ' Update statistics
    Dim searchTime As Double
    searchTime = Timer - startTime

    On Error Resume Next
    If Not lblSearchStats Is Nothing Then
        If UBound(results) >= 0 Then
            lblSearchStats.Caption = "Found " & (UBound(results) + 1) & " results in " & _
                                   Format(searchTime, "0.00") & " seconds"
        Else
            lblSearchStats.Caption = "No results found in " & Format(searchTime, "0.00") & " seconds"
        End If
    End If
    On Error GoTo 0

    HideSearchProgress
    searchInProgress = False
    Exit Sub

ErrorHandler:
    searchInProgress = False
    HideSearchProgress
    On Error Resume Next
    If Not lblSearchStats Is Nothing Then lblSearchStats.Caption = "Search error: " & Err.Description
    On Error GoTo 0
End Sub

Private Sub DisplaySearchResults(results() As SearchEngineV2.SearchResult)
    Dim i As Long
    Dim result As SearchEngineV2.SearchResult
    Dim fileName As String
    Dim displayRow As String

    On Error Resume Next

    ClearResults

    If UBound(results) = -1 Then
        If Not lstResults Is Nothing Then lstResults.AddItem "No results found"
        Exit Sub
    End If

    ' Store results for later use
    ReDim searchResults(LBound(results) To UBound(results))

    For i = LBound(results) To UBound(results)
        result = results(i)
        searchResults(i) = result

        fileName = GetFileNameFromPath(result.FilePath)

        displayRow = fileName & vbTab & _
                    result.FileType & vbTab & _
                    result.CustomerName & vbTab & _
                    result.ComponentCode & vbTab & _
                    result.MatchScore

        If Not lstResults Is Nothing Then lstResults.AddItem displayRow

        If i Mod 10 = 0 Then DoEvents
    Next i

    On Error GoTo 0
End Sub

Private Sub lstResults_Click()
    On Error Resume Next
    If Not lstResults Is Nothing Then
        If lstResults.ListIndex > 0 Then ' Skip header row
            ShowResultPreview lstResults.ListIndex - 1 ' Adjust for header
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub ShowResultPreview(resultIndex As Long)
    On Error Resume Next

    If resultIndex >= 0 And resultIndex <= UBound(searchResults) Then
        Dim result As SearchEngineV2.SearchResult
        result = searchResults(resultIndex)

        Dim previewText As String
        previewText = "File: " & result.FilePath & vbCrLf
        previewText = previewText & "Type: " & result.FileType & vbCrLf
        previewText = previewText & "Customer: " & result.CustomerName & vbCrLf
        previewText = previewText & "Component Code: " & result.ComponentCode & vbCrLf
        previewText = previewText & "Description: " & result.ComponentDesc & vbCrLf
        previewText = previewText & "Status: " & result.Status & vbCrLf
        previewText = previewText & "Match Score: " & result.MatchScore & vbCrLf
        previewText = previewText & "Modified: " & Format(result.ModDate, "yyyy-mm-dd hh:mm:ss") & vbCrLf

        If Not txtResultPreview Is Nothing Then txtResultPreview.Text = previewText
    End If

    On Error GoTo 0
End Sub

Private Sub lstResults_DblClick()
    On Error Resume Next
    If Not lstResults Is Nothing Then
        If lstResults.ListIndex > 0 Then
            OpenSelectedFile lstResults.ListIndex - 1
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub OpenSelectedFile(resultIndex As Long)
    On Error GoTo ErrorHandler

    If resultIndex >= 0 And resultIndex <= UBound(searchResults) Then
        Dim result As SearchEngineV2.SearchResult
        result = searchResults(resultIndex)

        ' Open the file
        Application.Workbooks.Open result.FilePath

        ' Close search form
        Unload Me
        Exit Sub
    End If
    Exit Sub

ErrorHandler:
    If resultIndex >= 0 And resultIndex <= UBound(searchResults) Then
        MsgBox "Unable to open file: " & searchResults(resultIndex).FilePath & vbCrLf & _
               "Error: " & Err.Description, vbExclamation, "File Open Error"
    End If
End Sub

Private Sub btnOpenFile_Click()
    On Error Resume Next
    If Not lstResults Is Nothing Then
        If lstResults.ListIndex > 0 Then
            OpenSelectedFile lstResults.ListIndex - 1
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub btnCopyPath_Click()
    On Error Resume Next
    If Not lstResults Is Nothing Then
        If lstResults.ListIndex > 0 Then
            Dim result As SearchEngineV2.SearchResult
            result = searchResults(lstResults.ListIndex - 1)

            ' Copy file path to clipboard (Windows API would be needed for full implementation)
            ' For now, show it in a message box
            MsgBox result.FilePath, vbInformation, "File Path"
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub btnShowInExplorer_Click()
    On Error Resume Next
    If Not lstResults Is Nothing Then
        If lstResults.ListIndex > 0 Then
            Dim result As SearchEngineV2.SearchResult
            result = searchResults(lstResults.ListIndex - 1)

            ' Open Windows Explorer to file location
            Shell "explorer.exe /select," & Chr(34) & result.FilePath & Chr(34), vbNormalFocus
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub btnAdvancedSearch_Click()
    ' Show advanced search options
    MsgBox "Advanced search options would go here", vbInformation, "Advanced Search"
End Sub

Private Sub ShowSearchProgress(message As String, percentage As Integer)
    On Error Resume Next
    If Not lblSearchStatus Is Nothing Then lblSearchStatus.Caption = message
    If Not prgSearch Is Nothing Then
        prgSearch.Caption = "Progress: " & percentage & "%"
        prgSearch.Visible = True
    End If
    DoEvents
    On Error GoTo 0
End Sub

Private Sub HideSearchProgress()
    On Error Resume Next
    If Not prgSearch Is Nothing Then prgSearch.Visible = False
    If Not lblSearchStatus Is Nothing Then lblSearchStatus.Caption = ""
    On Error GoTo 0
End Sub

Private Sub ClearResults()
    On Error Resume Next
    SetupResultsList
    If Not txtResultPreview Is Nothing Then txtResultPreview.Text = ""
    ReDim searchResults(1 To 0)
    On Error GoTo 0
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
    On Error Resume Next
    If Not lstResults Is Nothing Then
        If lstResults.ListIndex > 0 Then
            Dim result As SearchEngineV2.SearchResult
            result = searchResults(lstResults.ListIndex - 1)

            ' Pre-populate enquiry form with customer data
            MsgBox "Would create new enquiry for customer: " & result.CustomerName, vbInformation
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub btnConvertToQuote_Click()
    On Error Resume Next
    If Not lstResults Is Nothing Then
        If lstResults.ListIndex > 0 Then
            Dim result As SearchEngineV2.SearchResult
            result = searchResults(lstResults.ListIndex - 1)

            If result.FileType = "Enquiry" Then
                MsgBox "Would convert enquiry to quote: " & result.FilePath, vbInformation
            Else
                MsgBox "Selected item is not an enquiry", vbExclamation
            End If
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub btnCreateJob_Click()
    On Error Resume Next
    If Not lstResults Is Nothing Then
        If lstResults.ListIndex > 0 Then
            Dim result As SearchEngineV2.SearchResult
            result = searchResults(lstResults.ListIndex - 1)

            If result.FileType = "Quote" Then
                MsgBox "Would create job from quote: " & result.FilePath, vbInformation
            Else
                MsgBox "Selected item is not a quote", vbExclamation
            End If
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Cleanup code would go here
End Sub