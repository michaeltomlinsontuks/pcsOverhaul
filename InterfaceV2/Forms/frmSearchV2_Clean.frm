VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSearchV2
   Caption         =   "PCS Search V2"
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
Private searchResults() As DataTypes.SearchResult

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

    searchInProgress = False
    lastSearchTerm = ""
End Sub

Public Sub ExecuteSearch(searchTerm As String)
    Dim startTime As Double

    On Error GoTo ErrorHandler

    If Len(Trim(searchTerm)) < 2 Then
        Exit Sub
    End If

    If searchInProgress Then Exit Sub

    searchInProgress = True
    startTime = Timer
    lastSearchTerm = searchTerm

    ShowSearchProgress "Searching...", 10

    Dim results() As DataTypes.SearchResult
    results = SearchEngineV2.ExecuteSmartSearch(searchTerm)

    ShowSearchProgress "Processing results...", 70

    DisplaySearchResults results

    ShowSearchProgress "Complete", 100

    Dim searchTime As Double
    searchTime = Timer - startTime

    If UBound(results) >= 0 Then
        Debug.Print "Found " & (UBound(results) + 1) & " results in " & Format(searchTime, "0.00") & " seconds"
    Else
        Debug.Print "No results found in " & Format(searchTime, "0.00") & " seconds"
    End If

    HideSearchProgress
    searchInProgress = False
    Exit Sub

ErrorHandler:
    searchInProgress = False
    HideSearchProgress
    Debug.Print "Search error: " & Err.Description
End Sub

Private Sub DisplaySearchResults(results() As DataTypes.SearchResult)
    Dim i As Long
    Dim result As DataTypes.SearchResult
    Dim fileName As String

    Debug.Print "Displaying search results..."

    If UBound(results) = -1 Then
        Debug.Print "No results found"
        Exit Sub
    End If

    ReDim searchResults(LBound(results) To UBound(results))

    For i = LBound(results) To UBound(results)
        result = results(i)
        searchResults(i) = result

        fileName = GetFileNameFromPath(result.FilePath)

        Debug.Print fileName & " | " & result.FileType & " | " & result.CustomerName & " | " & result.ComponentCode & " | Score: " & result.MatchScore

        If i Mod 10 = 0 Then DoEvents
    Next i
End Sub

Public Sub ShowResultPreview(resultIndex As Long)
    On Error Resume Next

    If resultIndex >= 0 And resultIndex <= UBound(searchResults) Then
        Dim result As DataTypes.SearchResult
        result = searchResults(resultIndex)

        Debug.Print "=== File Preview ==="
        Debug.Print "File: " & result.FilePath
        Debug.Print "Type: " & result.FileType
        Debug.Print "Customer: " & result.CustomerName
        Debug.Print "Component Code: " & result.ComponentCode
        Debug.Print "Description: " & result.ComponentDesc
        Debug.Print "Status: " & result.Status
        Debug.Print "Match Score: " & result.MatchScore
        Debug.Print "Modified: " & Format(result.ModDate, "yyyy-mm-dd hh:mm:ss")
        Debug.Print "===================="
    End If

    On Error GoTo 0
End Sub

Public Sub OpenSelectedFile(resultIndex As Long)
    On Error GoTo ErrorHandler

    If resultIndex >= 0 And resultIndex <= UBound(searchResults) Then
        Dim result As DataTypes.SearchResult
        result = searchResults(resultIndex)

        Application.Workbooks.Open result.FilePath

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

Private Sub ShowSearchProgress(message As String, percentage As Integer)
    Debug.Print message & " (" & percentage & "%)"
    DoEvents
End Sub

Private Sub HideSearchProgress()
    Debug.Print "Search complete"
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

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Cleanup code
End Sub