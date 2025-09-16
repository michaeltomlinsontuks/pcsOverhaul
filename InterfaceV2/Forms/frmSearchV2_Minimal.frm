VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSearchV2
   Caption         =   "PCS Search V2"
   ClientHeight    =   4000
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSearchV2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private searchResults() As DataTypes.SearchResult

Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    Me.Caption = "PCS Search V2 - Enhanced Search Interface"
    Me.Width = 8000
    Me.Height = 4000

    CacheManager.InitializeCache

    MsgBox "Search interface initialized successfully!", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Error initializing search interface: " & Err.Description, vbCritical, "Initialization Error"
End Sub

Public Sub ExecuteSearch(searchTerm As String)
    On Error GoTo ErrorHandler

    If Len(Trim(searchTerm)) < 2 Then
        MsgBox "Please enter at least 2 characters to search", vbInformation
        Exit Sub
    End If

    Debug.Print "Searching for: " & searchTerm

    Dim results() As DataTypes.SearchResult
    results = SearchEngineV2.ExecuteSmartSearch(searchTerm)

    DisplaySearchResults results
    Exit Sub

ErrorHandler:
    MsgBox "Search error: " & Err.Description, vbExclamation, "Search Error"
End Sub

Private Sub DisplaySearchResults(results() As DataTypes.SearchResult)
    Dim i As Long

    Debug.Print "=== Search Results ==="

    If UBound(results) = -1 Then
        Debug.Print "No results found"
        MsgBox "No results found", vbInformation
        Exit Sub
    End If

    ReDim searchResults(LBound(results) To UBound(results))

    For i = LBound(results) To UBound(results)
        searchResults(i) = results(i)
        Debug.Print results(i).FilePath & " | Score: " & results(i).MatchScore
    Next i

    Debug.Print "====================="
    MsgBox "Found " & (UBound(results) + 1) & " results. Check Immediate Window for details.", vbInformation
End Sub

Private Sub UserForm_Click()
    Dim searchTerm As String
    searchTerm = InputBox("Enter search term:", "Quick Search", "")
    If Len(Trim(searchTerm)) > 0 Then
        ExecuteSearch searchTerm
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Cleanup
End Sub