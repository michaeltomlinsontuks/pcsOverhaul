Attribute VB_Name = "InterfaceLauncher"
Option Explicit

' Main launcher macros for PCS Interface V2

Public Sub OpenMainInterface()
    ' Launch the main PCS interface
    On Error GoTo ErrorHandler

    ' Initialize cache if needed
    CacheManager.InitializeCache

    ' Show the main form
    MainV2.Show
    Exit Sub

ErrorHandler:
    MsgBox "Error opening main interface: " & Err.Description, vbExclamation, "Interface Error"
End Sub

Public Sub OpenSearchInterface()
    ' Launch the search interface
    On Error GoTo ErrorHandler

    ' Initialize cache if needed
    CacheManager.InitializeCache

    ' Show the search form
    frmSearchV2.Show
    Exit Sub

ErrorHandler:
    MsgBox "Error opening search interface: " & Err.Description, vbExclamation, "Search Error"
End Sub

Public Sub ShowMainForm()
    ' Alternative launcher for main form
    OpenMainInterface
End Sub

Public Sub ShowSearchForm()
    ' Alternative launcher for search form
    OpenSearchInterface
End Sub

Public Sub QuickSearch()
    ' Quick search launcher
    Dim searchTerm As String

    searchTerm = InputBox("Enter search term:", "Quick Search", "")

    If Len(Trim(searchTerm)) > 0 Then
        frmSearchV2.Show
        ' Note: Would need to modify frmSearchV2 to accept initial search term
    End If
End Sub

Public Sub RefreshAllCaches()
    ' Refresh all search caches
    Dim response As VbMsgBoxResult

    response = MsgBox("This will rebuild all search caches. This may take several minutes. Continue?", _
                     vbYesNo + vbQuestion, "Rebuild All Caches")

    If response = vbYes Then
        Application.ScreenUpdating = False

        CacheManager.ClearCache
        CacheManager.BuildCacheInBackground

        Application.ScreenUpdating = True

        MsgBox "Cache rebuild completed.", vbInformation, "Cache Rebuild"
    End If
End Sub

Public Sub ShowCacheStatistics()
    ' Display cache statistics
    MsgBox CacheManager.GetCacheStats(), vbInformation, "Cache Statistics"
End Sub

' Helper function to check if interface is properly initialized
Public Function IsInterfaceReady() As Boolean
    On Error GoTo ErrorHandler

    ' Check if cache manager is available
    CacheManager.InitializeCache

    ' Check if required directories exist
    Dim basePath As String
    basePath = Application.ActiveWorkbook.Path

    If Dir(basePath & "\Enquiries\", vbDirectory) = "" Then
        MsgBox "Enquiries directory not found. Please ensure the folder structure is correct.", vbExclamation
        IsInterfaceReady = False
        Exit Function
    End If

    IsInterfaceReady = True
    Exit Function

ErrorHandler:
    IsInterfaceReady = False
End Function

' Setup function to create necessary folder structure
Public Sub SetupInterface()
    Dim basePath As String
    Dim folders() As String
    Dim i As Long

    basePath = Application.ActiveWorkbook.Path

    ReDim folders(1 To 4)
    folders(1) = basePath & "\Enquiries\"
    folders(2) = basePath & "\Quotes\"
    folders(3) = basePath & "\WIP\"
    folders(4) = basePath & "\Archive\"

    For i = 1 To UBound(folders)
        If Dir(folders(i), vbDirectory) = "" Then
            MkDir folders(i)
        End If
    Next i

    MsgBox "Interface setup completed. Folder structure created.", vbInformation, "Setup Complete"
End Sub