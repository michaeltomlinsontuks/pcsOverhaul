Attribute VB_Name = "SearchLauncher"
Option Explicit

' This module provides compatibility with the original search system
' while integrating it with the V2 architecture

Public Sub Show_Search_Menu()
    On Error GoTo Error_Handler

    ' This mimics the original Module1.bas Show_Search_Menu subroutine
    frmSearch.Show
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "Show_Search_Menu", "SearchLauncher"
End Sub

Public Sub LaunchIntegratedSearch()
    ' This provides a way to launch search directly from InterfaceVBA_V2
    On Error GoTo Error_Handler

    Dim SearchPath As String
    Dim wb As Workbook

    SearchPath = FileManager.GetRootPath & "\Search.xls"

    ' Check if Search.xls exists
    If Not FileManager.FileExists(SearchPath) Then
        MsgBox "Search database not found: " & SearchPath & vbCrLf & vbCrLf & _
               "Please ensure Search.xls exists in the root directory.", vbCritical, "Search Database Missing"
        Exit Sub
    End If

    ' Open Search.xls
    Set wb = FileManager.SafeOpenWorkbook(SearchPath)
    If wb Is Nothing Then
        MsgBox "Could not open search database: " & SearchPath, vbCritical, "Search Error"
        Exit Sub
    End If

    ' Activate the search worksheet
    On Error Resume Next
    wb.Worksheets(1).Activate
    If Err.Number <> 0 Then
        On Error GoTo Error_Handler
        MsgBox "Could not activate search worksheet.", vbCritical, "Search Error"
        Exit Sub
    End If
    On Error GoTo Error_Handler

    ' Show the search form
    frmSearch.Show
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "LaunchIntegratedSearch", "SearchLauncher"
End Sub

Public Sub RefreshSearchDatabase()
    ' This function provides database maintenance functionality
    On Error GoTo Error_Handler

    If SearchService.SortSearchDatabase() Then
        MsgBox "Search database refreshed successfully.", vbInformation, "Database Refresh"
    Else
        MsgBox "Failed to refresh search database.", vbCritical, "Database Refresh Error"
    End If
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "RefreshSearchDatabase", "SearchLauncher"
End Sub