Attribute VB_Name = "SearchModule"
Option Explicit

' Module to maintain compatibility with original Search_VBA interface
' Provides Show_Search_Menu() function that shows the refactored form

Public Sub Show_Search_Menu()
    On Error GoTo Error_Handler

    ' Show the refactored search form with optimized backend
    frmSearch.Show
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "Show_Search_Menu", "SearchModule"
End Sub