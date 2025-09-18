Attribute VB_Name = "SearchModule"
Option Explicit

' This module provides the exact same procedures as the original Search_VBA/Module1.bas
' to ensure compatibility when the search form is integrated

Sub Show_Search_Menu()
    On Error GoTo Error_Handler

    frmSearch.Show
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "Show_Search_Menu", "SearchModule"
End Sub