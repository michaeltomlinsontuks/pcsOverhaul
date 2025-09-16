Attribute VB_Name = "Delete_Sheet"
' Delete_Sheet
' Deletes a sheet without prompting to confirm delete

Option Explicit
Public Function DeleteSheet(SheetName As String)

    Application.DisplayAlerts = False
    Worksheets(SheetName).Delete
    Application.DisplayAlerts = True

End Function

