Attribute VB_Name = "Very_HiddenSheet"
' This example creates a new worksheet and then sets its
' Visible property to xlVeryHidden. To refer to the sheet,
' use its object variable, newSheet, as shown in the last
' line of the example. To use the newSheet object variable
' in another procedure, you must declare it as a public
' variable (Public newSheet As Object) in the first line of
' the module preceding any Sub or Function procedure.

Public Function VeryHiddenSheet(SheetNam As String)

        Sheets(SheetNam).Visible = xlVeryHidden

End Function

Public Function ShowSheet(SheetNam As String)
    
    Sheets(SheetNam).Visible = True

End Function

