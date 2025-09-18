Attribute VB_Name = "Open_Book"
Public Function OpenBook(File As String, RO As Boolean)

    Workbooks.Open Filename:= _
        File, _
        ReadOnly:=RO

End Function

