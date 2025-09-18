Attribute VB_Name = "GetValue"
Public Function GetValue(path, File, sheet, ref)
'   Retrieves a value from a closed workbook
    Dim arg As String

'   Make sure the file exists
    If Right(path, 1) <> "\" Then path = path & "\"
    If Dir(path & File) = "" Then
        GetValue = "File Not Found"
        Exit Function
    End If

'   Create the argument
    arg = "'" & path & "[" & File & "]" & sheet & "'!" & _
      Range(ref).Range("A1").Address(, , xlR1C1)

'   Execute an XLM macro
    GetValue = ExecuteExcel4Macro(arg)
End Function

Private Function TestGetValue()
    p = "c:\XLFiles\Budget"
    f = "99Budget.xls"
    s = "Sheet1"
    a = "A1"
    MsgBox GetValue(p, f, s, a)
End Function

'Another example is shown below. This procedure reads 1,200 values (100 rows and 12 columns) from a closed file, and places the values into the active worksheet.

Private Function TestGetValue2()
    
    p = "c:\XLFiles\Budget"
    f = "99Budget.xls"
    s = "Sheet1"
    Application.ScreenUpdating = False
    For r = 1 To 100
        For c = 1 To 12
            a = Cells(r, c).Address
            Cells(r, c) = GetValue(p, f, s, a)
        Next c
    Next r
    Application.ScreenUpdating = True

End Function


