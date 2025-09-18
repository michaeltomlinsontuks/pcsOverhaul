Attribute VB_Name = "a_ListFiles"
Public Function List_Files(path As String, frm As Object)
Dim Files(1 To 100000) As String
Dim FullFilePath As String, MyName As String
Dim GroupCount As Integer
'\* Check a Group folder exists
'FullFilePath = "C:\TEMP\Group*"

'fileextension = right(ucase(InputBox("Please enter file extension")),3)
fileextension = "*.*"

1:

'ChDrive (main.Main_MasterPath & path)

MyName = Dir(Main.Main_MasterPath & path & "\", vbDirectory)
    If MyName = "" Then
        MsgBox "Folder Not Found", vbOKOnly, "Test"
            Exit Function
    End If
'\* Store list of Group folder names

Do Until MyName = ""

    If MyName = "." Or MyName = ".." Then GoTo 2

    GroupCount = GroupCount + 1
    Files(GroupCount) = MyName
    
2:
    MyName = Dir
    
Loop

For i = 1 To GroupCount
    With frm
        x = Files(i)
        If path = "WIP" Then
            If GetValue(Main.Main_MasterPath.Value & path & "\", x, "ADMIN", "b88") = UCase("Quote Accepted") Then
                .AddItem Left(x, Len(x) - 4) & " *"
            Else
                .AddItem Left(x, Len(x) - 4)
            End If
        Else
            If path = "quotes" Then
                If GetValue(Main.Main_MasterPath & path & "\", x, "Admin", "b88") = "New Quote" Then
                    .AddItem Left(x, Len(x) - 4) & " *"
                Else
                    .AddItem Left(x, Len(x) - 4)
                End If
            Else
                .AddItem Left(x, Len(x) - 4)
            End If
        End If
    
    End With
Next i

End Function


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


