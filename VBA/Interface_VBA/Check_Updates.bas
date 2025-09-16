Attribute VB_Name = "Check_Updates"
Public NextCheck As Date

Public Function CheckUpdates()

If Main.Visible = False Or NextCheck > Now() Then
    If NextCheck = "12:00:00 AM" Then GoTo 5
    StopCheck
    Exit Function
End If

5:

On Error GoTo 8

If "Enquiries : " & Check_Files(Main.Main_MasterPath & "enquiries\") <> Main.Notice_Enquiries.Caption Then
    Main.Notice_Enquiries.Caption = "Enquiries : " & Check_Files(Main.Main_MasterPath & "enquiries\") & "*"
End If

If "Quotes : " & Check_Files(Main.Main_MasterPath & "Quotes\") <> Main.Notice_Quotes.Caption Then
    Main.Notice_Quotes.Caption = "Quotes : " & Check_Files(Main.Main_MasterPath & "Quotes\") & "*"
End If

If "WIP : " & Check_Files(Main.Main_MasterPath & "WIP\") <> Main.Notice_WIP.Caption Then
    Main.Notice_WIP.Caption = "WIP : " & Check_Files(Main.Main_MasterPath & "WIP\") & "*"
End If

NextCheck = Now + TimeValue("00:05:00")
Application.OnTime NextCheck, "CheckUpdates", NextCheck + TimeValue("00:01:00")

8:

End Function

Public Function StopCheck()

On Error Resume Next
Application.OnTime NextCheck, "CheckUpdates", , Schedule:=False


End Function
Public Function Check_Files(path As String)

Dim FullFilePath As String, MyName As String
Dim GroupCount As Integer
'\* Check a Group folder exists
'FullFilePath = "C:\TEMP\Group*"

GroupCount = 0
'fileextension = right(ucase(InputBox("Please enter file extension")),3)
fileextension = "*.*"

1:

'ChDrive (main.Main_MasterPath & path)

MyName = Dir(path, vbDirectory)
    If MyName = "" Then
        StopCheck
'        MsgBox "Folder Not Found", vbOKOnly, "Test"
            Exit Function
    End If
'\* Store list of Group folder names

Do Until MyName = ""

    If MyName = "." Or MyName = ".." Or MyName = "_Users.xls" Then GoTo 2
    
    GroupCount = GroupCount + 1

2:
    
    MyName = Dir
    
Loop
    
Check_Files = GroupCount

End Function




