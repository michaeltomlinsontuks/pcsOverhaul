Attribute VB_Name = "Calc_Numbers"
Public Function Calc_Next_Number(Typ As String)
Dim FullFilePath As String, MyName As String
Dim GroupCount As Integer
'\* Check a Group folder exists
'FullFilePath = "C:\TEMP\Group*"

'fileextension = right(ucase(InputBox("Please enter file extension")),3)
fileextension = "*.*"

'ChDrive (main.Main_MasterPath & path)

    path = "templates"
    
    MyName = Dir(Main.Main_MasterPath.Value & path & "\", vbDirectory)
        If MyName = "" Then
            MsgBox "Folder Not Found", vbOKOnly, "Test"
                Exit Function
        End If
    '\* Store list of Group folder names
     
    Do Until MyName = ""
    
        If MyName = "." Or MyName = ".." Then GoTo 2
                
            If Left(UCase(Typ), 1) = "E" And Left(MyName, 4) = "E - " Then
                Calc_Next_Number = Mid(MyName, InStr(1, MyName, "-", vbTextCompare) + 2, Len(MyName) - 8) + 1
                
                GoTo 8
                
            End If
            
            If Left(UCase(Typ), 1) = "J" And Left(MyName, 4) = "J - " Then
                Calc_Next_Number = Mid(MyName, InStr(1, MyName, "-", vbTextCompare) + 2, Len(MyName) - 8) + 1
                
                GoTo 8
                
            End If
                 
            If Left(UCase(Typ), 1) = "Q" And Left(MyName, 4) = "Q - " Then
                Calc_Next_Number = Mid(MyName, InStr(1, MyName, "-", vbTextCompare) + 2, Len(MyName) - 8) + 1
                
                GoTo 8
                
            End If
                
            GroupCount = GroupCount + 1
            
2:
        
        MyName = Dir
        
    Loop

8:

End Function

Public Function Confirm_Next_Number(Typ As String)
Dim FullFilePath As String, MyName As String
Dim GroupCount As Integer
'\* Check a Group folder exists
'FullFilePath = "C:\TEMP\Group*"

'fileextension = right(ucase(InputBox("Please enter file extension")),3)
fileextension = "*.*"

'ChDrive (main.Main_MasterPath & path)

    path = "templates"
    
    MyName = Dir(Main.Main_MasterPath.Value & path & "\", vbDirectory)
        If MyName = "" Then
            MsgBox "Folder Not Found", vbOKOnly, "Test"
                Exit Function
        End If
    '\* Store list of Group folder names
     
    Do Until MyName = ""
    
        If MyName = "." Or MyName = ".." Then GoTo 2
                
            If Left(UCase(Typ), 1) = "E" And Left(MyName, 4) = "E - " Then
                Confirm_Next_Number = Mid(MyName, InStr(1, MyName, "-", vbTextCompare) + 2, Len(MyName) - 8) + 1
                
                FileCopy Main.Main_MasterPath & path & "\" & MyName, Main.Main_MasterPath & path & "\" & "E - " & Confirm_Next_Number & ".TXT"
                Kill Main.Main_MasterPath & path & "\" & MyName

                GoTo 8
                
            End If
            
            If Left(UCase(Typ), 1) = "J" And Left(MyName, 4) = "J - " Then
                Confirm_Next_Number = Mid(MyName, InStr(1, MyName, "-", vbTextCompare) + 2, Len(MyName) - 8) + 1
                
                FileCopy Main.Main_MasterPath & path & "\" & MyName, Main.Main_MasterPath & path & "\" & "J - " & Confirm_Next_Number & ".TXT"
                Kill Main.Main_MasterPath & path & "\" & MyName

                GoTo 8
                
            End If
            
            If Left(UCase(Typ), 1) = "Q" And Left(MyName, 4) = "Q - " Then
                Confirm_Next_Number = Mid(MyName, InStr(1, MyName, "-", vbTextCompare) + 2, Len(MyName) - 8) + 1
                
                FileCopy Main.Main_MasterPath & path & "\" & MyName, Main.Main_MasterPath & path & "\" & "Q - " & Confirm_Next_Number & ".TXT"
                Kill Main.Main_MasterPath & path & "\" & MyName

                GoTo 8
                
            End If
                
            GroupCount = GroupCount + 1
            
2:
        
        MyName = Dir
        
    Loop

8:

End Function


