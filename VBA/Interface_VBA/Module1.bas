Attribute VB_Name = "Module1"
Sub Update_Search()
Dim Files(1 To 100000) As String
Dim FullFilePath As String, MyName As String
Dim GroupCount As Integer
Dim Folder_Name As String
'\* Check a Group folder exists
'FullFilePath = "C:\TEMP\Group*"


Workbooks.Open ActiveWorkbook.path & "\search.xls", ReadOnly:=False

Range("A:A").Font.Bold = True

fileextension = "*.xls"

1:

'ChDrive (main.Main_MasterPath & path)



GoTo SkipHERE       ' commment this out to list all files in the folders


Range("3:35000").Clear


For i = 1 To 4
    Select Case i
        Case 1
            Folder_Name = ActiveWorkbook.path & "\Archive"
        
        Case 2
            Folder_Name = ActiveWorkbook.path & "\Enquiries"
            
        Case 3
            Folder_Name = ActiveWorkbook.path & "\Quotes"
        
        Case 4
            Folder_Name = ActiveWorkbook.path & "\WIP"
            
    End Select

            
            
            MyName = Dir(Folder_Name & "\", vbDirectory)
                If MyName = "" Then
                    MsgBox "Folder Not Found", vbOKOnly, "Test"
                        Exit Sub
                End If
            '\* Store list of Group folder names
            
            Do Until MyName = ""
            
                If MyName = "." Or MyName = ".." Then GoTo 2
            
                Range("A1").Select
'
                Do
                    ActiveCell.Offset(1, 0).Select
                Loop Until ActiveCell.Value = "" Or ActiveCell.Value = Left(MyName, Len(MyName) - 4)
                
                ActiveCell.Value = Left(MyName, Len(MyName) - 4)
                
2:
                MyName = Dir
                
            Loop
Next i
            
Range("a3").Select
            
SkipHERE:
' change below if you want to skip to a specific row
Range("A" & InputBox("Please adjust if you wish to move to a specific row", "SKIP TO ROW", ActiveCell.Row)).Select

Do
    Folder_Name = ActiveWorkbook.path & "\Archive\"
    If Dir(Folder_Name & ActiveCell.Value & ".xls", vbNormal) <> "" Then GoTo CopyInfo
    Folder_Name = ActiveWorkbook.path & "\Enquiries\"
    If Dir(Folder_Name & ActiveCell.Value & ".xls", vbNormal) <> "" Then GoTo CopyInfo
    Folder_Name = ActiveWorkbook.path & "\Quotes\"
    If Dir(Folder_Name & ActiveCell.Value & ".xls", vbNormal) <> "" Then GoTo CopyInfo
    Folder_Name = ActiveWorkbook.path & "\WIP\"
    If Dir(Folder_Name & ActiveCell.Value & ".xls", vbNormal) <> "" Then GoTo CopyInfo
        
    MsgBox ("CANT FIND THE FILE")
    End
CopyInfo:
    i = 0
    Do
        i = i + 1
        ItemType = GetValue(Folder_Name, ActiveCell.Value & ".xls", "Admin", "A" & i)
        ItemValue = GetValue(Folder_Name, ActiveCell.Value & ".xls", "Admin", "B" & i)
        
        j = 0
        Do
            j = j + 1
            If UCase(Range("a1").Offset(0, j).Value) = UCase(ItemType) Then
                If ActiveCell.Offset(0, j).Value = "" Or UCase(ActiveCell.Offset(0, j).Value) = UCase(ItemValue) Then
                    ActiveCell.Offset(0, j).Value = UCase(ItemValue)
                Else
                    If InStr(1, ItemType, "DATE", vbTextCompare) > 0 Then
                        If CCur(ActiveCell.Offset(0, j).Value) = CCur(ItemValue) Then
                            ActiveCell.Offset(0, j).Value = UCase(ItemValue)
                        Else
                            If MsgBox("A Difference Exists with regards to - " & ItemType & vbNewLine & "Do you wish to replace : " & ActiveCell.Offset(0, j).Value & " with : " & CDate(ItemValue), vbYesNo) = vbYes Then
                                ActiveCell.Offset(0, j).Value = UCase(ItemValue)
                            Else
                                If MsgBox("Do you wish to continue?", vbYesNo) = vbNo Then
                                    End
                                End If
                            End If
                        End If
                    Else
                        If MsgBox("A Difference Exists with regards to - " & ItemType & vbNewLine & "Do you wish to replace : " & ActiveCell.Offset(0, j).Value & " with : " & ItemValue, vbYesNo) = vbYes Then
                            ActiveCell.Offset(0, j).Value = UCase(ItemValue)
                        Else
                            If MsgBox("Do you wish to continue?", vbYesNo) = vbNo Then
                                End
                            End If
                        End If
                    End If
                End If
                Selection.Font.Bold = False
                GoTo NextType
            End If
        Loop Until Range("a1").Offset(0, j + 1).Value = ""
NextType:
    Loop Until ItemType = ""

EndCopy:
    ActiveCell.Offset(1, 0).Select
Loop Until ActiveCell.Value = ""
ActiveWorkbook.Close True

End Sub

Public Function GetValue(path, File, sheet, ref)
'   Retrieves a value from a closed workbook
    Dim arg As String
On Error GoTo Err
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
    If IsError(GetValue) Then GetValue = ""
    If GetValue = 0 Then GetValue = ""
On Error GoTo UnknownErr:
    Exit Function
Err:
    'MsgBox ("Admin Sheet does not exist!")
    
    Resume
    End

UnknownErr:
    MsgBox ("Unknown Error - Pls debug")
    Resume

End Function



