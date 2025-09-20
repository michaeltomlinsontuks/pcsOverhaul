Attribute VB_Name = "BusinessLogic"
' **Purpose**: Main business operations and calculations for the PCS system
' **Consolidates**: a_Main.bas, RefreshMain.bas, Calc_Numbers.bas
' **Original Functionality**: All functions preserved exactly as original modules

Option Explicit

' **Purpose**: Show main menu (original functionality)
' **Parameters**: None
' **Returns**: None
' **Dependencies**: Main form, ActiveWorkbook.path
' **Side Effects**: Sets Main_MasterPath and shows Main form
' **Original Module**: a_Main.bas
Sub ShowMenu()
    Main.Main_MasterPath.Value = ActiveWorkbook.path & "\"
    Main.Show
End Sub

' **Purpose**: Refresh main interface (original functionality)
' **Parameters**: None
' **Returns**: None
' **Dependencies**: Main form, DirectoryHelpers.List_Files, DirectoryHelpers.Check_Files, DirectoryHelpers.CheckUpdates
' **Side Effects**: Clears and repopulates Main.lst, updates form controls
' **Original Module**: RefreshMain.bas
Public Function Refresh_Main()
    Main.lst.Clear

    If Main.Enquiries.Value = True Then
        x = DirectoryHelpers.List_Files("Enquiries", Main.lst)
    End If

    If Main.Quotes.Value = True Then
        x = DirectoryHelpers.List_Files("quotes", Main.lst)
        Main.Notice_Quotes.Caption = "Quotes : " & DirectoryHelpers.Check_Files(Main.Main_MasterPath & "Quotes\")
    End If

    If Main.WIP.Value = True Then
        x = DirectoryHelpers.List_Files("WIP", Main.lst)
    End If

    If Main.Archive.Value = True Then
        x = DirectoryHelpers.List_Files("Archive", Main.lst)
    End If

    If Main.Thirties.Value = True Then
        Main.Thirties.Value = False
        Main.Thirties.Value = True
    End If

        For Each ctl In Main.Controls
            If TypeName(ctl) = "Label" Then ctl.Caption = ""
            If UCase(TypeName(ctl)) = "TEXTBOX" And UCase(ctl.Name) <> "MAIN_MASTERPATH" Then ctl.Value = ""
        Next ctl

    DirectoryHelpers.CheckUpdates
End Function

' **Purpose**: Calculate next numbers for Enquiries, Quotes, or Jobs (original functionality)
' **Parameters**:
'   - Typ (String): Type prefix ("E" for Enquiries, "Q" for Quotes, "J" for Jobs)
' **Returns**: Long - Next available number for the specified type
' **Dependencies**: Main.Main_MasterPath, templates directory
' **Original Module**: Calc_Numbers.bas
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

' **Purpose**: Confirm and update next number by renaming template file (original functionality)
' **Parameters**:
'   - Typ (String): Type prefix ("E" for Enquiries, "Q" for Quotes, "J" for Jobs)
' **Returns**: Long - Confirmed next number with file operations completed
' **Dependencies**: Main.Main_MasterPath, templates directory, FileCopy, Kill
' **Side Effects**: Renames template files to increment number
' **Original Module**: Calc_Numbers.bas
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

' **Purpose**: Test function for number sequence operations (original functionality)
' **Parameters**: None
' **Returns**: None
' **Dependencies**: ActiveCell
' **Side Effects**: Modifies active cell values in sequence
' **Original Module**: a_Main.bas
Sub sadf()
    Do
        ActiveCell.Value = ActiveCell.Offset(-1, 0).Value - 1
        ActiveCell.Offset(1, 0).Select
    Loop Until ActiveCell.Offset(-1, 0).Value = 1011
End Sub