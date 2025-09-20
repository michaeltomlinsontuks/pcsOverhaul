Attribute VB_Name = "DirectoryHelpers"
' **Purpose**: Directory and file listing operations for the PCS system
' **Consolidates**: Check_Dir.bas, a_ListFiles.bas, Delete_Sheet.bas, Check_Updates.bas
' **Original Functionality**: All functions preserved exactly as original modules

Option Explicit

Public NextCheck As Date

' **Purpose**: Check if directory exists and create/change to it (original functionality)
' **Parameters**:
'   - Direc (String): Directory path to check/create
' **Returns**: None
' **Dependencies**: Dir, MkDir, ChDir
' **Side Effects**: Creates directory if it doesn't exist, changes current directory
' **Original Module**: Check_Dir.bas
Public Function CheckDir(Direc As String)
    If Dir(Direc, vbDirectory) = "" Then
        MkDir (Direc)
        ChDir (Direc)
    Else
        ChDir (Direc)
    End If
End Function

' **Purpose**: List files in directory and populate form control (original functionality)
' **Parameters**:
'   - path (String): Subdirectory name relative to Main_MasterPath
'   - frm (Object): Form control to populate with file list
' **Returns**: None
' **Dependencies**: Main.Main_MasterPath, CoreUtilities.GetValue
' **Side Effects**: Populates form control with file listing, adds status indicators
' **Original Module**: a_ListFiles.bas
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
                If CoreUtilities.GetValue(Main.Main_MasterPath.Value & path & "\", x, "ADMIN", "b88") = UCase("Quote Accepted") Then
                    .AddItem Left(x, Len(x) - 4) & " *"
                Else
                    .AddItem Left(x, Len(x) - 4)
                End If
            Else
                If path = "quotes" Then
                    If CoreUtilities.GetValue(Main.Main_MasterPath & path & "\", x, "Admin", "b88") = "New Quote" Then
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

' **Purpose**: Delete sheet without confirmation prompt (original functionality)
' **Parameters**:
'   - SheetName (String): Name of sheet to delete
' **Returns**: None
' **Dependencies**: Application.DisplayAlerts, Worksheets.Delete
' **Side Effects**: Deletes specified worksheet without user prompt
' **Original Module**: Delete_Sheet.bas
Public Function DeleteSheet(SheetName As String)
    Application.DisplayAlerts = False
    Worksheets(SheetName).Delete
    Application.DisplayAlerts = True
End Function

' **Purpose**: Check for file count updates and update UI (original functionality)
' **Parameters**: None
' **Returns**: None
' **Dependencies**: Main form, Check_Files function, Application.OnTime
' **Side Effects**: Updates notice captions on Main form, schedules next check
' **Original Module**: Check_Updates.bas
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

' **Purpose**: Stop scheduled file checking (original functionality)
' **Parameters**: None
' **Returns**: None
' **Dependencies**: Application.OnTime
' **Side Effects**: Cancels scheduled CheckUpdates execution
' **Original Module**: Check_Updates.bas
Public Function StopCheck()
    On Error Resume Next
    Application.OnTime NextCheck, "CheckUpdates", , Schedule:=False
End Function

' **Purpose**: Count files in specified directory (original functionality)
' **Parameters**:
'   - path (String): Directory path to count files in
' **Returns**: Integer - Number of files found (excluding system files)
' **Dependencies**: Dir function
' **Original Module**: Check_Updates.bas
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