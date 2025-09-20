Attribute VB_Name = "SearchOperations"
' **Purpose**: Search functionality - keeping original simple approach
' **Consolidates**: Search_Sync.bas, Update_Search from Module1.bas
' **Original Functionality**: All functions preserved exactly as original modules

Option Explicit

' **Purpose**: Update search database with file information (original functionality from Module1)
' **Parameters**: None
' **Returns**: None
' **Dependencies**: GetValue function, search.xls file
' **Side Effects**: Updates search.xls with file information from Archive, Enquiries, Quotes, WIP folders
' **Original Module**: Module1.bas
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
            ItemType = CoreUtilities.GetValue(Folder_Name, ActiveCell.Value & ".xls", "Admin", "A" & i)
            ItemValue = CoreUtilities.GetValue(Folder_Name, ActiveCell.Value & ".xls", "Admin", "B" & i)

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

' **Purpose**: Search synchronization (original 93-line functionality)
' **Parameters**: None
' **Returns**: None
' **Dependencies**: Search.xls, Search History.xls, password validation
' **Side Effects**: Updates Search History, creates backups, cleans old records
' **Original Module**: Search_Sync.bas
Sub SeachSYNC()
    Dim DCSData(0 To 30) As Variant
    Dim DelDate As Date

    If InputBox("PASSWORD") <> "KJB" Then
        MsgBox ("ERROR - INCORRECT")
        End
    End If

    Workbooks.Open ActiveWorkbook.path & "\Search.xls"
    Range("A3").Select
    ActiveWorkbook.SaveCopyAs ActiveWorkbook.path & "\Backups\" & Format(Now(), "yyyymmdd") & " - Search.xls"

    Workbooks.Open ActiveWorkbook.path & "\Search History.xls"
    Range("A3").Select
    ActiveWorkbook.SaveCopyAs ActiveWorkbook.path & "\Backups\" & Format(Now(), "yyyymmdd") & " - Search History.xls"

    Do
        Windows("Search").Activate
        JC = False
        QN = False
        en = False

        If ActiveCell.Offset(0, 3).Value <> "" Then
            JC = True
            GoTo SHist
        End If
        If ActiveCell.Offset(0, 2).Value <> "" Then
            QN = True
            GoTo SHist
        End If
        en = True

SHist:
        For i = 0 To 30
            DCSData(i) = ActiveCell.Offset(0, i).Value
        Next i

        Windows("Search History").Activate

        Range("A2").Select
        Do
            ActiveCell.Offset(1, 0).Select
            If JC = True And ActiveCell.Offset(0, 3).Value = DCSData(3) Then GoTo FillDSCData
            If QN = True And ActiveCell.Offset(0, 2).Value = DCSData(2) Then GoTo FillDSCData
            If en = True And ActiveCell.Offset(0, 1).Value = DCSData(1) Then GoTo FillDSCData
        Loop Until ActiveCell.Value = ""

FillDSCData:
        For i = 0 To 30
            ActiveCell.Offset(0, i).Value = DCSData(i)
        Next i

        Windows("Search").Activate
            ActiveCell.Offset(1, 0).Select
    Loop Until ActiveCell.Value = ""

    Workbooks("Search History.xls").Save
    Workbooks("Search.xls").Save

    Range("c3").Select
    Main.Main_MasterPath = ActiveWorkbook.path & "\"

    Do
        If ActiveCell.Value <> "" Then

            If ActiveCell.Offset(0, 1).Value <> "" Then
                If CCur(ActiveCell.Offset(0, 2).Value) < BusinessLogic.Calc_Next_Number("J") - 1000 Then
                   Selection.EntireRow.Delete
                Else
                    ActiveCell.Offset(1, 0).Select
                End If
            Else

               If CCur(ActiveCell.Offset(0, 2).Value) < BusinessLogic.Calc_Next_Number("Q") - 10000 Then
                   Selection.EntireRow.Delete
                Else
                    ActiveCell.Offset(1, 0).Select
                End If

            End If
        Else
            ActiveCell.Offset(1, 0).Select
        End If
    Loop Until Range("A" & ActiveCell.Row).Value = ""

    ActiveWorkbook.Close True

    MsgBox ("COMPLETED")
End Sub