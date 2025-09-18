Attribute VB_Name = "Search_Sync"
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
            If CCur(ActiveCell.Offset(0, 2).Value) < Calc_Next_Number("J") - 1000 Then
               Selection.EntireRow.Delete
            Else
                ActiveCell.Offset(1, 0).Select
            End If
        Else
        
           If CCur(ActiveCell.Offset(0, 2).Value) < Calc_Next_Number("Q") - 10000 Then
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
