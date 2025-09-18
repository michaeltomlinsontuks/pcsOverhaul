Attribute VB_Name = "a_Main"
Sub ShowMenu()

Main.Main_MasterPath.Value = ActiveWorkbook.path & "\"
Main.Show

End Sub

Sub sadf()

Do
    ActiveCell.Value = ActiveCell.Offset(-1, 0).Value - 1
    ActiveCell.Offset(1, 0).Select
Loop Until ActiveCell.Offset(-1, 0).Value = 1011

End Sub
