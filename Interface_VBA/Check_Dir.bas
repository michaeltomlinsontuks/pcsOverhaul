Attribute VB_Name = "Check_Dir"

' FUNCTION TO CHANGE DIRECTORY / CREATE DIRECTORY

Public Function CheckDir(Direc As String)

    If Dir(Direc, vbDirectory) = "" Then
        MkDir (Direc)
        ChDir (Direc)
    Else
        ChDir (Direc)
    End If
    
End Function

