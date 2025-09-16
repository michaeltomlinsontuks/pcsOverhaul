Attribute VB_Name = "RemoveCharacters"
Public Function Remove_Characters(Str As String)

For i = 1 To Len(Str)
    If Mid(Str, i, 1) = "/" Or Mid(Str, i, 1) = ":" Or Mid(Str, i, 1) = " " Then
        Str = Mid(Str, 1, i - 1) & Mid(Str, i + 1, Len(Str) - i)
    End If
Next i

Remove_Characters = Str

End Function

Public Function Insert_Characters(Str As String)

j = Len(Str)
i = 0

For i = 2 To j
    If Mid(Str, i, 1) = "_" Then
        Str = Mid(Str, 1, i - 1) & " " & Mid(Str, i + 1, Len(Str) - i)
        i = i + 1
    Else
        If UCase(Mid(Str, i, 1)) = Mid(Str, i, 1) Then
            Str = Mid(Str, 1, i - 1) & " " & Mid(Str, i, Len(Str) - i + 1)
            j = j + 1
            i = i + 1
        End If
    End If
Next i

If InStr(1, Str, "Component ", vbTextCompare) > 0 Then
    Str = Right(Str, Len(Str) - Len("Component "))
End If

Insert_Characters = Str

End Function


