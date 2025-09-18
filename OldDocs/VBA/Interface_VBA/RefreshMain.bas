Attribute VB_Name = "RefreshMain"
Public Function Refresh_Main()

Main.lst.Clear

If Main.Enquiries.Value = True Then
    x = List_Files("Enquiries", Main.lst)
End If

If Main.Quotes.Value = True Then
    x = List_Files("quotes", Main.lst)
    Main.Notice_Quotes.Caption = "Quotes : " & Check_Files(Main.Main_MasterPath & "Quotes\")
End If

If Main.WIP.Value = True Then
    x = List_Files("WIP", Main.lst)
End If

If Main.Archive.Value = True Then
    x = List_Files("Archive", Main.lst)
End If

If Main.Thirties.Value = True Then
    Main.Thirties.Value = False
    Main.Thirties.Value = True
End If

    For Each ctl In Main.Controls
        If TypeName(ctl) = "Label" Then ctl.Caption = ""
        If UCase(TypeName(ctl)) = "TEXTBOX" And UCase(ctl.Name) <> "MAIN_MASTERPATH" Then ctl.Value = ""
    Next ctl

CheckUpdates

End Function

