Attribute VB_Name = "Module3"
Sub ExportAllModules()
    Dim vbComp As Object
    Dim exportPath As String
    
    ' Set your export folder here (make sure it exists or gets created)
    exportPath = "C:\Users\Michael Tomlinson\Downloads\20081222\Interface_VBA\"
    
    If Dir(exportPath, vbDirectory) = "" Then
        MkDir exportPath
    End If
    
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
            Case 1 ' Module
                vbComp.Export exportPath & vbComp.Name & ".bas"
            Case 2 ' Class Module
                vbComp.Export exportPath & vbComp.Name & ".cls"
            Case 3 ' UserForm
                vbComp.Export exportPath & vbComp.Name & ".frm"
            Case Else
                ' Ignore document modules
        End Select
    Next vbComp
    
    MsgBox "Export complete to: " & exportPath
End Sub

