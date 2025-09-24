Private Sub Add_Enquiry_Click()
    On Error GoTo Error_Handler

    With FrmEnquiry
        .Enquiry_Date.Caption = Format(Now(), "dd mmm yyyy")
        .Component_Code = ""
        .Component_Description = ""
        .Customer = ""
        .Component_Grade = ""
        .Component_Quantity = ""
        .Show
    End With

    RefreshAllLists
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "Add_Enquiry_Click", "Main"
End Sub

Private Sub Archive_Click()
    On Error GoTo Error_Handler

    If Main.Archive.Value = True Then
        InterfaceManager.PopulateMainFileList Me, "Archive"
        InterfaceManager.ClearSupplementaryButtons Me
    End If
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "Archive_Click", "Main"
End Sub

Private Sub Enquiries_Click()
    On Error GoTo Error_Handler

    If Main.Enquiries.Value = True Then
        InterfaceManager.PopulateMainFileList Me, "Enquiries"
        InterfaceManager.ClearSupplementaryButtons Me
    End If
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "Enquiries_Click", "Main"
End Sub

Private Sub Quotes_Click()
    On Error GoTo Error_Handler

    If Main.Quotes.Value = True Then
        InterfaceManager.PopulateMainFileList Me, "Quotes"
        InterfaceManager.ClearSupplementaryButtons Me
    End If
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "Quotes_Click", "Main"
End Sub

Private Sub WIP_Click()
    On Error GoTo Error_Handler

    If Main.WIP.Value = True Then
        InterfaceManager.PopulateMainFileList Me, "WIP"
        InterfaceManager.ClearSupplementaryButtons Me
    End If
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "WIP_Click", "Main"
End Sub

Private Sub Make_Quote_Click()
    Dim SelectedFile As String
    Dim QuoteInfo As CoreFramework.QuoteData

    On Error GoTo Error_Handler

    SelectedFile = GetSelectedFileName()
    If SelectedFile = "" Then
        MsgBox "Please select an enquiry to convert to quote.", vbInformation
        Exit Sub
    End If

    With QuoteInfo
        .UnitPrice = 0
        .TotalPrice = 0
        .LeadTime = ""
        .ValidUntil = DateAdd("d", 30, Now)
        .Status = "Pending"
    End With

    With FQuote
        .LoadFromEnquiry SelectedFile
        .Show
    End With
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "Make_Quote_Click", "Main"
End Sub

Private Sub createjob_Click()
    Dim SelectedFile As String

    On Error GoTo Error_Handler

    SelectedFile = GetSelectedFileName()
    If SelectedFile = "" Then
        MsgBox "Please select a quote to accept.", vbInformation
        Exit Sub
    End If

    With FAcceptQuote
        .LoadFromQuote SelectedFile
        .Show
    End With
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "createjob_Click", "Main"
End Sub

Private Sub JumpTheGun_Click()
    On Error GoTo Error_Handler

    With FJG
        .Show
    End With
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "JumpTheGun_Click", "Main"
End Sub

Private Sub ContractWork_Click()
    On Error GoTo Error_Handler

    FJG.Show
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "ContractWork_Click", "Main"
End Sub

Private Sub but_CreateCTItem_Click()
    On Error GoTo Error_Handler

    FJG.Show
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "but_CreateCTItem_Click", "Main"
End Sub

Private Sub but_EditCTItem_Click()
    Dim ContractFiles As Variant
    Dim SelectedContract As String

    On Error GoTo Error_Handler

    ContractFiles = DataManager.GetFileList("Contracts")
    If Not IsArray(ContractFiles) Or UBound(ContractFiles) = -1 Then
        MsgBox "No contract templates found.", vbInformation
        Exit Sub
    End If

    With FList
        .PopulateList ContractFiles
        .Show
    End With

    SelectedContract = FList.GetSelectedItem()
    If SelectedContract <> "" Then
        Dim ContractPath As String
        ContractPath = DataManager.GetRootPath & "\Contracts\" & SelectedContract & ".xls"

        Dim wb As Workbook
        Set wb = DataManager.SafeOpenWorkbook(ContractPath)
        If Not wb Is Nothing Then
            ' Contract file opened successfully - user can edit it directly
        End If
    End If
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "but_EditCTItem_Click", "Main"
End Sub

Private Sub OpenJob_Click()
    Dim SelectedFile As String

    On Error GoTo Error_Handler

    SelectedFile = GetSelectedFileName()
    If SelectedFile = "" Then
        MsgBox "Please select a job to open.", vbInformation
        Exit Sub
    End If

    With FJobCard
        .LoadJob SelectedFile
        .Show
    End With
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "OpenJob_Click", "Main"
End Sub

Private Sub WIPReport_Click()
    On Error GoTo Error_Handler

    With fwip
        .Show
    End With
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "WIPReport_Click", "Main"
End Sub

Private Sub OpenWIP_Click()
    On Error GoTo Error_Handler

    Dim WIPPath As String
    WIPPath = DataManager.GetRootPath & "\WIP.xls"

    Dim wb As Workbook
    Set wb = DataManager.SafeOpenWorkbook(WIPPath)
    If wb Is Nothing Then
        MsgBox "Could not open WIP database.", vbCritical
    End If
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "OpenWIP_Click", "Main"
End Sub

Private Sub Search_Click()
    On Error GoTo Error_Handler

    ' Open search database directly for legacy compatibility
    Dim SearchPath As String
    SearchPath = DataManager.GetRootPath & "\Search.xls"

    Dim wb As Workbook
    Set wb = DataManager.SafeOpenWorkbook(SearchPath)
    If wb Is Nothing Then
        MsgBox "Could not open Search database.", vbCritical
    End If
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "Search_Click", "Main"
End Sub

Private Sub butEditSearch_Click()
    On Error GoTo Error_Handler

    If SearchManager.SortSearchDatabase() Then
        MsgBox "Search database sorted successfully.", vbInformation
    Else
        MsgBox "Failed to sort search database.", vbCritical
    End If

    Search_Click
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "butEditSearch_Click", "Main"
End Sub

Private Sub lst_Click()
    Dim SelectedFile As String
    Dim FilePath As String

    On Error GoTo Error_Handler

    SelectedFile = InterfaceManager.GetSelectedFileName(Me)
    If SelectedFile = "" Then Exit Sub

    FilePath = InterfaceManager.GetCurrentDirectoryPath(Me) & "\" & SelectedFile & ".xls"
    If DataManager.FileExists(FilePath) Then
        InterfaceManager.DisplayMainFileDetails Me, FilePath
    End If
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "lst_Click", "Main"
End Sub

Private Sub Lst_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim SelectedFile As String
    Dim FilePath As String

    On Error GoTo Error_Handler

    SelectedFile = GetSelectedFileName()
    If SelectedFile = "" Then Exit Sub

    FilePath = GetCurrentDirectoryPath() & "\" & SelectedFile & ".xls"

    Dim wb As Workbook
    Set wb = DataManager.SafeOpenWorkbook(FilePath)
    If wb Is Nothing Then
        MsgBox "Could not open file: " & SelectedFile, vbCritical
    End If
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "Lst_DblClick", "Main"
End Sub

Private Sub CloseJob_Click()
    Dim SelectedFile As String
    Dim JobNumber As String

    On Error GoTo Error_Handler

    SelectedFile = GetSelectedFileName()
    If SelectedFile = "" Then
        MsgBox "Please select a job to close.", vbInformation
        Exit Sub
    End If

    If MsgBox("Are you sure you want to close job " & SelectedFile & "?", vbYesNo + vbQuestion) = vbYes Then
        JobNumber = SelectedFile

        If BusinessController.CloseJob(JobNumber) Then
            MsgBox "Job " & JobNumber & " closed successfully.", vbInformation
            RefreshAllLists
        Else
            MsgBox "Failed to close job " & JobNumber & ".", vbCritical
        End If
    End If
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "CloseJob_Click", "Main"
End Sub

' **Purpose**: Business logic extracted to InterfaceManager module
' **CLAUDE.md Compliance**: Form functions now call module functions

' **Purpose**: Private functions extracted to InterfaceManager module
' **CLAUDE.md Compliance**: PopulateFileList, ClearOtherButtons, RefreshAllLists moved to InterfaceManager

Private Sub DisplayFileDetails(ByVal FilePath As String)
    Dim CustomerName As String
    Dim Description As String

    On Error GoTo Error_Handler

    CustomerName = DataManager.GetValue(FilePath, "ADMIN", "B3")
    Description = DataManager.GetValue(FilePath, "ADMIN", "B8")

    ' Display the information in form labels or text boxes
    On Error Resume Next
    Main.lblCustomer.Caption = CustomerName
    Main.lblDescription.Caption = Description
    On Error GoTo Error_Handler

    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "DisplayFileDetails", "Main"
End Sub