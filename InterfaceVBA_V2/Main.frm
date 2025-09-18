VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Main
   Caption         =   "Main"
   ClientHeight    =   11040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14625
   OleObjectBlob   =   "Main.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
    ErrorHandler.HandleStandardErrors Err.Number, "Add_Enquiry_Click", "Main"
End Sub

Private Sub Archive_Click()
    On Error GoTo Error_Handler

    If Main.Archive.Value = True Then
        Main.lst.Clear
        PopulateFileList "Archive"
        ClearOtherButtons
    End If
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "Archive_Click", "Main"
End Sub

Private Sub Enquiries_Click()
    On Error GoTo Error_Handler

    If Main.Enquiries.Value = True Then
        Main.lst.Clear
        PopulateFileList "Enquiries"
        ClearOtherButtons
    End If
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "Enquiries_Click", "Main"
End Sub

Private Sub Quotes_Click()
    On Error GoTo Error_Handler

    If Main.Quotes.Value = True Then
        Main.lst.Clear
        PopulateFileList "Quotes"
        ClearOtherButtons
    End If
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "Quotes_Click", "Main"
End Sub

Private Sub WIP_Click()
    On Error GoTo Error_Handler

    If Main.WIP.Value = True Then
        Main.lst.Clear
        PopulateFileList "WIP"
        ClearOtherButtons
    End If
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "WIP_Click", "Main"
End Sub

Private Sub Make_Quote_Click()
    Dim SelectedFile As String
    Dim QuoteInfo As QuoteData

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
    ErrorHandler.HandleStandardErrors Err.Number, "Make_Quote_Click", "Main"
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
    ErrorHandler.HandleStandardErrors Err.Number, "createjob_Click", "Main"
End Sub

Private Sub JumpTheGun_Click()
    On Error GoTo Error_Handler

    With FJG
        .Show
    End With
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "JumpTheGun_Click", "Main"
End Sub

Private Sub ContractWork_Click()
    On Error GoTo Error_Handler

    FJG.but_SaveAsCTItem.Visible = False
    FJG.butSaveJG.Visible = True
    FJG.Show
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "ContractWork_Click", "Main"
End Sub

Private Sub but_CreateCTItem_Click()
    On Error GoTo Error_Handler

    FJG.but_SaveAsCTItem.Visible = True
    FJG.butSaveJG.Visible = False
    FJG.Show
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "but_CreateCTItem_Click", "Main"
End Sub

Private Sub but_EditCTItem_Click()
    Dim ContractFiles As Variant
    Dim SelectedContract As String

    On Error GoTo Error_Handler

    ContractFiles = FileManager.GetFileList("Contracts")
    If UBound(ContractFiles) = -1 Then
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
        ContractPath = FileManager.GetRootPath & "\Contracts\" & SelectedContract & ".xls"

        Dim wb As Workbook
        Set wb = FileManager.SafeOpenWorkbook(ContractPath)
        If Not wb Is Nothing Then
            Unload Me
        End If
    End If
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "but_EditCTItem_Click", "Main"
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
    ErrorHandler.HandleStandardErrors Err.Number, "OpenJob_Click", "Main"
End Sub

Private Sub WIPReport_Click()
    On Error GoTo Error_Handler

    With fwip
        .Show
    End With
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "WIPReport_Click", "Main"
End Sub

Private Sub OpenWIP_Click()
    On Error GoTo Error_Handler

    Dim WIPPath As String
    WIPPath = FileManager.GetRootPath & "\WIP.xls"

    Dim wb As Workbook
    Set wb = FileManager.SafeOpenWorkbook(WIPPath)
    If wb Is Nothing Then
        MsgBox "Could not open WIP database.", vbCritical
    End If
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "OpenWIP_Click", "Main"
End Sub

Private Sub Search_Click()
    On Error GoTo Error_Handler

    FrmSearchV2.Show
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "Search_Click", "Main"
End Sub

Private Sub butEditSearch_Click()
    On Error GoTo Error_Handler

    If SearchService.SortSearchDatabase() Then
        MsgBox "Search database sorted successfully.", vbInformation
    Else
        MsgBox "Failed to sort search database.", vbCritical
    End If

    Search_Click
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "butEditSearch_Click", "Main"
End Sub

Private Sub lst_Click()
    Dim SelectedFile As String
    Dim FilePath As String

    On Error GoTo Error_Handler

    SelectedFile = GetSelectedFileName()
    If SelectedFile = "" Then Exit Sub

    FilePath = GetCurrentDirectoryPath() & "\" & SelectedFile & ".xls"
    If FileManager.FileExists(FilePath) Then
        DisplayFileDetails FilePath
    End If
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "lst_Click", "Main"
End Sub

Private Sub Lst_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim SelectedFile As String
    Dim FilePath As String

    On Error GoTo Error_Handler

    SelectedFile = GetSelectedFileName()
    If SelectedFile = "" Then Exit Sub

    FilePath = GetCurrentDirectoryPath() & "\" & SelectedFile & ".xls"

    Dim wb As Workbook
    Set wb = FileManager.SafeOpenWorkbook(FilePath)
    If wb Is Nothing Then
        MsgBox "Could not open file: " & SelectedFile, vbCritical
    End If
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "Lst_DblClick", "Main"
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

        If JobController.CloseJob(JobNumber) Then
            MsgBox "Job " & JobNumber & " closed successfully.", vbInformation
            RefreshAllLists
        Else
            MsgBox "Failed to close job " & JobNumber & ".", vbCritical
        End If
    End If
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "CloseJob_Click", "Main"
End Sub

Private Function GetSelectedFileName() As String
    On Error GoTo Error_Handler

    If Main.lst.ListIndex >= 0 Then
        Dim SelectedValue As String
        SelectedValue = Main.lst.Value

        If InStr(1, SelectedValue, "*") > 1 Then
            GetSelectedFileName = Left(SelectedValue, Len(SelectedValue) - 2)
        Else
            GetSelectedFileName = SelectedValue
        End If
    Else
        GetSelectedFileName = ""
    End If
    Exit Function

Error_Handler:
    GetSelectedFileName = ""
End Function

Private Function GetCurrentDirectoryPath() As String
    If Main.Enquiries.Value Then
        GetCurrentDirectoryPath = FileManager.GetRootPath & "\Enquiries"
    ElseIf Main.Quotes.Value Then
        GetCurrentDirectoryPath = FileManager.GetRootPath & "\Quotes"
    ElseIf Main.WIP.Value Then
        GetCurrentDirectoryPath = FileManager.GetRootPath & "\WIP"
    ElseIf Main.Archive.Value Then
        GetCurrentDirectoryPath = FileManager.GetRootPath & "\Archive"
    Else
        GetCurrentDirectoryPath = FileManager.GetRootPath
    End If
End Function

Private Sub PopulateFileList(ByVal DirectoryName As String)
    Dim FileList As Variant
    Dim i As Integer

    On Error GoTo Error_Handler

    FileList = FileManager.GetFileList(DirectoryName)

    For i = 0 To UBound(FileList)
        Main.lst.AddItem Left(FileList(i), Len(FileList(i)) - 4)
    Next i
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "PopulateFileList", "Main"
End Sub

Private Sub ClearOtherButtons()
    Main.Enquiries.Value = (Main.Enquiries.Value And True)
    Main.WIP.Value = (Main.WIP.Value And True)
    Main.Archive.Value = (Main.Archive.Value And True)
    Main.Quotes.Value = (Main.Quotes.Value And True)
    Main.Thirties.Value = False
    Main.JobsInWIP.Value = False
End Sub

Private Sub RefreshAllLists()
    If Main.WIP.Value = True Then
        Main.WIP.Value = False
        Main.WIP.Value = True
    End If

    If Main.Enquiries.Value = True Then
        Main.Enquiries.Value = False
        Main.Enquiries.Value = True
    End If

    If Main.Archive.Value = True Then
        Main.Archive.Value = False
        Main.Archive.Value = True
    End If

    If Main.Quotes.Value = True Then
        Main.Quotes.Value = False
        Main.Quotes.Value = True
    End If
End Sub

Private Sub DisplayFileDetails(ByVal FilePath As String)
    Dim CustomerName As String
    Dim Description As String

    CustomerName = DataUtilities.GetValue(FilePath, "ADMIN", "B3")
    Description = DataUtilities.GetValue(FilePath, "ADMIN", "B8")

End Sub