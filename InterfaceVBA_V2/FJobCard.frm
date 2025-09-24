Private CurrentJobPath As String

Private Sub SaveJobCard_Click()
    On Error GoTo Error_Handler

    If JobCardManager.SaveJobCard(Me, CurrentJobPath) Then
        ValidationFramework.ShowInformation "Job card saved successfully.", "Save Complete"
    End If
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "SaveJobCard_Click", "FJobCard"
End Sub

Private Sub CloseJobCard_Click()
    Unload Me
End Sub

Private Sub JobCardTemplates_Click()
    On Error GoTo Error_Handler

    JobCardManager.LoadJobTemplates Me
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "JobCardTemplates_Click", "FJobCard"
End Sub

Private Sub CopyFromJobCard_Click()
    Dim SourceJobNumber As String

    On Error GoTo Error_Handler

    SourceJobNumber = InputBox("Enter job number to copy operations from:", "Copy Operations")
    If SourceJobNumber = "" Then Exit Sub

    JobCardManager.CopyOperationsFromJob Me, SourceJobNumber
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "CopyFromJobCard_Click", "FJobCard"
End Sub

Private Sub AddPicture_Click()
    On Error GoTo Error_Handler

    JobCardManager.AddPictureToJob Me
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "AddPicture_Click", "FJobCard"
End Sub

' **Purpose**: Business logic extracted to JobCardManager module
    Dim JobInfo As CoreFramework.JobData

    On Error GoTo Error_Handler

    JobInfo = BusinessController.LoadJob(CurrentJobPath)
    If JobInfo.JobNumber = "" Then
        SaveCurrentJobCard = False
        Exit Function
    End If

    With JobInfo
        .AssignedOperator = Trim(Me.Assigned_Operator.Value)
        .Operations = GetOperationsFromForm()
        .Notes = Trim(Me.Notes.Value)
        .Pictures = Trim(Me.Pictures.Value)

        If IsDate(Me.Due_Date.Value) Then
            .DueDate = CDate(Me.Due_Date.Value)
        End If

        If IsDate(Me.Workshop_Due_Date.Value) Then
            .WorkshopDueDate = CDate(Me.Workshop_Due_Date.Value)
        End If

        If IsDate(Me.Customer_Due_Date.Value) Then
            .CustomerDueDate = CDate(Me.Customer_Due_Date.Value)
        End If

        .Status = Me.Job_Status.Value
    End With

    SaveCurrentJobCard = BusinessController.UpdateJob(JobInfo)

    ' Update search database for legacy compatibility
    If SaveCurrentJobCard Then
        On Error Resume Next
        SearchManager.SaveRowIntoSearch Me
        On Error GoTo Error_Handler
    End If
    Exit Function

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "SaveCurrentJobCard", "FJobCard"
    SaveCurrentJobCard = False
End Function

Public Sub LoadJob(ByVal JobFileName As String)
    Dim JobPath As String
    Dim JobInfo As CoreFramework.JobData

    On Error GoTo Error_Handler

    JobPath = DataManager.GetRootPath & "\WIP\" & JobFileName & ".xls"
    CurrentJobPath = JobPath

    JobInfo = BusinessController.LoadJob(JobPath)

    If JobInfo.JobNumber <> "" Then
        With Me
            .Job_Number.Value = JobInfo.JobNumber
            .Customer.Value = JobInfo.CustomerName
            .Component_Description.Value = JobInfo.ComponentDescription
            .Component_Code.Value = JobInfo.ComponentCode
            .Component_Grade.Value = JobInfo.MaterialGrade
            .Component_Quantity.Value = JobInfo.Quantity
            .Order_Value.Value = JobInfo.OrderValue
            .Due_Date.Value = Format(JobInfo.DueDate, "dd/mm/yyyy")
            .Workshop_Due_Date.Value = Format(JobInfo.WorkshopDueDate, "dd/mm/yyyy")
            .Customer_Due_Date.Value = Format(JobInfo.CustomerDueDate, "dd/mm/yyyy")
            .Assigned_Operator.Value = JobInfo.AssignedOperator
            .Job_Status.Value = JobInfo.Status
            .Notes.Value = JobInfo.Notes
            .Pictures.Value = JobInfo.Pictures

            PopulateOperationsToForm JobInfo.Operations
        End With
    End If
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "LoadJob", "FJobCard"
End Sub

' **Purpose**: PopulateOperationsToForm extracted to JobCardManager
    On Error GoTo Error_Handler

    If Operations <> "" Then
        Dim OpArray As Variant
        OpArray = Split(Operations, ";")

        Dim i As Integer
        For i = 0 To UBound(OpArray)
            If i < 15 And Trim(OpArray(i)) <> "" Then
                Me.Controls("Operation" & (i + 1)).Value = Trim(OpArray(i))
            End If
        Next i
    End If
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "PopulateOperationsToForm", "FJobCard"
End Sub

' **Purpose**: GetOperationsFromForm extracted to JobCardManager
    Dim Operations As String

    On Error GoTo Error_Handler

    Operations = ""
    Dim i As Integer
    Dim OperationValue As String

    For i = 1 To 15
        OperationValue = Trim(Me.Controls("Operation" & i).Value)
        If OperationValue <> "" Then
            Operations = Operations & OperationValue & ";"
        End If
    Next i

    If Len(Operations) > 0 Then
        Operations = Left(Operations, Len(Operations) - 1)
    End If

    GetOperationsFromForm = Operations
    Exit Function

Error_Handler:
    GetOperationsFromForm = ""
End Function

' **Purpose**: LoadJobTemplates extracted to JobCardManager
    Dim TemplatesPath As String
    Dim Templates As Variant

    On Error GoTo Error_Handler

    TemplatesPath = DataManager.GetRootPath & "\Job Templates\Operations.xls"

    If DataManager.FileExists(TemplatesPath) Then
        ' Use DataManager for consistency
        On Error Resume Next
        Templates = DataManager.GetRangeValues(TemplatesPath, "Sheet1", "A:A")
        On Error GoTo Error_Handler

        If UBound(Templates) >= 0 Then
            Dim TemplateList As String
            Dim i As Integer

            TemplateList = "Available Operation Templates:" & vbCrLf & vbCrLf
            For i = 0 To UBound(Templates)
                If Templates(i) <> "" Then
                    TemplateList = TemplateList & Templates(i) & vbCrLf
                End If
            Next i

            MsgBox TemplateList, vbInformation, "Job Templates"
        End If
    Else
        MsgBox "No job templates found.", vbInformation
    End If
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "LoadJobTemplates", "FJobCard"
End Sub

' **Purpose**: FindJobFile extracted to JobCardManager
    Dim WIPPath As String
    Dim ArchivePath As String

    On Error GoTo Error_Handler

    WIPPath = DataManager.GetRootPath & "\WIP\" & JobNumber & ".xls"
    ArchivePath = DataManager.GetRootPath & "\Archive\" & JobNumber & ".xls"

    If DataManager.FileExists(WIPPath) Then
        FindJobFile = WIPPath
    ElseIf DataManager.FileExists(ArchivePath) Then
        FindJobFile = ArchivePath
    Else
        FindJobFile = ""
    End If
    Exit Function

Error_Handler:
    FindJobFile = ""
End Function

' **Purpose**: CopyOperationsFromJob extracted to JobCardManager
    Dim SourceJobInfo As CoreFramework.JobData

    On Error GoTo Error_Handler

    SourceJobInfo = BusinessController.LoadJob(SourceJobPath)

    If SourceJobInfo.JobNumber <> "" Then
        PopulateOperationsToForm SourceJobInfo.Operations
        Me.Notes.Value = Me.Notes.Value & vbCrLf & "Operations copied from job: " & SourceJobInfo.JobNumber
    End If
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "CopyOperationsFromJob", "FJobCard"
End Sub

Private Sub Due_Date_Click()
    On Error GoTo Error_Handler

    Dim SelectedDate As Date
    SelectedDate = ShowCalendar()

    If SelectedDate <> 0 Then
        Me.Due_Date.Value = Format(SelectedDate, "dd/mm/yyyy")
    End If
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "Due_Date_Click", "FJobCard"
End Sub

Private Sub Workshop_Due_Date_Click()
    On Error GoTo Error_Handler

    Dim SelectedDate As Date
    SelectedDate = ShowCalendar()

    If SelectedDate <> 0 Then
        Me.Workshop_Due_Date.Value = Format(SelectedDate, "dd/mm/yyyy")
    End If
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "Workshop_Due_Date_Click", "FJobCard"
End Sub

Private Sub Customer_Due_Date_Click()
    On Error GoTo Error_Handler

    Dim SelectedDate As Date
    SelectedDate = ShowCalendar()

    If SelectedDate <> 0 Then
        Me.Customer_Due_Date.Value = Format(SelectedDate, "dd/mm/yyyy")
    End If
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "Customer_Due_Date_Click", "FJobCard"
End Sub

Private Function ShowCalendar() As Date
    On Error GoTo Error_Handler

    ShowCalendar = CDate(InputBox("Enter date (dd/mm/yyyy):", "Date Selection", Format(Now, "dd/mm/yyyy")))
    Exit Function

Error_Handler:
    ShowCalendar = 0
End Function

Private Sub UserForm_Initialize()
    On Error GoTo Error_Handler

    LoadJobStatusOptions
    LoadOperatorOptions
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "UserForm_Initialize", "FJobCard"
End Sub

Private Sub LoadJobStatusOptions()
    Me.Job_Status.AddItem "Active"
    Me.Job_Status.AddItem "On Hold"
    Me.Job_Status.AddItem "Completed"
    Me.Job_Status.AddItem "Cancelled"
End Sub

Private Sub LoadOperatorOptions()
    Dim OperatorsPath As String

    On Error GoTo Error_Handler

    OperatorsPath = DataManager.GetRootPath & "\Templates\Operators.xls"

    If DataManager.FileExists(OperatorsPath) Then
        Dim Operators As Variant
        ' Use DataManager for consistency
        On Error Resume Next
        Operators = DataManager.GetRangeValues(OperatorsPath, "Sheet1", "A:A")
        On Error GoTo Error_Handler

        If UBound(Operators) >= 0 Then
            Dim i As Integer
            For i = 0 To UBound(Operators)
                If Operators(i) <> "" Then
                    Me.Assigned_Operator.AddItem Operators(i)
                End If
            Next i
        End If
    End If
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "LoadOperatorOptions", "FJobCard"
End Sub

Private Sub Job_PicturePath_Change()
    On Error GoTo Error_Handler

    ' Handle picture path changes
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "Job_PicturePath_Change", "FJobCard"
End Sub