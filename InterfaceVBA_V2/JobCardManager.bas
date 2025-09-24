Attribute VB_Name = "JobCardManager"
' **Purpose**: Job card management and operations extracted from FJobCard.frm
' **Original**: Interface_VBA/FJobCard.frm business logic
' **CLAUDE.md Compliance**: Extract business logic from forms to modules
Option Explicit

' ===================================================================
' PUBLIC INTERFACE FUNCTIONS
' ===================================================================

' **Purpose**: Save current job card with form data
' **Original**: FJobCard.frm.SaveCurrentJobCard (lines 67-110+)
' **Parameters**:
'   - JobCardForm (Object): Form containing job card data
'   - CurrentJobPath (String): Path to current job file
' **Returns**: Boolean - True if save successful, False if failed
' **File Dependencies**: Job file from CurrentJobPath
' **Form Usage**: Extracted from FJobCard.frm to make it a thin wrapper
Public Function SaveJobCard(JobCardForm As Object, CurrentJobPath As String) As Boolean
    Dim JobInfo As CoreFramework.JobData

    On Error GoTo Error_Handler

    JobInfo = BusinessController.LoadJob(CurrentJobPath)
    If JobInfo.JobNumber = "" Then
        SaveJobCard = False
        Exit Function
    End If

    With JobInfo
        .AssignedOperator = Trim(JobCardForm.Assigned_Operator.Value)
        .Operations = GetOperationsFromForm(JobCardForm)
        .Notes = Trim(JobCardForm.Notes.Value)
        .Pictures = Trim(JobCardForm.Pictures.Value)

        If IsDate(JobCardForm.Due_Date.Value) Then
            .DueDate = CDate(JobCardForm.Due_Date.Value)
        End If

        If IsDate(JobCardForm.Workshop_Due_Date.Value) Then
            .WorkshopDueDate = CDate(JobCardForm.Workshop_Due_Date.Value)
        End If

        If IsDate(JobCardForm.Customer_Due_Date.Value) Then
            .CustomerDueDate = CDate(JobCardForm.Customer_Due_Date.Value)
        End If

        .Status = JobCardForm.Job_Status.Value
    End With

    SaveJobCard = BusinessController.UpdateJob(JobInfo)
    Exit Function

Error_Handler:
    CoreFramework.LogError Err.Number, "SaveJobCard", "JobCardManager", Err.Description
    SaveJobCard = False
End Function

' **Purpose**: Load job templates for operation selection
' **Original**: FJobCard.frm.LoadJobTemplates
' **Parameters**: JobCardForm (Object): Form to populate with template options
' **Returns**: Boolean - True if templates loaded successfully, False if failed
' **File Dependencies**: Template files in Job Templates directory
Public Function LoadJobTemplates(JobCardForm As Object) As Boolean
    Dim TemplatePath As String
    Dim TemplateFiles As Variant
    Dim i As Integer

    On Error GoTo Error_Handler

    TemplatePath = DataManager.GetRootPath & "\Job Templates"

    If Not DataManager.DirExists(TemplatePath) Then
        ValidationFramework.ShowWarning "Job Templates directory not found at: " & TemplatePath, "Templates Not Found"
        LoadJobTemplates = False
        Exit Function
    End If

    TemplateFiles = DataManager.GetFileList("Job Templates")

    If IsArray(TemplateFiles) Then
        ' Clear existing template list
        On Error Resume Next
        JobCardForm.TemplateList.Clear
        On Error GoTo Error_Handler

        ' Populate template list
        For i = 0 To UBound(TemplateFiles)
            If Right(TemplateFiles(i), 4) = ".xls" Then
                JobCardForm.TemplateList.AddItem Left(TemplateFiles(i), Len(TemplateFiles(i)) - 4)
            End If
        Next i
    End If

    LoadJobTemplates = True
    Exit Function

Error_Handler:
    CoreFramework.LogError Err.Number, "LoadJobTemplates", "JobCardManager", Err.Description
    LoadJobTemplates = False
End Function

' **Purpose**: Copy operations from another job card
' **Original**: FJobCard.frm.CopyOperationsFromJob
' **Parameters**:
'   - JobCardForm (Object): Form to populate with copied operations
'   - SourceJobNumber (String): Job number to copy from
' **Returns**: Boolean - True if copy successful, False if failed
' **File Dependencies**: Source job file
Public Function CopyOperationsFromJob(JobCardForm As Object, SourceJobNumber As String) As Boolean
    Dim SourceJobPath As String
    Dim SourceJobInfo As CoreFramework.JobData

    On Error GoTo Error_Handler

    SourceJobPath = FindJobFile(SourceJobNumber)
    If SourceJobPath = "" Then
        ValidationFramework.ShowWarning "Job " & SourceJobNumber & " not found.", "Job Not Found"
        CopyOperationsFromJob = False
        Exit Function
    End If

    SourceJobInfo = BusinessController.LoadJob(SourceJobPath)
    If SourceJobInfo.JobNumber = "" Then
        ValidationFramework.ShowWarning "Unable to load job data from " & SourceJobNumber, "Load Error"
        CopyOperationsFromJob = False
        Exit Function
    End If

    ' Copy operations to form
    PopulateFormWithOperations JobCardForm, SourceJobInfo.Operations

    ValidationFramework.ShowInformation "Operations copied successfully from Job " & SourceJobNumber, "Copy Complete"
    CopyOperationsFromJob = True
    Exit Function

Error_Handler:
    CoreFramework.LogError Err.Number, "CopyOperationsFromJob", "JobCardManager", Err.Description
    CopyOperationsFromJob = False
End Function

' **Purpose**: Add picture to job card
' **Original**: FJobCard.frm.AddPicture_Click (lines 51-65)
' **Parameters**: JobCardForm (Object): Form containing Pictures field
' **Returns**: Boolean - True if picture added successfully, False if cancelled/failed
Public Function AddPictureToJob(JobCardForm As Object) As Boolean
    Dim PicturePath As String

    On Error GoTo Error_Handler

    PicturePath = Application.GetOpenFilename("Image Files (*.jpg;*.jpeg;*.png;*.bmp),*.jpg;*.jpeg;*.png;*.bmp", , "Select Picture")

    If PicturePath <> "False" Then
        On Error Resume Next
        JobCardForm.Pictures.Value = JobCardForm.Pictures.Value & PicturePath & ";"
        On Error GoTo Error_Handler

        ValidationFramework.ShowInformation "Picture added to job.", "Picture Added"
        AddPictureToJob = True
    Else
        AddPictureToJob = False ' User cancelled
    End If
    Exit Function

Error_Handler:
    CoreFramework.LogError Err.Number, "AddPictureToJob", "JobCardManager", Err.Description
    AddPictureToJob = False
End Function

' **Purpose**: Load job data into job card form
' **Parameters**:
'   - JobCardForm (Object): Form to populate
'   - JobPath (String): Path to job file to load
' **Returns**: Boolean - True if load successful, False if failed
Public Function LoadJobCardData(JobCardForm As Object, JobPath As String) As Boolean
    Dim JobInfo As CoreFramework.JobData

    On Error GoTo Error_Handler

    JobInfo = BusinessController.LoadJob(JobPath)
    If JobInfo.JobNumber = "" Then
        LoadJobCardData = False
        Exit Function
    End If

    ' Populate form with job data
    With JobCardForm
        On Error Resume Next
        .Job_Number.Caption = JobInfo.JobNumber
        .Customer_Name.Caption = JobInfo.CustomerName
        .Component_Description.Caption = JobInfo.ComponentDescription
        .Component_Code.Caption = JobInfo.ComponentCode
        .Component_Quantity.Caption = CStr(JobInfo.Quantity)
        .Assigned_Operator.Value = JobInfo.AssignedOperator
        .Notes.Value = JobInfo.Notes
        .Pictures.Value = JobInfo.Pictures
        .Job_Status.Value = JobInfo.Status

        If JobInfo.DueDate <> 0 Then .Due_Date.Value = JobInfo.DueDate
        If JobInfo.WorkshopDueDate <> 0 Then .Workshop_Due_Date.Value = JobInfo.WorkshopDueDate
        If JobInfo.CustomerDueDate <> 0 Then .Customer_Due_Date.Value = JobInfo.CustomerDueDate
        On Error GoTo Error_Handler
    End With

    ' Load operations into form
    PopulateFormWithOperations JobCardForm, JobInfo.Operations

    LoadJobCardData = True
    Exit Function

Error_Handler:
    CoreFramework.LogError Err.Number, "LoadJobCardData", "JobCardManager", Err.Description
    LoadJobCardData = False
End Function

' ===================================================================
' PRIVATE HELPER FUNCTIONS
' ===================================================================

' **Purpose**: Find job file by job number
' **Parameters**: JobNumber (String): Job number to search for
' **Returns**: String - Full path to job file or empty string if not found
Private Function FindJobFile(JobNumber As String) As String
    Dim SearchPaths() As String
    Dim i As Integer
    Dim FilePath As String

    On Error GoTo Error_Handler

    ' Define directories to search
    ReDim SearchPaths(2)
    SearchPaths(0) = DataManager.GetRootPath & "\WIP"
    SearchPaths(1) = DataManager.GetRootPath & "\Archive"
    SearchPaths(2) = DataManager.GetRootPath & "\Quotes"

    For i = 0 To UBound(SearchPaths)
        FilePath = SearchPaths(i) & "\" & JobNumber & ".xls"
        If DataManager.FileExists(FilePath) Then
            FindJobFile = FilePath
            Exit Function
        End If
    Next i

    FindJobFile = ""
    Exit Function

Error_Handler:
    CoreFramework.LogError Err.Number, "FindJobFile", "JobCardManager", Err.Description
    FindJobFile = ""
End Function

' **Purpose**: Extract operations from form controls
' **Parameters**: JobCardForm (Object): Form containing operation controls
' **Returns**: String - Operations data formatted for storage
Private Function GetOperationsFromForm(JobCardForm As Object) As String
    Dim Operations As String
    Dim i As Integer
    Dim ControlName As String

    On Error GoTo Error_Handler

    For i = 1 To 15
        On Error Resume Next
        ControlName = "Operation" & Format(i, "00") & "_Type"
        If JobCardForm.Controls(ControlName).Value <> "" Then
            Operations = Operations & "Op" & i & ":" & JobCardForm.Controls(ControlName).Value

            ControlName = "Operation" & Format(i, "00") & "_Operator"
            If JobCardForm.Controls(ControlName).Value <> "" Then
                Operations = Operations & "|" & JobCardForm.Controls(ControlName).Value
            End If

            ControlName = "Operation" & Format(i, "00") & "_Comment"
            If JobCardForm.Controls(ControlName).Value <> "" Then
                Operations = Operations & "|" & JobCardForm.Controls(ControlName).Value
            End If

            Operations = Operations & ";"
        End If
        On Error GoTo Error_Handler
    Next i

    GetOperationsFromForm = Operations
    Exit Function

Error_Handler:
    CoreFramework.LogError Err.Number, "GetOperationsFromForm", "JobCardManager", Err.Description
    GetOperationsFromForm = ""
End Function

' **Purpose**: Populate form with operations data
' **Parameters**:
'   - JobCardForm (Object): Form to populate
'   - Operations (String): Operations data to populate from
Private Sub PopulateFormWithOperations(JobCardForm As Object, Operations As String)
    Dim OpParts() As String
    Dim OpDetails() As String
    Dim i As Integer
    Dim j As Integer

    On Error GoTo Error_Handler

    If Operations = "" Then Exit Sub

    OpParts = Split(Operations, ";")
    For i = 0 To UBound(OpParts)
        If OpParts(i) <> "" Then
            OpDetails = Split(OpParts(i), "|")
            If UBound(OpDetails) >= 0 Then
                ' Extract operation number from "Op1:" format
                If InStr(OpDetails(0), ":") > 0 Then
                    j = Val(Mid(OpDetails(0), 3, InStr(OpDetails(0), ":") - 3))

                    On Error Resume Next
                    If j >= 1 And j <= 15 Then
                        ' Set operation type
                        JobCardForm.Controls("Operation" & Format(j, "00") & "_Type").Value = _
                            Mid(OpDetails(0), InStr(OpDetails(0), ":") + 1)

                        ' Set operator if available
                        If UBound(OpDetails) >= 1 Then
                            JobCardForm.Controls("Operation" & Format(j, "00") & "_Operator").Value = OpDetails(1)
                        End If

                        ' Set comment if available
                        If UBound(OpDetails) >= 2 Then
                            JobCardForm.Controls("Operation" & Format(j, "00") & "_Comment").Value = OpDetails(2)
                        End If
                    End If
                    On Error GoTo Error_Handler
                End If
            End If
        End If
    Next i

    Exit Sub

Error_Handler:
    CoreFramework.LogError Err.Number, "PopulateFormWithOperations", "JobCardManager", Err.Description
End Sub