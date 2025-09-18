VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FJG
   Caption         =   "MEM: Jump The Gun"
   ClientHeight    =   9500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12000
   OleObjectBlob   =   "FJG.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FJG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private CurrentMode As String

Private Sub butSaveJG_Click()
    On Error GoTo Error_Handler

    If SaveDirectJob() Then
        MsgBox "Job created successfully.", vbInformation
        Unload Me
    End If
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "butSaveJG_Click", "FJG"
End Sub

Private Sub but_SaveAsCTItem_Click()
    On Error GoTo Error_Handler

    If SaveAsContract() Then
        MsgBox "Contract template saved successfully.", vbInformation
        Unload Me
    End If
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "but_SaveAsCTItem_Click", "FJG"
End Sub

Private Sub Cancel_Click()
    Unload Me
End Sub

Private Function SaveDirectJob() As Boolean
    Dim JobInfo As JobData
    Dim ValidationErrors As String

    On Error GoTo Error_Handler

    With JobInfo
        .CustomerName = Trim(Me.Customer.Value)
        .ComponentDescription = Trim(Me.Component_Description.Value)
        .ComponentCode = Trim(Me.Component_Code.Value)
        .MaterialGrade = Trim(Me.Component_Grade.Value)

        If IsNumeric(Me.Component_Quantity.Value) Then
            .Quantity = CLng(Me.Component_Quantity.Value)
        Else
            .Quantity = 0
        End If

        If IsDate(Me.Due_Date.Value) Then
            .DueDate = CDate(Me.Due_Date.Value)
        Else
            .DueDate = DateAdd("d", 14, Now)
        End If

        If IsDate(Me.Workshop_Due_Date.Value) Then
            .WorkshopDueDate = CDate(Me.Workshop_Due_Date.Value)
        Else
            .WorkshopDueDate = .DueDate
        End If

        If IsDate(Me.Customer_Due_Date.Value) Then
            .CustomerDueDate = CDate(Me.Customer_Due_Date.Value)
        Else
            .CustomerDueDate = .DueDate
        End If

        If IsNumeric(Me.Order_Value.Value) Then
            .OrderValue = CCur(Me.Order_Value.Value)
        Else
            .OrderValue = 0
        End If

        .AssignedOperator = Trim(Me.Assigned_Operator.Value)
        .Operations = GetOperationsString()
        .Notes = Trim(Me.Notes.Value)
        .Status = "Active"
    End With

    ValidationErrors = JobController.ValidateJobData(JobInfo)
    If ValidationErrors <> "" Then
        MsgBox "Please correct the following errors:" & vbCrLf & vbCrLf & ValidationErrors, vbExclamation
        SaveDirectJob = False
        Exit Function
    End If

    SaveDirectJob = JobController.CreateDirectJob(JobInfo)

    If SaveDirectJob Then
        Me.Job_Number.Value = JobInfo.JobNumber
        Me.File_Name.Value = JobInfo.JobNumber
        MsgBox "The Job Number is: " & JobInfo.JobNumber, vbInformation
    End If
    Exit Function

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "SaveDirectJob", "FJG"
    SaveDirectJob = False
End Function

Private Function SaveAsContract() As Boolean
    Dim ContractInfo As ContractData
    Dim ContractName As String

    On Error GoTo Error_Handler

    ContractName = Trim(InputBox("Enter contract template name:", "Contract Template"))
    If ContractName = "" Then
        SaveAsContract = False
        Exit Function
    End If

    With ContractInfo
        .ContractName = ContractName
        .CustomerName = Trim(Me.Customer.Value)
        .ComponentDescription = Trim(Me.Component_Description.Value)
        .StandardOperations = GetOperationsString()
        .LeadTime = "14 days"
        .FilePath = FileManager.GetRootPath & "\Contracts\" & ContractName & ".xls"
    End With

    SaveAsContract = CreateContractTemplate(ContractInfo)
    Exit Function

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "SaveAsContract", "FJG"
    SaveAsContract = False
End Function

Private Function CreateContractTemplate(ByVal ContractInfo As ContractData) As Boolean
    Dim TemplatePath As String
    Dim TemplateWB As Workbook

    On Error GoTo Error_Handler

    TemplatePath = FileManager.GetRootPath & "\Templates\_Enq.xls"

    Set TemplateWB = FileManager.SafeOpenWorkbook(TemplatePath)
    If TemplateWB Is Nothing Then
        CreateContractTemplate = False
        Exit Function
    End If

    PopulateContractTemplate TemplateWB, ContractInfo

    TemplateWB.SaveAs ContractInfo.FilePath
    FileManager.SafeCloseWorkbook TemplateWB

    CreateContractTemplate = True
    Exit Function

Error_Handler:
    If Not TemplateWB Is Nothing Then FileManager.SafeCloseWorkbook TemplateWB, False
    ErrorHandler.HandleStandardErrors Err.Number, "CreateContractTemplate", "FJG"
    CreateContractTemplate = False
End Function

Private Sub PopulateContractTemplate(ByVal wb As Workbook, ByVal ContractInfo As ContractData)
    Dim ws As Worksheet

    On Error GoTo Error_Handler

    Set ws = wb.Worksheets(1)

    With ws
        .Cells(2, 2).Value = "CONTRACT_TEMPLATE"
        .Cells(3, 2).Value = ContractInfo.CustomerName
        .Cells(8, 2).Value = ContractInfo.ComponentDescription
        .Cells(19, 2).Value = ContractInfo.StandardOperations
        .Cells(20, 2).Value = ContractInfo.LeadTime
        .Cells(21, 2).Value = Now
    End With

    Exit Sub

Error_Handler:
    ErrorHandler.LogError Err.Number, Err.Description, "PopulateContractTemplate", "FJG"
End Sub

Public Sub LoadFromContract(ByVal ContractFileName As String)
    Dim ContractPath As String
    Dim ContractInfo As ContractData

    On Error GoTo Error_Handler

    ContractPath = FileManager.GetRootPath & "\Contracts\" & ContractFileName & ".xls"

    If FileManager.FileExists(ContractPath) Then
        Dim wb As Workbook
        Set wb = FileManager.SafeOpenWorkbook(ContractPath)

        If Not wb Is Nothing Then
            Dim ws As Worksheet
            Set ws = wb.Worksheets(1)

            With Me
                .Customer.Value = ws.Cells(3, 2).Value
                .Component_Description.Value = ws.Cells(8, 2).Value
                .Operations_List.Value = ws.Cells(19, 2).Value
                .Due_Date.Value = Format(DateAdd("d", 14, Now), "dd/mm/yyyy")
                .Workshop_Due_Date.Value = Format(DateAdd("d", 14, Now), "dd/mm/yyyy")
                .Customer_Due_Date.Value = Format(DateAdd("d", 14, Now), "dd/mm/yyyy")
            End With

            FileManager.SafeCloseWorkbook wb, False
        End If
    End If
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "LoadFromContract", "FJG"
End Sub

Private Function GetOperationsString() As String
    Dim Operations As String

    On Error GoTo Error_Handler

    Operations = Trim(Me.Operations_List.Value)
    If Operations = "" Then
        Operations = "Standard Operations"
    End If

    GetOperationsString = Operations
    Exit Function

Error_Handler:
    GetOperationsString = "Standard Operations"
End Function

Private Sub LoadOperationTemplates()
    On Error GoTo Error_Handler

    Dim TemplatesPath As String
    TemplatesPath = FileManager.GetRootPath & "\Job Templates\Operations.xls"

    If FileManager.FileExists(TemplatesPath) Then
        Dim Operations As Variant
        Operations = DataUtilities.GetColumnData(TemplatesPath, "Sheet1", 1)

        If UBound(Operations) >= 0 Then
            Dim i As Integer
            For i = 0 To UBound(Operations)
                If Operations(i) <> "" Then
                    Me.Operations_List.AddItem Operations(i)
                End If
            Next i
        End If
    End If
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "LoadOperationTemplates", "FJG"
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
    ErrorHandler.HandleStandardErrors Err.Number, "Due_Date_Click", "FJG"
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
    ErrorHandler.HandleStandardErrors Err.Number, "Workshop_Due_Date_Click", "FJG"
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
    ErrorHandler.HandleStandardErrors Err.Number, "Customer_Due_Date_Click", "FJG"
End Sub

Private Function ShowCalendar() As Date
    On Error GoTo Error_Handler

    ShowCalendar = CDate(InputBox("Enter date (dd/mm/yyyy):", "Date Selection", Format(DateAdd("d", 14, Now), "dd/mm/yyyy")))
    Exit Function

Error_Handler:
    ShowCalendar = 0
End Function

Private Sub UserForm_Initialize()
    On Error GoTo Error_Handler

    Me.Due_Date.Value = Format(DateAdd("d", 14, Now), "dd/mm/yyyy")
    Me.Workshop_Due_Date.Value = Format(DateAdd("d", 14, Now), "dd/mm/yyyy")
    Me.Customer_Due_Date.Value = Format(DateAdd("d", 14, Now), "dd/mm/yyyy")

    LoadOperationTemplates
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "UserForm_Initialize", "FJG"
End Sub

Private Sub ClearForm()
    On Error GoTo Error_Handler

    Me.Job_Number.Value = ""
    Me.Customer.Value = ""
    Me.Component_Description.Value = ""
    Me.Component_Code.Value = ""
    Me.Component_Grade.Value = ""
    Me.Component_Quantity.Value = ""
    Me.Order_Value.Value = ""
    Me.Assigned_Operator.Value = ""
    Me.Operations_List.Value = ""
    Me.Notes.Value = ""
    Me.File_Name.Value = ""

    Me.Due_Date.Value = Format(DateAdd("d", 14, Now), "dd/mm/yyyy")
    Me.Workshop_Due_Date.Value = Format(DateAdd("d", 14, Now), "dd/mm/yyyy")
    Me.Customer_Due_Date.Value = Format(DateAdd("d", 14, Now), "dd/mm/yyyy")
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "ClearForm", "FJG"
End Sub

Private Sub Job_PicturePath_Change()
    On Error GoTo Error_Handler

    ' Handle picture path changes
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "Job_PicturePath_Change", "FJG"
End Sub

Private Sub Job_Urgency_Change()
    On Error GoTo Error_Handler

    ' Handle urgency changes
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "Job_Urgency_Change", "FJG"
End Sub