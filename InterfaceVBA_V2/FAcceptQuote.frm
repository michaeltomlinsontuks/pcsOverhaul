Private CurrentQuotePath As String

Private Sub butSAVE_Click()
    On Error GoTo Error_Handler

    If AcceptCurrentQuote() Then
        MsgBox "Quote accepted and job created successfully.", vbInformation
        Unload Me
    End If
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "butSAVE_Click", "FAcceptQuote"
End Sub

Private Sub Cancel_Click()
    Unload Me
End Sub

Private Function AcceptCurrentQuote() As Boolean
    Dim QuoteInfo As QuoteData
    Dim JobInfo As JobData
    Dim ValidationErrors As String

    On Error GoTo Error_Handler

    QuoteInfo = BusinessController.LoadQuote(CurrentQuotePath)
    If QuoteInfo.QuoteNumber = "" Then
        MsgBox "Could not load quote information.", vbCritical
        AcceptCurrentQuote = False
        Exit Function
    End If

    With JobInfo
        .QuoteNumber = QuoteInfo.QuoteNumber
        .CustomerName = QuoteInfo.CustomerName
        .ComponentDescription = QuoteInfo.ComponentDescription
        .ComponentCode = QuoteInfo.ComponentCode
        .MaterialGrade = QuoteInfo.MaterialGrade
        .Quantity = QuoteInfo.Quantity
        .OrderValue = QuoteInfo.TotalPrice

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

        .AssignedOperator = Trim(Me.Assigned_Operator.Value)
        .Operations = Trim(Me.Operations.Value)
        .Notes = Trim(Me.Notes.Value)
        .Status = "Active"
    End With

    ValidationErrors = BusinessController.ValidateJobData(JobInfo)
    If ValidationErrors <> "" Then
        MsgBox "Please correct the following errors:" & vbCrLf & vbCrLf & ValidationErrors, vbExclamation
        AcceptCurrentQuote = False
        Exit Function
    End If

    AcceptCurrentQuote = BusinessController.CreateJobFromQuote(JobInfo)

    If AcceptCurrentQuote Then
        Me.Job_Number.Value = JobInfo.JobNumber
        Me.File_Name.Value = JobInfo.JobNumber
        MsgBox "The Job Number is: " & JobInfo.JobNumber, vbInformation

        QuoteInfo.Status = "Accepted"
        BusinessController.UpdateQuote QuoteInfo
    End If
    Exit Function

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "AcceptCurrentQuote", "FAcceptQuote"
    AcceptCurrentQuote = False
End Function

Public Sub LoadFromQuote(ByVal QuoteFileName As String)
    Dim QuotePath As String
    Dim QuoteInfo As QuoteData

    On Error GoTo Error_Handler

    QuotePath = DataManager.GetRootPath & "\Quotes\" & QuoteFileName & ".xls"
    CurrentQuotePath = QuotePath

    QuoteInfo = BusinessController.LoadQuote(QuotePath)

    If QuoteInfo.QuoteNumber <> "" Then
        With Me
            .Quote_Number.Value = QuoteInfo.QuoteNumber
            .Enquiry_Number.Value = QuoteInfo.EnquiryNumber
            .Customer.Value = QuoteInfo.CustomerName
            .Component_Description.Value = QuoteInfo.ComponentDescription
            .Component_Code.Value = QuoteInfo.ComponentCode
            .Component_Grade.Value = QuoteInfo.MaterialGrade
            .Component_Quantity.Value = QuoteInfo.Quantity
            .Unit_Price.Value = QuoteInfo.UnitPrice
            .Total_Price.Value = QuoteInfo.TotalPrice
            .Lead_Time.Value = QuoteInfo.LeadTime

            .Due_Date.Value = Format(DateAdd("d", 14, Now), "dd/mm/yyyy")
            .Workshop_Due_Date.Value = Format(DateAdd("d", 14, Now), "dd/mm/yyyy")
            .Customer_Due_Date.Value = Format(DateAdd("d", 14, Now), "dd/mm/yyyy")

            .Assigned_Operator.Value = ""
            .Operations.Value = GetStandardOperations(QuoteInfo.ComponentCode)
            .Notes.Value = ""
        End With
    End If
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "LoadFromQuote", "FAcceptQuote"
End Sub

Private Function GetStandardOperations(ByVal ComponentCode As String) As String
    Dim OperationsPath As String
    Dim StandardOps As String

    On Error GoTo Error_Handler

    OperationsPath = DataManager.GetRootPath & "\Job Templates\Operations.xls"

    If DataManager.FileExists(OperationsPath) Then
        StandardOps = DataUtilities.GetValue(OperationsPath, "Sheet1", "A1")
        If StandardOps <> "" Then
            GetStandardOperations = StandardOps
        Else
            GetStandardOperations = "Standard Operations"
        End If
    Else
        GetStandardOperations = "Standard Operations"
    End If
    Exit Function

Error_Handler:
    GetStandardOperations = "Standard Operations"
End Function

Private Sub Due_Date_Click()
    On Error GoTo Error_Handler

    Dim SelectedDate As Date
    SelectedDate = ShowCalendar()

    If SelectedDate <> 0 Then
        Me.Due_Date.Value = Format(SelectedDate, "dd/mm/yyyy")
        AdjustOtherDates SelectedDate
    End If
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "Due_Date_Click", "FAcceptQuote"
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
    CoreFramework.HandleStandardErrors Err.Number, "Workshop_Due_Date_Click", "FAcceptQuote"
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
    CoreFramework.HandleStandardErrors Err.Number, "Customer_Due_Date_Click", "FAcceptQuote"
End Sub

Private Function ShowCalendar() As Date
    On Error GoTo Error_Handler

    ShowCalendar = CDate(InputBox("Enter date (dd/mm/yyyy):", "Date Selection", Format(DateAdd("d", 14, Now), "dd/mm/yyyy")))
    Exit Function

Error_Handler:
    ShowCalendar = 0
End Function

Private Sub AdjustOtherDates(ByVal DueDate As Date)
    On Error GoTo Error_Handler

    Dim WorkshopDate As Date
    Dim CustomerDate As Date

    WorkshopDate = DateAdd("d", -2, DueDate)
    CustomerDate = DueDate

    Me.Workshop_Due_Date.Value = Format(WorkshopDate, "dd/mm/yyyy")
    Me.Customer_Due_Date.Value = Format(CustomerDate, "dd/mm/yyyy")
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "AdjustOtherDates", "FAcceptQuote"
End Sub

Private Sub LoadOperators()
    On Error GoTo Error_Handler

    Dim OperatorsPath As String
    OperatorsPath = DataManager.GetRootPath & "\Templates\Operators.xls"

    If DataManager.FileExists(OperatorsPath) Then
        Dim Operators As Variant
        Operators = DataUtilities.GetColumnData(OperatorsPath, "Sheet1", 1)

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
    CoreFramework.HandleStandardErrors Err.Number, "LoadOperators", "FAcceptQuote"
End Sub

Private Sub LoadJobTemplates()
    On Error GoTo Error_Handler

    Dim TemplatesPath As String
    TemplatesPath = DataManager.GetRootPath & "\Job Templates\Standard_Operations.xls"

    If DataManager.FileExists(TemplatesPath) Then
        Dim Templates As Variant
        Templates = DataUtilities.GetColumnData(TemplatesPath, "Sheet1", 1)

        If UBound(Templates) >= 0 Then
            Dim i As Integer
            For i = 0 To UBound(Templates)
                If Templates(i) <> "" Then
                    Me.Operations.AddItem Templates(i)
                End If
            Next i
        End If
    End If
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "LoadJobTemplates", "FAcceptQuote"
End Sub

Private Sub UserForm_Initialize()
    On Error GoTo Error_Handler

    Me.Due_Date.Value = Format(DateAdd("d", 14, Now), "dd/mm/yyyy")
    Me.Workshop_Due_Date.Value = Format(DateAdd("d", 12, Now), "dd/mm/yyyy")
    Me.Customer_Due_Date.Value = Format(DateAdd("d", 14, Now), "dd/mm/yyyy")

    LoadOperators
    LoadJobTemplates
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "UserForm_Initialize", "FAcceptQuote"
End Sub

Private Sub ClearForm()
    On Error GoTo Error_Handler

    Me.Job_Number.Value = ""
    Me.Assigned_Operator.Value = ""
    Me.Operations.Value = ""
    Me.Notes.Value = ""
    Me.File_Name.Value = ""

    Me.Due_Date.Value = Format(DateAdd("d", 14, Now), "dd/mm/yyyy")
    Me.Workshop_Due_Date.Value = Format(DateAdd("d", 12, Now), "dd/mm/yyyy")
    Me.Customer_Due_Date.Value = Format(DateAdd("d", 14, Now), "dd/mm/yyyy")
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "ClearForm", "FAcceptQuote"
End Sub

Private Sub Job_Urgency_Change()
    On Error GoTo Error_Handler

    ' Handle urgency changes
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "Job_Urgency_Change", "FAcceptQuote"
End Sub