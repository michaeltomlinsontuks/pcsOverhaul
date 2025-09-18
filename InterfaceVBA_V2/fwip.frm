VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fwip
   Caption         =   "MEM: WIP Reports"
   ClientHeight    =   8000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10000
   OleObjectBlob   =   "fwip.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fwip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Go_Click()
    On Error GoTo Error_Handler

    If GenerateSelectedReports() Then
        MsgBox "Reports generated successfully.", vbInformation
        Unload Me
    End If
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "Go_Click", "fwip"
End Sub

Private Sub Cancel_Click()
    Unload Me
End Sub

Private Function GenerateSelectedReports() As Boolean
    Dim ReportsGenerated As Boolean

    On Error GoTo Error_Handler

    ReportsGenerated = False

    If Me.Operation_Reports.Value = True Then
        If WIPManager.GenerateWIPReport("OPERATION") Then
            ReportsGenerated = True
        End If
    End If

    If Me.Operator_Reports.Value = True Then
        If WIPManager.GenerateWIPReport("OPERATOR") Then
            ReportsGenerated = True
        End If
    End If

    If Me.Customer_Reports_Office.Value = True Then
        If WIPManager.GenerateWIPReport("CUSTOMER", "OFFICE") Then
            ReportsGenerated = True
        End If
    End If

    If Me.Customer_Reports_Workshop.Value = True Then
        If WIPManager.GenerateWIPReport("CUSTOMER", "WORKSHOP") Then
            ReportsGenerated = True
        End If
    End If

    If Me.Due_Date_Reports.Value = True Then
        Dim DueDateFilter As String
        DueDateFilter = Format(DateAdd("d", 7, Now), "dd/mm/yyyy")
        If WIPManager.GenerateWIPReport("DUEDATE", DueDateFilter) Then
            ReportsGenerated = True
        End If
    End If

    If Me.Job_Number_Reports_Office.Value = True Then
        If WIPManager.GenerateWIPReport("JOBNUMBER", "OFFICE") Then
            ReportsGenerated = True
        End If
    End If

    If Me.Job_Number_Reports_Workshop.Value = True Then
        If WIPManager.GenerateWIPReport("JOBNUMBER", "WORKSHOP") Then
            ReportsGenerated = True
        End If
    End If

    If Me.Custom_Report.Value = True Then
        If GenerateCustomReport() Then
            ReportsGenerated = True
        End If
    End If

    GenerateSelectedReports = ReportsGenerated
    Exit Function

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "GenerateSelectedReports", "fwip"
    GenerateSelectedReports = False
End Function

Private Function GenerateCustomReport() As Boolean
    Dim CustomerFilter As String
    Dim OperatorFilter As String
    Dim ReportType As String

    On Error GoTo Error_Handler

    CustomerFilter = Trim(Me.Customer_Filter.Value)
    OperatorFilter = Trim(Me.Operator_Filter.Value)

    If CustomerFilter <> "" Then
        ReportType = "CUSTOMER"
        GenerateCustomReport = WIPManager.GenerateWIPReport(ReportType, CustomerFilter)
    ElseIf OperatorFilter <> "" Then
        ReportType = "OPERATOR"
        GenerateCustomReport = WIPManager.GenerateWIPReport(ReportType, OperatorFilter)
    Else
        GenerateCustomReport = WIPManager.GenerateWIPReport("ALL")
    End If
    Exit Function

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "GenerateCustomReport", "fwip"
    GenerateCustomReport = False
End Function

Private Sub SelectAll_Click()
    On Error GoTo Error_Handler

    Me.Operation_Reports.Value = True
    Me.Operator_Reports.Value = True
    Me.Customer_Reports_Office.Value = True
    Me.Customer_Reports_Workshop.Value = True
    Me.Due_Date_Reports.Value = True
    Me.Job_Number_Reports_Office.Value = True
    Me.Job_Number_Reports_Workshop.Value = True
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "SelectAll_Click", "fwip"
End Sub

Private Sub ClearAll_Click()
    On Error GoTo Error_Handler

    Me.Operation_Reports.Value = False
    Me.Operator_Reports.Value = False
    Me.Customer_Reports_Office.Value = False
    Me.Customer_Reports_Workshop.Value = False
    Me.Due_Date_Reports.Value = False
    Me.Job_Number_Reports_Office.Value = False
    Me.Job_Number_Reports_Workshop.Value = False
    Me.Custom_Report.Value = False
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "ClearAll_Click", "fwip"
End Sub

Private Sub ViewWIPDatabase_Click()
    On Error GoTo Error_Handler

    Dim WIPPath As String
    WIPPath = FileManager.GetRootPath & "\WIP.xls"

    Dim wb As Workbook
    Set wb = FileManager.SafeOpenWorkbook(WIPPath)
    If wb Is Nothing Then
        MsgBox "Could not open WIP database.", vbCritical
    Else
        Me.Hide
    End If
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "ViewWIPDatabase_Click", "fwip"
End Sub

Private Sub RefreshWIPData_Click()
    On Error GoTo Error_Handler

    If RefreshWIPFromFiles() Then
        MsgBox "WIP database refreshed successfully.", vbInformation
    Else
        MsgBox "Failed to refresh WIP database.", vbCritical
    End If
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "RefreshWIPData_Click", "fwip"
End Sub

Private Function RefreshWIPFromFiles() As Boolean
    Dim WIPFiles As Variant
    Dim i As Integer
    Dim JobInfo As JobData
    Dim RefreshCount As Integer

    On Error GoTo Error_Handler

    WIPFiles = FileManager.GetFileList("WIP")
    RefreshCount = 0

    For i = 0 To UBound(WIPFiles)
        Dim JobPath As String
        JobPath = FileManager.GetRootPath & "\WIP\" & WIPFiles(i)

        JobInfo = JobController.LoadJob(JobPath)
        If JobInfo.JobNumber <> "" Then
            If WIPManager.UpdateJobInWIP(JobInfo) Then
                RefreshCount = RefreshCount + 1
            End If
        End If
    Next i

    If RefreshCount > 0 Then
        RefreshWIPFromFiles = True
        MsgBox RefreshCount & " jobs refreshed in WIP database.", vbInformation
    Else
        RefreshWIPFromFiles = False
    End If
    Exit Function

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "RefreshWIPFromFiles", "fwip"
    RefreshWIPFromFiles = False
End Function

Private Sub LoadCustomerList()
    Dim WIPJobs As Variant
    Dim i As Integer
    Dim UniqueCustomers As Collection
    Dim Customer As Variant

    On Error GoTo Error_Handler

    Set UniqueCustomers = New Collection
    WIPJobs = WIPManager.GetWIPJobs()

    For i = 0 To UBound(WIPJobs)
        On Error Resume Next
        UniqueCustomers.Add WIPJobs(i).CustomerName, WIPJobs(i).CustomerName
        On Error GoTo Error_Handler
    Next i

    Me.Customer_Filter.Clear
    For Each Customer In UniqueCustomers
        Me.Customer_Filter.AddItem Customer
    Next Customer
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "LoadCustomerList", "fwip"
End Sub

Private Sub LoadOperatorList()
    Dim WIPJobs As Variant
    Dim i As Integer
    Dim UniqueOperators As Collection
    Dim Operator As Variant

    On Error GoTo Error_Handler

    Set UniqueOperators = New Collection
    WIPJobs = WIPManager.GetWIPJobs()

    For i = 0 To UBound(WIPJobs)
        If WIPJobs(i).AssignedOperator <> "" Then
            On Error Resume Next
            UniqueOperators.Add WIPJobs(i).AssignedOperator, WIPJobs(i).AssignedOperator
            On Error GoTo Error_Handler
        End If
    Next i

    Me.Operator_Filter.Clear
    For Each Operator In UniqueOperators
        Me.Operator_Filter.AddItem Operator
    Next Operator
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "LoadOperatorList", "fwip"
End Sub

Private Sub UpdateJobCounts()
    Dim WIPJobs As Variant
    Dim ActiveJobs As Integer
    Dim OnHoldJobs As Integer
    Dim OverdueJobs As Integer
    Dim i As Integer

    On Error GoTo Error_Handler

    WIPJobs = WIPManager.GetWIPJobs()
    ActiveJobs = 0
    OnHoldJobs = 0
    OverdueJobs = 0

    For i = 0 To UBound(WIPJobs)
        Select Case UCase(WIPJobs(i).Status)
            Case "ACTIVE"
                ActiveJobs = ActiveJobs + 1
                If WIPJobs(i).DueDate < Date Then
                    OverdueJobs = OverdueJobs + 1
                End If
            Case "ON HOLD", "ONHOLD"
                OnHoldJobs = OnHoldJobs + 1
        End Select
    Next i

    Me.Active_Jobs_Count.Caption = "Active Jobs: " & ActiveJobs
    Me.OnHold_Jobs_Count.Caption = "On Hold: " & OnHoldJobs
    Me.Overdue_Jobs_Count.Caption = "Overdue: " & OverdueJobs
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "UpdateJobCounts", "fwip"
End Sub

Private Sub UserForm_Initialize()
    On Error GoTo Error_Handler

    LoadCustomerList
    LoadOperatorList
    UpdateJobCounts

    Me.Operation_Reports.Value = False
    Me.Operator_Reports.Value = False
    Me.Customer_Reports_Office.Value = False
    Me.Customer_Reports_Workshop.Value = False
    Me.Due_Date_Reports.Value = False
    Me.Job_Number_Reports_Office.Value = False
    Me.Job_Number_Reports_Workshop.Value = False
    Me.Custom_Report.Value = False
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "UserForm_Initialize", "fwip"
End Sub

Private Sub PreviewReport_Click()
    Dim ReportType As String
    Dim FilterValue As String

    On Error GoTo Error_Handler

    If Me.Operation_Reports.Value = True Then
        ReportType = "OPERATION"
    ElseIf Me.Operator_Reports.Value = True Then
        ReportType = "OPERATOR"
    ElseIf Me.Customer_Reports_Office.Value = True Then
        ReportType = "CUSTOMER"
        FilterValue = "OFFICE"
    ElseIf Me.Customer_Reports_Workshop.Value = True Then
        ReportType = "CUSTOMER"
        FilterValue = "WORKSHOP"
    ElseIf Me.Due_Date_Reports.Value = True Then
        ReportType = "DUEDATE"
        FilterValue = Format(DateAdd("d", 7, Now), "dd/mm/yyyy")
    Else
        MsgBox "Please select a report type to preview.", vbInformation
        Exit Sub
    End If

    If ShowReportPreview(ReportType, FilterValue) Then
        MsgBox "Report preview generated.", vbInformation
    End If
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "PreviewReport_Click", "fwip"
End Sub

Private Function ShowReportPreview(ByVal ReportType As String, Optional ByVal FilterValue As String = "") As Boolean
    Dim WIPJobs As Variant
    Dim PreviewText As String
    Dim i As Integer
    Dim Count As Integer

    On Error GoTo Error_Handler

    WIPJobs = WIPManager.GetWIPJobs()
    PreviewText = "Report Preview - " & ReportType & vbCrLf & String(50, "=") & vbCrLf & vbCrLf
    Count = 0

    For i = 0 To UBound(WIPJobs)
        If Count < 10 Then
            PreviewText = PreviewText & WIPJobs(i).JobNumber & " - " & WIPJobs(i).CustomerName & " - " & WIPJobs(i).ComponentDescription & vbCrLf
            Count = Count + 1
        End If
    Next i

    If Count = 10 Then
        PreviewText = PreviewText & vbCrLf & "... and " & (UBound(WIPJobs) - 9) & " more records"
    End If

    MsgBox PreviewText, vbInformation, "Report Preview"
    ShowReportPreview = True
    Exit Function

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "ShowReportPreview", "fwip"
    ShowReportPreview = False
End Function