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
    CoreFramework.HandleStandardErrors Err.Number, "Go_Click", "fwip"
End Sub

Private Sub Cancel_Click()
    Unload Me
End Sub

Private Function GenerateSelectedReports() As Boolean
    Dim ReportChoice As Integer
    Dim ReportsGenerated As Boolean

    On Error GoTo Error_Handler

    ReportsGenerated = False

    ' Show report selection menu
    ReportChoice = ShowReportMenu()

    Select Case ReportChoice
        Case 1 ' Operation Reports
            If BusinessController.GenerateWIPReport("BYOPERATOR") Then
                ReportsGenerated = True
            End If

        Case 2 ' Operator Reports
            If BusinessController.GenerateWIPReport("BYOPERATOR") Then
                ReportsGenerated = True
            End If

        Case 3 ' Customer Reports - Office
            If BusinessController.GenerateWIPReport("ALL") Then
                ReportsGenerated = True
            End If

        Case 4 ' Customer Reports - Workshop
            If BusinessController.GenerateWIPReport("ALL") Then
                ReportsGenerated = True
            End If

        Case 5 ' Due Date Reports
            Dim DueDateFilter As String
            DueDateFilter = Format(DateAdd("d", 7, Now), "dd/mm/yyyy")
            If BusinessController.GenerateWIPReport("BYDUEDATE", DueDateFilter) Then
                ReportsGenerated = True
            End If

        Case 6 ' Job Number Reports - Office
            If BusinessController.GenerateWIPReport("ALL") Then
                ReportsGenerated = True
            End If

        Case 7 ' Job Number Reports - Workshop
            If BusinessController.GenerateWIPReport("ALL") Then
                ReportsGenerated = True
            End If

        Case 8 ' All Reports
            If BusinessController.GenerateWIPReport("ALL") Then
                ReportsGenerated = True
            End If

        Case 0 ' Cancel
            ReportsGenerated = False
    End Select

    GenerateSelectedReports = ReportsGenerated
    Exit Function

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "GenerateSelectedReports", "fwip"
    GenerateSelectedReports = False
End Function

Private Function ShowReportMenu() As Integer
    Dim MenuItems As String
    Dim UserChoice As String

    On Error GoTo Error_Handler

    MenuItems = "Select WIP Report Type:" & vbCrLf & vbCrLf
    MenuItems = MenuItems & "1 - Operation Reports" & vbCrLf
    MenuItems = MenuItems & "2 - Operator Reports" & vbCrLf
    MenuItems = MenuItems & "3 - Customer Reports - Office" & vbCrLf
    MenuItems = MenuItems & "4 - Customer Reports - Workshop" & vbCrLf
    MenuItems = MenuItems & "5 - Due Date Reports" & vbCrLf
    MenuItems = MenuItems & "6 - Job Number Reports - Office" & vbCrLf
    MenuItems = MenuItems & "7 - Job Number Reports - Workshop" & vbCrLf
    MenuItems = MenuItems & "8 - All Reports" & vbCrLf & vbCrLf
    MenuItems = MenuItems & "Enter choice (1-8, 0 to cancel):"

    UserChoice = InputBox(MenuItems, "WIP Report Generator", "8")

    If IsNumeric(UserChoice) Then
        ShowReportMenu = CInt(UserChoice)
    Else
        ShowReportMenu = 0 ' Cancel
    End If

    Exit Function

Error_Handler:
    ShowReportMenu = 0
End Function

Private Function GenerateCustomReport() As Boolean
    Dim CustomerFilter As String
    Dim OperatorFilter As String
    Dim ReportType As String

    On Error GoTo Error_Handler

    ' Prompt user for custom filters
    CustomerFilter = Trim(InputBox("Enter customer name filter (leave blank for none):", "Custom Customer Filter"))

    If CustomerFilter <> "" Then
        ReportType = "CUSTOMER"
        GenerateCustomReport = BusinessController.GenerateWIPReport(ReportType, CustomerFilter)
    Else
        OperatorFilter = Trim(InputBox("Enter operator name filter (leave blank for all):", "Custom Operator Filter"))

        If OperatorFilter <> "" Then
            ReportType = "OPERATOR"
            GenerateCustomReport = BusinessController.GenerateWIPReport(ReportType, OperatorFilter)
        Else
            GenerateCustomReport = BusinessController.GenerateWIPReport("ALL")
        End If
    End If
    Exit Function

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "GenerateCustomReport", "fwip"
    GenerateCustomReport = False
End Function

Private Sub SelectAll_Click()
    On Error GoTo Error_Handler

    ' Generate all available reports
    Dim ReportsGenerated As Boolean
    ReportsGenerated = False

    If BusinessController.GenerateWIPReport("OPERATION") Then ReportsGenerated = True
    If BusinessController.GenerateWIPReport("OPERATOR") Then ReportsGenerated = True
    If BusinessController.GenerateWIPReport("CUSTOMER", "OFFICE") Then ReportsGenerated = True
    If BusinessController.GenerateWIPReport("CUSTOMER", "WORKSHOP") Then ReportsGenerated = True

    Dim DueDateFilter As String
    DueDateFilter = Format(DateAdd("d", 7, Now), "dd/mm/yyyy")
    If BusinessController.GenerateWIPReport("DUEDATE", DueDateFilter) Then ReportsGenerated = True

    If BusinessController.GenerateWIPReport("JOBNUMBER", "OFFICE") Then ReportsGenerated = True
    If BusinessController.GenerateWIPReport("JOBNUMBER", "WORKSHOP") Then ReportsGenerated = True

    If ReportsGenerated Then
        MsgBox "All reports generated successfully.", vbInformation
    Else
        MsgBox "No reports could be generated.", vbCritical
    End If
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "SelectAll_Click", "fwip"
End Sub

Private Sub ClearAll_Click()
    On Error GoTo Error_Handler

    ' Close the form - no controls to clear
    Unload Me
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "ClearAll_Click", "fwip"
End Sub

Private Sub ViewWIPDatabase_Click()
    On Error GoTo Error_Handler

    Dim WIPPath As String
    WIPPath = DataManager.GetRootPath & "\WIP.xls"

    Dim wb As Workbook
    Set wb = DataManager.SafeOpenWorkbook(WIPPath)
    If wb Is Nothing Then
        MsgBox "Could not open WIP database.", vbCritical
    Else
        Me.Hide
    End If
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "ViewWIPDatabase_Click", "fwip"
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
    CoreFramework.HandleStandardErrors Err.Number, "RefreshWIPData_Click", "fwip"
End Sub

Private Function RefreshWIPFromFiles() As Boolean
    Dim WIPFiles As Variant
    Dim i As Integer
    Dim JobInfo As JobData
    Dim RefreshCount As Integer

    On Error GoTo Error_Handler

    WIPFiles = DataManager.GetFileList("WIP")
    RefreshCount = 0

    For i = 0 To UBound(WIPFiles)
        Dim JobPath As String
        JobPath = DataManager.GetRootPath & "\WIP\" & WIPFiles(i)

        JobInfo = BusinessController.LoadJob(JobPath)
        If JobInfo.JobNumber <> "" Then
            If BusinessController.UpdateJobInWIP(JobInfo) Then
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
    CoreFramework.HandleStandardErrors Err.Number, "RefreshWIPFromFiles", "fwip"
    RefreshWIPFromFiles = False
End Function

Private Sub LoadCustomerList()
    Dim WIPJobs As Variant
    Dim i As Integer
    Dim UniqueCustomers As Collection
    Dim Customer As Variant
    Dim CustomerName As String

    On Error GoTo Error_Handler

    Set UniqueCustomers = New Collection
    WIPJobs = BusinessController.GetWIPJobs()

    If IsArray(WIPJobs) Then
        For i = 0 To UBound(WIPJobs, 1)
            CustomerName = WIPJobs(i, 1) ' CustomerName is in column 1
            If CustomerName <> "" Then
                On Error Resume Next
                UniqueCustomers.Add CustomerName, CustomerName
                On Error GoTo Error_Handler
            End If
        Next i
    End If

    Me.Customer_Filter.Clear
    For Each Customer In UniqueCustomers
        Me.Customer_Filter.AddItem Customer
    Next Customer
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "LoadCustomerList", "fwip"
End Sub

Private Sub LoadOperatorList()
    Dim WIPJobs As Variant
    Dim i As Integer
    Dim UniqueOperators As Collection
    Dim Operator As Variant
    Dim OperatorName As String

    On Error GoTo Error_Handler

    Set UniqueOperators = New Collection
    WIPJobs = BusinessController.GetWIPJobs()

    If IsArray(WIPJobs) Then
        For i = 0 To UBound(WIPJobs, 1)
            OperatorName = WIPJobs(i, 5) ' AssignedOperator is in column 5
            If OperatorName <> "" Then
                On Error Resume Next
                UniqueOperators.Add OperatorName, OperatorName
                On Error GoTo Error_Handler
            End If
        Next i
    End If

    Me.Operator_Filter.Clear
    For Each Operator In UniqueOperators
        Me.Operator_Filter.AddItem Operator
    Next Operator
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "LoadOperatorList", "fwip"
End Sub

Private Sub UpdateJobCounts()
    Dim WIPJobs As Variant
    Dim ActiveJobs As Integer
    Dim OnHoldJobs As Integer
    Dim OverdueJobs As Integer
    Dim i As Integer

    On Error GoTo Error_Handler

    WIPJobs = BusinessController.GetWIPJobs()
    ActiveJobs = 0
    OnHoldJobs = 0
    OverdueJobs = 0

    If IsArray(WIPJobs) Then
        For i = 0 To UBound(WIPJobs, 1)
            Select Case UCase(WIPJobs(i, 6)) ' Status is in column 6
                Case "ACTIVE"
                    ActiveJobs = ActiveJobs + 1
                    If IsDate(WIPJobs(i, 4)) And DateValue(WIPJobs(i, 4)) < Date Then ' DueDate is in column 4
                        OverdueJobs = OverdueJobs + 1
                    End If
                Case "ON HOLD", "ONHOLD"
                    OnHoldJobs = OnHoldJobs + 1
            End Select
        Next i
    End If

    Me.Active_Jobs_Count.Caption = "Active Jobs: " & ActiveJobs
    Me.OnHold_Jobs_Count.Caption = "On Hold: " & OnHoldJobs
    Me.Overdue_Jobs_Count.Caption = "Overdue: " & OverdueJobs
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "UpdateJobCounts", "fwip"
End Sub

Private Sub UserForm_Initialize()
    On Error GoTo Error_Handler

    LoadCustomerList
    LoadOperatorList
    UpdateJobCounts
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "UserForm_Initialize", "fwip"
End Sub

Private Sub PreviewReport_Click()
    Dim ReportType As String
    Dim FilterValue As String

    On Error GoTo Error_Handler

    ' Show report selection menu for preview
    Dim ReportChoice As Integer
    ReportChoice = ShowReportMenu()

    Select Case ReportChoice
        Case 1
            ReportType = "OPERATION"
        Case 2
            ReportType = "OPERATOR"
        Case 3
            ReportType = "CUSTOMER"
            FilterValue = "OFFICE"
        Case 4
            ReportType = "CUSTOMER"
            FilterValue = "WORKSHOP"
        Case 5
            ReportType = "DUEDATE"
            FilterValue = Format(DateAdd("d", 7, Now), "dd/mm/yyyy")
        Case 6
            ReportType = "JOBNUMBER"
            FilterValue = "OFFICE"
        Case 7
            ReportType = "JOBNUMBER"
            FilterValue = "WORKSHOP"
        Case 8
            ReportType = "ALL"
        Case 0
            Exit Sub
        Case Else
            MsgBox "Invalid selection.", vbInformation
            Exit Sub
    End Select

    If ShowReportPreview(ReportType, FilterValue) Then
        MsgBox "Report preview generated.", vbInformation
    End If
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "PreviewReport_Click", "fwip"
End Sub

Private Function ShowReportPreview(ByVal ReportType As String, Optional ByVal FilterValue As String = "") As Boolean
    Dim WIPJobs As Variant
    Dim PreviewText As String
    Dim i As Integer
    Dim Count As Integer

    On Error GoTo Error_Handler

    WIPJobs = BusinessController.GetWIPJobs()
    PreviewText = "Report Preview - " & ReportType & vbCrLf & String(50, "=") & vbCrLf & vbCrLf
    Count = 0

    If IsArray(WIPJobs) Then
        For i = 0 To UBound(WIPJobs, 1)
            If Count < 10 Then
                PreviewText = PreviewText & WIPJobs(i, 0) & " - " & WIPJobs(i, 1) & " - " & WIPJobs(i, 2) & vbCrLf ' JobNumber, CustomerName, ComponentDescription
                Count = Count + 1
            End If
        Next i

        If Count = 10 Then
            PreviewText = PreviewText & vbCrLf & "... and " & (UBound(WIPJobs, 1) - 9) & " more records"
        End If
    End If

    MsgBox PreviewText, vbInformation, "Report Preview"
    ShowReportPreview = True
    Exit Function

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "ShowReportPreview", "fwip"
    ShowReportPreview = False
End Function