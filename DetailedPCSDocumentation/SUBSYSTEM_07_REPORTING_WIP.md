# Subsystem 7: Reporting & WIP Management - PCS Original System

## 🎯 **Subsystem Purpose**

The Reporting & WIP Management subsystem provides **comprehensive work-in-progress reporting and job analysis** capabilities. This subsystem generates various reports for production management, operator tracking, and job scheduling based on the WIP database.

**Responsibility**: WIP report generation, job analysis, operator reports, due date tracking, and production scheduling support.

---

## 📁 **Module and Form Inventory**

### **Core Components**

| Component | Type | Lines | Purpose | Dependencies |
|-----------|------|-------|---------|-------------|
| `fwip.frm` | UserForm | 289+ | WIP reports interface | WIP.xls, Excel reporting |
| `fwip_modified.frm` | UserForm | 250+ | Enhanced WIP form | Same as fwip.frm |

**Note**: Two WIP forms exist - original and modified versions with different layouts and features.

---

## 📋 **WIP Reporting Interface**

### **fwip.frm** - Work-in-Progress Reports

#### **Report Type Options**

##### **Primary Report Categories**
```vba
' Report type radio buttons
ROperation      ' Group jobs by operation type
ROperator       ' Group jobs by assigned operator
RDueDate        ' Sort jobs by due date
RWIP            ' Complete WIP listing
```

##### **Sorting Options**
```vba
' Sorting criteria
Job_DueDate           ' Sort by customer delivery date
Office_Customer       ' Office view - sort by customer
Workshop_Customer     ' Workshop view - sort by customer
Office_JobNumber      ' Office view - sort by job number
Workshop_JobNumber    ' Workshop view - sort by job number
```

#### **Primary Event Handler**

##### **Go_Click() - Generate Report**
```vba
Private Sub Go_Click()
    ' 1. Validate report parameters selected
    If Not ReportTypeSelected() Then
        MsgBox "Please select a report type"
        Exit Sub
    End If
    
    ' 2. Open WIP database
    Call Open_Book.OpenBook(Main.Main_MasterPath.Value & "WIP.xls", True)
    
    ' 3. Generate report based on selected type
    If ROperation.Value Then
        Call GenerateOperationReport()
    ElseIf ROperator.Value Then
        Call GenerateOperatorReport()
    ElseIf RDueDate.Value Then
        Call GenerateDueDateReport()
    ElseIf RWIP.Value Then
        Call GenerateCompleteWIPReport()
    End If
    
    ' 4. Format and display report
    Call FormatReportOutput()
    
    ' 5. Save report to Templates directory
    Call SaveReportToFile()
End Sub
```

---

## 📁 **Jobs Data Structure**

### **Jobs Type Definition**

#### **Core Job Data Structure**
```vba
Private Type Jobs
    Dat As Date                    ' Job date
    Cust As String                 ' Customer name
    Job As String                  ' Job number
    Qty As String                  ' Quantity
    Desc As String                 ' Description
    Remarks As String              ' Additional notes
    DDat As String                 ' Due date
    OPs(1 To 15) As String        ' Operations (fwip.frm version)
    OperatorN(1 To 15) As String  ' Operator names (fwip_modified.frm)
    OperatorType(1 To 15) As String ' Operation types
End Type
```

#### **Array Declaration for Job Processing**
```vba
' Job data arrays for report processing
Dim JobsArray(1 To 500) As Jobs    ' Support up to 500 concurrent jobs
Dim JobCount As Integer             ' Actual number of jobs loaded
```

---

## 📊 **Report Generation Types**

### **Operation Reports**

#### **GenerateOperationReport() - Group by Operation Type**
```vba
Private Sub GenerateOperationReport()
    ' 1. Load all jobs from WIP database
    Call LoadJobsFromWIP()
    
    ' 2. Group jobs by operation type
    Dim operationGroups As Collection
    Set operationGroups = New Collection
    
    Dim i As Integer, j As Integer
    For i = 1 To JobCount
        For j = 1 To 15
            If JobsArray(i).OperatorType(j) <> "" Then
                ' Add job to operation group
                Call AddJobToOperationGroup(operationGroups, JobsArray(i), j)
            End If
        Next j
    Next i
    
    ' 3. Create report sheets for each operation
    Call CreateOperationReportSheets(operationGroups)
End Sub
```

#### **Operation Groupings**
```vba
' Common manufacturing operations
"Machining"     ' CNC, milling, turning operations
"Welding"       ' Arc, MIG, TIG welding operations
"Assembly"      ' Component assembly tasks
"Inspection"    ' Quality control checkpoints
"Grinding"      ' Surface finishing operations
"Drilling"      ' Hole creation operations
"Heat Treatment" ' Thermal processing
"Painting"      ' Coating and finishing
```

### **Operator Reports**

#### **GenerateOperatorReport() - Group by Assigned Operator**
```vba
Private Sub GenerateOperatorReport()
    ' 1. Load jobs and extract operator assignments
    Call LoadJobsFromWIP()
    
    ' 2. Create operator workload summary
    Dim operatorWorkload As Collection
    Set operatorWorkload = New Collection
    
    ' 3. Calculate workload per operator
    Dim i As Integer, j As Integer
    For i = 1 To JobCount
        For j = 1 To 15
            If JobsArray(i).OperatorN(j) <> "" Then
                Call AddJobToOperatorWorkload(operatorWorkload, JobsArray(i), j)
            End If
        Next j
    Next i
    
    ' 4. Generate operator-specific worksheets
    Call CreateOperatorReportSheets(operatorWorkload)
End Sub
```

#### **Operator Workload Analysis**
```vba
' Operator metrics calculated:
- Total jobs assigned
- Estimated completion time
- Current operation status
- Overdue jobs
- Efficiency ratings
```

### **Due Date Reports**

#### **GenerateDueDateReport() - Sort by Customer Delivery Dates**
```vba
Private Sub GenerateDueDateReport()
    ' 1. Load all WIP jobs
    Call LoadJobsFromWIP()
    
    ' 2. Sort jobs by due date (earliest first)
    Call SortJobsByDueDate(JobsArray, JobCount)
    
    ' 3. Categorize by urgency
    Dim overdueJobs As Collection
    Dim dueThisWeek As Collection
    Dim dueNextWeek As Collection
    
    Call CategorizeJobsByUrgency(overdueJobs, dueThisWeek, dueNextWeek)
    
    ' 4. Create urgency-based report sheets
    Call CreateDueDateReportSheets(overdueJobs, dueThisWeek, dueNextWeek)
End Sub
```

#### **Due Date Categories**
```vba
' Urgency classifications:
"OVERDUE"       ' Past customer delivery date
"DUE TODAY"     ' Due today
"DUE THIS WEEK" ' Due within 7 days
"DUE NEXT WEEK" ' Due within 14 days
"ON SCHEDULE"   ' Due beyond 14 days
```

### **Complete WIP Report**

#### **GenerateCompleteWIPReport() - Full WIP Listing**
```vba
Private Sub GenerateCompleteWIPReport()
    ' 1. Load all WIP jobs
    Call LoadJobsFromWIP()
    
    ' 2. Apply selected sorting criteria
    If Job_DueDate.Value Then
        Call SortJobsByDueDate(JobsArray, JobCount)
    ElseIf Office_Customer.Value Or Workshop_Customer.Value Then
        Call SortJobsByCustomer(JobsArray, JobCount)
    ElseIf Office_JobNumber.Value Or Workshop_JobNumber.Value Then
        Call SortJobsByJobNumber(JobsArray, JobCount)
    End If
    
    ' 3. Create comprehensive listing
    Call CreateCompleteWIPSheet()
    
    ' 4. Add summary statistics
    Call AddWIPSummaryStatistics()
End Sub
```

---

## 📈 **Report Formatting and Output**

### **Report Layout Standards**

#### **Standard Report Headers**
```vba
' Column headers for WIP reports
"Job Number"           ' J-prefix job identifier
"Customer"             ' Customer name
"Description"          ' Component description
"Quantity"             ' Required quantity
"Start Date"           ' Production start date
"Due Date"             ' Customer delivery date
"Current Operation"    ' Active operation
"Assigned Operator"    ' Operator responsible
"Status"               ' Job status
"Days Remaining"       ' Time to due date
```

#### **FormatReportOutput() - Apply Professional Formatting**
```vba
Private Sub FormatReportOutput()
    ' 1. Apply standard column widths
    With ActiveSheet
        .Columns("A:A").ColumnWidth = 12  ' Job Number
        .Columns("B:B").ColumnWidth = 20  ' Customer
        .Columns("C:C").ColumnWidth = 30  ' Description
        .Columns("D:D").ColumnWidth = 10  ' Quantity
        .Columns("E:E").ColumnWidth = 12  ' Start Date
        .Columns("F:F").ColumnWidth = 12  ' Due Date
    End With
    
    ' 2. Apply header formatting
    With Range("A1:J1")
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
        .Borders.Weight = xlMedium
    End With
    
    ' 3. Apply conditional formatting for due dates
    Call ApplyDueDateConditionalFormatting()
    
    ' 4. Add borders and gridlines
    Call ApplyReportBorders()
End Sub
```

#### **Conditional Formatting for Urgency**
```vba
Private Sub ApplyDueDateConditionalFormatting()
    ' Red background for overdue jobs
    ' Yellow background for jobs due this week
    ' Green background for jobs on schedule
    
    Dim dueDateColumn As Range
    Set dueDateColumn = Range("F:F")  ' Due date column
    
    ' Overdue formatting (red)
    With dueDateColumn.FormatConditions.Add(xlCellValue, xlLess, Date)
        .Interior.Color = RGB(255, 200, 200)
    End With
    
    ' Due this week formatting (yellow)
    With dueDateColumn.FormatConditions.Add(xlCellValue, xlBetween, Date, Date + 7)
        .Interior.Color = RGB(255, 255, 200)
    End With
End Sub
```

### **Report Export and Distribution**

#### **SaveReportToFile() - Save to Templates Directory**
```vba
Private Sub SaveReportToFile()
    Dim reportFileName As String
    Dim timestamp As String
    
    timestamp = Format(Now, "yyyy-mm-dd_hh-mm")
    
    ' Generate filename based on report type
    If ROperation.Value Then
        reportFileName = "WIP_Operations_" & timestamp & ".xls"
    ElseIf ROperator.Value Then
        reportFileName = "WIP_Operators_" & timestamp & ".xls"
    ElseIf RDueDate.Value Then
        reportFileName = "WIP_DueDates_" & timestamp & ".xls"
    Else
        reportFileName = "WIP_Complete_" & timestamp & ".xls"
    End If
    
    ' Save to Templates directory
    Dim savePath As String
    savePath = Main.Main_MasterPath.Value & "Templates\" & reportFileName
    
    ActiveWorkbook.SaveAs savePath
    
    MsgBox "Report saved as: " & reportFileName
End Sub
```

---

## 🔄 **Data Loading and Processing**

### **WIP Database Integration**

#### **LoadJobsFromWIP() - Extract Job Data**
```vba
Private Sub LoadJobsFromWIP()
    ' 1. Open WIP.xls database
    Dim wipWorkbook As Workbook
    Set wipWorkbook = Workbooks.Open(Main.Main_MasterPath.Value & "WIP.xls")
    
    Dim ws As Worksheet
    Set ws = wipWorkbook.Worksheets("WIP")
    
    ' 2. Determine number of jobs
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' 3. Load job data into array
    JobCount = 0
    Dim i As Long
    For i = 2 To lastRow  ' Skip header row
        If ws.Cells(i, 1).Value <> "" Then  ' Job number not empty
            JobCount = JobCount + 1
            
            ' Load core job data
            With JobsArray(JobCount)
                .Job = ws.Cells(i, 1).Value      ' Job number
                .Cust = ws.Cells(i, 2).Value     ' Customer
                .Desc = ws.Cells(i, 3).Value     ' Description
                .Qty = ws.Cells(i, 4).Value      ' Quantity
                .Dat = ws.Cells(i, 5).Value      ' Start date
                .DDat = ws.Cells(i, 6).Value     ' Due date
                
                ' Load operation data (columns 7-21 for 15 operations)
                Dim j As Integer
                For j = 1 To 15
                    .OPs(j) = ws.Cells(i, 6 + j).Value
                Next j
            End With
        End If
    Next i
    
    ' 4. Close WIP database
    wipWorkbook.Close SaveChanges:=False
End Sub
```

### **Sorting and Filtering Functions**

#### **SortJobsByDueDate() - Date-Based Sorting**
```vba
Private Sub SortJobsByDueDate(ByRef jobs() As Jobs, jobCount As Integer)
    ' Bubble sort implementation for job due dates
    Dim i As Integer, j As Integer
    Dim tempJob As Jobs
    
    For i = 1 To jobCount - 1
        For j = 1 To jobCount - i
            If CDate(jobs(j).DDat) > CDate(jobs(j + 1).DDat) Then
                ' Swap jobs
                tempJob = jobs(j)
                jobs(j) = jobs(j + 1)
                jobs(j + 1) = tempJob
            End If
        Next j
    Next i
End Sub
```

#### **FilterJobsByStatus() - Status-Based Filtering**
```vba
Private Function FilterJobsByStatus(status As String) As Collection
    Dim filteredJobs As New Collection
    Dim i As Integer
    
    For i = 1 To JobCount
        ' Check if job matches status criteria
        If JobMatchesStatus(JobsArray(i), status) Then
            filteredJobs.Add JobsArray(i)
        End If
    Next i
    
    Set FilterJobsByStatus = filteredJobs
End Function
```

---

## 📊 **Advanced Reporting Features**

### **Multi-Sheet Report Generation**

#### **CreateOperationReportSheets() - Operation-Specific Worksheets**
```vba
Private Sub CreateOperationReportSheets(operationGroups As Collection)
    ' Create summary sheet
    Dim summarySheet As Worksheet
    Set summarySheet = ActiveWorkbook.Worksheets.Add
    summarySheet.Name = "Operation Summary"
    
    ' Create individual sheets for each operation type
    Dim operation As Variant
    For Each operation In operationGroups
        Dim operationSheet As Worksheet
        Set operationSheet = ActiveWorkbook.Worksheets.Add
        operationSheet.Name = operation.Name
        
        ' Populate sheet with jobs for this operation
        Call PopulateOperationSheet(operationSheet, operation.Jobs)
    Next operation
End Sub
```

### **Statistical Analysis**

#### **AddWIPSummaryStatistics() - Report Statistics**
```vba
Private Sub AddWIPSummaryStatistics()
    ' Add summary statistics sheet
    Dim statsSheet As Worksheet
    Set statsSheet = ActiveWorkbook.Worksheets.Add
    statsSheet.Name = "Statistics"
    
    ' Calculate statistics
    Dim totalJobs As Integer
    Dim overdueJobs As Integer
    Dim averageLeadTime As Double
    
    totalJobs = JobCount
    overdueJobs = CountOverdueJobs()
    averageLeadTime = CalculateAverageLeadTime()
    
    ' Write statistics to sheet
    With statsSheet
        .Cells(1, 1).Value = "WIP Statistics"
        .Cells(2, 1).Value = "Total Jobs:"
        .Cells(2, 2).Value = totalJobs
        .Cells(3, 1).Value = "Overdue Jobs:"
        .Cells(3, 2).Value = overdueJobs
        .Cells(4, 1).Value = "Average Lead Time:"
        .Cells(4, 2).Value = averageLeadTime & " days"
    End With
End Sub
```

---

## ⚠️ **Error Handling and Validation**

### **Report Generation Error Handling**

#### **Safe Report Generation**
```vba
Private Function SafeGenerateReport() As Boolean
    On Error GoTo ErrorHandler
    
    ' 1. Validate WIP database exists
    If Dir(Main.Main_MasterPath.Value & "WIP.xls") = "" Then
        MsgBox "WIP database not found"
        SafeGenerateReport = False
        Exit Function
    End If
    
    ' 2. Check for jobs in WIP
    If CountWIPJobs() = 0 Then
        MsgBox "No jobs in WIP - no report to generate"
        SafeGenerateReport = False
        Exit Function
    End If
    
    ' 3. Generate report
    Call GenerateSelectedReport()
    
    SafeGenerateReport = True
    Exit Function
    
ErrorHandler:
    MsgBox "Error generating report: " & Err.Description
    SafeGenerateReport = False
End Function
```

### **Data Validation**

#### **ValidateJobData() - Data Quality Checks**
```vba
Private Function ValidateJobData(job As Jobs) As Boolean
    ValidateJobData = True
    
    ' Check required fields
    If job.Job = "" Then
        ValidateJobData = False
        Exit Function
    End If
    
    If job.Cust = "" Then
        ValidateJobData = False
        Exit Function
    End If
    
    ' Validate dates
    If Not IsDate(job.Dat) Or Not IsDate(job.DDat) Then
        ValidateJobData = False
        Exit Function
    End If
    
    ' Check for reasonable due dates
    If CDate(job.DDat) < CDate(job.Dat) Then
        ValidateJobData = False  ' Due date before start date
        Exit Function
    End If
End Function
```

---

## 🔧 **Development Guidelines**

### **Customizing WIP Reports**

#### **Adding New Report Types**
```vba
' 1. Add radio button to fwip.frm
' 2. Create report generation function
Private Sub GenerateCustomReport()
    ' Custom report logic
End Sub

' 3. Update Go_Click() event handler
Private Sub Go_Click()
    ' ... existing report types ...
    ElseIf RCustom.Value Then
        Call GenerateCustomReport()
    End If
End Sub
```

#### **Enhanced Data Analysis**
```vba
' Add calculated fields to reports
Private Function CalculateJobMetrics(job As Jobs) As Collection
    Dim metrics As New Collection
    
    ' Days in WIP
    metrics.Add DateDiff("d", job.Dat, Date), "DaysInWIP"
    
    ' Days until due
    metrics.Add DateDiff("d", Date, job.DDat), "DaysUntilDue"
    
    ' Operations completed
    metrics.Add CountCompletedOperations(job), "OperationsComplete"
    
    Set CalculateJobMetrics = metrics
End Function
```

### **Performance Optimization**

#### **Large Dataset Handling**
```vba
' Optimize for large WIP databases
Private Sub OptimizedDataLoading()
    ' Use arrays instead of collections for large datasets
    ' Implement paging for very large reports
    ' Cache frequently accessed data
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    ' ... data processing ...
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
```

---

## 🔍 **Next Steps**

After understanding Reporting & WIP Management:

1. **Study [Search Database](SUBSYSTEM_08_SEARCH_DATA.md)** - Complete the system understanding
2. **Practice Report Generation** - Create test WIP data and generate various reports
3. **Customize Report Formats** - Add new columns or modify layouts
4. **Integrate with Main Interface** - Understand how reports are launched from Main.frm
5. **Optimize Performance** - Test with large datasets and improve speed

**Ready for the final subsystem? Continue to [Search Database Subsystem](SUBSYSTEM_08_SEARCH_DATA.md)**