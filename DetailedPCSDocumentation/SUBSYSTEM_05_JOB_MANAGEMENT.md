# Subsystem 5: Job Management - PCS Original System

## 🎯 **Subsystem Purpose**

The Job Management subsystem handles the **complete job lifecycle** from quote acceptance through production completion. This subsystem manages job creation, work-in-progress tracking, job card operations, and job completion workflow.

**Responsibility**: Job acceptance from quotes, WIP database management, production tracking, job card operations, and job completion processing.

---

## 📁 **Module and Form Inventory**

### **Core Components**

| Component | Type | Lines | Purpose | Dependencies |
|-----------|------|-------|---------|-------------|
| `FAcceptQuote.frm` | UserForm | 200+ | Quote acceptance and job creation | Calc_Numbers, SaveWIPCode, SaveSearchCode |
| `FJG.frm` | UserForm | 250+ | Job generation with operations planning | Calc_Numbers, SaveWIPCode, Templates |
| `FJobCard.frm` | UserForm | 180+ | Job card management and completion | SaveWIPCode, Archive operations |
| `SaveWIPCode.bas` | Module | 45 | WIP database operations | WIP.xls, Open_Book |

**Total**: 675+ lines managing complete job lifecycle

---

## 🔄 **Job Creation Workflow**

### **FAcceptQuote.frm** - Quote Acceptance and Job Creation

#### **Form Purpose and Context**
```vba
' Quote acceptance form is opened when:
' 1. Customer accepts a submitted quote
' 2. Quote file is in Archive/ directory (quote submitted status)
' 3. User selects "Accept Quote" from Main interface
' 4. Form creates new job from accepted quote
```

#### **Key Controls and Data Fields**

##### **Pre-Populated from Quote**
- **Quote_Number** - Source quote (read-only)
- **Customer** - Customer name from quote
- **Component_Description** - Part description
- **Component_Quantity** - Required quantity
- **Component_Price** - Agreed pricing

##### **Job-Specific Fields**
- **Job_Number** - Auto-generated J-prefix number
- **CustomerOrderNumber** - Customer's PO number (required)
- **Job_Urgency** - Priority level (Normal, Break Down, Urgent)
- **Job_LeadTime** - Calculated based on urgency
- **Job_StartDate** - Production start date
- **CustomerDelivery_Date** - Customer delivery deadline
- **Job_WorkshopDueDate** - Internal due date

##### **Multi-Part Job Support**
- **Compilation_SequenceNumber** - Part number in assembly
- **Compilation_TotalNumber** - Total parts in assembly

#### **Primary Event Handlers**

##### **UserForm_Activate() - Form Initialization**
```vba
Private Sub UserForm_Activate()
    ' 1. Load quote data from Archive/ directory
    Dim quoteFile As String
    quoteFile = Main.lst.Value ' Selected quote from main interface

    ' 2. Read quote data
    Me.Quote_Number.Value = GetValue(archivePath, quoteFile, \"Admin\", \"Quote_Number\")
    Me.Customer.Value = GetValue(archivePath, quoteFile, \"Admin\", \"Customer\")
    ' ... populate other fields from quote

    ' 3. Calculate and display next job number
    Dim nextJobNumber As Long
    nextJobNumber = Calc_Numbers.Calc_Next_Number(\"J\")
    Me.Job_Number.Value = \"J\" & Format(nextJobNumber, \"0000\")

    ' 4. Set default values
    Me.Job_Urgency.Value = \"Normal\"
    Me.Job_StartDate.Value = Date
    Me.Job_LeadTime.Value = 14 ' Default lead time
End Sub
```

##### **butSAVE_Click() - Create Job**
```vba
Private Sub butSAVE_Click()
    ' 1. Validate job data
    If Not ValidateJobForm() Then Exit Sub

    ' 2. Confirm job number generation
    Dim confirmedJobNumber As Long
    confirmedJobNumber = Calc_Numbers.Confirm_Next_Number(\"J\")

    ' 3. Create job file from quote
    Dim quotePath As String, jobPath As String
    quotePath = Main.Main_MasterPath.Value & \"Archive\\\" & Me.Quote_Number.Value & \".xls\"
    jobPath = Main.Main_MasterPath.Value & \"WIP\\\" & \"J\" & Format(confirmedJobNumber, \"0000\") & \".xls\"

    ' 4. Copy quote file to WIP directory
    FileCopy quotePath, jobPath

    ' 5. Update job file with job-specific data
    Call PopulateJobFile(jobPath, confirmedJobNumber)

    ' 6. Update WIP database
    Call SaveWIPCode.SaveInfoIntoWIP(Me)

    ' 7. Update search database
    Call SaveSearchCode.SaveRowIntoSearch(Me)

    ' 8. Activate Job Card sheet for production
    Call ActivateJobCardSheet(jobPath)

    MsgBox \"Job \" & \"J\" & Format(confirmedJobNumber, \"0000\") & \" created successfully\"
    Me.Hide
End Sub
```

##### **Job_Urgency_Change() - Auto-Calculate Lead Times**
```vba
Private Sub Job_Urgency_Change()
    Select Case Me.Job_Urgency.Value
        Case \"Normal\"
            Me.Job_LeadTime.Value = 14
        Case \"Break Down\"
            Me.Job_LeadTime.Value = 7
        Case \"Urgent\"
            Me.Job_LeadTime.Value = 10
    End Select

    ' Update delivery dates based on new lead time
    Call CalculateDeliveryDates
End Sub
```

---

## 🏭 **Advanced Job Generation**

### **FJG.frm** - Job Generator with Operations Planning

#### **Enhanced Job Creation Features**
```vba
' FJG.frm provides advanced job creation with:
' 1. Detailed operation planning (Operation01-15)
' 2. Operator assignment
' 3. Technical drawing integration
' 4. Contract template support
' 5. Multi-part job coordination
```

#### **Operation Planning Controls**

##### **Operation Definition (15 Operations Supported)**
```vba
' For each operation (01-15):
Operation01_Type        ' Type of manufacturing operation
Operation01_Operator    ' Assigned operator
Operation01_Comment     ' Operation instructions
Operation01_EstTime     ' Estimated time
Operation01_Status      ' Operation status

' Example operations:
' - Machining, Welding, Assembly, Inspection
' - Cutting, Drilling, Grinding, Polishing
' - Heat Treatment, Coating, Packaging
```

##### **Key Event Handlers**

**butSaveJG_Click() - Save Job with Operations**
```vba
Private Sub butSaveJG_Click()
    ' 1. Create standard job (similar to FAcceptQuote)
    ' 2. Add detailed operation planning
    ' 3. Populate Job Card sheet with operations
    ' 4. Assign operators and time estimates
    ' 5. Set up production workflow
End Sub
```

**JobCardTemplates_Click() - Load Operation Templates**
```vba
Private Sub JobCardTemplates_Click()
    ' 1. Open template selection dialog
    ' 2. Load predefined operation sequences
    ' 3. Populate operation fields automatically
    ' 4. Allow customization of template
End Sub
```

**CopyFromJobCard_Click() - Copy Existing Job Operations**
```vba
Private Sub CopyFromJobCard_Click()
    ' 1. Browse existing completed jobs
    ' 2. Select similar job for copying
    ' 3. Copy operation definitions
    ' 4. Adapt for current job requirements
End Sub
```

#### **Technical Drawing Integration**
```vba
' Job_PicturePath field links to technical drawings
Job_PicturePath = Main.Main_MasterPath.Value & \"Images\\\" & drawingFile

' Supported formats:
' - PDF technical drawings
' - CAD files
' - Image files (JPG, PNG)
' - Specification documents
```

---

## 📊 **WIP Database Management**

### **SaveWIPCode.bas** - Work-in-Progress Tracking

#### **Primary Function**

##### **`SaveInfoIntoWIP(frm As Object)` - Update WIP Database**
```vba
Sub SaveInfoIntoWIP(frm As Object)
    ' 1. Open WIP.xls database
    Call Open_Book.OpenBook(Main.Main_MasterPath.Value & \"WIP.xls\", False)

    ' 2. Find or create row for this job
    Dim wipRow As Long
    wipRow = FindJobInWIP(frm.Job_Number.Value)

    If wipRow = 0 Then
        ' Create new row
        wipRow = GetNextWIPRow()
    End If

    ' 3. Update WIP database fields
    With ActiveWorkbook.Worksheets(\"WIP\")
        .Cells(wipRow, 1).Value = frm.Job_Number.Value
        .Cells(wipRow, 2).Value = frm.Customer.Value
        .Cells(wipRow, 3).Value = frm.Component_Description.Value
        .Cells(wipRow, 4).Value = frm.Component_Quantity.Value
        .Cells(wipRow, 5).Value = frm.Job_StartDate.Value
        .Cells(wipRow, 6).Value = frm.CustomerDelivery_Date.Value
        .Cells(wipRow, 7).Value = frm.Job_Urgency.Value
        ' ... additional fields
    End With

    ' 4. Save and close WIP database
    ActiveWorkbook.Save
    ActiveWorkbook.Close
End Sub
```

#### **WIP Database Schema**

**WIP.xls Structure**:
| Column | Field | Purpose | Data Type |
|--------|-------|---------|-----------|
| A | Job_Number | Primary key | String |
| B | Customer | Customer name | String |
| C | Component_Description | Part description | String |
| D | Component_Quantity | Required quantity | Number |
| E | Job_StartDate | Production start | Date |
| F | CustomerDelivery_Date | Delivery deadline | Date |
| G | Job_Urgency | Priority level | String |
| H | Job_WorkshopDueDate | Internal due date | Date |
| I | CurrentOperation | Active operation | String |
| J | PercentComplete | Progress percentage | Number |

#### **WIP Status Tracking**
```vba
' Job status values in WIP database
\"New Job\"          ' Just created, not started
\"In Progress\"      ' Production underway
\"On Hold\"          ' Temporarily stopped
\"Quality Check\"    ' In inspection
\"Ready to Ship\"    ' Completed, awaiting delivery
\"Shipped\"          ' Delivered to customer
```

---

## 🛠️ **Job Card Operations**

### **FJobCard.frm** - Production Job Card Management

#### **Job Card Purpose**
```vba
' Job cards provide:
' 1. Production instructions for operators
' 2. Operation sequence and timing
' 3. Quality checkpoints
' 4. Progress tracking
' 5. Completion documentation
```

#### **Key Controls and Functions**

##### **Operation Tracking (15 Operations)**
```vba
' Each operation has associated controls:
Operation01_Status      ' Not Started, In Progress, Complete
Operation01_ActualTime  ' Actual time taken
Operation01_Operator    ' Operator who performed work
Operation01_QualityOK   ' Quality check passed
Operation01_Notes       ' Operation-specific notes
```

##### **Primary Event Handlers**

**SaveJobCard_Click() - Complete Job**
```vba
Private Sub SaveJobCard_Click()
    ' 1. Validate all operations complete
    If Not AllOperationsComplete() Then
        MsgBox \"Cannot complete job - operations pending\"
        Exit Sub
    End If

    ' 2. Update job status to completed
    ' 3. Calculate actual vs estimated times
    ' 4. Record completion date
    ' 5. Move job from WIP to Archive
    ' 6. Remove from WIP database
    ' 7. Update search database with completion

    Call MoveJobToArchive()
    Call RemoveFromWIP()
    Call UpdateSearchWithCompletion()
End Sub
```

**CopyFromJobCard_Click() - Reuse Job Setup**
```vba
Private Sub CopyFromJobCard_Click()
    ' 1. Browse archived jobs for similar work
    ' 2. Load operation definitions from selected job
    ' 3. Copy time estimates and procedures
    ' 4. Adapt for current job requirements
End Sub
```

#### **Quality Control Integration**
```vba
' Quality checkpoints throughout production
Private Sub PerformQualityCheck(operationNumber As Integer)
    ' 1. Display quality checklist for operation
    ' 2. Record inspection results
    ' 3. Handle non-conformance if required
    ' 4. Update operation status based on quality

    If QualityCheckPassed Then
        Operation_Status(operationNumber) = \"Complete\"
    Else
        Operation_Status(operationNumber) = \"Rework Required\"
    End If
End Sub
```

---

## 🔄 **Complete Job Lifecycle**

### **Job State Transitions**

#### **Job Status Progression**
```
1. Quote Accepted → Job Creation (FAcceptQuote.frm)
   ↓
2. Job Planning → Operation Setup (FJG.frm optional)
   ↓
3. Production Start → WIP Tracking (WIP.xls)
   ↓
4. Operation Execution → Job Card Updates (FJobCard.frm)
   ↓
5. Quality Control → Final Inspection
   ↓
6. Job Completion → Archive and Remove from WIP
   ↓
7. Delivery → Customer Notification
```

#### **File Movement Through Directories**
```
Archive/Q####.xls (Accepted Quote)
   ↓ FileCopy
WIP/J####.xls (Active Job)
   ↓ Job Completion
Archive/J####.xls (Completed Job)
```

### **Database Updates Throughout Lifecycle**

#### **Job Creation Updates**
```vba
' When job is created:
1. Add record to WIP.xls
2. Update Search.xls with job details
3. Set System_Status = \"New Job\"
4. Initialize operation tracking
```

#### **Progress Updates**
```vba
' During production:
1. Update WIP.xls with progress
2. Record operation completions
3. Update time estimates vs actual
4. Track quality checkpoints
```

#### **Completion Updates**
```vba
' When job is completed:
1. Remove record from WIP.xls
2. Update Search.xls with completion date
3. Set System_Status = \"Completed\"
4. Archive production records
```

---

## ⚠️ **Error Handling and Business Rules**

### **Job Validation Rules**

#### **Customer Order Number Requirement**
```vba
Private Function ValidateJobForm() As Boolean
    ' Customer order number is mandatory for job creation
    If Me.CustomerOrderNumber.Value = \"\" Then
        MsgBox \"Customer Order Number is required for job creation\"
        Me.CustomerOrderNumber.SetFocus
        ValidateJobForm = False
        Exit Function
    End If

    ' Check for duplicate customer order numbers
    If DuplicateOrderNumber(Me.CustomerOrderNumber.Value) Then
        MsgBox \"Customer Order Number already exists\"
        ValidateJobForm = False
        Exit Function
    End If
End Function
```

#### **Lead Time Validation**
```vba
' Lead time business rules
Private Sub ValidateLeadTimes()
    ' Minimum lead times by urgency
    Select Case Me.Job_Urgency.Value
        Case \"Normal\"
            If Me.Job_LeadTime.Value < 10 Then
                MsgBox \"Normal jobs require minimum 10 days lead time\"
            End If
        Case \"Break Down\"
            If Me.Job_LeadTime.Value < 5 Then
                MsgBox \"Break down jobs require minimum 5 days\"
            End If
        Case \"Urgent\"
            If Me.Job_LeadTime.Value < 7 Then
                MsgBox \"Urgent jobs require minimum 7 days\"
            End If
    End Select
End Sub
```

### **Multi-Part Job Coordination**

#### **Assembly Job Management**
```vba
' For jobs with multiple components
Private Sub HandleMultiPartJob()
    If Me.Compilation_TotalNumber.Value > 1 Then
        ' Create individual jobs for each component
        Dim partNumber As Integer
        For partNumber = 1 To Me.Compilation_TotalNumber.Value
            Call CreateComponentJob(partNumber)
        Next partNumber

        ' Create assembly job to coordinate components
        Call CreateAssemblyJob()
    End If
End Sub
```

#### **Component Synchronization**
```vba
' Ensure all components ready before assembly
Private Function AllComponentsReady(assemblyJobNumber As String) As Boolean
    ' Check WIP database for component job status
    ' Return True only if all components completed
    ' Used to control assembly job start
End Function
```

---

## 📈 **Reporting and Analytics**

### **Job Performance Metrics**

#### **Time Tracking Analysis**
```vba
' Compare estimated vs actual times
Private Sub AnalyzeJobPerformance(jobNumber As String)
    Dim totalEstimated As Double, totalActual As Double
    Dim operationNum As Integer

    For operationNum = 1 To 15
        totalEstimated = totalEstimated + Operation_EstTime(operationNum)
        totalActual = totalActual + Operation_ActualTime(operationNum)
    Next operationNum

    Dim efficiency As Double
    efficiency = (totalEstimated / totalActual) * 100

    ' Store efficiency data for reporting
    Call RecordJobEfficiency(jobNumber, efficiency)
End Sub
```

#### **Operator Performance Tracking**
```vba
' Track operator productivity and quality
Private Sub RecordOperatorMetrics(operatorName As String, operationNum As Integer)
    ' Time taken for operation
    ' Quality results
    ' Rework requirements
    ' Used for operator evaluation and training
End Sub
```

---

## 🔧 **Development Guidelines**

### **Extending Job Management**

#### **Adding New Operation Types**
```vba
' To add new manufacturing operations:
' 1. Update operation type dropdown lists
' 2. Add operation-specific validation
' 3. Update time estimation formulas
' 4. Add quality checkpoints if needed

Private Sub LoadOperationTypes()
    With OperationType_Dropdown
        .AddItem \"Machining\"
        .AddItem \"Welding\"
        .AddItem \"Assembly\"
        .AddItem \"Inspection\"
        ' Add new operation types here
        .AddItem \"3D Printing\"
        .AddItem \"Laser Cutting\"
    End With
End Sub
```

#### **Custom Job Card Fields**
```vba
' Adding job-specific information
' 1. Add controls to FJobCard.frm
' 2. Update PopulateJobFile subroutine
' 3. Modify SaveWIPCode to include new fields
' 4. Update reporting as needed

Private Sub PopulateCustomFields()
    ' Example: Special handling requirements
    ws.Range(\"SpecialInstructions\").Value = Me.SpecialInstructions.Value
    ws.Range(\"SafetyRequirements\").Value = Me.SafetyRequirements.Value
End Sub
```

---

## 🔍 **Next Steps**

After understanding Job Management:

1. **Study [Interface Navigation](SUBSYSTEM_06_INTERFACE_NAVIGATION.md)** - See how jobs are displayed and managed in Main interface
2. **Review [Reporting & WIP](SUBSYSTEM_07_REPORTING_WIP.md)** - Understand WIP reporting and job analysis
3. **Examine [Search Database](SUBSYSTEM_08_SEARCH_DATA.md)** - See how jobs are indexed and tracked
4. **Practice Job Workflow** - Create test jobs and follow complete lifecycle
5. **Customize Operations** - Add operation types or modify job card fields

**Ready for interface navigation? Continue to [Interface Navigation Subsystem](SUBSYSTEM_06_INTERFACE_NAVIGATION.md)**