# PCS V2 Workflow Analysis Documentation

## Overview

PCS V2 maintains the complete Enquiry → Quote → Job → Archive workflow from the original system while providing cleaner code organization through module-based architecture. This document analyzes the complete process flows, data transformations, and system interactions.

## Primary Workflow Sequence

### **Core Business Process**
```
Enquiry Creation → Quote Generation → Quote Acceptance → Job Creation → Job Completion → Archive
```

---

## Workflow Stage 1: Enquiry Creation

### **Process Flow**

**Entry Points**:
- Main.frm → Add_Enquiry_Click() → UserInterface.AddEnquiry()
- Direct form launch: FEnquiry.frm or FrmEnquiry.frm

**Code Execution Chain**:
```
1. Form Event: FEnquiry.SaveQ_Click()
2. Module Call: WorkflowManagement.SaveEnquiry(Me)
3. Form Validation: WorkflowManagement.ValidateEnquiryFormData()
4. Data Population: Form controls → EnquiryData structure
5. Business Logic: BusinessLogic.CreateEnquiry(EnquiryInfo)
6. File Operations: Create Templates\_Enq.xls → Enquiries\E00001.xls
7. Search Update: BusinessLogic.UpdateSearchDatabase()
8. Customer Check: Create customer file if new
```

### **Data Transformations**

**Form → Data Structure**:
```vba
With EnquiryInfo
    .CustomerName = Trim(EnquiryForm.Customer.Value)
    .ContactPerson = Trim(EnquiryForm.Contact.Value)
    .CompanyPhone = Trim(EnquiryForm.Phone.Value)
    .CompanyFax = Trim(EnquiryForm.Fax.Value)
    .Email = Trim(EnquiryForm.Email.Value)
    .ComponentDescription = Trim(EnquiryForm.Component_Description.Value)
    .ComponentCode = Trim(EnquiryForm.Component_Code.Value)
    .MaterialGrade = Trim(EnquiryForm.Component_Grade.Value)
    .Quantity = CLng(EnquiryForm.Component_Quantity.Value)
    .SearchKeywords = .CustomerName & " " & .ComponentDescription & " " & .ComponentCode
End With
```

### **File System Operations**

**Files Created**:
- **Enquiry File**: `Enquiries\E00001.xls` (from `Templates\_Enq.xls`)
- **Customer File**: `Customers\CustomerName.xls` (if new customer)

**Files Updated**:
- **Search Database**: `Search.xls` (new enquiry record)
- **Number Tracking**: Updated via DataOperations.GetNextEnquiryNumber()

### **Excel File Structure (Enquiry)**

**Admin Sheet Data Mapping**:
```
Cell B2: EnquiryNumber (E00001)
Cell B3: CustomerName
Cell B4: ContactPerson
Cell B5: CompanyPhone
Cell B6: CompanyFax
Cell B7: Email
Cell B8: ComponentDescription
Cell B9: ComponentCode
Cell B10: MaterialGrade
Cell B11: Quantity
Cell B12: DateCreated
```

---

## Workflow Stage 2: Quote Generation

### **Process Flow**

**Entry Points**:
- Main.frm → Make_Quote_Click() → UserInterface.CreateQuote()
- Form: FQuote.LoadFromEnquiry(EnquiryPath)

**Code Execution Chain**:
```
1. Quote Form Launch: FQuote.LoadFromEnquiry()
2. Data Loading: WorkflowManagement.LoadQuoteFromEnquiry()
3. Form Initialization: WorkflowManagement.InitializeQuoteForm()
4. User Input: Pricing, lead times, validity dates
5. Form Event: FQuote.SaveQuote_Click()
6. Module Call: WorkflowManagement.SaveQuote(Me)
7. Business Logic: BusinessLogic.CreateQuote(QuoteInfo)
8. File Movement: Enquiries\E00001.xls → Quotes\Q00001.xls
9. Search Update: Status change to "Quote Created"
```

### **Data Transformations**

**Enquiry → Quote Data Enhancement**:
```vba
' Inherited from Enquiry
QuoteInfo.EnquiryNumber = EnquiryInfo.EnquiryNumber
QuoteInfo.CustomerName = EnquiryInfo.CustomerName
QuoteInfo.ComponentDescription = EnquiryInfo.ComponentDescription
QuoteInfo.ComponentCode = EnquiryInfo.ComponentCode
QuoteInfo.Quantity = EnquiryInfo.Quantity

' Added in Quote Stage
QuoteInfo.QuoteNumber = "Q00001"  ' Auto-generated
QuoteInfo.UnitPrice = [User Input]
QuoteInfo.TotalPrice = Quantity * UnitPrice
QuoteInfo.LeadTimeDays = [User Input, Default: 14]
QuoteInfo.ValidUntil = DateAdd("d", 30, Now)  ' Default 30 days
QuoteInfo.DateCreated = Now
```

### **Business Rules Applied**

**Quote Validation**:
- All enquiry fields must be valid
- Unit price must be > 0
- Lead time must be > 0
- Valid until date must be future date

**Automatic Calculations**:
- Total Price = Unit Price × Quantity
- Quote validity = Current Date + 30 days (default)

### **File System Operations**

**File Movement**:
```
Source: Enquiries\E00001.xls
Target: Quotes\Q00001.xls
Action: Move file and update all internal references
```

**Search Database Update**:
- Status: "To Quote" → "Quote Created"
- New fields: Quote number, pricing information
- Updated keywords: Include pricing terms

---

## Workflow Stage 3: Quote Acceptance and Job Creation

### **Process Flow**

**Entry Points**:
- Main.frm → AcceptQuote_Click() → UserInterface.AcceptQuote()
- Form: FAcceptQuote.LoadQuote(QuotePath)

**Code Execution Chain**:
```
1. Accept Quote Form: FAcceptQuote.LoadQuote()
2. Quote Loading: WorkflowManagement.LoadQuoteForAcceptance()
3. User Input: Customer order number, urgency, special requirements
4. Form Event: FAcceptQuote.butSAVE_Click()
5. Module Call: WorkflowManagement.AcceptQuote(Me, QuotePath)
6. Business Logic: BusinessLogic.CreateJobFromQuote()
7. File Movement: Quotes\Q00001.xls → WIP\J00001.xls
8. WIP Database: BusinessLogic.UpdateWIPDatabase()
9. Job Card Activation: Enable production worksheet
```

### **Data Transformations**

**Quote → Job Data Enhancement**:
```vba
' Inherited from Quote
JobInfo.EnquiryNumber = QuoteInfo.EnquiryNumber
JobInfo.QuoteNumber = QuoteInfo.QuoteNumber
JobInfo.CustomerName = QuoteInfo.CustomerName
JobInfo.ComponentInfo = [All component data]
JobInfo.Pricing = QuoteInfo.TotalPrice

' Added in Job Stage
JobInfo.JobNumber = "J00001"  ' Auto-generated
JobInfo.CustomerOrderNumber = [Required User Input]
JobInfo.Urgency = ["Normal", "Break Down", "Urgent"]
JobInfo.StartDate = Now
JobInfo.DueDate = CalculateDueDate(Urgency, LeadTime)
JobInfo.Status = "Active"
JobInfo.AssignedOperators = [Operations Planning]
```

### **Business Rules Applied**

**Job Creation Validation**:
- Customer order number is required
- Must reference valid quote
- Due date calculated based on urgency and lead time
- Job number must be unique

**Lead Time Calculations**:
```vba
Select Case Urgency
    Case "Normal": DueDate = StartDate + 14 days
    Case "Break Down": DueDate = StartDate + 7 days
    Case "Urgent": DueDate = StartDate + 10 days
End Select
```

### **File System Operations**

**File Movement**:
```
Source: Quotes\Q00001.xls
Target: WIP\J00001.xls
Action: Move file, update to job template structure
```

**Database Updates**:
- **Search.xls**: Status "Quote Created" → "Job Active"
- **WIP.xls**: New job record with all tracking information

---

## Workflow Stage 4: Production Management

### **Process Flow**

**Entry Points**:
- Main.frm → But_EditJC_Click() → UserInterface.EditJobCard()
- Form: FJobCard (Job Card Management)

**Code Execution Chain**:
```
1. Job Selection: From WIP file list in Main.frm
2. Job Card Form: FJobCard loads with job data
3. Production Planning: Operations, operators, schedules
4. Progress Updates: Track completion status
5. Job Completion: FJobCard.SaveJobCard_Click()
6. Archive Process: Move to Archive directory
7. WIP Update: Remove from active jobs list
```

### **Production Data Tracking**

**Operation Management**:
```vba
' Up to 15 operations per job
Operation01_Type = "Machining"
Operation01_Operator = "John Smith"
Operation01_Comment = "Rough turning to spec"
Operation01_Status = "Completed"
Operation01_Hours = 2.5

' ... Operations 02 through 15 as needed
```

**Job Progress Tracking**:
- Start date and actual start date
- Estimated vs actual hours per operation
- Material usage and waste tracking
- Quality control checkpoints

### **File System Operations**

**Active Job File**: `WIP\J00001.xls`
- Regular updates during production
- Progress tracking and time logging
- Operations completion status

---

## Workflow Stage 5: Job Completion and Archive

### **Process Flow**

**Entry Points**:
- Main.frm → CloseJob_Click() → UserInterface.CloseJob()
- FJobCard.SaveJobCard_Click() (completion)

**Code Execution Chain**:
```
1. Job Completion: All operations marked complete
2. Final Validation: Quality checks, delivery confirmation
3. File Archive: WIP\J00001.xls → Archive\J00001.xls
4. WIP Database: Remove from active jobs
5. Search Update: Status "Job Active" → "Job Completed"
6. Invoicing: Generate invoice reference (optional)
```

### **Completion Data**

**Final Job Data**:
```vba
JobInfo.CompletionDate = Now
JobInfo.Status = "Completed"
JobInfo.ActualHours = [Sum of all operation hours]
JobInfo.DeliveryDate = [Customer delivery date]
JobInfo.InvoiceNumber = [Generated if invoicing enabled]
JobInfo.FinalNotes = [Completion notes and customer feedback]
```

### **File System Operations**

**File Archive**:
```
Source: WIP\J00001.xls
Target: Archive\J00001.xls
Action: Move completed job to archive directory
```

**Database Cleanup**:
- **WIP.xls**: Remove completed job record
- **Search.xls**: Update status to "Archived"

---

## Cross-Workflow Operations

### **Search Database Management**

**Real-time Updates**: Every workflow stage updates `Search.xls`

**Search Record Structure**:
```vba
Type SearchRecord
    RecordType As String        ' "Enquiry", "Quote", "Job"
    Number As String           ' E00001, Q00001, J00001
    CustomerName As String
    ComponentDescription As String
    Status As String           ' Workflow stage indicator
    FilePath As String         ' Current file location
    Keywords As String         ' Searchable text
    DateCreated As Date
    LastModified As Date
End Type
```

**Status Progression**:
```
"To Quote" → "Quote Created" → "Quote Submitted" → "Job Active" → "Job Completed" → "Archived"
```

### **WIP Database Management**

**Active Job Tracking**: `WIP.xls` contains all jobs currently in production

**WIP Record Structure**:
- Job identification (numbers, customer)
- Production schedule (start, due dates)
- Operations breakdown (types, operators, status)
- Progress tracking (completion percentages)
- Priority and urgency indicators

### **Customer Database Integration**

**Automatic Customer Creation**:
```vba
If Not DataOperations.FileExists(CustomerPath) Then
    BusinessLogic.CreateNewCustomer(CustomerName)
End If
```

**Customer File Structure**: `Customers\CustomerName.xls`
- Contact information
- Job history references
- Credit and payment terms
- Special requirements and notes

---

## Error Handling and Recovery Patterns

### **Transaction-Like Operations**

**File Movement Safety**:
```vba
' Backup before moving
BackupPath = CreateBackup(SourceFile)
Try
    MoveFile(Source, Target)
    UpdateDatabases()
    DeleteBackup(BackupPath)
Catch Error
    RestoreFromBackup(BackupPath)
    RollbackDatabaseChanges()
End Try
```

### **Workflow State Recovery**

**Incomplete Transitions**:
- Files found in wrong directories are flagged for manual review
- Database inconsistencies are logged and reported
- Partial updates are completed or rolled back

**Data Validation at Each Stage**:
- Required fields validation before workflow advancement
- Business rule checking (dates, numbers, references)
- File integrity verification

---

## Performance Considerations

### **File Operation Optimization**

**Batch Operations**:
- Search database updates batched where possible
- Multiple file operations grouped together
- Reduced file open/close cycles

**Caching Strategies**:
- Customer lists cached during form initialization
- Component codes and grades cached for session
- Recently accessed files kept in memory

### **Database Maintenance**

**Search Database Optimization**:
- Regular cleanup of old records via `Search_Sync.SeachSYNC()`
- Index optimization for common search patterns
- Archival of historical data older than threshold

---

## Workflow Dependencies

### **Module Interaction Pattern**

```
Forms → WorkflowManagement → BusinessLogic → DataOperations
  ↓           ↓                    ↓             ↓
SystemCore ←---+--------------------+-------------+
(Error Handling, Validation, Data Types)
```

### **File System Dependencies**

**Required Directory Structure**:
```
Root/
├── Enquiries/          # Active enquiries
├── Quotes/             # Generated quotes
├── WIP/               # Active jobs
├── Archive/           # Completed jobs
├── Templates/         # System templates
├── Customers/         # Customer records
├── Contracts/         # Job templates
└── Images/           # Technical drawings
```

**Critical Files**:
- `Search.xls`: Master search database
- `WIP.xls`: Active jobs database
- `Templates\_Enq.xls`: Enquiry template
- `Templates\_Quote.xls`: Quote template
- `Templates\_Job.xls`: Job template

The V2 workflow system successfully maintains all original business processes while providing improved error handling, validation, and code organization through the modular architecture.