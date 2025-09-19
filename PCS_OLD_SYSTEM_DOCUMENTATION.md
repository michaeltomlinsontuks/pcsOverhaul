# PCS Interface System Documentation

## System Overview

The PCS Interface System is a VBA-based production control system that manages the complete workflow from enquiries through quotes to jobs and work-in-progress reporting. The system is built around Excel workbooks with VBA forms and modules, utilizing a file-based architecture where each record (enquiry, quote, job) is stored as a separate Excel file in categorized directories.

### Core Architecture
- **File Structure**: Main master path with subdirectories (Enquiries/, Quotes/, WIP/, Archive/, Templates/, Customers/, Contracts/, Images/)
- **Data Storage**: Each business record stored as individual Excel file with standardized "Admin" sheet containing metadata
- **User Interface**: Collection of VBA UserForms providing data entry and management interfaces
- **Search System**: Centralized search functionality via Search.xls with historical tracking
- **Reporting**: WIP reports with multiple sorting and filtering options

### Core Workflow
**Enquiries → Quotes → Jobs → Job Cards → WIP Reports → Archive**

---

## Subsystem 1: Core Utilities and Infrastructure

### Modules

#### a_Main.bas
**Purpose**: System entry point and initialization
- `ShowMenu()`: Main entry point that sets master path and displays main interface
- `sadf()`: Utility function for decrementing cell values (appears to be test/development code)

#### Open_Book.bas
**Purpose**: Workbook management utility
- `OpenBook(File As String, RO As Boolean)`: Opens Excel workbooks with read-only option
- **Dependencies**: Used throughout all subsystems for file access

#### Check_Dir.bas
**Purpose**: Directory management utility
- `CheckDir(Direc As String)`: Creates directories if they don't exist, then changes to directory
- **Used By**: File creation and organization functions across all subsystems

#### GetUserName64.bas / GetUserNameEx.bas
**Purpose**: User identification for 32/64-bit compatibility
- `Get_User_Name()`: Retrieves current Windows username
- **Note**: Two versions exist for 32-bit and 64-bit Excel compatibility
- **Dependencies**: Windows API calls (advapi32.dll)

#### GetValue.bas
**Purpose**: Reading data from closed Excel workbooks
- `GetValue(path, File, sheet, ref)`: Retrieves specific cell values from closed workbooks
- **Dependencies**: ExecuteExcel4Macro function
- **Used By**: All forms for loading data, search operations, template access

#### Very_HiddenSheet.bas
**Purpose**: Worksheet visibility management
- `VeryHiddenSheet(SheetNam As String)`: Hides sheets completely
- `ShowSheet(SheetNam As String)`: Makes sheets visible
- **Used By**: Template and data protection functions

#### Delete_Sheet.bas
**Purpose**: Sheet deletion without user prompts
- `DeleteSheet(SheetName As String)`: Deletes worksheets silently
- **Used By**: Cleanup operations, template processing

#### RemoveCharacters.bas
**Purpose**: String manipulation utilities
- `Remove_Characters(Str As String)`: Removes special characters (/, :, space)
- `Insert_Characters(Str As String)`: Formats display strings by adding spaces for readability
- **Used By**: Report generation, form display formatting

### Data Structures

#### Excel File Structure (Standard Template)
- **Admin Sheet**: Contains metadata in key-value pairs (Column A: Field Names, Column B: Values)
- **Job Card Sheet**: Production worksheet with operations, drawings, and job details
- **Additional Sheets**: Varies by file type (quotes may have additional calculation sheets)

### Key Field Mappings (Admin Sheet)
- **File_Name**: Primary identifier
- **Enquiry_Number**: E-prefix numbers
- **Quote_Number**: Q-prefix numbers
- **Job_Number**: J-prefix numbers
- **System_Status**: Workflow state tracking
- **Customer**: Client information
- **Component_***: Part specifications
- **Operation01-15_***: Manufacturing operations data

---

## Subsystem 2: Number Generation and Tracking

### Modules

#### Calc_Numbers.bas
**Purpose**: Sequential number generation for enquiries, quotes, and jobs

##### Functions
- `Calc_Next_Number(Typ As String)`: Calculates next available number for type
  - **Parameters**: Typ - "E" (Enquiry), "Q" (Quote), "J" (Job)
  - **Returns**: Integer - Next sequential number
  - **Logic**: Scans templates directory for highest existing number + 1
  - **Dependencies**: Main.Main_MasterPath, Dir() function

- `Confirm_Next_Number(Typ As String)`: Commits number by updating template file
  - **Parameters**: Typ - Type identifier
  - **Returns**: Integer - Confirmed number
  - **Side Effects**: Updates template file, deletes old template, creates new one
  - **Dependencies**: FileCopy, Kill operations

### Data Flow
1. Form requests next number via `Calc_Next_Number()`
2. System scans templates directory for pattern "E - ####.TXT", "Q - ####.TXT", "J - ####.TXT"
3. Extracts highest number and increments
4. `Confirm_Next_Number()` commits the reservation by updating template file

### File Dependencies
- **Templates Directory**: Contains number tracking files
- **Naming Convention**: "[Type] - [Number].TXT"

---

## Subsystem 3: Enquiry Management

### Forms

#### FEnquiry.frm / FrmEnquiry.frm
**Purpose**: Enquiry data entry and management (Note: Two similar forms exist - likely versioning issue)

##### Key Controls
- **Enquiry_Number**: Auto-generated E-prefix number
- **Customer**: Dropdown populated from Customers directory
- **Component_Code**: Dropdown from price list
- **Component_Grade**: Dropdown from Component_Grades.xls
- **Component_Description**: Part description
- **Component_Quantity**: Quantity required
- **Enquiry_Date**: Date picker
- **ContactPerson**: Populated based on customer selection
- **Notes**: Free text field with template

##### Key Events
- `UserForm_Activate()`: Loads customer list, component codes, grades
- `SaveQ_Click()`: Saves enquiry to Enquiries directory
- `AddMore_Click()`: Saves current and starts new enquiry
- `AddNewClient_Click()`: Creates new customer record
- `Customer_Change()`: Populates contact person dropdown

##### Dependencies
- **Calc_Numbers.bas**: For enquiry number generation
- **List_Files()**: For populating customer dropdown
- **SaveSearchCode.bas**: For updating central search index
- **Templates**: _Enq.xls, _client.xls, price list.xls, Component_Grades.xls

### Modules

#### SaveSearchCode.bas
**Purpose**: Central search index management
- `SaveRowIntoSearch(frm As Object)`: Updates Search.xls with form data
- **Process**: Opens Search.xls, finds/creates row, maps form controls to columns, sorts by date
- **Dependencies**: OpenBook(), Excel sorting functions

### Data Flow
1. User opens enquiry form
2. System loads templates and reference data
3. User enters enquiry details
4. System generates enquiry number
5. Data saved to Enquiries/[EnquiryNumber].xls
6. Search index updated
7. File moved through workflow: Enquiries → Quotes → WIP → Archive

### Field Mappings
| Form Control | Excel Field | Type | Notes |
|--------------|-------------|------|-------|
| Enquiry_Number | Enquiry_Number | String | Auto-generated E-prefix |
| Customer | Customer | String | From customers directory |
| Component_Code | Component_Code | String | From price list |
| Component_Description | Component_Description | String | Part description |
| Component_Quantity | Component_Quantity | Integer | Required quantity |
| System_Status | System_Status | String | "To Quote" |

---

## Subsystem 4: Quote Management

### Forms

#### FQuote.frm
**Purpose**: Convert enquiries to quotes with pricing

##### Key Controls
- **Quote_Number**: Auto-generated Q-prefix number
- **Enquiry_Number**: Source enquiry (read-only)
- **Component_Price**: Pricing information
- **Job_LeadTime**: Delivery timeframe
- **Quote_Date**: Auto-filled current date

##### Key Events
- `UserForm_Activate()`: Loads data from source enquiry file
- `SaveQuote_Click()`: Converts enquiry to quote
- `Search_Component_code_Click()`: Opens search functionality

##### Dependencies
- **Source File**: Enquiry file from Enquiries directory
- **Target Location**: Quotes directory
- **Search Update**: Updates Search.xls index

### Data Flow
1. User selects enquiry from main interface
2. Quote form opens with enquiry data pre-populated
3. User adds pricing and lead time information
4. System generates quote number
5. File moved from Enquiries to Quotes directory
6. Search index updated with new status
7. Original enquiry file deleted

### Field Transformations
| Field | Enquiry → Quote |
|-------|----------------|
| System_Status | "To Quote" → "New Quote" |
| Quote_Number | Generated |
| Component_Price | Added by user |
| Job_LeadTime | Added by user (default: 14) |

---

## Subsystem 5: Job Creation and Acceptance

### Forms

#### FAcceptQuote.frm
**Purpose**: Accept quotes and convert to jobs

##### Key Controls
- **Job_Number**: Auto-generated J-prefix number
- **CustomerOrderNumber**: Required field for job creation
- **Job_Urgency**: Dropdown (Normal, Break Down, Urgent)
- **Job_LeadTime**: Auto-calculated based on urgency
- **Compilation_SequenceNumber/TotalNumber**: Multi-part job support

##### Key Events
- `butSAVE_Click()`: Main save operation
- `Job_Urgency_Change()`: Auto-sets lead times
- `UserForm_Activate()`: Loads quote data

##### Dependencies
- **Source**: Archive directory (quotes are moved here when submitted)
- **Target**: WIP directory
- **Job Card**: Activates Job Card sheet for production

#### FJG.frm (Job Generator)
**Purpose**: Advanced job creation with operations planning

##### Key Controls
- **Operation01-15_Type**: Manufacturing operation types
- **Operation01-15_Operator**: Assigned operators
- **Operation01-15_Comment**: Operation instructions
- **Job_PicturePath**: Technical drawing reference
- **Compilation controls**: Multi-part job handling

##### Key Events
- `butSaveJG_Click()`: Save job with operations
- `but_SaveAsCTItem_Click()`: Save as contract template
- `JobCardTemplates_Click()`: Load operation templates
- `CopyFromJobCard_Click()`: Copy operations from existing job

### Modules

#### SaveWIPCode.bas
**Purpose**: WIP reporting database management
- `SaveInfoIntoWIP(frm As Object)`: Updates WIP.xls master list
- **Process**: Opens WIP.xls, finds/creates row, updates job status

### Data Flow
1. Quote marked as "Quote Submitted" in archive
2. Accept Quote form loaded with quote data
3. Customer order number entered (required)
4. Job number generated
5. Job urgency set (affects lead time calculation)
6. For multi-part jobs: sequence tracking maintained
7. File moved to WIP directory
8. WIP master list updated
9. Job Card sheet activated for production

### Lead Time Calculations
| Urgency | Lead Time (Days) |
|---------|------------------|
| Normal | 14 |
| Break Down | 7 |
| Urgent | 10 |

---

## Subsystem 6: Production Management

### Forms

#### FJobCard.frm
**Purpose**: Production job card management

##### Key Controls
- **Operation01-15 controls**: Production operations
- **Job_PicturePath**: Technical drawings
- **System_Status**: Job status tracking

##### Key Events
- `SaveJobCard_Click()`: Complete job and move to archive
- `CopyFromJobCard_Click()`: Copy operations from existing job
- `JobCardTemplates_Click()`: Load operation templates

#### Main.frm
**Purpose**: Central system interface and navigation

##### Key Controls
- **lst**: Main file listing
- **Main_MasterPath**: System root directory
- **Checkboxes**: Enquiries, Quotes, WIP, Archive, JobsInWIP, Thirties
- **Notice_***: File count displays

##### Key Events
- **File Type Toggles**: Enquiries_Click(), Quotes_Click(), WIP_Click(), Archive_Click()
- **Actions**: Make_Quote_Click(), createjob_Click(), CloseJob_Click()
- **Navigation**: lst_Click(), Lst_DblClick()
- **Reports**: WIPReport_Click(), Search_Click()

### Modules

#### RefreshMain.bas
**Purpose**: Main interface data refresh
- `Refresh_Main()`: Updates file listings based on selected categories
- **Dependencies**: List_Files(), Check_Files(), CheckUpdates()

#### Check_Updates.bas
**Purpose**: Real-time file monitoring
- `CheckUpdates()`: Scheduled update checker (5-minute intervals)
- `Check_Files(path As String)`: Counts files in directory
- `StopCheck()`: Cancels scheduled updates
- **Dependencies**: Application.OnTime for scheduling

#### a_ListFiles.bas
**Purpose**: Directory file enumeration
- `List_Files(path As String, frm As Object)`: Populates listboxes with files
- **Special Handling**: Marks new quotes (*), accepted quotes (*) in WIP
- **Dependencies**: GetValue() for status checking

### Data Flow
1. Main interface displays available files by category
2. User selects file type (Enquiries, Quotes, WIP, Archive)
3. System scans appropriate directory
4. Files displayed with status indicators
5. User selects file for editing/viewing
6. Appropriate form loaded with file data
7. Changes saved back to file and search index updated

---

## Subsystem 7: Reporting and WIP Management

### Forms

#### fwip.frm
**Purpose**: Work-in-Progress reporting interface

##### Key Controls
- **Report Type Options**: ROperation, ROperator, RDueDate, RWIP
- **Sorting Options**: Job_DueDate, Office_Customer, Workshop_Customer
- **Job Number Sorting**: Office_JobNumber, Workshop_JobNumber

##### Key Events
- `Go_Click()`: Main report generation function

### Functions
- **Operation Reports**: Groups jobs by operation type
- **Operator Reports**: Groups jobs by assigned operator
- **Due Date Reports**: Sorts by customer delivery dates
- **Customer Reports**: Sorts by customer (Office vs Workshop views)

#### FList.frm
**Purpose**: Generic list selection dialog
- **Used By**: Template selection, file browsing operations

### Modules

#### Search_Sync.bas
**Purpose**: Search database maintenance
- `SeachSYNC()`: Synchronizes search database with search history
- **Password Protected**: Requires "KJB" password
- **Cleanup**: Removes old records based on number thresholds

### Data Flow
1. User selects report type and parameters
2. System loads WIP.xls master file
3. Data sorted according to selected criteria
4. New workbook created with filtered/sorted data
5. Multiple sheets created for different groupings
6. Report saved to Templates directory
7. Formatting applied (borders, headers, column sizing)

---

## Subsystem 8: Search and Navigation

### Core Search Files
- **Search.xls**: Main search database
- **Search History.xls**: Historical search records

### Search Integration
All forms include search functionality through:
- **SaveSearchCode.bas**: Updates main search index
- **Form Integration**: Each save operation updates search
- **Real-time Updates**: Search updated on every record change

### Search Fields
- File_Name (primary key)
- Enquiry_Number, Quote_Number, Job_Number
- Customer, System_Status
- Component information
- Dates (enquiry, quote, start, due)

---

## Cross-System Analysis

### Function Call Chains

#### Primary Workflow Chain
1. **a_Main.ShowMenu()** → Main.frm
2. **Main.Add_Enquiry_Click()** → FrmEnquiry.frm
3. **FrmEnquiry.SaveQ_Click()** → Calc_Numbers.Calc_Next_Number() → SaveSearchCode.SaveRowIntoSearch()
4. **Main.Make_Quote_Click()** → FQuote.frm
5. **FQuote.SaveQuote_Click()** → Calc_Numbers.Confirm_Next_Number() → SaveSearchCode.SaveRowIntoSearch()
6. **Main.createjob_Click()** → FAcceptQuote.frm
7. **FAcceptQuote.butSAVE_Click()** → SaveWIPCode.SaveInfoIntoWIP()
8. **Main.OpenJob_Click()** → FJobCard.frm
9. **FJobCard.SaveJobCard_Click()** → Archive process

#### Utility Function Dependencies
- **GetValue()**: Used by ALL forms for data loading (defined in 4 different modules - code duplication)
- **OpenBook()**: Used by ALL file operations
- **CheckDir()**: Used by save operations
- **List_Files()**: Used by Main interface and form population
- **Insert_Characters()**: Used by forms for display formatting

### Data Flow Mapping

#### File Movement Workflow
```
Templates/[Type] - [Number].TXT (Number tracking)
    ↓
Enquiries/[EnquiryNumber].xls
    ↓ (Quote creation)
Quotes/[QuoteNumber].xls
    ↓ (Quote submission)
Archive/[QuoteNumber].xls
    ↓ (Job acceptance)
WIP/[JobNumber].xls
    ↓ (Job completion)
Archive/[JobNumber].xls
```

#### Search Database Flow
```
Form Data → SaveSearchCode.SaveRowIntoSearch() → Search.xls
Search.xls → Search_Sync.SeachSYNC() → Search History.xls
```

#### WIP Reporting Flow
```
WIP/[JobNumber].xls → SaveWIPCode.SaveInfoIntoWIP() → WIP.xls
WIP.xls → fwip.Go_Click() → Templates/[ReportType].xls
```

### Field Mismatches and Data Inconsistencies

#### Critical Issues Identified

##### 1. Duplicate Function Definitions
- **GetValue()**: Defined in 4 separate modules (GetValue.bas, Module1.bas, FAcceptQuote.frm, FQuote.frm)
- **Insert_Characters()**: Defined in RemoveCharacters.bas and FJG.frm
- **Risk**: Inconsistent behavior, maintenance issues

##### 2. Form Duplication
- **FEnquiry.frm vs FrmEnquiry.frm**: Two enquiry forms with similar but not identical functionality
- **fwip.frm vs fwip_modified.frm**: Two WIP report forms
- **Risk**: User confusion, data inconsistency

##### 3. Inconsistent Field Naming
- **Quote_Number vs Quote_Nmber**: Typo in SaveWIPCode.bas line 19
- **Various Date Fields**: Job_StartDate, Enquiry_Date, Quote_Date (inconsistent naming pattern)

##### 4. Hard-coded Paths and Dependencies
- **Module3.bas**: Contains hard-coded export path "C:\Users\Michael Tomlinson\Downloads\20081222\Interface_VBA\"
- **Multiple Forms**: Reference specific template files without error handling

##### 5. Search Field Inconsistencies
- **Column Mapping**: Forms use control name matching for search updates, but no validation of column existence
- **Sort Keys**: Different forms use different sort columns (Range("e2"), Range("A2"))

##### 6. Number Generation Race Conditions
- **Calc_Numbers.bas**: Gap between Calc_Next_Number() and Confirm_Next_Number() could cause number conflicts in multi-user environment

##### 7. File Locking Issues
- **Read-Only Loops**: Multiple functions contain infinite loops checking for read-only files without timeout

##### 8. Error Handling Inconsistencies
- **Some functions**: Comprehensive error handling with Resume
- **Other functions**: Basic or no error handling
- **GetValue functions**: Different error handling strategies across implementations

### Dead Code Identification

#### Potentially Unused Code
1. **Module2.bas**: Contains only commented-out Leeora() function - appears to be legacy security code
2. **sadf() function** in a_Main.bas: Appears to be test code
3. **Component_code_Change()**: Commented out in both enquiry forms
4. **Multiple TestGetValue functions**: In GetValue.bas - development/testing code
5. **Hard-coded user exclusions**: In Module2.bas - legacy security bypasses

#### Template-Related Dead Code
- **Calendar Functions**: References to ShowCalender function that may not exist
- **Price List Integration**: Some forms reference Price List.xls but don't always close it properly

### Performance Issues

#### Identified Bottlenecks
1. **Search Operations**: Linear search through all files for each operation
2. **File I/O**: Repeated opening/closing of same files
3. **Form Loading**: Multiple file reads during form initialization
4. **WIP Reports**: Large dataset processing without optimization

#### Optimization Opportunities
1. **Caching**: Master path and frequently accessed data
2. **Batch Operations**: Group file operations
3. **Index Usage**: Better search indexing strategy
4. **Connection Pooling**: Reduce file open/close operations

---

## Recommendations for Refactoring

### High Priority
1. **Consolidate Duplicate Functions**: Create single authoritative version of GetValue(), Insert_Characters()
2. **Remove Duplicate Forms**: Standardize on single enquiry form, single WIP form
3. **Fix Field Name Typos**: Correct Quote_Nmber and other inconsistencies
4. **Implement Proper Error Handling**: Standardized error handling across all modules
5. **Remove Hard-coded Paths**: Make all paths relative to master path

### Medium Priority
1. **Standardize Naming Conventions**: Consistent field naming across all forms
2. **Improve File Locking**: Add timeouts to read-only loops
3. **Optimize Search Performance**: Implement proper indexing
4. **Clean Up Dead Code**: Remove unused functions and commented code

### Low Priority
1. **Performance Optimization**: Implement caching and batch operations
2. **Enhanced Logging**: Add comprehensive audit trail
3. **User Interface Improvements**: Standardize form layouts and behaviors

---

## Field Reference Guide

### Complete Field Mapping
| Field Name | Type | Used In | Purpose |
|------------|------|---------|---------|
| File_Name | String | All | Primary identifier |
| Enquiry_Number | String | Enquiry, Quote, Job | Enquiry tracking |
| Quote_Number | String | Quote, Job | Quote tracking |
| Job_Number | String | Job, WIP | Job tracking |
| System_Status | String | All | Workflow state |
| Customer | String | All | Client identification |
| ContactPerson | String | Enquiry, Quote | Client contact |
| Component_Code | String | All | Part number |
| Component_Description | String | All | Part description |
| Component_Quantity | Integer | All | Required quantity |
| Component_Price | Currency | Quote, Job | Pricing |
| Component_Grade | String | All | Material specification |
| Component_DrawingNumber_SampleNumber | String | All | Technical reference |
| Enquiry_Date | Date | Enquiry | Initial inquiry date |
| Quote_Date | Date | Quote | Quote creation date |
| Job_StartDate | Date | Job | Production start |
| CustomerDelivery_Date | Date | Job | Customer delivery deadline |
| Job_WorkshopDueDate | Date | Job | Internal due date |
| Job_LeadTime | Integer | Quote, Job | Lead time in days |
| Job_Urgency | String | Job | Priority level |
| CustomerOrderNumber | String | Job | Client PO number |
| Operation01-15_Type | String | Job | Manufacturing operations |
| Operation01-15_Operator | String | Job | Assigned operators |
| Operation01-15_Comment | String | Job | Operation instructions |
| Job_PicturePath | String | Job | Technical drawing path |
| Invoice_Number | String | Closed Jobs | Billing reference |
| Invoice_Date | Date | Closed Jobs | Billing date |
| Compilation_SequenceNumber | Integer | Multi-part Jobs | Part sequence |
| Compilation_TotalNumber | Integer | Multi-part Jobs | Total parts |

This documentation provides a comprehensive overview of the PCS Interface System structure, dependencies, and identified issues. The system follows a clear workflow pattern but requires refactoring to eliminate code duplication, fix field inconsistencies, and improve maintainability while preserving all existing functionality.