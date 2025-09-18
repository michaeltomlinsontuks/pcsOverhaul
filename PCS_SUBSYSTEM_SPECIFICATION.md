# PCS Interface Subsystem Specification

## Overview

This document maps every VBA file, function, and button in the PCS Interface system according to the defined subsystems based on the business workflow:
- **Enquiry → Quote → Jobs**
- **Jobs → Job Cards → WIP Reports**
- **Contracts (Job Templates)**
- **Search (Finds anything in the system)**

---

## SUBSYSTEM 1: ENQUIRY MANAGEMENT

### Core Files
- **FEnquiry.frm** - Main enquiry input form
- **FrmEnquiry.frm** - Alternative enquiry form (appears to be duplicate/legacy)

### Supporting Modules
- **Calc_Numbers.bas** - Generates sequential enquiry numbers (E-series)
- **SaveSearchCode.bas** - Saves enquiry data to search system
- **a_ListFiles.bas** - Lists enquiry files in directory

### Button Mappings
| Button/Action | Form | Function | Dependencies |
|---------------|------|----------|--------------|
| Add_Enquiry_Click | Main.frm | Opens FrmEnquiry form | Templates/_Enq.xls |
| AddMore_Click | FEnquiry.frm | Saves enquiry and creates new one | Calc_Numbers, SaveSearchCode |
| SaveQ_Click | FEnquiry.frm | Saves single enquiry | Calc_Numbers, SaveSearchCode |
| AddNewClient_Click | FEnquiry.frm | Creates new customer record | Templates/_client.xls |
| Dat_Click | FEnquiry.frm | Date picker functionality | ShowCalender function |

### Data Flow
1. **Input**: Customer details, component specifications, quantities
2. **Processing**: Assigns unique enquiry number (E-series)
3. **Storage**: Templates/_Enq.xls → Enquiries/ directory
4. **Integration**: Updates Search.xls for global search

### File Dependencies
- `Templates/_Enq.xls` - Enquiry template
- `Templates/_client.xls` - Customer template
- `Templates/price list.xls` - Component codes/descriptions
- `Templates/Component_Grades.xls` - Material grades
- `Customers/` directory - Customer files
- `Search.xls` - Global search database

---

## SUBSYSTEM 2: QUOTE MANAGEMENT

### Core Files
- **FQuote.frm** - Quote creation form

### Supporting Modules
- **Calc_Numbers.bas** - Generates sequential quote numbers (Q-series)
- **SaveSearchCode.bas** - Updates search system

### Button Mappings
| Button/Action | Form | Function | Dependencies |
|---------------|------|----------|--------------|
| Make_Quote_Click | Main.frm | Converts enquiry to quote | FQuote.frm |
| SaveQuote_Click | FQuote.frm | Finalizes quote | Calc_Numbers, Search.xls |

### Data Flow
1. **Input**: Enquiry file + pricing + lead time
2. **Processing**: Converts enquiry to quote with Q-series number
3. **Storage**: Enquiries/ → Quotes/ directory
4. **Integration**: Updates Search.xls

### File Dependencies
- Source: `Enquiries/` directory
- Target: `Quotes/` directory
- `Search.xls` - Global search database

---

## SUBSYSTEM 3: JOB MANAGEMENT

### Core Files
- **FJG.frm** - Job generation form (Jump the Gun)
- **FAcceptQuote.frm** - Quote acceptance and job creation

### Supporting Modules
- **Calc_Numbers.bas** - Generates job numbers (J-series)
- **SaveSearchCode.bas** - Updates search system
- **SaveWIPCode.bas** - Manages WIP database

### Button Mappings
| Button/Action | Form | Function | Dependencies |
|---------------|------|----------|--------------|
| createjob_Click | Main.frm | Accepts quote and creates job | FAcceptQuote.frm |
| butSAVE_Click | FAcceptQuote.frm | Finalizes job creation | Calc_Numbers, SaveWIP |
| butSaveJG_Click | FJG.frm | Direct job creation (bypass quote) | All job systems |
| ContractWork_Click | Main.frm | Creates job from contract template | FJG.frm, Contracts/ |
| JumpTheGun_Click | Main.frm | Emergency job creation | FJG.frm |

### Data Flow
1. **Input**: Accepted quote OR contract template OR direct job
2. **Processing**: Assigns J-series number, job card setup
3. **Storage**: Archive/ → WIP/ directory
4. **Integration**: Updates WIP.xls and Search.xls

### File Dependencies
- Source: `Quotes/` or `Contracts/` directory
- Target: `WIP/` directory
- `WIP.xls` - Work-in-progress database
- `Search.xls` - Global search database

---

## SUBSYSTEM 4: JOB CARD MANAGEMENT

### Core Files
- **FJobCard.frm** - Job card editing and operation planning

### Supporting Modules
- **SaveWIPCode.bas** - Updates WIP tracking
- **SaveSearchCode.bas** - Updates search system

### Button Mappings
| Button/Action | Form | Function | Dependencies |
|---------------|------|----------|--------------|
| OpenJob_Click | Main.frm | Opens job for editing | FJobCard.frm |
| SaveJobCard_Click | FJobCard.frm | Saves job card changes | WIP.xls, Search.xls |
| JobCardTemplates_Click | FJobCard.frm | Load operation templates | Job Templates/ |
| CopyFromJobCard_Click | FJobCard.frm | Copy operations from existing job | All directories |

### Data Flow
1. **Input**: Job operations, operators, dates, pictures
2. **Processing**: Operation planning and resource allocation
3. **Storage**: Updates files in WIP/ directory
4. **Integration**: Updates WIP.xls tracking

### File Dependencies
- Source/Target: `WIP/` directory
- `Job Templates/` - Operation templates
- `Operations.xls` - Available operation types
- `images/` - Job pictures/drawings
- `WIP.xls` - Work-in-progress database

---

## SUBSYSTEM 5: WIP REPORTS

### Core Files
- **fwip.frm** - WIP report generator
- **fwip_modified.frm** - Modified WIP reports

### Supporting Modules
- **SaveWIPCode.bas** - WIP data management

### Button Mappings
| Button/Action | Form | Function | Dependencies |
|---------------|------|----------|--------------|
| WIPReport_Click | Main.frm | Opens WIP report options | fwip.frm |
| Go_Click | fwip.frm | Generates selected reports | WIP.xls |
| OpenWIP_Click | Main.frm | Opens WIP database directly | WIP.xls |
| CloseJob_Click | Main.frm | Closes completed job | Multiple files |

### Report Types Generated
- **Operation Reports** - Jobs by operation type
- **Operator Reports** - Jobs by assigned operator
- **Due Date Reports** - Jobs by delivery date
- **Customer Reports** - Jobs by customer (Office & Workshop views)
- **Job Number Reports** - Jobs by number sequence

### Data Flow
1. **Input**: WIP.xls database
2. **Processing**: Sorts and filters by various criteria
3. **Output**: Formatted Excel reports in Templates/ directory

### File Dependencies
- `WIP.xls` - Primary data source
- Output: `Templates/` directory (various .xls files)

---

## SUBSYSTEM 6: CONTRACT MANAGEMENT

### Core Files
- **FJG.frm** - Used for contract job creation
- **FList.frm** - Contract selection list

### Supporting Modules
- **a_ListFiles.bas** - Lists available contracts

### Button Mappings
| Button/Action | Form | Function | Dependencies |
|---------------|------|----------|--------------|
| but_CreateCTItem_Click | Main.frm | Creates new contract template | FJG.frm |
| but_EditCTItem_Click | Main.frm | Edits existing contract | FList.frm, FJG.frm |
| but_SaveAsCTItem_Click | FJG.frm | Saves job as contract template | Contracts/ |

### Data Flow
1. **Creation**: Job specifications saved as reusable templates
2. **Storage**: `Contracts/` directory
3. **Usage**: Templates used to create new jobs via ContractWork_Click

### File Dependencies
- `Contracts/` directory - Contract templates
- `Templates/_Enq.xls` - Base template for contracts

---

## SUBSYSTEM 7: SEARCH SYSTEM

### Core Files
- **Search.xls** (external file, managed by interface)

### Supporting Modules
- **SaveSearchCode.bas** - Core search data management
- **Module1.bas** - Search database updates
- **Search_Sync.bas** - Search synchronization

### Button Mappings
| Button/Action | Form | Function | Dependencies |
|---------------|------|----------|--------------|
| Search_Click | Main.frm | Opens search interface | Search.xls |
| butEditSearch_Click | Main.frm | Edit search database | Search.xls |
| butSearchHistory_Click | Main.frm | Opens search history | Search History.xls |
| butJobHistory_Click | Main.frm | Opens job history | Job History.xls |
| butQuoteHistory_Click | Main.frm | Opens quote history | Quote History.xls |
| butSortSearch_Click | Main.frm | Sorts search database | Search.xls |

### Data Flow
1. **Input**: All subsystems automatically feed data
2. **Processing**: Maintains searchable database of all records
3. **Output**: Global search across all enquiries, quotes, jobs

### File Dependencies
- `Search.xls` - Main search database
- `Search History.xls` - Search history
- `Job History.xls` - Job search history
- `Quote History.xls` - Quote search history

---

## SUBSYSTEM 8: MAIN INTERFACE & NAVIGATION

### Core Files
- **Main.frm** - Primary navigation interface
- **a_Main.bas** - Main interface launcher

### Supporting Modules
- **RefreshMain.bas** - Interface refresh functionality
- **Check_Updates.bas** - Real-time file monitoring
- **a_ListFiles.bas** - File listing functionality

### Button Mappings
| Button/Action | Form | Function | Dependencies |
|---------------|------|----------|--------------|
| Enquiries_Click | Main.frm | Lists enquiry files | List_Files("Enquiries") |
| Quotes_Click | Main.frm | Lists quote files | List_Files("quotes") |
| WIP_Click | Main.frm | Lists WIP files | List_Files("WIP") |
| Archive_Click | Main.frm | Lists archived files | List_Files("Archive") |
| JobsInWIP_Click | Main.frm | Lists jobs in WIP database | WIP.xls |
| Thirties_Click | Main.frm | Lists jobs 30000+ series | Search.xls |
| lst_Click | Main.frm | Displays selected file details | GetValue function |
| Lst_DblClick | Main.frm | Opens selected file | OpenBook function |

### Data Flow
- **Central Hub**: All subsystems accessible from Main interface
- **Real-time Updates**: Monitors file changes across all directories
- **File Management**: Provides access to all system files

---

## UTILITY MODULES

### Core Utility Functions

#### **Calc_Numbers.bas**
- `Calc_Next_Number()` - Generates sequential numbers for E, Q, J series
- `Confirm_Next_Number()` - Commits number assignment

#### **Check_Updates.bas**
- `CheckUpdates()` - Real-time monitoring of directory changes
- `Check_Files()` - Counts files in directories
- `StopCheck()` - Stops monitoring

#### **SaveSearchCode.bas**
- `SaveRowIntoSearch()` - Updates search database

#### **SaveWIPCode.bas**
- `SaveInfoIntoWIP()` - Updates WIP database

#### **a_ListFiles.bas**
- `List_Files()` - Populates listboxes with directory contents
- `GetValue()` - Retrieves data from closed workbooks

#### **RefreshMain.bas**
- `Refresh_Main()` - Refreshes main interface display

#### **Module1.bas**
- `Update_Search()` - Bulk search database updates
- `GetValue()` - External file data retrieval

#### **Search_Sync.bas**
- `Search_Sync()` - Synchronizes search data

### Supporting Utilities

#### **Open_Book.bas**
- `OpenBook()` - Standardized file opening

#### **GetUserNameEx.bas** / **GetUserName64.bas**
- User identification functions (32/64 bit compatibility)

#### **RemoveCharacters.bas**
- `Remove_Characters()` - String cleaning for filenames

#### **Very_HiddenSheet.bas**
- `ShowSheet()` - Sheet visibility management
- `DeleteSheet()` - Sheet removal

#### **SaveFileCode.bas**
- Generic file saving functionality

#### **GetValue.bas**
- External workbook data retrieval

#### **Delete_Sheet.bas**
- Sheet management utilities

#### **Check_Dir.bas**
- Directory validation

---

## UNUSED/LEGACY CODE IDENTIFICATION

### Potentially Unused Files
- **FrmEnquiry.frm** - Appears to duplicate FEnquiry.frm functionality
- **fwip_modified.frm** - Modified version, unclear if active
- **Module2.bas** / **Module3.bas** - Generic module names, minimal code

### Dead Code Analysis Required
- Multiple `GetValue()` function definitions across files
- Duplicate form handling in Main.frm
- Unused variables and functions in fwip.frm
- Legacy error handling code

---

## CRITICAL DEPENDENCIES

### File System Structure
```
Root/
├── Enquiries/          # E-series files
├── Quotes/             # Q-series files
├── WIP/                # J-series active jobs
├── Archive/            # Completed jobs
├── Contracts/          # Job templates
├── Customers/          # Customer database
├── Templates/          # System templates
├── Job Templates/      # Operation templates
├── images/             # Job pictures/drawings
├── Search.xls          # Global search database
├── WIP.xls            # Work-in-progress tracking
└── Operations.xls      # Available operations
```

### Inter-Subsystem Dependencies
1. **All subsystems** → Search System (for tracking)
2. **Enquiry** → Quote → Job → Job Card (main workflow)
3. **Job Card** → WIP Reports (for monitoring)
4. **Contracts** → Job Creation (for templates)
5. **All subsystems** → Main Interface (for navigation)

### External File Dependencies
- Excel workbooks in all directories
- Template files for data structures
- Image files for job documentation
- History files for audit trails

---

## REFACTORING RECOMMENDATIONS

### Immediate Improvements Needed
1. **Consolidate GetValue() functions** - Single utility module
2. **Remove duplicate forms** - Standardize on single enquiry form
3. **Centralize search updates** - Single SaveSearch module
4. **Standardize error handling** - Common error management
5. **Clean up unused variables** - Remove dead code

### Modular Structure for New Code
1. **Core Data Models** - Enquiry, Quote, Job classes
2. **File Management Service** - Centralized file operations
3. **Search Service** - Unified search functionality
4. **Number Generation Service** - Centralized ID management
5. **UI Controllers** - Separate business logic from forms

This specification provides the complete mapping of the current PCS Interface system and serves as the foundation for the planned refactoring while maintaining all existing functionality.