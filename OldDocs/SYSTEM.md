# PCS System Architecture Documentation

## System Overview

The PCS system is a Microsoft Excel VBA-based business management application used for tracking customer enquiries, creating quotes, managing jobs and contracts, and tracking work-in-progress (WIP). The system follows a modular approach with interconnected Excel files that share data and functionality.

## Technical Architecture

### Core Components

1. **Main Interface (_Interface.xls)**
   - Central navigation hub containing 28 VBA modules
   - Provides access to all system functionality through custom forms
   - Manages application initialization and system-wide settings
   - Contains core business logic for the main workflows

2. **Data Storage**
   - Uses file-based storage rather than a database
   - Individual Excel files store customer, quote, job, and WIP data
   - Directory structure provides organization for different data types
   - Data flows between files through VBA code rather than linked cells

3. **Search Engine (Search.xls)**
   - Custom implementation of search functionality
   - Indexes key data across all system files
   - Provides cross-system data access for reporting and lookups

4. **Template System**
   - Standardized templates for creating new records
   - Templates enforce data structure and business rules
   - Separate templates for workshop, office, and customer-facing documents

## Core Modules & Functionality

### User Interface Components

1. **Main Form (Main module)**
   - Primary navigation interface
   - Displays lists of enquiries, quotes, and jobs
   - Allows filtering by status and type
   - Contains buttons for primary actions (create quote, create job, search)

2. **Form Modules**
   - **FEnquiry/FrmEnquiry**: Captures customer enquiry details
   - **FQuote**: Creates and manages quotes from enquiries
   - **FJobCard/FJG**: Handles job creation and management
   - **FAcceptQuote**: Processes quote acceptance into jobs
   - **FWIP/fwip**: Manages work-in-progress tracking
   - **FList**: Displays lists of records for selection

### Core Business Logic

1. **Workflow Management**
   - **Check_Updates**: Background process that monitors file system changes
   - **RefreshMain**: Updates the main interface when data changes
   - **Calc_Numbers**: Generates sequential numbers for new records
   - **a_ListFiles/a_Main**: Core file management utilities
   - **Search_Sync**: Synchronizes search functionality

2. **Data Management**
   - **SaveFileCode**: Handles saving files with appropriate naming
   - **SaveSearchCode**: Updates search indices when records change
   - **SaveWIPCode**: Specialized saving for WIP records
   - **GetValue**: Retrieves values from closed workbooks (key for performance)
   - **Check_Dir**: Directory verification and creation

3. **Utility Functions**
   - **Open_Book**: Standardized file opening with error handling
   - **Delete_Sheet**: Removes sheets without confirmation prompts
   - **RemoveCharacters**: Text manipulation for consistent formatting
   - **GetUserNameEx**: User identification for audit trails

## System Workflows

### 1. Enquiry Management
```
Customer Enquiry → _Enq.xls template → 
Saved as unique number (e.g., 384776.xls) → 
Listed in Main interface
```

### 2. Quote Generation
```
Select enquiry → FQuote form → 
Price calculation → Quote acceptance process → 
Saved to Quotes directory
```

### 3. Job Creation
```
Accepted quote → FAcceptQuote form → 
Job number assignment (via Calc_Numbers) → 
Job card creation (FJobCard) → 
Saved to WIP directory
```

### 4. Work Tracking
```
Active job → WIP tracking → 
Status updates via fwip module → 
Completion and archiving
```

## Technical Implementation Details

### Data Flow Architecture

The system uses a file-centric approach where:
- Each major record type has its own Excel file
- Files are organized in directories by type
- VBA code reads/writes between files rather than using cell links
- The `GetValue()` function is critical for accessing data across files

### State Management

- System state is primarily file-based
- User interface is refreshed via polling (`CheckUpdates`)
- File timestamps are used to detect changes
- The main interface displays notifications of changes

### Search Implementation

- Custom implementation rather than Excel's built-in search
- Creates and maintains indices of key fields
- Full-text search across multiple record types
- Results display in a unified interface

### Security Model

- Based on file system security
- No explicit user permission model in the VBA code
- Relies on Excel's workbook protection for formula integrity
- No evidence of encryption for sensitive data

## Technical Limitations

1. **Performance Constraints**
   - File-based architecture limits scalability
   - Search performance likely degrades with volume
   - VBA's single-threaded nature limits concurrent operations

2. **Maintenance Challenges**
   - Distributed business logic across multiple files
   - Limited modularization in some core components
   - Tight coupling between user interface and business logic

3. **Error Handling**
   - Basic error handling with message boxes
   - Limited logging of system errors
   - Some error suppression without diagnostic information

## System Dependencies

- Microsoft Excel (likely 2003-2010 era based on the .xls format)
- Windows OS (contains Windows API calls for file management)
- Local file system access (no network database connectivity)
- No external dependencies beyond Excel

---

This document represents a technical analysis of the PCS system based on examination of the VBA code extracted from the Excel files.
