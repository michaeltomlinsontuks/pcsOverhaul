# PCS - Excel VBA Business Management System

## System Overview

This folder contains an integrated business management system implemented in Microsoft Excel using VBA (Visual Basic for Applications). The system is designed to manage various aspects of business operations including enquiries, quotes, jobs, and work-in-progress tracking.

## Directory Structure

- **Main Interface Files**
  - `_Interface.xls` - Main entry point and navigation hub
  - `Search.xls` - Comprehensive search functionality
  - `Operation.xls` / `Operations.xls` - Operational workflows
  - `Wip.xls` - Work in Progress item management

- **Data Organization**
  - `/Customers/` - Individual Excel files for customer data
  - `/Contracts/` - Contract files stored as separate Excel files
  - `/Enquiries/` - Enquiry records stored as individual files
  - `/Job Templates/` - Templates for different job types
  - `/Templates/` - System templates for various functions
  - `/WIP/` - Work in Progress tracking files
  - `/Quotes/` - Quote management (directory structure)
  - `/VBA/` - Extracted VBA code from all Excel files

## Functional Components

Based on the folder structure and files, the system handles:

1. **Customer Management**:
   - Individual customer files stored in the `Customers` directory
   - Each customer has their own Excel file (e.g., "ACRON ENGINEERING.xls")

2. **Contract Management**:
   - Contracts stored in the `Contracts` directory
   - Files are named with contract references or descriptions

3. **Job Processing System**:
   - `Job Templates` directory contains standardized job templates
   - Various template files for different workflows for job processing
   - Job creation from quotes via the main interface

4. **Work-in-Progress Tracking**:
   - `Wip.xls` and the `WIP` directory track ongoing work
   - Reports for Work-in-Progress available through the interface

5. **Quoting System**:
   - The `Quotes` directory manages quotes
   - Functionality to convert enquiries to quotes
   - Quote acceptance process leading to job creation

6. **Enquiry Management**:
   - The `Enquiries` directory contains individual enquiry files
   - System to create and track customer enquiries

7. **Search Functionality**:
   - The `Search.xls` file provides comprehensive search capabilities
   - Ability to search across the entire system's data

## System Architecture

The system operates as a modular Excel VBA application where:

1. Users start with the `_Interface.xls` file as the main dashboard
2. From there, they navigate to different functional areas (customer management, job management, etc.)
3. Templates are used to create standardized documents and records
4. Customer and contract data is stored in individual files for organization
5. Search functionality allows finding information across the system

## VBA Code Structure

The core VBA code has been extracted to the `/VBA/` directory using the oledump.py tool. The main modules include:

- `_Interface.xls.vba` (176KB) - The core application with all the main modules and forms
- Functional modules for search, WIP tracking, and enquiry management
- Templates for job handling, office processes, and workshop operations
- Individual contract and work item files with their specific VBA code

## Key Features (From VBA Code Analysis)

- **User Interface**: Forms-based navigation system with a main menu
- **File Management**: Directory-based organization of data files
- **Workflow Management**: Processes for handling enquiries → quotes → jobs → WIP
- **Reporting**: Various reporting functions including WIP reports
- **Search Functionality**: Comprehensive search across all system data
- **Real-time Updates**: Periodic checks for file system changes

## Tools for Code Analysis

VBA code was extracted using oledump.py, a powerful tool for analyzing OLE files including Excel documents:

1. **Installation**: 
   ```
   mkdir -p /Users/athollt/tools/oledump
   cd /Users/athollt/tools/oledump
   curl -L -o oledump.zip https://didierstevens.com/files/software/oledump_V0_0_76.zip
   unzip -o oledump.zip
   rm oledump.zip
   python3 -m venv venv
   source venv/bin/activate
   pip install olefile yara-python
   ```

2. **Extracting VBA Code**:
   ```
   /path/to/oledump/run_oledump.sh -s [MODULE_NUMBER] -v [EXCEL_FILE_PATH]
   ```

3. **Listing All Modules**:
   ```
   /path/to/oledump/run_oledump.sh [EXCEL_FILE_PATH]
   ```

## Usage Notes

The system appears designed to handle a complete business workflow from initial customer enquiry through to job completion:

1. Customer enquiries are captured and stored
2. Enquiries can be converted to quotes
3. Quotes can be accepted and converted to jobs
4. Jobs are tracked through production/delivery
5. Search functionality helps locate information across the system

---

*Note: This README was created based on analysis of the file and directory structure, as well as examination of the VBA code extracted from the Excel files.*
