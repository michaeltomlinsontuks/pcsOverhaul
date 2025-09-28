# Subsystem 2: Number Generation - PCS Original System

## 🎯 **Subsystem Purpose**

The Number Generation subsystem provides **sequential number allocation** for the three core business entities in the PCS system: Enquiries (E-prefix), Quotes (Q-prefix), and Jobs (J-prefix). This subsystem ensures unique, sequential identification for all business records throughout the system lifecycle.

**Responsibility**: Sequential number generation, reservation, and confirmation for business entity identification.

---

## 📁 **Module Inventory**

### **1 Core Module**

| Module | Lines | Functions | Purpose | Dependencies |
|--------|-------|-----------|---------|--------------|
| `Calc_Numbers.bas` | 126 | 2 | E/Q/J number calculation and confirmation | Templates/ directory, Main.Main_MasterPath |

**Total**: 126 lines managing critical business numbering

---

## 🔢 **Core Numbering Architecture**

### **Numbering System Design**

#### **Business Entity Prefixes**
- **E-prefix**: Enquiries (E0001, E0002, E0003...)
- **Q-prefix**: Quotes (Q0001, Q0002, Q0003...)
- **J-prefix**: Jobs (J0001, J0002, J0003...)

#### **Number Tracking Method**
```
Templates/ Directory Structure:
├── E - 2003.TXT    # Next available enquiry number (2003)
├── Q - 1052.TXT    # Next available quote number (1052)
├── J - 0152.TXT    # Next available job number (0152)
└── Other template files...

Format: [Type] - [Number].TXT
```

**Storage Method**: Physical text files in Templates/ directory track the next available number for each type.

---

## 🔧 **Function Reference**

### **Calc_Numbers.bas** - Number Generation Engine

#### **Primary Functions**

##### **`Calc_Next_Number(Typ As String) As Long` - Calculate Next Available Number**

```vba
Public Function Calc_Next_Number(Typ As String)
Dim FullFilePath As String, MyName As String
Dim GroupCount As Integer

fileextension = "*.*"
path = "templates"

MyName = Dir(Main.Main_MasterPath.Value & path & "\", vbDirectory)
    If MyName = "" Then
        MsgBox "Folder Not Found", vbOKOnly, "Test"
            Exit Function
    End If

Do Until MyName = ""
    If MyName = "." Or MyName = ".." Then GoTo 2

        If Left(UCase(Typ), 1) = "E" And Left(MyName, 4) = "E - " Then
            Calc_Next_Number = Mid(MyName, InStr(1, MyName, "-", vbTextCompare) + 2, Len(MyName) - 8) + 1
            GoTo 8
        End If

        If Left(UCase(Typ), 1) = "J" And Left(MyName, 4) = "J - " Then
            Calc_Next_Number = Mid(MyName, InStr(1, MyName, "-", vbTextCompare) + 2, Len(MyName) - 8) + 1
            GoTo 8
        End If

        If Left(UCase(Typ), 1) = "Q" And Left(MyName, 4) = "Q - " Then
            Calc_Next_Number = Mid(MyName, InStr(1, MyName, "-", vbTextCompare) + 2, Len(MyName) - 8) + 1
            GoTo 8
        End If

        GroupCount = GroupCount + 1

2:
    MyName = Dir
Loop

8:
End Function
```

**Purpose**: Scans Templates/ directory to find current number tracker file and calculates next available number
**Parameters**:
- `Typ As String` - Entity type identifier ("E", "Q", or "J")
**Returns**: `Long` - Next available number for the specified type
**Algorithm**:
1. Scan Templates/ directory for files matching pattern "[Type] - ####.TXT"
2. Extract current number from filename
3. Add 1 to get next available number
4. Return calculated number (does NOT reserve it)

##### **`Confirm_Next_Number(Typ As String) As Long` - Reserve Number and Update Tracker**

```vba
Public Function Confirm_Next_Number(Typ As String)
Dim FullFilePath As String, MyName As String
Dim GroupCount As Integer

fileextension = "*.*"
path = "templates"

MyName = Dir(Main.Main_MasterPath.Value & path & "\", vbDirectory)
    If MyName = "" Then
        MsgBox "Folder Not Found", vbOKOnly, "Test"
            Exit Function
    End If

Do Until MyName = ""
    If MyName = "." Or MyName = ".." Then GoTo 2

        If Left(UCase(Typ), 1) = "E" And Left(MyName, 4) = "E - " Then
            Confirm_Next_Number = Mid(MyName, InStr(1, MyName, "-", vbTextCompare) + 2, Len(MyName) - 8) + 1

            FileCopy Main.Main_MasterPath & path & "\" & MyName, Main.Main_MasterPath & path & "\" & "E - " & Confirm_Next_Number & ".TXT"
            Kill Main.Main_MasterPath & path & "\" & MyName

            GoTo 8
        End If

        If Left(UCase(Typ), 1) = "J" And Left(MyName, 4) = "J - " Then
            Confirm_Next_Number = Mid(MyName, InStr(1, MyName, "-", vbTextCompare) + 2, Len(MyName) - 8) + 1

            FileCopy Main.Main_MasterPath & path & "\" & MyName, Main.Main_MasterPath & path & "\" & "J - " & Confirm_Next_Number & ".TXT"
            Kill Main.Main_MasterPath & path & "\" & MyName

            GoTo 8
        End If

        If Left(UCase(Typ), 1) = "Q" And Left(MyName, 4) = "Q - " Then
            Confirm_Next_Number = Mid(MyName, InStr(1, MyName, "-", vbTextCompare) + 2, Len(MyName) - 8) + 1

            FileCopy Main.Main_MasterPath & path & "\" & MyName, Main.Main_MasterPath & path & "\" & "Q - " & Confirm_Next_Number & ".TXT"
            Kill Main.Main_MasterPath & path & "\" & MyName

            GoTo 8
        End If

        GroupCount = GroupCount + 1

2:
    MyName = Dir
Loop

8:
End Function
```

**Purpose**: Reserves the next number by updating the tracker file in Templates/ directory
**Parameters**:
- `Typ As String` - Entity type identifier ("E", "Q", or "J")
**Returns**: `Long` - Confirmed/reserved number for the specified type
**Algorithm**:
1. Find current tracker file (e.g., "E - 1050.TXT")
2. Calculate next number (1051)
3. Create new tracker file with next number ("E - 1051.TXT")
4. Delete old tracker file
5. Return confirmed number

**Critical Operations**:
- **FileCopy**: Creates new tracker file
- **Kill**: Removes old tracker file
- **Atomic Update**: File operations ensure number reservation

---

## 🔄 **Number Generation Workflow**

### **Two-Phase Number Allocation**

#### **Phase 1: Number Calculation (Non-Committing)**
```vba
' Get next available number without reserving it
Dim nextEnquiry As Long
nextEnquiry = Calc_Next_Number("E")

' At this point, number is NOT reserved
' Other processes could potentially get the same number
```

#### **Phase 2: Number Confirmation (Committing)**
```vba
' Reserve the number and update tracker
Dim confirmedNumber As Long
confirmedNumber = Confirm_Next_Number("E")

' Number is now reserved and tracker file updated
' Subsequent calls will get the next number
```

### **Complete Usage Pattern in Forms**

#### **Standard Form Implementation**
```vba
' In FEnquiry.frm - SaveQ_Click() event
Private Sub SaveQ_Click()
    ' 1. Calculate next number (preview)
    Dim previewNumber As Long
    previewNumber = Calc_Next_Number("E")

    ' 2. Show number to user for confirmation
    Me.Enquiry_Number.Value = "E" & Format(previewNumber, "0000")

    ' 3. User confirms - now reserve the number
    Dim confirmedNumber As Long
    confirmedNumber = Confirm_Next_Number("E")

    ' 4. Use confirmed number for file operations
    Dim fileName As String
    fileName = "E" & Format(confirmedNumber, "0000") & ".xls"

    ' 5. Continue with file save operations...
End Sub
```

### **Number Sequencing Example**

#### **Initial State**
```
Templates/E - 1050.TXT    # Current enquiry tracker
```

#### **After Calc_Next_Number("E")**
```
Templates/E - 1050.TXT    # Unchanged - just calculated 1051
Return Value: 1051        # Next available number
```

#### **After Confirm_Next_Number("E")**
```
Templates/E - 1051.TXT    # New tracker file created
                          # Old E - 1050.TXT deleted
Return Value: 1051        # Confirmed reserved number
```

#### **Next Call to Calc_Next_Number("E")**
```
Templates/E - 1051.TXT    # Current tracker
Return Value: 1052        # Next available number
```

---

## 🗂️ **File System Dependencies**

### **Critical Directory Structure**

#### **Templates/ Directory Requirements**
```
Templates/
├── E - ####.TXT     # Enquiry number tracker (REQUIRED)
├── Q - ####.TXT     # Quote number tracker (REQUIRED)
├── J - ####.TXT     # Job number tracker (REQUIRED)
├── _Enq.xls         # Enquiry template
├── _Quote.xls       # Quote template
├── _Job.xls         # Job template
└── Other templates...
```

**Tracker File Format**: `[TYPE] - [NUMBER].TXT`
- **TYPE**: Single letter (E, Q, J)
- **NUMBER**: Current allocated number (4+ digits)
- **Extension**: Always .TXT

#### **File Content**
```
Tracker files are typically empty or contain minimal data
The filename itself contains the numbering information
```

### **Dependencies on Other Subsystems**

#### **Core Infrastructure Dependencies**
- **Main.Main_MasterPath.Value** - Base directory path from system initialization
- **File System Access** - Dir(), FileCopy, Kill operations
- **Directory Operations** - Templates/ directory must exist

#### **Usage by Business Logic Subsystems**
- **Enquiry Management** - Calls for E-prefix numbers
- **Quote Management** - Calls for Q-prefix numbers
- **Job Management** - Calls for J-prefix numbers

---

## ⚠️ **Error Handling and Limitations**

### **Error Conditions**

#### **Missing Templates Directory**
```vba
MyName = Dir(Main.Main_MasterPath.Value & path & "\", vbDirectory)
If MyName = "" Then
    MsgBox "Folder Not Found", vbOKOnly, "Test"
    Exit Function
End If
```

**Problem**: Function exits without returning value if Templates/ directory missing
**Impact**: Calling code receives uninitialized variable (0 or Empty)
**Solution**: Create Templates/ directory and initial tracker files

#### **Missing Tracker Files**
**Symptom**: Function returns 0 or Empty when no tracker file exists
**Cause**: No file matching "[Type] - ####.TXT" pattern in Templates/
**Resolution**: Manually create initial tracker files (e.g., "E - 1000.TXT")

#### **File System Permissions**
**Operations at Risk**:
- **Dir()** - Reading directory contents
- **FileCopy** - Creating new tracker file
- **Kill** - Deleting old tracker file

**Failure Mode**: Runtime errors if insufficient permissions

### **Race Condition Vulnerabilities**

#### **Multi-User Environment Issues**
```vba
' POTENTIAL RACE CONDITION:
' User A calls Calc_Next_Number("E") -> Gets 1051
' User B calls Calc_Next_Number("E") -> Gets 1051 (same number!)
' User A calls Confirm_Next_Number("E") -> Reserves 1051
' User B calls Confirm_Next_Number("E") -> May get error or skip numbers
```

**Current System Limitation**: No concurrency protection
**Mitigation**: Single-user system design assumption

#### **File Locking Issues**
- **FileCopy Operation**: May fail if file is locked by another process
- **Kill Operation**: May fail if file is in use
- **Directory Scanning**: May miss files during concurrent updates

### **Number Gap Scenarios**

#### **Intentional Gaps**
```vba
' User gets number 1051, but cancels operation
' Number 1051 is reserved but never used
' Next user gets 1052, creating gap at 1051
```

#### **Error-Induced Gaps**
- **FileCopy Succeeds, Kill Fails**: Both tracker files exist temporarily
- **Application Crash**: Between FileCopy and Kill operations
- **Network Issues**: File operations on network drives

---

## 🔧 **Algorithm Analysis**

### **String Parsing Logic**

#### **Filename Pattern Matching**
```vba
' Check for E-prefix files
If Left(UCase(Typ), 1) = "E" And Left(MyName, 4) = "E - " Then
    ' Extract number from filename
    Calc_Next_Number = Mid(MyName, InStr(1, MyName, "-", vbTextCompare) + 2, Len(MyName) - 8) + 1
End If
```

**Pattern Analysis**:
- `Left(UCase(Typ), 1) = "E"` - Checks first character of type parameter
- `Left(MyName, 4) = "E - "` - Verifies filename starts with "E - "
- `Mid(MyName, InStr(1, MyName, "-", vbTextCompare) + 2, Len(MyName) - 8)` - Extracts number portion

#### **Number Extraction Formula**
```
Filename: "E - 1050.TXT" (Length = 12)
Start Position: InStr(1, "E - 1050.TXT", "-") + 2 = 5
Extract Length: Len("E - 1050.TXT") - 8 = 4
Result: Mid("E - 1050.TXT", 5, 4) = "1050"
Add 1: 1050 + 1 = 1051
```

### **Directory Scanning Performance**

#### **Linear Search Algorithm**
```vba
Do Until MyName = ""
    ' Check each file against pattern
    If [pattern matching] Then
        ' Found match - extract number and exit
        GoTo 8
    End If
    MyName = Dir  ' Get next file
Loop
```

**Performance Characteristics**:
- **O(n)** complexity where n = number of files in Templates/
- **Best Case**: Target file is first in directory listing
- **Worst Case**: Target file is last or doesn't exist
- **Typical Case**: Fast for small Templates/ directory

### **File Operation Atomicity**

#### **Two-Step Update Process**
```vba
' Step 1: Create new tracker file
FileCopy Main.Main_MasterPath & path & "\" & MyName, _
         Main.Main_MasterPath & path & "\" & "E - " & Confirm_Next_Number & ".TXT"

' Step 2: Delete old tracker file
Kill Main.Main_MasterPath & path & "\" & MyName
```

**Atomicity Analysis**:
- **Not Atomic**: Two separate file operations
- **Failure Window**: Between FileCopy and Kill
- **Recovery**: Manual cleanup may be required

---

## 🎯 **Usage Patterns Throughout System**

### **Form Integration Examples**

#### **FEnquiry.frm - Enquiry Creation**
```vba
Private Sub SaveQ_Click()
    ' Calculate next enquiry number
    Dim nextNumber As Long
    nextNumber = Calc_Numbers.Calc_Next_Number("E")

    ' Display to user
    Me.Enquiry_Number.Value = "E" & Format(nextNumber, "0000")

    ' Confirm and reserve number
    Dim confirmedNumber As Long
    confirmedNumber = Calc_Numbers.Confirm_Next_Number("E")

    ' Use for file operations
    Dim fileName As String
    fileName = "E" & Format(confirmedNumber, "0000") & ".xls"
    ' ... continue with save operation
End Sub
```

#### **FQuote.frm - Quote Creation**
```vba
Private Sub SaveQuote_Click()
    ' Generate quote number
    Dim quoteNumber As Long
    quoteNumber = Calc_Numbers.Confirm_Next_Number("Q")

    ' Update form
    Me.Quote_Number.Value = "Q" & Format(quoteNumber, "0000")

    ' Move file from Enquiries to Quotes directory
    ' ... file operations using quote number
End Sub
```

#### **FAcceptQuote.frm - Job Creation**
```vba
Private Sub butSAVE_Click()
    ' Generate job number
    Dim jobNumber As Long
    jobNumber = Calc_Numbers.Confirm_Next_Number("J")

    ' Update form
    Me.Job_Number.Value = "J" & Format(jobNumber, "0000")

    ' Create job file in WIP directory
    ' ... file operations using job number
End Sub
```

### **Number Formatting Standards**

#### **Display Format Conventions**
```vba
' Standard 4-digit format with leading zeros
"E" & Format(number, "0000")  ' Results: E0001, E0010, E0100, E1000

' Alternative padding method
"E" & Right("0000" & number, 4)  ' Same result, different approach
```

#### **File Naming Conventions**
```vba
' Excel files use full formatted number
fileName = "E" & Format(enquiryNumber, "0000") & ".xls"  ' E1051.xls

' Tracker files use plain number
trackerName = "E - " & enquiryNumber & ".TXT"  ' E - 1051.TXT
```

---

## 🔄 **Integration with Other Subsystems**

### **Enquiry Management Integration**
```vba
' Enquiry form workflow
1. User opens FEnquiry.frm
2. Form calls Calc_Next_Number("E") for preview
3. User enters enquiry data
4. Form calls Confirm_Next_Number("E") on save
5. Enquiry saved with confirmed number
```

### **Quote Management Integration**
```vba
' Quote form workflow
1. User selects enquiry for quoting
2. FQuote.frm opens with enquiry data
3. Form calls Confirm_Next_Number("Q") for quote
4. Quote saved with new Q-number
5. Original enquiry file moved to quotes directory
```

### **Job Management Integration**
```vba
' Job acceptance workflow
1. User accepts submitted quote
2. FAcceptQuote.frm opens with quote data
3. Form calls Confirm_Next_Number("J") for job
4. Job created with new J-number
5. Quote archived, job moved to WIP
```

### **Search Database Integration**
```vba
' Search database updates
SaveSearchCode.SaveRowIntoSearch(frm)
' Updates Search.xls with:
' - New enquiry number (E####)
' - New quote number (Q####)
' - New job number (J####)
```

---

## 🛠️ **Maintenance and Troubleshooting**

### **Common Issues and Solutions**

#### **Missing Tracker Files**
**Symptoms**: Functions return 0, new numbers not generated
**Diagnosis**: Check Templates/ directory for [Type] - ####.TXT files
**Resolution**:
```vba
' Manually create missing tracker files
' Create file: Templates/E - 1000.TXT (starting enquiry number)
' Create file: Templates/Q - 1000.TXT (starting quote number)
' Create file: Templates/J - 1000.TXT (starting job number)
```

#### **Number Gaps**
**Symptoms**: Missing numbers in sequence (E1050, E1052 - missing E1051)
**Cause**: Cancelled operations or system errors
**Resolution**: Document gaps or backfill if necessary

#### **Duplicate Numbers**
**Symptoms**: Multiple files with same number
**Cause**: Race conditions or interrupted operations
**Resolution**: Renumber duplicates and update references

### **System Recovery Procedures**

#### **Tracker File Corruption**
```vba
' Find highest numbered file in each directory
1. Scan Enquiries/ directory for highest E#### file
2. Scan Quotes/ directory for highest Q#### file
3. Scan WIP/ and Archive/ directories for highest J#### file
4. Create new tracker files with next numbers
```

#### **Directory Scanning Issues**
```vba
' If Dir() function fails or returns unexpected results
1. Check Templates/ directory exists
2. Verify file permissions
3. Check for special characters in filenames
4. Restart Excel application if needed
```

---

## 🎯 **Development Best Practices**

### **Safe Number Generation Usage**

#### **Error Handling Pattern**
```vba
Public Function SafeGetNextNumber(entityType As String) As Long
    On Error GoTo ErrorHandler

    ' Validate input
    If entityType <> "E" And entityType <> "Q" And entityType <> "J" Then
        MsgBox "Invalid entity type: " & entityType
        SafeGetNextNumber = 0
        Exit Function
    End If

    ' Get next number
    SafeGetNextNumber = Calc_Numbers.Calc_Next_Number(entityType)

    ' Validate result
    If SafeGetNextNumber <= 0 Then
        MsgBox "Error generating number for type: " & entityType
        SafeGetNextNumber = 0
    End If

    Exit Function

ErrorHandler:
    MsgBox "Error in number generation: " & Err.Description
    SafeGetNextNumber = 0
End Function
```

#### **Validation Before Confirmation**
```vba
' Always validate before confirming numbers
Dim nextNumber As Long
nextNumber = Calc_Numbers.Calc_Next_Number("E")

If nextNumber > 0 Then
    ' Proceed with confirmation
    Dim confirmedNumber As Long
    confirmedNumber = Calc_Numbers.Confirm_Next_Number("E")

    If confirmedNumber = nextNumber Then
        ' Success - use the number
    Else
        ' Error - numbers don't match
        MsgBox "Number generation error"
    End If
Else
    MsgBox "Could not calculate next number"
End If
```

### **Testing Number Generation**

#### **Unit Test Functions**
```vba
Sub TestNumberGeneration()
    ' Test enquiry numbers
    Dim e1 As Long, e2 As Long
    e1 = Calc_Numbers.Calc_Next_Number("E")
    e2 = Calc_Numbers.Calc_Next_Number("E")

    If e1 = e2 Then
        MsgBox "Enquiry calculation test passed"
    Else
        MsgBox "Enquiry calculation test failed"
    End If

    ' Test confirmation
    Dim confirmed As Long
    confirmed = Calc_Numbers.Confirm_Next_Number("E")

    If confirmed = e1 Then
        MsgBox "Enquiry confirmation test passed"
    Else
        MsgBox "Enquiry confirmation test failed"
    End If
End Sub
```

---

## 🔍 **Next Steps**

After understanding Number Generation:

1. **Study [Enquiry Management](SUBSYSTEM_03_ENQUIRY_MANAGEMENT.md)** - See how E-numbers are used
2. **Review [Quote Management](SUBSYSTEM_04_QUOTE_MANAGEMENT.md)** - See how Q-numbers are generated
3. **Examine [Job Management](SUBSYSTEM_05_JOB_MANAGEMENT.md)** - See how J-numbers are created
4. **Practice Number Generation** - Create test functions with safe error handling
5. **Review Templates Directory** - Understand the file structure requirements

**Ready for enquiry management? Continue to [Enquiry Management Subsystem](SUBSYSTEM_03_ENQUIRY_MANAGEMENT.md)**