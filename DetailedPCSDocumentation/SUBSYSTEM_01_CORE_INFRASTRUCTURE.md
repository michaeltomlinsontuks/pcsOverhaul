# Subsystem 1: Core Infrastructure - PCS Original System

## 🎯 **Subsystem Purpose**

The Core Infrastructure subsystem provides the **foundational building blocks** for the entire PCS Interface System. This subsystem handles system initialization, file operations, API functions, and essential utilities that all other subsystems depend on.

**Responsibility**: Foundation layer for system entry, file management, directory operations, and platform compatibility.

---

## 📁 **Module Inventory**

### **9 Core Infrastructure Modules**

| Module | Lines | Purpose | Dependencies |
|--------|-------|---------|--------------|
| `a_Main.bas` | 17 | System entry point | Main.frm |
| `Open_Book.bas` | 10 | Workbook management | Excel Application |
| `Check_Dir.bas` | 16 | Directory operations | File System |
| `GetUserNameEx.bas` | 14 | 32-bit Windows API | advapi32.dll |
| `GetUserName64.bas` | 16 | 64-bit Windows API | advapi32.dll |
| `GetValue.bas` | 25 | Closed workbook access | Excel ExecuteExcel4Macro |
| `Very_HiddenSheet.bas` | 22 | Worksheet visibility | Excel Worksheets |
| `Delete_Sheet.bas` | 12 | Worksheet deletion | Excel Worksheets |
| `RemoveCharacters.bas` | 35 | String utilities | None |

**Total**: 167 lines of core infrastructure code

---

## 🚀 **System Entry Point**

### **a_Main.bas** - Application Initialization

#### **Primary Functions**

##### **`ShowMenu()` - System Entry Point**
```vba
Sub ShowMenu()
    Main.Main_MasterPath.Value = ActiveWorkbook.path & "\"
    Main.Show
End Sub
```

**Purpose**: Primary application entry point that initializes the system
**Parameters**: None
**Returns**: None (displays main interface)
**Side Effects**:
- Sets global master path for file operations
- Displays Main.frm interface
- Initializes system for user interaction

**Dependencies**:
- Main.frm must exist and be accessible
- ActiveWorkbook must be the PCS system workbook

##### **`sadf()` - Development/Testing Function**
```vba
Sub sadf()
    Do
        ActiveCell.Value = ActiveCell.Offset(-1, 0).Value - 1
        ActiveCell.Offset(1, 0).Select
    Loop Until ActiveCell.Offset(-1, 0).Value = 1011
End Sub
```

**Purpose**: Appears to be development/testing code for cell value manipulation
**Parameters**: None
**Returns**: None
**Note**: This function appears to be legacy development code and may not be used in production

#### **Usage Pattern**
```vba
' System startup sequence
1. User opens _Interface.xls
2. Macro execution triggers a_Main.ShowMenu()
3. Master path set to workbook directory
4. Main interface displayed to user
```

---

## 📂 **File Operations**

### **Open_Book.bas** - Workbook Management

#### **Primary Function**

##### **`OpenBook(File As String, RO As Boolean)` - Workbook Opening**
```vba
Public Function OpenBook(File As String, RO As Boolean)
    Workbooks.Open Filename:=File, ReadOnly:=RO
End Function
```

**Purpose**: Opens Excel workbooks with read-only option
**Parameters**:
- `File As String` - Full path to workbook file
- `RO As Boolean` - True for read-only, False for read-write
**Returns**: None (opens workbook in Excel application)
**Usage**: Standard file opening throughout system

**Critical Dependencies**:
- File path must be valid and accessible
- Excel application must be available
- User must have file system permissions

#### **Usage Examples**
```vba
' Open for editing
Call OpenBook("C:\PCS\Enquiries\E1001.xls", False)

' Open for read-only access
Call OpenBook("C:\PCS\Templates\_Enq.xls", True)

' Used throughout system for:
' - Template access
' - Data file editing
' - Search database updates
' - WIP database management
```

### **GetValue.bas** - Closed Workbook Data Access

#### **Primary Function**

##### **`GetValue(path, File, sheet, ref)` - Read From Closed Files**
```vba
Public Function GetValue(path, File, sheet, ref) As Variant
    Dim arg As String

    ' Check if already open
    If file_opened(File) = True Then
        GetValue = ActiveWorkbook.Worksheets(sheet).Range(ref).Value
        Exit Function
    End If

    ' Build Excel 4 macro formula for closed file access
    If Right(path, 1) <> "\" Then
        path = path & "\"
    End If

    arg = "'" & path & "[" & File & "]" & sheet & "'!" & Range(ref).Address(, , xlR1C1)
    GetValue = ExecuteExcel4Macro(arg)
End Function
```

**Purpose**: Retrieves specific cell values from closed Excel workbooks
**Parameters**:
- `path` - Directory path (string)
- `File` - Workbook filename (string)
- `sheet` - Worksheet name (string)
- `ref` - Cell reference (string)
**Returns**: `Variant` - Cell value from closed workbook
**Dependencies**: Excel ExecuteExcel4Macro function

#### **Supporting Function**

##### **`file_opened(File_name As String)` - Check If Workbook Open**
```vba
Private Function file_opened(File_name As String) As Boolean
    Dim i As Integer
    file_opened = False

    For i = 1 To Workbooks.Count
        If Workbooks(i).Name = File_name Then
            file_opened = True
            Exit For
        End If
    Next i
End Function
```

**Purpose**: Checks if workbook is already open in Excel
**Parameters**: `File_name As String` - Workbook filename
**Returns**: `Boolean` - True if open, False if closed

#### **Usage Throughout System**
```vba
' Load customer data
customerName = GetValue(masterPath & "Customers\", "CustomerList.xls", "Data", "A1")

' Read template values
templateValue = GetValue(masterPath & "Templates\", "_Enq.xls", "Admin", "B5")

' Access search database
searchRecord = GetValue(masterPath, "Search.xls", "SearchData", "C10")
```

---

## 📁 **Directory Management**

### **Check_Dir.bas** - Directory Operations

#### **Primary Function**

##### **`CheckDir(Direc As String)` - Directory Creation and Navigation**
```vba
Public Function CheckDir(Direc As String)
    If Dir(Direc, vbDirectory) = "" Then
        MkDir (Direc)
        ChDir (Direc)
    Else
        ChDir (Direc)
    End If
End Function
```

**Purpose**: Creates directory if it doesn't exist, then changes to that directory
**Parameters**: `Direc As String` - Directory path to create/navigate to
**Returns**: None
**Side Effects**:
- Creates directory if missing
- Changes current working directory
- May throw error if permissions insufficient

#### **Usage Pattern Throughout System**
```vba
' Ensure Enquiries directory exists before saving
Call CheckDir(Main.Main_MasterPath.Value & "Enquiries\")

' Create customer directory structure
Call CheckDir(Main.Main_MasterPath.Value & "Customers\")

' Prepare WIP directory for job files
Call CheckDir(Main.Main_MasterPath.Value & "WIP\")
```

**Critical for System Operation**:
- All file save operations depend on directory existence
- System fails gracefully if directories are missing
- Maintains 20081222/ directory structure integrity

---

## 🖥️ **Platform Compatibility**

### **32/64-bit Windows API Functions**

#### **GetUserNameEx.bas** - 32-bit Windows API

##### **API Declaration**
```vba
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
                                      (ByVal lpBuffer As String, _
                                       nSize As Long) As Long
```

##### **`Get_User_Name()` - Retrieve Windows Username (32-bit)**
```vba
Public Function Get_User_Name()
    Dim lpBuff As String * 25
    Dim ret As Long, UserName As String
    ret = GetUserName(lpBuff, 25)
    Get_User_Name = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)
End Function
```

**Purpose**: Gets current Windows username using 32-bit API
**Parameters**: None
**Returns**: `String` - Current Windows username
**Dependencies**: advapi32.dll (Windows system library)
**Architecture**: 32-bit Excel only

#### **GetUserName64.bas** - 64-bit Windows API

##### **API Declaration**
```vba
Private Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
                                              (ByVal lpBuffer As String, _
                                               nSize As LongPtr) As Long
```

##### **`Get_User_Name()` - Retrieve Windows Username (64-bit)**
```vba
Public Function Get_User_Name()
    Dim lpBuff As String * 25
    Dim ret As Long, UserName As String
    ret = GetUserName(lpBuff, 25)
    Get_User_Name = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)
End Function
```

**Purpose**: Gets current Windows username using 64-bit API
**Parameters**: None
**Returns**: `String` - Current Windows username
**Dependencies**: advapi32.dll (Windows system library)
**Architecture**: 64-bit Excel only
**Key Differences**: Uses `PtrSafe` and `LongPtr` for 64-bit compatibility

#### **Universal Compatibility Pattern**
```vba
' V2 System consolidates both approaches:
#If VBA7 Then
    Private Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
        (ByVal lpBuffer As String, nSize As LongPtr) As Long
#Else
    Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
        (ByVal lpBuffer As String, nSize As Long) As Long
#End If
```

---

## 🛠️ **Worksheet Management**

### **Very_HiddenSheet.bas** - Worksheet Visibility Control

#### **Primary Functions**

##### **`VeryHiddenSheet(SheetNam As String)` - Hide Worksheet Completely**
```vba
Public Function VeryHiddenSheet(SheetNam As String)
    Worksheets(SheetNam).Visible = xlSheetVeryHidden
End Function
```

**Purpose**: Hides worksheet completely (not visible in sheet tabs or menus)
**Parameters**: `SheetNam As String` - Name of worksheet to hide
**Returns**: None
**Usage**: Protecting system templates and sensitive data sheets

##### **`ShowSheet(SheetNam As String)` - Make Worksheet Visible**
```vba
Public Function ShowSheet(SheetNam As String)
    Worksheets(SheetNam).Visible = xlSheetVisible
End Function
```

**Purpose**: Makes hidden worksheet visible and accessible
**Parameters**: `SheetNam As String` - Name of worksheet to show
**Returns**: None
**Usage**: Revealing sheets for user interaction

#### **Visibility States in Excel**
- `xlSheetVisible` - Normal visibility (default)
- `xlSheetHidden` - Hidden from tabs but accessible via code
- `xlSheetVeryHidden` - Completely hidden, only accessible via VBA

### **Delete_Sheet.bas** - Worksheet Deletion

#### **Primary Function**

##### **`DeleteSheet(SheetName As String)` - Remove Worksheet**
```vba
Public Function DeleteSheet(SheetName As String)
    Application.DisplayAlerts = False
    Worksheets(SheetName).Delete
    Application.DisplayAlerts = True
End Function
```

**Purpose**: Deletes worksheet without user confirmation prompts
**Parameters**: `SheetName As String` - Name of worksheet to delete
**Returns**: None
**Side Effects**:
- Permanently removes worksheet
- Suppresses Excel confirmation dialogs
- Restores alert settings

**Critical Safety Note**: No undo available - deletion is permanent

---

## 🔤 **String Utilities**

### **RemoveCharacters.bas** - String Processing

#### **Primary Functions**

##### **`Remove_Characters(Str As String)` - Clean Strings for Filenames**
```vba
Public Function Remove_Characters(Str As String) As String
    Str = Replace(Str, "/", "")
    Str = Replace(Str, ":", "")
    Str = Replace(Str, " ", "")
    Remove_Characters = Str
End Function
```

**Purpose**: Removes problematic characters for file naming
**Parameters**: `Str As String` - Input string to clean
**Returns**: `String` - Cleaned string safe for filenames
**Removes**: Forward slash (/), colon (:), space ( )
**Usage**: Preparing customer names and descriptions for file paths

##### **`Insert_Characters(Str As String)` - Format for Display**
```vba
Public Function Insert_Characters(Str As String) As String
    Str = Replace(Str, "REF", "REF: ")
    Str = Replace(Str, "QTY", "QTY: ")
    Insert_Characters = Str
End Function
```

**Purpose**: Adds formatting spaces for better display readability
**Parameters**: `Str As String` - Input string to format
**Returns**: `String` - Formatted string for display
**Usage**: Improving form display and report formatting

#### **Usage Examples**
```vba
' Clean customer name for file saving
Dim cleanName As String
cleanName = Remove_Characters("ABC/Company: Ltd.")
' Result: "ABCCompanyLtd."

' Format for display
Dim displayText As String
displayText = Insert_Characters("REF123QTY50")
' Result: "REF: 123QTY: 50"
```

---

## 🔗 **Subsystem Dependencies**

### **Core Infrastructure Provides Services To:**

#### **All Other Subsystems Depend On:**
- **System Entry**: `a_Main.ShowMenu()` - Application startup
- **File Operations**: `OpenBook()`, `GetValue()` - File access throughout system
- **Directory Management**: `CheckDir()` - Directory preparation for all file operations
- **User Identification**: `Get_User_Name()` - User tracking and logging
- **String Processing**: `Remove_Characters()`, `Insert_Characters()` - Data formatting

#### **Specific Subsystem Dependencies:**

##### **Number Generation Subsystem**
```vba
' Depends on directory operations
Call CheckDir(Main.Main_MasterPath.Value & "Templates\")

' Depends on file operations
Call OpenBook(templateFile, False)
```

##### **Enquiry Management Subsystem**
```vba
' Depends on user identification
currentUser = Get_User_Name()

' Depends on string utilities
cleanCustomerName = Remove_Characters(customerInput)

' Depends on file access
customerData = GetValue(customerPath, customerFile, "Data", "A1")
```

##### **All Form Subsystems**
```vba
' Depend on master path from a_Main initialization
filePath = Main.Main_MasterPath.Value & "Enquiries\" & enquiryNumber & ".xls"
```

### **External Dependencies:**

#### **Windows System Dependencies**
- **advapi32.dll** - Windows API library for user functions
- **File System** - Directory and file operations
- **Excel Application** - Workbook and worksheet management

#### **Excel Dependencies**
- **ExecuteExcel4Macro** - Closed workbook access
- **Workbooks Collection** - Workbook management
- **Worksheets Collection** - Worksheet operations
- **Application Object** - Excel application control

---

## ⚠️ **Error Handling and Limitations**

### **Error Handling Patterns**

#### **Minimal Error Handling in Original System**
```vba
' Most functions lack comprehensive error handling
Public Function OpenBook(File As String, RO As Boolean)
    Workbooks.Open Filename:=File, ReadOnly:=RO
    ' No error handling - will crash if file doesn't exist
End Function
```

#### **Basic Error Handling Where Present**
```vba
' GetValue includes basic error checking
Private Function file_opened(File_name As String) As Boolean
    ' Checks for open workbooks to avoid errors
    For i = 1 To Workbooks.Count
        If Workbooks(i).Name = File_name Then
            file_opened = True
            Exit For
        End If
    Next i
End Function
```

### **Common Failure Points**

#### **File System Errors**
- **Missing Files**: `OpenBook()` crashes if file doesn't exist
- **Permission Errors**: `CheckDir()` fails if insufficient rights
- **Path Issues**: Invalid paths cause runtime errors

#### **API Function Errors**
- **Architecture Mismatch**: Wrong API version causes compile errors
- **DLL Missing**: advapi32.dll issues on non-Windows systems
- **Buffer Overflow**: Fixed-size string buffers may truncate

#### **Excel Object Errors**
- **Worksheet Missing**: Direct worksheet access without validation
- **Application State**: Errors when Excel is in protected mode
- **Memory Issues**: Large file operations without cleanup

### **Defensive Programming Recommendations**

#### **Enhanced Error Handling Pattern**
```vba
Public Function SafeOpenBook(File As String, RO As Boolean) As Boolean
    On Error GoTo ErrorHandler

    ' Check file exists first
    If Dir(File) = "" Then
        MsgBox "File not found: " & File
        SafeOpenBook = False
        Exit Function
    End If

    Workbooks.Open Filename:=File, ReadOnly:=RO
    SafeOpenBook = True
    Exit Function

ErrorHandler:
    MsgBox "Error opening file: " & File & vbNewLine & Err.Description
    SafeOpenBook = False
End Function
```

---

## 🎯 **Development Guidelines**

### **Working with Core Infrastructure**

#### **1. Master Path Usage**
```vba
' ALWAYS use Main.Main_MasterPath.Value for relative paths
Dim fullPath As String
fullPath = Main.Main_MasterPath.Value & "Enquiries\" & fileName

' NEVER hardcode paths
' WRONG: fullPath = "C:\PCS\Enquiries\" & fileName
```

#### **2. File Operation Safety**
```vba
' Check directory exists before file operations
Call CheckDir(Main.Main_MasterPath.Value & "NewDirectory\")

' Use read-only when possible
Call OpenBook(templateFile, True)  ' Read-only for templates
```

#### **3. API Function Usage**
```vba
' Use appropriate version for target Excel
' 32-bit: GetUserNameEx.bas
' 64-bit: GetUserName64.bas
' Universal: Use conditional compilation (V2 approach)
```

#### **4. String Processing**
```vba
' Clean user input before file operations
Dim safeFileName As String
safeFileName = Remove_Characters(userInput) & ".xls"

' Format for display
Dim displayText As String
displayText = Insert_Characters(rawData)
```

### **Testing Core Infrastructure**

#### **File Operation Testing**
```vba
Sub TestCoreInfrastructure()
    ' Test directory operations
    Call CheckDir(Main.Main_MasterPath.Value & "Test\")

    ' Test file operations
    Dim testFile As String
    testFile = Main.Main_MasterPath.Value & "Test\test.txt"

    ' Test user identification
    Dim currentUser As String
    currentUser = Get_User_Name()

    ' Test string utilities
    Dim cleaned As String
    cleaned = Remove_Characters("Test/File: Name")

    MsgBox "Core infrastructure test completed"
End Sub
```

---

## 🔍 **Next Steps**

After understanding Core Infrastructure:

1. **Study [Number Generation](SUBSYSTEM_02_NUMBER_GENERATION.md)** - Learn sequential numbering system
2. **Review [Interface Navigation](SUBSYSTEM_06_INTERFACE_NAVIGATION.md)** - See how Main.frm uses core functions
3. **Examine Usage Patterns** - Look at how other subsystems call core functions
4. **Practice API Functions** - Work with 32/64-bit compatibility
5. **Test File Operations** - Create sample file operation scripts

**Ready for number generation? Continue to [Number Generation Subsystem](SUBSYSTEM_02_NUMBER_GENERATION.md)**