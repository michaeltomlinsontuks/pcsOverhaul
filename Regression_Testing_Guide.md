# Regression Testing Guide

## Overview
This regression testing framework runs **identical operations** on both your original VBA system and your new replacement, then compares all results to ensure they match exactly.

## Prerequisites

1. **Working original system** (your current VBA interface)
2. **New system** (your rebuilt interface with new forms)
3. **Test data** (sample enquiries, quotes, jobs)
4. **Both systems pointing to separate data folders**

## Step-by-Step Process

### Phase 1: Environment Setup

1. **Import the testing modules:**
   ```vba
   ' Import these into your VBA project:
   VBA_Regression_Testing.bas
   Regression_Testing_Implementation.bas
   ```

2. **Prepare two separate data environments:**
   ```
   C:\PCS_Original\    (your current working system)
   C:\PCS_New\         (your new system copy)
   ```

3. **Run environment setup:**
   ```vba
   SetupRegressionEnvironment
   ```
   - This will prompt for both system paths
   - Creates backup of original data
   - Prepares testing framework

### Phase 2: Customize Test Functions

**Edit these functions to call your actual code:**

#### In `VBA_Regression_Testing.bas`:

```vba
' Replace mock functions with real calls:

Private Function RunOriginalEnquiryCreation(testData As String) As String
    ' Call your original Main.Add_Enquiry_Click functionality
    ' Capture the actual results (file created, numbers generated, etc.)
    ' Return standardized result string
End Function

Private Function RunNewEnquiryCreation(testData As String) As String
    ' Call your new form's enquiry creation
    ' Should produce identical results
    ' Return standardized result string
End Function
```

#### Key functions to customize:
- `RunOriginalMakeQuote()` → Call original quote creation
- `RunNewMakeQuote()` → Call new quote creation
- `RunOriginalCreateJob()` → Call original job creation
- `RunNewCreateJob()` → Call new job creation
- `GetOriginalFileList()` → Call original List_Files function
- `GetNewFileList()` → Call new list population logic

### Phase 3: Execute Regression Tests

#### Basic Test (Simulated):
```vba
RunRegressionTests
```
This runs with mock data to verify the framework works.

#### Real System Test:
```vba
ExecuteRealRegressionTest
```
This runs actual operations on both systems and compares results.

### Phase 4: Interpret Results

#### Perfect Match (Goal):
```
✓ PASS: Enquiry Creation
✓ PASS: Make Quote
✓ PASS: Create Job
✓ PASS: File Migration
✓ PASS: List Population
✓ PASS: Status Updates

🎉 PERFECT MATCH! New system behavior is identical to original.
✅ Safe to deploy new interface.
```

#### Typical Issues Found:
```
✗ FAIL: Enquiry Number Generation
    Original: ENQ20241201001
    New:      ENQ20241201002

✗ FAIL: File Migration
    Original: enquiries->quotes->wip->archive
    New:      enquiries->quotes->archive (missing WIP step)
```

## Test Scenarios Covered

### 1. **Workflow Tests**
- Enquiry → Quote → Job → Archive progression
- File movement between folders
- Status updates at each step
- Number generation consistency

### 2. **Data Integrity Tests**
- Search.xls database updates
- WIP.xls tracking updates
- File content preservation
- Named range consistency

### 3. **UI Behavior Tests**
- List population when toggles clicked
- Form control value updates
- Filter behavior (WIP, Enquiries, etc.)
- Status display updates

### 4. **File Operation Tests**
- Template usage and copying
- Excel file creation and saving
- Directory navigation
- File reading/writing

## Common Customizations Needed

### 1. **Path Configuration**
```vba
Private Function GetOriginalMasterPath() As String
    GetOriginalMasterPath = "C:\YourActualPath\"
End Function
```

### 2. **Number Generation Testing**
```vba
Private Function GenerateOriginalEnquiryNumber() As String
    ' Call your actual Calc_Next_Number function
    GenerateOriginalEnquiryNumber = Calc_Next_Number("ENQ")
End Function
```

### 3. **Form Integration**
```vba
Private Function RunOriginalEnquiryCreation(testData As String) As String
    ' Parse test data
    Dim customer As String, component As String
    customer = ExtractValue(testData, "Customer")
    component = ExtractValue(testData, "Component")

    ' Set up original form
    Main.Customer.Value = customer
    Main.Component_Description.Value = component

    ' Trigger original workflow
    Main.Add_Enquiry_Click

    ' Capture results
    Dim result As String
    result = "EnquiryNumber=" & Main.Enquiry_Number.Value
    result = result & "|Status=" & Main.System_Status.Value

    Return result
End Function
```

## Validation Checklist

Before deploying your new interface, ensure:

- [ ] **100% workflow compatibility** - All enquiry→quote→job→archive flows identical
- [ ] **File structure identical** - Same folders, same file naming, same templates
- [ ] **Database updates identical** - Search.xls, WIP.xls updated identically
- [ ] **Number generation identical** - Sequential numbering preserved
- [ ] **Status tracking identical** - All status changes happen same way
- [ ] **List behavior identical** - Toggle buttons populate lists identically
- [ ] **Template usage identical** - Same Excel templates used same way

## Troubleshooting

### "Function not found" errors
- Import all original `.bas` modules
- Check function names match exactly
- Verify modules are in same VBA project

### "Path not found" errors
- Run `SetupRegressionEnvironment` first
- Ensure both system folders exist
- Check folder permissions

### Results always mismatch
- Add debugging to see what's actually returned
- Check date/time dependencies (use fixed test dates)
- Verify both systems use same Master Path

## Best Practices

1. **Start with simple tests** - Test one operation at a time
2. **Use fixed test data** - Avoid random dates/numbers during comparison
3. **Test incrementally** - Add one new control at a time to new interface
4. **Document differences** - Track any intentional changes/improvements
5. **Backup everything** - Keep original system safe during testing

## Success Criteria

**Ready for deployment when:**
- 95%+ regression tests pass
- All critical workflows (enquiry/quote/job) work identically
- File operations produce identical results
- Database updates are consistent
- No data loss or corruption in testing

The goal is **behavioral equivalence** - users should not notice any difference in functionality between old and new interfaces.