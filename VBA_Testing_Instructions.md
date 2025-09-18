# VBA Framework Testing Instructions

## Step-by-Step Testing Process

### Phase 1: Setup Test Environment

1. **Import the test modules into your VBA project:**
   - `VBA_Test_Framework.bas`
   - `Create_Test_Templates.bas`

2. **Create directory structure:**
   ```vba
   CreateTestDirectoryStructure
   ```
   - This will prompt for a base path (e.g., `C:\PCS_Test\`)
   - Creates all required folders

3. **Create template files:**
   ```vba
   CreateAllTestTemplates
   CreateAdditionalTemplates
   ```
   - Creates minimal Excel templates with proper structure
   - Includes the critical `_Enq.xls`, `Search.xls`, `WIP.xls` files

### Phase 2: Import Existing VBA Modules

**Import all your existing `.bas` files:**
- `a_ListFiles.bas`
- `Open_Book.bas`
- `Check_Updates.bas`
- `RefreshMain.bas`
- `RemoveCharacters.bas`
- `Calc_Numbers.bas`
- `GetValue.bas`
- All other modules from `/VBA/Interface_VBA/`

### Phase 3: Run Comprehensive Tests

```vba
RunAllTests
```

**This will systematically test:**
- ✅ Directory structure exists
- ✅ All required template files present
- ✅ Core functions are callable
- ✅ File operations work
- ✅ Data access functions work
- ✅ String utilities work
- ✅ Template files have correct structure

### Phase 4: Interpret Results

The test will show:
- **Green ✓ PASS**: Function works correctly
- **Red ✗ FAIL**: Issue that needs fixing
- **Final Score**: Percentage of tests passed

### Expected Results

**If 90%+ tests pass:** ✅ Framework is solid, proceed with interface
**If 70-90% pass:** ⚠️ Minor issues, review failed tests
**If <70% pass:** ❌ Major issues, fix before building interface

## Common Issues & Solutions

### "Missing directory" errors
**Solution:** Run `CreateTestDirectoryStructure` again

### "Template file missing" errors
**Solution:** Run `CreateAllTestTemplates` and `CreateAdditionalTemplates`

### "Function not found" errors
**Solution:** Import the missing `.bas` module files

### "Main form not found"
**Expected:** This will fail until you build the form

## Testing Individual Components

### Test just file operations:
```vba
TestFileOperations
```

### Test just string functions:
```vba
TestStringFunctions
```

### Test directory structure only:
```vba
TestDirectoryStructure
```

## Next Steps After Testing

1. **If tests pass:** Use the HTML mockup to build your VBA form
2. **Map each HTML control** to the VBA control specification
3. **Wire up event handlers** using the existing `.frm` code as reference
4. **Test incrementally** as you add each control

## Test Data Structure Created

The test setup creates this structure:
```
C:\PCS_Test\
├── enquiries\          # (empty, ready for new enquiries)
├── quotes\             # (empty, ready for quotes)
├── wip\               # (empty, ready for WIP jobs)
├── archive\           # (empty, ready for completed jobs)
├── contracts\         # (empty, ready for contracts)
├── customers\         # (empty, ready for customer files)
├── templates\         # Template files created
│   ├── _Enq.xls      # Main enquiry template
│   ├── _client.xls   # Customer template
│   ├── price list.xls # Price list
│   └── Component_Grades.xls
├── Search.xls         # Central search database
├── WIP.xls           # WIP tracking
├── search History.xls # Search history
├── Job History.xls   # Job history
└── Quote History.xls # Quote history
```

## Final Validation

Once your form is built, test the complete workflow:

1. **Add Enquiry** → Creates file in enquiries folder
2. **Make Quote** → Moves file to quotes folder
3. **Create Job** → Moves file to WIP folder
4. **Close Job** → Moves file to archive folder

Each step should update the central Search.xls database and maintain data integrity.