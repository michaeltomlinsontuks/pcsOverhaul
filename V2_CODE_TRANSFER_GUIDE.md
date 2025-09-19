# V2 Code Transfer Guide: File Import Instructions

## 🎯 **Current Status**
- ✅ 5 consolidated modules ready for import
- ✅ 9 updated forms ready for import
- ✅ All legacy V2 modules removed
- ✅ All forms updated to use consolidated modules

## 📁 **Files Ready for Import**

**Location**: `InterfaceVBA_V2/`

**Core Modules (5)**:
- CoreFramework.bas
- DataManager.bas
- SearchManager.bas
- BusinessController.bas
- InterfaceManager.bas

**Updated Forms (9)**:
- Main.frm
- FEnquiry.frm
- FQuote.frm
- FAcceptQuote.frm
- FJobCard.frm
- FJG.frm
- fwip.frm
- frmSearch.frm
- frmSearchNew.frm

## 🔄 **Import Order & Steps**

### Step 1: Open Interface.xls
```
1. Open Interface.xls
2. Press Alt+F11 to open VBA Editor
3. Backup current project before importing
```

### Step 2: Import Core Modules (Order Matters)
```
Import in this exact order:

1. CoreFramework.bas       (Foundation - types, errors, utilities)
2. DataManager.bas         (File operations, Excel access)
3. SearchManager.bas       (Search functionality)
4. BusinessController.bas  (Business logic - depends on above)
5. InterfaceManager.bas    (System integration - depends on all above)
```

**Import Method**:
- Right-click project → Import File... → Select .bas file
- Or drag .bas files into VBA Editor

### Step 3: Import Updated Forms
```
Import all 9 .frm files (order doesn't matter):

- Main.frm
- FEnquiry.frm
- FQuote.frm
- FAcceptQuote.frm
- FJobCard.frm
- FJG.frm
- fwip.frm
- frmSearch.frm
- frmSearchNew.frm
```

**Import Method**:
- Right-click project → Import File... → Select .frm file
- Replace existing forms when prompted

### Step 4: Compile & Test
```
1. Debug → Compile VBAProject
2. Fix any compilation errors
3. Test basic form loading
```

### Step 5: Remove Legacy Modules
```
After successful testing:

1. Remove all old Interface_VBA/ modules
2. Remove any remaining old module references
3. Delete Search.xls (functionality now in Interface.xls)
```

## ⚠️ **Important Notes**

**Module Dependencies**:
- CoreFramework.bas must be imported first (other modules depend on it)
- BusinessController.bas depends on CoreFramework, DataManager, SearchManager
- InterfaceManager.bas depends on all other modules

**Form Updates**:
- Forms are already updated to use consolidated modules
- No additional code changes needed after import
- All forms reference: CoreFramework, BusinessController, DataManager, SearchManager

**Compilation**:
- Must compile successfully before testing
- All forms depend on the 5 core modules being present
- Test each major workflow after import

## 🎯 **Success Criteria**

✅ All 5 modules imported without errors
✅ All 9 forms imported without errors
✅ Project compiles successfully
✅ Basic form functionality works
✅ Enquiry → Quote → Jobs workflow functions
✅ Search functionality works
✅ WIP reports generate correctly

## 🚨 **Rollback Plan**

If import fails:
1. Restore Interface.xls backup
2. Restore Search.xls backup
3. Check import order and dependencies
4. Re-attempt import following exact steps above