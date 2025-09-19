# PCS Overhaul Development Rules

## Project Scope and Goals

**PRIMARY GOAL**: Refactor existing VBA interface code to make it cleaner and more maintainable while preserving all existing functionality.

### Hard Rules

1. **NO NEW FORMS**: Do not create new UserForms or interfaces. Work only with existing forms and functionality.

2. **COMPATIBILITY REQUIREMENTS**:
   - All code must work with both 32-bit and 64-bit Excel
   - Must maintain compatibility with existing directory structure
   - Tens of thousands of files depend on current structure - DO NOT CHANGE IT

3. **EXISTING FRAMEWORK PRESERVATION**:
   - Maintain the current subsystem flow: Enquiry → Quote → Jobs
   - Preserve Jobs → Job Cards → WIP Reports workflow
   - Keep Contracts (Job Templates) functionality intact
   - Maintain Search functionality (finds anything in the system)

4. **CODE QUALITY OBJECTIVES**:
   - Make code more modular and maintainable
   - Fix formatting issues in reports
   - Remove unused/dead code
   - Improve code organization and structure

5. **DEVELOPMENT APPROACH**:
   - Export existing code using macros (completed)
   - Build framework around existing subsystems, classes, and modules
   - Identify and remove dead code
   - Remap existing forms to use new/refactored functions
   - Maintain backward compatibility throughout

6. **FORBIDDEN ACTIONS**:
   - Creating new UserForms or interfaces
   - Changing directory structure
   - Breaking compatibility with existing file storage system
   - Removing functionality without replacement

## Testing Requirements

When making changes:
- Test with both 32-bit and 64-bit Excel
- Verify all existing workflows still function
- Ensure file paths and directory access remain intact
- Test all forms and reports for proper functionality

## Documentation Requirements

### Mandatory Documentation Updates

**WHEN**: Documentation MUST be updated immediately when:
- Adding new functions or subroutines
- Modifying function parameters or return values
- Changing business logic or workflows
- Adding or removing Excel sheet columns/fields
- Modifying data structures or types
- Changing file paths or directory interactions
- Updating error handling patterns

### Documentation Locations

**PRIMARY SYSTEM DOCUMENTATION**:
- `PCS_V2_SYSTEM_DOCUMENTATION.md` - Complete V2 system reference
- `PCS_OLD_SYSTEM_DOCUMENTATION.md` - Legacy system reference (read-only)

**FUNCTION-LEVEL DOCUMENTATION**:
Each function MUST include doxygen-style comments:
```vba
' **Purpose**: Brief description of what function does
' **Parameters**:
'   - paramName (Type): Description
' **Returns**: Type and description of return value
' **Dependencies**: List of called functions/modules
' **Side Effects**: Files created, sheets modified, etc.
' **Errors**: Error handling approach
```

### Documentation Standards

**ABSOLUTE REQUIREMENTS**:
1. **Function Signatures**: All public functions documented with parameters and return types
2. **Data Structures**: All Type definitions documented with field purposes
3. **Excel Schema**: Sheet structures with column mappings and business rules
4. **Workflow Maps**: Business process flows maintained with function call chains
5. **Field Mappings**: Cross-reference of all data fields used across sheets
6. **Error Codes**: All error handling patterns documented with recovery steps

**UPDATE TRIGGERS**:
- Code changes → Update function documentation immediately
- Schema changes → Update Excel structure sections
- Workflow changes → Update process flow diagrams
- New modules → Add to appropriate subsystem documentation
- Bug fixes → Document in error handling sections

### Documentation Verification

**BEFORE COMMITTING CODE**:
1. Verify all modified functions have updated documentation
2. Check that Excel schema changes are reflected in documentation
3. Ensure workflow changes are mapped in process documentation
4. Validate that all new error patterns are documented

**DOCUMENTATION DEBT**:
- Never commit code without corresponding documentation updates
- Missing documentation is considered a blocker for code acceptance
- Documentation accuracy is as critical as code functionality

## Success Criteria

- Cleaner, more maintainable codebase
- All existing functionality preserved
- Improved modularity for future maintenance
- No breaking changes to user workflows or file storage