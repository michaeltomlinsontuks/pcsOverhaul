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

## Success Criteria

- Cleaner, more maintainable codebase
- All existing functionality preserved
- Improved modularity for future maintenance
- No breaking changes to user workflows or file storage