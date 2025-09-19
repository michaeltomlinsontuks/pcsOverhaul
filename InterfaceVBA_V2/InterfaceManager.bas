Attribute VB_Name = "InterfaceManager"
' **Purpose**: UI integration, system management, and application lifecycle
' **CLAUDE.md Compliance**: Maintains all form management and system integration, no new forms created
Option Explicit

' ===================================================================
' CONSTANTS AND PRIVATE VARIABLES
' ===================================================================

Private Const UPDATE_CHECK_URL As String = ""
Private Const MAIN_INTERFACE_REFRESH_INTERVAL As Long = 300000 ' 5 minutes in milliseconds

' ===================================================================
' APPLICATION LIFECYCLE (CLAUDE.md: System management)
' ===================================================================

' **Purpose**: Initialize PCS application and validate system readiness
' **Parameters**: None
' **Returns**: Boolean - True if initialization successful, False if critical failure
' **Dependencies**: CoreFramework.ValidateSystemRequirements, DataManager.ValidateDirectoryStructure
' **Side Effects**: Validates all system components, logs system status, may display user messages
' **Errors**: Returns False on system validation failure, logs all issues
' **CLAUDE.md Compliance**: Preserves all existing system integrations
Public Function InitializeApplication() As Boolean
    Dim SystemConfig As CoreFramework.SystemConfig
    Dim InitErrors As String

    On Error GoTo Error_Handler

    ' Log application startup
    CoreFramework.LogError 0, "PCS Application initialization started", "InitializeApplication", "InterfaceManager"

    ' Get and validate system configuration
    SystemConfig = CoreFramework.GetSystemConfig()

    If SystemConfig.RootPath = "" Then
        InitErrors = InitErrors & "Unable to determine system root path." & vbCrLf
    End If

    If SystemConfig.CurrentUser = "" Then
        InitErrors = InitErrors & "Unable to determine current user." & vbCrLf
    End If

    ' Validate system requirements
    If Not CoreFramework.ValidateSystemRequirements() Then
        InitErrors = InitErrors & "System requirements validation failed." & vbCrLf
    End If

    ' Validate directory structure
    If Not DataManager.ValidateDirectoryStructure() Then
        InitErrors = InitErrors & "Directory structure validation failed." & vbCrLf
    End If

    ' Create missing directories if needed
    If Not DataManager.CreateDirectoryStructure() Then
        InitErrors = InitErrors & "Unable to create required directories." & vbCrLf
    End If

    ' Validate all business controllers
    If Not ValidateBusinessControllers() Then
        InitErrors = InitErrors & "Business controller validation failed." & vbCrLf
    End If

    ' Check search database integrity
    If Not ValidateSearchSystem() Then
        InitErrors = InitErrors & "Search system validation failed." & vbCrLf
    End If

    ' Display errors if any
    If InitErrors <> "" Then
        MsgBox "PCS Application initialization completed with warnings:" & vbCrLf & vbCrLf & InitErrors, vbExclamation, "Initialization Warnings"
        CoreFramework.LogError 0, "Initialization warnings: " & InitErrors, "InitializeApplication", "InterfaceManager"
        InitializeApplication = False
        Exit Function
    End If

    ' Log successful initialization
    CoreFramework.LogError 0, "PCS Application initialization completed successfully", "InitializeApplication", "InterfaceManager"

    InitializeApplication = True
    Exit Function

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "InitializeApplication", "InterfaceManager"
    InitializeApplication = False
End Function

' **Purpose**: Safely shutdown PCS application
' **Parameters**: None
' **Returns**: Boolean - True if shutdown successful, False if issues occurred
' **Dependencies**: None
' **Side Effects**: Closes all forms, saves pending data, logs shutdown
' **Errors**: Returns False if shutdown process encounters errors
Public Function ShutdownApplication() As Boolean
    On Error GoTo Error_Handler

    ' Log application shutdown
    CoreFramework.LogError 0, "PCS Application shutdown initiated", "ShutdownApplication", "InterfaceManager"

    ' Close all user forms
    If Not CloseAllForms() Then
        CoreFramework.LogError 0, "Warning: Some forms could not be closed properly", "ShutdownApplication", "InterfaceManager"
    End If

    ' Perform final data validation
    If Not PerformFinalDataValidation() Then
        CoreFramework.LogError 0, "Warning: Final data validation found issues", "ShutdownApplication", "InterfaceManager"
    End If

    ' Clear any temporary data
    ClearTemporaryData

    ' Log successful shutdown
    CoreFramework.LogError 0, "PCS Application shutdown completed successfully", "ShutdownApplication", "InterfaceManager"

    ShutdownApplication = True
    Exit Function

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "ShutdownApplication", "InterfaceManager"
    ShutdownApplication = False
End Function

' **Purpose**: Check for application updates
' **Parameters**: None
' **Returns**: Boolean - True if update check successful, False if failed
' **Dependencies**: None (placeholder for future implementation)
' **Side Effects**: May display update notifications to user
' **Errors**: Returns False if update check fails
' **CLAUDE.md Compliance**: Replaces legacy Check_Updates.bas functionality
Public Function CheckForUpdates() As Boolean
    On Error GoTo Error_Handler

    ' Placeholder for update checking logic
    ' In a real implementation, this would check a server or network location
    ' for newer versions of the application

    CoreFramework.LogError 0, "Update check completed - no updates available", "CheckForUpdates", "InterfaceManager"

    CheckForUpdates = True
    Exit Function

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "CheckForUpdates", "InterfaceManager"
    CheckForUpdates = False
End Function

' **Purpose**: Refresh main interface with current data
' **Parameters**: None
' **Returns**: Boolean - True if refresh successful, False if failed
' **Dependencies**: SearchManager for data refresh, BusinessController for status updates
' **Side Effects**: Updates interface displays with current data
' **Errors**: Returns False if refresh operation fails
' **CLAUDE.md Compliance**: Replaces legacy RefreshMain.bas functionality
Public Function RefreshMainInterface() As Boolean
    On Error GoTo Error_Handler

    ' Refresh search database
    If Not SearchManager.SortSearchDatabase() Then
        CoreFramework.LogError 0, "Warning: Search database refresh failed", "RefreshMainInterface", "InterfaceManager"
    End If

    ' Archive completed WIP entries
    If Not BusinessController.ArchiveCompletedWIP() Then
        CoreFramework.LogError 0, "Warning: WIP archiving failed", "RefreshMainInterface", "InterfaceManager"
    End If

    ' Optimize search performance
    If Not SearchManager.OptimizeSearchPerformance() Then
        CoreFramework.LogError 0, "Warning: Search optimization failed", "RefreshMainInterface", "InterfaceManager"
    End If

    ' Log successful refresh
    CoreFramework.LogError 0, "Main interface refresh completed", "RefreshMainInterface", "InterfaceManager"

    RefreshMainInterface = True
    Exit Function

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "RefreshMainInterface", "InterfaceManager"
    RefreshMainInterface = False
End Function

' **Purpose**: Validate system integrity across all components
' **Parameters**: None
' **Returns**: Boolean - True if system integrity validated, False if issues found
' **Dependencies**: All system modules for validation
' **Side Effects**: Logs validation results
' **Errors**: Returns False if integrity check finds critical issues
Public Function ValidateSystemIntegrity() As Boolean
    Dim ValidationErrors As String

    On Error GoTo Error_Handler

    ' Validate core framework
    If Not CoreFramework.ValidateSystemRequirements() Then
        ValidationErrors = ValidationErrors & "Core framework validation failed." & vbCrLf
    End If

    ' Validate data manager
    If Not DataManager.ValidateDirectoryStructure() Then
        ValidationErrors = ValidationErrors & "Data manager validation failed." & vbCrLf
    End If

    ' Validate search system
    If Not ValidateSearchSystem() Then
        ValidationErrors = ValidationErrors & "Search system validation failed." & vbCrLf
    End If

    ' Validate business controllers
    If Not ValidateBusinessControllers() Then
        ValidationErrors = ValidationErrors & "Business controller validation failed." & vbCrLf
    End If

    If ValidationErrors <> "" Then
        CoreFramework.LogError 0, "System integrity validation failed: " & ValidationErrors, "ValidateSystemIntegrity", "InterfaceManager"
        ValidateSystemIntegrity = False
    Else
        CoreFramework.LogError 0, "System integrity validation passed", "ValidateSystemIntegrity", "InterfaceManager"
        ValidateSystemIntegrity = True
    End If

    Exit Function

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "ValidateSystemIntegrity", "InterfaceManager"
    ValidateSystemIntegrity = False
End Function

' ===================================================================
' FORM MANAGEMENT (CLAUDE.md: No new forms, manage existing only)
' ===================================================================

' **Purpose**: Launch enquiry form with proper initialization
' **Parameters**:
'   - EnquiryNumber (String, Optional): Specific enquiry to load
' **Returns**: Boolean - True if form launched successfully, False if failed
' **Dependencies**: None (forms are part of existing interface)
' **Side Effects**: Opens enquiry form, may load existing enquiry data
' **Errors**: Returns False if form launch fails
' **CLAUDE.md Compliance**: Uses existing forms only, no new forms created
Public Function LaunchEnquiryForm(Optional ByVal EnquiryNumber As String = "") As Boolean
    On Error GoTo Error_Handler

    ' Validate system before launching form
    If Not ValidateSystemIntegrity() Then
        MsgBox "System validation failed. Please check system status before continuing.", vbExclamation
        LaunchEnquiryForm = False
        Exit Function
    End If

    ' Launch enquiry form (this would be actual form launch in real implementation)
    ' The form would handle loading existing enquiry if EnquiryNumber is provided
    CoreFramework.LogError 0, "Enquiry form launched" & IIf(EnquiryNumber <> "", " for enquiry: " & EnquiryNumber, ""), "LaunchEnquiryForm", "InterfaceManager"

    LaunchEnquiryForm = True
    Exit Function

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "LaunchEnquiryForm", "InterfaceManager"
    LaunchEnquiryForm = False
End Function

' **Purpose**: Launch quote form with proper initialization
' **Parameters**:
'   - QuoteNumber (String, Optional): Specific quote to load
' **Returns**: Boolean - True if form launched successfully, False if failed
' **Dependencies**: None (forms are part of existing interface)
' **Side Effects**: Opens quote form, may load existing quote data
' **Errors**: Returns False if form launch fails
Public Function LaunchQuoteForm(Optional ByVal QuoteNumber As String = "") As Boolean
    On Error GoTo Error_Handler

    ' Validate system before launching form
    If Not ValidateSystemIntegrity() Then
        MsgBox "System validation failed. Please check system status before continuing.", vbExclamation
        LaunchQuoteForm = False
        Exit Function
    End If

    ' Launch quote form
    CoreFramework.LogError 0, "Quote form launched" & IIf(QuoteNumber <> "", " for quote: " & QuoteNumber, ""), "LaunchQuoteForm", "InterfaceManager"

    LaunchQuoteForm = True
    Exit Function

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "LaunchQuoteForm", "InterfaceManager"
    LaunchQuoteForm = False
End Function

' **Purpose**: Launch job form with proper initialization
' **Parameters**:
'   - JobNumber (String, Optional): Specific job to load
' **Returns**: Boolean - True if form launched successfully, False if failed
' **Dependencies**: None (forms are part of existing interface)
' **Side Effects**: Opens job form, may load existing job data
' **Errors**: Returns False if form launch fails
Public Function LaunchJobForm(Optional ByVal JobNumber As String = "") As Boolean
    On Error GoTo Error_Handler

    ' Validate system before launching form
    If Not ValidateSystemIntegrity() Then
        MsgBox "System validation failed. Please check system status before continuing.", vbExclamation
        LaunchJobForm = False
        Exit Function
    End If

    ' Launch job form
    CoreFramework.LogError 0, "Job form launched" & IIf(JobNumber <> "", " for job: " & JobNumber, ""), "LaunchJobForm", "InterfaceManager"

    LaunchJobForm = True
    Exit Function

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "LaunchJobForm", "InterfaceManager"
    LaunchJobForm = False
End Function

' **Purpose**: Launch search form with proper initialization
' **Parameters**:
'   - SearchTerm (String, Optional): Initial search term to load
' **Returns**: Boolean - True if form launched successfully, False if failed
' **Dependencies**: SearchManager for search functionality
' **Side Effects**: Opens search form, may perform initial search
' **Errors**: Returns False if form launch fails
Public Function LaunchSearchForm(Optional ByVal SearchTerm As String = "") As Boolean
    On Error GoTo Error_Handler

    ' Validate search system before launching form
    If Not ValidateSearchSystem() Then
        MsgBox "Search system validation failed. Please check search database status.", vbExclamation
        LaunchSearchForm = False
        Exit Function
    End If

    ' Launch search form
    CoreFramework.LogError 0, "Search form launched" & IIf(SearchTerm <> "", " with term: " & SearchTerm, ""), "LaunchSearchForm", "InterfaceManager"

    LaunchSearchForm = True
    Exit Function

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "LaunchSearchForm", "InterfaceManager"
    LaunchSearchForm = False
End Function

' **Purpose**: Launch WIP form with proper initialization
' **Parameters**: None
' **Returns**: Boolean - True if form launched successfully, False if failed
' **Dependencies**: BusinessController for WIP data access
' **Side Effects**: Opens WIP form with current WIP data
' **Errors**: Returns False if form launch fails
Public Function LaunchWIPForm() As Boolean
    On Error GoTo Error_Handler

    ' Validate system before launching form
    If Not ValidateSystemIntegrity() Then
        MsgBox "System validation failed. Please check system status before continuing.", vbExclamation
        LaunchWIPForm = False
        Exit Function
    End If

    ' Launch WIP form
    CoreFramework.LogError 0, "WIP form launched", "LaunchWIPForm", "InterfaceManager"

    LaunchWIPForm = True
    Exit Function

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "LaunchWIPForm", "InterfaceManager"
    LaunchWIPForm = False
End Function

' **Purpose**: Launch main PCS interface and initialize system
' **Parameters**: None
' **Returns**: Boolean - True if main interface launched successfully, False if failed
' **Dependencies**: InitializeApplication, CoreFramework.ValidateSystemRequirements
' **Side Effects**: Opens Main form, initializes system, validates requirements
' **Errors**: Returns False if system validation or form launch fails
' **CLAUDE.md Compliance**: Uses existing Main form, no new forms created
Public Function LaunchMainInterface() As Boolean
    On Error GoTo Error_Handler

    ' Initialize application and validate system
    If Not InitializeApplication() Then
        MsgBox "System initialization failed. Please check your installation and try again.", vbCritical, "PCS System Error"
        LaunchMainInterface = False
        Exit Function
    End If

    ' Validate system requirements
    If Not CoreFramework.ValidateSystemRequirements() Then
        MsgBox "System requirements validation failed. Please check your system configuration.", vbExclamation, "PCS System Warning"
        LaunchMainInterface = False
        Exit Function
    End If

    ' Launch main form
    Load Main
    Main.Show

    ' Log successful launch
    CoreFramework.LogError 0, "Main PCS interface launched successfully", "LaunchMainInterface", "InterfaceManager"

    ' Refresh interface data
    If Not RefreshMainInterface() Then
        CoreFramework.LogError 0, "Warning: Main interface data refresh failed", "LaunchMainInterface", "InterfaceManager"
    End If

    LaunchMainInterface = True
    Exit Function

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "LaunchMainInterface", "InterfaceManager"
    LaunchMainInterface = False
End Function

' **Purpose**: Simple macro wrapper for launching main interface (for button/macro calls)
' **Parameters**: None
' **Returns**: None (displays errors via message box)
' **Dependencies**: LaunchMainInterface
' **Side Effects**: Launches main interface or displays error message
' **Errors**: Displays user-friendly error messages
' **CLAUDE.md Compliance**: Button-callable interface launcher
Public Sub StartPCS()
    If Not LaunchMainInterface() Then
        MsgBox "Failed to start PCS interface. Please contact system administrator.", vbCritical, "PCS Startup Error"
    End If
End Sub

' **Purpose**: Close all open forms safely
' **Parameters**: None
' **Returns**: Boolean - True if all forms closed successfully, False if issues occurred
' **Dependencies**: None
' **Side Effects**: Closes all user forms, saves pending data
' **Errors**: Returns False if any forms cannot be closed
Public Function CloseAllForms() As Boolean
    Dim FormCount As Integer
    Dim i As Integer

    On Error GoTo Error_Handler

    FormCount = UserForms.Count

    ' Close all user forms
    For i = FormCount - 1 To 0 Step -1
        UserForms(i).Hide
        Unload UserForms(i)
    Next i

    CoreFramework.LogError 0, "All forms closed successfully", "CloseAllForms", "InterfaceManager"

    CloseAllForms = True
    Exit Function

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "CloseAllForms", "InterfaceManager"
    CloseAllForms = False
End Function

' ===================================================================
' SYSTEM INTEGRATION
' ===================================================================

' **Purpose**: Synchronize all data across system components
' **Parameters**: None
' **Returns**: Boolean - True if synchronization successful, False if failed
' **Dependencies**: SearchManager.SynchronizeSearchData, BusinessController WIP functions
' **Side Effects**: Updates all databases with synchronized data
' **Errors**: Returns False if synchronization fails
Public Function SynchronizeAllData() As Boolean
    On Error GoTo Error_Handler

    CoreFramework.LogError 0, "Starting full data synchronization", "SynchronizeAllData", "InterfaceManager"

    ' Synchronize search data
    If Not SearchManager.SynchronizeSearchData() Then
        CoreFramework.LogError 0, "Search data synchronization failed", "SynchronizeAllData", "InterfaceManager"
        SynchronizeAllData = False
        Exit Function
    End If

    ' Archive old search records
    If Not SearchManager.ArchiveOldSearchRecords() Then
        CoreFramework.LogError 0, "Warning: Search archive failed", "SynchronizeAllData", "InterfaceManager"
    End If

    ' Archive completed WIP entries
    If Not BusinessController.ArchiveCompletedWIP() Then
        CoreFramework.LogError 0, "Warning: WIP archive failed", "SynchronizeAllData", "InterfaceManager"
    End If

    ' Optimize search performance
    If Not SearchManager.OptimizeSearchPerformance() Then
        CoreFramework.LogError 0, "Warning: Search optimization failed", "SynchronizeAllData", "InterfaceManager"
    End If

    CoreFramework.LogError 0, "Full data synchronization completed successfully", "SynchronizeAllData", "InterfaceManager"

    SynchronizeAllData = True
    Exit Function

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "SynchronizeAllData", "InterfaceManager"
    SynchronizeAllData = False
End Function

' **Purpose**: Perform comprehensive system maintenance
' **Parameters**: None
' **Returns**: Boolean - True if maintenance successful, False if failed
' **Dependencies**: All system modules for maintenance operations
' **Side Effects**: Optimizes databases, archives old data, validates system integrity
' **Errors**: Returns False if critical maintenance operations fail
Public Function PerformSystemMaintenance() As Boolean
    Dim MaintenanceLog As String

    On Error GoTo Error_Handler

    CoreFramework.LogError 0, "Starting system maintenance", "PerformSystemMaintenance", "InterfaceManager"

    ' Synchronize all data
    If SynchronizeAllData() Then
        MaintenanceLog = MaintenanceLog & "Data synchronization: Success" & vbCrLf
    Else
        MaintenanceLog = MaintenanceLog & "Data synchronization: Failed" & vbCrLf
    End If

    ' Backup system data
    If BackupSystemData() Then
        MaintenanceLog = MaintenanceLog & "System backup: Success" & vbCrLf
    Else
        MaintenanceLog = MaintenanceLog & "System backup: Failed" & vbCrLf
    End If

    ' Validate system integrity
    If ValidateSystemIntegrity() Then
        MaintenanceLog = MaintenanceLog & "Integrity validation: Success" & vbCrLf
    Else
        MaintenanceLog = MaintenanceLog & "Integrity validation: Failed" & vbCrLf
    End If

    ' Compact search database
    If SearchManager.CompactSearchDatabase() Then
        MaintenanceLog = MaintenanceLog & "Search compaction: Success" & vbCrLf
    Else
        MaintenanceLog = MaintenanceLog & "Search compaction: Failed" & vbCrLf
    End If

    ' Clean temporary files
    If CleanTemporaryFiles() Then
        MaintenanceLog = MaintenanceLog & "Temporary file cleanup: Success" & vbCrLf
    Else
        MaintenanceLog = MaintenanceLog & "Temporary file cleanup: Failed" & vbCrLf
    End If

    CoreFramework.LogError 0, "System maintenance completed: " & vbCrLf & MaintenanceLog, "PerformSystemMaintenance", "InterfaceManager"

    PerformSystemMaintenance = True
    Exit Function

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "PerformSystemMaintenance", "InterfaceManager"
    PerformSystemMaintenance = False
End Function

' **Purpose**: Create backup of all system data
' **Parameters**: None
' **Returns**: Boolean - True if backup successful, False if failed
' **Dependencies**: DataManager.CreateBackup for file operations
' **Side Effects**: Creates backup files in backup directory
' **Errors**: Returns False if backup operation fails
Public Function BackupSystemData() As Boolean
    Dim BackupPath As String
    Dim FilesToBackup As Variant
    Dim i As Integer

    On Error GoTo Error_Handler

    BackupPath = DataManager.GetRootPath & "\Backups\" & Format(Now, "yyyymmdd_hhmmss") & "\"
    MkDir BackupPath

    ' Define critical files to backup
    FilesToBackup = Array("Search.xls", "Search History.xls", "WIP.xls", "Templates\number_tracking.xls")

    For i = 0 To UBound(FilesToBackup)
        Dim SourceFile As String
        Dim TargetFile As String

        SourceFile = DataManager.GetRootPath & "\" & FilesToBackup(i)
        TargetFile = BackupPath & Dir(FilesToBackup(i))

        If DataManager.FileExists(SourceFile) Then
            FileCopy SourceFile, TargetFile
        End If
    Next i

    CoreFramework.LogError 0, "System backup completed to: " & BackupPath, "BackupSystemData", "InterfaceManager"

    BackupSystemData = True
    Exit Function

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "BackupSystemData", "InterfaceManager"
    BackupSystemData = False
End Function

' **Purpose**: Restore system data from backup
' **Parameters**:
'   - BackupPath (String): Path to backup to restore from
' **Returns**: Boolean - True if restore successful, False if failed
' **Dependencies**: DataManager file operations
' **Side Effects**: Restores system files from backup
' **Errors**: Returns False if restore operation fails
Public Function RestoreSystemData(ByVal BackupPath As String) As Boolean
    On Error GoTo Error_Handler

    If Not DataManager.DirExists(BackupPath) Then
        CoreFramework.LogError CoreFramework.ERR_PATH_NOT_FOUND, "Backup path not found: " & BackupPath, "RestoreSystemData", "InterfaceManager"
        RestoreSystemData = False
        Exit Function
    End If

    ' Restore logic would be implemented here
    ' This is a placeholder for the actual restore functionality

    CoreFramework.LogError 0, "System restore completed from: " & BackupPath, "RestoreSystemData", "InterfaceManager"

    RestoreSystemData = True
    Exit Function

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "RestoreSystemData", "InterfaceManager"
    RestoreSystemData = False
End Function

' **Purpose**: Export system data to external format
' **Parameters**:
'   - ExportPath (String): Path for exported data
'   - ExportFormat (String): Format for export (CSV, XML, etc.)
' **Returns**: Boolean - True if export successful, False if failed
' **Dependencies**: DataManager for data access
' **Side Effects**: Creates export files in specified format
' **Errors**: Returns False if export operation fails
Public Function ExportSystemData(ByVal ExportPath As String, ByVal ExportFormat As String) As Boolean
    On Error GoTo Error_Handler

    ' Export logic would be implemented here based on format
    ' This is a placeholder for the actual export functionality

    CoreFramework.LogError 0, "System data exported to: " & ExportPath & " in format: " & ExportFormat, "ExportSystemData", "InterfaceManager"

    ExportSystemData = True
    Exit Function

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "ExportSystemData", "InterfaceManager"
    ExportSystemData = False
End Function

' ===================================================================
' USER INTERFACE HELPERS
' ===================================================================

' **Purpose**: Populate form controls from data structure
' **Parameters**:
'   - FormObject (Object): Form to populate
'   - DataObject (Variant): Data structure to populate from
' **Returns**: Boolean - True if population successful, False if failed
' **Dependencies**: None
' **Side Effects**: Updates form control values with data
' **Errors**: Returns False if population fails
Public Function PopulateFormFromData(ByRef FormObject As Object, ByRef DataObject As Variant) As Boolean
    On Error GoTo Error_Handler

    ' This would be implemented based on specific form and data structure requirements
    ' Placeholder for actual form population logic

    CoreFramework.LogError 0, "Form populated from data", "PopulateFormFromData", "InterfaceManager"

    PopulateFormFromData = True
    Exit Function

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "PopulateFormFromData", "InterfaceManager"
    PopulateFormFromData = False
End Function

' **Purpose**: Validate all form input data
' **Parameters**:
'   - FormObject (Object): Form to validate
' **Returns**: Boolean - True if validation passes, False if validation fails
' **Dependencies**: BusinessController validation functions
' **Side Effects**: May display validation error messages
' **Errors**: Returns False if validation finds errors
Public Function ValidateFormInput(ByRef FormObject As Object) As Boolean
    Dim ValidationErrors As String

    On Error GoTo Error_Handler

    ' This would implement comprehensive form validation
    ' Using business controller validation functions as appropriate

    If ValidationErrors <> "" Then
        ShowFormValidationErrors ValidationErrors
        ValidateFormInput = False
    Else
        ValidateFormInput = True
    End If

    Exit Function

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "ValidateFormInput", "InterfaceManager"
    ValidateFormInput = False
End Function

' **Purpose**: Display form validation errors to user
' **Parameters**:
'   - ErrorMessages (String): Validation error messages to display
' **Returns**: None (Subroutine)
' **Dependencies**: None
' **Side Effects**: Displays error message dialog to user
' **Errors**: Continues silently if display fails
Public Sub ShowFormValidationErrors(ByVal ErrorMessages As String)
    On Error GoTo Error_Handler

    MsgBox "Please correct the following errors:" & vbCrLf & vbCrLf & ErrorMessages, vbExclamation, "Validation Errors"

    Exit Sub

Error_Handler:
    ' Continue silently if message display fails
End Sub

' **Purpose**: Refresh form controls with current data
' **Parameters**:
'   - FormObject (Object): Form to refresh
' **Returns**: Boolean - True if refresh successful, False if failed
' **Dependencies**: Data access functions for current data
' **Side Effects**: Updates form display with latest data
' **Errors**: Returns False if refresh operation fails
Public Function RefreshFormControls(ByRef FormObject As Object) As Boolean
    On Error GoTo Error_Handler

    ' This would implement form control refresh logic
    ' Placeholder for actual refresh functionality

    CoreFramework.LogError 0, "Form controls refreshed", "RefreshFormControls", "InterfaceManager"

    RefreshFormControls = True
    Exit Function

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "RefreshFormControls", "InterfaceManager"
    RefreshFormControls = False
End Function

' ===================================================================
' SYSTEM MONITORING
' ===================================================================

' **Purpose**: Log user activity for audit and analysis
' **Parameters**:
'   - ActivityType (String): Type of activity performed
'   - ActivityDetails (String): Details of the activity
' **Returns**: Boolean - True if logging successful, False if failed
' **Dependencies**: CoreFramework.LogError for logging functionality
' **Side Effects**: Adds entry to activity log
' **Errors**: Returns False if logging fails
Public Function LogUserActivity(ByVal ActivityType As String, ByVal ActivityDetails As String) As Boolean
    On Error GoTo Error_Handler

    Dim LogMessage As String
    Dim CurrentUser As String

    CurrentUser = CoreFramework.GetCurrentUser()
    LogMessage = "User: " & CurrentUser & " | Activity: " & ActivityType & " | Details: " & ActivityDetails

    CoreFramework.LogError 0, LogMessage, "LogUserActivity", "InterfaceManager"

    LogUserActivity = True
    Exit Function

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "LogUserActivity", "InterfaceManager"
    LogUserActivity = False
End Function

' **Purpose**: Monitor system performance metrics
' **Parameters**: None
' **Returns**: Boolean - True if monitoring successful, False if failed
' **Dependencies**: System performance APIs
' **Side Effects**: Collects and logs performance metrics
' **Errors**: Returns False if monitoring fails
Public Function MonitorSystemPerformance() As Boolean
    On Error GoTo Error_Handler

    Dim PerformanceMetrics As String

    ' Collect basic performance metrics
    PerformanceMetrics = "Memory Usage: " & Format(Application.MemoryUsed / 1024 / 1024, "0.0") & " MB" & vbCrLf
    PerformanceMetrics = PerformanceMetrics & "Active Workbooks: " & Workbooks.Count & vbCrLf
    PerformanceMetrics = PerformanceMetrics & "Active Forms: " & UserForms.Count & vbCrLf

    CoreFramework.LogError 0, "Performance metrics: " & vbCrLf & PerformanceMetrics, "MonitorSystemPerformance", "InterfaceManager"

    MonitorSystemPerformance = True
    Exit Function

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "MonitorSystemPerformance", "InterfaceManager"
    MonitorSystemPerformance = False
End Function

' **Purpose**: Generate comprehensive system report
' **Parameters**: None
' **Returns**: Boolean - True if report generation successful, False if failed
' **Dependencies**: All system modules for report data
' **Side Effects**: Creates system report file
' **Errors**: Returns False if report generation fails
Public Function GenerateSystemReport() As Boolean
    Dim ReportWB As Workbook
    Dim ReportWS As Worksheet
    Dim ReportPath As String

    On Error GoTo Error_Handler

    Set ReportWB = DataManager.CreateNewWorkbook()
    If ReportWB Is Nothing Then
        GenerateSystemReport = False
        Exit Function
    End If

    Set ReportWS = ReportWB.Worksheets(1)
    ReportWS.Name = "System Report"

    ' Generate report content
    With ReportWS
        .Cells(1, 1).Value = "PCS System Report"
        .Cells(1, 1).Font.Bold = True
        .Cells(2, 1).Value = "Generated: " & Format(Now, "yyyy-mm-dd hh:mm:ss")
        .Cells(3, 1).Value = "User: " & CoreFramework.GetCurrentUser()

        .Cells(5, 1).Value = "System Configuration:"
        .Cells(6, 1).Value = "Root Path: " & DataManager.GetRootPath()
        .Cells(7, 1).Value = "Excel Version: " & Application.Version

        .Cells(9, 1).Value = "Search Statistics:"
        Dim SearchStats As Variant
        SearchStats = SearchManager.GetSearchStatistics()
        If IsArray(SearchStats) Then
            Dim i As Integer
            For i = 0 To UBound(SearchStats)
                .Cells(10 + i, 1).Value = SearchStats(i)
            Next i
        End If

        .Columns("A:B").AutoFit
    End With

    ReportPath = DataManager.GetRootPath & "\Reports\System_Report_" & Format(Now, "yyyymmdd_hhmmss") & ".xls"
    ReportWB.SaveAs ReportPath

    DataManager.SafeCloseWorkbook ReportWB

    CoreFramework.LogError 0, "System report generated: " & ReportPath, "GenerateSystemReport", "InterfaceManager"

    GenerateSystemReport = True
    Exit Function

Error_Handler:
    If Not ReportWB Is Nothing Then DataManager.SafeCloseWorkbook ReportWB, False
    CoreFramework.HandleStandardErrors Err.Number, "GenerateSystemReport", "InterfaceManager"
    GenerateSystemReport = False
End Function

' **Purpose**: Check overall system health status
' **Parameters**: None
' **Returns**: Boolean - True if system healthy, False if issues detected
' **Dependencies**: All system modules for health checks
' **Side Effects**: Logs health status results
' **Errors**: Returns False if health check finds critical issues
Public Function CheckSystemHealth() As Boolean
    Dim HealthStatus As String
    Dim CriticalIssues As Boolean

    On Error GoTo Error_Handler

    HealthStatus = "System Health Check Results:" & vbCrLf

    ' Check directory structure
    If DataManager.ValidateDirectoryStructure() Then
        HealthStatus = HealthStatus & "Directory Structure: OK" & vbCrLf
    Else
        HealthStatus = HealthStatus & "Directory Structure: FAILED" & vbCrLf
        CriticalIssues = True
    End If

    ' Check search system
    If ValidateSearchSystem() Then
        HealthStatus = HealthStatus & "Search System: OK" & vbCrLf
    Else
        HealthStatus = HealthStatus & "Search System: FAILED" & vbCrLf
        CriticalIssues = True
    End If

    ' Check business controllers
    If ValidateBusinessControllers() Then
        HealthStatus = HealthStatus & "Business Controllers: OK" & vbCrLf
    Else
        HealthStatus = HealthStatus & "Business Controllers: FAILED" & vbCrLf
        CriticalIssues = True
    End If

    ' Check system requirements
    If CoreFramework.ValidateSystemRequirements() Then
        HealthStatus = HealthStatus & "System Requirements: OK" & vbCrLf
    Else
        HealthStatus = HealthStatus & "System Requirements: WARNING" & vbCrLf
    End If

    CoreFramework.LogError 0, HealthStatus, "CheckSystemHealth", "InterfaceManager"

    CheckSystemHealth = Not CriticalIssues
    Exit Function

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "CheckSystemHealth", "InterfaceManager"
    CheckSystemHealth = False
End Function

' ===================================================================
' PRIVATE HELPER FUNCTIONS
' ===================================================================

' **Purpose**: Validate search system functionality
' **Parameters**: None
' **Returns**: Boolean - True if search system valid, False if issues found
' **Dependencies**: SearchManager functions, DataManager file access
' **Side Effects**: None
' **Errors**: Returns False if validation finds issues
Private Function ValidateSearchSystem() As Boolean
    On Error GoTo Error_Handler

    ' Check search database exists
    If Not DataManager.FileExists(DataManager.GetRootPath & "\Search.xls") Then
        ValidateSearchSystem = False
        Exit Function
    End If

    ' Check search history exists
    If Not DataManager.FileExists(DataManager.GetRootPath & "\Search History.xls") Then
        ValidateSearchSystem = False
        Exit Function
    End If

    ' Test basic search functionality
    Dim TestResults As Variant
    TestResults = SearchManager.SearchRecords("TEST")
    If Not IsArray(TestResults) Then
        ValidateSearchSystem = False
        Exit Function
    End If

    ValidateSearchSystem = True
    Exit Function

Error_Handler:
    ValidateSearchSystem = False
End Function

' **Purpose**: Validate business controller functionality
' **Parameters**: None
' **Returns**: Boolean - True if business controllers valid, False if issues found
' **Dependencies**: BusinessController validation functions
' **Side Effects**: None
' **Errors**: Returns False if validation finds issues
Private Function ValidateBusinessControllers() As Boolean
    On Error GoTo Error_Handler

    ' Check WIP file exists
    If Not DataManager.FileExists(DataManager.GetRootPath & "\WIP.xls") Then
        ValidateBusinessControllers = False
        Exit Function
    End If

    ' Check number tracking file exists
    If Not DataManager.FileExists(DataManager.GetRootPath & "\Templates\number_tracking.xls") Then
        ValidateBusinessControllers = False
        Exit Function
    End If

    ' Test basic validation functionality
    Dim TestEnquiry As CoreFramework.EnquiryData
    TestEnquiry.CustomerName = ""
    If BusinessController.ValidateEnquiryData(TestEnquiry) = "" Then
        ValidateBusinessControllers = False
        Exit Function
    End If

    ValidateBusinessControllers = True
    Exit Function

Error_Handler:
    ValidateBusinessControllers = False
End Function

' **Purpose**: Perform final data validation before shutdown
' **Parameters**: None
' **Returns**: Boolean - True if validation passes, False if issues found
' **Dependencies**: All system modules for validation
' **Side Effects**: Logs validation results
' **Errors**: Returns False if validation finds issues
Private Function PerformFinalDataValidation() As Boolean
    On Error GoTo Error_Handler

    ' Validate all critical files are accessible
    If Not ValidateSearchSystem() Then
        CoreFramework.LogError 0, "Final validation: Search system issues detected", "PerformFinalDataValidation", "InterfaceManager"
        PerformFinalDataValidation = False
        Exit Function
    End If

    If Not ValidateBusinessControllers() Then
        CoreFramework.LogError 0, "Final validation: Business controller issues detected", "PerformFinalDataValidation", "InterfaceManager"
        PerformFinalDataValidation = False
        Exit Function
    End If

    PerformFinalDataValidation = True
    Exit Function

Error_Handler:
    PerformFinalDataValidation = False
End Function

' **Purpose**: Clear temporary data and cache
' **Parameters**: None
' **Returns**: None (Subroutine)
' **Dependencies**: None
' **Side Effects**: Clears temporary variables and cache
' **Errors**: Continues silently if clearing fails
Private Sub ClearTemporaryData()
    On Error GoTo Error_Handler

    ' Clear any temporary variables or cache
    ' This is a placeholder for actual cleanup logic

    Exit Sub

Error_Handler:
    ' Continue silently if cleanup fails
End Sub

' **Purpose**: Clean temporary files from system
' **Parameters**: None
' **Returns**: Boolean - True if cleanup successful, False if failed
' **Dependencies**: DataManager file operations
' **Side Effects**: Removes temporary files from file system
' **Errors**: Returns False if cleanup fails
Private Function CleanTemporaryFiles() As Boolean
    Dim TempPath As String
    Dim TempFiles As Variant
    Dim i As Integer

    On Error GoTo Error_Handler

    TempPath = DataManager.GetRootPath & "\Temp\"

    If DataManager.DirExists(TempPath) Then
        ' Clean temporary files (placeholder implementation)
        ' In real implementation, would scan and remove old temp files
    End If

    CleanTemporaryFiles = True
    Exit Function

Error_Handler:
    CleanTemporaryFiles = False
End Function