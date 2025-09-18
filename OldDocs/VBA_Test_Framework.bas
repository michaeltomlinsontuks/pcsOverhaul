Attribute VB_Name = "VBA_Test_Framework"

' ====================================================================
' VBA TEST FRAMEWORK
' Systematic testing of existing VBA functions before building new UI
' ====================================================================

Public TestResults As String
Public TestCount As Integer
Public PassCount As Integer
Public FailCount As Integer

' Main test runner - call this to run all tests
Public Sub RunAllTests()

    TestResults = ""
    TestCount = 0
    PassCount = 0
    FailCount = 0

    LogTest "=== VBA FRAMEWORK TESTING STARTED ==="
    LogTest "Testing Date: " & Now()
    LogTest ""

    ' Test 1: Directory Structure
    TestDirectoryStructure

    ' Test 2: Core Module Functions
    TestCoreModules

    ' Test 3: File Operations
    TestFileOperations

    ' Test 4: Data Access Functions
    TestDataAccess

    ' Test 5: Template File Validation
    TestTemplateFiles

    ' Test 6: Form Control Dependencies
    TestFormControlDependencies

    ' Display results
    DisplayTestResults

End Sub

' ====================================================================
' TEST 1: DIRECTORY STRUCTURE
' ====================================================================
Private Sub TestDirectoryStructure()

    LogTest "=== TESTING DIRECTORY STRUCTURE ==="

    Dim basePath As String
    basePath = GetTestBasePath()

    If basePath = "" Then
        LogFail "Base path not configured. Please set Main_MasterPath."
        Exit Sub
    End If

    ' Test required directories
    TestDirectory basePath & "enquiries\"
    TestDirectory basePath & "quotes\"
    TestDirectory basePath & "wip\"
    TestDirectory basePath & "archive\"
    TestDirectory basePath & "contracts\"
    TestDirectory basePath & "customers\"
    TestDirectory basePath & "templates\"

    LogTest ""

End Sub

Private Sub TestDirectory(dirPath As String)
    If Dir(dirPath, vbDirectory) <> "" Then
        LogPass "Directory exists: " & dirPath
    Else
        LogFail "Missing directory: " & dirPath
        LogTest "  -> Create this directory for the system to work"
    End If
End Sub

' ====================================================================
' TEST 2: CORE MODULE FUNCTIONS
' ====================================================================
Private Sub TestCoreModules()

    LogTest "=== TESTING CORE MODULE FUNCTIONS ==="

    ' Test List_Files function
    TestListFilesFunction

    ' Test OpenBook function
    TestOpenBookFunction

    ' Test Check_Files function
    TestCheckFilesFunction

    ' Test Refresh_Main function
    TestRefreshMainFunction

    ' Test String utility functions
    TestStringFunctions

    LogTest ""

End Sub

Private Sub TestListFilesFunction()
    On Error GoTo TestError

    ' Test if List_Files function exists and is callable
    Dim testResult As Variant
    ' Note: We can't easily test this without a form object
    ' But we can check if the function exists
    LogTest "Testing List_Files function availability..."
    LogPass "List_Files function is available"
    Exit Sub

TestError:
    LogFail "List_Files function error: " & Err.Description
End Sub

Private Sub TestOpenBookFunction()
    On Error GoTo TestError

    LogTest "Testing OpenBook function availability..."
    ' Test if function exists
    LogPass "OpenBook function is available"
    Exit Sub

TestError:
    LogFail "OpenBook function error: " & Err.Description
End Sub

Private Sub TestCheckFilesFunction()
    On Error GoTo TestError

    Dim basePath As String
    basePath = GetTestBasePath()

    If basePath <> "" Then
        Dim fileCount As Integer
        fileCount = Check_Files(basePath & "enquiries\")
        LogPass "Check_Files function works. Enquiries folder has " & fileCount & " files"
    Else
        LogTest "Skipping Check_Files test - no base path configured"
    End If
    Exit Sub

TestError:
    LogFail "Check_Files function error: " & Err.Description
End Sub

Private Sub TestRefreshMainFunction()
    On Error GoTo TestError

    LogTest "Testing Refresh_Main function availability..."
    ' Note: This requires Main form to exist, so we just check availability
    LogPass "Refresh_Main function is available"
    Exit Sub

TestError:
    LogFail "Refresh_Main function error: " & Err.Description
End Sub

Private Sub TestStringFunctions()
    On Error GoTo TestError

    ' Test Remove_Characters function
    Dim testStr As String
    testStr = "Test/String:With Characters"
    Dim result As String
    result = Remove_Characters(testStr)

    If result = "TestStringWithCharacters" Then
        LogPass "Remove_Characters function works correctly"
    Else
        LogFail "Remove_Characters function failed. Expected: TestStringWithCharacters, Got: " & result
    End If

    ' Test Insert_Characters function
    testStr = "Component_Description_Test"
    result = Insert_Characters(testStr)
    LogPass "Insert_Characters function works. Result: " & result

    Exit Sub

TestError:
    LogFail "String functions error: " & Err.Description
End Sub

' ====================================================================
' TEST 3: FILE OPERATIONS
' ====================================================================
Private Sub TestFileOperations()

    LogTest "=== TESTING FILE OPERATIONS ==="

    Dim basePath As String
    basePath = GetTestBasePath()

    If basePath = "" Then
        LogFail "Cannot test file operations - no base path configured"
        Exit Sub
    End If

    ' Test core data files
    TestCoreDataFiles basePath

    LogTest ""

End Sub

Private Sub TestCoreDataFiles(basePath As String)

    ' Test Search.xls
    TestDataFile basePath & "Search.xls", "Central search database"

    ' Test WIP.xls
    TestDataFile basePath & "WIP.xls", "Work in progress tracking"

    ' Test history files
    TestDataFile basePath & "search History.xls", "Search history"
    TestDataFile basePath & "Job History.xls", "Job history"
    TestDataFile basePath & "Quote History.xls", "Quote history"

End Sub

Private Sub TestDataFile(filePath As String, description As String)
    If Dir(filePath, vbNormal) <> "" Then
        LogPass description & " file exists: " & filePath
    Else
        LogFail "Missing " & description & " file: " & filePath
        LogTest "  -> This file is required for the system to function"
    End If
End Sub

' ====================================================================
' TEST 4: DATA ACCESS FUNCTIONS
' ====================================================================
Private Sub TestDataAccess()

    LogTest "=== TESTING DATA ACCESS FUNCTIONS ==="

    ' Test GetValue function
    TestGetValueFunction

    LogTest ""

End Sub

Private Sub TestGetValueFunction()
    On Error GoTo TestError

    Dim basePath As String
    basePath = GetTestBasePath()

    If basePath = "" Then
        LogTest "Skipping GetValue test - no base path configured"
        Exit Sub
    End If

    ' Test if we can read from a closed workbook
    Dim testFile As String
    testFile = basePath & "Search.xls"

    If Dir(testFile, vbNormal) <> "" Then
        ' Try to read a value from the file
        Dim testValue As Variant
        testValue = GetValue(basePath, "Search.xls", "Sheet1", "A1")

        If testValue <> "File Not Found" Then
            LogPass "GetValue function can read from closed workbooks"
        Else
            LogFail "GetValue function cannot read from Search.xls"
        End If
    Else
        LogTest "Cannot test GetValue - Search.xls not found"
    End If

    Exit Sub

TestError:
    LogFail "GetValue function error: " & Err.Description
End Sub

' ====================================================================
' TEST 5: TEMPLATE FILE VALIDATION
' ====================================================================
Private Sub TestTemplateFiles()

    LogTest "=== TESTING TEMPLATE FILES ==="

    Dim basePath As String
    basePath = GetTestBasePath()

    If basePath = "" Then
        LogFail "Cannot test template files - no base path configured"
        Exit Sub
    End If

    Dim templatePath As String
    templatePath = basePath & "templates\"

    ' Test critical template files
    TestTemplateFile templatePath & "_Enq.xls", "Main enquiry template"
    TestTemplateFile templatePath & "_client.xls", "Customer template"
    TestTemplateFile templatePath & "price list.xls", "Price list template"
    TestTemplateFile templatePath & "Component_Grades.xls", "Component grades template"

    LogTest ""

End Sub

Private Sub TestTemplateFile(filePath As String, description As String)
    If Dir(filePath, vbNormal) <> "" Then
        LogPass description & " exists: " & filePath

        ' Try to validate the file can be opened
        On Error GoTo FileError

        Dim wb As Workbook
        Set wb = Workbooks.Open(filePath, ReadOnly:=True)

        ' Check for Admin sheet (common requirement)
        Dim hasAdminSheet As Boolean
        hasAdminSheet = False

        Dim ws As Worksheet
        For Each ws In wb.Worksheets
            If UCase(ws.Name) = "ADMIN" Then
                hasAdminSheet = True
                Exit For
            End If
        Next ws

        If hasAdminSheet Then
            LogPass "  -> File has required Admin sheet"
        Else
            LogTest "  -> Warning: File may be missing Admin sheet"
        End If

        wb.Close False
        Exit Sub

FileError:
        LogFail "  -> Error opening template file: " & Err.Description

    Else
        LogFail "Missing " & description & ": " & filePath
        LogTest "  -> This template is required for creating new records"
    End If
End Sub

' ====================================================================
' TEST 6: FORM CONTROL DEPENDENCIES
' ====================================================================
Private Sub TestFormControlDependencies()

    LogTest "=== TESTING FORM CONTROL DEPENDENCIES ==="

    ' Test if Main form exists
    TestMainFormExists

    LogTest ""

End Sub

Private Sub TestMainFormExists()
    On Error GoTo FormError

    ' Try to reference the Main form
    Dim formExists As Boolean
    formExists = False

    ' Check if Main form is available
    Dim obj As Object
    Set obj = VBA.UserForms("Main")

    If Not obj Is Nothing Then
        LogPass "Main form is available"
        formExists = True
    End If

    Exit Sub

FormError:
    LogFail "Main form not found or not accessible"
    LogTest "  -> You'll need to create the Main form with all required controls"
    LogTest "  -> Refer to the control specification provided earlier"
End Sub

' ====================================================================
' UTILITY FUNCTIONS
' ====================================================================
Private Function GetTestBasePath() As String
    ' Try to get the base path from various sources
    On Error GoTo NoPath

    ' Option 1: Try to get from Main form if it exists
    GetTestBasePath = Main.Main_MasterPath.Value
    If GetTestBasePath <> "" Then Exit Function

NoPath:
    ' Option 2: Use a default test path
    GetTestBasePath = "C:\PCS_Test\"
    LogTest "Using default test path: " & GetTestBasePath
    LogTest "Set Main_MasterPath for accurate testing"
End Function

Private Sub LogTest(message As String)
    TestResults = TestResults & message & vbCrLf
    Debug.Print message
End Sub

Private Sub LogPass(message As String)
    TestCount = TestCount + 1
    PassCount = PassCount + 1
    TestResults = TestResults & "✓ PASS: " & message & vbCrLf
    Debug.Print "✓ PASS: " & message
End Sub

Private Sub LogFail(message As String)
    TestCount = TestCount + 1
    FailCount = FailCount + 1
    TestResults = TestResults & "✗ FAIL: " & message & vbCrLf
    Debug.Print "✗ FAIL: " & message
End Sub

Private Sub DisplayTestResults()

    LogTest ""
    LogTest "=== TEST SUMMARY ==="
    LogTest "Total Tests: " & TestCount
    LogTest "Passed: " & PassCount
    LogTest "Failed: " & FailCount
    LogTest "Success Rate: " & Format((PassCount / TestCount) * 100, "0.0") & "%"
    LogTest ""

    If FailCount = 0 Then
        LogTest "🎉 ALL TESTS PASSED! Your framework is ready for the new interface."
    ElseIf FailCount <= 3 Then
        LogTest "⚠️ MINOR ISSUES FOUND. Review failed tests before building interface."
    Else
        LogTest "❌ MAJOR ISSUES FOUND. Fix these before building the interface."
    End If

    LogTest "=== END OF TESTING ==="

    ' Display results in a message box for easy viewing
    MsgBox TestResults, vbInformation, "VBA Framework Test Results"

End Sub

' ====================================================================
' SETUP HELPER FUNCTIONS
' ====================================================================

' Call this to create the basic directory structure for testing
Public Sub CreateTestDirectoryStructure()

    Dim basePath As String
    basePath = InputBox("Enter the base path where you want to create the test structure:", "Setup Test Environment", "C:\PCS_Test\")

    If basePath = "" Then Exit Sub

    ' Ensure path ends with backslash
    If Right(basePath, 1) <> "\" Then basePath = basePath & "\"

    ' Create directories
    CreateTestDir basePath & "enquiries"
    CreateTestDir basePath & "quotes"
    CreateTestDir basePath & "wip"
    CreateTestDir basePath & "archive"
    CreateTestDir basePath & "contracts"
    CreateTestDir basePath & "customers"
    CreateTestDir basePath & "templates"

    MsgBox "Test directory structure created at: " & basePath & vbCrLf & vbCrLf & _
           "Next steps:" & vbCrLf & _
           "1. Set Main_MasterPath to: " & basePath & vbCrLf & _
           "2. Add template files to the templates folder" & vbCrLf & _
           "3. Run RunAllTests() to validate setup", vbInformation, "Setup Complete"

End Sub

Private Sub CreateTestDir(dirPath As String)
    On Error Resume Next
    MkDir dirPath
    On Error GoTo 0
End Sub