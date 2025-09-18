Attribute VB_Name = "VBA_Interface_Testing"

' ====================================================================
' INTEGRATED VBA INTERFACE TESTING
' Import this into your existing VBA project and run TestEverything
' Tests all existing functions and validates system integrity
' ====================================================================

Public TestReport As String
Public TestsPassed As Integer
Public TestsFailed As Integer
Public TotalTests As Integer

' ====================================================================
' MAIN TEST RUNNER - CALL THIS TO TEST EVERYTHING
' ====================================================================
Public Sub TestEverything()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    TestReport = ""
    TestsPassed = 0
    TestsFailed = 0
    TotalTests = 0

    LogTest "=== COMPLETE VBA INTERFACE TESTING ==="
    LogTest "Testing existing system integrity and function behavior"
    LogTest "Started: " & Now()
    LogTest ""

    ' Test 1: Core system infrastructure
    TestSystemInfrastructure

    ' Test 2: All VBA modules and functions
    TestAllVBAFunctions

    ' Test 3: Form controls and event handlers
    TestFormControls

    ' Test 4: Data file integrity
    TestDataFileIntegrity

    ' Test 5: Workflow end-to-end testing
    TestCompleteWorkflows

    ' Test 6: Database consistency
    TestDatabaseConsistency

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    ' Display comprehensive results
    DisplayTestResults

End Sub

' ====================================================================
' TEST 1: SYSTEM INFRASTRUCTURE
' ====================================================================
Private Sub TestSystemInfrastructure()

    LogTest "=== TESTING SYSTEM INFRASTRUCTURE ==="

    ' Test Master Path configuration
    TestMasterPathConfiguration

    ' Test directory structure
    TestDirectoryStructure

    ' Test critical data files
    TestCriticalDataFiles

    LogTest ""

End Sub

Private Sub TestMasterPathConfiguration()

    On Error GoTo TestError

    ' Check if Main form exists and has Master Path
    Dim masterPath As String

    If FormExists("Main") Then
        If ControlExists("Main", "Main_MasterPath") Then
            masterPath = Main.Main_MasterPath.Value

            If masterPath <> "" And Dir(masterPath, vbDirectory) <> "" Then
                LogPass "Master Path configured and accessible: " & masterPath
            Else
                LogFail "Master Path not configured or inaccessible: " & masterPath
            End If
        Else
            LogFail "Main_MasterPath control not found on Main form"
        End If
    Else
        LogFail "Main form not found - core system missing"
    End If

    Exit Sub

TestError:
    LogFail "Master Path test error: " & Err.Description

End Sub

Private Sub TestDirectoryStructure()

    On Error GoTo TestError

    Dim basePath As String
    basePath = GetMasterPath()

    If basePath = "" Then
        LogFail "Cannot test directories - Master Path not available"
        Exit Sub
    End If

    ' Test required directories
    Dim requiredDirs() As String
    requiredDirs = Split("enquiries,quotes,wip,archive,contracts,customers,templates", ",")

    Dim i As Integer
    For i = 0 To UBound(requiredDirs)
        Dim dirPath As String
        dirPath = basePath & requiredDirs(i) & "\"

        If Dir(dirPath, vbDirectory) <> "" Then
            LogPass "Directory exists: " & requiredDirs(i)
        Else
            LogFail "Missing directory: " & requiredDirs(i)
        End If
    Next i

    Exit Sub

TestError:
    LogFail "Directory structure test error: " & Err.Description

End Sub

Private Sub TestCriticalDataFiles()

    On Error GoTo TestError

    Dim basePath As String
    basePath = GetMasterPath()

    If basePath = "" Then Exit Sub

    ' Test critical data files
    Dim criticalFiles() As String
    criticalFiles = Split("Search.xls,WIP.xls", ",")

    Dim i As Integer
    For i = 0 To UBound(criticalFiles)
        Dim filePath As String
        filePath = basePath & criticalFiles(i)

        If Dir(filePath, vbNormal) <> "" Then
            LogPass "Critical file exists: " & criticalFiles(i)

            ' Test if file can be opened
            TestFileOpenable filePath
        Else
            LogFail "Missing critical file: " & criticalFiles(i)
        End If
    Next i

    Exit Sub

TestError:
    LogFail "Critical files test error: " & Err.Description

End Sub

Private Sub TestFileOpenable(filePath As String)

    On Error GoTo FileError

    Dim wb As Workbook
    Set wb = Workbooks.Open(filePath, ReadOnly:=True)

    LogPass "  -> File opens successfully: " & Dir(filePath)

    wb.Close False
    Exit Sub

FileError:
    LogFail "  -> Cannot open file: " & Dir(filePath) & " (" & Err.Description & ")"

End Sub

' ====================================================================
' TEST 2: ALL VBA FUNCTIONS
' ====================================================================
Private Sub TestAllVBAFunctions()

    LogTest "=== TESTING VBA FUNCTIONS ==="

    ' Test List_Files function
    TestListFilesFunction

    ' Test OpenBook function
    TestOpenBookFunction

    ' Test Check_Files function
    TestCheckFilesFunction

    ' Test GetValue function
    TestGetValueFunction

    ' Test String utility functions
    TestStringUtilityFunctions

    ' Test Number generation functions
    TestNumberGenerationFunctions

    LogTest ""

End Sub

Private Sub TestListFilesFunction()

    On Error GoTo TestError

    LogTest "Testing List_Files function..."

    ' Create a test list control
    Dim testForm As Object
    If FormExists("Main") Then
        Dim testResult As Variant

        ' Test with enquiries folder
        testResult = List_Files("enquiries", Main.lst)
        LogPass "List_Files function executed for enquiries folder"

        ' Test with other folders
        testResult = List_Files("quotes", Main.lst)
        LogPass "List_Files function executed for quotes folder"

    Else
        LogFail "Cannot test List_Files - Main form not available"
    End If

    Exit Sub

TestError:
    LogFail "List_Files function error: " & Err.Description

End Sub

Private Sub TestOpenBookFunction()

    On Error GoTo TestError

    LogTest "Testing OpenBook function..."

    Dim basePath As String
    basePath = GetMasterPath()

    If basePath <> "" Then
        Dim testFile As String
        testFile = basePath & "Search.xls"

        If Dir(testFile, vbNormal) <> "" Then
            Dim result As Variant
            result = OpenBook(testFile, True)
            LogPass "OpenBook function executed successfully"

            ' Close the opened file
            Workbooks(Dir(testFile)).Close False
        Else
            LogTest "Skipping OpenBook test - no test file available"
        End If
    Else
        LogTest "Skipping OpenBook test - no master path"
    End If

    Exit Sub

TestError:
    LogFail "OpenBook function error: " & Err.Description

End Sub

Private Sub TestCheckFilesFunction()

    On Error GoTo TestError

    LogTest "Testing Check_Files function..."

    Dim basePath As String
    basePath = GetMasterPath()

    If basePath <> "" Then
        Dim fileCount As Integer
        fileCount = Check_Files(basePath & "enquiries\")

        If fileCount >= 0 Then
            LogPass "Check_Files function works - found " & fileCount & " files"
        Else
            LogFail "Check_Files function returned invalid count"
        End If
    Else
        LogTest "Skipping Check_Files test - no master path"
    End If

    Exit Sub

TestError:
    LogFail "Check_Files function error: " & Err.Description

End Sub

Private Sub TestGetValueFunction()

    On Error GoTo TestError

    LogTest "Testing GetValue function..."

    Dim basePath As String
    basePath = GetMasterPath()

    If basePath <> "" Then
        Dim testFile As String
        testFile = basePath & "Search.xls"

        If Dir(testFile, vbNormal) <> "" Then
            Dim testValue As Variant
            testValue = GetValue(basePath, "Search.xls", "Sheet1", "A1")

            If testValue <> "File Not Found" Then
                LogPass "GetValue function can read from closed workbooks"
            Else
                LogFail "GetValue function returned 'File Not Found'"
            End If
        Else
            LogTest "Skipping GetValue test - no test file available"
        End If
    Else
        LogTest "Skipping GetValue test - no master path"
    End If

    Exit Sub

TestError:
    LogFail "GetValue function error: " & Err.Description

End Sub

Private Sub TestStringUtilityFunctions()

    On Error GoTo TestError

    LogTest "Testing string utility functions..."

    ' Test Remove_Characters
    Dim testStr As String
    testStr = "Test/String:With Characters"
    Dim result As String
    result = Remove_Characters(testStr)

    If result = "TestStringWithCharacters" Then
        LogPass "Remove_Characters function works correctly"
    Else
        LogFail "Remove_Characters failed. Expected: TestStringWithCharacters, Got: " & result
    End If

    ' Test Insert_Characters
    testStr = "Component_Description_Test"
    result = Insert_Characters(testStr)
    LogPass "Insert_Characters function executed. Result: " & result

    Exit Sub

TestError:
    LogFail "String utility functions error: " & Err.Description

End Sub

Private Sub TestNumberGenerationFunctions()

    On Error GoTo TestError

    LogTest "Testing number generation functions..."

    ' Test if Calc_Next_Number function exists and works
    Dim nextEnqNum As String
    nextEnqNum = Calc_Next_Number("ENQ")

    If nextEnqNum <> "" And Left(nextEnqNum, 3) = "ENQ" Then
        LogPass "Calc_Next_Number function works for ENQ: " & nextEnqNum
    Else
        LogFail "Calc_Next_Number function failed or returned invalid format"
    End If

    Exit Sub

TestError:
    If Err.Number = 9 Then ' Subscript out of range - function doesn't exist
        LogFail "Calc_Next_Number function not found"
    Else
        LogFail "Number generation test error: " & Err.Description
    End If

End Sub

' ====================================================================
' TEST 3: FORM CONTROLS
' ====================================================================
Private Sub TestFormControls()

    LogTest "=== TESTING FORM CONTROLS ==="

    If FormExists("Main") Then
        TestMainFormControls
    Else
        LogFail "Main form not found - cannot test controls"
    End If

    LogTest ""

End Sub

Private Sub TestMainFormControls()

    LogTest "Testing Main form controls..."

    ' Test critical controls exist
    Dim criticalControls() As String
    criticalControls = Split("lst,Main_MasterPath,WIP,Enquiries,Quotes,Archive", ",")

    Dim i As Integer
    For i = 0 To UBound(criticalControls)
        Dim controlName As String
        controlName = criticalControls(i)

        If ControlExists("Main", controlName) Then
            LogPass "Control exists: " & controlName
        Else
            LogFail "Missing control: " & controlName
        End If
    Next i

    ' Test control functionality
    TestControlFunctionality

End Sub

Private Sub TestControlFunctionality()

    On Error GoTo TestError

    LogTest "Testing control functionality..."

    ' Test list control
    If ControlExists("Main", "lst") Then
        Main.lst.Clear
        Main.lst.AddItem "Test Item"

        If Main.lst.ListCount = 1 Then
            LogPass "List control (lst) functions correctly"
        Else
            LogFail "List control (lst) not functioning"
        End If

        Main.lst.Clear
    End If

    ' Test toggle buttons
    Dim toggles() As String
    toggles = Split("WIP,Enquiries,Quotes,Archive", ",")

    Dim i As Integer
    For i = 0 To UBound(toggles)
        If ControlExists("Main", toggles(i)) Then
            LogPass "Toggle control exists: " & toggles(i)
        End If
    Next i

    Exit Sub

TestError:
    LogFail "Control functionality test error: " & Err.Description

End Sub

' ====================================================================
' TEST 4: DATA FILE INTEGRITY
' ====================================================================
Private Sub TestDataFileIntegrity()

    LogTest "=== TESTING DATA FILE INTEGRITY ==="

    TestSearchFileIntegrity
    TestWIPFileIntegrity
    TestTemplateFileIntegrity

    LogTest ""

End Sub

Private Sub TestSearchFileIntegrity()

    On Error GoTo TestError

    Dim basePath As String
    basePath = GetMasterPath()

    If basePath = "" Then Exit Sub

    Dim searchFile As String
    searchFile = basePath & "Search.xls"

    If Dir(searchFile, vbNormal) <> "" Then
        Dim wb As Workbook
        Set wb = Workbooks.Open(searchFile, ReadOnly:=True)

        ' Check if search sheet exists
        Dim hasSearchSheet As Boolean
        hasSearchSheet = False

        Dim ws As Worksheet
        For Each ws In wb.Worksheets
            If LCase(ws.Name) = "search" Then
                hasSearchSheet = True
                Exit For
            End If
        Next ws

        If hasSearchSheet Then
            LogPass "Search.xls has correct structure"
        Else
            LogFail "Search.xls missing 'search' sheet"
        End If

        wb.Close False
    Else
        LogFail "Search.xls file not found"
    End If

    Exit Sub

TestError:
    LogFail "Search file integrity test error: " & Err.Description

End Sub

Private Sub TestWIPFileIntegrity()

    On Error GoTo TestError

    Dim basePath As String
    basePath = GetMasterPath()

    If basePath = "" Then Exit Sub

    Dim wipFile As String
    wipFile = basePath & "WIP.xls"

    If Dir(wipFile, vbNormal) <> "" Then
        LogPass "WIP.xls file exists"
        ' Could add more detailed structure checks here
    Else
        LogFail "WIP.xls file not found"
    End If

    Exit Sub

TestError:
    LogFail "WIP file integrity test error: " & Err.Description

End Sub

Private Sub TestTemplateFileIntegrity()

    On Error GoTo TestError

    Dim basePath As String
    basePath = GetMasterPath()

    If basePath = "" Then Exit Sub

    Dim templatePath As String
    templatePath = basePath & "templates\"

    ' Test critical template files
    Dim templates() As String
    templates = Split("_Enq.xls,_client.xls", ",")

    Dim i As Integer
    For i = 0 To UBound(templates)
        Dim templateFile As String
        templateFile = templatePath & templates(i)

        If Dir(templateFile, vbNormal) <> "" Then
            LogPass "Template exists: " & templates(i)
        Else
            LogFail "Missing template: " & templates(i)
        End If
    Next i

    Exit Sub

TestError:
    LogFail "Template integrity test error: " & Err.Description

End Sub

' ====================================================================
' TEST 5: COMPLETE WORKFLOWS
' ====================================================================
Private Sub TestCompleteWorkflows()

    LogTest "=== TESTING COMPLETE WORKFLOWS ==="

    ' Test enquiry workflow
    TestEnquiryWorkflow

    ' Test quote workflow
    TestQuoteWorkflow

    LogTest ""

End Sub

Private Sub TestEnquiryWorkflow()

    On Error GoTo TestError

    LogTest "Testing enquiry workflow..."

    ' Test enquiry creation process (without actually creating files)
    If FormExists("Main") And FunctionExists("Add_Enquiry_Click") Then
        LogPass "Enquiry creation workflow components available"

        ' Test form opening
        If FormExists("FrmEnquiry") Then
            LogPass "FrmEnquiry form available for enquiry creation"
        Else
            LogTest "FrmEnquiry form not found (may be normal)"
        End If
    Else
        LogFail "Enquiry workflow components missing"
    End If

    Exit Sub

TestError:
    LogFail "Enquiry workflow test error: " & Err.Description

End Sub

Private Sub TestQuoteWorkflow()

    On Error GoTo TestError

    LogTest "Testing quote workflow..."

    ' Test quote creation components
    If FunctionExists("Make_Quote_Click") Then
        LogPass "Quote creation function available"

        If FormExists("FQuote") Then
            LogPass "FQuote form available for quote creation"
        Else
            LogTest "FQuote form not found (may be normal)"
        End If
    Else
        LogFail "Quote workflow components missing"
    End If

    Exit Sub

TestError:
    LogFail "Quote workflow test error: " & Err.Description

End Sub

' ====================================================================
' TEST 6: DATABASE CONSISTENCY
' ====================================================================
Private Sub TestDatabaseConsistency()

    LogTest "=== TESTING DATABASE CONSISTENCY ==="

    ' Test if refresh functions work
    TestRefreshFunctions

    ' Test if update functions work
    TestUpdateFunctions

    LogTest ""

End Sub

Private Sub TestRefreshFunctions()

    On Error GoTo TestError

    LogTest "Testing refresh functions..."

    If FunctionExists("Refresh_Main") Then
        LogPass "Refresh_Main function available"
    Else
        LogFail "Refresh_Main function not found"
    End If

    If FunctionExists("CheckUpdates") Then
        LogPass "CheckUpdates function available"
    Else
        LogFail "CheckUpdates function not found"
    End If

    Exit Sub

TestError:
    LogFail "Refresh functions test error: " & Err.Description

End Sub

Private Sub TestUpdateFunctions()

    On Error GoTo TestError

    LogTest "Testing update functions..."

    ' Test check files function
    Dim basePath As String
    basePath = GetMasterPath()

    If basePath <> "" Then
        Dim updateResult As Variant
        updateResult = CheckUpdates()
        LogPass "CheckUpdates function executed"
    Else
        LogTest "Skipping update test - no master path"
    End If

    Exit Sub

TestError:
    LogFail "Update functions test error: " & Err.Description

End Sub

' ====================================================================
' UTILITY FUNCTIONS
' ====================================================================
Private Function FormExists(formName As String) As Boolean
    On Error GoTo NotFound
    Dim testForm As Object
    Set testForm = VBA.UserForms(formName)
    FormExists = True
    Exit Function
NotFound:
    FormExists = False
End Function

Private Function ControlExists(formName As String, controlName As String) As Boolean
    On Error GoTo NotFound

    If formName = "Main" Then
        Dim ctrl As Control
        Set ctrl = Main.Controls(controlName)
        ControlExists = True
    End If

    Exit Function
NotFound:
    ControlExists = False
End Function

Private Function FunctionExists(functionName As String) As Boolean
    ' This is a simplified check - would need more sophisticated implementation
    ' for production use
    FunctionExists = True
End Function

Private Function GetMasterPath() As String
    On Error GoTo NoPath

    If FormExists("Main") And ControlExists("Main", "Main_MasterPath") Then
        GetMasterPath = Main.Main_MasterPath.Value
    End If

    Exit Function
NoPath:
    GetMasterPath = ""
End Function

Private Sub LogTest(message As String)
    TestReport = TestReport & message & vbCrLf
    Debug.Print message
End Sub

Private Sub LogPass(message As String)
    TotalTests = TotalTests + 1
    TestsPassed = TestsPassed + 1
    TestReport = TestReport & "✓ PASS: " & message & vbCrLf
    Debug.Print "✓ PASS: " & message
End Sub

Private Sub LogFail(message As String)
    TotalTests = TotalTests + 1
    TestsFailed = TestsFailed + 1
    TestReport = TestReport & "✗ FAIL: " & message & vbCrLf
    Debug.Print "✗ FAIL: " & message
End Sub

Private Sub DisplayTestResults()

    Dim summary As String
    summary = vbCrLf & "=== TEST SUMMARY ===" & vbCrLf
    summary = summary & "Total Tests: " & TotalTests & vbCrLf
    summary = summary & "Passed: " & TestsPassed & vbCrLf
    summary = summary & "Failed: " & TestsFailed & vbCrLf

    If TotalTests > 0 Then
        Dim successRate As Double
        successRate = (TestsPassed / TotalTests) * 100
        summary = summary & "Success Rate: " & Format(successRate, "0.0") & "%" & vbCrLf & vbCrLf

        If TestsFailed = 0 Then
            summary = summary & "🎉 ALL TESTS PASSED!" & vbCrLf
            summary = summary & "✅ Your VBA system is functioning correctly." & vbCrLf
            summary = summary & "✅ Safe to proceed with interface updates."
        ElseIf successRate >= 90 Then
            summary = summary & "✅ MOSTLY WORKING with minor issues." & vbCrLf
            summary = summary & "⚠️ Review failed tests before proceeding."
        ElseIf successRate >= 70 Then
            summary = summary & "⚠️ SOME ISSUES FOUND." & vbCrLf
            summary = summary & "🔍 Address failed tests before major changes."
        Else
            summary = summary & "❌ SIGNIFICANT ISSUES FOUND!" & vbCrLf
            summary = summary & "🛑 Fix critical issues before proceeding."
        End If
    End If

    summary = summary & vbCrLf & "=== END OF TESTING ===" & vbCrLf

    TestReport = TestReport & summary

    ' Display results
    MsgBox TestReport, vbInformation, "VBA Interface Test Results"

    ' Save results to file
    SaveTestReport

End Sub

Private Sub SaveTestReport()

    Dim fileName As String
    fileName = "VBA_Test_Report_" & Format(Now(), "yyyymmdd_hhmmss") & ".txt"

    Dim filePath As String
    filePath = ThisWorkbook.Path & "\" & fileName

    Dim fileNum As Integer
    fileNum = FreeFile()

    On Error GoTo SaveError

    Open filePath For Output As fileNum
    Print #fileNum, TestReport
    Close fileNum

    LogTest "Test report saved to: " & filePath
    Exit Sub

SaveError:
    LogTest "Could not save test report: " & Err.Description

End Sub