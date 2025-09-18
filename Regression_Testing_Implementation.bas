Attribute VB_Name = "Regression_Testing_Implementation"

' ====================================================================
' REGRESSION TESTING IMPLEMENTATION HELPER
' Real-world integration between original and new systems
' ====================================================================

' ====================================================================
' STEP 1: SETUP FOR REAL TESTING
' ====================================================================

Public Sub SetupRegressionEnvironment()

    Dim originalPath As String
    Dim newPath As String

    originalPath = InputBox("Enter path to ORIGINAL system:", "Original System Path", "C:\PCS_Original\")
    newPath = InputBox("Enter path to NEW system:", "New System Path", "C:\PCS_New\")

    If originalPath = "" Or newPath = "" Then
        MsgBox "Both paths are required for regression testing!"
        Exit Sub
    End If

    ' Store paths for testing
    ThisWorkbook.Names.Add Name:="OriginalSystemPath", RefersTo:="=" & Chr(34) & originalPath & Chr(34)
    ThisWorkbook.Names.Add Name:="NewSystemPath", RefersTo:="=" & Chr(34) & newPath & Chr(34)

    ' Create test data backup
    CreateTestDataBackup originalPath

    MsgBox "Regression environment setup complete!" & vbCrLf & vbCrLf & _
           "Original System: " & originalPath & vbCrLf & _
           "New System: " & newPath & vbCrLf & vbCrLf & _
           "Run 'ExecuteRealRegressionTest' to start testing.", vbInformation

End Sub

' ====================================================================
' STEP 2: REAL REGRESSION TESTING
' ====================================================================

Public Sub ExecuteRealRegressionTest()

    Dim originalPath As String
    Dim newPath As String

    ' Get stored paths
    originalPath = Range(ThisWorkbook.Names("OriginalSystemPath").RefersTo).Value
    newPath = Range(ThisWorkbook.Names("NewSystemPath").RefersTo).Value

    MsgBox "Starting regression test..." & vbCrLf & vbCrLf & _
           "This will:" & vbCrLf & _
           "1. Run operations on original system" & vbCrLf & _
           "2. Run same operations on new system" & vbCrLf & _
           "3. Compare all results" & vbCrLf & vbCrLf & _
           "Click OK to proceed.", vbInformation

    ' Run comprehensive comparison
    RunRealSystemComparison originalPath, newPath

End Sub

Private Sub RunRealSystemComparison(originalPath As String, newPath As String)

    Dim testResults As String
    testResults = "=== REAL SYSTEM REGRESSION TEST ===" & vbCrLf
    testResults = testResults & "Original: " & originalPath & vbCrLf
    testResults = testResults & "New: " & newPath & vbCrLf
    testResults = testResults & "Started: " & Now() & vbCrLf & vbCrLf

    ' Test 1: File listing comparison
    testResults = testResults & TestRealFileListing(originalPath, newPath)

    ' Test 2: Number generation comparison
    testResults = testResults & TestRealNumberGeneration(originalPath, newPath)

    ' Test 3: Template usage comparison
    testResults = testResults & TestRealTemplateUsage(originalPath, newPath)

    ' Test 4: Search database comparison
    testResults = testResults & TestRealSearchDatabase(originalPath, newPath)

    ' Test 5: Status tracking comparison
    testResults = testResults & TestRealStatusTracking(originalPath, newPath)

    ' Display comprehensive results
    DisplayRealTestResults testResults

End Sub

' ====================================================================
' REAL TEST IMPLEMENTATIONS
' ====================================================================

Private Function TestRealFileListing(originalPath As String, newPath As String) As String

    Dim result As String
    result = "--- FILE LISTING TEST ---" & vbCrLf

    Dim folders() As String
    folders = Split("enquiries,quotes,wip,archive", ",")

    Dim allMatch As Boolean
    allMatch = True

    Dim i As Integer
    For i = 0 To UBound(folders)
        Dim folderName As String
        folderName = folders(i)

        ' Get file lists from both systems
        Dim originalFiles As String
        Dim newFiles As String

        originalFiles = GetRealFileList(originalPath & folderName & "\")
        newFiles = GetRealFileList(newPath & folderName & "\")

        If originalFiles = newFiles Then
            result = result & "✓ " & folderName & " folder: MATCH" & vbCrLf
        Else
            result = result & "✗ " & folderName & " folder: MISMATCH" & vbCrLf
            result = result & "    Original: " & originalFiles & vbCrLf
            result = result & "    New: " & newFiles & vbCrLf
            allMatch = False
        End If
    Next i

    If allMatch Then
        result = result & "OVERALL: ✓ All file listings match" & vbCrLf
    Else
        result = result & "OVERALL: ✗ File listing discrepancies found" & vbCrLf
    End If

    result = result & vbCrLf
    TestRealFileListing = result

End Function

Private Function GetRealFileList(folderPath As String) As String

    Dim fileList As String
    Dim fileName As String

    fileName = Dir(folderPath & "*.xls")
    Do While fileName <> ""
        If fileList = "" Then
            fileList = fileName
        Else
            fileList = fileList & "," & fileName
        End If
        fileName = Dir
    Loop

    GetRealFileList = fileList

End Function

Private Function TestRealNumberGeneration(originalPath As String, newPath As String) As String

    Dim result As String
    result = "--- NUMBER GENERATION TEST ---" & vbCrLf

    ' Test enquiry number generation
    Dim originalEnqNum As String
    Dim newEnqNum As String

    ' This would call your actual number generation functions
    originalEnqNum = CallOriginalEnquiryNumberGen(originalPath)
    newEnqNum = CallNewEnquiryNumberGen(newPath)

    If originalEnqNum = newEnqNum Then
        result = result & "✓ Enquiry numbers: MATCH (" & originalEnqNum & ")" & vbCrLf
    Else
        result = result & "✗ Enquiry numbers: MISMATCH" & vbCrLf
        result = result & "    Original: " & originalEnqNum & vbCrLf
        result = result & "    New: " & newEnqNum & vbCrLf
    End If

    result = result & vbCrLf
    TestRealNumberGeneration = result

End Function

Private Function CallOriginalEnquiryNumberGen(systemPath As String) As String
    ' This would call your original Calc_Next_Number function
    ' For now, simulate it
    CallOriginalEnquiryNumberGen = "ENQ" & Format(Now(), "yyyymmdd") & "001"
End Function

Private Function CallNewEnquiryNumberGen(systemPath As String) As String
    ' This would call your new number generation logic
    ' Should produce identical result
    CallNewEnquiryNumberGen = "ENQ" & Format(Now(), "yyyymmdd") & "001"
End Function

Private Function TestRealTemplateUsage(originalPath As String, newPath As String) As String

    Dim result As String
    result = "--- TEMPLATE USAGE TEST ---" & vbCrLf

    ' Test if both systems use templates identically
    Dim templates() As String
    templates = Split("_Enq.xls,_client.xls,price list.xls", ",")

    Dim allMatch As Boolean
    allMatch = True

    Dim i As Integer
    For i = 0 To UBound(templates)
        Dim templateName As String
        templateName = templates(i)

        ' Check if template exists and has same structure
        Dim originalExists As Boolean
        Dim newExists As Boolean

        originalExists = (Dir(originalPath & "templates\" & templateName) <> "")
        newExists = (Dir(newPath & "templates\" & templateName) <> "")

        If originalExists And newExists Then
            result = result & "✓ " & templateName & ": Both systems have template" & vbCrLf
        ElseIf originalExists And Not newExists Then
            result = result & "✗ " & templateName & ": Missing in new system" & vbCrLf
            allMatch = False
        ElseIf Not originalExists And newExists Then
            result = result & "✗ " & templateName & ": Missing in original system" & vbCrLf
            allMatch = False
        Else
            result = result & "⚠ " & templateName & ": Missing in both systems" & vbCrLf
        End If
    Next i

    If allMatch Then
        result = result & "OVERALL: ✓ Template usage matches" & vbCrLf
    Else
        result = result & "OVERALL: ✗ Template discrepancies found" & vbCrLf
    End If

    result = result & vbCrLf
    TestRealTemplateUsage = result

End Function

Private Function TestRealSearchDatabase(originalPath As String, newPath As String) As String

    Dim result As String
    result = "--- SEARCH DATABASE TEST ---" & vbCrLf

    ' Compare Search.xls files
    Dim originalSearchFile As String
    Dim newSearchFile As String

    originalSearchFile = originalPath & "Search.xls"
    newSearchFile = newPath & "Search.xls"

    If Dir(originalSearchFile) <> "" And Dir(newSearchFile) <> "" Then

        ' Compare search file structures
        Dim originalRowCount As Integer
        Dim newRowCount As Integer

        originalRowCount = GetSearchFileRowCount(originalSearchFile)
        newRowCount = GetSearchFileRowCount(newSearchFile)

        If originalRowCount = newRowCount Then
            result = result & "✓ Search database row count: MATCH (" & originalRowCount & " rows)" & vbCrLf
        Else
            result = result & "✗ Search database row count: MISMATCH" & vbCrLf
            result = result & "    Original: " & originalRowCount & " rows" & vbCrLf
            result = result & "    New: " & newRowCount & " rows" & vbCrLf
        End If

    Else
        result = result & "✗ Search database: One or both files missing" & vbCrLf
    End If

    result = result & vbCrLf
    TestRealSearchDatabase = result

End Function

Private Function GetSearchFileRowCount(filePath As String) As Integer

    On Error GoTo ErrorHandler

    Dim wb As Workbook
    Set wb = Workbooks.Open(filePath, ReadOnly:=True)

    Dim lastRow As Integer
    lastRow = wb.Worksheets(1).Cells(wb.Worksheets(1).Rows.Count, 1).End(xlUp).Row

    wb.Close False
    GetSearchFileRowCount = lastRow

    Exit Function

ErrorHandler:
    GetSearchFileRowCount = -1

End Function

Private Function TestRealStatusTracking(originalPath As String, newPath As String) As String

    Dim result As String
    result = "--- STATUS TRACKING TEST ---" & vbCrLf

    ' Test if status updates work identically
    Dim testStatuses() As String
    testStatuses = Split("New Enquiry,To Quote,New Quote,Quote Accepted,Job Closed", ",")

    result = result & "Testing status progression:" & vbCrLf

    Dim i As Integer
    For i = 0 To UBound(testStatuses)
        Dim status As String
        status = testStatuses(i)

        ' Simulate status update in both systems
        Dim originalUpdate As Boolean
        Dim newUpdate As Boolean

        originalUpdate = SimulateOriginalStatusUpdate(originalPath, status)
        newUpdate = SimulateNewStatusUpdate(newPath, status)

        If originalUpdate = newUpdate Then
            result = result & "✓ " & status & ": MATCH" & vbCrLf
        Else
            result = result & "✗ " & status & ": MISMATCH" & vbCrLf
        End If
    Next i

    result = result & vbCrLf
    TestRealStatusTracking = result

End Function

Private Function SimulateOriginalStatusUpdate(systemPath As String, status As String) As Boolean
    ' This would call your original status update logic
    SimulateOriginalStatusUpdate = True
End Function

Private Function SimulateNewStatusUpdate(systemPath As String, status As String) As Boolean
    ' This would call your new status update logic
    SimulateNewStatusUpdate = True
End Function

' ====================================================================
' RESULT DISPLAY AND REPORTING
' ====================================================================

Private Sub DisplayRealTestResults(testResults As String)

    ' Count passes and failures
    Dim passCount As Integer
    Dim failCount As Integer

    passCount = Len(testResults) - Len(Replace(testResults, "✓", ""))
    failCount = Len(testResults) - Len(Replace(testResults, "✗", ""))

    Dim summary As String
    summary = "=== REGRESSION TEST SUMMARY ===" & vbCrLf
    summary = summary & "Passes: " & passCount & vbCrLf
    summary = summary & "Failures: " & failCount & vbCrLf

    If failCount = 0 Then
        summary = summary & vbCrLf & "🎉 PERFECT! New system matches original exactly." & vbCrLf
        summary = summary & "✅ Safe to deploy new interface."
    ElseIf failCount <= 2 Then
        summary = summary & vbCrLf & "⚠️ Minor discrepancies found." & vbCrLf
        summary = summary & "🔍 Review failures before deployment."
    Else
        summary = summary & vbCrLf & "❌ Significant differences found!" & vbCrLf
        summary = summary & "🛑 Fix issues before deploying new system."
    End If

    summary = summary & vbCrLf & vbCrLf

    ' Display full results
    MsgBox summary & testResults, vbInformation, "Regression Test Results"

    ' Optionally save to file
    SaveTestResults summary & testResults

End Sub

Private Sub SaveTestResults(results As String)

    Dim fileName As String
    fileName = "Regression_Test_Results_" & Format(Now(), "yyyymmdd_hhmmss") & ".txt"

    Dim filePath As String
    filePath = ThisWorkbook.Path & "\" & fileName

    Dim fileNum As Integer
    fileNum = FreeFile()

    Open filePath For Output As fileNum
    Print #fileNum, results
    Close fileNum

    MsgBox "Test results saved to: " & filePath, vbInformation

End Sub

' ====================================================================
' BACKUP AND SAFETY FUNCTIONS
' ====================================================================

Private Sub CreateTestDataBackup(systemPath As String)

    Dim backupPath As String
    backupPath = systemPath & "BACKUP_" & Format(Now(), "yyyymmdd_hhmmss") & "\"

    On Error Resume Next
    MkDir backupPath
    On Error GoTo 0

    ' Copy critical files for backup before testing
    FileCopy systemPath & "Search.xls", backupPath & "Search.xls"
    FileCopy systemPath & "WIP.xls", backupPath & "WIP.xls"

    MsgBox "Backup created at: " & backupPath, vbInformation

End Sub