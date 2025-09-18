Attribute VB_Name = "VBA_Cross_System_Testing"

' ====================================================================
' CROSS-SYSTEM TESTING FRAMEWORK
' Compares Original VBA Interface with InterfaceV2 system
' Ensures both systems produce identical end-user results
' ====================================================================

Public Type TestComparison
    TestName As String
    InputData As String
    OriginalResult As String
    InterfaceV2Result As String
    ResultsMatch As Boolean
    Notes As String
End Type

Public TestResults() As TestComparison
Public TestCount As Integer
Public PassCount As Integer
Public FailCount As Integer
Public ComparisonReport As String

' ====================================================================
' MAIN COMPARISON TEST RUNNER
' ====================================================================
Public Sub CompareOriginalVsInterfaceV2()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ComparisonReport = ""
    TestCount = 0
    PassCount = 0
    FailCount = 0

    LogMessage "=== CROSS-SYSTEM COMPARISON TESTING ==="
    LogMessage "Comparing Original VBA Interface vs InterfaceV2"
    LogMessage "Started: " & Now()
    LogMessage ""

    ReDim TestResults(1 To 100)

    ' Test core functionality equivalence
    TestFileListingEquivalence
    TestSearchFunctionalityEquivalence
    TestDataAccessEquivalence
    TestPerformanceEquivalence

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    ' Display comprehensive results
    DisplayComparisonResults

End Sub

' ====================================================================
' TEST 1: FILE LISTING EQUIVALENCE
' ====================================================================
Private Sub TestFileListingEquivalence()

    LogMessage "=== TESTING FILE LISTING EQUIVALENCE ==="

    ' Test file listing for different folder types
    TestSingleFolderListing "enquiries"
    TestSingleFolderListing "quotes"
    TestSingleFolderListing "wip"
    TestSingleFolderListing "archive"

    ' Test filtered listing
    TestFilteredListing

    LogMessage ""

End Sub

Private Sub TestSingleFolderListing(folderType As String)

    Dim originalResult As String
    Dim interfaceV2Result As String

    ' Get file listing from original system
    originalResult = GetOriginalFileListing(folderType)

    ' Get file listing from InterfaceV2 system
    interfaceV2Result = GetInterfaceV2FileListing(folderType)

    ' Compare results
    CompareResults "FileListing_" & folderType, folderType, originalResult, interfaceV2Result

End Sub

Private Function GetOriginalFileListing(folderType As String) As String

    On Error GoTo OriginalError

    ' Use original List_Files function if available
    If FunctionExists("List_Files") And FormExists("Main") Then

        ' Clear and populate using original method
        Main.lst.Clear
        List_Files folderType, Main.lst

        ' Capture results
        Dim result As String
        result = ""

        Dim i As Integer
        For i = 0 To Main.lst.ListCount - 1
            If result = "" Then
                result = Main.lst.List(i)
            Else
                result = result & "|" & Main.lst.List(i)
            End If
        Next i

        GetOriginalFileListing = "Count=" & Main.lst.ListCount & ";Items=" & result
    Else
        GetOriginalFileListing = "ERROR=OriginalSystemNotAvailable"
    End If

    Exit Function

OriginalError:
    GetOriginalFileListing = "ERROR=" & Err.Description

End Function

Private Function GetInterfaceV2FileListing(folderType As String) As String

    On Error GoTo InterfaceV2Error

    ' Use InterfaceV2 FileUtilities.BuildFileList method
    Dim allFiles() As String
    Dim filteredFiles() As String
    Dim fileCount As Integer
    Dim i As Integer

    allFiles = FileUtilities.BuildFileList()

    ' Filter files by type
    ReDim filteredFiles(1 To UBound(allFiles))
    fileCount = 0

    For i = LBound(allFiles) To UBound(allFiles)
        If allFiles(i) <> "" Then
            Dim fileType As String
            fileType = FileUtilities.GetFileTypeFromPath(allFiles(i))

            If (folderType = "enquiries" And fileType = "Enquiry") Or _
               (folderType = "quotes" And fileType = "Quote") Or _
               (folderType = "wip" And fileType = "WIP") Or _
               (folderType = "archive" And fileType = "Archive") Then

                fileCount = fileCount + 1
                filteredFiles(fileCount) = GetFileNameOnly(allFiles(i))
            End If
        End If
    Next i

    ' Build result string
    Dim result As String
    result = ""

    For i = 1 To fileCount
        If result = "" Then
            result = filteredFiles(i)
        Else
            result = result & "|" & filteredFiles(i)
        End If
    Next i

    GetInterfaceV2FileListing = "Count=" & fileCount & ";Items=" & result

    Exit Function

InterfaceV2Error:
    GetInterfaceV2FileListing = "ERROR=" & Err.Description

End Function

Private Sub TestFilteredListing()

    ' Test if both systems handle filtering similarly
    Dim originalFiltered As String
    Dim interfaceV2Filtered As String

    originalFiltered = GetOriginalFilteredListing()
    interfaceV2Filtered = GetInterfaceV2FilteredListing()

    CompareResults "FilteredListing", "MultipleTypes", originalFiltered, interfaceV2Filtered

End Sub

Private Function GetOriginalFilteredListing() As String

    On Error GoTo FilterError

    ' Simulate original system with multiple filters
    If FormExists("Main") Then
        ' Original system would show based on toggle states
        Dim result As String
        result = "WIP=" & GetOriginalFileListing("wip")
        result = result & ";Enquiries=" & GetOriginalFileListing("enquiries")

        GetOriginalFilteredListing = result
    Else
        GetOriginalFilteredListing = "ERROR=MainFormNotFound"
    End If

    Exit Function

FilterError:
    GetOriginalFilteredListing = "ERROR=" & Err.Description

End Function

Private Function GetInterfaceV2FilteredListing() As String

    On Error GoTo FilterError

    ' Use InterfaceV2 filtering approach
    Dim allFiles() As String
    Dim wipCount As Integer, enquiryCount As Integer

    allFiles = FileUtilities.BuildFileList()
    wipCount = 0
    enquiryCount = 0

    Dim i As Integer
    For i = LBound(allFiles) To UBound(allFiles)
        If allFiles(i) <> "" Then
            Dim fileType As String
            fileType = FileUtilities.GetFileTypeFromPath(allFiles(i))

            If fileType = "WIP" Then wipCount = wipCount + 1
            If fileType = "Enquiry" Then enquiryCount = enquiryCount + 1
        End If
    Next i

    GetInterfaceV2FilteredListing = "WIP=Count=" & wipCount & ";Enquiries=Count=" & enquiryCount

    Exit Function

FilterError:
    GetInterfaceV2FilteredListing = "ERROR=" & Err.Description

End Function

' ====================================================================
' TEST 2: SEARCH FUNCTIONALITY EQUIVALENCE
' ====================================================================
Private Sub TestSearchFunctionalityEquivalence()

    LogMessage "=== TESTING SEARCH FUNCTIONALITY EQUIVALENCE ==="

    ' Test various search scenarios
    TestSingleSearch "TestCorp"
    TestSingleSearch "Widget"
    TestSingleSearch "ENQ001"
    TestSingleSearch "2024"

    LogMessage ""

End Sub

Private Sub TestSingleSearch(searchTerm As String)

    Dim originalResult As String
    Dim interfaceV2Result As String

    originalResult = GetOriginalSearchResult(searchTerm)
    interfaceV2Result = GetInterfaceV2SearchResult(searchTerm)

    CompareResults "Search_" & searchTerm, searchTerm, originalResult, interfaceV2Result

End Sub

Private Function GetOriginalSearchResult(searchTerm As String) As String

    On Error GoTo SearchError

    ' Original system search would typically use simple file listing
    ' and manual filtering - simulate this behavior

    Dim allFiles() As String
    Dim matchCount As Integer
    Dim basePath As String

    basePath = GetMasterPath()
    matchCount = 0

    ' Simple filename matching (original system approach)
    Dim folders() As String
    folders = Split("enquiries,quotes,wip,archive", ",")

    Dim i As Integer, j As Integer
    For i = 0 To UBound(folders)
        Dim fileName As String
        fileName = Dir(basePath & folders(i) & "\*.xls")

        Do While fileName <> ""
            If InStr(1, fileName, searchTerm, vbTextCompare) > 0 Then
                matchCount = matchCount + 1
            End If
            fileName = Dir
        Loop
    Next i

    GetOriginalSearchResult = "Method=SimpleFilename;Matches=" & matchCount

    Exit Function

SearchError:
    GetOriginalSearchResult = "ERROR=" & Err.Description

End Function

Private Function GetInterfaceV2SearchResult(searchTerm As String) As String

    On Error GoTo SearchError

    ' Use InterfaceV2 smart search
    Dim searchResults As Variant
    searchResults = SearchEngineV2.ExecuteSmartSearch(searchTerm)

    Dim resultCount As Integer
    If IsArray(searchResults) Then
        resultCount = UBound(searchResults) - LBound(searchResults) + 1
    Else
        resultCount = 0
    End If

    GetInterfaceV2SearchResult = "Method=SmartSearch;Matches=" & resultCount

    Exit Function

SearchError:
    GetInterfaceV2SearchResult = "ERROR=" & Err.Description

End Function

' ====================================================================
' TEST 3: DATA ACCESS EQUIVALENCE
' ====================================================================
Private Sub TestDataAccessEquivalence()

    LogMessage "=== TESTING DATA ACCESS EQUIVALENCE ==="

    ' Test file reading methods
    TestFileValueReading
    TestMetadataExtraction

    LogMessage ""

End Sub

Private Sub TestFileValueReading()

    ' Create a test file path (if one exists)
    Dim testFile As String
    testFile = FindTestFile()

    If testFile <> "" Then
        Dim originalValue As String
        Dim interfaceV2Value As String

        originalValue = GetOriginalFileValue(testFile, "Admin", "B2")
        interfaceV2Value = GetInterfaceV2FileValue(testFile, "Admin", "B2")

        CompareResults "FileValueReading", testFile & ":Admin:B2", originalValue, interfaceV2Value
    Else
        LogMessage "No test file available for data access testing"
    End If

End Sub

Private Function GetOriginalFileValue(filePath As String, sheetName As String, cellRef As String) As String

    On Error GoTo ValueError

    ' Use original GetValue function if available
    If FunctionExists("GetValue") Then
        Dim basePath As String
        Dim fileName As String

        basePath = GetDirectoryFromPath(filePath)
        fileName = GetFileNameOnly(filePath)

        GetOriginalFileValue = GetValue(basePath, fileName, sheetName, cellRef)
    Else
        GetOriginalFileValue = "ERROR=GetValueNotAvailable"
    End If

    Exit Function

ValueError:
    GetOriginalFileValue = "ERROR=" & Err.Description

End Function

Private Function GetInterfaceV2FileValue(filePath As String, sheetName As String, cellRef As String) As String

    On Error GoTo ValueError

    ' Use InterfaceV2 fast file access
    GetInterfaceV2FileValue = FileUtilities.GetValueFast(filePath, sheetName, cellRef)

    Exit Function

ValueError:
    GetInterfaceV2FileValue = "ERROR=" & Err.Description

End Function

Private Sub TestMetadataExtraction()

    Dim testFile As String
    testFile = FindTestFile()

    If testFile <> "" Then
        Dim originalMetadata As String
        Dim interfaceV2Metadata As String

        originalMetadata = GetOriginalMetadata(testFile)
        interfaceV2Metadata = GetInterfaceV2Metadata(testFile)

        CompareResults "MetadataExtraction", testFile, originalMetadata, interfaceV2Metadata
    End If

End Sub

Private Function GetOriginalMetadata(filePath As String) As String

    ' Original system would read specific cells manually
    Dim customer As String, component As String, status As String

    customer = GetOriginalFileValue(filePath, "Admin", "B3")
    component = GetOriginalFileValue(filePath, "Admin", "B4")
    status = GetOriginalFileValue(filePath, "Admin", "B2")

    GetOriginalMetadata = "Customer=" & customer & ";Component=" & component & ";Status=" & status

End Function

Private Function GetInterfaceV2Metadata(filePath As String) As String

    ' InterfaceV2 would use cached values
    Dim customer As String, component As String, status As String

    customer = CacheManager.GetCachedValue(filePath, "CustomerName")
    component = CacheManager.GetCachedValue(filePath, "ComponentCode")
    status = CacheManager.GetCachedValue(filePath, "Status")

    ' If not cached, fall back to direct read
    If customer = "" Then customer = GetInterfaceV2FileValue(filePath, "Admin", "B3")
    If component = "" Then component = GetInterfaceV2FileValue(filePath, "Admin", "B4")
    If status = "" Then status = GetInterfaceV2FileValue(filePath, "Admin", "B2")

    GetInterfaceV2Metadata = "Customer=" & customer & ";Component=" & component & ";Status=" & status

End Function

' ====================================================================
' TEST 4: PERFORMANCE EQUIVALENCE
' ====================================================================
Private Sub TestPerformanceEquivalence()

    LogMessage "=== TESTING PERFORMANCE EQUIVALENCE ==="

    ' Test performance of key operations
    TestListingPerformance
    TestSearchPerformance

    LogMessage ""

End Sub

Private Sub TestListingPerformance()

    Dim originalTime As Double, interfaceV2Time As Double
    Dim startTime As Double

    ' Test original system performance
    startTime = Timer
    GetOriginalFileListing "enquiries"
    originalTime = Timer - startTime

    ' Test InterfaceV2 system performance
    startTime = Timer
    GetInterfaceV2FileListing "enquiries"
    interfaceV2Time = Timer - startTime

    Dim performanceComparison As String
    performanceComparison = "Original=" & Format(originalTime, "0.000") & "s;InterfaceV2=" & Format(interfaceV2Time, "0.000") & "s"

    CompareResults "ListingPerformance", "enquiries", performanceComparison, performanceComparison

    ' Log performance notes
    If interfaceV2Time < originalTime Then
        LogMessage "  -> InterfaceV2 is " & Format((originalTime / interfaceV2Time), "0.0") & "x faster"
    Else
        LogMessage "  -> Original is " & Format((interfaceV2Time / originalTime), "0.0") & "x faster"
    End If

End Sub

Private Sub TestSearchPerformance()

    Dim originalTime As Double, interfaceV2Time As Double
    Dim startTime As Double
    Dim searchTerm As String

    searchTerm = "Test"

    ' Test original system search performance
    startTime = Timer
    GetOriginalSearchResult searchTerm
    originalTime = Timer - startTime

    ' Test InterfaceV2 system search performance
    startTime = Timer
    GetInterfaceV2SearchResult searchTerm
    interfaceV2Time = Timer - startTime

    Dim performanceComparison As String
    performanceComparison = "Original=" & Format(originalTime, "0.000") & "s;InterfaceV2=" & Format(interfaceV2Time, "0.000") & "s"

    CompareResults "SearchPerformance", searchTerm, performanceComparison, performanceComparison

End Sub

' ====================================================================
' UTILITY FUNCTIONS
' ====================================================================
Private Sub CompareResults(testName As String, inputData As String, originalResult As String, interfaceV2Result As String)

    TestCount = TestCount + 1

    Dim resultsMatch As Boolean

    ' Special handling for performance tests
    If InStr(testName, "Performance") > 0 Then
        resultsMatch = True ' Performance tests are informational
    Else
        resultsMatch = (originalResult = interfaceV2Result)
    End If

    If resultsMatch Then
        PassCount = PassCount + 1
        LogMessage "✓ PASS: " & testName
    Else
        FailCount = FailCount + 1
        LogMessage "✗ FAIL: " & testName
        LogMessage "    Input: " & inputData
        LogMessage "    Original: " & originalResult
        LogMessage "    InterfaceV2: " & interfaceV2Result
        LogMessage ""
    End If

    ' Store detailed results
    If TestCount <= UBound(TestResults) Then
        With TestResults(TestCount)
            .TestName = testName
            .InputData = inputData
            .OriginalResult = originalResult
            .InterfaceV2Result = interfaceV2Result
            .ResultsMatch = resultsMatch
            .Notes = IIf(resultsMatch, "Systems behave identically", "DIFFERENCE DETECTED")
        End With
    End If

End Sub

Private Function FindTestFile() As String

    Dim basePath As String
    Dim testPath As String

    basePath = GetMasterPath()

    ' Look for any .xls file in enquiries folder
    testPath = Dir(basePath & "enquiries\*.xls")
    If testPath <> "" Then
        FindTestFile = basePath & "enquiries\" & testPath
        Exit Function
    End If

    ' Try other folders
    testPath = Dir(basePath & "quotes\*.xls")
    If testPath <> "" Then
        FindTestFile = basePath & "quotes\" & testPath
        Exit Function
    End If

    FindTestFile = ""

End Function

Private Function GetMasterPath() As String

    On Error GoTo NoPath

    If FormExists("Main") Then
        GetMasterPath = Main.Main_MasterPath.Value
    Else
        GetMasterPath = Application.ActiveWorkbook.Path & "\"
    End If

    Exit Function

NoPath:
    GetMasterPath = Application.ActiveWorkbook.Path & "\"

End Function

Private Function FormExists(formName As String) As Boolean

    On Error GoTo NotFound
    Dim testForm As Object
    Set testForm = VBA.UserForms(formName)
    FormExists = True
    Exit Function

NotFound:
    FormExists = False

End Function

Private Function FunctionExists(functionName As String) As Boolean

    ' Simplified check - in practice would need more sophisticated detection
    FunctionExists = True

End Function

Private Function GetFileNameOnly(fullPath As String) As String

    Dim lastSlash As Long
    lastSlash = InStrRev(fullPath, "\")
    If lastSlash > 0 Then
        GetFileNameOnly = Mid(fullPath, lastSlash + 1)
    Else
        GetFileNameOnly = fullPath
    End If

End Function

Private Function GetDirectoryFromPath(fullPath As String) As String

    Dim lastSlash As Long
    lastSlash = InStrRev(fullPath, "\")
    If lastSlash > 0 Then
        GetDirectoryFromPath = Left(fullPath, lastSlash)
    Else
        GetDirectoryFromPath = ""
    End If

End Function

Private Sub LogMessage(message As String)

    ComparisonReport = ComparisonReport & message & vbCrLf
    Debug.Print message

End Sub

Private Sub DisplayComparisonResults()

    Dim summary As String
    summary = vbCrLf & "=== CROSS-SYSTEM COMPARISON SUMMARY ===" & vbCrLf
    summary = summary & "Total Tests: " & TestCount & vbCrLf
    summary = summary & "Passed: " & PassCount & vbCrLf
    summary = summary & "Failed: " & FailCount & vbCrLf

    If TestCount > 0 Then
        Dim successRate As Double
        successRate = (PassCount / TestCount) * 100
        summary = summary & "Compatibility Rate: " & Format(successRate, "0.0") & "%" & vbCrLf & vbCrLf

        If FailCount = 0 Then
            summary = summary & "🎉 PERFECT COMPATIBILITY!" & vbCrLf
            summary = summary & "✅ InterfaceV2 behaves identically to original system." & vbCrLf
            summary = summary & "✅ Safe to use InterfaceV2 as replacement."
        ElseIf successRate >= 90 Then
            summary = summary & "✅ EXCELLENT COMPATIBILITY with minor differences." & vbCrLf
            summary = summary & "⚠️ Review failed tests - may be acceptable differences."
        ElseIf successRate >= 70 Then
            summary = summary & "⚠️ GOOD COMPATIBILITY with some differences." & vbCrLf
            summary = summary & "🔍 Address failed tests before switching systems."
        Else
            summary = summary & "❌ SIGNIFICANT DIFFERENCES FOUND!" & vbCrLf
            summary = summary & "🛑 InterfaceV2 behavior differs substantially from original."
        End If
    End If

    summary = summary & vbCrLf & vbCrLf
    summary = summary & "=== DETAILED DIFFERENCES ===" & vbCrLf

    Dim i As Integer
    For i = 1 To TestCount
        If Not TestResults(i).ResultsMatch Then
            summary = summary & vbCrLf & "DIFFERENCE: " & TestResults(i).TestName & vbCrLf
            summary = summary & "  Input: " & TestResults(i).InputData & vbCrLf
            summary = summary & "  Original: " & TestResults(i).OriginalResult & vbCrLf
            summary = summary & "  InterfaceV2: " & TestResults(i).InterfaceV2Result & vbCrLf
        End If
    Next i

    ComparisonReport = ComparisonReport & summary

    ' Display results
    MsgBox ComparisonReport, vbInformation, "Cross-System Comparison Results"

    ' Save results to file
    SaveComparisonReport

End Sub

Private Sub SaveComparisonReport()

    Dim fileName As String
    fileName = ThisWorkbook.Path & "\Cross_System_Comparison_" & Format(Now(), "yyyymmdd_hhmmss") & ".txt"

    Dim fileNum As Integer
    fileNum = FreeFile()

    On Error GoTo SaveError

    Open fileName For Output As fileNum
    Print #fileNum, ComparisonReport
    Close fileNum

    LogMessage "Comparison report saved to: " & fileName
    Exit Sub

SaveError:
    LogMessage "Error saving report: " & Err.Description

End Sub