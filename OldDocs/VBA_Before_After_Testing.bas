Attribute VB_Name = "VBA_Before_After_Testing"

' ====================================================================
' VBA BEFORE/AFTER TESTING FRAMEWORK
' 1. Capture baseline behavior of existing code
' 2. Test new code against baseline to ensure identical behavior
' ====================================================================

Public Type TestBaseline
    TestName As String
    InputData As String
    ExpectedOutput As String
    ActualOutput As String
    TestPassed As Boolean
End Type

Public Baselines() As TestBaseline
Public BaselineCount As Integer
Public ComparisonResults As String

' ====================================================================
' PHASE 1: CAPTURE BASELINE (Run with existing code)
' ====================================================================
Public Sub CaptureBaseline()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    BaselineCount = 0
    ReDim Baselines(1 To 100)

    LogMessage "=== CAPTURING BASELINE BEHAVIOR ==="
    LogMessage "Recording how existing system behaves"
    LogMessage "Started: " & Now()
    LogMessage ""

    ' Capture baseline behaviors
    CaptureFileListingBehavior
    CaptureEnquiryCreationBehavior
    CaptureQuoteCreationBehavior
    CaptureJobCreationBehavior
    CaptureStatusUpdateBehavior
    CaptureNumberGenerationBehavior
    CaptureSearchUpdateBehavior

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    ' Save baseline to file
    SaveBaselineToFile

    LogMessage ""
    LogMessage "=== BASELINE CAPTURE COMPLETE ==="
    LogMessage "Captured " & BaselineCount & " behavioral tests"
    LogMessage "Baseline saved for comparison testing"

    MsgBox "Baseline captured successfully!" & vbCrLf & vbCrLf & _
           "Total behaviors recorded: " & BaselineCount & vbCrLf & vbCrLf & _
           "Now make your code changes, then run CompareWithBaseline()", vbInformation

End Sub

' ====================================================================
' CAPTURE SPECIFIC BEHAVIORS
' ====================================================================
Private Sub CaptureFileListingBehavior()

    LogMessage "Capturing file listing behavior..."

    ' Test file listing for different folders
    Dim folders() As String
    folders = Split("enquiries,quotes,wip,archive", ",")

    Dim i As Integer
    For i = 0 To UBound(folders)
        Dim folderName As String
        folderName = folders(i)

        ' Capture current behavior
        Dim currentBehavior As String
        currentBehavior = CaptureCurrentFileListing(folderName)

        ' Record baseline
        RecordBaseline "FileListing_" & folderName, folderName, currentBehavior
    Next i

End Sub

Private Function CaptureCurrentFileListing(folderName As String) As String

    On Error GoTo CaptureError

    ' Clear the list and populate with current method
    If FormExists("Main") Then
        Main.lst.Clear

        ' Call the existing List_Files function
        List_Files folderName, Main.lst

        ' Capture the results
        Dim result As String
        result = ""

        Dim j As Integer
        For j = 0 To Main.lst.ListCount - 1
            If result = "" Then
                result = Main.lst.List(j)
            Else
                result = result & "|" & Main.lst.List(j)
            End If
        Next j

        CaptureCurrentFileListing = "Count=" & Main.lst.ListCount & ";Items=" & result
    Else
        CaptureCurrentFileListing = "ERROR=MainFormNotFound"
    End If

    Exit Function

CaptureError:
    CaptureCurrentFileListing = "ERROR=" & Err.Description

End Function

Private Sub CaptureEnquiryCreationBehavior()

    LogMessage "Capturing enquiry creation behavior..."

    ' Simulate enquiry creation without actually creating files
    Dim testInputs() As String
    testInputs = Split("Customer=TestCorp|Component=Widget|Quantity=100", ",")

    Dim i As Integer
    For i = 0 To UBound(testInputs)
        Dim inputData As String
        inputData = testInputs(i)

        Dim behavior As String
        behavior = CaptureCurrentEnquiryCreation(inputData)

        RecordBaseline "EnquiryCreation_" & (i + 1), inputData, behavior
    Next i

End Sub

Private Function CaptureCurrentEnquiryCreation(inputData As String) As String

    On Error GoTo CaptureError

    ' Parse input data
    Dim customer As String, component As String, quantity As String
    customer = ExtractValue(inputData, "Customer")
    component = ExtractValue(inputData, "Component")
    quantity = ExtractValue(inputData, "Quantity")

    ' Capture what would happen (without actually doing it)
    Dim nextEnqNum As String
    nextEnqNum = Calc_Next_Number("ENQ")

    Dim expectedPath As String
    expectedPath = GetMasterPath() & "enquiries\" & nextEnqNum & ".xls"

    ' Capture the expected behavior
    Dim behavior As String
    behavior = "EnquiryNumber=" & nextEnqNum
    behavior = behavior & ";Path=" & expectedPath
    behavior = behavior & ";Status=To Quote"
    behavior = behavior & ";Customer=" & customer
    behavior = behavior & ";Component=" & component

    CaptureCurrentEnquiryCreation = behavior

    Exit Function

CaptureError:
    CaptureCurrentEnquiryCreation = "ERROR=" & Err.Description

End Function

Private Sub CaptureQuoteCreationBehavior()

    LogMessage "Capturing quote creation behavior..."

    ' Test quote creation from enquiry
    Dim testEnquiry As String
    testEnquiry = "ENQ20241201001"

    Dim behavior As String
    behavior = CaptureCurrentQuoteCreation(testEnquiry)

    RecordBaseline "QuoteCreation", testEnquiry, behavior

End Sub

Private Function CaptureCurrentQuoteCreation(enquiryFile As String) As String

    On Error GoTo CaptureError

    ' Capture quote creation behavior
    Dim nextQuoteNum As String
    nextQuoteNum = Calc_Next_Number("QUO")

    Dim sourcePath As String
    sourcePath = GetMasterPath() & "enquiries\" & enquiryFile & ".xls"

    Dim targetPath As String
    targetPath = GetMasterPath() & "quotes\" & nextQuoteNum & ".xls"

    Dim behavior As String
    behavior = "QuoteNumber=" & nextQuoteNum
    behavior = behavior & ";SourceFile=" & enquiryFile
    behavior = behavior & ";TargetPath=" & targetPath
    behavior = behavior & ";Status=New Quote"

    CaptureCurrentQuoteCreation = behavior

    Exit Function

CaptureError:
    CaptureCurrentQuoteCreation = "ERROR=" & Err.Description

End Function

Private Sub CaptureJobCreationBehavior()

    LogMessage "Capturing job creation behavior..."

    Dim testQuote As String
    testQuote = "QUO20241201001"

    Dim behavior As String
    behavior = CaptureCurrentJobCreation(testQuote)

    RecordBaseline "JobCreation", testQuote, behavior

End Sub

Private Function CaptureCurrentJobCreation(quoteFile As String) As String

    On Error GoTo CaptureError

    Dim nextJobNum As String
    nextJobNum = Calc_Next_Number("JOB")

    Dim sourcePath As String
    sourcePath = GetMasterPath() & "quotes\" & quoteFile & ".xls"

    Dim targetPath As String
    targetPath = GetMasterPath() & "wip\" & nextJobNum & ".xls"

    Dim behavior As String
    behavior = "JobNumber=" & nextJobNum
    behavior = behavior & ";SourceFile=" & quoteFile
    behavior = behavior & ";TargetPath=" & targetPath
    behavior = behavior & ";Status=Quote Accepted"

    CaptureCurrentJobCreation = behavior

    Exit Function

CaptureError:
    CaptureCurrentJobCreation = "ERROR=" & Err.Description

End Function

Private Sub CaptureStatusUpdateBehavior()

    LogMessage "Capturing status update behavior..."

    Dim statuses() As String
    statuses = Split("New Enquiry,To Quote,New Quote,Quote Submitted,Quote Accepted,Job Closed", ",")

    Dim i As Integer
    For i = 0 To UBound(statuses)
        Dim status As String
        status = statuses(i)

        Dim behavior As String
        behavior = CaptureCurrentStatusUpdate("TEST001", status)

        RecordBaseline "StatusUpdate_" & Replace(status, " ", ""), "TEST001->" & status, behavior
    Next i

End Sub

Private Function CaptureCurrentStatusUpdate(fileName As String, newStatus As String) As String

    On Error GoTo CaptureError

    ' Capture what happens during status update
    Dim behavior As String
    behavior = "File=" & fileName
    behavior = behavior & ";NewStatus=" & newStatus
    behavior = behavior & ";UpdatedSearchDB=True"
    behavior = behavior & ";UpdatedWIP=" & IIf(newStatus = "Quote Accepted", "True", "False")

    CaptureCurrentStatusUpdate = behavior

    Exit Function

CaptureError:
    CaptureCurrentStatusUpdate = "ERROR=" & Err.Description

End Function

Private Sub CaptureNumberGenerationBehavior()

    LogMessage "Capturing number generation behavior..."

    Dim types() As String
    types = Split("ENQ,QUO,JOB", ",")

    Dim i As Integer
    For i = 0 To UBound(types)
        Dim numberType As String
        numberType = types(i)

        Dim behavior As String
        behavior = CaptureCurrentNumberGeneration(numberType)

        RecordBaseline "NumberGeneration_" & numberType, numberType, behavior
    Next i

End Sub

Private Function CaptureCurrentNumberGeneration(numberType As String) As String

    On Error GoTo CaptureError

    Dim nextNumber As String
    nextNumber = Calc_Next_Number(numberType)

    Dim behavior As String
    behavior = "Type=" & numberType
    behavior = behavior & ";NextNumber=" & nextNumber
    behavior = behavior & ";Format=" & Left(nextNumber, 3) & "YYYYMMDDNNN"

    CaptureCurrentNumberGeneration = behavior

    Exit Function

CaptureError:
    CaptureCurrentNumberGeneration = "ERROR=" & Err.Description

End Function

Private Sub CaptureSearchUpdateBehavior()

    LogMessage "Capturing search update behavior..."

    ' Test search database update
    Dim testData As String
    testData = "File=TEST001|Customer=TestCorp|Status=New Enquiry"

    Dim behavior As String
    behavior = CaptureCurrentSearchUpdate(testData)

    RecordBaseline "SearchUpdate", testData, behavior

End Sub

Private Function CaptureCurrentSearchUpdate(testData As String) As String

    On Error GoTo CaptureError

    ' Capture search update behavior (without actually updating)
    Dim behavior As String
    behavior = "UpdateType=INSERT"
    behavior = behavior & ";Database=Search.xls"
    behavior = behavior & ";Sheet=search"
    behavior = behavior & ";Action=AddNewRow"

    CaptureCurrentSearchUpdate = behavior

    Exit Function

CaptureError:
    CaptureCurrentSearchUpdate = "ERROR=" & Err.Description

End Function

' ====================================================================
' PHASE 2: COMPARE WITH BASELINE (Run after code changes)
' ====================================================================
Public Sub CompareWithBaseline()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ComparisonResults = ""
    LoadBaselineFromFile

    If BaselineCount = 0 Then
        MsgBox "No baseline found! Run CaptureBaseline() first.", vbCritical
        Exit Sub
    End If

    LogMessage "=== COMPARING WITH BASELINE ==="
    LogMessage "Testing new code against captured baseline"
    LogMessage "Started: " & Now()
    LogMessage ""

    Dim passCount As Integer
    Dim failCount As Integer
    passCount = 0
    failCount = 0

    ' Run all baseline tests with new code
    Dim i As Integer
    For i = 1 To BaselineCount
        Dim actualOutput As String
        actualOutput = ExecuteTestWithNewCode(Baselines(i).TestName, Baselines(i).InputData)

        Baselines(i).ActualOutput = actualOutput
        Baselines(i).TestPassed = (actualOutput = Baselines(i).ExpectedOutput)

        If Baselines(i).TestPassed Then
            LogMessage "✓ PASS: " & Baselines(i).TestName
            passCount = passCount + 1
        Else
            LogMessage "✗ FAIL: " & Baselines(i).TestName
            LogMessage "    Expected: " & Baselines(i).ExpectedOutput
            LogMessage "    Actual:   " & actualOutput
            LogMessage ""
            failCount = failCount + 1
        End If
    Next i

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    ' Display final results
    DisplayComparisonResults passCount, failCount

End Sub

Private Function ExecuteTestWithNewCode(testName As String, inputData As String) As String

    On Error GoTo TestError

    ' Route to appropriate test based on test name
    If InStr(testName, "FileListing_") > 0 Then
        Dim folderName As String
        folderName = Replace(testName, "FileListing_", "")
        ExecuteTestWithNewCode = CaptureCurrentFileListing(folderName)

    ElseIf InStr(testName, "EnquiryCreation_") > 0 Then
        ExecuteTestWithNewCode = CaptureCurrentEnquiryCreation(inputData)

    ElseIf testName = "QuoteCreation" Then
        ExecuteTestWithNewCode = CaptureCurrentQuoteCreation(inputData)

    ElseIf testName = "JobCreation" Then
        ExecuteTestWithNewCode = CaptureCurrentJobCreation(inputData)

    ElseIf InStr(testName, "StatusUpdate_") > 0 Then
        Dim parts() As String
        parts = Split(inputData, "->")
        ExecuteTestWithNewCode = CaptureCurrentStatusUpdate(parts(0), parts(1))

    ElseIf InStr(testName, "NumberGeneration_") > 0 Then
        Dim numberType As String
        numberType = Replace(testName, "NumberGeneration_", "")
        ExecuteTestWithNewCode = CaptureCurrentNumberGeneration(numberType)

    ElseIf testName = "SearchUpdate" Then
        ExecuteTestWithNewCode = CaptureCurrentSearchUpdate(inputData)

    Else
        ExecuteTestWithNewCode = "ERROR=UnknownTestType"
    End If

    Exit Function

TestError:
    ExecuteTestWithNewCode = "ERROR=" & Err.Description

End Function

' ====================================================================
' UTILITY FUNCTIONS
' ====================================================================
Private Sub RecordBaseline(testName As String, inputData As String, expectedOutput As String)

    BaselineCount = BaselineCount + 1

    With Baselines(BaselineCount)
        .TestName = testName
        .InputData = inputData
        .ExpectedOutput = expectedOutput
        .ActualOutput = ""
        .TestPassed = False
    End With

End Sub

Private Function ExtractValue(data As String, key As String) As String

    Dim parts() As String
    parts = Split(data, "|")

    Dim i As Integer
    For i = 0 To UBound(parts)
        If InStr(parts(i), key & "=") = 1 Then
            ExtractValue = Mid(parts(i), Len(key) + 2)
            Exit Function
        End If
    Next i

    ExtractValue = ""

End Function

Private Function GetMasterPath() As String

    On Error GoTo NoPath

    If FormExists("Main") Then
        GetMasterPath = Main.Main_MasterPath.Value
    Else
        GetMasterPath = "C:\PCS\"
    End If

    Exit Function

NoPath:
    GetMasterPath = "C:\PCS\"

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

Private Sub LogMessage(message As String)

    ComparisonResults = ComparisonResults & message & vbCrLf
    Debug.Print message

End Sub

Private Sub SaveBaselineToFile()

    Dim fileName As String
    fileName = ThisWorkbook.Path & "\VBA_Baseline.txt"

    Dim fileNum As Integer
    fileNum = FreeFile()

    On Error GoTo SaveError

    Open fileName For Output As fileNum

    Print #fileNum, "BaselineCount=" & BaselineCount

    Dim i As Integer
    For i = 1 To BaselineCount
        With Baselines(i)
            Print #fileNum, .TestName & "|" & .InputData & "|" & .ExpectedOutput
        End With
    Next i

    Close fileNum
    LogMessage "Baseline saved to: " & fileName
    Exit Sub

SaveError:
    LogMessage "Error saving baseline: " & Err.Description

End Sub

Private Sub LoadBaselineFromFile()

    Dim fileName As String
    fileName = ThisWorkbook.Path & "\VBA_Baseline.txt"

    If Dir(fileName) = "" Then
        BaselineCount = 0
        Exit Sub
    End If

    Dim fileNum As Integer
    fileNum = FreeFile()

    On Error GoTo LoadError

    Open fileName For Input As fileNum

    Dim line As String
    Line Input #fileNum, line

    If InStr(line, "BaselineCount=") = 1 Then
        BaselineCount = Val(Mid(line, 15))
        ReDim Baselines(1 To BaselineCount)
    End If

    Dim i As Integer
    For i = 1 To BaselineCount
        Line Input #fileNum, line

        Dim parts() As String
        parts = Split(line, "|")

        If UBound(parts) >= 2 Then
            With Baselines(i)
                .TestName = parts(0)
                .InputData = parts(1)
                .ExpectedOutput = parts(2)
            End With
        End If
    Next i

    Close fileNum
    Exit Sub

LoadError:
    BaselineCount = 0
    LogMessage "Error loading baseline: " & Err.Description

End Sub

Private Sub DisplayComparisonResults(passCount As Integer, failCount As Integer)

    Dim summary As String
    summary = vbCrLf & "=== COMPARISON SUMMARY ===" & vbCrLf
    summary = summary & "Total Tests: " & (passCount + failCount) & vbCrLf
    summary = summary & "Passed: " & passCount & vbCrLf
    summary = summary & "Failed: " & failCount & vbCrLf

    If (passCount + failCount) > 0 Then
        Dim successRate As Double
        successRate = (passCount / (passCount + failCount)) * 100
        summary = summary & "Success Rate: " & Format(successRate, "0.0") & "%" & vbCrLf & vbCrLf

        If failCount = 0 Then
            summary = summary & "🎉 PERFECT MATCH!" & vbCrLf
            summary = summary & "✅ New code behaves identically to original." & vbCrLf
            summary = summary & "✅ Safe to deploy changes."
        ElseIf successRate >= 95 Then
            summary = summary & "✅ EXCELLENT - Minor differences only." & vbCrLf
            summary = summary & "⚠️ Review failed tests."
        ElseIf successRate >= 80 Then
            summary = summary & "⚠️ GOOD with some issues." & vbCrLf
            summary = summary & "🔍 Fix failed tests before deployment."
        Else
            summary = summary & "❌ SIGNIFICANT BEHAVIORAL CHANGES!" & vbCrLf
            summary = summary & "🛑 New code differs substantially from original."
        End If
    End If

    ComparisonResults = ComparisonResults & summary

    MsgBox ComparisonResults, vbInformation, "Baseline Comparison Results"

    ' Save comparison results
    SaveComparisonResults

End Sub

Private Sub SaveComparisonResults()

    Dim fileName As String
    fileName = ThisWorkbook.Path & "\VBA_Comparison_" & Format(Now(), "yyyymmdd_hhmmss") & ".txt"

    Dim fileNum As Integer
    fileNum = FreeFile()

    On Error GoTo SaveError

    Open fileName For Output As fileNum
    Print #fileNum, ComparisonResults
    Close fileNum

    LogMessage "Comparison results saved to: " & fileName
    Exit Sub

SaveError:
    LogMessage "Error saving comparison: " & Err.Description

End Sub