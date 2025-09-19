Attribute VB_Name = "ValidationTesting"
' **Purpose**: Testing module for ValidationFramework popup validation system
' **Dependencies**: ValidationFramework
' **Side Effects**: Shows test validation dialogs
' **Errors**: None (testing only)
Option Explicit

' **Purpose**: Comprehensive test of all validation functions
' **Parameters**: None
' **Returns**: None
' **Dependencies**: ValidationFramework
' **Side Effects**: Shows multiple test validation popups
' **Errors**: None (testing function)
Public Sub TestAllValidations()
    Dim testResult As Boolean

    ' Test 1: Required Field Validation
    Debug.Print "=== Testing Required Field Validation ==="
    testResult = ValidationFramework.ValidateRequired("", "Test Field")
    Debug.Print "Empty string test (should fail): " & testResult

    testResult = ValidationFramework.ValidateRequired("Valid Value", "Test Field")
    Debug.Print "Valid string test (should pass): " & testResult

    ' Test 2: Numeric Validation
    Debug.Print "=== Testing Numeric Validation ==="
    testResult = ValidationFramework.ValidateNumeric("ABC", "Test Number")
    Debug.Print "Non-numeric test (should fail): " & testResult

    testResult = ValidationFramework.ValidateNumeric("123.45", "Test Number")
    Debug.Print "Valid numeric test (should pass): " & testResult

    ' Test 3: Positive Number Validation
    Debug.Print "=== Testing Positive Number Validation ==="
    testResult = ValidationFramework.ValidatePositiveNumber("-5", "Test Amount")
    Debug.Print "Negative number test (should fail): " & testResult

    testResult = ValidationFramework.ValidatePositiveNumber("10.5", "Test Amount")
    Debug.Print "Positive number test (should pass): " & testResult

    ' Test 4: Date Validation
    Debug.Print "=== Testing Date Validation ==="
    testResult = ValidationFramework.ValidateDate("Invalid Date", "Test Date")
    Debug.Print "Invalid date test (should fail): " & testResult

    testResult = ValidationFramework.ValidateDate("01/01/2024", "Test Date")
    Debug.Print "Valid date test (should pass): " & testResult

    ' Test 5: Special Date Caption Validation
    Debug.Print "=== Testing Special Date Caption ==="
    testResult = ValidationFramework.ValidateSpecialDateCaption("Please click here to insert a date", "Test Date")
    Debug.Print "Special date caption test (user decides): " & testResult

    ' Test 6: Confirmation Dialog
    Debug.Print "=== Testing Confirmation Dialog ==="
    testResult = ValidationFramework.ShowConfirmation("Do you want to continue with testing?", "Test Confirmation")
    Debug.Print "Confirmation dialog test (user decides): " & testResult

    ' Test 7: Information, Warning, and Error Messages
    Debug.Print "=== Testing Message Types ==="
    ValidationFramework.ShowInformation "This is an information message for testing.", "Test Information"
    ValidationFramework.ShowWarning "This is a warning message for testing.", "Test Warning"
    ValidationFramework.ShowError "This is an error message for testing.", "Test Error"

    Debug.Print "=== All Validation Tests Completed ==="
End Sub

' **Purpose**: Test validation framework in context of enquiry form simulation
' **Parameters**: None
' **Returns**: None
' **Dependencies**: ValidationFramework
' **Side Effects**: Shows enquiry validation test popups
' **Errors**: None (testing function)
Public Sub TestEnquiryFormValidation()
    Dim customer As String
    Dim description As String
    Dim quantity As String
    Dim dateCaption As String

    Debug.Print "=== Testing Enquiry Form Validation Scenario ==="

    ' Simulate empty form submission
    customer = ""
    description = ""
    quantity = ""
    dateCaption = "Please click here to insert a date"

    Debug.Print "Testing empty form scenario..."

    ' Test customer validation
    If Not ValidationFramework.ValidateRequired(customer, "Customer") Then
        Debug.Print "Customer validation failed (expected)"
    End If

    ' Test description validation
    If Not ValidationFramework.ValidateRequired(description, "Component Description") Then
        Debug.Print "Description validation failed (expected)"
    End If

    ' Test quantity validation
    If Not ValidationFramework.ValidatePositiveNumber(quantity, "Component Quantity") Then
        Debug.Print "Quantity validation failed (expected)"
    End If

    ' Test date validation
    If Not ValidationFramework.ValidateSpecialDateCaption(dateCaption, "Enquiry Date") Then
        Debug.Print "Date validation cancelled by user"
    End If

    Debug.Print "=== Enquiry Form Validation Test Completed ==="
End Sub

' **Purpose**: Test file existence validation
' **Parameters**: None
' **Returns**: None
' **Dependencies**: ValidationFramework
' **Side Effects**: Shows file validation test popups
' **Errors**: None (testing function)
Public Sub TestFileValidation()
    Debug.Print "=== Testing File Validation ==="

    ' Test with non-existent file
    Dim testResult As Boolean
    testResult = ValidationFramework.ValidateFileExists("C:\NonExistentFile.txt", "Test File")
    Debug.Print "Non-existent file test (should fail): " & testResult

    ' Test with system file that should exist
    testResult = ValidationFramework.ValidateFileExists("C:\Windows\System32\notepad.exe", "Notepad")
    Debug.Print "Existing file test (should pass): " & testResult

    Debug.Print "=== File Validation Test Completed ==="
End Sub

' **Purpose**: Demonstrates validation workflow integration
' **Parameters**: None
' **Returns**: None
' **Dependencies**: ValidationFramework
' **Side Effects**: Shows workflow validation example
' **Errors**: None (testing function)
Public Sub DemonstrateValidationWorkflow()
    Debug.Print "=== Demonstrating Complete Validation Workflow ==="

    ' Simulate a complete form validation process
    Dim isValid As Boolean
    isValid = True

    ValidationFramework.ShowInformation "Starting validation workflow demonstration...", "Validation Demo"

    ' Step 1: Validate required fields
    If Not ValidationFramework.ValidateRequired("John Doe Manufacturing", "Customer") Then
        isValid = False
    End If

    ' Step 2: Validate numeric fields
    If isValid And Not ValidationFramework.ValidatePositiveNumber("100", "Quantity") Then
        isValid = False
    End If

    ' Step 3: Validate business logic
    If isValid Then
        If ValidationFramework.ShowConfirmation("All validations passed. Continue with save operation?", "Validation Success") Then
            ValidationFramework.ShowInformation "Data saved successfully!", "Save Complete"
        Else
            ValidationFramework.ShowWarning "Save operation cancelled by user.", "Save Cancelled"
        End If
    Else
        ValidationFramework.ShowError "Validation failed. Please correct errors and try again.", "Validation Failed"
    End If

    Debug.Print "=== Validation Workflow Demonstration Completed ==="
End Sub