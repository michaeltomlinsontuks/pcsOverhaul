Attribute VB_Name = "ValidationFramework"
' **Purpose**: Core validation framework providing standardized popup validation messages
' **Dependencies**: None (uses built-in VBA functions only)
' **Side Effects**: Shows MsgBox dialogs for validation feedback
' **Errors**: Returns False on validation failure, True on success
' **32/64-bit Notes**: Compatible with both architectures
Option Explicit

' **Purpose**: Validates required field and shows popup if empty
' **Parameters**:
'   - fieldValue (Variant): Value to validate
'   - fieldName (String): Display name for field in error message
'   - setFocusControl (Object): Optional control to focus on error
' **Returns**: Boolean - True if field has value, False if empty
' **Dependencies**: None
' **Side Effects**: Shows MsgBox popup on validation failure
' **Errors**: Returns False on validation failure
Public Function ValidateRequired(fieldValue As Variant, fieldName As String, Optional setFocusControl As Object = Nothing) As Boolean
    ValidateRequired = True

    If IsNull(fieldValue) Or Trim(CStr(fieldValue)) = "" Then
        MsgBox "Please enter " & fieldName & ".", vbExclamation + vbOKOnly, "Required Field Missing"
        If Not setFocusControl Is Nothing Then setFocusControl.SetFocus
        ValidateRequired = False
    End If
End Function

' **Purpose**: Validates numeric field and shows popup if invalid
' **Parameters**:
'   - fieldValue (Variant): Value to validate
'   - fieldName (String): Display name for field in error message
'   - setFocusControl (Object): Optional control to focus on error
' **Returns**: Boolean - True if field is numeric, False if invalid
' **Dependencies**: None
' **Side Effects**: Shows MsgBox popup on validation failure
' **Errors**: Returns False on validation failure
Public Function ValidateNumeric(fieldValue As Variant, fieldName As String, Optional setFocusControl As Object = Nothing) As Boolean
    ValidateNumeric = True

    If Not IsNumeric(fieldValue) Or Trim(CStr(fieldValue)) = "" Then
        MsgBox fieldName & " must be a valid number.", vbExclamation + vbOKOnly, "Invalid Number"
        If Not setFocusControl Is Nothing Then setFocusControl.SetFocus
        ValidateNumeric = False
    End If
End Function

' **Purpose**: Validates positive number field and shows popup if invalid
' **Parameters**:
'   - fieldValue (Variant): Value to validate
'   - fieldName (String): Display name for field in error message
'   - setFocusControl (Object): Optional control to focus on error
' **Returns**: Boolean - True if field is positive number, False if invalid
' **Dependencies**: ValidateNumeric
' **Side Effects**: Shows MsgBox popup on validation failure
' **Errors**: Returns False on validation failure
Public Function ValidatePositiveNumber(fieldValue As Variant, fieldName As String, Optional setFocusControl As Object = Nothing) As Boolean
    ValidatePositiveNumber = True

    If Not ValidateNumeric(fieldValue, fieldName, setFocusControl) Then
        ValidatePositiveNumber = False
        Exit Function
    End If

    If CDbl(fieldValue) <= 0 Then
        MsgBox fieldName & " must be greater than zero.", vbExclamation + vbOKOnly, "Invalid Value"
        If Not setFocusControl Is Nothing Then setFocusControl.SetFocus
        ValidatePositiveNumber = False
    End If
End Function

' **Purpose**: Validates list selection and shows popup if none selected
' **Parameters**:
'   - listControl (Object): ListBox or ComboBox control to validate
'   - fieldName (String): Display name for field in error message
' **Returns**: Boolean - True if selection made, False if no selection
' **Dependencies**: None
' **Side Effects**: Shows MsgBox popup on validation failure, sets focus to control
' **Errors**: Returns False on validation failure
Public Function ValidateListSelection(listControl As Object, fieldName As String) As Boolean
    ValidateListSelection = True

    If listControl.ListIndex < 0 Then
        MsgBox "Please select a " & fieldName & ".", vbExclamation + vbOKOnly, "Selection Required"
        listControl.SetFocus
        ValidateListSelection = False
    End If
End Function

' **Purpose**: Validates date field and shows popup if invalid
' **Parameters**:
'   - fieldValue (Variant): Value to validate as date
'   - fieldName (String): Display name for field in error message
'   - setFocusControl (Object): Optional control to focus on error
' **Returns**: Boolean - True if valid date, False if invalid
' **Dependencies**: None
' **Side Effects**: Shows MsgBox popup on validation failure
' **Errors**: Returns False on validation failure
Public Function ValidateDate(fieldValue As Variant, fieldName As String, Optional setFocusControl As Object = Nothing) As Boolean
    ValidateDate = True

    If Not IsDate(fieldValue) Then
        MsgBox fieldName & " must be a valid date.", vbExclamation + vbOKOnly, "Invalid Date"
        If Not setFocusControl Is Nothing Then setFocusControl.SetFocus
        ValidateDate = False
    End If
End Function

' **Purpose**: Validates file exists and shows popup if missing
' **Parameters**:
'   - filePath (String): Full path to file to check
'   - fieldName (String): Display name for field in error message
' **Returns**: Boolean - True if file exists, False if missing
' **Dependencies**: None
' **Side Effects**: Shows MsgBox popup on validation failure
' **Errors**: Returns False on validation failure or file access error
Public Function ValidateFileExists(filePath As String, fieldName As String) As Boolean
    ValidateFileExists = True

    On Error GoTo FileError

    If Trim(filePath) = "" Or Dir(filePath) = "" Then
        MsgBox "The file specified for " & fieldName & " does not exist." & vbCrLf & "Path: " & filePath, vbExclamation + vbOKOnly, "File Not Found"
        ValidateFileExists = False
    End If

    Exit Function

FileError:
    MsgBox "Cannot access file for " & fieldName & "." & vbCrLf & "Path: " & filePath, vbCritical + vbOKOnly, "File Access Error"
    ValidateFileExists = False
End Function

' **Purpose**: Shows confirmation dialog with Yes/No options
' **Parameters**:
'   - message (String): Message to display
'   - title (String): Dialog title
' **Returns**: Boolean - True if user clicked Yes, False if No
' **Dependencies**: None
' **Side Effects**: Shows MsgBox confirmation dialog
' **Errors**: Returns False if dialog fails
Public Function ShowConfirmation(message As String, title As String) As Boolean
    ShowConfirmation = (MsgBox(message, vbYesNo + vbQuestion, title) = vbYes)
End Function

' **Purpose**: Shows information popup message
' **Parameters**:
'   - message (String): Message to display
'   - title (String): Dialog title
' **Returns**: None
' **Dependencies**: None
' **Side Effects**: Shows MsgBox information dialog
' **Errors**: None
Public Sub ShowInformation(message As String, title As String)
    MsgBox message, vbInformation + vbOKOnly, title
End Sub

' **Purpose**: Shows warning popup message
' **Parameters**:
'   - message (String): Message to display
'   - title (String): Dialog title
' **Returns**: None
' **Dependencies**: None
' **Side Effects**: Shows MsgBox warning dialog
' **Errors**: None
Public Sub ShowWarning(message As String, title As String)
    MsgBox message, vbExclamation + vbOKOnly, title
End Sub

' **Purpose**: Shows error popup message
' **Parameters**:
'   - message (String): Message to display
'   - title (String): Dialog title
' **Returns**: None
' **Dependencies**: None
' **Side Effects**: Shows MsgBox error dialog
' **Errors**: None
Public Sub ShowError(message As String, title As String)
    MsgBox message, vbCritical + vbOKOnly, title
End Sub

' **Purpose**: Validates special date caption (used in enquiry forms)
' **Parameters**:
'   - dateCaption (String): Caption text to check
'   - fieldName (String): Display name for field
' **Returns**: Boolean - True if user wants to continue without date, False to cancel
' **Dependencies**: ShowConfirmation
' **Side Effects**: Shows confirmation dialog for missing date
' **Errors**: Returns False if user cancels
Public Function ValidateSpecialDateCaption(dateCaption As String, fieldName As String) As Boolean
    ValidateSpecialDateCaption = True

    If dateCaption = "Please click here to insert a date" Then
        ValidateSpecialDateCaption = ShowConfirmation("No " & fieldName & " has been entered. Continue without date?", "Missing Date")
    End If
End Function