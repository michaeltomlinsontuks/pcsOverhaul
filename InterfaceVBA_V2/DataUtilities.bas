Attribute VB_Name = "DataUtilities"
' **Purpose**: Data utility functions for pricing, lookups, and data transformations
' **CLAUDE.md Compliance**: New module to support existing functionality requirements
Option Explicit

' ===================================================================
' PRICING UTILITIES
' ===================================================================

' **Purpose**: Get standard price for component from price list
' **Parameters**:
'   - PriceListPath (String): Full path to price list Excel file
'   - ComponentCode (String): Component code to look up
' **Returns**: Currency - Standard price for component, 0 if not found
' **Dependencies**: DataManager.SafeOpenWorkbook for file access
' **Side Effects**: Opens and closes price list file
' **Errors**: Returns 0 if file access fails or component not found
Public Function GetStandardPrice(ByVal PriceListPath As String, ByVal ComponentCode As String) As Currency
    Dim PriceWB As Workbook
    Dim PriceWS As Worksheet
    Dim LastRow As Long
    Dim i As Long
    Dim FoundPrice As Currency

    On Error GoTo Error_Handler

    GetStandardPrice = 0

    ' Validate inputs
    If Trim(ComponentCode) = "" Then Exit Function
    If Not DataManager.FileExists(PriceListPath) Then Exit Function

    Set PriceWB = DataManager.SafeOpenWorkbook(PriceListPath)
    If PriceWB Is Nothing Then Exit Function

    Set PriceWS = PriceWB.Worksheets(1)
    LastRow = PriceWS.Cells(PriceWS.Rows.Count, 1).End(xlUp).Row

    ' Search for component code in first column, price in second column
    For i = 2 To LastRow ' Skip header row
        If UCase(Trim(PriceWS.Cells(i, 1).Value)) = UCase(Trim(ComponentCode)) Then
            If IsNumeric(PriceWS.Cells(i, 2).Value) Then
                FoundPrice = CCur(PriceWS.Cells(i, 2).Value)
                Exit For
            End If
        End If
    Next i

    DataManager.SafeCloseWorkbook PriceWB, False
    GetStandardPrice = FoundPrice
    Exit Function

Error_Handler:
    If Not PriceWB Is Nothing Then DataManager.SafeCloseWorkbook PriceWB, False
    CoreFramework.HandleStandardErrors Err.Number, "GetStandardPrice", "DataUtilities"
    GetStandardPrice = 0
End Function

' **Purpose**: Get component description from component database
' **Parameters**:
'   - ComponentDatabasePath (String): Full path to component database
'   - ComponentCode (String): Component code to look up
' **Returns**: String - Component description, empty if not found
' **Dependencies**: DataManager.SafeOpenWorkbook for file access
' **Side Effects**: Opens and closes component database file
' **Errors**: Returns empty string if file access fails or component not found
Public Function GetComponentDescription(ByVal ComponentDatabasePath As String, ByVal ComponentCode As String) As String
    Dim ComponentWB As Workbook
    Dim ComponentWS As Worksheet
    Dim LastRow As Long
    Dim i As Long
    Dim FoundDescription As String

    On Error GoTo Error_Handler

    GetComponentDescription = ""

    ' Validate inputs
    If Trim(ComponentCode) = "" Then Exit Function
    If Not DataManager.FileExists(ComponentDatabasePath) Then Exit Function

    Set ComponentWB = DataManager.SafeOpenWorkbook(ComponentDatabasePath)
    If ComponentWB Is Nothing Then Exit Function

    Set ComponentWS = ComponentWB.Worksheets(1)
    LastRow = ComponentWS.Cells(ComponentWS.Rows.Count, 1).End(xlUp).Row

    ' Search for component code in first column, description in second column
    For i = 2 To LastRow ' Skip header row
        If UCase(Trim(ComponentWS.Cells(i, 1).Value)) = UCase(Trim(ComponentCode)) Then
            FoundDescription = Trim(ComponentWS.Cells(i, 2).Value)
            Exit For
        End If
    Next i

    DataManager.SafeCloseWorkbook ComponentWB, False
    GetComponentDescription = FoundDescription
    Exit Function

Error_Handler:
    If Not ComponentWB Is Nothing Then DataManager.SafeCloseWorkbook ComponentWB, False
    CoreFramework.HandleStandardErrors Err.Number, "GetComponentDescription", "DataUtilities"
    GetComponentDescription = ""
End Function

' **Purpose**: Get material grade information from materials database
' **Parameters**:
'   - MaterialDatabasePath (String): Full path to materials database
'   - MaterialGrade (String): Material grade to look up
' **Returns**: String - Material specifications, empty if not found
' **Dependencies**: DataManager.SafeOpenWorkbook for file access
' **Side Effects**: Opens and closes materials database file
' **Errors**: Returns empty string if file access fails or grade not found
Public Function GetMaterialSpecifications(ByVal MaterialDatabasePath As String, ByVal MaterialGrade As String) As String
    Dim MaterialWB As Workbook
    Dim MaterialWS As Worksheet
    Dim LastRow As Long
    Dim i As Long
    Dim FoundSpecs As String

    On Error GoTo Error_Handler

    GetMaterialSpecifications = ""

    ' Validate inputs
    If Trim(MaterialGrade) = "" Then Exit Function
    If Not DataManager.FileExists(MaterialDatabasePath) Then Exit Function

    Set MaterialWB = DataManager.SafeOpenWorkbook(MaterialDatabasePath)
    If MaterialWB Is Nothing Then Exit Function

    Set MaterialWS = MaterialWB.Worksheets(1)
    LastRow = MaterialWS.Cells(MaterialWS.Rows.Count, 1).End(xlUp).Row

    ' Search for material grade in first column, specifications in second column
    For i = 2 To LastRow ' Skip header row
        If UCase(Trim(MaterialWS.Cells(i, 1).Value)) = UCase(Trim(MaterialGrade)) Then
            FoundSpecs = Trim(MaterialWS.Cells(i, 2).Value)
            Exit For
        End If
    Next i

    DataManager.SafeCloseWorkbook MaterialWB, False
    GetMaterialSpecifications = FoundSpecs
    Exit Function

Error_Handler:
    If Not MaterialWB Is Nothing Then DataManager.SafeCloseWorkbook MaterialWB, False
    CoreFramework.HandleStandardErrors Err.Number, "GetMaterialSpecifications", "DataUtilities"
    GetMaterialSpecifications = ""
End Function

' ===================================================================
' DATA TRANSFORMATION UTILITIES
' ===================================================================

' **Purpose**: Format currency for display in forms
' **Parameters**:
'   - Amount (Currency): Amount to format
'   - IncludeSymbol (Boolean, Optional): Include currency symbol (default True)
' **Returns**: String - Formatted currency string
' **Dependencies**: None
' **Side Effects**: None
' **Errors**: Returns "£0.00" if formatting fails
Public Function FormatCurrencyDisplay(ByVal Amount As Currency, Optional ByVal IncludeSymbol As Boolean = True) As String
    On Error GoTo Error_Handler

    If IncludeSymbol Then
        FormatCurrencyDisplay = Format(Amount, "£#,##0.00")
    Else
        FormatCurrencyDisplay = Format(Amount, "#,##0.00")
    End If
    Exit Function

Error_Handler:
    FormatCurrencyDisplay = IIf(IncludeSymbol, "£0.00", "0.00")
End Function

' **Purpose**: Format quantity for display in forms
' **Parameters**:
'   - Quantity (Long): Quantity to format
' **Returns**: String - Formatted quantity string
' **Dependencies**: None
' **Side Effects**: None
' **Errors**: Returns "0" if formatting fails
Public Function FormatQuantityDisplay(ByVal Quantity As Long) As String
    On Error GoTo Error_Handler

    FormatQuantityDisplay = Format(Quantity, "#,##0")
    Exit Function

Error_Handler:
    FormatQuantityDisplay = "0"
End Function

' **Purpose**: Parse currency value from user input
' **Parameters**:
'   - InputValue (Variant): User input to parse
' **Returns**: Currency - Parsed currency value, 0 if invalid
' **Dependencies**: None
' **Side Effects**: None
' **Errors**: Returns 0 if parsing fails
Public Function ParseCurrencyInput(ByVal InputValue As Variant) As Currency
    Dim CleanValue As String

    On Error GoTo Error_Handler

    ' Remove currency symbols and spaces
    CleanValue = Replace(Replace(Replace(CStr(InputValue), "£", ""), "$", ""), " ", "")

    If IsNumeric(CleanValue) Then
        ParseCurrencyInput = CCur(CleanValue)
    Else
        ParseCurrencyInput = 0
    End If
    Exit Function

Error_Handler:
    ParseCurrencyInput = 0
End Function

' **Purpose**: Parse quantity value from user input
' **Parameters**:
'   - InputValue (Variant): User input to parse
' **Returns**: Long - Parsed quantity value, 0 if invalid
' **Dependencies**: None
' **Side Effects**: None
' **Errors**: Returns 0 if parsing fails
Public Function ParseQuantityInput(ByVal InputValue As Variant) As Long
    Dim CleanValue As String

    On Error GoTo Error_Handler

    ' Remove commas and spaces
    CleanValue = Replace(Replace(CStr(InputValue), ",", ""), " ", "")

    If IsNumeric(CleanValue) Then
        ParseQuantityInput = CLng(CleanValue)
    Else
        ParseQuantityInput = 0
    End If
    Exit Function

Error_Handler:
    ParseQuantityInput = 0
End Function

' ===================================================================
' LOOKUP UTILITIES
' ===================================================================

' **Purpose**: Lookup value in any Excel table
' **Parameters**:
'   - TablePath (String): Full path to lookup table file
'   - SearchValue (Variant): Value to search for
'   - SearchColumn (Long, Optional): Column to search in (default 1)
'   - ReturnColumn (Long, Optional): Column to return value from (default 2)
' **Returns**: Variant - Found value, empty string if not found
' **Dependencies**: DataManager.SafeOpenWorkbook for file access
' **Side Effects**: Opens and closes lookup table file
' **Errors**: Returns empty string if file access fails or value not found
Public Function LookupValue(ByVal TablePath As String, ByVal SearchValue As Variant, Optional ByVal SearchColumn As Long = 1, Optional ByVal ReturnColumn As Long = 2) As Variant
    Dim TableWB As Workbook
    Dim TableWS As Worksheet
    Dim LastRow As Long
    Dim i As Long
    Dim FoundValue As Variant

    On Error GoTo Error_Handler

    LookupValue = ""

    ' Validate inputs
    If Not DataManager.FileExists(TablePath) Then Exit Function
    If SearchColumn < 1 Or ReturnColumn < 1 Then Exit Function

    Set TableWB = DataManager.SafeOpenWorkbook(TablePath)
    If TableWB Is Nothing Then Exit Function

    Set TableWS = TableWB.Worksheets(1)
    LastRow = TableWS.Cells(TableWS.Rows.Count, SearchColumn).End(xlUp).Row

    ' Search for value
    For i = 2 To LastRow ' Skip header row
        If TableWS.Cells(i, SearchColumn).Value = SearchValue Then
            FoundValue = TableWS.Cells(i, ReturnColumn).Value
            Exit For
        End If
    Next i

    DataManager.SafeCloseWorkbook TableWB, False
    LookupValue = FoundValue
    Exit Function

Error_Handler:
    If Not TableWB Is Nothing Then DataManager.SafeCloseWorkbook TableWB, False
    CoreFramework.HandleStandardErrors Err.Number, "LookupValue", "DataUtilities"
    LookupValue = ""
End Function

' **Purpose**: Get list of values from lookup table column
' **Parameters**:
'   - TablePath (String): Full path to lookup table file
'   - ColumnNumber (Long, Optional): Column to extract values from (default 1)
' **Returns**: Variant - Array of values, empty array if failed
' **Dependencies**: DataManager.SafeOpenWorkbook for file access
' **Side Effects**: Opens and closes lookup table file
' **Errors**: Returns empty array if file access fails
Public Function GetColumnValues(ByVal TablePath As String, Optional ByVal ColumnNumber As Long = 1) As Variant
    Dim TableWB As Workbook
    Dim TableWS As Worksheet
    Dim LastRow As Long
    Dim Values() As String
    Dim i As Long
    Dim ValueCount As Long

    On Error GoTo Error_Handler

    GetColumnValues = Array()

    ' Validate inputs
    If Not DataManager.FileExists(TablePath) Then Exit Function
    If ColumnNumber < 1 Then Exit Function

    Set TableWB = DataManager.SafeOpenWorkbook(TablePath)
    If TableWB Is Nothing Then Exit Function

    Set TableWS = TableWB.Worksheets(1)
    LastRow = TableWS.Cells(TableWS.Rows.Count, ColumnNumber).End(xlUp).Row

    ' Extract values (skip header row)
    ValueCount = 0
    For i = 2 To LastRow
        If Trim(TableWS.Cells(i, ColumnNumber).Value) <> "" Then
            ReDim Preserve Values(ValueCount)
            Values(ValueCount) = Trim(TableWS.Cells(i, ColumnNumber).Value)
            ValueCount = ValueCount + 1
        End If
    Next i

    DataManager.SafeCloseWorkbook TableWB, False

    If ValueCount > 0 Then
        GetColumnValues = Values
    Else
        GetColumnValues = Array()
    End If
    Exit Function

Error_Handler:
    If Not TableWB Is Nothing Then DataManager.SafeCloseWorkbook TableWB, False
    CoreFramework.HandleStandardErrors Err.Number, "GetColumnValues", "DataUtilities"
    GetColumnValues = Array()
End Function