Attribute VB_Name = "CoreUtilities"
' **Purpose**: Basic utility functions used throughout the PCS system
' **Consolidates**: RemoveCharacters.bas, GetValue.bas, Very_HiddenSheet.bas, GetUserNameEx.bas, GetUserName64.bas
' **Original Functionality**: All functions preserved exactly as original modules

Option Explicit

' 32/64-bit compatibility declarations
#If VBA7 Then
    Private Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
                                                            (ByVal lpBuffer As String, _
                                                            nSize As LongPtr) As Long
#Else
    Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
                                                            (ByVal lpBuffer As String, _
                                                            nSize As Long) As Long
#End If

' **Purpose**: Get current Windows username with 32/64-bit compatibility
' **Parameters**: None
' **Returns**: String - Current Windows username
' **Dependencies**: Windows API GetUserName
' **Original Module**: GetUserNameEx.bas + GetUserName64.bas (merged)
Public Function Get_User_Name()
    Dim lpBuff As String * 25
    Dim ret As Long, UserName As String

    #If VBA7 Then
        Dim nSize As LongPtr
        nSize = 25
        ret = GetUserName(lpBuff, nSize)
    #Else
        ret = GetUserName(lpBuff, 25)
    #End If

    Get_User_Name = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)
End Function

' **Purpose**: Remove invalid characters from strings for file names
' **Parameters**:
'   - Str (String): Input string to clean
' **Returns**: String - Cleaned string with invalid characters removed
' **Dependencies**: None
' **Original Module**: RemoveCharacters.bas
Public Function Remove_Characters(Str As String)
    For i = 1 To Len(Str)
        If Mid(Str, i, 1) = "/" Or Mid(Str, i, 1) = ":" Or Mid(Str, i, 1) = " " Then
            Str = Mid(Str, 1, i - 1) & Mid(Str, i + 1, Len(Str) - i)
        End If
    Next i

    Remove_Characters = Str
End Function

' **Purpose**: Insert spaces and format characters in strings
' **Parameters**:
'   - Str (String): Input string to format
' **Returns**: String - Formatted string with spaces inserted
' **Dependencies**: None
' **Original Module**: RemoveCharacters.bas
Public Function Insert_Characters(Str As String)
    j = Len(Str)
    i = 0

    For i = 2 To j
        If Mid(Str, i, 1) = "_" Then
            Str = Mid(Str, 1, i - 1) & " " & Mid(Str, i + 1, Len(Str) - i)
            i = i + 1
        Else
            If UCase(Mid(Str, i, 1)) = Mid(Str, i, 1) Then
                Str = Mid(Str, 1, i - 1) & " " & Mid(Str, i, Len(Str) - i + 1)
                j = j + 1
                i = i + 1
            End If
        End If
    Next i

    If InStr(1, Str, "Component ", vbTextCompare) > 0 Then
        Str = Right(Str, Len(Str) - Len("Component "))
    End If

    Insert_Characters = Str
End Function

' **Purpose**: Get value from closed workbook (original functionality)
' **Parameters**:
'   - path: File path
'   - File: File name
'   - sheet: Sheet name
'   - ref: Cell reference
' **Returns**: Value from specified cell in closed workbook
' **Dependencies**: ExecuteExcel4Macro
' **Original Module**: GetValue.bas
Public Function GetValue(path, File, sheet, ref)
    '   Retrieves a value from a closed workbook
    Dim arg As String

    '   Make sure the file exists
    If Right(path, 1) <> "\" Then path = path & "\"
    If Dir(path & File) = "" Then
        GetValue = "File Not Found"
        Exit Function
    End If

    '   Create the argument
    arg = "'" & path & "[" & File & "]" & sheet & "'!" & _
      Range(ref).Range("A1").Address(, , xlR1C1)

    '   Execute an XLM macro
    GetValue = ExecuteExcel4Macro(arg)
End Function

' **Purpose**: Hide sheet with xlVeryHidden property
' **Parameters**:
'   - SheetNam (String): Name of sheet to hide
' **Returns**: None
' **Dependencies**: None
' **Original Module**: Very_HiddenSheet.bas
Public Function VeryHiddenSheet(SheetNam As String)
    Sheets(SheetNam).Visible = xlVeryHidden
End Function

' **Purpose**: Show previously hidden sheet
' **Parameters**:
'   - SheetNam (String): Name of sheet to show
' **Returns**: None
' **Dependencies**: None
' **Original Module**: Very_HiddenSheet.bas
Public Function ShowSheet(SheetNam As String)
    Sheets(SheetNam).Visible = True
End Function

' **Purpose**: Test function for GetValue functionality
' **Parameters**: None
' **Returns**: None (displays message box)
' **Dependencies**: GetValue, MsgBox
' **Original Module**: GetValue.bas
Private Function TestGetValue()
    p = "c:\XLFiles\Budget"
    f = "99Budget.xls"
    s = "Sheet1"
    a = "A1"
    MsgBox GetValue(p, f, s, a)
End Function

' **Purpose**: Test function for bulk GetValue operations
' **Parameters**: None
' **Returns**: None (populates active worksheet)
' **Dependencies**: GetValue
' **Original Module**: GetValue.bas
Private Function TestGetValue2()
    p = "c:\XLFiles\Budget"
    f = "99Budget.xls"
    s = "Sheet1"
    Application.ScreenUpdating = False
    For r = 1 To 100
        For c = 1 To 12
            a = Cells(r, c).Address
            Cells(r, c) = GetValue(p, f, s, a)
        Next c
    Next r
    Application.ScreenUpdating = True
End Function