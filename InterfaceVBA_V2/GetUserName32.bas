Attribute VB_Name = "GetUserName32"
' **Purpose**: 32-bit Windows API user authentication functions
' **Original**: Interface_VBA/GetUserName.bas and GetUserNameEx.bas
' **CLAUDE.md Compliance**: 32/64-bit API compatibility - NOT backwards compatible by design
' **Deployment**: For 32-bit Excel systems ONLY
Option Explicit

' **Purpose**: Windows API declaration for getting current user name (32-bit)
' **Dependencies**: advapi32.dll
' **32/64-bit Notes**: 32-bit ONLY - will not compile on 64-bit systems
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long

' **Purpose**: Get current Windows username for 32-bit systems
' **Original**: GetUserName.bas.Get_User_Name() and GetUserNameEx.bas.Get_User_Name()
' **Parameters**: None
' **Returns**: String - Current Windows username
' **Dependencies**: Windows API (advapi32.dll)
' **Side Effects**: None
' **Errors**: Returns "Unknown" if API call fails
' **32/64-bit Notes**: 32-bit Excel ONLY - designed to NOT be backwards compatible
Public Function Get_User_Name() As String
    Dim UserName As String
    Dim UserNameSize As Long

    On Error GoTo Error_Handler

    UserNameSize = 255
    UserName = Space$(UserNameSize)

    If GetUserName(UserName, UserNameSize) <> 0 Then
        Get_User_Name = Left$(UserName, UserNameSize - 1)
    Else
        Get_User_Name = "Unknown"
    End If
    Exit Function

Error_Handler:
    Get_User_Name = "Unknown"
End Function

' **Purpose**: Alternative user name function for compatibility
' **Returns**: String - Current Windows username (same as Get_User_Name)
' **CLAUDE.md Compliance**: Maintains exact function signatures from original system
Public Function GetCurrentUser() As String
    GetCurrentUser = Get_User_Name()
End Function