Option Explicit
    Private Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
                                                            (ByVal lpBuffer As String, _
                                                            nSize As LongPtr) As Long
Public Function Get_User_Name()

    Dim lpBuff As String * 25
    Dim ret As Long, UserName As String
    Dim nSize As LongPtr
    nSize = 25
    ret = GetUserName(lpBuff, nSize)
    Get_User_Name = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)

End Function