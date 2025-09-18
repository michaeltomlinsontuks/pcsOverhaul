Option Explicit
    Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
                                                            (ByVal lpBuffer As String, _
                                                            nSize As Long) As Long
Public Function Get_User_Name()
    
    Dim lpBuff As String * 25
    Dim ret As Long, UserName As String
    ret = GetUserName(lpBuff, 25)
    Get_User_Name = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)
    
End Function

