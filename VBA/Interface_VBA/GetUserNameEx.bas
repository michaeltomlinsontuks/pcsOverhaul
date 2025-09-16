Attribute VB_Name = "GetUserNameEx"
Option Explicit
#If VBA7 Then
    Private Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
                                                            (ByVal lpBuffer As String, _
                                                            nSize As LongPtr) As Long
#Else
    Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
                                                            (ByVal lpBuffer As String, _
                                                            nSize As Long) As Long
#End If
Public Function Get_User_Name()
    
    Dim lpBuff As String * 25
    Dim ret As Long, UserName As String
    ret = GetUserName(lpBuff, 25)
    Get_User_Name = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)
    
End Function

