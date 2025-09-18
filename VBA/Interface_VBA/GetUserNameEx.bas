Attribute VB_Name = "GetUserNameEx"
Option Explicit

' Windows API declarations for 32-bit and 64-bit compatibility
#If VBA7 Then
    ' 64-bit Excel (Office 2010+) - Uses PtrSafe and LongPtr for pointers
    Private Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
                                                            (ByVal lpBuffer As String, _
                                                            nSize As LongPtr) As Long
    Private Declare PtrSafe Function GetUserNameW Lib "advapi32.dll" _
                                                            (ByVal lpBuffer As String, _
                                                            nSize As LongPtr) As Long
#Else
    ' 32-bit Excel (Pre-2010) - Uses Long for pointers
    Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
                                                            (ByVal lpBuffer As String, _
                                                            nSize As Long) As Long
    Private Declare Function GetUserNameW Lib "advapi32.dll" _
                                                            (ByVal lpBuffer As String, _
                                                            nSize As Long) As Long
#End If

' Enhanced GetUserName function with proper 32/64-bit compatibility
Public Function Get_User_Name() As String
    Dim lpBuff As String
    Dim bufferSize As Long
    Dim result As Long
    Dim actualLength As Long

    ' Set initial buffer size (Windows recommends 256 for usernames)
    bufferSize = 256
    lpBuff = String(bufferSize, vbNullChar)

    ' Handle the size parameter correctly for both architectures
    #If VBA7 Then
        ' 64-bit: Use LongPtr for size parameter
        Dim nSize As LongPtr
        nSize = CLngPtr(bufferSize)
        result = GetUserName(lpBuff, nSize)
        actualLength = CLng(nSize)
    #Else
        ' 32-bit: Use Long for size parameter
        Dim nSize As Long
        nSize = bufferSize
        result = GetUserName(lpBuff, nSize)
        actualLength = nSize
    #End If

    ' Check if the API call was successful
    If result <> 0 Then
        ' Extract the username (remove null terminator)
        actualLength = InStr(lpBuff, vbNullChar) - 1
        If actualLength > 0 Then
            Get_User_Name = Left(lpBuff, actualLength)
        Else
            Get_User_Name = Trim(lpBuff)
        End If
    Else
        ' Fallback: Use environment variable if API fails
        Get_User_Name = Environ("USERNAME")
        If Get_User_Name = "" Then
            Get_User_Name = "Unknown User"
        End If
    End If
End Function

' Alternative function using Unicode version for better character support
Public Function Get_User_Name_Unicode() As String
    Dim lpBuff As String
    Dim bufferSize As Long
    Dim result As Long
    Dim actualLength As Long

    ' Set buffer size for Unicode (twice as large)
    bufferSize = 512
    lpBuff = String(bufferSize, vbNullChar)

    ' Handle the size parameter correctly for both architectures
    #If VBA7 Then
        ' 64-bit: Use LongPtr for size parameter
        Dim nSize As LongPtr
        nSize = CLngPtr(bufferSize)
        result = GetUserNameW(lpBuff, nSize)
        actualLength = CLng(nSize)
    #Else
        ' 32-bit: Use Long for size parameter
        Dim nSize As Long
        nSize = bufferSize
        result = GetUserNameW(lpBuff, nSize)
        actualLength = nSize
    #End If

    ' Process result
    If result <> 0 Then
        actualLength = InStr(lpBuff, vbNullChar) - 1
        If actualLength > 0 Then
            Get_User_Name_Unicode = Left(lpBuff, actualLength)
        Else
            Get_User_Name_Unicode = Trim(lpBuff)
        End If
    Else
        ' Fallback to ANSI version
        Get_User_Name_Unicode = Get_User_Name()
    End If
End Function

' Utility function to get detailed user information
Public Function Get_User_Info() As String
    Dim userName As String
    Dim computerName As String
    Dim domainName As String

    userName = Get_User_Name()
    computerName = Environ("COMPUTERNAME")
    domainName = Environ("USERDOMAIN")

    Get_User_Info = "User: " & userName & vbCrLf & _
                   "Computer: " & computerName & vbCrLf & _
                   "Domain: " & domainName & vbCrLf & _
                   "Excel Architecture: " & GetExcelArchitecture()
End Function

' Helper function to determine Excel architecture
Private Function GetExcelArchitecture() As String
    #If VBA7 Then
        #If Win64 Then
            GetExcelArchitecture = "64-bit"
        #Else
            GetExcelArchitecture = "32-bit (VBA7)"
        #End If
    #Else
        GetExcelArchitecture = "32-bit (Legacy)"
    #End If
End Function

