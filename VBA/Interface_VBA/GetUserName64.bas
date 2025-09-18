Attribute VB_Name = "GetUserName64"
Option Explicit

' GetUserName64.bas - Clean 64-bit VBA7 Implementation
' Designed for modern Excel (Office 2010+) - No backwards compatibility
' For legacy systems, use the original GetUserNameEx.bas

' Windows API declarations for modern VBA7 (Office 2010+)
' All VBA7 declarations require PtrSafe, regardless of architecture
#If Win64 Then
    ' True 64-bit Excel - Uses LongPtr for pointer parameters
    Private Declare PtrSafe Function GetUserNameA Lib "advapi32.dll" _
                                                   (ByVal lpBuffer As String, _
                                                   nSize As LongPtr) As Long
    Private Declare PtrSafe Function GetUserNameW Lib "advapi32.dll" _
                                                   (ByVal lpBuffer As String, _
                                                   nSize As LongPtr) As Long
#Else
    ' 32-bit Excel on VBA7 - Uses Long but still requires PtrSafe
    Private Declare PtrSafe Function GetUserNameA Lib "advapi32.dll" _
                                                   (ByVal lpBuffer As String, _
                                                   nSize As Long) As Long
    Private Declare PtrSafe Function GetUserNameW Lib "advapi32.dll" _
                                                   (ByVal lpBuffer As String, _
                                                   nSize As Long) As Long
#End If

' Primary function - Get current Windows username
Public Function GetUserName() As String
    Dim buffer As String
    Dim bufferSize As Long
    Dim result As Long
    Dim actualLength As Long

    ' Allocate buffer (Windows standard size for usernames)
    bufferSize = 256
    buffer = String(bufferSize, vbNullChar)

    ' Call Windows API with appropriate parameter types
    #If Win64 Then
        ' 64-bit: Use LongPtr for size parameter
        Dim nSize As LongPtr
        nSize = CLngPtr(bufferSize)
        result = GetUserNameA(buffer, nSize)
        actualLength = CLng(nSize)
    #Else
        ' 32-bit on VBA7: Use Long for size parameter
        Dim nSize As Long
        nSize = bufferSize
        result = GetUserNameA(buffer, nSize)
        actualLength = nSize
    #End If

    ' Process API result
    If result <> 0 Then
        ' Success: Extract username (remove null terminator)
        actualLength = InStr(buffer, vbNullChar)
        If actualLength > 1 Then
            GetUserName = Left(buffer, actualLength - 1)
        Else
            GetUserName = Trim(buffer)
        End If
    Else
        ' Fallback: Use environment variable
        GetUserName = Environ("USERNAME")
        If GetUserName = "" Then
            GetUserName = "Unknown"
        End If
    End If
End Function

' Unicode version for international character support
Public Function GetUserNameUnicode() As String
    Dim buffer As String
    Dim bufferSize As Long
    Dim result As Long
    Dim actualLength As Long

    ' Allocate Unicode buffer (larger for wide characters)
    bufferSize = 512
    buffer = String(bufferSize, vbNullChar)

    ' Call Unicode Windows API
    #If Win64 Then
        Dim nSize As LongPtr
        nSize = CLngPtr(bufferSize)
        result = GetUserNameW(buffer, nSize)
        actualLength = CLng(nSize)
    #Else
        Dim nSize As Long
        nSize = bufferSize
        result = GetUserNameW(buffer, nSize)
        actualLength = nSize
    #End If

    ' Process Unicode result
    If result <> 0 Then
        actualLength = InStr(buffer, vbNullChar)
        If actualLength > 1 Then
            GetUserNameUnicode = Left(buffer, actualLength - 1)
        Else
            GetUserNameUnicode = Trim(buffer)
        End If
    Else
        ' Fallback to ANSI version
        GetUserNameUnicode = GetUserName()
    End If
End Function

' Get comprehensive user and system information
Public Function GetUserInfo() As String
    Dim info As String

    info = "Username: " & GetUserName() & vbCrLf
    info = info & "Computer: " & Environ("COMPUTERNAME") & vbCrLf
    info = info & "Domain: " & Environ("USERDOMAIN") & vbCrLf
    info = info & "User Profile: " & Environ("USERPROFILE") & vbCrLf
    info = info & "Excel Version: " & Application.Version & vbCrLf
    info = info & "Architecture: " & GetArchitecture() & vbCrLf
    info = info & "OS Version: " & Environ("OS") & vbCrLf
    info = info & "Processor: " & Environ("PROCESSOR_IDENTIFIER")

    GetUserInfo = info
End Function

' Determine Excel architecture
Public Function GetArchitecture() As String
    #If Win64 Then
        GetArchitecture = "64-bit Excel on 64-bit Windows"
    #Else
        GetArchitecture = "32-bit Excel (VBA7 on " & IIf(InStr(Environ("PROCESSOR_ARCHITECTURE"), "64") > 0, "64-bit", "32-bit") & " Windows)"
    #End If
End Function

' Simple compatibility check
Public Function IsModernExcel() As Boolean
    ' This module requires VBA7 (Office 2010+)
    IsModernExcel = True  ' If this compiles, we're on VBA7+
End Function

' Performance test function
Public Function TestUserNamePerformance(iterations As Long) As String
    Dim startTime As Double
    Dim endTime As Double
    Dim i As Long
    Dim userName As String

    startTime = Timer

    For i = 1 To iterations
        userName = GetUserName()
    Next i

    endTime = Timer

    TestUserNamePerformance = "Retrieved username '" & userName & "' " & iterations & " times in " & _
                             Format(endTime - startTime, "0.000") & " seconds" & vbCrLf & _
                             "Average: " & Format((endTime - startTime) / iterations * 1000, "0.00") & "ms per call"
End Function

' Diagnostic function for troubleshooting
Public Function DiagnoseUserName() As String
    Dim diag As String

    diag = "=== GetUserName64 Diagnostics ===" & vbCrLf
    diag = diag & "Module: GetUserName64.bas (VBA7 Modern)" & vbCrLf
    diag = diag & "Architecture: " & GetArchitecture() & vbCrLf
    diag = diag & "Excel Version: " & Application.Version & vbCrLf & vbCrLf

    diag = diag & "=== Username Tests ===" & vbCrLf
    diag = diag & "ANSI GetUserName(): '" & GetUserName() & "'" & vbCrLf
    diag = diag & "Unicode GetUserName(): '" & GetUserNameUnicode() & "'" & vbCrLf
    diag = diag & "Environment USERNAME: '" & Environ("USERNAME") & "'" & vbCrLf & vbCrLf

    diag = diag & "=== System Environment ===" & vbCrLf
    diag = diag & "COMPUTERNAME: " & Environ("COMPUTERNAME") & vbCrLf
    diag = diag & "USERDOMAIN: " & Environ("USERDOMAIN") & vbCrLf
    diag = diag & "USERPROFILE: " & Environ("USERPROFILE") & vbCrLf
    diag = diag & "OS: " & Environ("OS") & vbCrLf
    diag = diag & "PROCESSOR_ARCHITECTURE: " & Environ("PROCESSOR_ARCHITECTURE") & vbCrLf

    ' Performance test
    diag = diag & vbCrLf & "=== Performance Test ===" & vbCrLf
    diag = diag & TestUserNamePerformance(1000)

    DiagnoseUserName = diag
End Function