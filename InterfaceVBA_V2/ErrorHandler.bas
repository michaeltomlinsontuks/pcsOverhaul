Attribute VB_Name = "ErrorHandler"
Option Explicit

Public Const ERR_FILE_NOT_FOUND As Long = 53
Public Const ERR_PATH_NOT_FOUND As Long = 76
Public Const ERR_PERMISSION_DENIED As Long = 70
Public Const ERR_DISK_FULL As Long = 61

Public Sub LogError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String, ByVal ProcedureName As String, Optional ByVal ModuleName As String = "")
    Dim ErrorMsg As String
    Dim LogPath As String
    Dim FileNum As Integer

    On Error GoTo ErrorLogging_Error

    ErrorMsg = Format(Now, "yyyy-mm-dd hh:mm:ss") & " - "
    If ModuleName <> "" Then ErrorMsg = ErrorMsg & ModuleName & "."
    ErrorMsg = ErrorMsg & ProcedureName & " - Error " & ErrorNumber & ": " & ErrorDescription

    LogPath = ThisWorkbook.Path & "\error_log.txt"

    FileNum = FreeFile
    Open LogPath For Append As #FileNum
    Print #FileNum, ErrorMsg
    Close #FileNum

    Exit Sub

ErrorLogging_Error:
    MsgBox "Critical Error: Unable to log error to file." & vbCrLf & _
           "Original Error: " & ErrorNumber & " - " & ErrorDescription, vbCritical
End Sub

Public Function HandleStandardErrors(ByVal ErrorNumber As Long, ByVal ProcedureName As String, Optional ByVal ModuleName As String = "") As Boolean
    Dim UserMsg As String

    Select Case ErrorNumber
        Case ERR_FILE_NOT_FOUND
            UserMsg = "File not found. Please check the file path and try again."
        Case ERR_PATH_NOT_FOUND
            UserMsg = "Directory not found. Please verify the directory structure."
        Case ERR_PERMISSION_DENIED
            UserMsg = "Access denied. Please check file permissions or close any open files."
        Case ERR_DISK_FULL
            UserMsg = "Disk full. Please free up space and try again."
        Case Else
            HandleStandardErrors = False
            Exit Function
    End Select

    LogError ErrorNumber, Err.Description, ProcedureName, ModuleName
    MsgBox UserMsg & vbCrLf & vbCrLf & "Technical Details: Error " & ErrorNumber, vbExclamation
    HandleStandardErrors = True
End Function

Public Sub ClearError()
    Err.Clear
End Sub

Public Function GetLastErrorInfo() As String
    GetLastErrorInfo = "Error " & Err.Number & ": " & Err.Description
End Function