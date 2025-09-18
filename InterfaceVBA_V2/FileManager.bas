Attribute VB_Name = "FileManager"
Option Explicit

Private Const ROOT_PATH As String = ""

Public Function GetRootPath() As String
    If ROOT_PATH = "" Then
        GetRootPath = ThisWorkbook.Path
    Else
        GetRootPath = ROOT_PATH
    End If
End Function

Public Function ValidateDirectoryStructure() As Boolean
    Dim RequiredDirs As Variant
    Dim i As Integer

    On Error GoTo Error_Handler

    RequiredDirs = Array("Enquiries", "Quotes", "WIP", "Archive", "Contracts", _
                        "Customers", "Templates", "Job Templates", "images")

    For i = 0 To UBound(RequiredDirs)
        If Not DirExists(GetRootPath & "\" & RequiredDirs(i)) Then
            ValidateDirectoryStructure = False
            ErrorHandler.LogError 0, "Missing directory: " & RequiredDirs(i), "ValidateDirectoryStructure", "FileManager"
            Exit Function
        End If
    Next i

    ValidateDirectoryStructure = True
    Exit Function

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "ValidateDirectoryStructure", "FileManager"
    ValidateDirectoryStructure = False
End Function

Public Function DirExists(ByVal DirPath As String) As Boolean
    On Error GoTo Error_Handler
    DirExists = (Dir(DirPath, vbDirectory) <> "")
    Exit Function

Error_Handler:
    DirExists = False
End Function

Public Function FileExists(ByVal FilePath As String) As Boolean
    On Error GoTo Error_Handler
    FileExists = (Dir(FilePath) <> "")
    Exit Function

Error_Handler:
    FileExists = False
End Function

Public Function SafeOpenWorkbook(ByVal FilePath As String) As Workbook
    Dim wb As Workbook

    On Error GoTo Error_Handler

    If Not FileExists(FilePath) Then
        ErrorHandler.LogError ERR_FILE_NOT_FOUND, "File not found: " & FilePath, "SafeOpenWorkbook", "FileManager"
        Set SafeOpenWorkbook = Nothing
        Exit Function
    End If

    Set wb = Workbooks.Open(FilePath, ReadOnly:=False, UpdateLinks:=False)
    Set SafeOpenWorkbook = wb
    Exit Function

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "SafeOpenWorkbook", "FileManager"
    Set SafeOpenWorkbook = Nothing
End Function

Public Function SafeCloseWorkbook(ByRef wb As Workbook, Optional ByVal SaveChanges As Boolean = True) As Boolean
    On Error GoTo Error_Handler

    If Not wb Is Nothing Then
        wb.Close SaveChanges:=SaveChanges
        Set wb = Nothing
        SafeCloseWorkbook = True
    End If
    Exit Function

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "SafeCloseWorkbook", "FileManager"
    SafeCloseWorkbook = False
End Function

Public Function GetFileList(ByVal DirectoryName As String) As Variant
    Dim DirPath As String
    Dim FileName As String
    Dim FileList() As String
    Dim FileCount As Integer

    On Error GoTo Error_Handler

    DirPath = GetRootPath & "\" & DirectoryName & "\"

    If Not DirExists(DirPath) Then
        ErrorHandler.LogError ERR_PATH_NOT_FOUND, "Directory not found: " & DirPath, "GetFileList", "FileManager"
        GetFileList = Array()
        Exit Function
    End If

    FileName = Dir(DirPath & "*.xls*")
    FileCount = 0

    Do While FileName <> ""
        ReDim Preserve FileList(FileCount)
        FileList(FileCount) = FileName
        FileCount = FileCount + 1
        FileName = Dir
    Loop

    If FileCount > 0 Then
        GetFileList = FileList
    Else
        GetFileList = Array()
    End If
    Exit Function

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "GetFileList", "FileManager"
    GetFileList = Array()
End Function

Public Function CreateBackup(ByVal FilePath As String) As Boolean
    Dim BackupPath As String
    Dim BackupDir As String

    On Error GoTo Error_Handler

    BackupDir = GetRootPath & "\Backups\"
    If Not DirExists(BackupDir) Then
        MkDir BackupDir
    End If

    BackupPath = BackupDir & Format(Now, "yyyymmdd_hhmmss_") & Dir(FilePath)

    FileCopy FilePath, BackupPath
    CreateBackup = True
    Exit Function

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "CreateBackup", "FileManager"
    CreateBackup = False
End Function

Public Function GetNextFileName(ByVal DirectoryName As String, ByVal Prefix As String, ByVal Extension As String) As String
    Dim DirPath As String
    Dim Counter As Integer
    Dim FileName As String

    DirPath = GetRootPath & "\" & DirectoryName & "\"
    Counter = 1

    Do
        FileName = Prefix & Format(Counter, "0000") & Extension
        Counter = Counter + 1
    Loop While FileExists(DirPath & FileName)

    GetNextFileName = FileName
End Function