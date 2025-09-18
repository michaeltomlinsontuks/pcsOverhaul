Attribute VB_Name = "InterfaceLauncher"
Option Explicit

Public Sub ShowMenu()
    On Error GoTo Error_Handler

    ' Initialize the system if needed
    If Not FileManager.ValidateDirectoryStructure() Then
        MsgBox "Warning: Some required directories are missing. Please check the system setup.", vbExclamation
    End If

    ' Show the main interface
    Main.Show
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "ShowMenu", "InterfaceLauncher"
End Sub

Public Sub LaunchMainInterface()
    ShowMenu
End Sub

Public Sub InitializeSystem()
    On Error GoTo Error_Handler

    ' Validate directory structure
    If Not FileManager.ValidateDirectoryStructure() Then
        MsgBox "System initialization failed: Directory structure validation failed.", vbCritical
        Exit Sub
    End If

    ' Show main interface
    ShowMenu
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "InitializeSystem", "InterfaceLauncher"
End Sub

Public Sub RefreshInterface()
    On Error GoTo Error_Handler

    ' If main interface is already loaded, refresh it
    If Not Main Is Nothing Then
        Main.RefreshAllLists
        Main.Show
    Else
        ' Otherwise launch it
        ShowMenu
    End If
    Exit Sub

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "RefreshInterface", "InterfaceLauncher"
End Sub