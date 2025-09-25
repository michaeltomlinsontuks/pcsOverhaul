VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FWIP
   Caption         =   "WIP Reports"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4470
   OleObjectBlob   =   "fwip.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fwip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' **Purpose**: Thin wrapper for WIP report generation - calls ReportingSystem module
' **Original**: fwip.frm contained 289 lines of complex report logic
' **CLAUDE.md Compliance**: Form now only handles UI events, business logic moved to module
Private Sub Go_Click()
    On Error GoTo Error_Handler

    ' Validate form selections
    If Not SystemCore.ValidateReportSelection(Me) Then Exit Sub

    ' Call module to do the actual work
    ReportingSystem.GenerateWIPReports Me
    Exit Sub

Error_Handler:
    SystemCore.HandleStandardErrors Err.Number, "Go_Click", "fwip"
End Sub

Private Sub UserForm_Initialize()
    ' Initialize the form when it loads
    fwip.Label1.Caption = "Ready - Select report types and click Go"
End Sub


