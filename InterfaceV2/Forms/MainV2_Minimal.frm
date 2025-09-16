VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainV2
   Caption         =   "PCS Interface V2"
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainV2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private currentFilters As DataTypes.FilterState
Private lastRefreshTime As Date
Private refreshInProgress As Boolean

Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    Me.Caption = "PCS Interface V2 - Enhanced Performance Dashboard"
    Me.Width = 12000
    Me.Height = 6000

    CacheManager.InitializeCache
    InitializeInterface
    RefreshListSmart

    MsgBox "Interface initialized successfully!", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Error initializing interface: " & Err.Description, vbCritical, "Initialization Error"
End Sub

Private Sub InitializeInterface()
    currentFilters.NewEnquiries = True
    currentFilters.QuotesToSubmit = True
    currentFilters.WIPToSequence = True
    currentFilters.JobsInWIP = True
    currentFilters.ShowArchived = False
    currentFilters.DateRangeStart = DateAdd("m", -3, Now)
    currentFilters.DateRangeEnd = Now

    refreshInProgress = False
End Sub

Public Function RefreshListSmart() As Boolean
    On Error GoTo ErrorHandler

    If refreshInProgress Then
        RefreshListSmart = False
        Exit Function
    End If

    refreshInProgress = True

    Dim fileList() As String
    fileList = FileUtilities.BuildFileList()

    Dim i As Long
    Debug.Print "=== File List ==="
    For i = LBound(fileList) To UBound(fileList)
        If fileList(i) <> "" Then
            Debug.Print fileList(i)
        End If
    Next i
    Debug.Print "=================="

    lastRefreshTime = Now
    refreshInProgress = False
    RefreshListSmart = True
    Exit Function

ErrorHandler:
    refreshInProgress = False
    MsgBox "Error refreshing list: " & Err.Description, vbExclamation, "Refresh Error"
    RefreshListSmart = False
End Function

Private Sub UserForm_Click()
    RefreshListSmart
End Sub

Private Sub UserForm_Terminate()
    CacheManager.SaveCacheToFile
End Sub