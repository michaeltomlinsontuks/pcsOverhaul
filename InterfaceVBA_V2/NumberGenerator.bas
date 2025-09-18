Attribute VB_Name = "NumberGenerator"
Option Explicit

Private Const NUMBERS_FILE As String = "Templates\number_tracking.xls"

Public Function GetNextEnquiryNumber() As String
    GetNextEnquiryNumber = GetNextNumber("E")
End Function

Public Function GetNextQuoteNumber() As String
    GetNextQuoteNumber = GetNextNumber("Q")
End Function

Public Function GetNextJobNumber() As String
    GetNextJobNumber = GetNextNumber("J")
End Function

Private Function GetNextNumber(ByVal Prefix As String) As String
    Dim NumbersWB As Workbook
    Dim NumbersWS As Worksheet
    Dim LastNumber As Long
    Dim NextNumber As Long
    Dim NumbersFile As String

    On Error GoTo Error_Handler

    NumbersFile = FileManager.GetRootPath & "\" & NUMBERS_FILE

    If Not FileManager.FileExists(NumbersFile) Then
        CreateNumbersFile NumbersFile
    End If

    Set NumbersWB = FileManager.SafeOpenWorkbook(NumbersFile)
    If NumbersWB Is Nothing Then
        GetNextNumber = ""
        Exit Function
    End If

    Set NumbersWS = NumbersWB.Worksheets(1)

    LastNumber = GetLastNumberFromSheet(NumbersWS, Prefix)
    NextNumber = LastNumber + 1

    UpdateNumberInSheet NumbersWS, Prefix, NextNumber

    NumbersWB.Save
    FileManager.SafeCloseWorkbook NumbersWB

    GetNextNumber = Prefix & Format(NextNumber, "00000")
    Exit Function

Error_Handler:
    If Not NumbersWB Is Nothing Then FileManager.SafeCloseWorkbook NumbersWB, False
    ErrorHandler.HandleStandardErrors Err.Number, "GetNextNumber", "NumberGenerator"
    GetNextNumber = ""
End Function

Private Function GetLastNumberFromSheet(ByVal ws As Worksheet, ByVal Prefix As String) As Long
    Dim i As Long

    On Error GoTo Error_Handler

    For i = 1 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If ws.Cells(i, 1).Value = Prefix Then
            GetLastNumberFromSheet = ws.Cells(i, 2).Value
            Exit Function
        End If
    Next i

    GetLastNumberFromSheet = 0
    Exit Function

Error_Handler:
    GetLastNumberFromSheet = 0
End Function

Private Sub UpdateNumberInSheet(ByVal ws As Worksheet, ByVal Prefix As String, ByVal Number As Long)
    Dim i As Long
    Dim Found As Boolean

    On Error GoTo Error_Handler

    For i = 1 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If ws.Cells(i, 1).Value = Prefix Then
            ws.Cells(i, 2).Value = Number
            ws.Cells(i, 3).Value = Now
            Found = True
            Exit For
        End If
    Next i

    If Not Found Then
        i = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        ws.Cells(i, 1).Value = Prefix
        ws.Cells(i, 2).Value = Number
        ws.Cells(i, 3).Value = Now
    End If

    Exit Sub

Error_Handler:
    ErrorHandler.LogError Err.Number, Err.Description, "UpdateNumberInSheet", "NumberGenerator"
End Sub

Private Sub CreateNumbersFile(ByVal FilePath As String)
    Dim NewWB As Workbook
    Dim NewWS As Worksheet

    On Error GoTo Error_Handler

    Set NewWB = Workbooks.Add
    Set NewWS = NewWB.Worksheets(1)

    With NewWS
        .Name = "NumberTracking"
        .Cells(1, 1).Value = "Prefix"
        .Cells(1, 2).Value = "Last Number"
        .Cells(1, 3).Value = "Last Updated"

        .Cells(2, 1).Value = "E"
        .Cells(2, 2).Value = 0
        .Cells(2, 3).Value = Now

        .Cells(3, 1).Value = "Q"
        .Cells(3, 2).Value = 0
        .Cells(3, 3).Value = Now

        .Cells(4, 1).Value = "J"
        .Cells(4, 2).Value = 0
        .Cells(4, 3).Value = Now

        .Range("A1:C1").Font.Bold = True
        .Columns("A:C").AutoFit
    End With

    NewWB.SaveAs FilePath
    NewWB.Close
    Set NewWB = Nothing

    Exit Sub

Error_Handler:
    If Not NewWB Is Nothing Then
        NewWB.Close SaveChanges:=False
        Set NewWB = Nothing
    End If
    ErrorHandler.HandleStandardErrors Err.Number, "CreateNumbersFile", "NumberGenerator"
End Sub

Public Function ValidateNumber(ByVal Number As String, ByVal ExpectedPrefix As String) As Boolean
    If Len(Number) < 6 Then
        ValidateNumber = False
        Exit Function
    End If

    If Left(Number, 1) <> ExpectedPrefix Then
        ValidateNumber = False
        Exit Function
    End If

    If Not IsNumeric(Mid(Number, 2)) Then
        ValidateNumber = False
        Exit Function
    End If

    ValidateNumber = True
End Function

Public Function ReserveNumber(ByVal Prefix As String) As String
    ReserveNumber = GetNextNumber(Prefix)
End Function

Public Function ConfirmNumberUsage(ByVal Number As String) As Boolean
    ConfirmNumberUsage = True
End Function