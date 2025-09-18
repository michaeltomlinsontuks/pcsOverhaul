Attribute VB_Name = "QuoteController"
Option Explicit

Public Function CreateQuoteFromEnquiry(ByVal EnquiryFilePath As String, ByRef QuoteInfo As QuoteData) As Boolean
    Dim QuoteNumber As String
    Dim NewFilePath As String
    Dim EnquiryWB As Workbook
    Dim SearchRecord As SearchRecord

    On Error GoTo Error_Handler

    QuoteNumber = NumberGenerator.GetNextQuoteNumber()
    If QuoteNumber = "" Then
        CreateQuoteFromEnquiry = False
        Exit Function
    End If

    QuoteInfo.QuoteNumber = QuoteNumber
    QuoteInfo.DateCreated = Now

    Set EnquiryWB = FileManager.SafeOpenWorkbook(EnquiryFilePath)
    If EnquiryWB Is Nothing Then
        CreateQuoteFromEnquiry = False
        Exit Function
    End If

    NewFilePath = FileManager.GetRootPath & "\Quotes\" & QuoteNumber & ".xls"

    PopulateQuoteFromEnquiry EnquiryWB, QuoteInfo

    EnquiryWB.SaveAs NewFilePath
    FileManager.SafeCloseWorkbook EnquiryWB

    QuoteInfo.FilePath = NewFilePath

    SearchRecord = SearchService.CreateSearchRecord(rtQuote, QuoteNumber, QuoteInfo.CustomerName, QuoteInfo.ComponentDescription, NewFilePath)
    SearchService.UpdateSearchDatabase SearchRecord

    CreateQuoteFromEnquiry = True
    Exit Function

Error_Handler:
    If Not EnquiryWB Is Nothing Then FileManager.SafeCloseWorkbook EnquiryWB, False
    ErrorHandler.HandleStandardErrors Err.Number, "CreateQuoteFromEnquiry", "QuoteController"
    CreateQuoteFromEnquiry = False
End Function

Private Sub PopulateQuoteFromEnquiry(ByVal wb As Workbook, ByRef QuoteInfo As QuoteData)
    Dim ws As Worksheet

    On Error GoTo Error_Handler

    Set ws = wb.Worksheets(1)

    With ws
        .Cells(2, 2).Value = QuoteInfo.QuoteNumber
        .Cells(13, 2).Value = QuoteInfo.UnitPrice
        .Cells(14, 2).Value = QuoteInfo.TotalPrice
        .Cells(15, 2).Value = QuoteInfo.LeadTime
        .Cells(16, 2).Value = QuoteInfo.ValidUntil
        .Cells(17, 2).Value = QuoteInfo.DateCreated
        .Cells(18, 2).Value = QuoteInfo.Status
    End With

    Exit Sub

Error_Handler:
    ErrorHandler.LogError Err.Number, Err.Description, "PopulateQuoteFromEnquiry", "QuoteController"
End Sub

Public Function LoadQuote(ByVal FilePath As String) As QuoteData
    Dim QuoteWB As Workbook
    Dim ws As Worksheet
    Dim QuoteInfo As QuoteData

    On Error GoTo Error_Handler

    Set QuoteWB = FileManager.SafeOpenWorkbook(FilePath)
    If QuoteWB Is Nothing Then
        Exit Function
    End If

    Set ws = QuoteWB.Worksheets(1)

    With QuoteInfo
        .QuoteNumber = ws.Cells(2, 2).Value
        .EnquiryNumber = ws.Cells(2, 2).Value
        .CustomerName = ws.Cells(3, 2).Value
        .ComponentDescription = ws.Cells(8, 2).Value
        .ComponentCode = ws.Cells(9, 2).Value
        .MaterialGrade = ws.Cells(10, 2).Value
        .Quantity = ws.Cells(11, 2).Value
        .UnitPrice = ws.Cells(13, 2).Value
        .TotalPrice = ws.Cells(14, 2).Value
        .LeadTime = ws.Cells(15, 2).Value
        .ValidUntil = ws.Cells(16, 2).Value
        .DateCreated = ws.Cells(17, 2).Value
        .Status = ws.Cells(18, 2).Value
        .FilePath = FilePath
    End With

    FileManager.SafeCloseWorkbook QuoteWB, False
    LoadQuote = QuoteInfo
    Exit Function

Error_Handler:
    If Not QuoteWB Is Nothing Then FileManager.SafeCloseWorkbook QuoteWB, False
    ErrorHandler.HandleStandardErrors Err.Number, "LoadQuote", "QuoteController"
End Function

Public Function UpdateQuote(ByRef QuoteInfo As QuoteData) As Boolean
    Dim QuoteWB As Workbook

    On Error GoTo Error_Handler

    Set QuoteWB = FileManager.SafeOpenWorkbook(QuoteInfo.FilePath)
    If QuoteWB Is Nothing Then
        UpdateQuote = False
        Exit Function
    End If

    PopulateQuoteFromEnquiry QuoteWB, QuoteInfo

    QuoteWB.Save
    FileManager.SafeCloseWorkbook QuoteWB

    UpdateQuote = True
    Exit Function

Error_Handler:
    If Not QuoteWB Is Nothing Then FileManager.SafeCloseWorkbook QuoteWB, False
    ErrorHandler.HandleStandardErrors Err.Number, "UpdateQuote", "QuoteController"
    UpdateQuote = False
End Function

Public Function AcceptQuote(ByVal QuoteFilePath As String) As String
    Dim QuoteInfo As QuoteData
    Dim JobInfo As JobData

    On Error GoTo Error_Handler

    QuoteInfo = LoadQuote(QuoteFilePath)
    If QuoteInfo.QuoteNumber = "" Then
        AcceptQuote = ""
        Exit Function
    End If

    With JobInfo
        .QuoteNumber = QuoteInfo.QuoteNumber
        .CustomerName = QuoteInfo.CustomerName
        .ComponentDescription = QuoteInfo.ComponentDescription
        .ComponentCode = QuoteInfo.ComponentCode
        .MaterialGrade = QuoteInfo.MaterialGrade
        .Quantity = QuoteInfo.Quantity
        .OrderValue = QuoteInfo.TotalPrice
        .Status = "Active"
    End With

    If JobController.CreateJobFromQuote(JobInfo) Then
        AcceptQuote = JobInfo.JobNumber
    Else
        AcceptQuote = ""
    End If

    Exit Function

Error_Handler:
    ErrorHandler.HandleStandardErrors Err.Number, "AcceptQuote", "QuoteController"
    AcceptQuote = ""
End Function

Public Function ValidateQuoteData(ByRef QuoteInfo As QuoteData) As String
    Dim ValidationErrors As String

    If Trim(QuoteInfo.CustomerName) = "" Then
        ValidationErrors = ValidationErrors & "Customer name is required." & vbCrLf
    End If

    If QuoteInfo.UnitPrice <= 0 Then
        ValidationErrors = ValidationErrors & "Unit price must be greater than zero." & vbCrLf
    End If

    If QuoteInfo.Quantity <= 0 Then
        ValidationErrors = ValidationErrors & "Quantity must be greater than zero." & vbCrLf
    End If

    If QuoteInfo.ValidUntil < Date Then
        ValidationErrors = ValidationErrors & "Valid until date cannot be in the past." & vbCrLf
    End If

    ValidateQuoteData = ValidationErrors
End Function

Public Function CalculateTotalPrice(ByVal UnitPrice As Currency, ByVal Quantity As Long) As Currency
    CalculateTotalPrice = UnitPrice * Quantity
End Function