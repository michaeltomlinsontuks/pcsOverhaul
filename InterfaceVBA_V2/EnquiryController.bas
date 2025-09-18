Attribute VB_Name = "EnquiryController"
Option Explicit

Public Function CreateNewEnquiry(ByRef EnquiryInfo As EnquiryData) As Boolean
    Dim EnquiryNumber As String
    Dim TemplatePath As String
    Dim NewFilePath As String
    Dim TemplateWB As Workbook
    Dim SearchRecord As SearchRecord

    On Error GoTo Error_Handler

    EnquiryNumber = NumberGenerator.GetNextEnquiryNumber()
    If EnquiryNumber = "" Then
        CreateNewEnquiry = False
        Exit Function
    End If

    EnquiryInfo.EnquiryNumber = EnquiryNumber
    EnquiryInfo.DateCreated = Now

    TemplatePath = FileManager.GetRootPath & "\Templates\_Enq.xls"
    NewFilePath = FileManager.GetRootPath & "\Enquiries\" & EnquiryNumber & ".xls"

    If Not FileManager.FileExists(TemplatePath) Then
        ErrorHandler.LogError ERR_FILE_NOT_FOUND, "Enquiry template not found: " & TemplatePath, "CreateNewEnquiry", "EnquiryController"
        CreateNewEnquiry = False
        Exit Function
    End If

    Set TemplateWB = FileManager.SafeOpenWorkbook(TemplatePath)
    If TemplateWB Is Nothing Then
        CreateNewEnquiry = False
        Exit Function
    End If

    PopulateEnquiryTemplate TemplateWB, EnquiryInfo

    TemplateWB.SaveAs NewFilePath
    FileManager.SafeCloseWorkbook TemplateWB

    EnquiryInfo.FilePath = NewFilePath

    SearchRecord = SearchService.CreateSearchRecord(rtEnquiry, EnquiryNumber, EnquiryInfo.CustomerName, EnquiryInfo.ComponentDescription, NewFilePath, EnquiryInfo.SearchKeywords)
    SearchService.UpdateSearchDatabase SearchRecord

    CreateNewEnquiry = True
    Exit Function

Error_Handler:
    If Not TemplateWB Is Nothing Then FileManager.SafeCloseWorkbook TemplateWB, False
    ErrorHandler.HandleStandardErrors Err.Number, "CreateNewEnquiry", "EnquiryController"
    CreateNewEnquiry = False
End Function

Private Sub PopulateEnquiryTemplate(ByVal wb As Workbook, ByRef EnquiryInfo As EnquiryData)
    Dim ws As Worksheet

    On Error GoTo Error_Handler

    Set ws = wb.Worksheets(1)

    With ws
        .Cells(2, 2).Value = EnquiryInfo.EnquiryNumber
        .Cells(3, 2).Value = EnquiryInfo.CustomerName
        .Cells(4, 2).Value = EnquiryInfo.ContactPerson
        .Cells(5, 2).Value = EnquiryInfo.CompanyPhone
        .Cells(6, 2).Value = EnquiryInfo.CompanyFax
        .Cells(7, 2).Value = EnquiryInfo.Email
        .Cells(8, 2).Value = EnquiryInfo.ComponentDescription
        .Cells(9, 2).Value = EnquiryInfo.ComponentCode
        .Cells(10, 2).Value = EnquiryInfo.MaterialGrade
        .Cells(11, 2).Value = EnquiryInfo.Quantity
        .Cells(12, 2).Value = EnquiryInfo.DateCreated
    End With

    Exit Sub

Error_Handler:
    ErrorHandler.LogError Err.Number, Err.Description, "PopulateEnquiryTemplate", "EnquiryController"
End Sub

Public Function LoadEnquiry(ByVal FilePath As String) As EnquiryData
    Dim EnquiryWB As Workbook
    Dim ws As Worksheet
    Dim EnquiryInfo As EnquiryData

    On Error GoTo Error_Handler

    Set EnquiryWB = FileManager.SafeOpenWorkbook(FilePath)
    If EnquiryWB Is Nothing Then
        Exit Function
    End If

    Set ws = EnquiryWB.Worksheets(1)

    With EnquiryInfo
        .EnquiryNumber = ws.Cells(2, 2).Value
        .CustomerName = ws.Cells(3, 2).Value
        .ContactPerson = ws.Cells(4, 2).Value
        .CompanyPhone = ws.Cells(5, 2).Value
        .CompanyFax = ws.Cells(6, 2).Value
        .Email = ws.Cells(7, 2).Value
        .ComponentDescription = ws.Cells(8, 2).Value
        .ComponentCode = ws.Cells(9, 2).Value
        .MaterialGrade = ws.Cells(10, 2).Value
        .Quantity = ws.Cells(11, 2).Value
        .DateCreated = ws.Cells(12, 2).Value
        .FilePath = FilePath
    End With

    FileManager.SafeCloseWorkbook EnquiryWB, False
    LoadEnquiry = EnquiryInfo
    Exit Function

Error_Handler:
    If Not EnquiryWB Is Nothing Then FileManager.SafeCloseWorkbook EnquiryWB, False
    ErrorHandler.HandleStandardErrors Err.Number, "LoadEnquiry", "EnquiryController"
End Function

Public Function UpdateEnquiry(ByRef EnquiryInfo As EnquiryData) As Boolean
    Dim EnquiryWB As Workbook

    On Error GoTo Error_Handler

    Set EnquiryWB = FileManager.SafeOpenWorkbook(EnquiryInfo.FilePath)
    If EnquiryWB Is Nothing Then
        UpdateEnquiry = False
        Exit Function
    End If

    PopulateEnquiryTemplate EnquiryWB, EnquiryInfo

    EnquiryWB.Save
    FileManager.SafeCloseWorkbook EnquiryWB

    UpdateEnquiry = True
    Exit Function

Error_Handler:
    If Not EnquiryWB Is Nothing Then FileManager.SafeCloseWorkbook EnquiryWB, False
    ErrorHandler.HandleStandardErrors Err.Number, "UpdateEnquiry", "EnquiryController"
    UpdateEnquiry = False
End Function

Public Function CreateNewCustomer(ByVal CustomerName As String) As Boolean
    Dim TemplatePath As String
    Dim NewFilePath As String
    Dim TemplateWB As Workbook

    On Error GoTo Error_Handler

    TemplatePath = FileManager.GetRootPath & "\Templates\_client.xls"
    NewFilePath = FileManager.GetRootPath & "\Customers\" & CustomerName & ".xls"

    If FileManager.FileExists(NewFilePath) Then
        CreateNewCustomer = True
        Exit Function
    End If

    Set TemplateWB = FileManager.SafeOpenWorkbook(TemplatePath)
    If TemplateWB Is Nothing Then
        CreateNewCustomer = False
        Exit Function
    End If

    TemplateWB.Worksheets(1).Cells(1, 1).Value = CustomerName
    TemplateWB.SaveAs NewFilePath
    FileManager.SafeCloseWorkbook TemplateWB

    CreateNewCustomer = True
    Exit Function

Error_Handler:
    If Not TemplateWB Is Nothing Then FileManager.SafeCloseWorkbook TemplateWB, False
    ErrorHandler.HandleStandardErrors Err.Number, "CreateNewCustomer", "EnquiryController"
    CreateNewCustomer = False
End Function

Public Function ValidateEnquiryData(ByRef EnquiryInfo As EnquiryData) As String
    Dim ValidationErrors As String

    If Trim(EnquiryInfo.CustomerName) = "" Then
        ValidationErrors = ValidationErrors & "Customer name is required." & vbCrLf
    End If

    If Trim(EnquiryInfo.ComponentDescription) = "" Then
        ValidationErrors = ValidationErrors & "Component description is required." & vbCrLf
    End If

    If EnquiryInfo.Quantity <= 0 Then
        ValidationErrors = ValidationErrors & "Quantity must be greater than zero." & vbCrLf
    End If

    ValidateEnquiryData = ValidationErrors
End Function