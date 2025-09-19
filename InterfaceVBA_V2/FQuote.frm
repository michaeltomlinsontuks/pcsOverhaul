VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FQuote
   Caption         =   "MEM: Quote"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11220
   OleObjectBlob   =   "FQuote.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FQuote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private CurrentEnquiryPath As String

Private Sub SaveQuote_Click()
    On Error GoTo Error_Handler

    If SaveCurrentQuote() Then
        MsgBox "Quote saved successfully.", vbInformation
        Unload Me
    End If
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "SaveQuote_Click", "FQuote"
End Sub

Private Sub Cancel_Click()
    Unload Me
End Sub

Private Sub UnitPrice_Change()
    On Error GoTo Error_Handler

    CalculateTotalPrice
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "UnitPrice_Change", "FQuote"
End Sub

Private Sub Quantity_Change()
    On Error GoTo Error_Handler

    CalculateTotalPrice
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "Quantity_Change", "FQuote"
End Sub

Private Function SaveCurrentQuote() As Boolean
    Dim QuoteInfo As QuoteData
    Dim ValidationErrors As String

    On Error GoTo Error_Handler

    With QuoteInfo
        .EnquiryNumber = Trim(Me.Enquiry_Number.Value)
        .CustomerName = Trim(Me.Customer.Value)
        .ComponentDescription = Trim(Me.Component_Description.Value)
        .ComponentCode = Trim(Me.Component_Code.Value)
        .MaterialGrade = Trim(Me.Component_Grade.Value)

        If IsNumeric(Me.Component_Quantity.Value) Then
            .Quantity = CLng(Me.Component_Quantity.Value)
        Else
            .Quantity = 0
        End If

        If IsNumeric(Me.Unit_Price.Value) Then
            .UnitPrice = CCur(Me.Unit_Price.Value)
        Else
            .UnitPrice = 0
        End If

        If IsNumeric(Me.Total_Price.Value) Then
            .TotalPrice = CCur(Me.Total_Price.Value)
        Else
            .TotalPrice = 0
        End If

        .LeadTime = Trim(Me.Lead_Time.Value)

        If IsDate(Me.Valid_Until.Value) Then
            .ValidUntil = CDate(Me.Valid_Until.Value)
        Else
            .ValidUntil = DateAdd("d", 30, Now)
        End If

        .Status = "Active"
    End With

    ValidationErrors = BusinessController.ValidateQuoteData(QuoteInfo)
    If ValidationErrors <> "" Then
        MsgBox "Please correct the following errors:" & vbCrLf & vbCrLf & ValidationErrors, vbExclamation
        SaveCurrentQuote = False
        Exit Function
    End If

    SaveCurrentQuote = BusinessController.CreateQuoteFromEnquiry(CurrentEnquiryPath, QuoteInfo)

    If SaveCurrentQuote Then
        Me.Quote_Number.Value = QuoteInfo.QuoteNumber
        Me.File_Name.Value = QuoteInfo.QuoteNumber
        MsgBox "The Quote Number is: " & QuoteInfo.QuoteNumber, vbInformation
    End If
    Exit Function

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "SaveCurrentQuote", "FQuote"
    SaveCurrentQuote = False
End Function

Public Sub LoadFromEnquiry(ByVal EnquiryFileName As String)
    Dim EnquiryPath As String
    Dim EnquiryInfo As EnquiryData

    On Error GoTo Error_Handler

    EnquiryPath = DataManager.GetRootPath & "\Enquiries\" & EnquiryFileName & ".xls"
    CurrentEnquiryPath = EnquiryPath

    EnquiryInfo = BusinessController.LoadEnquiry(EnquiryPath)

    If EnquiryInfo.EnquiryNumber <> "" Then
        With Me
            .Enquiry_Number.Value = EnquiryInfo.EnquiryNumber
            .Customer.Value = EnquiryInfo.CustomerName
            .Component_Description.Value = EnquiryInfo.ComponentDescription
            .Component_Code.Value = EnquiryInfo.ComponentCode
            .Component_Grade.Value = EnquiryInfo.MaterialGrade
            .Component_Quantity.Value = EnquiryInfo.Quantity
            .Enquiry_Date.Caption = Format(EnquiryInfo.DateCreated, "dd mmm yyyy")
            .Quote_Date.Caption = Format(Now, "dd mmm yyyy")

            .Unit_Price.Value = 0
            .Total_Price.Value = 0
            .Lead_Time.Value = ""
            .Valid_Until.Value = Format(DateAdd("d", 30, Now), "dd/mm/yyyy")
        End With
    End If
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "LoadFromEnquiry", "FQuote"
End Sub

Public Sub LoadQuote(ByVal QuoteFileName As String)
    Dim QuotePath As String
    Dim QuoteInfo As QuoteData

    On Error GoTo Error_Handler

    QuotePath = DataManager.GetRootPath & "\Quotes\" & QuoteFileName & ".xls"
    QuoteInfo = BusinessController.LoadQuote(QuotePath)

    If QuoteInfo.QuoteNumber <> "" Then
        With Me
            .Quote_Number.Value = QuoteInfo.QuoteNumber
            .Enquiry_Number.Value = QuoteInfo.EnquiryNumber
            .Customer.Value = QuoteInfo.CustomerName
            .Component_Description.Value = QuoteInfo.ComponentDescription
            .Component_Code.Value = QuoteInfo.ComponentCode
            .Component_Grade.Value = QuoteInfo.MaterialGrade
            .Component_Quantity.Value = QuoteInfo.Quantity
            .Unit_Price.Value = QuoteInfo.UnitPrice
            .Total_Price.Value = QuoteInfo.TotalPrice
            .Lead_Time.Value = QuoteInfo.LeadTime
            .Valid_Until.Value = Format(QuoteInfo.ValidUntil, "dd/mm/yyyy")
            .Quote_Date.Caption = Format(QuoteInfo.DateCreated, "dd mmm yyyy")
            .File_Name.Value = QuoteInfo.QuoteNumber
        End With
    End If
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "LoadQuote", "FQuote"
End Sub

Private Sub CalculateTotalPrice()
    Dim UnitPrice As Currency
    Dim Quantity As Long
    Dim TotalPrice As Currency

    On Error GoTo Error_Handler

    If IsNumeric(Me.Unit_Price.Value) And IsNumeric(Me.Component_Quantity.Value) Then
        UnitPrice = CCur(Me.Unit_Price.Value)
        Quantity = CLng(Me.Component_Quantity.Value)
        TotalPrice = BusinessController.CalculateTotalPrice(UnitPrice, Quantity)
        Me.Total_Price.Value = TotalPrice
    End If
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "CalculateTotalPrice", "FQuote"
End Sub

Private Sub LoadPricing()
    On Error GoTo Error_Handler

    Dim PriceListPath As String
    PriceListPath = DataManager.GetRootPath & "\Templates\Price List.xls"

    If DataManager.FileExists(PriceListPath) Then
        Dim StandardPrice As Currency
        StandardPrice = DataUtilities.GetStandardPrice(PriceListPath, Me.Component_Code.Value)

        If StandardPrice > 0 Then
            Me.Unit_Price.Value = StandardPrice
            CalculateTotalPrice
        End If
    End If
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "LoadPricing", "FQuote"
End Sub

Private Sub Component_Code_Change()
    On Error GoTo Error_Handler

    If Len(Me.Component_Code.Value) > 0 Then
        LoadPricing
    End If
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "Component_Code_Change", "FQuote"
End Sub

Private Sub ValidUntil_Click()
    On Error GoTo Error_Handler

    Dim SelectedDate As Date
    SelectedDate = ShowCalendar()

    If SelectedDate <> 0 Then
        Me.Valid_Until.Value = Format(SelectedDate, "dd/mm/yyyy")
    End If
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "ValidUntil_Click", "FQuote"
End Sub

Private Function ShowCalendar() As Date
    On Error GoTo Error_Handler

    ShowCalendar = CDate(InputBox("Enter date (dd/mm/yyyy):", "Date Selection", Format(DateAdd("d", 30, Now), "dd/mm/yyyy")))
    Exit Function

Error_Handler:
    ShowCalendar = 0
End Function

Private Sub ClearForm()
    On Error GoTo Error_Handler

    Me.Quote_Number.Value = ""
    Me.Unit_Price.Value = ""
    Me.Total_Price.Value = ""
    Me.Lead_Time.Value = ""
    Me.Valid_Until.Value = Format(DateAdd("d", 30, Now), "dd/mm/yyyy")
    Me.File_Name.Value = ""
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "ClearForm", "FQuote"
End Sub

Private Sub Search_Component_code_Click()
    On Error GoTo Error_Handler

    ' Original functionality - search for component codes
    Dim PriceListPath As String
    PriceListPath = DataManager.GetRootPath & "\Templates\Price List.xls"

    If DataManager.FileExists(PriceListPath) Then
        Dim wb As Workbook
        Set wb = DataManager.SafeOpenWorkbook(PriceListPath)
        If Not wb Is Nothing Then
            Me.Hide
        End If
    Else
        MsgBox "Price list not found.", vbExclamation
    End If
    Exit Sub

Error_Handler:
    CoreFramework.HandleStandardErrors Err.Number, "Search_Component_code_Click", "FQuote"
End Sub