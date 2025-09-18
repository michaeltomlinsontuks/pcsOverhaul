Attribute VB_Name = "Create_Test_Templates"

' ====================================================================
' CREATE TEST TEMPLATES
' Creates minimal template files needed for testing the VBA framework
' ====================================================================

Public Sub CreateAllTestTemplates()

    Dim basePath As String
    basePath = InputBox("Enter your test base path (where directories were created):", "Create Templates", "C:\PCS_Test\")

    If basePath = "" Then Exit Sub
    If Right(basePath, 1) <> "\" Then basePath = basePath & "\"

    Application.DisplayAlerts = False

    ' Create core data files
    CreateSearchTemplate basePath
    CreateWIPTemplate basePath
    CreateHistoryTemplates basePath

    ' Create enquiry template
    CreateEnquiryTemplate basePath

    Application.DisplayAlerts = True

    MsgBox "Test templates created successfully!" & vbCrLf & vbCrLf & _
           "You can now run RunAllTests() to validate the framework.", vbInformation, "Templates Created"

End Sub

' ====================================================================
' CREATE SEARCH.XLS TEMPLATE
' ====================================================================
Private Sub CreateSearchTemplate(basePath As String)

    Dim wb As Workbook
    Set wb = Workbooks.Add

    ' Create search sheet structure
    With wb.Worksheets(1)
        .Name = "search"

        ' Add headers for search functionality
        .Range("A1").Value = "File_Name"
        .Range("B1").Value = "System_Status"
        .Range("C1").Value = "Customer"
        .Range("D1").Value = "Component_Description"
        .Range("E1").Value = "Date_Created"
        .Range("F1").Value = "Job_Number"
        .Range("G1").Value = "Quote_Number"
        .Range("H1").Value = "Enquiry_Number"
        .Range("I1").Value = "Invoice_Number"
        .Range("J1").Value = "Invoice_Date"

        ' Format headers
        .Range("A1:J1").Font.Bold = True
        .Range("A1:J1").Interior.Color = RGB(220, 220, 220)

        ' Add some sample data
        .Range("A2").Value = "TEST001_Sample_Job"
        .Range("B2").Value = "IN PROGRESS"
        .Range("C2").Value = "Test Customer"
        .Range("D2").Value = "Sample Component"
        .Range("E2").Value = Now()
    End With

    wb.SaveAs basePath & "Search.xls"
    wb.Close

End Sub

' ====================================================================
' CREATE WIP.XLS TEMPLATE
' ====================================================================
Private Sub CreateWIPTemplate(basePath As String)

    Dim wb As Workbook
    Set wb = Workbooks.Add

    With wb.Worksheets(1)
        .Name = "WIP"

        ' Add WIP tracking headers
        .Range("A1").Value = "Date"
        .Range("B1").Value = "Customer"
        .Range("C1").Value = "Job_Number"
        .Range("D1").Value = "Description"
        .Range("E1").Value = "Status"
        .Range("F1").Value = "Due_Date"
        .Range("G1").Value = "Operator"

        ' Format headers
        .Range("A1:G1").Font.Bold = True
        .Range("A1:G1").Interior.Color = RGB(200, 230, 200)

        ' Add sample WIP entry
        .Range("A2").Value = Now()
        .Range("B2").Value = "Test Customer"
        .Range("C2").Value = "J2024-001"
        .Range("D2").Value = "Sample WIP Job"
        .Range("E2").Value = "Quote Accepted"
        .Range("F2").Value = Now() + 7
    End With

    wb.SaveAs basePath & "WIP.xls"
    wb.Close

End Sub

' ====================================================================
' CREATE HISTORY TEMPLATES
' ====================================================================
Private Sub CreateHistoryTemplates(basePath As String)

    ' Search History
    CreateHistoryFile basePath & "search History.xls", "Search History"

    ' Job History
    CreateHistoryFile basePath & "Job History.xls", "Job History"

    ' Quote History
    CreateHistoryFile basePath & "Quote History.xls", "Quote History"

End Sub

Private Sub CreateHistoryFile(filePath As String, sheetName As String)

    Dim wb As Workbook
    Set wb = Workbooks.Add

    With wb.Worksheets(1)
        .Name = sheetName

        ' Basic history structure
        .Range("A1").Value = "Date"
        .Range("B1").Value = "Action"
        .Range("C1").Value = "File_Name"
        .Range("D1").Value = "Details"

        ' Format headers
        .Range("A1:D1").Font.Bold = True
        .Range("A1:D1").Interior.Color = RGB(255, 240, 200)

        ' Sample entry
        .Range("A2").Value = Now()
        .Range("B2").Value = "System Test"
        .Range("C2").Value = "TEST001"
        .Range("D2").Value = "Template created for testing"
    End With

    wb.SaveAs filePath
    wb.Close

End Sub

' ====================================================================
' CREATE ENQUIRY TEMPLATE (_ENQ.XLS)
' ====================================================================
Private Sub CreateEnquiryTemplate(basePath As String)

    Dim wb As Workbook
    Set wb = Workbooks.Add

    ' Create Admin sheet
    With wb.Worksheets(1)
        .Name = "Admin"

        ' Create the admin data structure that the VBA code expects
        .Range("A1").Value = "File_Name"
        .Range("B1").Value = ""
        .Range("A2").Value = "System_Status"
        .Range("B2").Value = "New Enquiry"
        .Range("A3").Value = "Customer"
        .Range("B3").Value = ""
        .Range("A4").Value = "Component_Description"
        .Range("B4").Value = ""
        .Range("A5").Value = "Component_Quantity"
        .Range("B5").Value = ""
        .Range("A6").Value = "Component_Grade"
        .Range("B6").Value = ""
        .Range("A7").Value = "Job_Number"
        .Range("B7").Value = ""
        .Range("A8").Value = "Quote_Number"
        .Range("B8").Value = ""
        .Range("A9").Value = "Enquiry_Number"
        .Range("B9").Value = ""
        .Range("A10").Value = "Invoice_Number"
        .Range("B10").Value = ""
        .Range("A11").Value = "Invoice_Date"
        .Range("B11").Value = ""

        ' Important: Cell B88 is used for status checking
        .Range("B88").Value = "New Enquiry"

        ' Format the admin section
        .Range("A1:A11").Font.Bold = True
        .Range("A1:B11").Borders.LineStyle = xlContinuous
    End With

    ' Create Job Card sheet
    wb.Worksheets.Add After:=wb.Worksheets(1)
    With wb.Worksheets(2)
        .Name = "Job Card"

        .Range("A1").Value = "JOB CARD TEMPLATE"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14

        .Range("A3").Value = "Customer:"
        .Range("A4").Value = "Job Number:"
        .Range("A5").Value = "Description:"
        .Range("A6").Value = "Quantity:"
        .Range("A7").Value = "Due Date:"

        ' Create named ranges that the VBA code expects
        .Range("B4").Name = "Job_Number"
        .Range("B3").Name = "Customer"
        .Range("B88").Name = "system_Status"
        .Range("B10").Name = "Invoice_Number"
    End With

    wb.SaveAs basePath & "templates\_Enq.xls"
    wb.Close

End Sub

' ====================================================================
' CREATE OTHER REQUIRED TEMPLATES
' ====================================================================
Public Sub CreateAdditionalTemplates()

    Dim basePath As String
    basePath = InputBox("Enter your test base path:", "Create Additional Templates", "C:\PCS_Test\")

    If basePath = "" Then Exit Sub
    If Right(basePath, 1) <> "\" Then basePath = basePath & "\"

    Application.DisplayAlerts = False

    ' Create client template
    CreateClientTemplate basePath

    ' Create price list template
    CreatePriceListTemplate basePath

    ' Create component grades template
    CreateComponentGradesTemplate basePath

    Application.DisplayAlerts = True

    MsgBox "Additional templates created!", vbInformation

End Sub

Private Sub CreateClientTemplate(basePath As String)

    Dim wb As Workbook
    Set wb = Workbooks.Add

    With wb.Worksheets(1)
        .Range("A1").Value = "Customer Information Template"
        .Range("A1").Font.Bold = True

        .Range("A3").Value = "Company_Name"
        .Range("B3").Value = ""
        .Range("B3").Name = "company_Name"

        .Range("A4").Value = "Contact_Person"
        .Range("A5").Value = "Phone"
        .Range("A6").Value = "Email"
        .Range("A7").Value = "Address"
    End With

    wb.SaveAs basePath & "templates\_client.xls"
    wb.Close

End Sub

Private Sub CreatePriceListTemplate(basePath As String)

    Dim wb As Workbook
    Set wb = Workbooks.Add

    ' Create Component_Descriptions sheet
    With wb.Worksheets(1)
        .Name = "Component_Descriptions"

        .Range("A1").Value = "Component Code"
        .Range("B1").Value = "Description"
        .Range("C1").Value = "Unit Price"

        .Range("A2").Value = "COMP001"
        .Range("B2").Value = "Standard Component"
        .Range("C2").Value = 100

        .Range("A3").Value = "COMP002"
        .Range("B3").Value = "Premium Component"
        .Range("C3").Value = 150
    End With

    wb.SaveAs basePath & "templates\price list.xls"
    wb.Close

End Sub

Private Sub CreateComponentGradesTemplate(basePath As String)

    Dim wb As Workbook
    Set wb = Workbooks.Add

    With wb.Worksheets(1)
        .Range("A1").Value = "Grade"
        .Range("A2").Value = "Standard"
        .Range("A3").Value = "Premium"
        .Range("A4").Value = "Custom"
    End With

    wb.SaveAs basePath & "templates\Component_Grades.xls"
    wb.Close

End Sub