# Subsystem 4: Quote Management - PCS Original System

## 🎯 **Subsystem Purpose**

The Quote Management subsystem handles the **conversion of enquiries to formal quotes** with pricing and lead time information. This subsystem manages the transition from initial customer interest (enquiry) to formal quotation with commercial terms.

**Responsibility**: Quote creation from enquiries, pricing data entry, file movement from Enquiries/ to Quotes/ directory, and search database updates.

---

## 📁 **Module and Form Inventory**

### **Primary Components**

| Component | Type | Lines | Purpose | Dependencies |
|-----------|------|-------|---------|-------------|
| `FQuote.frm` | UserForm | 150+ | Quote creation and management | Calc_Numbers, SaveSearchCode, Open_Book |

**Note**: Quote management is primarily form-driven, relying on shared modules for number generation and database updates.

---

## 📋 **Quote Creation Workflow**

### **FQuote.frm** - Quote Management Form

#### **Form Purpose and Context**
```vba
' Quote form is opened from Main interface when user selects:
' 1. Enquiry file from Enquiries/ directory
' 2. "Make Quote" button on Main.frm
' 3. Form loads existing enquiry data and allows quote creation
```

#### **Key Controls and Data Fields**

##### **Pre-Populated from Enquiry**
- **Enquiry_Number** - Source enquiry (read-only)
- **Customer** - Customer name from enquiry
- **Component_Description** - Part description from enquiry
- **Component_Quantity** - Quantity from enquiry
- **Component_Code** - Part number from enquiry
- **Component_Grade** - Material specification from enquiry

##### **Quote-Specific Fields**
- **Quote_Number** - Auto-generated Q-prefix number
- **Component_Price** - Pricing information (user input)
- **Job_LeadTime** - Delivery timeframe in days
- **Quote_Date** - Auto-filled with current date
- **Quote_Notes** - Additional quote-specific information

#### **Primary Event Handlers**

##### **UserForm_Activate() - Form Initialization**
```vba
Private Sub UserForm_Activate()
    ' 1. Load enquiry data from selected file
    Dim enquiryFile As String
    enquiryFile = Main.lst.Value ' Selected from main interface

    ' 2. Read enquiry data using GetValue
    Me.Customer.Value = GetValue(enquiryPath, enquiryFile, "Admin", "Customer")
    Me.Component_Description.Value = GetValue(enquiryPath, enquiryFile, "Admin", "Component_Description")
    ' ... populate other fields from enquiry

    ' 3. Calculate and display next quote number
    Dim nextQuoteNumber As Long
    nextQuoteNumber = Calc_Numbers.Calc_Next_Number("Q")
    Me.Quote_Number.Value = "Q" & Format(nextQuoteNumber, "0000")

    ' 4. Set current date
    Me.Quote_Date.Value = Date

    ' 5. Set default lead time
    Me.Job_LeadTime.Value = 14 ' Default 14 days
End Sub
```

##### **SaveQuote_Click() - Create Quote**
```vba
Private Sub SaveQuote_Click()
    ' 1. Validate quote data
    If Not ValidateQuoteForm() Then Exit Sub

    ' 2. Confirm quote number generation
    Dim confirmedQuoteNumber As Long
    confirmedQuoteNumber = Calc_Numbers.Confirm_Next_Number("Q")

    ' 3. Create quote file from enquiry
    Dim enquiryPath As String, quotePath As String
    enquiryPath = Main.Main_MasterPath.Value & "Enquiries\" & Me.Enquiry_Number.Value & ".xls"
    quotePath = Main.Main_MasterPath.Value & "Quotes\" & "Q" & Format(confirmedQuoteNumber, "0000") & ".xls"

    ' 4. Copy enquiry file to quotes directory
    FileCopy enquiryPath, quotePath

    ' 5. Update quote file with new data
    Call PopulateQuoteFile(quotePath, confirmedQuoteNumber)

    ' 6. Update search database
    Call SaveSearchCode.SaveRowIntoSearch(Me)

    ' 7. Delete original enquiry file
    Kill enquiryPath

    ' 8. Show success message and close
    MsgBox "Quote " & "Q" & Format(confirmedQuoteNumber, "0000") & " created successfully"
    Me.Hide
End Sub
```

##### **Search_Component_code_Click() - Component Lookup**
```vba
Private Sub Search_Component_code_Click()
    ' Open component search functionality
    ' Allow user to browse price list and select components
    ' Populate component code and pricing information
End Sub
```

---

## 💰 **Pricing and Commercial Data**

### **Quote-Specific Information**

#### **Pricing Fields**
```vba
' Component pricing structure
Component_Price         ' Unit price for the component
Job_LeadTime           ' Delivery time in days
Quote_Date             ' Date quote was created
Quote_Notes            ' Additional terms and conditions
```

#### **Lead Time Calculation**
```vba
' Default lead times based on component complexity
Standard_Components = 14    ' Days for standard items
Custom_Components = 21      ' Days for custom manufacturing
Rush_Jobs = 7              ' Days for urgent requirements

' Lead time can be manually adjusted by user
```

#### **Pricing Integration**
```vba
' Price lookup from price list database
Private Sub LoadComponentPricing()
    Dim priceListPath As String
    priceListPath = Main.Main_MasterPath.Value & "price list.xls"

    ' Look up component code in price list
    Dim unitPrice As Variant
    unitPrice = GetValue(Main.Main_MasterPath.Value, "price list.xls", "Prices", componentCell)

    If Not IsEmpty(unitPrice) Then
        Me.Component_Price.Value = unitPrice
    End If
End Sub
```

---

## 📂 **File Operations and Data Flow**

### **Enquiry to Quote Conversion Process**

#### **File Movement Workflow**
```
1. Source: Enquiries/E####.xls
   ↓
2. Copy to: Quotes/Q####.xls
   ↓
3. Update quote file with new data:
   - Quote_Number
   - Component_Price
   - Job_LeadTime
   - Quote_Date
   - System_Status = "New Quote"
   ↓
4. Delete original enquiry file
   ↓
5. Update Search.xls database
```

#### **PopulateQuoteFile() - Update Quote Data**
```vba
Private Sub PopulateQuoteFile(quotePath As String, quoteNumber As Long)
    ' Open the newly created quote file
    Call Open_Book.OpenBook(quotePath, False)

    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets("Admin")

    ' Update quote-specific fields
    ws.Range("Quote_Number").Value = "Q" & Format(quoteNumber, "0000")
    ws.Range("Component_Price").Value = Me.Component_Price.Value
    ws.Range("Job_LeadTime").Value = Me.Job_LeadTime.Value
    ws.Range("Quote_Date").Value = Me.Quote_Date.Value
    ws.Range("System_Status").Value = "New Quote"
    ws.Range("Quote_Notes").Value = Me.Quote_Notes.Value

    ' Save and close
    ActiveWorkbook.Save
    ActiveWorkbook.Close
End Sub
```

### **Search Database Updates**

#### **Status Transition in Search.xls**
```vba
' Before quote creation:
System_Status = "To Quote"     ' Enquiry ready for quoting

' After quote creation:
System_Status = "New Quote"    ' Quote created, awaiting customer response
```

#### **Search Record Mapping for Quotes**
```vba
' Search database fields updated during quote creation
File_Name = "Q" & Format(quoteNumber, "0000")
Quote_Number = "Q" & Format(quoteNumber, "0000")
Component_Price = Me.Component_Price.Value
Job_LeadTime = Me.Job_LeadTime.Value
Quote_Date = Me.Quote_Date.Value
System_Status = "New Quote"

' Enquiry fields preserved:
Enquiry_Number = Original enquiry number
Customer = Customer name
Component_Description = Part description
Component_Quantity = Required quantity
```

---

## 🔄 **Integration with Other Subsystems**

### **Upstream Dependencies**

#### **Enquiry Management Integration**
```vba
' Quote form depends on existing enquiry data
1. User selects enquiry from Main.frm file listing
2. FQuote.frm opens with enquiry file reference
3. Form loads all enquiry data for quote conversion
4. User adds pricing and commercial terms
5. Quote replaces enquiry in system
```

#### **Number Generation Dependency**
```vba
' Quote numbering follows same pattern as enquiries
Dim nextQuoteNumber As Long
nextQuoteNumber = Calc_Numbers.Calc_Next_Number("Q")     ' Preview
confirmedNumber = Calc_Numbers.Confirm_Next_Number("Q")  ' Reserve
```

### **Downstream Integration**

#### **Job Management Connection**
```vba
' Quotes become source for job creation
1. Quote saved to Quotes/ directory
2. Customer response received
3. If accepted, quote moved to Archive/
4. Job creation initiated from archived quote
5. FAcceptQuote.frm processes quote acceptance
```

#### **Search System Integration**
```vba
' Quote data searchable through central system
1. All quotes indexed in Search.xls
2. Search by customer, component, price, date
3. Status tracking shows quote workflow progress
4. Historical quote data maintained
```

---

## ⚠️ **Error Handling and Validation**

### **Quote Form Validation**

#### **ValidateQuoteForm() - Data Validation**
```vba
Private Function ValidateQuoteForm() As Boolean
    ValidateQuoteForm = True

    ' Component price must be positive
    If Not IsNumeric(Me.Component_Price.Value) Or Me.Component_Price.Value <= 0 Then
        MsgBox "Please enter a valid component price"
        Me.Component_Price.SetFocus
        ValidateQuoteForm = False
        Exit Function
    End If

    ' Lead time must be positive integer
    If Not IsNumeric(Me.Job_LeadTime.Value) Or Me.Job_LeadTime.Value <= 0 Then
        MsgBox "Please enter a valid lead time in days"
        Me.Job_LeadTime.SetFocus
        ValidateQuoteForm = False
        Exit Function
    End If

    ' Quote date must be valid
    If Not IsDate(Me.Quote_Date.Value) Then
        MsgBox "Please enter a valid quote date"
        Me.Quote_Date.SetFocus
        ValidateQuoteForm = False
        Exit Function
    End If
End Function
```

### **File Operation Error Handling**

#### **Safe File Movement**
```vba
Private Function SafeCopyEnquiryToQuote(sourcePath As String, targetPath As String) As Boolean
    On Error GoTo ErrorHandler

    ' Check source file exists
    If Dir(sourcePath) = "" Then
        MsgBox "Source enquiry file not found: " & sourcePath
        SafeCopyEnquiryToQuote = False
        Exit Function
    End If

    ' Check target directory exists
    Call CheckDir(Main.Main_MasterPath.Value & "Quotes\")

    ' Perform file copy
    FileCopy sourcePath, targetPath

    ' Verify copy successful
    If Dir(targetPath) = "" Then
        MsgBox "Failed to create quote file"
        SafeCopyEnquiryToQuote = False
        Exit Function
    End If

    SafeCopyEnquiryToQuote = True
    Exit Function

ErrorHandler:
    MsgBox "Error copying enquiry to quote: " & Err.Description
    SafeCopyEnquiryToQuote = False
End Function
```

#### **Safe File Deletion**
```vba
Private Sub SafeDeleteEnquiry(enquiryPath As String)
    On Error GoTo ErrorHandler

    ' Verify quote was created successfully first
    Dim quotePath As String
    quotePath = Main.Main_MasterPath.Value & "Quotes\" & Me.Quote_Number.Value & ".xls"

    If Dir(quotePath) <> "" Then
        ' Quote exists - safe to delete enquiry
        Kill enquiryPath
    Else
        MsgBox "Warning: Quote file not found - enquiry not deleted"
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Error deleting enquiry file: " & Err.Description
    ' Continue execution - manual cleanup may be required
End Sub
```

---

## 💼 **Business Rules and Workflow**

### **Quote Approval Workflow**

#### **Quote Status Progression**
```
1. "New Quote"        → Quote created, sent to customer
2. "Quote Submitted"  → Customer received quote
3. "Quote Accepted"   → Customer accepts, ready for job creation
4. "Quote Rejected"   → Customer declines quote
5. "Quote Expired"    → Quote validity period ended
```

#### **Quote Follow-up Process**
```vba
' Manual status updates through Main interface
1. User reviews quotes in Quotes/ directory
2. Customer response received (phone, email, etc.)
3. User updates quote status in file
4. Search database reflects current status
5. Accepted quotes moved to Archive/ for job creation
```

### **Pricing and Lead Time Guidelines**

#### **Standard Pricing Rules**
```vba
' Component pricing considerations
Base_Price = Material_Cost + Labor_Cost + Overhead
Margin = Base_Price * Profit_Margin
Total_Price = Base_Price + Margin

' Lead time factors
Processing_Time = 2-3 days     ' Order processing and setup
Manufacturing_Time = Variable   ' Based on component complexity
Buffer_Time = 2-3 days         ' Contingency buffer
```

#### **Quote Validity Period**
```vba
' Standard quote terms
Quote_Validity = 30 days       ' Standard validity period
Price_Hold_Period = 60 days    ' Price protection period
Delivery_Commitment = Lead_Time ' Firm delivery commitment
```

---

## 🔧 **Development Guidelines**

### **Customizing Quote Forms**

#### **Adding New Quote Fields**
```vba
' 1. Add control to FQuote.frm
' 2. Update validation function
Private Function ValidateQuoteForm() As Boolean
    ' Add validation for new field
    If NewField.Value = "" Then
        MsgBox "New field is required"
        NewField.SetFocus
        ValidateQuoteForm = False
        Exit Function
    End If
End Function

' 3. Update file population
Private Sub PopulateQuoteFile(quotePath As String, quoteNumber As Long)
    ' Add new field to Admin sheet
    ws.Range("NewFieldCell").Value = Me.NewField.Value
End Sub

' 4. Update search database
' Modify SaveSearchCode.bas to include new field
```

#### **Pricing Integration Enhancement**
```vba
' Enhanced price lookup with multiple pricing tiers
Private Sub LoadTieredPricing()
    Dim quantity As Long
    quantity = Me.Component_Quantity.Value

    ' Determine pricing tier based on quantity
    Dim priceTier As String
    Select Case quantity
        Case 1 To 10: priceTier = "Small"
        Case 11 To 100: priceTier = "Medium"
        Case Is > 100: priceTier = "Large"
    End Select

    ' Load appropriate price
    Dim priceCell As String
    priceCell = "Price_" & priceTier
    Me.Component_Price.Value = GetValue(pricePath, priceFile, "Prices", priceCell)
End Sub
```

### **Testing Quote Workflow**

#### **End-to-End Quote Testing**
```vba
Sub TestQuoteWorkflow()
    ' 1. Create test enquiry
    ' 2. Open quote form with test data
    FQuote.Show

    ' 3. Populate quote fields
    FQuote.Component_Price.Value = "100.00"
    FQuote.Job_LeadTime.Value = "14"
    FQuote.Quote_Date.Value = Date

    ' 4. Test validation
    If ValidateQuoteForm() Then
        MsgBox "Quote validation passed"
    End If

    ' 5. Test file operations (with test directory)
    ' 6. Verify search database updates
End Sub
```

---

## 🔍 **Next Steps**

After understanding Quote Management:

1. **Study [Job Management](SUBSYSTEM_05_JOB_MANAGEMENT.md)** - See how quotes become jobs
2. **Review [Search Database](SUBSYSTEM_08_SEARCH_DATA.md)** - Understand quote indexing and status tracking
3. **Examine [Interface Navigation](SUBSYSTEM_06_INTERFACE_NAVIGATION.md)** - See how quotes are displayed and managed
4. **Practice Quote Customization** - Add pricing tiers or delivery options
5. **Test Integration** - Follow complete workflow from enquiry through quote to job

**Ready for job management? Continue to [Job Management Subsystem](SUBSYSTEM_05_JOB_MANAGEMENT.md)**