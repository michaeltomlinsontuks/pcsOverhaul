# Subsystem 3: Enquiry Management - PCS Original System

## 🎯 **Subsystem Purpose**

The Enquiry Management subsystem handles the **initial customer enquiry capture** and processing workflow. This subsystem manages customer data entry, component specification, enquiry numbering, and search database integration for the first stage of the PCS business process.

**Responsibility**: Customer enquiry data entry, validation, filing, and search database integration.

---

## 📁 **Module and Form Inventory**

### **Forms and Modules**

| Component | Type | Lines | Purpose | Dependencies |
|-----------|------|-------|---------|-------------|
| `FEnquiry.frm` | UserForm | 200+ | Primary enquiry data entry | Calc_Numbers, SaveSearchCode |
| `FrmEnquiry.frm` | UserForm | 150+ | Alternative enquiry form | Same as FEnquiry |
| `SaveSearchCode.bas` | Module | 45 | Search database updates | Search.xls, Open_Book |

**Note**: Two enquiry forms exist in the original system - likely due to version evolution. Both provide similar functionality.

---

## 📝 **Form Components**

### **FEnquiry.frm** - Primary Enquiry Form

#### **Key Controls and Data Fields**

##### **Business Information**
```vba
' Core enquiry data controls
Private Sub UserForm_Activate()
    ' Customer dropdown population
    ' Component code dropdown
    ' Material grade selection
    ' Contact person lookup
End Sub
```

**Primary Controls**:
- **Enquiry_Number** - Auto-generated E-prefix number
- **Customer** - Dropdown populated from Customers/ directory
- **ContactPerson** - Auto-populated based on customer selection
- **Component_Code** - Dropdown from price list database
- **Component_Description** - Free text part description
- **Component_Quantity** - Numeric quantity required
- **Component_Grade** - Material specification dropdown
- **Enquiry_Date** - Date picker for enquiry receipt
- **Notes** - Free text field for additional information

##### **Event Handlers**

**UserForm_Activate() - Form Initialization**
```vba
Private Sub UserForm_Activate()
    ' Load customer list from Customers/ directory
    ' Populate component codes from price list
    ' Load material grades from Component_Grades.xls
    ' Set default values and prepare form for input
End Sub
```

**SaveQ_Click() - Save Enquiry**
```vba
Private Sub SaveQ_Click()
    ' 1. Generate enquiry number
    Dim enquiryNumber As Long
    enquiryNumber = Calc_Numbers.Confirm_Next_Number("E")
    
    ' 2. Create enquiry file from template
    ' 3. Populate file with form data
    ' 4. Save to Enquiries/ directory
    ' 5. Update search database
    
    Call SaveSearchCode.SaveRowIntoSearch(Me)
End Sub
```

**Customer_Change() - Customer Selection Handler**
```vba
Private Sub Customer_Change()
    ' Populate contact person dropdown based on selected customer
    ' Load customer-specific pricing or preferences
End Sub
```

**AddMore_Click() - Continue Adding Enquiries**
```vba
Private Sub AddMore_Click()
    ' Save current enquiry
    ' Clear form for next enquiry
    ' Keep user in enquiry entry mode
End Sub
```

**AddNewClient_Click() - New Customer Creation**
```vba
Private Sub AddNewClient_Click()
    ' Open customer creation interface
    ' Create new customer record in Customers/ directory
    ' Refresh customer dropdown
End Sub
```

### **Data Validation and Business Rules**

#### **Required Field Validation**
```vba
Private Function ValidateEnquiryForm() As Boolean
    ' Customer must be selected
    If Customer.Value = "" Then
        MsgBox "Please select a customer"
        Customer.SetFocus
        ValidateEnquiryForm = False
        Exit Function
    End If
    
    ' Component description required
    If Component_Description.Value = "" Then
        MsgBox "Please enter component description"
        Component_Description.SetFocus
        ValidateEnquiryForm = False
        Exit Function
    End If
    
    ' Quantity must be positive number
    If Not IsNumeric(Component_Quantity.Value) Or Component_Quantity.Value <= 0 Then
        MsgBox "Please enter valid quantity"
        Component_Quantity.SetFocus
        ValidateEnquiryForm = False
        Exit Function
    End If
    
    ValidateEnquiryForm = True
End Function
```

---

## 🗃️ **Search Database Integration**

### **SaveSearchCode.bas** - Search Database Management

#### **Primary Function**

##### **`SaveRowIntoSearch(frm As Object)` - Update Master Search Index**

```vba
Sub SaveRowIntoSearch(frm As Object)
    ' Open Search.xls database
    Call Open_Book.OpenBook(Main.Main_MasterPath.Value & "Search.xls", False)
    
    ' Find or create row for this enquiry
    ' Map form controls to search database columns
    ' Update search record with enquiry data
    
    ' Column mapping:
    ' A: File_Name
    ' B: Enquiry_Number
    ' C: Customer
    ' D: Component_Description
    ' E: Component_Quantity
    ' F: Enquiry_Date
    ' G: System_Status
    
    ' Sort search database by date
    ' Save and close search database
End Sub
```

**Purpose**: Updates the central Search.xls database with enquiry information
**Parameters**: `frm As Object` - Reference to enquiry form containing data
**Dependencies**: 
- Search.xls must exist and be accessible
- Open_Book.bas for file operations
- Form controls must match expected names

#### **Search Database Schema**

**Search.xls Structure**:
| Column | Field | Purpose | Data Type |
|--------|-------|---------|----------|
| A | File_Name | Unique file identifier | String |
| B | Enquiry_Number | E-prefix number | String |
| C | Customer | Customer name | String |
| D | Component_Description | Part description | String |
| E | Component_Quantity | Required quantity | Number |
| F | Enquiry_Date | Date of enquiry | Date |
| G | System_Status | Workflow status | String |
| H | ContactPerson | Customer contact | String |
| I | Component_Code | Part number | String |
| J | Component_Grade | Material spec | String |

**Status Values for Enquiries**:
- "New Enquiry" - Just created
- "To Quote" - Ready for quote generation
- "Quoted" - Quote has been created

---

## 📂 **File System Integration**

### **Template and File Management**

#### **Template Usage**
```vba
' Enquiry creation process
1. Copy _Enq.xls template from Templates/ directory
2. Rename to E####.xls format (e.g., E1051.xls)
3. Populate "Admin" sheet with form data
4. Save to Enquiries/ directory
5. Update search database
```

#### **File Structure**

**Template File**: `Templates/_Enq.xls`
- **Admin Sheet**: Contains metadata fields
- **Job Card Sheet**: Placeholder for future job card
- **Additional Sheets**: Customer-specific or component-specific data

**Enquiry File**: `Enquiries/E####.xls`
- **Same structure as template**
- **Admin Sheet populated** with enquiry data
- **File_Name field**: Set to "E####"
- **System_Status**: Set to "To Quote"

#### **Admin Sheet Field Mapping**
```vba
' Form control -> Excel cell mapping
File_Name = "E" & Format(enquiryNumber, "0000")
Enquiry_Number = "E" & Format(enquiryNumber, "0000")
Customer = FEnquiry.Customer.Value
Component_Description = FEnquiry.Component_Description.Value
Component_Quantity = FEnquiry.Component_Quantity.Value
Component_Code = FEnquiry.Component_Code.Value
Component_Grade = FEnquiry.Component_Grade.Value
Enquiry_Date = FEnquiry.Enquiry_Date.Caption
ContactPerson = FEnquiry.ContactPerson.Value
System_Status = "To Quote"
```

---

## 🔗 **Dependencies and Integration**

### **Upstream Dependencies**

#### **Core Infrastructure**
- **a_Main.ShowMenu()** - System initialization
- **Main.Main_MasterPath.Value** - Base directory path
- **Open_Book.OpenBook()** - File operations
- **CheckDir()** - Directory validation

#### **Number Generation**
- **Calc_Numbers.Calc_Next_Number("E")** - Preview next enquiry number
- **Calc_Numbers.Confirm_Next_Number("E")** - Reserve enquiry number

#### **Reference Data**
- **Customers/ Directory** - Customer database files
- **price list.xls** - Component codes and pricing
- **Component_Grades.xls** - Material specifications
- **Templates/_Enq.xls** - Enquiry template file

### **Downstream Integration**

#### **Quote Management**
- Enquiry files in Enquiries/ directory are source for quote creation
- FQuote.frm reads enquiry data and creates quotes
- Files move from Enquiries/ to Quotes/ directory

#### **Search System**
- All enquiries automatically indexed in Search.xls
- Search functionality can locate enquiries by any field
- Status tracking shows enquiry workflow progress

#### **Main Interface**
- Main.frm displays enquiry counts and listings
- Users can browse and select enquiries for processing
- Status indicators show enquiry workflow state

---

## 🔄 **Complete Enquiry Workflow**

### **Step-by-Step Process**

#### **1. User Initiates Enquiry Creation**
```vba
' From Main.frm
Private Sub Add_Enquiry_Click()
    FEnquiry.Show
End Sub
```

#### **2. Form Initialization**
```vba
' FEnquiry.frm loads reference data
- Customer list from Customers/ directory
- Component codes from price list.xls
- Material grades from Component_Grades.xls
- Calculate preview enquiry number
```

#### **3. User Data Entry**
```vba
' User fills form fields:
- Select customer (triggers contact person lookup)
- Enter component description
- Select component code (optional)
- Select material grade
- Enter quantity
- Add notes
```

#### **4. Form Validation**
```vba
' System validates:
- Customer selected
- Component description entered
- Quantity is positive number
- Date is valid
```

#### **5. Number Generation and File Creation**
```vba
' On SaveQ_Click():
1. Confirm enquiry number (reserve it)
2. Copy template file (_Enq.xls)
3. Rename to E####.xls
4. Populate Admin sheet with form data
5. Save to Enquiries/ directory
```

#### **6. Search Database Update**
```vba
' Update master search index:
1. Open Search.xls
2. Find or create row for enquiry
3. Map form data to search columns
4. Set System_Status to "To Quote"
5. Sort by date and save
```

#### **7. User Notification and Cleanup**
```vba
' Provide user feedback:
- Display success message
- Clear form for next enquiry (if AddMore)
- Close form (if single enquiry)
- Refresh main interface counts
```

---

## ⚠️ **Error Handling and Edge Cases**

### **Common Error Scenarios**

#### **Missing Reference Data**
```vba
' Customer directory empty or missing
If Dir(Main.Main_MasterPath.Value & "Customers\") = "" Then
    MsgBox "Customer directory not found"
    ' Create directory or provide error guidance
End If

' Price list file missing
If Dir(Main.Main_MasterPath.Value & "price list.xls") = "" Then
    MsgBox "Price list not found - component codes unavailable"
    ' Continue with limited functionality
End If
```

#### **Number Generation Failures**
```vba
' Template directory issues
If Calc_Numbers.Calc_Next_Number("E") = 0 Then
    MsgBox "Cannot generate enquiry number - check Templates directory"
    Exit Sub
End If

' Number confirmation mismatch
Dim preview As Long, confirmed As Long
preview = Calc_Numbers.Calc_Next_Number("E")
confirmed = Calc_Numbers.Confirm_Next_Number("E")
If preview <> confirmed Then
    MsgBox "Number generation error - contact administrator"
End If
```

#### **File Operation Errors**
```vba
' Template file access
On Error GoTo TemplateError
Call FileCopy(templatePath, enquiryPath)
GoTo ContinueProcess

TemplateError:
    MsgBox "Cannot access enquiry template: " & templatePath
    Exit Sub

ContinueProcess:
' Continue with file population
```

#### **Search Database Issues**
```vba
' Search.xls locked or missing
On Error GoTo SearchError
Call Open_Book.OpenBook(searchPath, False)
GoTo UpdateSearch

SearchError:
    MsgBox "Cannot update search database - enquiry saved but not indexed"
    ' Continue without search update
    GoTo CompleteProcess

UpdateSearch:
' Proceed with search database update
```

---

## 🎯 **Development Guidelines**

### **Form Customization Best Practices**

#### **Adding New Fields**
```vba
' 1. Add control to form in VBA IDE
' 2. Update validation function
Private Function ValidateEnquiryForm() As Boolean
    ' Add validation for new field
    If NewField.Value = "" Then
        MsgBox "New field is required"
        NewField.SetFocus
        ValidateEnquiryForm = False
        Exit Function
    End If
End Function

' 3. Update file mapping
Private Sub PopulateEnquiryFile()
    ' Add new field to Admin sheet
    ws.Range("NewFieldCell").Value = FEnquiry.NewField.Value
End Sub

' 4. Update search database mapping
Private Sub UpdateSearchDatabase()
    ' Add new field to search record
    searchWS.Range("NewColumn" & rowNum).Value = FEnquiry.NewField.Value
End Sub
```

#### **Reference Data Management**
```vba
' Loading dropdown data
Private Sub LoadCustomerList()
    Dim customerPath As String
    Dim fileName As String
    
    customerPath = Main.Main_MasterPath.Value & "Customers\"
    fileName = Dir(customerPath & "*.xls")
    
    FEnquiry.Customer.Clear
    Do While fileName <> ""
        ' Extract customer name from filename
        FEnquiry.Customer.AddItem Left(fileName, Len(fileName) - 4)
        fileName = Dir
    Loop
End Sub
```

### **Integration Testing**

#### **End-to-End Testing**
```vba
Sub TestEnquiryWorkflow()
    ' 1. Test form initialization
    FEnquiry.Show
    DoEvents
    
    ' 2. Simulate user input
    FEnquiry.Customer.Value = "Test Customer"
    FEnquiry.Component_Description.Value = "Test Component"
    FEnquiry.Component_Quantity.Value = "10"
    
    ' 3. Test validation
    If ValidateEnquiryForm() Then
        MsgBox "Validation passed"
    Else
        MsgBox "Validation failed"
    End If
    
    ' 4. Test save operation
    ' (Use test data directory)
End Sub
```

---

## 🔍 **Next Steps**

After understanding Enquiry Management:

1. **Study [Quote Management](SUBSYSTEM_04_QUOTE_MANAGEMENT.md)** - See how enquiries become quotes
2. **Review [Search Database](SUBSYSTEM_08_SEARCH_DATA.md)** - Understand central indexing system
3. **Examine [Interface Navigation](SUBSYSTEM_06_INTERFACE_NAVIGATION.md)** - See how enquiries are displayed and selected
4. **Practice Form Customization** - Add fields or modify validation
5. **Test Integration** - Follow complete workflow from enquiry to quote

**Ready for quote management? Continue to [Quote Management Subsystem](SUBSYSTEM_04_QUOTE_MANAGEMENT.md)**