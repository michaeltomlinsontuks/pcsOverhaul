Attribute VB_Name = "DataTypes"
Option Explicit

Public Type EnquiryData
    EnquiryNumber As String
    CustomerName As String
    ContactPerson As String
    CompanyPhone As String
    CompanyFax As String
    Email As String
    ComponentDescription As String
    ComponentCode As String
    MaterialGrade As String
    Quantity As Long
    DateCreated As Date
    FilePath As String
    SearchKeywords As String
End Type

Public Type QuoteData
    QuoteNumber As String
    EnquiryNumber As String
    CustomerName As String
    ComponentDescription As String
    ComponentCode As String
    MaterialGrade As String
    Quantity As Long
    UnitPrice As Currency
    TotalPrice As Currency
    LeadTime As String
    ValidUntil As Date
    DateCreated As Date
    FilePath As String
    Status As String
End Type

Public Type JobData
    JobNumber As String
    QuoteNumber As String
    CustomerName As String
    ComponentDescription As String
    ComponentCode As String
    MaterialGrade As String
    Quantity As Long
    DueDate As Date
    WorkshopDueDate As Date
    CustomerDueDate As Date
    OrderValue As Currency
    DateCreated As Date
    FilePath As String
    Status As String
    AssignedOperator As String
    Operations As String
    Pictures As String
    Notes As String
End Type

Public Type ContractData
    ContractName As String
    CustomerName As String
    ComponentDescription As String
    StandardOperations As String
    LeadTime As String
    FilePath As String
    DateCreated As Date
    LastUsed As Date
End Type

Public Type SearchRecord
    RecordType As String
    RecordNumber As String
    CustomerName As String
    Description As String
    DateCreated As Date
    FilePath As String
    Keywords As String
End Type

Public Enum RecordType
    rtEnquiry = 1
    rtQuote = 2
    rtJob = 3
    rtContract = 4
End Enum

Public Enum JobStatus
    jsActive = 1
    jsOnHold = 2
    jsCompleted = 3
    jsCancelled = 4
End Enum