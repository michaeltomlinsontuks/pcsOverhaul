Attribute VB_Name = "DataTypes"
Option Explicit

' Public type definitions for PCS Interface V2
Public Type SearchResult
    FilePath As String
    CustomerName As String
    ComponentCode As String
    ComponentDesc As String
    Status As String
    MatchScore As Integer
    FileType As String
    ModDate As Date
End Type

Public Type FilterState
    NewEnquiries As Boolean
    QuotesToSubmit As Boolean
    WIPToSequence As Boolean
    JobsInWIP As Boolean
    ShowArchived As Boolean
    DateRangeStart As Date
    DateRangeEnd As Date
End Type

Public Type FileInfo
    FullPath As String
    ModDate As Date
    Size As Long
    IsValid As Boolean
End Type