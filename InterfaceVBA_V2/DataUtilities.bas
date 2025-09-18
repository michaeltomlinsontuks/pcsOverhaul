Attribute VB_Name = "DataUtilities"
Option Explicit

Public Function GetValue(ByVal FilePath As String, ByVal SheetName As String, ByVal CellAddress As String) As Variant
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim CellValue As Variant

    On Error GoTo Error_Handler

    If Not FileManager.FileExists(FilePath) Then
        GetValue = ""
        Exit Function
    End If

    Set wb = FileManager.SafeOpenWorkbook(FilePath)
    If wb Is Nothing Then
        GetValue = ""
        Exit Function
    End If

    Set ws = wb.Worksheets(SheetName)
    CellValue = ws.Range(CellAddress).Value

    FileManager.SafeCloseWorkbook wb, False

    GetValue = CellValue
    Exit Function

Error_Handler:
    If Not wb Is Nothing Then FileManager.SafeCloseWorkbook wb, False
    ErrorHandler.HandleStandardErrors Err.Number, "GetValue", "DataUtilities"
    GetValue = ""
End Function

Public Function GetValueFromClosedWorkbook(ByVal FilePath As String, ByVal SheetName As String, ByVal CellAddress As String) As Variant
    Dim Formula As String
    Dim TempCell As Range

    On Error GoTo Error_Handler

    Set TempCell = ThisWorkbook.Worksheets(1).Cells(1, 1)

    Formula = "='" & FilePath & "'![" & SheetName & "]!" & CellAddress

    TempCell.Formula = Formula
    GetValueFromClosedWorkbook = TempCell.Value
    TempCell.Clear

    Exit Function

Error_Handler:
    If Not TempCell Is Nothing Then TempCell.Clear
    ErrorHandler.HandleStandardErrors Err.Number, "GetValueFromClosedWorkbook", "DataUtilities"
    GetValueFromClosedWorkbook = ""
End Function

Public Function SetValue(ByVal FilePath As String, ByVal SheetName As String, ByVal CellAddress As String, ByVal Value As Variant) As Boolean
    Dim wb As Workbook
    Dim ws As Worksheet

    On Error GoTo Error_Handler

    Set wb = FileManager.SafeOpenWorkbook(FilePath)
    If wb Is Nothing Then
        SetValue = False
        Exit Function
    End If

    Set ws = wb.Worksheets(SheetName)
    ws.Range(CellAddress).Value = Value

    wb.Save
    FileManager.SafeCloseWorkbook wb

    SetValue = True
    Exit Function

Error_Handler:
    If Not wb Is Nothing Then FileManager.SafeCloseWorkbook wb, False
    ErrorHandler.HandleStandardErrors Err.Number, "SetValue", "DataUtilities"
    SetValue = False
End Function

Public Function GetRowData(ByVal FilePath As String, ByVal SheetName As String, ByVal RowNumber As Long) As Variant
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim LastCol As Long
    Dim RowData As Variant

    On Error GoTo Error_Handler

    Set wb = FileManager.SafeOpenWorkbook(FilePath)
    If wb Is Nothing Then
        GetRowData = Array()
        Exit Function
    End If

    Set ws = wb.Worksheets(SheetName)
    LastCol = ws.Cells(RowNumber, ws.Columns.Count).End(xlToLeft).Column

    If LastCol > 0 Then
        RowData = ws.Range(ws.Cells(RowNumber, 1), ws.Cells(RowNumber, LastCol)).Value
    Else
        RowData = Array()
    End If

    FileManager.SafeCloseWorkbook wb, False

    GetRowData = RowData
    Exit Function

Error_Handler:
    If Not wb Is Nothing Then FileManager.SafeCloseWorkbook wb, False
    ErrorHandler.HandleStandardErrors Err.Number, "GetRowData", "DataUtilities"
    GetRowData = Array()
End Function

Public Function GetColumnData(ByVal FilePath As String, ByVal SheetName As String, ByVal ColumnNumber As Long) As Variant
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim ColumnData As Variant

    On Error GoTo Error_Handler

    Set wb = FileManager.SafeOpenWorkbook(FilePath)
    If wb Is Nothing Then
        GetColumnData = Array()
        Exit Function
    End If

    Set ws = wb.Worksheets(SheetName)
    LastRow = ws.Cells(ws.Rows.Count, ColumnNumber).End(xlUp).Row

    If LastRow > 0 Then
        ColumnData = ws.Range(ws.Cells(1, ColumnNumber), ws.Cells(LastRow, ColumnNumber)).Value
    Else
        ColumnData = Array()
    End If

    FileManager.SafeCloseWorkbook wb, False

    GetColumnData = ColumnData
    Exit Function

Error_Handler:
    If Not wb Is Nothing Then FileManager.SafeCloseWorkbook wb, False
    ErrorHandler.HandleStandardErrors Err.Number, "GetColumnData", "DataUtilities"
    GetColumnData = Array()
End Function

Public Function GetRangeData(ByVal FilePath As String, ByVal SheetName As String, ByVal RangeAddress As String) As Variant
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim RangeData As Variant

    On Error GoTo Error_Handler

    Set wb = FileManager.SafeOpenWorkbook(FilePath)
    If wb Is Nothing Then
        GetRangeData = Array()
        Exit Function
    End If

    Set ws = wb.Worksheets(SheetName)
    RangeData = ws.Range(RangeAddress).Value

    FileManager.SafeCloseWorkbook wb, False

    GetRangeData = RangeData
    Exit Function

Error_Handler:
    If Not wb Is Nothing Then FileManager.SafeCloseWorkbook wb, False
    ErrorHandler.HandleStandardErrors Err.Number, "GetRangeData", "DataUtilities"
    GetRangeData = Array()
End Function

Public Function FindValue(ByVal FilePath As String, ByVal SheetName As String, ByVal SearchValue As Variant, Optional ByVal SearchColumn As Long = 1) As Long
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim FoundCell As Range

    On Error GoTo Error_Handler

    Set wb = FileManager.SafeOpenWorkbook(FilePath)
    If wb Is Nothing Then
        FindValue = 0
        Exit Function
    End If

    Set ws = wb.Worksheets(SheetName)
    Set FoundCell = ws.Columns(SearchColumn).Find(SearchValue, LookIn:=xlValues, LookAt:=xlWhole)

    If Not FoundCell Is Nothing Then
        FindValue = FoundCell.Row
    Else
        FindValue = 0
    End If

    FileManager.SafeCloseWorkbook wb, False

    Exit Function

Error_Handler:
    If Not wb Is Nothing Then FileManager.SafeCloseWorkbook wb, False
    ErrorHandler.HandleStandardErrors Err.Number, "FindValue", "DataUtilities"
    FindValue = 0
End Function

Public Function CleanFileName(ByVal FileName As String) As String
    Dim InvalidChars As String
    Dim i As Integer

    InvalidChars = "\/:*?""<>|"

    CleanFileName = FileName

    For i = 1 To Len(InvalidChars)
        CleanFileName = Replace(CleanFileName, Mid(InvalidChars, i, 1), "_")
    Next i

    CleanFileName = Trim(CleanFileName)
End Function

Public Function FormatCurrency(ByVal Amount As Currency) As String
    FormatCurrency = Format(Amount, "$#,##0.00")
End Function

Public Function FormatDate(ByVal DateValue As Date) As String
    FormatDate = Format(DateValue, "dd/mm/yyyy")
End Function