VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FWIP
   Caption         =   "WIP Reports"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4470
   OleObjectBlob   =   "fwip.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fwip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Type Jobs
    Dat As Date
    Cust As String
    Job As String
    JobD As Double
    Qty As String
    Cod As String
    Desc As String
    Remarks As String
    DDat As String

    OperatorN(1 To 15) As String
    OperatorType(1 To 15) As String
End Type

Private Sub Go_Click()
    Dim TempSheet As String
    Dim Job(1 To 5000) As Jobs
    Dim i As Integer
    Dim col As Integer
    Dim OPP As String
    Dim j As Integer
    Dim k As Integer
    Dim x As Integer
    Dim sh As Worksheet
    Dim Sortcol As String, Sortcol1 As String, Sortcol2 As String

    On Error GoTo Error_Handler

    fwip.Label1.Caption = "Please Wait"
    Application.DisplayAlerts = False

    ' Load WIP data from individual job files in WIP directory
    Dim WIPDirPath As String
    WIPDirPath = DataManager.GetRootPath & "\WIP\"

    ' Check if WIP directory exists
    If Not DataManager.DirExists(WIPDirPath) Then
        MsgBox "WIP directory not found at: " & WIPDirPath & vbCrLf & vbCrLf & _
               "Please create the WIP directory and add some job files to generate reports.", vbCritical, "WIP Directory Not Found"
        Exit Sub
    End If

    ' Get list of Excel files in WIP directory
    Dim WIPFiles As Variant
    WIPFiles = DataManager.GetFileList("WIP")

    If IsArray(WIPFiles) And UBound(WIPFiles) >= 0 Then
        ' Process each WIP file
        i = LoadWIPDataFromFiles(WIPFiles, Job, WIPDirPath)

        If i = 0 Then
            MsgBox "No valid WIP data found in: " & WIPDirPath & vbCrLf & vbCrLf & _
                   "Please ensure there are job Excel files in the WIP directory.", vbInformation, "No WIP Data"
            Exit Sub
        End If
    Else
        MsgBox "No Excel files found in WIP directory: " & WIPDirPath & vbCrLf & vbCrLf & _
               "Please add job files (.xls) to the WIP directory.", vbInformation, "No WIP Files"
        Exit Sub
    End If

    fwip.Hide

    ' Operation Reports
    If ROperation.Value = True Then
        Workbooks.Add

        OPP = ""
        If MsgBox("Specific Operation?", vbYesNo) = vbYes Then
            OPP = InputBox("Which Operation")
        End If

        For j = 1 To i
            With Job(j)
                For k = 1 To 15
                    If OPP <> "" Then
                        If Trim(UCase(.OperatorType(k))) <> Trim(UCase(OPP)) Then GoTo SkipOPP
                    End If

                    If .OperatorType(k) <> "" Then
                        TempSheet = "OPERATION - " & .OperatorType(k)
                        On Error GoTo AddSheet
                        Sheets(CoreFramework.RemoveInvalidCharacters(Trim(TempSheet))).Select
                        On Error GoTo Error_Handler

                        ActiveCell.FormulaR1C1 = .Dat
                        ActiveCell.Offset(0, 1).FormulaR1C1 = .Cust
                        ActiveCell.Offset(0, 2).FormulaR1C1 = .Job
                        ActiveCell.Offset(0, 3).FormulaR1C1 = .JobD
                        ActiveCell.Offset(0, 4).FormulaR1C1 = .Qty
                        ActiveCell.Offset(0, 5).FormulaR1C1 = .Cod
                        ActiveCell.Offset(0, 6).FormulaR1C1 = .Desc
                        ActiveCell.Offset(0, 7).FormulaR1C1 = .Remarks
                        ActiveCell.Offset(0, 8).FormulaR1C1 = .DDat

                        If k > 1 Then
                            If .OperatorType(k - 1) = "" Then
                                ActiveCell.Offset(0, 9).FormulaR1C1 = "*"
                                Selection.EntireRow.Font.Bold = True
                            End If
                        End If
                        ActiveCell.Offset(1, 0).Select
                    End If
                    TempSheet = ""
SkipOPP:
                Next k
            End With
        Next j

        ' Format all operation sheets
        For Each sh In Sheets
            sh.Select
            Cells.EntireColumn.AutoFit
            Range("A1:i5000").Select
            Selection.Sort Key1:=Range("H2"), Order1:=xlAscending, Key2:=Range("G2") _
                , Order2:=xlAscending, Header:=xlYes, OrderCustom:=1, MatchCase:= _
                False, Orientation:=xlTopToBottom

            With ActiveSheet.PageSetup
                .CenterHeader = ActiveSheet.Name
                .RightHeader = "&D &T"
            End With

            FormatWorksheetBorders
            Range("A1").Select
            Range("a:a").NumberFormat = "DD MMM YYYY"
            Range("i:i").NumberFormat = "DD MMM YYYY"
        Next sh

        DeleteSheet ("sheet1")
        DeleteSheet ("sheet2")
        DeleteSheet ("sheet3")

        Application.DisplayAlerts = False
        ActiveWorkbook.SaveAs (DataManager.GetRootPath & "\TEMPLATES\Operation.xls")
    End If

    ' Operator Reports
    If ROperator.Value = True Then
        Workbooks.Add

        For j = 1 To i
            With Job(j)
                For k = 1 To 15
                    If Trim(.OperatorN(k)) <> "" Then
                        TempSheet = CoreFramework.RemoveInvalidCharacters("OPERATOR - " & Trim(.OperatorN(k)))
                        On Error GoTo AddSheet
                        Sheets(TempSheet).Select
                        On Error GoTo Error_Handler

                        ActiveCell.FormulaR1C1 = .Dat
                        ActiveCell.Offset(0, 1).FormulaR1C1 = .Cust
                        ActiveCell.Offset(0, 2).FormulaR1C1 = .Job
                        ActiveCell.Offset(0, 3).FormulaR1C1 = .JobD
                        ActiveCell.Offset(0, 4).FormulaR1C1 = .Qty
                        ActiveCell.Offset(0, 5).FormulaR1C1 = .Cod
                        ActiveCell.Offset(0, 6).FormulaR1C1 = .Desc
                        ActiveCell.Offset(0, 7).FormulaR1C1 = .Remarks
                        ActiveCell.Offset(0, 8).FormulaR1C1 = .DDat

                        If k > 1 Then
                            If .OperatorN(k - 1) = "" Then
                                ActiveCell.Offset(0, 9).FormulaR1C1 = "*"
                                Selection.EntireRow.Font.Bold = True
                            End If
                        End If

                        ActiveCell.Offset(1, 0).Select
                    End If
                    TempSheet = ""
                Next k
            End With
        Next j

        ' Format all operator sheets
        For Each sh In Sheets
            sh.Select
            Cells.EntireColumn.AutoFit
            Range("A1:i5000").Select
            Selection.Sort Key1:=Range("H2"), Order1:=xlAscending, Key2:=Range("G2") _
                , Order2:=xlAscending, Header:=xlYes, OrderCustom:=1, MatchCase:= _
                False, Orientation:=xlTopToBottom

            With ActiveSheet.PageSetup
                .CenterHeader = ActiveSheet.Name
                .RightHeader = "&D &T"
            End With

            FormatWorksheetBorders
            Range("a:a").NumberFormat = "DD MMM YYYY"
            Range("i:i").NumberFormat = "DD MMM YYYY"
            Range("A1").Select
        Next sh

        DeleteSheet ("sheet1")
        DeleteSheet ("sheet2")
        DeleteSheet ("sheet3")

        Application.DisplayAlerts = False
        ActiveWorkbook.SaveAs (DataManager.GetRootPath & "\TEMPLATES\Operator.xls")
    End If

    ' Generate consolidated reports from loaded job data
    ' (All the report types now use the job data loaded from individual files)

    Application.DisplayAlerts = True
    Unload fwip
    Exit Sub

AddSheet:
    Sheets.Add
    ActiveSheet.Name = CoreFramework.RemoveInvalidCharacters(TempSheet)
    ActiveSheet.PageSetup.CenterHeader = TempSheet
    ActiveCell.FormulaR1C1 = "DATE"
    ActiveCell.Offset(0, 1).FormulaR1C1 = "CUSTOMER"
    ActiveCell.Offset(0, 2).FormulaR1C1 = "JOB"
    ActiveCell.Offset(0, 3).FormulaR1C1 = "JOB"
    ActiveCell.Offset(0, 4).FormulaR1C1 = "QTY"
    ActiveCell.Offset(0, 5).FormulaR1C1 = "COMPONENT CODE"
    ActiveCell.Offset(0, 6).FormulaR1C1 = "COMPONENT DESCRIPTION"
    ActiveCell.Offset(0, 7).FormulaR1C1 = "REMARKS"
    ActiveCell.Offset(0, 8).FormulaR1C1 = "DUE DATE"
    Columns("h:h").NumberFormat = "dd mmm"
    Columns("A:A").NumberFormat = "dd mmm"
    Selection.EntireRow.Font.Bold = True

    Columns("A:A").ColumnWidth = 10
    Columns("b:b").ColumnWidth = 18
    Columns("c:c").ColumnWidth = 10
    Columns("e:e").ColumnWidth = 6
    Columns("g:g").ColumnWidth = 30
    Columns("h:h").ColumnWidth = 20
    Columns("i:i").ColumnWidth = 10
    Cells.RowHeight = 30

    ActiveCell.Offset(1, 0).Select
    Resume

Error_Handler:
    Application.DisplayAlerts = True
    If Not WIPWB Is Nothing Then DataManager.SafeCloseWorkbook WIPWB, False
    CoreFramework.HandleStandardErrors Err.Number, "Go_Click", "fwip"
End Sub

Private Function ParseJobNumberForSorting(jobString As String) As Double
    Dim parts() As String
    Dim mainPart As Double
    Dim subPart As Double

    If InStr(jobString, "-") > 0 Then
        parts = Split(jobString, "-")
        mainPart = Val(parts(0))
        If UBound(parts) > 0 Then
            subPart = Val(parts(1))
            ParseJobNumberForSorting = mainPart + (subPart / 1000000)
        Else
            ParseJobNumberForSorting = mainPart
        End If
    Else
        ParseJobNumberForSorting = Val(jobString)
    End If
End Function

Private Function ShowOfficeCols()
    Range("A1").Select
    Do
        Selection.EntireColumn.Hidden = True

        Select Case ActiveCell.Value
            Case "Job_StartDate"
                Selection.EntireColumn.Hidden = False
            Case "Job_Urgency"
                Selection.EntireColumn.Hidden = False
            Case "CUSTOMER"
                Selection.EntireColumn.Hidden = False
            Case "Job_Number"
                Selection.EntireColumn.Hidden = False
            Case "Component_Quantity"
                Selection.EntireColumn.Hidden = False
            Case "Component_Code"
                Selection.EntireColumn.Hidden = False
            Case "Component_Description"
                Selection.EntireColumn.Hidden = False
            Case "Component_Comments"
                Selection.EntireColumn.Hidden = False
            Case "CustomerDelivery_Date"
                Selection.EntireColumn.Hidden = False
            Case "CustomerOrderNumber"
                Selection.EntireColumn.Hidden = False
            Case "Component_Price"
                Selection.EntireColumn.Hidden = False
            Case "Component_DrawingNumber_SampleNumber"
                Selection.EntireColumn.Hidden = False
        End Select
        ActiveCell.Offset(0, 1).Select

    Loop Until ActiveCell.Value = ""
End Function

Private Function ShowWorkshopCols()
    Range("A1").Select
    Do
        Selection.EntireColumn.Hidden = True

        Select Case ActiveCell.Value
            Case "Job_StartDate"
                Selection.EntireColumn.Hidden = False
            Case "Job_Urgency"
                Selection.EntireColumn.Hidden = False
            Case "CUSTOMER"
                Selection.EntireColumn.Hidden = False
            Case "Job_Number"
                Selection.EntireColumn.Hidden = False
            Case "Job_WorkshopDueDate"
                Selection.EntireColumn.Hidden = False
            Case "Component_Quantity"
                Selection.EntireColumn.Hidden = False
            Case "Component_Code"
                Selection.EntireColumn.Hidden = False
            Case "Component_Description"
                Selection.EntireColumn.Hidden = False
            Case "Component_Comments"
                Selection.EntireColumn.Hidden = False
            Case " "
                Selection.EntireColumn.Hidden = False
            Case "Component_DrawingNumber_SampleNumber"
                Selection.EntireColumn.Hidden = False
            Case "Operation01_Type"
                Selection.EntireColumn.Hidden = False
            Case "Operation01_Operator"
                Selection.EntireColumn.Hidden = False
            Case "Operation02_Type"
                Selection.EntireColumn.Hidden = False
            Case "Operation02_Operator"
                Selection.EntireColumn.Hidden = False
            Case "Operation03_Type"
                Selection.EntireColumn.Hidden = False
            Case "Operation03_Operator"
                Selection.EntireColumn.Hidden = False
            Case "Operation04_Type"
                Selection.EntireColumn.Hidden = False
            Case "Operation04_Operator"
                Selection.EntireColumn.Hidden = False
            Case "Operation05_Type"
                Selection.EntireColumn.Hidden = False
            Case "Operation05_Operator"
                Selection.EntireColumn.Hidden = False
            Case "Operation06_Type"
                Selection.EntireColumn.Hidden = False
            Case "Operation06_Operator"
                Selection.EntireColumn.Hidden = False
            Case "Operation07_Type"
                Selection.EntireColumn.Hidden = False
            Case "Operation07_Operator"
                Selection.EntireColumn.Hidden = False
            Case "Operation08_Type"
                Selection.EntireColumn.Hidden = False
            Case "Operation08_Operator"
                Selection.EntireColumn.Hidden = False
            Case "Operation09_Type"
                Selection.EntireColumn.Hidden = False
            Case "Operation09_Operator"
                Selection.EntireColumn.Hidden = False
            Case "Operation10_Type"
                Selection.EntireColumn.Hidden = False
            Case "Operation10_Operator"
                Selection.EntireColumn.Hidden = False
            Case "Operation11_Type"
                Selection.EntireColumn.Hidden = False
            Case "Operation11_Operator"
                Selection.EntireColumn.Hidden = False
            Case "Operation12_Type"
                Selection.EntireColumn.Hidden = False
            Case "Operation12_Operator"
                Selection.EntireColumn.Hidden = False
            Case "Operation13_Type"
                Selection.EntireColumn.Hidden = False
            Case "Operation13_Operator"
                Selection.EntireColumn.Hidden = False
            Case "Operation14_Type"
                Selection.EntireColumn.Hidden = False
            Case "Operation14_Operator"
                Selection.EntireColumn.Hidden = False
            Case "Operation15_Type"
                Selection.EntireColumn.Hidden = False
            Case "Operation15_Operator"
                Selection.EntireColumn.Hidden = False
        End Select
        ActiveCell.Offset(0, 1).Select

    Loop Until ActiveCell.Value = ""
End Function

Private Sub FormatWorksheetBorders()
    Cells.Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
End Sub

Private Sub DeleteSheet(SheetName As String)
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets(SheetName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub

Private Function LoadWIPDataFromFiles(FileList As Variant, ByRef JobArray() As Jobs, WIPDirPath As String) As Integer
    Dim JobCount As Integer
    Dim CurrentWB As Workbook
    Dim CurrentWS As Worksheet
    Dim FileName As String
    Dim FilePath As String
    Dim i As Integer
    Dim OpNum As Integer

    On Error GoTo LoadError

    JobCount = 0

    ' Process each file in the WIP directory
    For i = LBound(FileList) To UBound(FileList)
        FileName = FileList(i)
        FilePath = WIPDirPath & FileName

        ' Try to open the job file
        Set CurrentWB = DataManager.SafeOpenWorkbook(FilePath)
        If Not CurrentWB Is Nothing Then
            ' Try to find the ADMIN sheet (standard job file structure)
            Set CurrentWS = Nothing
            On Error Resume Next
            Set CurrentWS = CurrentWB.Worksheets("ADMIN")
            On Error GoTo LoadError

            If Not CurrentWS Is Nothing Then
                ' Extract job data from ADMIN sheet
                JobCount = JobCount + 1
                If JobCount <= 5000 Then ' Prevent array overflow
                    With JobArray(JobCount)
                        ' Extract basic job information
                        .Dat = GetCellValueByName(CurrentWS, "Date_Created")
                        .Cust = GetCellValueByName(CurrentWS, "Customer_Name")
                        .Job = GetCellValueByName(CurrentWS, "Job_Description")
                        .JobD = ParseJobNumberForSorting(GetCellValueByName(CurrentWS, "Job_Number"))
                        .Qty = GetCellValueByName(CurrentWS, "Component_Quantity")
                        .Cod = GetCellValueByName(CurrentWS, "Component_Code")
                        .Desc = GetCellValueByName(CurrentWS, "Component_Description")
                        .Remarks = GetCellValueByName(CurrentWS, "Component_Comments")
                        .DDat = GetCellValueByName(CurrentWS, "Customer_DueDate")

                        ' Extract operation information
                        For OpNum = 1 To 15
                            .OperatorType(OpNum) = GetCellValueByName(CurrentWS, "Operation" & Format(OpNum, "00") & "_Type")
                            .OperatorN(OpNum) = GetCellValueByName(CurrentWS, "Operation" & Format(OpNum, "00") & "_Operator")
                        Next OpNum
                    End With
                End If
            End If

            ' Close the workbook
            DataManager.SafeCloseWorkbook CurrentWB, False
            Set CurrentWB = Nothing
        End If
    Next i

    LoadWIPDataFromFiles = JobCount
    Exit Function

LoadError:
    If Not CurrentWB Is Nothing Then
        DataManager.SafeCloseWorkbook CurrentWB, False
        Set CurrentWB = Nothing
    End If
    CoreFramework.HandleStandardErrors Err.Number, "LoadWIPDataFromFiles", "fwip"
    LoadWIPDataFromFiles = JobCount ' Return what we have so far
End Function

Private Function GetCellValueByName(WS As Worksheet, FieldName As String) As Variant
    ' Find field by name in column A and return value from column B
    Dim SearchRange As Range
    Dim FoundCell As Range

    On Error GoTo GetCellError

    Set SearchRange = WS.Range("A:A")
    Set FoundCell = SearchRange.Find(FieldName, LookIn:=xlValues, LookAt:=xlWhole)

    If Not FoundCell Is Nothing Then
        GetCellValueByName = WS.Cells(FoundCell.Row, 2).Value ' Column B
    Else
        GetCellValueByName = ""
    End If

    Exit Function

GetCellError:
    GetCellValueByName = ""
End Function

