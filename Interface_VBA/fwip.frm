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

Private Function GetWIPFileName() As String
    Dim BasePath As String
    BasePath = Main.Main_MasterPath

    If Dir(BasePath & "WIP.xls") <> "" Then
        GetWIPFileName = "WIP.xls"
    ElseIf Dir(BasePath & "Wip.xls") <> "" Then
        GetWIPFileName = "Wip.xls"
    ElseIf Dir(BasePath & "wip.xls") <> "" Then
        GetWIPFileName = "wip.xls"
    Else
        GetWIPFileName = "WIP.xls"
    End If
End Function

Private Function GetWIPWorkbook() As Workbook
    Dim WIPFileName As String
    Dim wb As Workbook

    WIPFileName = GetWIPFileName()

    ' Try to find the workbook that's already open
    For Each wb In Workbooks
        If UCase(wb.Name) = UCase(WIPFileName) Then
            Set GetWIPWorkbook = wb
            Exit Function
        End If
    Next wb

    ' If not found, return Nothing
    Set GetWIPWorkbook = Nothing
End Function

Private Sub ActivateWIPWorkbook()
    Dim wb As Workbook
    Set wb = GetWIPWorkbook()

    If Not wb Is Nothing Then
        wb.Activate
    Else
        ' Try to open the file if it's not already open
        Workbooks.Open Main.Main_MasterPath & GetWIPFileName(), ReadOnly:=True
        Set wb = GetWIPWorkbook()
        If Not wb Is Nothing Then
            wb.Activate
        End If
    End If
End Sub

Private Sub Go_Click()
Dim TempSheet As String
Dim Job(1 To 5000) As Jobs

fwip.Label1.Caption = "Please Wait"

Application.DisplayAlerts = False

'OpenBook (main.Main_MasterPath & "Wip.xls")

Workbooks.Open Main.Main_MasterPath & GetWIPFileName(), ReadOnly:=True
Range("bb1").Select
Selection.End(xlToLeft).Select
col = ActiveCell.Column

Range("A1").Select
Selection.End(xlDown).Select

Range("A2", Range("A2").Offset(ActiveCell.Row, col - 1).Address).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Sort Key1:=Range("h3"), Order1:=xlAscending, Header:=xlYes, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
            
Range("A3").Select
fwip.Hide
Main.Hide
i = 0
If ActiveCell.FormulaR1C1 <> "" Then
    Do
        i = i + 1
        With Job(i)
            .Dat = ActiveCell.Offset(0, 0).Value
            .Cust = ActiveCell.Offset(0, 1).Value
            .Job = ActiveCell.Offset(0, 2).Value
            .JobD = ParseJobNumberForSorting(ActiveCell.Offset(0, 3).Value)
            .Qty = ActiveCell.Offset(0, 4).Value
            .Cod = ActiveCell.Offset(0, 5).Value
            .Desc = ActiveCell.Offset(0, 6).Value
            .Remarks = ActiveCell.Offset(0, 8).Value
            .DDat = ActiveCell.Offset(0, 12).Value
            x = 0
            For j = 1 To 30 Step 2
                x = x + 1
                .OperatorType(x) = ActiveCell.Offset(0, 14 + j).Value
            Next j
            x = 0
            For j = 1 To 30 Step 2
                x = x + 1
                .OperatorN(x) = ActiveCell.Offset(0, 15 + j).Value
            Next j
        End With
        ActiveCell.Offset(1, 0).Select
    Loop Until ActiveCell.FormulaR1C1 = ""
Else
    End
End If

fwip.Hide
Main.Hide
'Operation - Enhanced with CustomReports
If ROperation.Value = True Then
    ' Ask for specific operation filter (preserving original behavior)
    OPP = ""
    If MsgBox("Specific Operation?", vbYesNo) = vbYes Then
        OPP = InputBox("Which Operation")
    End If

    ' Generate enhanced operation report using CustomReports module
    CustomReports.GenerateWIPReport "Operation", OPP
End If

'Operator - Enhanced with CustomReports
If ROperator.Value = True Then
    ' Generate enhanced operator report using CustomReports module
    CustomReports.GenerateWIPReport "Operator"
End If

'End With

ActivateWIPWorkbook
If fwip.RDueDate.Value = True Then
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs (Main.Main_MasterPath & "TEMPLATES\Due Date.xls")
    Range("a1").Select
Else
    ActiveWorkbook.Close False
End If

If fwip.RWIP.Value = True Then
    Workbooks.Open Main.Main_MasterPath & GetWIPFileName(), ReadOnly:=True

    Range("A2", Range("A2").Offset(ActiveCell.Row, col - 1).Address).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Sort Key1:=Range("A3"), Order1:=xlAscending, Header:=xlYes, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom

    Range("A1").Select

    ' Generate enhanced WIP_NEW report using CustomReports module
    CustomReports.GenerateWIPReport "WIP"
End If

'New

If fwip.Job_DueDate.Value = True Then
    Workbooks.Open Main.Main_MasterPath & GetWIPFileName(), ReadOnly:=True
    
    ActivateWIPWorkbook
    Range("A1").Select
    Do
        ActiveCell.Offset(0, 1).Select
    Loop Until ActiveCell.Value = "CustomerDelivery_Date" Or ActiveCell.FormulaR1C1 = ""
    
    Sortcol = ActiveCell.Address
    
    Range("A3", Range("A3").Offset(0, col - 1).Address).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Sort Key1:=Range(Range(Sortcol).Offset(2, 0).Address), Order1:=xlAscending, Header:=xlNo, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
    
    ShowOfficeCols
        Range("b1").Select
    Application.DisplayAlerts = False
                    
    With ActiveSheet.PageSetup
    
        .CenterHeader = "OFFICE DUE DATE"
        .RightHeader = "&D &T"
        
    End With
    
    ActiveWorkbook.SaveAs (Main.Main_MasterPath & "TEMPLATES\CustomerDelivery_Date.xls")

   
End If

If fwip.Office_Customer.Value = True Then
    Workbooks.Open Main.Main_MasterPath & GetWIPFileName(), ReadOnly:=True
    
    ActivateWIPWorkbook
    Range("A1").Select
    Do
        ActiveCell.Offset(0, 1).Select
        If UCase(ActiveCell.Value) = UCase("Customer") Then Sortcol1 = ActiveCell.Address
        If UCase(ActiveCell.Value) = UCase("Job_Number") Then Sortcol2 = ActiveCell.Address
        'If ActiveCell.Value = "" Then Sortcol2 = ActiveCell.Address
        
    Loop Until ActiveCell.FormulaR1C1 = ""

    
    Range("A3", Range("A3").Offset(0, col - 1).Address).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Sort Key1:=Range(Range(Sortcol1).Offset(2, 0).Address), Order1:=xlAscending, Header:=xlNo, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom _
        , Key2:=Range(Range(Sortcol2).Offset(2, 0).Address), Order2:=xlAscending '_
        ', Key3:=Range(Range(Sortcol3).Offset(2, 0).Address), Order3:=xlAscending
    
    ShowOfficeCols
        Range("b1").Select
                            
    With ActiveSheet.PageSetup
    
        .CenterHeader = "OFFICE CUSTOMER"
        .RightHeader = "&D &T"
        
    End With
    
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs (Main.Main_MasterPath & "TEMPLATES\Office_Customer.xls")
End If

'CUSTOMER
If fwip.Workshop_Customer.Value = True Then
    Workbooks.Open Main.Main_MasterPath & GetWIPFileName(), ReadOnly:=True
    
    ActivateWIPWorkbook
    Range("A1").Select
    Do
        ActiveCell.Offset(0, 1).Select
         
        If UCase(ActiveCell.Value) = UCase("Customer") Then Sortcol1 = ActiveCell.Address
        If UCase(ActiveCell.Value) = UCase("Job_Number") Then Sortcol2 = ActiveCell.Address
        'If ActiveCell.Value = "" Then Sortcol2 = ActiveCell.Address
        
    Loop Until ActiveCell.FormulaR1C1 = ""
    
'    Sortcol = ActiveCell.Address
    
    Range("A3", Range("A3").Offset(0, col - 1).Address).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Sort Key1:=Range(Range(Sortcol1).Offset(2, 0).Address), Order1:=xlAscending, Header:=xlNo, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom _
        , Key2:=Range(Range(Sortcol2).Offset(2, 0).Address), Order2:=xlAscending '_
        ', Key3:=Range(Range(Sortcol3).Offset(2, 0).Address), Order3:=xlAscending
    
    ShowWorkshopCols
        Range("b1").Select
                            
    With ActiveSheet.PageSetup
    
        .CenterHeader = "WORKSHOP CUSTOMER"
        .RightHeader = "&D &T"
        
    End With

    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs (Main.Main_MasterPath & "TEMPLATES\Workshop_Customer.xls")
End If

'OFFICE JOB NUMBER
If fwip.Office_JobNumber.Value = True Then
    Workbooks.Open Main.Main_MasterPath & GetWIPFileName(), ReadOnly:=True
    
    ActivateWIPWorkbook
    Range("A1").Select
    Do
        ActiveCell.Offset(0, 1).Select
    Loop Until ActiveCell.Value = "Converted_JN" Or ActiveCell.FormulaR1C1 = ""
    
    Sortcol = ActiveCell.Address
    
    Range("A3", Range("A3").Offset(0, col - 1).Address).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Sort Key1:=Range(Range(Sortcol).Offset(2, 0).Address), Order1:=xlAscending, Header:=xlNo, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortTextAsNumbers
        
    
    ShowOfficeCols
            Range("b1").Select
                                    
    With ActiveSheet.PageSetup
    
        .CenterHeader = "OFFICE JOB NUMBER"
        .RightHeader = "&D &T"
        
    End With
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs (Main.Main_MasterPath & "TEMPLATES\Office_JobNumber.xls")

   
End If

'WORKSHOP JOB NUMBER
If fwip.Workshop_JobNumber.Value = True Then
    Workbooks.Open Main.Main_MasterPath & GetWIPFileName(), ReadOnly:=True
    
    ActivateWIPWorkbook
    Range("A1").Select
    Do
        ActiveCell.Offset(0, 1).Select
    Loop Until ActiveCell.Value = "Converted_JN" Or ActiveCell.FormulaR1C1 = ""
    
    Sortcol = ActiveCell.Address
    
    Range("A3", Range("A3").Offset(0, col - 1).Address).Select
    Range(Selection, Selection.End(xlDown)).Select

    Selection.Sort Key1:=Range(Range(Sortcol).Offset(2, 0).Address), Order1:=xlAscending, Header:=xlNo, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortTextAsNumbers
        
    ShowWorkshopCols
        
    Application.DisplayAlerts = False
        Range("b1").Select
                        
    With ActiveSheet.PageSetup
    
        .CenterHeader = "WORKSHOP JOB NUMBER"
        .RightHeader = "&D &T"
        
    End With
    ActiveWorkbook.SaveAs (Main.Main_MasterPath & "TEMPLATES\Workshop_JobNumber.xls")

End If

'DUEDATE
If fwip.Job_WorkshopDueDate.Value = True Then
    Workbooks.Open Main.Main_MasterPath & GetWIPFileName(), ReadOnly:=True
    
    ActivateWIPWorkbook
    Range("A1").Select
    Do
        ActiveCell.Offset(0, 1).Select
    Loop Until ActiveCell.Value = "Job_WorkshopDueDate" Or ActiveCell.FormulaR1C1 = ""
    
    Sortcol = ActiveCell.Address
    
    Range("A3", Range("A3").Offset(0, col - 1).Address).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Sort Key1:=Range(Range(Sortcol).Offset(2, 0).Address), Order1:=xlAscending, Header:=xlNo, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
    
    ShowWorkshopCols
    
    Application.DisplayAlerts = False
        Range("b1").Select
                            
    With ActiveSheet.PageSetup
    
        .CenterHeader = "WORKSHOP DUE DATE"
        .RightHeader = "&D &T"
        
    End With
    
    ActiveWorkbook.SaveAs (Main.Main_MasterPath & "TEMPLATES\Job_WorkshopDueDate.xls")
   
End If

' If no specific report options are selected, just open the base WIP report
If Not (ROperation.Value Or ROperator.Value Or RDueDate.Value Or RWIP.Value Or _
        Job_DueDate.Value Or Office_Customer.Value Or Workshop_Customer.Value Or _
        Office_JobNumber.Value Or Workshop_JobNumber.Value Or Job_WorkshopDueDate.Value) Then

    ' Open base WIP file for viewing
    Workbooks.Open Main.Main_MasterPath & GetWIPFileName(), ReadOnly:=True
    Range("A1").Select
End If

Application.DisplayAlerts = True

Unload fwip
Unload Main

Exit Sub

End Sub

Private Function ParseJobNumberForSorting(ByVal JobNumber As Variant) As Double
    On Error Resume Next
    ParseJobNumberForSorting = CDbl(JobNumber)
    If Err.Number <> 0 Then
        ParseJobNumberForSorting = 0
        Err.Clear
    End If
    On Error GoTo 0
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
            Case "Component_Comments"
                Selection.EntireColumn.Hidden = False
            Case "Component_Comments"
                Selection.EntireColumn.Hidden = False
            Case "Component_Comments"
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
              Case "Component_Comments"
                Selection.EntireColumn.Hidden = False
            Case "Component_Comments"
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
            Case "Operation13_Type"
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

