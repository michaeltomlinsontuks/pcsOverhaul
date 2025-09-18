# VBA code extracted from Price List.xls
# Extraction date: Mon Jun  2 11:09:14 SAST 2025

# =======================================================
# Module 8
# =======================================================
Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Cells.Select
    Selection.Sort Key1:=Range("A2"), Order1:=xlAscending, Key2:=Range("B2") _
        , Order2:=xlAscending, Header:=True, OrderCustom:=1, MatchCase:= _
        False, Orientation:=xlTopToBottom
    Range("A1").Select
End Sub


