Sub Macro1()
'
' Macro1 Macro
' Macro recorded 02/03/2009 by K J Bigham
'

'
    Range("A8891").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Sort Key1:=Range("A3"), Order1:=xlDescending, Header:=xlYes, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal


    Selection.Sort Key1:=Range("A3"), Order1:=xlAscending, Header:=xlYes, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal

    Range("A6").Select
    Selection.End(xlUp).Select
    Range("A5").Select
End Sub
Sub Macro2()
'
' Macro2 Macro
' Macro recorded 02/03/2009 by K J Bigham
'

'
    Selection.Sort Key1:=Range("A2"), Order1:=xlAscending, Header:=xlYes, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortTextAsNumbers
End Sub
