Attribute VB_Name = "Module1"
Sub filterDemo()
Attribute filterDemo.VB_ProcData.VB_Invoke_Func = "p\n14"
'
' filterDemo ����
'
' �ֳt��: Ctrl+p
'
    Range("A2:B414").Select
    ActiveWindow.SmallScroll Down:=-414
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Add2 Key:=Range("B2:B414") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("�u�@��1").Sort
        .SetRange Range("A2:B414")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "=SUM(R2C2:R414C2)"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "=AVERAGE(R2C2:R414C2)"
    Range("H3").Select
End Sub
