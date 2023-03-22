Attribute VB_Name = "Module2"
Sub makro8()
Attribute makro8.VB_ProcData.VB_Invoke_Func = "I\n14"
'
' makro8 Makro
'
' Klawisz skrotu: Ctrl+Shift+I
'
    Range("E2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(R[0]C[-2]<=100,AND(R[0]C[-2]+R[0]C[-1]>=60)),1000,""Brak premii"")"
    Selection.Columns.AutoFit
    Range("E2").Select
    Selection.AutoFill Destination:=Range("E2:E26")
    Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Selection.NumberFormat = _
        "_-* #,##0.00 [$zł-pl-PL]_-;-* #,##0.00 [$zł-pl-PL]_-;_-* ""-""?? [$zł-pl-PL]_-;_-@_-"
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "Suma A"
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "Suma B"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "=SUM(C[-5])"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "=SUM(C[-5])"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "=SUM(C[-2])"
    Range("G1").Select
    Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Selection.NumberFormat = _
        "_-* #,##0.00 [$zł-pl-PL]_-;-* #,##0.00 [$zł-pl-PL]_-;_-* ""-""?? [$zł-pl-PL]_-;_-@_-"
    Range("H1:I2").Select
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=Sheets(ActiveSheet.Select).Range(Selection)
End Sub

