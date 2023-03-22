Attribute VB_Name = "Module1"
Sub makro1()
Attribute makro1.VB_ProcData.VB_Invoke_Func = "Q\n14"
'
' makro1 Makro
'
' Klawisz skrotu: Ctrl+Shift+Q
'
    ActiveCell.FormulaR1C1 = "=NOW()"
    Selection.NumberFormat = "[$-x-sysdate]dddd, mmmm dd, yyyy"
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
Sub makro2()
Attribute makro2.VB_ProcData.VB_Invoke_Func = "W\n14"
'
' makro2 Makro
'
' Klawisz skrotu: Ctrl+Shift+W
'
    ActiveCell.FormulaR1C1 = "Kamila Kopacz"
    ActiveCell.Select
    With Selection.Font
        .Name = "Agency FB"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
Sub makro3()
Attribute makro3.VB_ProcData.VB_Invoke_Func = "E\n14"
'
' makro3 Makro
'
' Klawisz skrotu: Ctrl+Shift+E
'
    With Selection.Font
        .Name = "Times New Roman"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Font
        .Name = "Times New Roman"
        .Size = 28
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection.Font
        .Color = -16711681
        .TintAndShade = 0
    End With
    Selection.Columns.AutoFit
    Range("Tabela4[#Headers]").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
Sub makro4()
Attribute makro4.VB_ProcData.VB_Invoke_Func = "R\n14"
'
' makro4 Makro
'
' Klawisz skrotu: Ctrl+Shift+R
'
    Range("D7").Select
    ActiveCell.FormulaR1C1 = "=NOW()"
    Selection.NumberFormat = "0.00"
End Sub
Sub makro5()
Attribute makro5.VB_ProcData.VB_Invoke_Func = "T\n14"
'
' makro5 Makro
'
' Klawisz skrotu: Ctrl+Shift+T
'
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Application.CutCopyMode = False
    Selection.FormulaR1C1 = "=R[-1]C"
End Sub
Sub makro6()
Attribute makro6.VB_ProcData.VB_Invoke_Func = "Y\n14"
'
' makro6 Makro
'
' Klawisz skrotu: Ctrl+Shift+Y
'
    ActiveCell.FormulaR1C1 = "Oceny"
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]>90,5,IF(RC[-1]>70,4,IF(RC[-1]>50,3,2)))"
    Range("Tabela6[[#Headers],[ID]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Columns.AutoFit
    Range("E3").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=AVERAGE(C[-2])"
    Columns("C:C").Select
    ActiveSheet.Shapes.AddChart2(366, xlHistogram).Select
    ActiveSheet.ChartObjects("Wykres 1").Activate
End Sub
Sub makro7()
Attribute makro7.VB_ProcData.VB_Invoke_Func = "U\n14"
'
' makro7 Makro
'
' Klawisz skrotu: Ctrl+Shift+U
'
    ActiveSheet.Shapes.AddChart2(227, xlLine).Select
    ActiveChart.SetSourceData Source:=Sheets(ActiveSheet.Select).Range(Selection)
End Sub
