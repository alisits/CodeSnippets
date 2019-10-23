Attribute VB_Name = "format_new"
Sub format_2_colors()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True
Application.StatusBar = "Working..."

Dim i, j, ini, fin, c As Integer
Dim myrange As Range

i = 4

Do While Cells(i, 3) <> Empty
 
 With Rows(i).Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
 End With
 Set myrange = Range(Cells(i, "F"), Cells(i, "K"))
  For Each cell In myrange
        cell.Value = UCase(cell.Value)
    Next cell
 With myrange
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .Font.Bold = True
End With
Set myrange = Range(Cells(i, "AB"), Cells(i, "AE"))
  For Each cell In myrange
        cell.Value = UCase(cell.Value)
    Next cell
With myrange
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .Font.Bold = True
End With
 i = i + 1
Loop

i = 4

Do While Cells(i, 3) <> Empty
 ini = i
 fin = ini

 With Range(ini & ":" & fin).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -4.99893185216834E-02
        .PatternTintAndShade = 0
 End With
 

 i = fin
 i = i + 2
 
Loop

Application.ScreenUpdating = True
Application.StatusBar = "Done."

End Sub



Sub format_Quality_DS()
Application.ScreenUpdating = False
Application.DisplayStatusBar = True
Application.StatusBar = "Working..."

Dim i, a, b As Integer
Dim myrange As Range

Sheets("1.2.1 Standards respected DS").Select

a = 12
b = 2
i = 0


Do While Cells(a, 1) <> Empty

 Select Case Cells(a, 1).Value
    Case "Design Guidelines"
         Set myrange = Range(Cells(a, 1), Cells(a, 26))
        myrange.Merge
        With myrange.Font
            .Size = 38
            .Bold = True
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .TintAndShade = 0
        End With
        myrange.RowHeight = 50
    Case "Components"
         Set myrange = Range(Cells(a, 1), Cells(a, 26))
        myrange.Merge
        With myrange.Font
            .Size = 38
            .Bold = True
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .TintAndShade = 0
        End With
        myrange.RowHeight = 50
    Case "Design Elements"
         Set myrange = Range(Cells(a, 1), Cells(a, 26))
        myrange.Merge
        With myrange.Font
            .Size = 38
            .Bold = True
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .TintAndShade = 0
        End With
        myrange.RowHeight = 50
    Case "Functional assemblies"
         Set myrange = Range(Cells(a, 1), Cells(a, 26))
        myrange.Merge
        With myrange.Font
            .Size = 38
            .Bold = True
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .TintAndShade = 0
        End With
        myrange.RowHeight = 50
     Case "Material Specification"
         Set myrange = Range(Cells(a, 1), Cells(a, 26))
        myrange.Merge
        With myrange.Font
            .Size = 38
            .Bold = True
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .TintAndShade = 0
        End With
        myrange.RowHeight = 50
    Case "Drawing templates"
         Set myrange = Range(Cells(a, 1), Cells(a, 26))
        myrange.Merge
        With myrange.Font
            .Size = 38
            .Bold = True
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .TintAndShade = 0
        End With
        myrange.RowHeight = 50
                  
    Case Else
         Range("K" & a & ":N" & a).MergeCells = True
         Range("R" & a & ":Z" & a).MergeCells = True
         Set myrange = Range(Cells(a, 1), Cells(a, 26))
         With myrange.Font
            .Bold = False
            .Size = 36
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .TintAndShade = 0
        End With
        With myrange.Borders
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
      '  Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    '    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With myrange.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With myrange.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With myrange.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With myrange.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With myrange.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With myrange.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With myrange.Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        myrange.RowHeight = 100
    End Select
    
 a = a + 1
Loop
a = a - 1
Set myrange = Range(Cells(13, "O"), Cells(a, "Q"))
    For Each c In myrange
        c.Value = UCase(c.Value)
    Next c
With myrange
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .Font.Bold = False
End With


Range(Cells(12, 2), Cells(a, 26)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
    Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    
Rows(a + 1).Select
With Selection
    .RowHeight = 80
End With

Rows(a + 5).Select
With Selection
    .RowHeight = 80
End With

Cells(a + 2, 2).Select
With Selection.Font
    .Bold = True
    .Strikethrough = False
    .Superscript = False
    .Subscript = False
    .OutlineFont = False
    .Shadow = False
    .Underline = xlUnderlineStyleNone
    .TintAndShade = 0
End With
With Selection
    .RowHeight = 30
End With

Cells(a + 3, 2).Select
With Selection.Font
    .Size = 36
    .Bold = True
    .Italic = False
    .Strikethrough = False
    .Superscript = False
    .Subscript = False
    .OutlineFont = False
    .Shadow = False
    .Underline = xlUnderlineStyleNone
    .TintAndShade = 0
End With
With Selection
    .RowHeight = 50
End With

Cells(a + 4, 2).Select
With Selection.Font
    .Bold = True
    .Italic = False
    .Strikethrough = False
    .Superscript = False
    .Subscript = False
    .OutlineFont = False
    .Shadow = False
    .Underline = xlUnderlineStyleNone
    .TintAndShade = 0
End With
With Selection
    .RowHeight = 30
End With

Application.ScreenUpdating = True
Application.StatusBar = "Done."
End Sub



Sub format_Quality_FT()
Application.ScreenUpdating = False
Application.DisplayStatusBar = True
Application.StatusBar = "Working..."

Dim i, a, b As Integer
Dim myrange As Range

Sheets("1.2.1 Standards respected FT").Select

a = 12
b = 2
i = 0


Do While Cells(a, 1) <> Empty

 Select Case Cells(a, 1).Value
    Case "Design Guidelines"
        Set myrange = Range(Cells(a, 1), Cells(a, 26))
        myrange.Merge
        With myrange.Font
            .Size = 38
            .Bold = True
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .TintAndShade = 0
        End With
        myrange.RowHeight = 50
    Case "Components"
        Set myrange = Range(Cells(a, 1), Cells(a, 26))
        myrange.Merge
        With myrange.Font
            .Size = 38
            .Bold = True
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .TintAndShade = 0
        End With
        myrange.RowHeight = 50
    Case "Design Elements"
        Set myrange = Range(Cells(a, 1), Cells(a, 26))
        myrange.Merge
        With myrange.Font
            .Size = 38
            .Bold = True
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .TintAndShade = 0
        End With
        myrange.RowHeight = 50
    Case "Functional Assemblies"
        Set myrange = Range(Cells(a, 1), Cells(a, 26))
        myrange.Merge
        With myrange.Font
            .Size = 38
            .Bold = True
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .TintAndShade = 0
        End With
        myrange.RowHeight = 50
    Case "Material Specification"
        Set myrange = Range(Cells(a, 1), Cells(a, 26))
        myrange.Merge
        With myrange.Font
            .Size = 38
            .Bold = True
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .TintAndShade = 0
        End With
        myrange.RowHeight = 50
    Case "Drawing templates"
        Set myrange = Range(Cells(a, 1), Cells(a, 26))
        myrange.Merge
        With myrange.Font
            .Size = 38
            .Bold = True
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .TintAndShade = 0
        End With
        myrange.RowHeight = 50
    
    Case Else
         Range("K" & a & ":N" & a).MergeCells = True
         Range("R" & a & ":Z" & a).MergeCells = True
         Set myrange = Range(Cells(a, 1), Cells(a, 26))
         With myrange.Font
            .Bold = False
            .Size = 36
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .TintAndShade = 0
        End With
        With myrange.Borders
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
      '  Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    '    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With myrange.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With myrange.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With myrange.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With myrange.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With myrange.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With myrange.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With myrange.Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        myrange.RowHeight = 100
    End Select
    
 a = a + 1
Loop
a = a - 1

Set myrange = Range(Cells(13, "O"), Cells(a, "Q"))
    For Each c In myrange
        c.Value = UCase(c.Value)
    Next c
With myrange
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .Font.Bold = False
End With


Range(Cells(12, 2), Cells(a, 26)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
    Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous

Rows(a + 1).Select
With Selection
    .RowHeight = 80
End With

Rows(a + 5).Select
With Selection
    .RowHeight = 80
End With

Cells(a + 2, 2).Select
With Selection.Font
    .Bold = True
    .Strikethrough = False
    .Superscript = False
    .Subscript = False
    .OutlineFont = False
    .Shadow = False
    .Underline = xlUnderlineStyleNone
    .TintAndShade = 0
End With
With Selection
    .RowHeight = 30
End With

Cells(a + 3, 2).Select
With Selection.Font
    .Size = 36
    .Bold = True
    .Italic = False
    .Strikethrough = False
    .Superscript = False
    .Subscript = False
    .OutlineFont = False
    .Shadow = False
    .Underline = xlUnderlineStyleNone
    .TintAndShade = 0
End With
With Selection
    .RowHeight = 50
End With

Cells(a + 4, 2).Select
With Selection.Font
    .Bold = True
    .Italic = False
    .Strikethrough = False
    .Superscript = False
    .Subscript = False
    .OutlineFont = False
    .Shadow = False
    .Underline = xlUnderlineStyleNone
    .TintAndShade = 0
End With
With Selection
    .RowHeight = 30
End With

Application.ScreenUpdating = True
Application.StatusBar = "Done."
End Sub

Sub FormatSTDList()
    
    Dim DimCol
    DimCol = Find_the_last_row_of_the_column("STD-List")
    DimCol = DimCol - 1
    Dim a As Integer
    Dim b As Integer
    
    
    
    With Range(Cells(4, 1), Cells(DimCol, 34)).Borders
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
     End With
    
    
    With Range(Cells(4, 1), Cells(DimCol, 1))  ' Category settings
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .RowHeight = 25.5
        .ColumnWidth = 14
    End With
    
    With Range(Cells(4, 1), Cells(DimCol, 1)).Font   ' General settings
        .Name = "Calibri"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    
    
    With Range(Cells(4, 2), Cells(DimCol, 2))   ' ID Settings
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .RowHeight = 25.5
        .ColumnWidth = 13
    End With
    
    With Range(Cells(4, 2), Cells(DimCol, 2)).Font
        .Name = "Calibri"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    
    
    With Range(Cells(4, 3), Cells(DimCol, 3)).Font  ' Title settings
        .Bold = True
        .Name = "Calibri"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    
    With Range(Cells(4, 3), Cells(DimCol, 3))
         .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .ColumnWidth = 45.33
    End With


    With Range(Cells(4, 4), Cells(DimCol, 4))   ' Rev. settings
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .ColumnWidth = 6.56
    End With
    
    With Range(Cells(4, 4), Cells(DimCol, 4)).Font
        .Name = "Calibri"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    
    
    With Range(Cells(4, 5), Cells(DimCol, 5))   ' Former ID's settings
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .ColumnWidth = 25.5
    End With
    
    With Range(Cells(4, 5), Cells(DimCol, 5)).Font
        .Name = "Calibri"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    
    
    With Range(Cells(4, 6), Cells(DimCol, 6))   ' Status settings
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .ColumnWidth = 12
    End With
    
    With Range(Cells(4, 6), Cells(DimCol, 6)).Font
        .Name = "Calibri"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    
    With Range(Cells(4, 7), Cells(DimCol, 9))   ' Idea : Review settings
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .ColumnWidth = 9
    End With
    
    With Range(Cells(4, 7), Cells(DimCol, 9)).Font
        .Name = "Calibri"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    
    a = 10                                  ' Release : "Review Date" settings
    b = 13
    With Range(Cells(4, a), Cells(DimCol, b))
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .ColumnWidth = 10
    End With
    With Range(Cells(4, a), Cells(DimCol, b)).Font
        .Name = "Calibri"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    
    a = 14                                  ' Owner : Department responsibility settings
    b = 16
    With Range(Cells(4, a), Cells(DimCol, b))
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .ColumnWidth = 10
    End With
    With Range(Cells(4, a), Cells(DimCol, b)).Font
        .Name = "Calibri"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With

    a = 17                                   ' Competence group design settings
    b = 17
    With Range(Cells(4, a), Cells(DimCol, b))
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .ColumnWidth = 20
    End With
    With Range(Cells(4, a), Cells(DimCol, b)).Font
        .Name = "Calibri"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With

    a = 18                                    ' Competence group application settings
    b = 18
    With Range(Cells(4, a), Cells(DimCol, b))
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .ColumnWidth = 13.33
    End With
    With Range(Cells(4, a), Cells(DimCol, b)).Font
        .Name = "Calibri"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    
    a = 19                                  ' Material Engineering : Supporting Others settings
    b = 28
    With Range(Cells(4, a), Cells(DimCol, b))
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .ColumnWidth = 10
    End With
    With Range(Cells(4, a), Cells(DimCol, b)).Font
        .Name = "Calibri"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With

    a = 29                                  ' Division : AP settings
    b = 32
    With Range(Cells(4, a), Cells(DimCol, b))
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .ColumnWidth = 4.56
    End With
    With Range(Cells(4, a), Cells(DimCol, b)).Font
        .Name = "Calibri"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
        .Bold = True
    End With
    
    a = 33                                  ' Storage settings
    b = 33
    With Range(Cells(4, a), Cells(DimCol, b))
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .ColumnWidth = 12
    End With
    With Range(Cells(4, a), Cells(DimCol, b)).Font
        .Name = "Calibri"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
        .Bold = True
    End With
    
    a = 34                                  ' Task to do settings
    b = 34
    With Range(Cells(4, a), Cells(DimCol, b))
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .ColumnWidth = 38
    End With
    With Range(Cells(4, a), Cells(DimCol, b)).Font
        .Name = "Calibri"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
        .Bold = False
    End With
    
End Sub






