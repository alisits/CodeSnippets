Attribute VB_Name = "order_by_date_new"
Sub order_by_date()
'
' order_by_date Macro
'not in use - too many loops/not worth
'

'
Application.ScreenUpdating = False
Application.DisplayStatusBar = True
Application.StatusBar = "Working..."
Dim a, b, c, d As Integer
a = 4
b = 33
Call Findposition
Call Expandrevisions

Do While Cells(a, posTitle) <> Empty
 a = a + 1
Loop
c = a - 1
If Cells(c, posTitle) = "End" Then a = c - 1


    Range(Cells(4, 1), Cells(a, b)).Select
    ActiveWorkbook.Worksheets(ActiveSheet.Name).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(ActiveSheet.Name).Sort.SortFields.Add Key:=Range(Cells(4, 6), Cells(a, 6)) _
    , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(ActiveSheet.Name).Sort
        .SetRange Range(Cells(4, 1), Cells(a, b))
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Range(Cells(4, 1), Cells(a, b)).Select
    ActiveWorkbook.Worksheets(ActiveSheet.Name).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(ActiveSheet.Name).Sort.SortFields.Add Key:=Range(Cells(4, 7), Cells(a, 7)) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(ActiveSheet.Name).Sort
        .SetRange Range(Cells(4, 1), Cells(a, b))
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Range(Cells(4, 1), Cells(a, b)).Select
    ActiveWorkbook.Worksheets(ActiveSheet.Name).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(ActiveSheet.Name).Sort.SortFields.Add Key:=Range(Cells(4, 8), Cells(a, 8)) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(ActiveSheet.Name).Sort
        .SetRange Range(Cells(4, 1), Cells(a, b))
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Range(Cells(4, 1), Cells(a, b)).Select
    ActiveWorkbook.Worksheets(ActiveSheet.Name).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(ActiveSheet.Name).Sort.SortFields.Add Key:=Range(Cells(4, 9), Cells(a, 9)) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(ActiveSheet.Name).Sort
        .SetRange Range(Cells(4, 1), Cells(a, b))
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Range(Cells(4, 1), Cells(a, b)).Select
    ActiveWorkbook.Worksheets(ActiveSheet.Name).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(ActiveSheet.Name).Sort.SortFields.Add Key:=Range(Cells(4, 10), Cells(a, 10)) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(ActiveSheet.Name).Sort
        .SetRange Range(Cells(4, 1), Cells(a, b))
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Range(Cells(4, 1), Cells(a, b)).Select
    ActiveWorkbook.Worksheets(ActiveSheet.Name).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(ActiveSheet.Name).Sort.SortFields.Add Key:=Range(Cells(4, 11), Cells(a, 11)) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(ActiveSheet.Name).Sort
        .SetRange Range(Cells(4, 1), Cells(a, b))
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

Call collapserevisions

Application.ScreenUpdating = True
Application.StatusBar = "Done."
End Sub


Sub Format_Dates()

Selection.NumberFormat = “dd / mm / yyyy”

End Sub
