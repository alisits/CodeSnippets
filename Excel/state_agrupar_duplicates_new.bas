Attribute VB_Name = "state_agrupar_duplicates_new"
Option Compare Text

Sub State()

Dim DimCol As Long

DimCol = Find_the_last_row_of_the_column("STD-List")

'Dim Table() As Variant
'ReDim Table(dimCol)
With Worksheets("STD-List")
For i = 4 To DimCol
    If Range("J" & i) <> Empty Then
        Range("F" & i) = "RELEASED"
    ElseIf Range("J" & i) = Empty Then
        Range("F" & i) = "NOT RELEASED"
    End If
    If Range("L" & i) Like "Obsolete" Then
          Range("F" & i) = "OBSOLETE"
    End If
Next i
End With
End Sub

Sub Agrupar()

'currently not in use because revisions are no longer in the list

'there is a bug maybe
Dim a, b, c As Integer
Dim cel1, cel2
a = 4
b = a + 1
c = a + 1

Cells.Select
Selection.Rows.Ungroup

Do While Cells(a, 3) <> Empty
 b = a + 1
 cel1 = Cells(a, 3)
 cel2 = Cells(b, 3)
 If cel1 = cel2 Then
 c = a + 1
 While cel1 = cel2
  b = b + 1
  cel2 = Cells(b, 3)
 Wend
  c = a + 1
  b = b - 1
  Rows(c & ":" & b).Select
  Selection.Rows.Group
 End If
  
  a = b
 Loop
End Sub

Sub duplicates()


Application.ScreenUpdating = False
Application.DisplayStatusBar = True
Application.StatusBar = "Working..."

Dim a, b  As Integer
Dim cel1, cel2, cel3


b = 4
a = b + 1
Do While Cells(b, 3) <> Empty

cel1 = Cells(b, 3)
cel2 = Cells(b, 4)
cel3 = Cells(b, 2)
a = 4

Do While Cells(a, 3) <> Empty
 If a = b Then a = a + 1
 If cel1 = Cells(a, 3) Then
  If cel2 = Cells(a, 4) Then
   If cel3 = Cells(a, 2) Then
   
    Rows(a).Select
    With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 192
            .TintAndShade = 0
            .PatternTintAndShade = 0
    End With
    Exit Do
   End If
  End If
 End If
 a = a + 1
Loop

b = b + 1
Loop
   
MsgBox ("The rows marked in dark red are duplicated")
   
Application.ScreenUpdating = True
Application.StatusBar = "Done."
 
End Sub

