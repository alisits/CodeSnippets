Attribute VB_Name = "CheckResult_F"
Option Explicit
Option Base 1
Option Compare Text

Public Sub CheckResult(ByVal sheetName As String)

    Dim DimCol As Long
    Dim i As Long
    Dim Counter As Long
    Dim z As Long
    Const NormalRowHeight = 50
    
    Counter = 0
    
    DimCol = Find_the_last_row_of_the_column(sheetName)
    z = DimCol
    
    With Worksheets(sheetName)
        
        For i = 8 To DimCol
                
        If .Range("D" & i) <> "x" And .Range("D" & i) <> "" Or .Range("E" & i) <> "x" And .Range("E" & i) <> "" Or _
         .Range("G" & i) <> "x" And .Range("G" & i) <> "" Or .Range("F" & i) <> "x" And .Range("F" & i) <> "" Then
            MsgBox "Please use only 'x' or leave empty cell in cells related to status of the standard"
            .Range("C" & z) = "???"
            Exit Sub
        End If
                
            If .Range("D" & i) = "x" Then
                If .Range("D" & i) <> .Range("E" & i) Then Counter = Counter + 1
                    If .Range("E" & i) = "x" And .Range("F" & i) = "x" Or .Range("E" & i) = "x" And .Range("G" & i) = "x" _
                    Or .Range("F" & i) = "x" And .Range("G" & i) = "x" Then
                        MsgBox "Status of the standatd should be or OK or NOK or Not Done. Please check your answers."
                        .Range("C" & z) = "???"
                        Exit Sub
                    End If
            End If
        
        Next i
    
    If Counter = 0 Then
        
        With .Range("C" & z)
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeRight).Weight = xlMedium
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlEdgeLeft).Weight = xlMedium
        .Interior.Color = 5287936               ' Green
        .RowHeight = 2 * NormalRowHeight
        End With
    
        .Range("C" & z).HorizontalAlignment = xlCenter
        .Range("C" & z).VerticalAlignment = xlCenter
        .Range("C" & z) = "OK"
        .Range("C" & z).Font.Bold = True
        .Range("C" & z).Font.Size = 38
        
    Else
        
        With .Range("C" & z)
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeRight).Weight = xlMedium
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlEdgeLeft).Weight = xlMedium
        .Interior.Color = 255                   ' Red
        .RowHeight = 2 * NormalRowHeight
        End With
    
        .Range("C" & z).HorizontalAlignment = xlCenter
        .Range("C" & z).VerticalAlignment = xlCenter
        .Range("C" & z) = "NOT OK"
        .Range("C" & z).Font.Bold = True
        .Range("C" & z).Font.Size = 38
        
    End If
    
    End With
    
End Sub
