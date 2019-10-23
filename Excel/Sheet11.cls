VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Private Sub Worksheet_Change(ByVal Target As Range) '
' Format Macro for review date
'On Error Resume Next
'Application.ScreenUpdating = False

'Dim myRange As Range
'Dim cell As Range


'Set myRange = Range("l4:L50000")
'If Not Target.Column = 12 Then
 '   Exit Sub
'Else
'For Each cell In myRange
    'If (cell.Value = "30.12.1901") Then
    'MsgBox "celda"
     '   cell.Interior.ColorIndex = 0
      '  cell.Interior.Font.ColorIndex = 0
    'End If
'Next
'End If

'Application.ScreenUpdating = True
'End Sub



'Private Sub Worksheet_SelectionChange(ByVal Target As Range)
 
 'Macro that collapses revisions after removing a filter
 'it runs all the time...or only when selecting certain cells...
' es una mierda y no se como arreglarlo
 
'Application.ScreenUpdating = False
'Dim mg As Range
'Set hoja = Worksheets("STD-1")
'Set mg = Me.Range("C4:C4000")
'If Not Intersect(Target, mg) Is Nothing Then
'If hoja.FilterMode = False Then
'Call collapserevisions
'End If
'End If


'Application.ScreenUpdating = True

'End Sub
Private Sub CommandButton1_Click()
    
    Call OrderOther
    Call OrderOther2
    
End Sub

Private Sub CommandButton2_Click()
    
    Call Show_released
    
End Sub

Private Sub CommandButton3_Click()
    
    Call Show_all
    
End Sub

Private Sub CommandButton4_Click()
    
    Call Collapse
    
End Sub
