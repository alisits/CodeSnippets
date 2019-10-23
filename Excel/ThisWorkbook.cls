VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

'Private Sub CloseBook()
 '   Application.DisplayAlerts = False
  '  ActiveWorkbook.Close
   ' Application.DisplayAlerts = True
'End Sub


'Private Sub Workbook_BeforeClose(cancel As Boolean)



'Dim Response As Long
'Response = MsgBox("If you made any modifications you should ORDER the list and UPDATE THE D-QUALITY sheet." & vbCrLf & "If you want to QUIT anyway press YES", vbYesNo)
'If Response = vbNo Then
' cancel = True
'Else
 'cancel = False

'End If
'End Sub








Private Sub Workbook_Open()

Sheets("Help").Select

MsgBox ("Please, take a look at the Help Sheet to get familiar with the document options.")


End Sub
