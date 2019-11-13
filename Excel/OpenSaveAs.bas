Attribute VB_Name = "Module3"
Public R, L, grades As Integer

Sub Macro3()
'
' Macro for creation of ideal parison bcs file. To be used in B-Sim
'Alicia Garcia
'TI Automotive Rastatt
'algarcia@de.tiauto.com


Dim Fullfilename
Dim book1, book2 As Workbook
Dim i, j, k As Long
Dim Pi As Double
Dim gridpath, thefilename, finalname  As String
Dim opthickness As Long


Pi = 3.14159265358979


Application.ScreenUpdating = False
Application.DisplayStatusBar = True
Application.StatusBar = "Working..."
Application.Calculation = xlCalculationManual

On Error Resume Next
Application.EnableCancelKey = xlErrorHandler


Set book1 = ThisWorkbook

thefilepath = ActiveWorkbook.Path
gridpath = thefilepath & "\Grid\"
'creates the grid folder if it doesn't exist
If Dir(gridpath, vbDirectory) = "" Then
MkDir gridpath
End If


selectfilemould.Show
 
'Opens the dialog and allows to load the file
Fullfilename = Application.GetOpenFilename
If Fullfilename = "False" Then
    MsgBox "Action cancelled"
    Application.StatusBar = "Done."
    Exit Sub
End If
    
Workbooks.OpenText Filename:=Fullfilename, Origin:=xlMSDOS, StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, Semicolon:=False, Comma:=False, Space:=True, Other:=True, OtherChar:="|", DecimalSeparator:=".", ThousandsSeparator:="'"
'Bug: parameters are not alligned-not important


Set book2 = ActiveWorkbook

'code to delete empty columns
i = 3
j = i
k = i
Do While Cells(j, "B") <> -111
    If Cells(j, "A") = Empty Then k = j
    If Cells(j, "A") = -111 Then Exit Do
    j = j + 1
Loop

If k <> i Then ActiveSheet.Range("A" & i, "A" & k).Delete

j = j + 1
i = j
k = i
Do While Cells(j, "B") <> -111
    If Cells(j, "A") = Empty Then k = j
    If Cells(j, "A") = -111 Then Exit Do
    j = j + 1
Loop
If k <> i Then ActiveSheet.Range("A" & i, "A" & k).Delete

Optimalthicknessvalue.Show

Unload Optimalthicknessvalue

thefilename = ActiveWorkbook.Name & "_ideal.xls" 'gets the name from the txt and adds "ideal"
finalname = gridpath & thefilename 'and the directory to Grid

Application.DisplayAlerts = False
ActiveWorkbook.SaveAs Filename:=finalname
ActiveWorkbook.Close Saved = True


Application.ScreenUpdating = True
Application.DisplayStatusBar = True
Application.StatusBar = "Done."
Application.Calculation = xlCalculationAutomatic



End Sub

