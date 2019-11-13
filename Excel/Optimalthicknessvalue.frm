VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Optimalthicknessvalue 
   Caption         =   "Optimal Thickness Value"
   ClientHeight    =   3504
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   4608
   OleObjectBlob   =   "Optimalthicknessvalue.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Optimalthicknessvalue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public Sub CommandButton1_Click()

Dim bookparison, bookmolde As Workbook
Dim opthickness, parisonthick, mouldthick, temperature, result, upperlimit As Long
Dim i, j, k, m, n, f, c, L, R, unodes, vnodes, fila As Long
Pi = 3.14159265358979
vnodes = 129
unodes = 120
upperlimit = 15


Set bookmolde = ActiveWorkbook

opthickness = TextBox1.Value
opthickness = CDbl(opthickness)


If (TextBox1 = Empty) Then
    MsgBox "Please insert the value."
    Exit Sub
End If

If Not IsNumeric(opthickness) Then
    MsgBox "Please, enter a correct number"
    Exit Sub
End If


selectfilemacro2.Show 'to load the previous bcs file
Fullfilename = Application.GetOpenFilename
If Fullfilename = "False" Then
    MsgBox "Action cancelled"
    Application.StatusBar = "Done."
    End
End If
Workbooks.OpenText Filename:=Fullfilename, Origin:=xlMSDOS, StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, Semicolon:=False, Comma:=False, Space:=True, Other:=True, OtherChar:="|", DecimalSeparator:=".", ThousandsSeparator:="'"
Optimalthicknessvalue.Hide

Set bookparison = ActiveWorkbook
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


'Read values for temperature, radius and length
temperature = Cells(3, "F").Value
Rotated.Show
If grades = 0 Then
 fila = 30725
Else
 grades = 360 - grades
 fila = grades * unodes / 360
 fila = Round(fila, 0)
 fila = fila + 30724
End If

R = Cells(fila, "B").Value
L = -Cells(fila, "D").Value


Unload Rotated
'Start calculating optimal parison

i = 3
Do While Cells(i, "A") <> -111
    If Cells(i, "A") = "Char." Then Exit Do
    bookparison.Activate
    parisonthick = Cells(i, "E").Value
    bookmolde.Activate
    mouldthick = Cells(i, "E").Value
    Cells(i, "E").Value = opthickness * parisonthick / (mouldthick)
    If Cells(i, "E").Value < opthickness Then
        Cells(i, "E").Value = opthickness
    Else
        If Cells(i, "E").Value > upperlimit Then Cells(i, "E").Value = upperlimit
    End If
    Cells(i, "F").Value = temperature
    i = i + 1
 Loop

'Coordinates
bookmolde.Activate


For m = 0 To (vnodes - 1)
    For n = 1 To unodes
        f = i + 1 + n + 120 * m
        c = 2
        Cells(f, c).Value = R * Cos((3 * (n - 1)) * Pi / 180)
        Cells(f, c + 1).Value = R * Sin((3 * (n - 1)) * Pi / 180)
        Cells(f, c + 2).Value = -L + L * (m) / 128
    Next
Next


bookmolde.Activate
End Sub
Function isnumber(ByVal Value As String) As Boolean
    Dim DP As String
    DP = Format$(0, ",")
    isnumber = Not (Value Like "" & DP & "")
 
End Function


Private Sub Label3_Click()

End Sub

Private Sub UserForm_Click()

End Sub
