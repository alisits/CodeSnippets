Attribute VB_Name = "Module2"
Sub Macro1()
'
' Macro1 Macro
'Creates the original parison taken imput values from txt file
'Alicia Garcia


Dim Fullfilename
Dim book1, book2 As Workbook
Dim Rno, Rma, Swe, R, uped, loed, L, Rang, Strk, Bdgap, unodes, vnodes, Bdgapmm, z0, carrera, sfdr, pwds, pwds1, pwds2  As Long
Dim j, f, c, m, n, i, k, odd, t1, t2, t, s, countrow, fila1, fila2, fila3, fila4, a, cm As Long
Dim Pi, Anma, z1, z2, xb, yb, difab, utheta, uzero, angle  As Double
Dim gridpath, thefilename, finalname  As String

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


selectfile.Show
 
'Opens the dialog and allows to load the file
Fullfilename = Application.GetOpenFilename
    
Workbooks.OpenText Filename:=Fullfilename, Origin:=xlMSDOS, StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, Semicolon:=False, Comma:=False, Space:=True, Other:=True, OtherChar:="|", DecimalSeparator:=".", ThousandsSeparator:="'"
'Bug: parameters are not alligned-not important


Set book2 = ActiveWorkbook

Application.ScreenUpdating = False
Application.DisplayStatusBar = True
Application.StatusBar = "Working..."

'temperature:
j = 7
book2.Activate
ActiveSheet.Cells(16, "F").Select
Selection.Copy
book1.Activate
Worksheets("Data").Select
For i = 3 To 30722
    Cells(i, j).Select
    ActiveSheet.Paste
Next

'nodes coordinates
book2.Activate
'reads the input data from .txt file in book2 and stores them in variants
Rno = Cells(7, "F").Value 'Radius of nozzle (mm)
Rma = Cells(8, "F").Value 'Radius of mandrel (mm)
Anma = Cells(9, "F").Value 'Angle of mandrel (°)
Anma = Anma * Pi / 180 'Angle of mandrel(rad)
Swe = Cells(26, "E").Value 'Swelling factor
R = ((Rno - Rma) / 2 + Rma) * Swe 'Parison radius (mm)
uped = Cells(11, "G").Value 'upper edge of parison (mm)
loed = Cells(12, "G").Value 'Lower edge of parison (mm)
L = uped - loed 'Parison length (mm)
Rang = Cells(22, "E").Value / 100 'Range of servovalve
Strk = Cells(23, "E").Value  'Max stroke of servovalve (mm)
Bdgap = Cells(24, "E").Value / 100 'Basic die gap (%)
Bdgapmm = Strk * Bdgap 'Basic die gap(mm)
unodes = Cells(27, "I").Value  'Number of nodes in u direction
vnodes = Cells(28, "I").Value  'Number of nodes in v direction
pwdsstrk = 4
z0 = pwdsstrk / Sin(Anma)

book1.Activate
Worksheets("Data").Select

For m = 0 To (vnodes - 1)
    For n = 1 To unodes
        f = 30724 + n + 120 * m
        c = 3
        Cells(f, c).Value = R * Cos((3 * (n - 1)) * Pi / 180)
        Cells(f, c + 1).Value = R * Sin((3 * (n - 1)) * Pi / 180)
        Cells(f, c + 2).Value = -L + L * (m) / 128
    Next
Next
'Actions for wall thickness values:
'Assignation of thickness values to the elements in the mesh only attending to VWDS.
carrera = Strk * Rang

For countrow = 1 To ((vnodes - 1) / 2)
    book2.Activate
    fila1 = 36 + countrow
    fila2 = fila1 + 1
    fila3 = 200
    fila4 = 36 + countrow
    
    z1 = Cells(fila1, 4).Value
    ActiveSheet.Cells(fila2, "D").Select
    If ActiveCell.Value = Empty Then
        z2 = z1
    Else
        z2 = ActiveCell.Value
    End If
    z1 = (z1 * carrera / 100) + Bdgapmm + z0
    z2 = (z2 * carrera / 100) + Bdgapmm + z0
    t1 = z1 * Sin(Anma)
    t2 = z2 * Sin(Anma)
    t = (t1 + t2) / 2
    
    For m = 1 To 2
        For n = 1 To (2 * unodes)
            elnum = (countrow - 1) * 4 * unodes + (m - 1) * 2 * unodes + n
            odd = (m Mod 2)
            angle = (n - 1) * 3
            
            'SFDR
            book2.Activate
            Select Case n 'sfdr values
                Case 3, 123
                 fila3 = 109
                Case 7, 127
                 fila3 = 110
                Case 10, 130
                 fila3 = 111
                Case 13, 133
                 fila3 = 112
                Case 17, 137
                 fila3 = 113
                Case 20, 140
                 fila3 = 114
                Case 23, 143
                 fila3 = 115
                Case 27, 147
                 fila3 = 116
                Case 30, 150
                 fila3 = 117
                Case 33, 153
                 fila3 = 118
                Case 37, 157
                 fila3 = 119
                Case 40, 160
                 fila3 = 120
                Case 43, 163
                 fila3 = 121
                Case 47, 167
                 fila3 = 122
                Case 50, 170
                 fila3 = 123
                Case 53, 173
                 fila3 = 124
                Case 57, 177
                 fila3 = 125
                Case 60, 180
                 fila3 = 126
                Case 63, 183
                 fila3 = 127
                Case 67, 187
                 fila3 = 128
                Case 70, 190
                 fila3 = 129
                Case 73, 193
                 fila3 = 130
                Case 77, 197
                 fila3 = 131
                Case 80, 200
                 fila3 = 132
                Case 83, 203
                 fila3 = 133
                Case 87, 207
                 fila3 = 134
                Case 90, 210
                 fila3 = 135
                Case 93, 213
                 fila3 = 136
                Case 97, 217
                 fila3 = 137
                Case 100, 220
                 fila3 = 138
                Case 103, 223
                 fila3 = 139
                Case 107, 227
                 fila3 = 140
                Case 110, 230
                 fila3 = 141
                Case 113, 233
                 fila3 = 142
                Case 117, 237
                 fila3 = 143
                Case 120, 240
                 fila3 = 108
                End Select
                
            sfdr = Cells(fila3, 4).Value 'same value is taken for 3/4 nodes (10degrees)
            
        'PWDS
            pwds1 = Cells(fila4, 6).Value
            pwds2 = Cells(fila4, 8).Value
            a = (pwds2 - pwds1) * pwdsstrk / 100
            xb = (Cos(angle / 180 * Pi)) ^ 2 * (a + ((a ^ 2) - ((a ^ 2) - (Rno ^ 2)) / (Cos(angle / 180 * Pi) ^ 2)) ^ 0.5)
            yb = xb * Tan(angle / 180 * Pi)
            difab = (((((xb ^ 2) + (yb ^ 2)) ^ 0.5) - Rno) * Cos(Anma))
            uzero = (100 - pwds1 - pwds2) * pwdsstrk / 100
            cm = uzero / ((1 / Pi) - (Pi / 8))
            utheta = cm * ((1 / Pi) + ((angle / 180 * Pi - Pi / 2) / 4 * Cos(angle / 180 * Pi) - (Sin(angle / 180 * Pi) / 4))) * Cos(Anma)
            pwds = difab + uzero
            
         'data writing
            If odd = False Then 'if the number is even, average value must be used
                book1.Activate
                Worksheets("Data").Select
                Cells((elnum + 2), ("F")).Value = (t - sfdr + pwds) * Swe
            Else 'if the number is odd, there is a control point in its row
                book1.Activate
                Worksheets("Data").Select
                Cells((elnum + 2), ("F")).Value = (t1 - sfdr + pwds) * Swe
            End If
        Next
     Next
Next
        

book2.Activate
thefilename = ActiveWorkbook.Name & ".xls" 'gets the name from the txt and changes the extension
finalname = gridpath & thefilename 'and the directory to Grid
ActiveWorkbook.Close Saved = True
Application.DisplayAlerts = False
ActiveWorkbook.SaveAs Filename:=finalname, FileFormat:=56

Application.ScreenUpdating = True
Application.DisplayStatusBar = True
Application.StatusBar = "Done."
Application.Calculation = xlCalculationAutomatic

ActiveWorkbook.Close Saved = True

    
End Sub

Sub macro2()

toexit.Show

End Sub



