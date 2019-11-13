Attribute VB_Name = "Module1"

Sub SaveAsTXT()

Dim Mypath, myname As String

'myname = ActiveWorkbook.Name da fallo porque mete también la extensión

Mypath = ActiveWorkbook.Path

ChDir (Mypath)

ActiveWorkbook.SaveAs Filename:="prueba2", FileFormat:=xlText, CreateBackup:=False

ThisWorkbook.Saved = True
Application.Quit

End Sub
