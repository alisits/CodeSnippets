Attribute VB_Name = "Module1"
   Public Function IsoWeekNum(d1 As Date) As Integer
        ' Provided by Daniel Maher.
           'd1=valor de cada elda de AG
              
           Dim d2 As Long
           
           Dim i As Integer
           
        
           
            d2 = DateSerial(Year(d1 - Weekday(d1 - 1) + 4), 1, 3)
           IsoWeekNum = Int((d1 - d2 + Weekday(d2) + 5) / 7) 'numero de la semana es la variable de la funcion
                     
        
           
        End Function


Sub prueba()

 Dim newPowerPoint As PowerPoint.Application
        Dim activeSlide As PowerPoint.Slide
        Dim i As Long
        
     
     'Look for existing instance
        On Error Resume Next
        Set newPowerPoint = GetObject("", "PowerPoint.Application")
        On Error GoTo 0
     
    'Let's create a new PowerPoint
        If newPowerPoint Is Nothing Then
            Set newPowerPoint = New PowerPoint.Application
        End If
        'Else???
        
    'Make a presentation in PowerPoint    esto no funciona
     If newPowerPoint.Presentations.Count = 0 Then
        newPowerPoint.Visible = msoTrue
    Set newPowerPoint = newPowerPoint.Presentations.Open("http://tiweb.intranet.tiauto.com/erc/brandcenter/Shared%20Documents/TI%20Automotive%20Presentation%20Template.pptx", msoCTrue, msoCTrue, msoCTrue)
    'Set newPowerPoint = newPowerPoint.Presentations.Open("T:\Projects\Internal\Design Follow-Up List\Template.pptx", msoCTrue, msoCTrue, msoCTrue)
       '    newPowerPoint.Presentations.Open "T:\Projects\Internal\Design Follow-Up List\Template.potx", Untitled:=msoTrue
           
        End If
    'Show the PowerPoint
        newPowerPoint.Visible = True


End Sub

Sub test()

Dim a As Date
Dim b As Integer

a = "10 / 10 / 2018"
b = IsoWeekNum(a)


End Sub


