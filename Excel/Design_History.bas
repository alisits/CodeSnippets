Attribute VB_Name = "Design_History"
'Module for Design History creation20

'Sub to create a ppt out of the excel list with the completed points
'would be good to use TI template

'Not in use

Sub ExcelToNewPowerPoint()
    Dim PPApp As PowerPoint.Application
    Dim PPPres As PowerPoint.Presentation
    Dim PPSlide As PowerPoint.Slide
    Dim Libro As Workbook
    Dim Hoja1 As Worksheet
    
    Set Libro = Active.Workbook
    Set Hoja1 = Active.Worksheet
       

    ' Create instance of PowerPoint
    Set PPApp = CreateObject("Powerpoint.Application")

    ' For automation to work, PowerPoint must be visible
    ' (alternatively, other extraordinary measures must be taken)
    PPApp.Visible = True

    ' Create a presentation
    Set PPPres = PPApp.Presentations.Add
    'seguramente sea aqui donde hay que elegir la template

    ' Some PowerPoint actions work best in normal slide view
    PPApp.ActiveWindow.ViewType = ppViewSlide

    ' Add first slide to presentation
    Set PPSlide = PPPres.Slides.Add(1, ppLayoutTitleOnly) 'Layout??

    ''---------------------
    '' Do Some Stuff Here
    ''---------------------

   Libro.Activate
   Hoja1.Activate
   
   While (Cells(k, 9) <> Empty)
    If Cells(k, 15) = "Completed" Then
        'copy lines with status completed to the ppt
         Rows(k).Copy
         Set PPSlide = PPPres.Slides.Add(1, 11)
        
        
    End If
    
    k = k + 1
    
   Wend
    



    ' Save and close presentation
   ' With PPPres
  '      .SaveAs "C:\My Documents\MyPreso.ppt"
    '    .Close
   ' End With

    ' Quit PowerPoint
  '  PPApp.Quit

    ' Clean up
  '  Set PPSlide = Nothing
   ' Set PPPres = Nothing
  '  Set PPApp = Nothing

End Sub



Sub CreatePowerPoint()

 'Add a reference to the Microsoft PowerPoint Library by:
    '1. Go to Tools in the VBA menu
    '2. Click on Reference
    '3. Scroll down to Microsoft PowerPoint X.0 Object Library, check the box, and press Okay
 
    'First we declare the variables we will be using
        Dim newPowerPoint As PowerPoint.Application
        Dim activeSlide As PowerPoint.Slide
        Dim i As Long
        
     
     'Look for existing instance
        On Error Resume Next
        Set newPowerPoint = GetObject("", "PowerPoint.Application")
        On Error Resume Next
  
    'Let's create a new PowerPoint
        If newPowerPoint Is Nothing Then
            Set newPowerPoint = New PowerPoint.Application
        End If
        'Else???
         newPowerPoint.Visible = True
    'Make a presentation in PowerPoint    esto no funciona
     If newPowerPoint.Presentations.Count = 0 Then
          '  newPowerPoint.Presentations.Add
           ' newPowerPoint.Activate
            
       'Set newPowerPoint = newPowerPoint.Presentations.Open("T:\Projects\Internal\Design Follow-Up List\Template.pptx", msoCTrue, msoCTrue, msoCTrue)
          '  ActivePresentation.ApplyTemplate ("T:\Projects\Internal\Design Follow-Up List\Template.potx")
       Set newPowerPoint = newPowerPoint.Presentations.Open("http://tiweb.intranet.tiauto.com/erc/brandcenter/Shared%20Documents/TI%20Automotive%20Presentation%20Template.pptx", msoCTrue, msoCTrue, msoCTrue)
        End If
    'Show the PowerPoint
       
    
    'Loop through each chart in the Excel worksheet and paste them into the PowerPoint
    i = 11
    'for each completed task do
        While Cells(i, 10) <> Empty
            If Cells(i, 16) = "Completed" Then
                'Add a new slide
            newPowerPoint.ActivePresentation.Slides.Add newPowerPoint.ActivePresentation.Slides.Count + 1, ppLayoutText
            newPowerPoint.ActiveWindow.View.GotoSlide newPowerPoint.ActivePresentation.Slides.Count
      Set activeSlide = newPowerPoint.ActivePresentation.Slides(newPowerPoint.ActivePresentation.Slides.Count)

                
                'Set the title of the slide the same as the action name
            activeSlide.Shapes(1).TextFrame.TextRange.Text = Cells(i, 10).Text
            
                    
            End If
            i = i + 1
        
        Wend
        
        
     
    AppActivate ("Microsoft PowerPoint")
    Set activeSlide = Nothing
    Set newPowerPoint = Nothing
     
End Sub
