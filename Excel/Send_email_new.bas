Attribute VB_Name = "Send_email_new"




Sub Send_email()
'On Error Resume Next
Dim CurrentWB As Workbook
Dim CurrentWS As Worksheet
Set CurrentWB = ActiveWorkbook
Set CurrentWS = ActiveWorkbook.Sheets("Data_base")

Dim Email_Subject, Email_Send_From, Email_Body As String, i As Integer
Dim Mail_Object, nameList As String, o As Variant, ws As Worksheet, ws1 As Worksheet
Dim email1, email2 As String
Call Findposition


    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With
    
CurrentWS.Activate
i = 2
Do While Cells(i, 17) <> Empty
 '   Select.sheets("data_base".cells(i,20)

    If Sheets("Data_base").Cells(i, 20).Value <> "" Then
        nameList = nameList & ";" & Sheets("data_base").Cells(i, 20).Value
    End If
    i = i + 1
Loop

Set CurrentWS = ActiveWorkbook.Sheets("STD-List")
CurrentWS.Activate


' Set ws = Worksheets("Data_Base")
    'show next 12 weeks only
' code for next 12 weeks
 ddate = DateAdd("ww", 12, Date)
 
 a = 4
 b = 12
 
  Do While Cells(a, posTitle) <> Empty
'if there is no date in the cell
  
  If IsDate(Cells(a, posRevDate)) = False Then
   review = "31.12.3000"
  Else: review = Cells(a, posRevDate)
  End If
  
   If (review < ddate) Then
    Rows(a).Hidden = False
   Else: Rows(a).Hidden = True
   End If
   a = a + 1
   Loop
   'ask before sending
    ans = MsgBox("Are you sure you want to send review emails to the standards champions ??", vbYesNo)
        If ans = vbNo Then
         Call Show_all
         Exit Sub
        End If

  'Copy the sheet STD1
CurrentWB.Activate
ActiveWorkbook.Sheets("STD-List").Copy

'Send Email to defined adresses
  '      Application.CutCopyMode = False
    TempFilePath = Environ$("temp") & "\"
    TempFileName = "temp1"
    FileExtStr = ".xlsx"
    FileFormatNum = 51
    Set CurrentWB = ActiveWorkbook
    CurrentWB.Activate
    Application.DisplayAlerts = False
        With ActiveWorkbook
            .SaveAs TempFilePath & TempFileName & FileExtStr, FileFormat:=FileFormatNum
        End With
   Application.DisplayAlerts = True
        Set Mail_Object = CreateObject("Outlook.Application")
        With Mail_Object.CreateItem(o)
            .Subject = "Standards to review soon - " & Format(Date, "dd/mm/yyyy")
           ' .to = ws.Cells(17, i).Value
            .To = nameList
            .Body = "Dear Standards Champions, attached a list of the standards in need of review. Please consider the ones you are responsible of. Thanks."
            .Attachments.Add ActiveWorkbook.FullName
            .display
            '.Send 'Will send straight away use .display to send manually
        End With
        With ActiveWorkbook
            .Close SaveChanges:=False
        End With
         
    Kill TempFilePath & TempFileName & FileExtStr
    Set OutMail = Nothing
    Set OutApp = Nothing
    
    Set CurrentWB = ActiveWorkbook
    CurrentWB.Activate
    
    Call Show_all
    

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
End Sub

