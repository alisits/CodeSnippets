Attribute VB_Name = "Order_new"
Sub Order()


Dim count, start, round  As Integer
Dim almAud, almTrain, almRel, almRev, almCreate, almIdea As Integer
Dim almOngR, almOngC As Integer


count = 4
start = 4


Application.ScreenUpdating = False
Application.DisplayStatusBar = True
Application.StatusBar = "Working..."

Call Findposition
Call Expandall
Call State

Application.ScreenUpdating = False
Application.DisplayStatusBar = True
Application.StatusBar = "Working..."

If ActiveSheet.FilterMode Then
    ActiveSheet.ShowAllData
  End If

'Delete previous End line and insert new one
ActiveSheet.Cells(count, posTitle).Select
Do While ActiveCell <> Empty
 If ActiveCell = "End" Then
  Rows(start).Select
  Selection.Delete
  start = start - 1
 End If
 start = start + 1
 Cells(start, posTitle).Select
Loop
Cells(start, 1) = "End"
Cells(start, 2) = "End"
Cells(start, 3) = "End"
Cells(start, 4) = "End"
Cells(start, 5) = "End"
Cells(start, 6) = "End"
Cells(start, 7) = "End"
Cells(start, 8) = "End"
Cells(start, 9) = "End"
Cells(start, 10) = "End"
Cells(start, 11) = "End"
Cells(start, 12) = "End"
Cells(start, 13) = "End"


'**************************************************************************************
'move obsolete to the start
start = 4
almObs = start
Do While Cells(almObs, posAud) = "Obsolete" 'Avoids to move rows that are well placed
    almObs = almObs + 1
Loop
count = almObs

Do While Cells(count, posTitle) <> "End"
     If ((Cells(count, posAud) = "Obsolete") Or ((Cells(count, posStatus) = "Obsolete"))) Then
        On Error Resume Next
        Rows(count).Select
        Selection.Cut
        Rows(start).Select
        Selection.Insert shift:=xlDown
               
    Else
    count = count + 1
    End If
        
Loop

'**************************************************************************************
'Order from Audit Status

almAud = 4

count = almAud

Do While Cells(count, posTitle) <> "End"
    If ((Cells(count, posAud) <> Empty) And (Cells(count, posAud) <> "Obsolete")) Then 'Looks for rows with Audit Status and moves them to the top
        On Error Resume Next
        Rows(count).Select
        Selection.Cut
        Rows(almAud).Select
        Selection.Insert shift:=xlDown
        
        count = count + 1
        
    Else
    count = count + 1
    
    End If
Loop

'****************************************************************************************
'Order from Train Status

start = 4
almObs = 4

Do While Cells(almObs, posAud) <> "Obsolete" 'store first obsolete row
    almObs = almObs + 1
Loop

almObs2 = almObs

Do While Cells(almObs2, posAud) = "Obsolete" 'store obsolete row
    almObs2 = almObs2 + 1
Loop

almTrain = almObs

'Do While Cells(almTrain, posTrain) <> Empty 'Avoids to move rows that are well placed
 '   almTrain = almTrain + 1
'Loop

count = almObs2

Do While Cells(count, posTitle) <> "End"
    If Cells(count, posTrain) <> Empty Then
    On Error Resume Next
        Rows(count).Select
        Selection.Cut
        Rows(almTrain).Select
        Selection.Insert shift:=xlDown
       
        count = count + 1
        
    Else
    count = count + 1
    End If
        
Loop

'*******************************************************************
'Order from Release status with date
start = 4
almObs = 4

Do While Cells(almObs, posAud) <> "Obsolete" 'store first obsolete row
    almObs = almObs + 1
Loop

almObs2 = almObs

Do While Cells(almObs2, posAud) = "Obsolete" 'store obsolete row
    almObs2 = almObs2 + 1
Loop

almRevD = almObs
count = almObs2

Do While Cells(count, posTitle) <> "End"
    If Cells(count, posRel) <> Empty And IsDate(Cells(count, posRel)) Then
    On Error Resume Next
        Rows(count).Select
        Selection.Cut
        Rows(almRevD).Select
        Selection.Insert shift:=xlDown
       
        count = count + 1
        
    Else
    count = count + 1
    End If
        
Loop



'*******************************************************************
'Order from Release status with X
start = 4
almObs = 4

Do While Cells(almObs, posAud) <> "Obsolete" 'store first obsolete row
    almObs = almObs + 1
Loop

almObs2 = almObs

Do While Cells(almObs2, posAud) = "Obsolete" 'store obsolete row
    almObs2 = almObs2 + 1
Loop

almRevD = almObs
count = almObs2

Do While Cells(count, posTitle) <> "End"
    If Cells(count, posRel) <> Empty Then
    On Error Resume Next
        Rows(count).Select
        Selection.Cut
        Rows(almRevD).Select
        Selection.Insert shift:=xlDown
       
        count = count + 1
        
    Else
    count = count + 1
    End If
        
Loop

'*******************************************************************
'Order from Idea status

start = 4
almObs = 4

Do While Cells(almObs, posAud) <> "Obsolete" 'store first obsolete row
    almObs = almObs + 1
Loop

almObs2 = almObs

Do While Cells(almObs2, posAud) = "Obsolete" 'store obsolete row
    almObs2 = almObs2 + 1
Loop

almRevD = almObs
count = almObs2

Do While Cells(count, posTitle) <> "End"
    If Cells(count, posIdea) = "X" Then
    On Error Resume Next
        Rows(count).Select
        Selection.Cut
        Rows(almObs2).Select
        Selection.Insert shift:=xlDown
       
        count = count + 1
        
    Else
    count = count + 1
    End If
        
Loop

'*******************************************************************
'Order from create ongoing status

start = 4
almObs = 4

Do While Cells(almObs, posAud) <> "Obsolete" 'store first obsolete row
    almObs = almObs + 1
Loop

almObs2 = almObs

Do While Cells(almObs2, posAud) = "Obsolete" 'store obsolete row
    almObs2 = almObs2 + 1
Loop

almRevD = almObs
count = almObs2

Do While Cells(count, posTitle) <> "End"
    If Cells(count, posCreate) = "ONGOING" Then
    On Error Resume Next
        Rows(count).Select
        Selection.Cut
        Rows(almObs2).Select
        Selection.Insert shift:=xlDown
       
        count = count + 1
        
    Else
    count = count + 1
    End If
        
Loop

'*******************************************************************
'Order from create status

start = 4
almObs = 4

Do While Cells(almObs, posAud) <> "Obsolete" 'store first obsolete row
    almObs = almObs + 1
Loop

almObs2 = almObs

Do While Cells(almObs2, posAud) = "Obsolete" 'store obsolete row
    almObs2 = almObs2 + 1
Loop

almRevD = almObs
count = almObs2

Do While Cells(count, posTitle) <> "End"
    If Cells(count, posCreate) = "X" Then
    On Error Resume Next
        Rows(count).Select
        Selection.Cut
        Rows(almObs2).Select
        Selection.Insert shift:=xlDown
       
        count = count + 1
        
    Else
    count = count + 1
    End If
        
Loop
'*******************************************************************
'Order from Review status with "ongoing"
start = 4

almObs = 4

Do While Cells(almObs, posAud) <> "Obsolete" 'store first obsolete row
    almObs = almObs + 1
Loop

almObs2 = almObs

Do While Cells(almObs2, posAud) = "Obsolete" 'store obsolete row
    almObs2 = almObs2 + 1
Loop

almRevD = almObs
count = almObs2

Do While Cells(count, posTitle) <> "End"
    If Cells(count, posReview) = "ONGOING" Then
    On Error Resume Next
        Rows(count).Select
        Selection.Cut
        Rows(almObs2).Select
        Selection.Insert shift:=xlDown
       
        count = count + 1
        
    Else
    count = count + 1
    End If
        
Loop

'*******************************************************************
'Order from Review status with x
start = 4
almObs = 4

Do While Cells(almObs, posAud) <> "Obsolete" 'store first obsolete row
    almObs = almObs + 1
Loop

almObs2 = almObs

Do While Cells(almObs2, posAud) = "Obsolete" 'store obsolete row
    almObs2 = almObs2 + 1
Loop

almRevD = almObs
count = almObs2

Do While Cells(count, posTitle) <> "End"
    If Cells(count, posReview) = "X" Then
    On Error Resume Next
        Rows(count).Select
        Selection.Cut
        Rows(almObs2).Select
        Selection.Insert shift:=xlDown
       
        count = count + 1
        
    Else
    count = count + 1
    End If
        
Loop




'Move end line
start = 4

Do While Cells(start, posTitle) <> "End"
 start = start + 1
Loop
round = 4
Do While Cells(round, posTitle) <> Empty
 round = round + 1
Loop
On Error Resume Next
Rows(start).Select
Selection.Cut
Rows(round).Select
Selection.Insert shift:=xlDown

'Call order_by_date
Call Undo_all
Call format_2_colors



Application.ScreenUpdating = True
Application.StatusBar = "Done."
Application.DisplayAlerts = True
End Sub

