Attribute VB_Name = "Order_other2"
Option Explicit
Option Base 1
Option Compare Text

Public Sub OrderOther2()

Application.ScreenUpdating = False
Application.DisplayStatusBar = True
Application.StatusBar = "Working..."

Dim i As Long
Dim j As Long   'counter
Dim DimCol As Long
Dim k As Long
Dim l As Long
Dim m As Long
Dim n As Long
Dim o As Long
Dim p As Long
Dim r As Long
Dim q As Long
Dim s As Long
Dim t As Long
Dim w As Long
Dim total As Long

Const start = 4
DimCol = Find_the_last_row_of_the_column("STD-List")  '2.1
DimCol = DimCol - 1
Const dimRow = 34

Dim Decision As Boolean

Const StatusNum = 6
Const CreateNum = 8
Const ReviewNum = 9
Const ReleaseNum = 10
Const TrainNum = 11
Const AuditNum = 12

'                                               ' Following part of the code has no influent for performance, but can be used to check if constant value above have right values...
'Dim Names(18) As String                             ' 18 is number of column we would like to control during run of the code
'
'Names(1) = "Category"                       '2.2
'Names(2) = "ID"
'Names(3) = "Title"
'Names(4) = "Rev."
'Names(5) = "Former ID's"
'Names(6) = "Status"
'Names(7) = "Idea"
'Names(8) = "Create"
'Names(9) = "Review"
'Names(10) = "Release"
'Names(11) = "Train"
'Names(12) = "Audit"
'Names(13) = "Review date"
'Names(14) = "Division"
'Names(15) = "EU"
'Names(16) = "NA"
'Names(17) = "AP"
'Names(18) = "Tasks to do"
'
'Dim Positiontable() As Integer
'ReDim Positiontable(UBound(Names)) As Integer
'
'For i = 1 To UBound(Names)                                '2.3
'    Positiontable(i) = Findposition_NEW(Names(i))
'        If Positiontable(i) = 0 Then
'            Exit Sub
'        End If
'Next i

Call State                                          '2.4
'Positiontable(8) = 8           ' Hesitant point. R:"Creation" instead of "create" ("I1")

Dim EntireDataOutput() As Variant
Dim EntireData() As Variant
ReDim Preserve EntireData(dimRow, DimCol - 3)
ReDim Preserve EntireDataOutput(dimRow, DimCol - 3)

For j = 4 To DimCol                                 '2.5
    For i = 1 To dimRow
    
        EntireData(i, j - 3) = Cells(j, i)
    
    Next i
Next j

Dim ObsoleteTable() As Variant
Dim ReleasedTable() As Variant
Dim ReleasedTableDateAudit() As Variant
Dim ReleasedTableDateTrain() As Variant
Dim ReleasedTableDate() As Variant
Dim NotReleasedTableIdea() As Variant
Dim NotReleasedTableCreate() As Variant
Dim NotReleasedTableReview() As Variant
Dim NotReleasedTableCreateONGOING() As Variant
Dim NotReleasedTableReviewONGOING() As Variant

l = 0
m = 0
n = 0
o = 0
p = 0
r = 0
q = 0
s = 0
t = 0
w = 0

For j = 1 To DimCol - 3                                 '2.6
    
    If EntireData(12, j) = "Obsolete" Then
        l = l + 1
        ReDim Preserve ObsoleteTable(1 To dimRow, 1 To l)
    
            For k = 1 To dimRow
                ObsoleteTable(k, l) = EntireData(k, j)
            Next k
    ElseIf EntireData(StatusNum, j) = "Not Released" And EntireData(CreateNum, j) = "X" And EntireData(ReviewNum, j) = "X" Then
        m = m + 1
        ReDim Preserve NotReleasedTableReview(1 To dimRow, 1 To m)
            
            For k = 1 To dimRow
                NotReleasedTableReview(k, m) = EntireData(k, j)
            Next k
    ElseIf EntireData(StatusNum, j) = "Not Released" And EntireData(CreateNum, j) = "X" And EntireData(ReviewNum, j) = "ONGOING" Then
        o = o + 1
        ReDim Preserve NotReleasedTableReviewONGOING(1 To dimRow, 1 To o)
            
            For k = 1 To dimRow
                NotReleasedTableReviewONGOING(k, o) = EntireData(k, j)
            Next k
    ElseIf EntireData(StatusNum, j) = "Not Released" And EntireData(CreateNum, j) = "X" Then
        p = p + 1
        ReDim Preserve NotReleasedTableCreate(1 To dimRow, 1 To p)
            
            For k = 1 To dimRow
                NotReleasedTableCreate(k, p) = EntireData(k, j)
            Next k
    ElseIf EntireData(StatusNum, j) = "Not Released" And EntireData(CreateNum, j) = "ONGOING" Then
        n = n + 1
        ReDim Preserve NotReleasedTableCreateONGOING(1 To dimRow, 1 To n)
            
            For k = 1 To dimRow
                NotReleasedTableCreateONGOING(k, n) = EntireData(k, j)
            Next k
    ElseIf EntireData(StatusNum, j) = "NOT RELEASED" Then
        r = r + 1
        ReDim Preserve NotReleasedTableIdea(1 To dimRow, 1 To r)
            
            For k = 1 To dimRow
                NotReleasedTableIdea(k, r) = EntireData(k, j)
            Next k
    ElseIf EntireData(StatusNum, j) = "Released" And EntireData(ReleaseNum, j) <> "" And EntireData(AuditNum, j) <> "" Then
        q = q + 1
        ReDim Preserve ReleasedTableDateAudit(1 To dimRow, 1 To q)
            
            For k = 1 To dimRow
                ReleasedTableDateAudit(k, q) = EntireData(k, j)
            Next k
    ElseIf EntireData(StatusNum, j) = "Released" And EntireData(ReleaseNum, j) <> "" And EntireData(TrainNum, j) <> "" Then
        s = s + 1
        ReDim Preserve ReleasedTableDateTrain(1 To dimRow, 1 To s)
            
            For k = 1 To dimRow
                ReleasedTableDateTrain(k, s) = EntireData(k, j)
            Next k
    ElseIf EntireData(StatusNum, j) = "Released" And EntireData(ReleaseNum, j) <> "" And EntireData(ReleaseNum, j) <> "X" Then
        t = t + 1
        ReDim Preserve ReleasedTableDate(1 To dimRow, 1 To t)
            
            For k = 1 To dimRow
                ReleasedTableDate(k, t) = EntireData(k, j)
            Next k
    ElseIf EntireData(StatusNum, j) = "Released" And EntireData(ReleaseNum, j) <> "" Then
        w = w + 1
        ReDim Preserve ReleasedTable(1 To dimRow, 1 To w)
            
            For k = 1 To dimRow
                ReleasedTable(k, w) = EntireData(k, j)
            Next k
    Else                                                             '2.6
        MsgBox "Some rows are not accurately fullfiled"
        Exit Sub
    End If
Next j

total = l + m + n + o + p + r + q + s + t + w

ObsoleteTable = Application.Transpose(ObsoleteTable)
NotReleasedTableIdea = Application.Transpose(NotReleasedTableIdea)
NotReleasedTableCreate = Application.Transpose(NotReleasedTableCreate)
NotReleasedTableReview = Application.Transpose(NotReleasedTableReview)
NotReleasedTableCreateONGOING = Application.Transpose(NotReleasedTableCreateONGOING)
NotReleasedTableReviewONGOING = Application.Transpose(NotReleasedTableReviewONGOING)
ReleasedTableDateAudit = Application.Transpose(ReleasedTableDateAudit)
ReleasedTable = Application.Transpose(ReleasedTable)
ReleasedTableDateTrain = Application.Transpose(ReleasedTableDateTrain)
ReleasedTableDate = Application.Transpose(ReleasedTableDate)

'2.7

EntireDataOutput = addTables2(ReleasedTableDateAudit, ReleasedTableDateTrain, ReleasedTableDate, ReleasedTable, ObsoleteTable, _
NotReleasedTableReview, NotReleasedTableReviewONGOING, NotReleasedTableCreate, NotReleasedTableCreateONGOING, NotReleasedTableIdea)


    For i = 1 To dimRow                 '2.8
        
                Range(Cells(4, i), Cells(DimCol, dimRow)) = OneColumn(EntireDataOutput, i)
        
    Next i
    
    Range("M4").Formula = "=IF(L4="""",IF(K4="""",IF(J4="""","""",J4+1095),IF(K4=""x"", J4+1095, K4+1095)),L4+1095)"
    
    Range("M4").AutoFill Destination:=Range("M4:M" & DimCol), Type:=xlFillDefault
    
'FormatingQuestion.Show                  '2.9
'UpdateQuestion.Show

'Unload FormatingQuestion
'If FormatingQuestion.decision Then
'    Call FormatSTDList
'End If


Application.ScreenUpdating = True
Application.StatusBar = "Done."
Application.DisplayAlerts = True

End Sub


