Attribute VB_Name = "D_Quality_new"
Option Base 1
Option Explicit

Public InputTable() As Variant
Public DataReleased() As Variant
Public DataRELEASEDDivBy7() As Variant
Public DataRELEASEDdim As Long
Public DataRELEASEDdimFT As Long
Public DataRELEASEDdimDS As Long
Public DataRELEASEDft() As Variant
Public DataRELEASEDds() As Variant
Public DataReleasedCatDS() As Variant
Public DataRELEASEDcat1DS() As Variant
Public DataRELEASEDcat2DS() As Variant
Public DataRELEASEDcat3DS() As Variant
Public DataRELEASEDcat4DS() As Variant
Public DataRELEASEDcat5DS() As Variant
Public DataRELEASEDcat6DS() As Variant
Public DataReleasedCatFT() As Variant
Public DataRELEASEDcat1FT() As Variant
Public DataRELEASEDcat2FT() As Variant
Public DataRELEASEDcat3FT() As Variant
Public DataRELEASEDcat4FT() As Variant
Public DataRELEASEDcat5FT() As Variant
Public DataRELEASEDcat6FT() As Variant
Public DataRelCat() As Variant

Dim RangeSubtitle As Range
Dim NumberColumn As String
Dim TitleColumn As String
Dim LastColumn As String
Dim FirstRow As Integer
Dim MandColumn As String
Dim OKColumn As String

Dim Heading(6) As String
Dim NOKColumn As String
Dim NotDoneColumn As String
Dim EUColumn As String
Dim NAColumn As String
Dim APColumn As String
Dim RelDocColumn As String

Const TextSize = 28
Const HeadingSize = 38
Const NormalRowHeight = 50

Public LineNumber As Integer

Sub D_Quality_Update()

Application.ScreenUpdating = False


Dim i As Long
Dim j As Long
Dim z As Long
Dim y As Long
Dim LastRowDS As Long
Dim LastRowFT As Long
Dim Names(8) As String
Dim Position As Integer

Names(1) = "ID"
Names(2) = "Status"
Names(3) = "Title"
Names(4) = "EU"
Names(5) = "NA"
Names(6) = "AP"
Names(7) = "Division"
Names(8) = "Category"



NumberColumn = "A"
TitleColumn = "B"
RelDocColumn = "C"
MandColumn = "D"
OKColumn = "E"
NOKColumn = "F"
NotDoneColumn = "G"
EUColumn = "I"
NAColumn = "J"
APColumn = "K"
LastColumn = "L"

FirstRow = 7
LineNumber = 1

Dim Heading(6) As String

Heading(1) = "Design Guideline"
Heading(2) = "Component"
Heading(3) = "Design Element"
Heading(4) = "Functional assembly"
Heading(5) = "Material Specification"
Heading(6) = "Drawing Template"

Dim InputSizeRow As Long
Const InputSizeColumn = 8
Dim Positiontable(InputSizeColumn) As Long


InputSizeRow = Find_the_last_row_of_the_column("STD-List") - 4

'TransferData2Table
Call CreateInputTable(InputSizeRow)

'Select only released data
Call CreateDataReleased(InputSizeRow, InputSizeColumn)

'Resize
DataRELEASEDdim = UBound(DataReleased, 1)

'Select DS/FT standards
Call CreateDataReleasedDS(InputSizeColumn)
Call CreateDataReleasedFT(InputSizeColumn)

DataRELEASEDdimDS = UBound(DataRELEASEDds, 1)
DataRELEASEDdimFT = UBound(DataRELEASEDft, 1)

'Tables per category
'Hay que recortar las ultimas filas pero hace falta trasponerlo todo porque preserve solo funciona con la ultima dimension
Call CreateDataReleasedCatDS(InputSizeColumn, Heading(1))
DataRELEASEDcat1DS = DataReleasedCatDS
DataRELEASEDcat1DS = DeleteEmptyRows(DataRELEASEDcat1DS)

Call CreateDataReleasedCatDS(InputSizeColumn, Heading(2))
DataRELEASEDcat2DS = DataReleasedCatDS
DataRELEASEDcat2DS = DeleteEmptyRows(DataRELEASEDcat2DS)

Call CreateDataReleasedCatDS(InputSizeColumn, Heading(3))
DataRELEASEDcat3DS = DataReleasedCatDS
DataRELEASEDcat3DS = DeleteEmptyRows(DataRELEASEDcat3DS)

Call CreateDataReleasedCatDS(InputSizeColumn, Heading(4))
DataRELEASEDcat4DS = DataReleasedCatDS
DataRELEASEDcat4DS = DeleteEmptyRows(DataRELEASEDcat4DS)

Call CreateDataReleasedCatDS(InputSizeColumn, Heading(5))
DataRELEASEDcat5DS = DataReleasedCatDS
DataRELEASEDcat5DS = DeleteEmptyRows(DataRELEASEDcat5DS)

Call CreateDataReleasedCatDS(InputSizeColumn, Heading(6))
DataRELEASEDcat6DS = DataReleasedCatDS
DataRELEASEDcat6DS = DeleteEmptyRows(DataRELEASEDcat6DS)

Call CreateDataReleasedCatFT(InputSizeColumn, Heading(1))
DataRELEASEDcat1FT = DataReleasedCatFT
DataRELEASEDcat1FT = DeleteEmptyRows(DataRELEASEDcat1FT)

Call CreateDataReleasedCatFT(InputSizeColumn, Heading(2))
DataRELEASEDcat2FT = DataReleasedCatFT
DataRELEASEDcat2FT = DeleteEmptyRows(DataRELEASEDcat2FT)

Call CreateDataReleasedCatFT(InputSizeColumn, Heading(3))
DataRELEASEDcat3FT = DataReleasedCatFT
DataRELEASEDcat3FT = DeleteEmptyRows(DataRELEASEDcat3FT)

Call CreateDataReleasedCatFT(InputSizeColumn, Heading(4))
DataRELEASEDcat4FT = DataReleasedCatFT
DataRELEASEDcat4FT = DeleteEmptyRows(DataRELEASEDcat4FT)

Call CreateDataReleasedCatFT(InputSizeColumn, Heading(5))
DataRELEASEDcat5FT = DataReleasedCatFT
DataRELEASEDcat5FT = DeleteEmptyRows(DataRELEASEDcat5FT)

Call CreateDataReleasedCatFT(InputSizeColumn, Heading(6))
DataRELEASEDcat6FT = DataReleasedCatFT
DataRELEASEDcat6FT = DeleteEmptyRows(DataRELEASEDcat6FT)

'Update sheets
Dim FTSheet As Worksheet
Dim DSSheet As Worksheet
Set FTSheet = Worksheets("1.2.1 Standards respected FT")
Set DSSheet = Worksheets("1.2.1 Standards respected DS")

LastRowDS = Find_the_last_row_of_the_column("1.2.1 Standards respected DS")
DSSheet.Range("A12", "Z" & LastRowDS).Delete shift:=xlUp
LastRowFT = Find_the_last_row_of_the_column("1.2.1 Standards respected FT")
FTSheet.Range("A12", "Z" & LastRowFT).Delete shift:=xlUp

Call UpdateDQualityFT(FTSheet)
Call UpdateDQualityDS(DSSheet)
    
Application.ScreenUpdating = True

End Sub

Function CreateInputTable(InputDataSizeRow As Long) As Variant

    Dim i As Long
    Dim j As Integer
    Dim Names(8) As String
    Dim Positiontable(8) As Integer
    Dim Position As Integer
    
    Names(1) = "ID"
    Names(2) = "Status"
    Names(3) = "Title"
    Names(4) = "EU"
    Names(5) = "NA"
    Names(6) = "AP"
    Names(7) = "Division"
    Names(8) = "Category"

    ReDim InputTable(InputDataSizeRow, 8)
    For j = 1 To 8
         Positiontable(j) = Findposition_NEW(Names(j))
         For i = 1 To InputDataSizeRow
           
            'ReDim Preserve InputTable(i, j)
                InputTable((i), (j)) = Sheets("STD-List").Cells(i + 3, Positiontable(j))
            
        Next i
    Next j
     
End Function
Function CreateDataReleased(InputDataSizeRow As Long, InputDataSizeColumn As Long) As Variant
    Dim i, j As Long
    
    ReDim Preserve DataReleased(InputDataSizeRow, InputDataSizeColumn)
    For i = 1 To InputDataSizeRow
        
        For j = 1 To InputDataSizeColumn
       
            If InputTable(i, 2) = "RELEASED" Then
                                
                DataReleased((i), (j)) = InputTable(i, j)
                         
            End If
        Next j
    Next i

End Function
Function CreateDataReleasedDS(InputDataSizeColumn) As Variant
    Dim i, j, z As Integer
    z = 1
      ReDim DataRELEASEDds(DataRELEASEDdim, InputDataSizeColumn)
    'ReDim DataRELEASEDds(InputDataSizeColumn, z)
    For i = 1 To DataRELEASEDdim
        If DataReleased(i, 7) = "DS" Or DataReleased(i, 7) = "FTDS" Then
            
            For j = 1 To InputDataSizeColumn
        
                DataRELEASEDds(z, j) = DataReleased(i, j)
            
            Next j
           
            z = z + 1
         '   ReDim Preserve DataRELEASEDds(InputDataSizeColumn, z)
        End If
    Next i
'
End Function
Function CreateDataReleasedFT(InputDataSizeColumn) As Variant
    Dim i, j, z As Integer
    z = 1
    ReDim DataRELEASEDft(DataRELEASEDdim, InputDataSizeColumn)
    For i = 1 To DataRELEASEDdim
        If DataReleased(i, 7) = "FT" Or DataReleased(i, 7) = "FTDS" Then

            For j = 1 To InputDataSizeColumn
        
                DataRELEASEDft(z, j) = DataReleased(i, j)
        
            Next j
            z = z + 1
        End If
    Next i
'    ReDim Preserve DataRELEASEDft(z - 1, InputDataSizeColumn)
End Function
Function CreateDataReleasedCatDS(InputDataSizeColumn As Integer, Heading As String) As Variant
    Dim i, j, z As Integer
    z = 1
    ReDim DataReleasedCatDS(DataRELEASEDdimDS, InputDataSizeColumn)
    For i = 1 To DataRELEASEDdimDS
        If DataRELEASEDds(i, 8) = Heading Then
            
            For j = 1 To InputDataSizeColumn
           
                DataReleasedCatDS(z, j) = DataRELEASEDds(i, j)
        
            Next j
            z = z + 1
        End If
    Next i
'    ReDim Preserve DataReleasedCatDS(z - 1, InputDataSizeColumn)

End Function
Function CreateDataReleasedCatFT(InputDataSizeColumn As Integer, Heading As String) As Variant
    Dim i, j, z As Integer
    z = 1
    ReDim DataReleasedCatFT(DataRELEASEDdimFT, InputDataSizeColumn)
    For i = 1 To DataRELEASEDdimFT
        If DataRELEASEDft(i, 8) = Heading Then
            For j = 1 To InputDataSizeColumn
           
                DataReleasedCatFT(z, j) = DataRELEASEDft(i, j)
        
            Next j
            z = z + 1
          End If
    Next i
'   DataReleasedCatFT = Application.Transpose(DataReleasedCatFT)
'   ReDim Preserve DataReleasedCatFT(InputDataSizeColumn, z)
'   DataReleasedCatFT = Application.Transpose(DataReleasedCatFT)
End Function
Public Function UpdateDQualityFT(Worksheet As Worksheet)

Dim z As Integer
Dim Subtitle As String
Dim DataRelCategory() As Variant

z = FirstRow
Heading(1) = "Design Guideline"
Heading(2) = "Component"
Heading(3) = "Design Element"
Heading(4) = "Functional assembly"
Heading(5) = "Material Specification"
Heading(6) = "Drawing Template"

Worksheet.Activate
 
    Set RangeSubtitle = Range(TitleColumn & z, TitleColumn & z)
    Subtitle = Heading(1)
    Call EditSubtitle(RangeSubtitle, Subtitle)
    DataRelCategory = DataRELEASEDcat1FT
    Call FillStandards(z, DataRelCategory, Worksheet)
    Call FormatStandardRows(z, DataRelCategory)
    
    z = z + UBound(DataRelCategory, 1)
    DataRelCategory = DataRELEASEDcat2FT
    Set RangeSubtitle = Range(TitleColumn & z, TitleColumn & z)
    Subtitle = Heading(2)
    Call EditSubtitle(RangeSubtitle, Subtitle)
    Call FillStandards(z, DataRelCategory, Worksheet)
    Call FormatStandardRows(z, DataRelCategory)
     
    z = z + UBound(DataRelCategory, 1)
    DataRelCategory = DataRELEASEDcat3FT
    Set RangeSubtitle = Range(TitleColumn & z, TitleColumn & z)
    Subtitle = Heading(3)
    Call EditSubtitle(RangeSubtitle, Subtitle)
    Call FillStandards(z, DataRelCategory, Worksheet)
    Call FormatStandardRows(z, DataRelCategory)
    
    z = z + UBound(DataRelCategory, 1)
    DataRelCategory = DataRELEASEDcat4FT
    Set RangeSubtitle = Range(TitleColumn & z, TitleColumn & z)
    Subtitle = Heading(4)
    Call EditSubtitle(RangeSubtitle, Subtitle)
    Call FillStandards(z, DataRelCategory, Worksheet)
    Call FormatStandardRows(z, DataRelCategory)
    
    z = z + UBound(DataRelCategory, 1)
    DataRelCategory = DataRELEASEDcat5FT
    Set RangeSubtitle = Range(TitleColumn & z, TitleColumn & z)
    Subtitle = Heading(5)
    Call EditSubtitle(RangeSubtitle, Subtitle)
    Call FillStandards(z, DataRelCategory, Worksheet)
    Call FormatStandardRows(z, DataRelCategory)
    
    z = z + UBound(DataRelCategory, 1)
    DataRelCategory = DataRELEASEDcat6FT
    Set RangeSubtitle = Range(TitleColumn & z, TitleColumn & z)
    Subtitle = Heading(6)
    Call EditSubtitle(RangeSubtitle, Subtitle)
    Call FillStandards(z, DataRelCategory, Worksheet)
    Call FormatStandardRows(z, DataRelCategory)
    
    z = z + UBound(DataRelCategory, 1)
  
    Call FormatDQualitySheet(Worksheet, z)
    
 End Function
Function FormatDQualitySheet(Worksheet As Worksheet, z As Integer)

    With Worksheet
            .PageSetup.PrintArea = Range("A1", Cells(z, LastColumn)).Address
            .Range(TitleColumn & FirstRow, TitleColumn & z).HorizontalAlignment = xlLeft
            
        With Range(TitleColumn & z)
            .Value = "Result"
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .WrapText = True
            .RowHeight = 100
            With .Font
                .Bold = True
                .Size = HeadingSize
            End With
        End With
        
        With Range(NumberColumn & FirstRow, NumberColumn & z).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        
        With Range(TitleColumn & FirstRow, TitleColumn & z).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        
       
        With Range(MandColumn & FirstRow, MandColumn & z).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        
        With Range(OKColumn & FirstRow, OKColumn & z).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        
        With Range(NOKColumn & FirstRow, NOKColumn & z).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        
        With Range(NotDoneColumn & FirstRow, NotDoneColumn & z).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    
        
        With Range(EUColumn & FirstRow, EUColumn & z).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        
        With Range(NAColumn & FirstRow, NAColumn & z).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        
        With Range(APColumn & FirstRow, APColumn & z).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
         
        With Range(RelDocColumn & FirstRow, RelDocColumn & z).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    End With


End Function
Public Function UpdateDQualityDS(Worksheet As Worksheet)

Dim z As Integer
Dim Subtitle As String
Dim DataRelCategory() As Variant

z = FirstRow
Heading(1) = "Design Guideline"
Heading(2) = "Component"
Heading(3) = "Design Element"
Heading(4) = "Functional assembly"
Heading(5) = "Material Specification"
Heading(6) = "Drawing Template"
 
Worksheet.Activate

    Set RangeSubtitle = Worksheet.Range(TitleColumn & z, TitleColumn & z)
    Subtitle = Heading(1)
    Call EditSubtitle(RangeSubtitle, Subtitle)
    DataRelCategory = DataRELEASEDcat1DS
    Call FillStandards(z, DataRelCategory, Worksheet)
    Call FormatStandardRows(z, DataRelCategory)
    
    z = z + UBound(DataRelCategory, 1)
    DataRelCategory = DataRELEASEDcat2DS
    Set RangeSubtitle = Worksheet.Range(TitleColumn & z, TitleColumn & z)
    Subtitle = Heading(2)
    Call EditSubtitle(RangeSubtitle, Subtitle)
    Call FillStandards(z, DataRelCategory, Worksheet)
    Call FormatStandardRows(z, DataRelCategory)
     
    z = z + UBound(DataRelCategory, 1)
    DataRelCategory = DataRELEASEDcat3DS
    Set RangeSubtitle = Worksheet.Range(TitleColumn & z, TitleColumn & z)
    Subtitle = Heading(3)
    Call EditSubtitle(RangeSubtitle, Subtitle)
    Call FillStandards(z, DataRelCategory, Worksheet)
    Call FormatStandardRows(z, DataRelCategory)
    
    z = z + UBound(DataRelCategory, 1)
    DataRelCategory = DataRELEASEDcat4DS
    Set RangeSubtitle = Worksheet.Range(TitleColumn & z, TitleColumn & z)
    Subtitle = Heading(4)
    Call EditSubtitle(RangeSubtitle, Subtitle)
    Call FillStandards(z, DataRelCategory, Worksheet)
    Call FormatStandardRows(z, DataRelCategory)
    
    z = z + UBound(DataRelCategory, 1)
    DataRelCategory = DataRELEASEDcat5DS
    Set RangeSubtitle = Worksheet.Range(TitleColumn & z, TitleColumn & z)
    Subtitle = Heading(5)
    Call EditSubtitle(RangeSubtitle, Subtitle)
    Call FillStandards(z, DataRelCategory, Worksheet)
    Call FormatStandardRows(z, DataRelCategory)
    
    z = z + UBound(DataRelCategory, 1)
    DataRelCategory = DataRELEASEDcat6DS
    Set RangeSubtitle = Worksheet.Range(TitleColumn & z, TitleColumn & z)
    Subtitle = Heading(6)
    Call EditSubtitle(RangeSubtitle, Subtitle)
    Call FillStandards(z, DataRelCategory, Worksheet)
    Call FormatStandardRows(z, DataRelCategory)
    
    z = z + UBound(DataRelCategory, 1)
       
    Call FormatDQualitySheet(Worksheet, z)


End Function
Public Function EditSubtitle(Rango As Range, Subtitle As String)
    
   Dim numberLastColumn As Integer
   Dim numberFirstColumn As Integer
  
   numberFirstColumn = 1
   numberLastColumn = 11

    Rango = Subtitle
    With Rango
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .WrapText = True
            .RowHeight = 100
            With .Font
                .Name = "Arial"
                .Size = HeadingSize
                .Bold = True
                .Italic = True
            End With
            With .Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
    End With

    
End Function
Public Function FillStandards(z As Integer, DataRelCat() As Variant, Worksheet As Worksheet)

    Dim LineNumberTable() As Variant
    Dim i As Integer
    i = 1
    If (Worksheet.Name = "1.2.1 Standards respected DS" And z = 7) Then LineNumber = 1 'otherwise the ds sheets starts with the last number from ft
    
    ReDim LineNumberTable(UBound(DataRelCat, 1), 1)
    
    For i = 1 To UBound(DataRelCat, 1)
        LineNumberTable(i, 1) = LineNumber
        LineNumber = LineNumber + 1
    Next i
            
    
    Worksheet.Range(NumberColumn & z, NumberColumn & z + UBound(DataRelCat, 1) - 1).Clear
    z = z + 1
    Worksheet.Range(NumberColumn & z, NumberColumn & z + UBound(DataRelCat, 1) - 1) = OneColumn(LineNumberTable, 1)
    
    Worksheet.Range(TitleColumn & z, TitleColumn & z + UBound(DataRelCat, 1) - 1) = OneColumn(DataRelCat, 3)
    Worksheet.Range(RelDocColumn & z, RelDocColumn & z + UBound(DataRelCat, 1) - 1) = OneColumn(DataRelCat, 1)
    Worksheet.Range(EUColumn & z, EUColumn & z + UBound(DataRelCat, 1) - 1) = OneColumn(DataRelCat, 4)
    Worksheet.Range(NAColumn & z, NAColumn & z + UBound(DataRelCat, 1) - 1) = OneColumn(DataRelCat, 5)
    Worksheet.Range(APColumn & z, APColumn & z + UBound(DataRelCat, 1) - 1) = OneColumn(DataRelCat, 6)

End Function

Public Function FormatStandardRows(z As Integer, DataRelCat() As Variant)
    Dim cel As Range
    With Range(NumberColumn & z, LastColumn & z + UBound(DataRelCat, 1) - 1).Font
            .Name = "Arial"
            .Size = TextSize
            .Bold = False
        End With
        
        With Range(NumberColumn & z, LastColumn & z + UBound(DataRelCat, 1) - 1)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .RowHeight = 100
        End With
        
     
    For Each cel In Range(NumberColumn & z, LastColumn & z + UBound(DataRelCat, 1)).Cells
        With cel.Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    Next cel
       
End Function

Public Function DeleteEmptyRows(Table2Fix As Variant) As Variant
    Dim i, FirstEmptyRow As Integer
    Dim Columnsize, Rowsize As Long
    
    Rowsize = UBound(Table2Fix)
    Columnsize = UBound(Table2Fix, 2)
    
    Do While i < Rowsize
        i = i + 1
        If Table2Fix(i, 1) = Empty Then
            FirstEmptyRow = i
            Exit Do
        End If
    Loop
    
    Table2Fix = Application.Transpose(Table2Fix)
    ReDim Preserve Table2Fix(Columnsize, FirstEmptyRow - 1)
    DeleteEmptyRows = Application.Transpose(Table2Fix)
    
End Function
