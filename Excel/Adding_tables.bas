Attribute VB_Name = "Adding_tables"
Option Base 1

Public Function addTables(Tables1() As Variant, Tables2() As Variant) As Variant()
    
    Dim a As Long
    Dim b As Long
    Dim Dim11 As Long
    Dim Dim12 As Long
    Dim Dim21 As Long
    Dim Dim22 As Long
    Dim Tables3() As Variant
    
    Dim11 = UBound(Tables1(), 1)
    Dim12 = UBound(Tables1(), 2)
    Dim21 = UBound(Tables2(), 1)
    Dim22 = UBound(Tables2(), 2)
    
    ReDim Preserve Tables3(Dim11 + Dim21, Dim12)
    
    'addTables = Tables1
            
        For a = 1 To Dim11
            For b = 1 To Dim12
                
                Tables3(a, b) = Tables1(a, b)
                
            Next b
        Next a
        
        For a = 1 To Dim21
            For b = 1 To Dim22
                
                Tables3(a + Dim11, b) = Tables2(a, b)
                
            Next b
        Next a
        
       addTables = Tables3
        
End Function

Public Function addTables2(ParamArray Table() As Variant) As Variant()
    
    Dim c As Long
    Dim d As Long
    Dim e As Long
    Dim ConstDim As Long
   Dim Leng As Integer

   Leng = 10
   Dim VariableDim() As Long
   ReDim VariableDim(Leng)
   Dim Dim11() As Long
   ReDim Dim11(Leng)
   Dim Dim12() As Long
   ReDim Dim12(Leng)

   Dim Tables() As Variant
   
        For c = 1 To Leng
            
            Dim11(c) = UBound(Table(c - 1), 1)
            Dim12(c) = UBound(Table(c - 1), 2)

        Next c
   
   VariableDim(1) = 0
   
   For c = 1 To Leng
   
        ConstDim = ConstDim + Dim11(c)
            If c > 1 Then
                VariableDim(c) = VariableDim(c - 1) + Dim11(c - 1)

            End If
   Next c
   
   ReDim Tables(ConstDim, Dim12(1))
   
   For c = 0 To Leng - 1
        For d = 1 To Dim11(c + 1)
            For e = 1 To Dim12(c + 1)
                    
                    Tables(VariableDim(c + 1) + d, e) = Table(c)(d, e)
                    
            Next e
        Next d
    Next c
    addTables2 = Tables
End Function

