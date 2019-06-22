VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Copy_of_2014()
'

    
    Sheets("Copy_of_2014").Columns(1).Copy Destination:=Sheets("Copy_of_2014").Columns(9)
   
    ActiveSheet.Range("$I$2:$I$705714").RemoveDuplicates Columns:=1, Header:= _
        xlNo
 
  
    
    For i = 2 To 2836
        Worksheets("Copy_of_2014").Range("J" & i).Formula = "=SUMIF(C[-9],RC[-1],C[-3])"

      
    Next i
    
   
End Sub


