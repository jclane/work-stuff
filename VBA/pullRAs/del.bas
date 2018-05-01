Option Private Module
Sub RA_DEL()
       
    'Filter for location, MFG Warr, brand, and status
    ActiveSheet.Range("$A$1:$AT$1").AutoFilter Field:=1, Criteria1:= _
        "<>1320"
    ActiveSheet.Range("$A$1:$AT$1").AutoFilter Field:=15, Criteria1:= _
        "MFG Warranty"
    ActiveSheet.Range("$A$1:$AT$1").AutoFilter Field:=16, Criteria1:= _
        "MFG Warranty"
    ActiveSheet.Range("$A$1:$AT$1").AutoFilter Field:=29, Criteria1:= _
        "=ALIENWARE CORP", Operator:=xlOr, Criteria2:="=DELL DIRECT SALES LP"
    ActiveSheet.Range("$A$1:$AT$1").AutoFilter Field:=31, Criteria1:= _
        "<>Shipped", Operator:=xlAnd
    
    SetDaysForm.Show vbModal 'Show form to get days to filter for
    
End Sub
Sub format_DEL()

    ' Format for Dell
    Sheets("Sheet1").Range("B:V,X:Y,AB:AE,AG:AG,AI:AL,AO:AU").EntireColumn.Delete
        
    'Remove dupes
    ActiveSheet.Range("$A:$H").RemoveDuplicates Columns:=8, Header:=xlYes
    
End Sub
