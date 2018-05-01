Option Private Module
Sub RA_LNV()
    
    'Filter for location, MFG warr, brand, and status
    ActiveSheet.Range("$A$1:$AT$1").AutoFilter Field:=1, Criteria1:= _
        "<>1320"
    ActiveSheet.Range("$A$1:$AT$1").AutoFilter Field:=15, Criteria1:= _
        "MFG Warranty"
    ActiveSheet.Range("$A$1:$AT$1").AutoFilter Field:=16, Criteria1:= _
        "MFG Warranty"
    ActiveSheet.Range("$A$1:$AT$1").AutoFilter Field:=29, Criteria1:= _
        "LENOVO UNITED STATES"
    ActiveSheet.Range("$A$1:$AT$1").AutoFilter Field:=31, Criteria1:= _
        "<>Shipped", Operator:=xlAnd
    
    SetDaysForm.Show vbModal 'Show form to get days to filter for

End Sub
Sub format_LNV()

        'Format for Lenovo
        Sheets("Sheet1").Range("B:I,K:Y,AA:AE,AG:AG,AI:AL,AO:AU").EntireColumn.Delete
        
        'Remove dupes
        ActiveSheet.Range("$A:$AT").RemoveDuplicates Columns:=6, Header:= _
            xlYes
            
End Sub
