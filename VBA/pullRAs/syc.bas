Option Private Module
Sub RA_SYC()
       
    ' Filter for location, MFG Warr, brand, and status
    ActiveSheet.Range("$A$1:$AT$1").AutoFilter Field:=1, Criteria1:= _
        "1320"
    ActiveSheet.Range("$A$1:$AT$1").AutoFilter Field:=15, Criteria1:= _
        "MFG Warranty"
    ActiveSheet.Range("$A$1:$AT$1").AutoFilter Field:=16, Criteria1:= _
        "MFG Warranty"
    ActiveSheet.Range("$A$1:$AT$1").AutoFilter Field:=22, Criteria1:= _
        "SYC"
    ActiveSheet.Range("$A$1:$AT$1").AutoFilter Field:=29, Criteria1:= _
        "NATIONAL PARTS INC"
    ActiveSheet.Range("$A$1:$AT$1").AutoFilter Field:=31, Criteria1:= _
        "<>Shipped", Operator:=xlAnd
        
    SetDaysForm.Show vbModal 'Show form to get days to filter for
    
End Sub
Sub format_SYC()

    ' Format for Sony
    Sheets("Sheet1").Range("A:G,I:I,K:K,N:Q,T:V,X:Y,AB:AD,AG:AG,AI:AL,AO:AT").EntireColumn.Delete
    
    ' Set format for serial number column
    ActiveSheet.Range("N:N").NumberFormat = "0"
    
    ' Remove dupes
    ActiveSheet.Range("$A:$N").RemoveDuplicates Columns:=4, Header:= _
        xlYes
        
    ' Sort based on SKU
    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Sort.SortFields.Add Key:=Range("A:A"), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveSheet.Sort
        .SetRange Range("A:O")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Dim zSKUStart As Integer, zSKUEnd As Integer, skuStart As Integer, skuEnd As Integer
    Dim rngC As Range
    Dim foundZSKU As Boolean
    Dim foundSKU As Boolean
        
    For Each rngC In Intersect(ActiveSheet.UsedRange.Offset(1, 0), ActiveSheet.Range("A:A").SpecialCells(xlCellTypeVisible))
        If rngC.Value Like "a*" Then foundZSKU = True
        If rngC.Value Like "#*" Then foundSKU = True
    Next rngC
    
    ' Find first and last rows the 0 SKUs
    With ActiveSheet
        If foundZSKU = True Then
            zSKUStart = .Range("A:A").Find(what:="a*", after:=.Range("A1")).Row
            zSKUEnd = .Range("A:A").Find(what:="a*", after:=.Range("A1"), searchdirection:=xlPrevious).Row
        End If
        If foundSKU = True Then
            skuStart = Range("A2").Row
            If foundZSKU = True Then
                skuEnd = zSKUStart - 1
            Else
                skuEnd = .Range("A:A").Find(what:="*", after:=.Range("A1"), searchdirection:=xlPrevious).Row
            End If
        End If
    End With
    

    If foundZSKU = True Then
        ' Move 0 SKUS down to create empty row between them and non-0 SKUs
        Rows(zSKUStart).Insert shift:=xlShiftDown
        zSKUEnd = zSKUEnd + 1 'Add 1 to zSKUEnd to account for move
        ActiveSheet.Range("L" & zSKUStart).Value = "Age" 'Add header for 0 SKUs
            
        ' Sort 0 SKUs by age
        ActiveSheet.Sort.SortFields.Clear
        ActiveSheet.Sort.SortFields.Add Key:=Range("L" & zSKUStart & ":L" & zSKUEnd), _
            SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        With ActiveSheet.Sort
            .SetRange Range("A" & zSKUStart & ":N" & zSKUEnd)
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End If

    If foundSKU = True Then
        ' Sort non-0 SKUs by age
        ActiveSheet.Sort.SortFields.Clear
        ActiveSheet.Sort.SortFields.Add Key:=Range("L" & skuStart & ":L" & skuEnd), _
            SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        With ActiveSheet.Sort
            .SetRange Range("A1" & ":N" & skuEnd)
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End If

End Sub
