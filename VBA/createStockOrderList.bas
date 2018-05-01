Public PlanBook As Workbook
Public StockBook As Workbook

Sub superslow(Optional dummy_var As Integer)
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    'Application.Calculation = xlCalculationAutomatic
    ActiveSheet.DisplayPageBreaks = True
End Sub

Sub superfast(Optional dummy_var As Integer)
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    'Application.Calculation = xlCalculationManual
    ActiveSheet.DisplayPageBreaks = False
End Sub
Private Sub pullGPC()
    ' Apply autofilter and check for results
    ActiveSheet.Range("A5:AD5").AutoFilter Field:=23, Criteria1:=">=1", Operator:=xlFilterValues
    
    Dim lastRow As Integer
    
    If ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row > 0 Then
        lastRow = ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
    End If
    
    If lastRow > 5 Then
        ActiveSheet.AutoFilter.Range.Copy
        StockBook.Sheets("GPC").Activate
        Worksheets("GPC").Paste
        
        ' Format
        Sheets("GPC").Range("C:T,V:AB").EntireColumn.Delete
        Sheets("GPC").Range("A1:C1").Delete
        If lastRow = 0 Then
            Range("A1").Value = "NONE"
        Else
            ' Max_Order_Qty_10 for the GPC parts
            Dim RowToTest As Long
            Dim last_row As Integer
            last_row = WorksheetFunction.CountA(Worksheets("GPC").Range("C:C"))
            
            For i = last_row To 1 Step -1
                    If Worksheets("GPC").Columns(3).Rows(i).Value < 1 Then
                        Rows(i).EntireRow.Delete
                    End If
            Next i
            
            Dim original_QTY_Ary As Variant
            Dim last_row_order As Integer
            
            last_row = WorksheetFunction.CountA(Worksheets("GPC").Range("C:C"))
            original_QTY_Ary = Worksheets("GPC").Range("A1:C" & last_row)
            last_row_order = WorksheetFunction.CountA(Worksheets("GPC").Range("A:A"))
            
            For i = 1 To (last_row)
                If original_QTY_Ary(i, 3) < 10 Then
                    last_row_order = WorksheetFunction.CountA(Worksheets("GPC").Range("A:A"))
                    Worksheets("GPC").Range("A" & (last_row_order + 1)).Value = original_QTY_Ary(i, 1)
                    Worksheets("GPC").Range("B" & (last_row_order + 1)).Value = original_QTY_Ary(i, 2)
                    Worksheets("GPC").Range("C" & (last_row_order + 1)).Value = original_QTY_Ary(i, 3)
                ElseIf original_QTY_Ary(i, 3) Mod 10 <> 0 Then
                    For j = 0 To Int((original_QTY_Ary(i, 3) / 10))
                        last_row_order = WorksheetFunction.CountA(Worksheets("GPC").Range("A:A"))
                        Worksheets("GPC").Range("A" & (last_row_order + 1)).Value = original_QTY_Ary(i, 1)
                        Worksheets("GPC").Range("B" & (last_row_order + 1)).Value = original_QTY_Ary(i, 2)
                        Worksheets("GPC").Range("C" & (last_row_order + 1)).Value = 10
                    Next
                    Worksheets("GPC").Range("A" & (last_row_order + 1)).Value = original_QTY_Ary(i, 1)
                    Worksheets("GPC").Range("B" & (last_row_order + 1)).Value = original_QTY_Ary(i, 2)
                    Worksheets("GPC").Range("C" & (last_row_order + 1)).Value = original_QTY_Ary(i, 3) Mod 10
                ElseIf original_QTY_Ary(i, 3) Mod 10 = 0 Then
                    For j = 1 To (original_QTY_Ary(i, 3) / 10)
                        last_row_order = WorksheetFunction.CountA(Worksheets("GPC").Range("A:A"))
                        Worksheets("GPC").Range("A" & (last_row_order + 1)).Value = original_QTY_Ary(i, 1)
                        Worksheets("GPC").Range("B" & (last_row_order + 1)).Value = original_QTY_Ary(i, 2)
                        Worksheets("GPC").Range("C" & (last_row_order + 1)).Value = 10
                    Next
                End If
            Next
            
            Range("A1:C" & last_row).Rows.Delete
        End If
    Else
        StockBook.Sheets("GPC").Activate
        Range("A1").Value = "NONE"
    End If
End Sub
Private Sub saveIt()
    Dim Path As String
    Dim Filename As String
    
    ' Order worksheets because having them out of order bothers me
    Sheets("Sheet2").Move after:=Sheets("Sheet1")
    Sheets("Sheet3").Move after:=Sheets("Sheet2")
    StockBook.Sheets("Sheet2").Activate
    
    ' Add comment to file properties so I know if stock order file was created with this macro
    StockBook.BuiltinDocumentProperties("Comments").Value = "Created with macro."
    
    ' Set "Path" to user's desktop
    Path = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"
    
    ' Set "Filename"
    Filename = "Stock " & Date$ & ".xlsx"
            
    ' Save it
    ActiveWorkbook.SaveAs Path & Filename
    MsgBox ("File " & Filename & " has been saved to your desktop.")
    
End Sub

Sub Max_Order_Qty_10()
    
    Call superfast
    
    ' Set Planning report as PlanBook
    Set PlanBook = ActiveWorkbook
    
    ' Set sheet with relevant data as active
    Sheets("OrderConsolidation(ExcludesGPC)").Activate
    
    ' Copy and paste data to new workbook
    ActiveSheet.UsedRange.Offset(4).Copy
    Workbooks.Add.Worksheets(1).Paste
    
    ' Set Stock order list as StockBook
    Set StockBook = ActiveWorkbook

    ' Add 3 sheets and name last "GPC"
    Worksheets.Add after:=Worksheets(Worksheets.Count), Count:=3
    Sheets("Sheet4").Name = "GPC"
    
    ' Go back to Planning report and copy GPC data
    PlanBook.Activate
    Sheets("GPC").Activate
    
    Call pullGPC
    
    Dim original_QTY_Ary As Variant
    Dim last_row As Integer
    Dim last_row_order As Integer
        
    last_row = WorksheetFunction.CountA(Worksheets("Sheet1").Range("E:E"))
    original_QTY_Ary = Worksheets("Sheet1").Range("A2:E" & last_row)
    last_row_order = WorksheetFunction.CountA(Worksheets("Sheet2").Range("A:A"))
    Worksheets("Sheet2").Range("A:E").Clear
    Worksheets("Sheet2").Range("A1").Value = "Part Num"
    Worksheets("Sheet2").Range("B1").Value = "QTY"
    For i = 1 To (last_row - 1)
        If original_QTY_Ary(i, 5) < 10 Then
            last_row_order = WorksheetFunction.CountA(Worksheets("Sheet2").Range("A:A"))
            Worksheets("Sheet2").Range("A" & (last_row_order + 1)).Value = original_QTY_Ary(i, 1)
            Worksheets("Sheet2").Range("B" & (last_row_order + 1)).Value = original_QTY_Ary(i, 5)
        ElseIf original_QTY_Ary(i, 5) Mod 10 <> 0 Then
            For j = 0 To Int((original_QTY_Ary(i, 5) / 10))
                last_row_order = WorksheetFunction.CountA(Worksheets("Sheet2").Range("A:A"))
                Worksheets("Sheet2").Range("A" & (last_row_order + 1)).Value = original_QTY_Ary(i, 1)
                Worksheets("Sheet2").Range("B" & (last_row_order + 1)).Value = 10
            Next
                Worksheets("Sheet2").Range("A" & (last_row_order + 1)).Value = original_QTY_Ary(i, 1)
                Worksheets("Sheet2").Range("B" & (last_row_order + 1)).Value = original_QTY_Ary(i, 5) Mod 10
        ElseIf original_QTY_Ary(i, 5) Mod 10 = 0 Then
            For j = 1 To (original_QTY_Ary(i, 5) / 10)
                last_row_order = WorksheetFunction.CountA(Worksheets("Sheet2").Range("A:A"))
                Worksheets("Sheet2").Range("A" & (last_row_order + 1)).Value = original_QTY_Ary(i, 1)
                Worksheets("Sheet2").Range("B" & (last_row_order + 1)).Value = 10
            Next
        End If
    Next
    Worksheets("Sheet3").Range("A:J").Clear
    Worksheets("Sheet3").Range("A1").Value = "Part Num"
    Worksheets("Sheet3").Range("B1").Value = "Original QTY"
    Worksheets("Sheet3").Range("C1").Value = "Order QTY"
    For i = 1 To UBound(original_QTY_Ary, 1)
        last_row_order = WorksheetFunction.CountA(Worksheets("Sheet3").Range("A:A"))
        Worksheets("Sheet3").Range("A" & (last_row_order + 1)).Value = original_QTY_Ary(i, 1)
        Worksheets("Sheet3").Range("B" & (last_row_order + 1)).Value = original_QTY_Ary(i, 5)
        Worksheets("Sheet3").Range("C" & (last_row_order + 1)).Value = "=SUMIF(Sheet2!A:A,A" & i + 1 & ",Sheet2!B:B)"
    Next
    For i = 2 To WorksheetFunction.CountA(Worksheets("Sheet3").Range("A:A"))
        If Worksheets("Sheet3").Range("B" & i).Value = Worksheets("Sheet3").Range("C" & i).Value Then
            Worksheets("Sheet3").Range("A" & i).Interior.Color = (RGB(0, 128, 0))
        End If
    Next
    
    last_row_order = WorksheetFunction.CountA(Worksheets("Sheet3").Range("A:A"))
    For i = 2 To last_row_order
        If Worksheets("Sheet3").Range("B" & i).Value = Worksheets("Sheet3").Range("c" & i).Value Then
            Worksheets("Sheet3").Range("A" & i).Interior.Color = RGB(0, 128, 0)
        End If
    Next
    
    Worksheets("Sheet3").Range("G2").Value = "Checks"
    Worksheets("Sheet3").Range("G3").Value = "Original count of part numbers"
    Worksheets("Sheet3").Range("G4").Value = "Order count of part numbers"
    Worksheets("Sheet3").Range("G6").Value = "Original Sum of part orders"
    Worksheets("Sheet3").Range("G7").Value = "Post processing count of part orders"
    Worksheets("Sheet3").Range("H3").Value = "=COUNTA(Sheet1!A:A)"
    Worksheets("Sheet3").Range("H4").Value = "=COUNTA(Sheet3!A:A)"
    Worksheets("Sheet3").Range("H6").Value = "=SUM(Sheet1!E:E)"
    Worksheets("Sheet3").Range("H7").Value = "=SUM(Sheet2!B:B)"
    
    
    ' Sort "Sheet2" by quantity (high to low)
    StockBook.Sheets("Sheet2").Activate
    
    ActiveWorkbook.Worksheets("Sheet2").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet2").Sort.SortFields.Add Key:=Range("B1"), _
    SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet2").Sort
        .SetRange Range("A:B")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Call saveIt
    
Call superslow
End Sub
