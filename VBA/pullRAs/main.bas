Public MainWkBk As Workbook
Public btnCancel As Boolean, fileSaved As Boolean
Public days As Long
Public addDays As Long
'Public addDaysCheck As Boolean
Public currBrand As String
Function IsWorkBookOpen(Name As String) As Boolean
    Dim xWb As Workbook
    On Error Resume Next
    Set xWb = Application.Workbooks.item(Name)
    IsWorkBookOpen = (Not xWb Is Nothing)
End Function
Private Sub saveIt()
    Dim Path As String
    Dim Filename As String
    
    ' Set "Path" to user's desktop
    Path = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"
    
    ' Set "Filename" to correct format based on brand
    Select Case currBrand
        Case "Dell"
            Filename = "Dell non-GSC " & Date$ & ".xlsx"
        Case "Lenovo"
            Filename = "Lenovo non-GSC RA " & Date$ & ".xlsx"
        Case "Toshiba"
            Filename = "Toshiba 7+ Days " & Date$ & ".xlsx"
        Case "Sony"
            Filename = "Sony 0 SKUs " & Date$ & ".xlsx"
    End Select
        
    ' Save it
    Dim book As Boolean
    book = IsWorkBookOpen(Filename)
    If book = False Then
        ActiveWorkbook.SaveAs Path & Filename
        MsgBox ("File " & Filename & " has been saved to your desktop.")
        fileSaved = True
    Else
        MsgBox "ALERT!"
    End If
    
End Sub
Private Sub noRAToday()
        ActiveSheet.AutoFilterMode = False
        MsgBox ("There are no " & currBrand & " RAs for today.")
End Sub
Private Sub deleteFile()
      
    Dim ask As String, ans As Variant
    
    ' Ask if user wants delete the file
    ask = "Would you like to delete the Open PO report from your desktop?"

    ans = MsgBox(ask, vbYesNo)
    
    Select Case ans
        Case vbYes 'Close and delete
            Workbooks("openPO " & Date$).Close SaveChanges:=False
            Dim Path As String
            ' Set "Path" to user's desktop
            Path = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"
            
            If Len(Dir(Path & "openPO " & Date$ & ".xlsx")) > 0 Then Kill Path & "openPO " & Date$ & ".xlsx"
        Case vbNo ' Do nothin
    End Select
        
End Sub

Private Sub copyData()
        ActiveSheet.AutoFilter.Range.Copy
        Workbooks.Add.Worksheets(1).Paste
End Sub
Private Sub bigDaddyLoop()
    
    If btnCancel = False Then
        
        Dim daysArr() As String
        ageFound = False
        
        If currBrand = "Sony" Then
            For Each rngC In Intersect(ActiveSheet.UsedRange.Offset(1, 0), ActiveSheet.Range("AH:AH").SpecialCells(xlCellTypeVisible))
                If rngC.Value >= days Then
                    ageFound = True
                    Exit For
                End If
            Next rngC
            ' Filter for 10+ days
            ActiveSheet.Range("$A$1:$AU$1").AutoFilter Field:=34, Criteria1:=">=" & days, Operator:=xlFilterValues
        Else
            ' Create an array of numbers to filter the Age column
            Dim i As Long
            ReDim daysArr(0 To addDays) As String
            
            For i = LBound(daysArr) To UBound(daysArr)
                daysArr(i) = days + i
            Next i
            
            Dim c As Long
            
            ' Filter for days in array
            ActiveSheet.Range("$A$1:$AU$1").AutoFilter Field:=34, Criteria1:=daysArr, Operator:=xlFilterValues
            
            ' Check if age is present
            For Each rngC In Intersect(ActiveSheet.UsedRange, ActiveSheet.Range("AH:AH").SpecialCells(xlCellTypeVisible))
                If ageFound = True Then
                    Exit For
                Else
                    For c = LBound(daysArr) To UBound(daysArr)
                        If rngC.Value = daysArr(c) Then
                            ageFound = True
                            Exit For
                        End If
                    Next c
                End If
            Next rngC
        
        End If
        

            
        ' If age was found then complete the list of RAs for the vendor
        If ageFound = True Then
             
            ' Copy results, create new sheet, paste
            Call copyData
                
            ' Call correct format macro
            Select Case currBrand
                Case "Dell"
                    Call format_DEL
                Case "Lenovo"
                    Call format_LNV
                Case "Toshiba"
                    Call format_TSC
                Case "Sony"
                    Call format_SYC
            End Select
                
            ' Save the file to the desktop
            Call saveIt
        Else
            Call noRAToday
        End If
    End If

    days = 0
    MainWkBk.Activate
    ActiveSheet.AutoFilterMode = False
    
End Sub
Sub getRAs()

    Set MainWkBk = ActiveWorkbook

    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual
    ActiveSheet.DisplayPageBreaks = False

    currBrand = "Dell"
    Call RA_DEL
    Call bigDaddyLoop
    
    currBrand = "Lenovo"
    Call RA_LNV
    Call bigDaddyLoop
    
    currBrand = "Toshiba"
    Call RA_TSC
    Call bigDaddyLoop

    currBrand = "Sony"
    Call RA_SYC
    Call bigDaddyLoop
    
    MainWkBk.Activate
    
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.Calculation = xlCalculationAutomatic
    ActiveSheet.DisplayPageBreaks = True
    
    If fileSaved = True Then
        MsgBox ("The RAs have been pulled and are available on your desktop.")
    Else
        MsgBox ("There were no RAs today.")
    End If
    
    Call deleteFile
    
End Sub
