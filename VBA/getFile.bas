Private Sub makeTheCall()
    Call Stock.Max_Order_Qty_10
End Sub
Private Sub gotOPReport()
    Dim strFileToOpen As String
    strFileToOpen = Application.GetOpenFilename _
    (Title:="Please choose a file to open", _
    FileFilter:="Excel Files *.xls* (*.xls*),")
    If strFileToOpen = "False" Then
        MsgBox "No file selected.", vbExclamation, "Sorry!"
        Exit Sub
    Else
        Workbooks.Open Filename:=strFileToOpen
        Call makeTheCall
    End If
End Sub
Private Sub getOPReport()

    Dim ask As String, ans As Variant, gotPO As Boolean
    
    ' Ask if user wants to import report
    ask = "Do you wish to save the the OpenPO Report to your desktop and pull the RAs now?" _

    ans = MsgBox(ask, vbYesNo)
    
    Select Case ans
        Case vbYes 'Import and run
        
        
            Dim Path As String, Share As String, fn As String
            
            ' Set "Path" to user's desktop
            Path = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"
            
            ' Set "Share" to path to Planning report
            Share = "!!REDACTED!!\Stock Replenishment Rpts\"
            
            ' Set "fn" to planning report file name
            fn = Dir(Share & "Parts Planning *-GSC.xlsm")
            
            ' If file found copy to desktop, open and run Max_Order_Qty_10
            If Len(fn) > 0 Then
                FileCopy Share & fn, Path & "Parts Planning.xlsm"
                Workbooks.Open Path & "Parts Planning.xlsm"
                Call makeTheCall
            End If
        Case vbNo ' Do nothin
    End Select
        
End Sub
