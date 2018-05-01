Private Sub makeTheCall()
    Call MAIN.getRAs
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
        
            
                Dim Path As String
                ' Set "Path" to user's desktop
                Path = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"
                FileCopy "!!REDACTED!!\rsCorpOpenPO.xlsx", Path & "openPO " & Date$ & ".xlsx"
                Workbooks.Open Path & "openPO " & Date$ & ".xlsx"
                Call makeTheCall
        

        Case vbNo ' Do nothin if user says NO
    End Select
        
End Sub
