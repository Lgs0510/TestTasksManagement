
Private Sub Workbook_BeforeClose(Cancel As Boolean)
     Call sbDeleteFile
     Call restoreOldUserConfig
     
    ProgressLoadBarModule.hideProgressBar
    
End Sub

Private Sub Workbook_Deactivate()

End Sub

Private Sub Workbook_Open()
    Call checkBadClosure
    Call sbCopyFile
    g_vbaIsRunning = False
    ThisWorkbook.Sheets("Trace").Activate
    ThisWorkbook.Sheets("TestCases").Activate

    ProgressLoadBarModule.hideProgressBar
End Sub