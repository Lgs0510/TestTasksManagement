Private Sub Workbook_Activate()
    Call checkBadClosure
    Call sbCopyFile
    ThisWorkbook.Sheets("Trace").Activate
    ThisWorkbook.Sheets("TestCases").Activate
End Sub

Private Sub Workbook_Deactivate()
     Call sbDeleteFile
     Call restoreOldUserConfig
End Sub