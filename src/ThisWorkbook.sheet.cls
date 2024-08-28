Private Sub Workbook_Activate()
    Call checkBadClosure
    Call sbCopyFile
End Sub

Private Sub Workbook_Deactivate()
     Call sbDeleteFile
     Call restoreOldUserConfig
End Sub