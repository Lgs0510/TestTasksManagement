Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Target.Column = 2 Then
        Target = toggleTestCaseStatus(Target)
    Else
        TestCaseInsertion.Show
    End If
    Cancel = True
End Sub
Private Sub Worksheet_Change(ByVal Target As Range)
    Application.EnableEvents = False
    protectionStatus = ActiveSheet.ProtectContents

    For Each cell In Target
        If cell <> "" Then
            UnprotectSheet (protectionStatus)
            Cells(cell.Row, 5).value = Date
            ProtectSheet (protectionStatus)
        End If
    Next
    Cancel = True
    Application.EnableEvents = True
End Sub

Private Function toggleTestCaseStatus(curCell) As String
    Select Case curCell
        Case ""
            curCell = "OK"
        Case "OK"
            curCell = "NOK"
        Case Else
            curCell = ""
    End Select
    toggleTestCaseStatus = curCell
End Function