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
    Dim calcPrevStatus As XlCalculation
    Dim screenPrevUpdate As Boolean

    calcPrevStatus = Application.Calculation
    screenPrevUpdate = Application.ScreenUpdating
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    UnprotectSheet (protectionStatus)
    For Each cell In Target
        If cell <> "" Then
            
            Cells(cell.Row, 5).value = Date
        End If
    Next
    Application.Calculation = Application.Calculation
    Application.ScreenUpdating = screenPrevUpdate
    ProtectSheet (protectionStatus)
    Application.EnableEvents = True
    Cancel = True
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