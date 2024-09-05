Private Sub Worksheet_Activate()
    If Not g_vbaIsRunning Then
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        ActiveSheet.Protect _
            Password:=sheetsProtectionPassword, _
            AllowFiltering:=True, _
            AllowSorting:=True
    End If
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    g_vbaIsRunning = False
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
    g_vbaIsRunning = True

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
    Application.Calculation = calcPrevStatus
    Application.ScreenUpdating = screenPrevUpdate
    ProtectSheet (protectionStatus)
    Application.EnableEvents = True
    g_vbaIsRunning = False
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
