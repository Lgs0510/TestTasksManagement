Private Sub Worksheet_Activate()
    If Not g_vbaIsRunning Then
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
    End If
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    g_vbaIsRunning = False
End Sub