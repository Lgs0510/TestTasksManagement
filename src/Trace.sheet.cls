Private Sub Worksheet_Activate()
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    ActiveSheet.Protect _
        Password:=sheetsProtectionPassword, _
        AllowFiltering:=True, _
        AllowSorting:=True
End Sub
