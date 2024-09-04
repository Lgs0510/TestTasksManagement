Attribute VB_Name = "UpdateOldCvsWithNewCvs"
Sub UpdateOldCvWithNewCv()
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    For Each Worksheet In ActiveWorkbook.Worksheets
        If InStr(Worksheet.Name, "CV-") Then
            Worksheet.Activate
            numberOfRows = lastRowNumber
            For iRow = 2 To numberOfRows
                Range(TESTCASES_NewCvCL + CStr(iRow)).Select
                If ActiveCell.value <> "" Then
                    If InStr(ActiveCell.value, "CV-") Then
                        Range(TESTCASES_TestCvCL + CStr(iRow)).value = Range(TESTCASES_NewCvCL + CStr(iRow)).value
                    End If
                End If
            Next
        End If
    Next
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
