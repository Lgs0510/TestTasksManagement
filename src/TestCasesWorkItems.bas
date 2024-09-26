Attribute VB_Name = "TestCasesWorkItems"
Sub UpdateOldCvWithNewCv()
    Dim statusToDelete As New list
    Dim calcPrevStatus As XlCalculation
    
    calcPrevStatus = Application.Calculation
    GenericFunctions.uiDisable
    statusToDelete.letList = getTestCasesStatusToRemove
    For Each Worksheet In ActiveWorkbook.Worksheets
        If InStr(Worksheet.Name, "CV-") Then
            Worksheet.Activate
            numberOfRows = lastRowNumber
            For iRow = 2 To numberOfRows
                Range(CVs_SHEETS_NewCvCL + CStr(iRow)).Select
                If ActiveCell.value <> "" Then
                    If InStr(ActiveCell.value, "CV-") Then
                        Range(CVs_SHEETS_TestCvCL + CStr(iRow)).value = Range(CVs_SHEETS_NewCvCL + CStr(iRow)).value
                    End If
                ElseIf statusToDelete.Contains(Range(CVs_SHEETS_StatusCL + CStr(iRow)).value) Then
                    Range(CVs_SHEETS_StatusCL + CStr(iRow)).EntireRow.Delete
                    iRow = iRow - 1
                End If
            Next
        End If
    Next

    GenericFunctions.uiEnable(calcPrevStatus)
End Sub


Function getTestCasesStatusToRemove() As String()
    Dim statusList As New list
    For Each Status In Split(testCaseStatusToDELETE, ",")
        statusList.Add (Replace(Status, " ", ""))
    Next
    getTestCasesStatusToRemove = statusList.getList
End Function

