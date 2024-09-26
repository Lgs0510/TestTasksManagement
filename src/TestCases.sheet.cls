Private Sub Worksheet_Activate()

    ActiveWorkbook.Sheets("TestCases").ProgressBarTestCases.Visible = False
    ActiveWorkbook.Sheets("TestCases").ProgressBar_Label.Visible = False
    ActiveWorkbook.Sheets("TestCases").ProgressBar_percentage.Visible = False
    
    If Not g_vbaIsRunning Then
        GenericFunctions.uiEnable(True)
        GenericFunctions.ProtectSheet (True)
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
    
    calcPrevStatus = Application.Calculation
    GenericFunctions.uiDisable
    GenericFunctions.UnprotectSheet
    For Each cell In Target
        If cell <> "" Then
            Cells(cell.Row, 5).value = Date
        End If
    Next
    GenericFunctions.ProtectSheet (protectionStatus)
    GenericFunctions.uiEnable (calcPrevStatus)
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
