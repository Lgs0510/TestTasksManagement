Attribute VB_Name = "ProgressLoadBarModule"
Dim previousProgress As Integer

Public Sub ProgressLoad(ByVal curValue As Integer, ByVal maxValue As Integer, ByVal progressLabel As String)

    curProgress = CInt((curValue / maxValue) * 1000) / 10
    
    If curProgress > 100 Then
        curProgress = 100
    End If
    If (curProgress - previousProgress) > 1 Then
        previousProgress = curProgress
        Application.ScreenUpdating = True
        GenericFunctions.UnprotectSheet
        openProgressBar
        ActiveWorkbook.Sheets("TestCases").ProgressBar_Label.Caption = progressLabel
        ActiveWorkbook.Sheets("TestCases").ProgressBar_percentage.Text = 0
        
    
        ActiveWorkbook.Sheets("TestCases").ProgressBarLoad.value = curProgress
        ActiveWorkbook.Sheets("TestCases").ProgressBar_percentage.Text = CStr(curProgress) + "%"
        DoEvents
        Application.ScreenUpdating = False
    End If
End Sub

Public Sub closeProgressBar()
    previousProgress = 0
    ActiveWorkbook.Sheets("TestCases").ProgressBarLoad.Visible = False
    ActiveWorkbook.Sheets("TestCases").ProgressBar_Label.Visible = False
    ActiveWorkbook.Sheets("TestCases").ProgressBar_percentage.Visible = False
End Sub


Private Sub openProgressBar()
        ActiveWorkbook.Sheets("TestCases").ProgressBarLoad.Visible = True
        ActiveWorkbook.Sheets("TestCases").ProgressBar_Label.Visible = True
        ActiveWorkbook.Sheets("TestCases").ProgressBar_percentage.Visible = True
End Sub


