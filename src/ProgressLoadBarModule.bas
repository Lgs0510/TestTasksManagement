Attribute VB_Name = "ProgressLoadBarModule"
Dim previousProgress As Double
Dim ProgressBarStatus As Boolean
Const milisecond As Double = 0.000000011574



Public Sub ProgressLoad(ByVal curValue As Integer, ByVal maxValue As Integer, ByVal progressLabel As String)

    curProgress = CInt((curValue / maxValue) * 1000) / 10
    
    If curProgress > 100 Then
        curProgress = 100
    End If
    If (curProgress - previousProgress) >= 0.1 Then
        previousProgress = curProgress
        DoEvents
        DoEvents
        DoEvents
    End If
    If Not ProgressBarStatus Then
        openProgressBar (progressLabel)
    End If
    ActiveWorkbook.Sheets("TestCases").ProgressBarLoad.value = previousProgress
    ActiveWorkbook.Sheets("TestCases").ProgressBar_percentage.Text = CStr(previousProgress) + "%"
End Sub
Public Sub closeProgressBar()

    previousProgress = 0
    hideProgressBar
    ProgressBarStatus = False
    ActiveWorkbook.AutoSaveOn = True
End Sub

Public Sub hideProgressBar()
    If ActiveWorkbook.Sheets("TestCases").ProgressBarLoad.Visible Then
        ActiveWorkbook.Sheets("TestCases").ProgressBarLoad.Visible = False
    End If
    
    If ActiveWorkbook.Sheets("TestCases").ProgressBar_Label.Visible Then
        ActiveWorkbook.Sheets("TestCases").ProgressBar_Label.Visible = False
    End If
    
    If ActiveWorkbook.Sheets("TestCases").ProgressBar_percentage.Visible Then
        ActiveWorkbook.Sheets("TestCases").ProgressBar_percentage.Visible = False
    End If
End Sub


Private Sub openProgressBar(barLabel As String)
        ActiveWorkbook.AutoSaveOn = False
        DoEvents
        GenericFunctions.UnprotectSheet
        Application.ScreenUpdating = True
        ActiveSheet.Range("A2").Activate
        ActiveWorkbook.Sheets("TestCases").ProgressBar_Label.Caption = barLabel
        ActiveWorkbook.Sheets("TestCases").ProgressBar_percentage.Text = "0%"
        
        ActiveWorkbook.Sheets("TestCases").ProgressBarLoad.Visible = True
        ActiveWorkbook.Sheets("TestCases").ProgressBarLoad.Top = 100
        ActiveWorkbook.Sheets("TestCases").ProgressBar_Label.Visible = True
        ActiveWorkbook.Sheets("TestCases").ProgressBar_Label.Top = 60
        ActiveWorkbook.Sheets("TestCases").ProgressBar_percentage.Visible = True
        ActiveWorkbook.Sheets("TestCases").ProgressBar_percentage.Top = 80
        Application.Wait (Now + 10 * milisecond)
        ProgressBarStatus = True
End Sub


