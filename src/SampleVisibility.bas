Attribute VB_Name = "SampleVisibility"
Sub ShowSample()
    InitializeWorkBook.UnhideSheet ("Sample")
End Sub

Sub HideSample()
    InitializeWorkBook.VeryHiddenSheet ("Sample")
End Sub
