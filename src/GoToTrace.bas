Attribute VB_Name = "GoToTrace"

'--------------------------------------------------------
'--------------------- Private Subs ---------------------
'--------------------------------------------------------



'-----------------------------------Go To Begining---------------------------------
'Sub Name:GoToBegining
'Description: This Sub is responsible for change the focus of the open woorksheet to the TRACE's sheet top
'Inputs: ---
'-----------------------------------------------------------------------------------
Sub GoToTrace()
    ActiveWorkbook.Worksheets("Trace").Activate
    Range("A2").Select
End Sub


