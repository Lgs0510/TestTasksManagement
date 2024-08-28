Attribute VB_Name = "GoToEnd"

'--------------------------------------------------------
'--------------------- Private Subs ---------------------
'--------------------------------------------------------


'-----------------------------------Go To End---------------------------------
'Sub Name:GoToEnd
'Description: This Sub is responsible for change the focus of the open sheet to the sheet's bottom
'Inputs: ---;
'-----------------------------------------------------------------------------------
Sub GoToEnd()
    addr = Cells(Range("A" & Rows.count).End(xlUp).Row, 1).Address
    Range(addr).Select
End Sub


