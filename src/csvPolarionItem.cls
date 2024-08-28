VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "csvPolarionItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class module

Public value As String
Public colIndex As Integer
Public colName As String


' Event - triggered when class created
Private Sub Class_Initialize()

End Sub

Public Sub cellInit(curID As String)
    value = ""
    colIndex = 0
    colName = curID
End Sub

