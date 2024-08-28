VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "csvLastUpdatedClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'csvLastUpdatedClass module
Public lastUpdated As New csvPolarionItem

' Event - triggered when class created
Private Sub Class_Initialize()
    lastUpdated.cellInit ("lastUpdated")
End Sub

Public Property Get getValue() As String
    getValue = lastUpdated.value
End Property

Public Property Get getIndex() As Integer
    getIndex = lastUpdated.colIndex
End Property

Public Property Get getName() As String
    getName = lastUpdated.colName
End Property



Public Sub initClass(colIndex As Integer)
    lastUpdated.colIndex = colIndex
End Sub


Public Sub setValue(value)
    lastUpdated.value = value
End Sub
