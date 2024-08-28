VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "csvIDClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'csvIDClass module
Private id As New csvPolarionItem

' Event - triggered when class created
Private Sub Class_Initialize()
    id.cellInit ("ID")
End Sub

Public Property Get getValue() As String
    getValue = id.value
End Property

Public Property Get getIndex() As Integer
    getIndex = id.colIndex
End Property

Public Property Get getName() As String
    getName = id.colName
End Property

Public Sub initClass(colIndex As Integer)
    id.colIndex = colIndex
End Sub

Public Sub setValue(value)
    id.value = value
End Sub
