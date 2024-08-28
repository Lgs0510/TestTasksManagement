VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "csvLinkedWorkItensClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'csvLinkedWorkItensClass module
Public linkedWorkItens As New csvPolarionItem

' Event - triggered when class created
Private Sub Class_Initialize()
    linkedWorkItens.cellInit ("Linked Work Items")
End Sub


Public Property Get getValue() As String
    getValue = linkedWorkItens.value
End Property

Public Property Get getIndex() As Integer
    getIndex = linkedWorkItens.colIndex
End Property

Public Property Get getName() As String
    getName = linkedWorkItens.colName
End Property



Public Sub initClass(colIndex As Integer)
    linkedWorkItens.colIndex = colIndex
End Sub



Public Sub setValue(value)
    linkedWorkItens.value = value
End Sub
