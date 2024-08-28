VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "csvTitleClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'csvTitleClass module
Public Title As New csvPolarionItem

' Event - triggered when class created
Private Sub Class_Initialize()
    Title.cellInit ("Title")
End Sub

Public Property Get getValue() As String
    getValue = Title.value
End Property

Public Property Get getIndex() As Integer
    getIndex = Title.colIndex
End Property

Public Property Get getName() As String
    getName = Title.colName
End Property





Public Sub initClass(colIndex As Integer)
    Title.colIndex = colIndex
End Sub


Public Sub setValue(value)
    Title.value = value
End Sub
