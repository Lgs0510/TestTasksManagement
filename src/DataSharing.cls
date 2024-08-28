VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataSharing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private strData As String

Public Property Get getData() As Variant
    getData = strData
End Property

Public Property Let letArray(newData)
    strData = newData
End Property
