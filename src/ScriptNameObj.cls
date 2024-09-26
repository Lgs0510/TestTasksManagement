VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ScriptNameObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class module
Public cvNumber As String
Public ScriptName As String

Public Sub Add(cvObj)
    cvNumber = cvObj.cvNumber
    ScriptName = cvObj.ScriptName
End Sub

Public Function retrieve() As ScriptNameObj
    Dim cvObj As New ScriptNameObj
    
    cvObj.cvNumber = cvNumber
    cvObj.ScriptName = ScriptName
    Set retrieve = cvObj
End Function

