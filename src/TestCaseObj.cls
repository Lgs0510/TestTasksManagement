VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestCaseObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class module
Public cvNumber As String
Public testStatus As String
Public cvOld As String

Public Property Let letTestCase(newTestCase)
    cvNumber = newTestCase.cvNumber
    testStatus = newTestCase.testStatus
    cvOld = newTestCase.cvOld
End Property


Public Sub Add(cvObj)
    cvNumber = cvObj.cvNumber
    testStatus = cvObj.testStatus
    cvOld = cvObj.cvOld
End Sub

