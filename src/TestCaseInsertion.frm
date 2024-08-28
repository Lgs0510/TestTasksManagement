VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TestCaseInsertion 
   Caption         =   "Test Cases Insertion"
   ClientHeight    =   4320
   ClientLeft      =   360
   ClientTop       =   1410
   ClientWidth     =   14040
   OleObjectBlob   =   "TestCaseInsertion.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TestCaseInsertion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub btnInsertion_Click()
    'Insert Button
    Dim curReqList As New list
    Dim testCasesArray() As String
    Dim testCaseCv As New TestCaseObj
    
    curRow = Selection.Row
    If Not IsNumeric(txtBoxCvNumber) & Len(txtBoxCvNumber) > 0 Then
        MsgBox ("CV Number invalid! Only numbers!")
        Exit Sub
    ElseIf Len(txtBoxCvNumber) = 0 Then
        Unload Me
        Exit Sub
    End If
    
    If Not IsNumeric(txtBoxOldCvNumber) & Len(txtBoxOldCvNumber) > 0 Then
        MsgBox ("CV Number invalid! Only numbers!")
        Exit Sub
    End If

    wholeList = "A" + CStr(Range("J5").value)
    wholeTestCasesList = Range("A2", wholeList)
    
    testCasesNumber = SizeOfArray(wholeTestCasesList)
    ReDim testCasesArray(testCasesNumber - 1)
    i = 0
    If testCasesNumber > 1 Then
        For Each cv In wholeTestCasesList
            testCasesArray(i) = cv
            i = i + 1
        Next
    Else
        testCasesArray(0) = wholeTestCasesList
    End If
    curReqList.letList = testCasesArray

    Cells(curRow, 1).Select
    If curReqList.Contains("CV-" + txtBoxCvNumber) Then
        If IsEmpty(Cells(Selection.Row, 1)) Or (StrComp("CV-" + txtBoxCvNumber, Cells(Selection.Row, 1).value) <> 0) Then
            MsgBox "This requirement is already on the list!      Line: " + CStr(curReqList.Find("CV-" + txtBoxCvNumber) + 2)
            Exit Sub
        End If
    End If
    
    ActiveSheet.Unprotect
    Cells(Selection.Row, 1) = "CV-" + txtBoxCvNumber
    If Len(txtBoxOldCvNumber) > 0 Then
        Cells(Selection.Row, 3) = "CV-" + txtBoxOldCvNumber
    End If
    
    Select Case btnTestResult
        Case True
            Cells(Selection.Row, 2) = "OK"
        Case False
            Cells(Selection.Row, 2) = "NOK"
        Case Else
            Cells(Selection.Row, 2) = ""
    End Select

    strTestResult = Cells(Selection.Row, 2)
    ActiveSheet.Protect _
        AllowFiltering:=True, _
        AllowSorting:=True
        
    testCaseCv.cvNumber = "CV-" + txtBoxCvNumber
    If Len(txtBoxOldCvNumber) > 0 Then
        testCaseCv.cvOld = "CV-" + txtBoxOldCvNumber
    Else
        testCaseCv.cvOld = ""
    End If
    testCaseCv.testStatus = strTestResult
    updateTestCasesCVs testCaseCv
    Unload Me
End Sub

Private Sub btnTestResult_Change()
    If btnTestResult = True Then
        btnTestResult.BackColor = &H80FF80
        btnTestResult.Caption = "OK"
    ElseIf btnTestResult = False Then
        btnTestResult = False
        btnTestResult.BackColor = &H8080FF
        btnTestResult.Caption = "NOK"
    Else
        btnTestResult.BackColor = &HE0E0E0
        btnTestResult.Caption = ""
    End If
End Sub


Private Sub UserForm_Initialize()
    txtBoxCvNumber = Cells(Selection.Row, 1)
    
    If Not IsNumeric(txtBoxCvNumber) Then
        cvLinePos = InStr(txtBoxCvNumber, "CV-")
        cvNumberLenght = 6
        While (Not IsNumeric(Mid(txtBoxCvNumber, cvLinePos + 2 + cvNumberLenght, 1))) And (cvNumberLenght > 0)
            cvNumberLenght = cvNumberLenght - 1
        Wend
        txtBoxCvNumber = Mid(txtBoxCvNumber, cvLinePos + 3, cvNumberLenght)
    End If
    
    txtBoxOldCvNumber = Cells(Selection.Row, 3)
    If Not IsNumeric(txtBoxOldCvNumber) Then
        cvLinePos = InStr(txtBoxOldCvNumber, "CV-")
        cvNumberLenght = 6
        While (Not IsNumeric(Mid(txtBoxOldCvNumber, cvLinePos + 2 + cvNumberLenght, 1))) And (cvNumberLenght > 0)
            cvNumberLenght = cvNumberLenght - 1
        Wend
        txtBoxOldCvNumber = Mid(txtBoxOldCvNumber, cvLinePos + 3, cvNumberLenght)
    End If
    
    testResult = Cells(Selection.Row, 2)
    If testResult = "OK" Then
        btnTestResult = True
        btnTestResult.BackColor = &H80FF80
    ElseIf testResult = "NOK" Then
        btnTestResult = False
        btnTestResult.BackColor = &H8080FF
    Else
        btnTestResult.BackColor = &HE0E0E0
    End If
        btnTestResult.Caption = testResult
End Sub
