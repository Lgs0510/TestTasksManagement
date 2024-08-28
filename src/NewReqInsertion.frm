VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewReqInsertion 
   Caption         =   "New Requirement Insertion"
   ClientHeight    =   3345
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15660
   OleObjectBlob   =   "NewReqInsertion.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NewReqInsertion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub btnInsertion_Click()
    Dim curReqList As New list
    Dim testCasesArray() As String
    Dim overwriteAnswer As VbMsgBoxResult
    
    If Not IsNumeric(txtBoxCvNumber) & Len(txtBoxCvNumber) > 0 Then
        MsgBox ("CV Number invalid! Only numbers!")
        Exit Sub
    ElseIf Len(txtBoxCvNumber) = 0 Then
        Unload Me
        Exit Sub
    End If
 
    LastRow = ActiveSheet.Range("A" & ActiveSheet.Rows.count).End(xlUp).Row
    wholeTestCasesList = Range("A2", "A" + CStr(LastRow))
    
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
    
    If Not curReqList.Contains(txtBoxCvNumber) Then
        rowToUpdate = ActiveCell.Row
    Else
        overwriteAnswer = MsgBox("This requirement is already on the list!" + vbCrLf + "Do you want to update it?", vbYesNo, "WorkItem already Exist!")
        If overwriteAnswer = vbYes Then
            rowToUpdate = curReqList.Find(txtBoxCvNumber) + 2
            If sheetExist("CV-" + CStr(txtBoxCvNumber)) Then
                Application.DisplayAlerts = False
                ActiveWorkbook.Worksheets("CV-" + CStr(txtBoxCvNumber)).Delete
                Application.DisplayAlerts = True
            End If
        Else
            Exit Sub
        End If
    End If
    ActiveSheet.Unprotect
    If overwriteAnswer = 0 Then
        ActiveCell.value = txtBoxCvNumber
    End If
    Cells(rowToUpdate, 8).value = txtBoxLinkedWI
    ActiveSheet.Protect _
        AllowFiltering:=True, _
        AllowSorting:=True
    InitializeWorkBook.InitializeWorkBook
    Unload Me
End Sub

