Attribute VB_Name = "Statistics"

'--------------------------------------------------------
'-------------------- Private Subs ----------------------
'--------------------------------------------------------

'----------------------------------Get Test Statistics--------------------------------
'Sub Name:getTestStatistict
'Description: This sub is responsible for read all CVs in this workbook to calculate the amount of tests checked and to be test.
'              At the end, update the respective cell with the current amount of CVs
'Inputs: ---
'-------------------------------------------------------------------------------------
Sub getTestStatistict()
    Dim cvList As New TestCasesList
    Dim testCase As New TestCaseObj

    For Each Worksheet In ActiveWorkbook.Worksheets
        If InStr(Worksheet.Name, "CV-") Then
            Worksheet.Activate
            Range("B2").Select
            While Not (IsEmpty(ActiveCell.value))
                testCase.cvNumber = ActiveCell.value
                ActiveCell.Offset(0, 1).Select
                testCase.testStatus = ActiveCell.value
                cvList.Add testCase
                ActiveCell.Offset(1, -1).Select
            Wend
        End If
    Next
    cvList.RemoveDuplicates
    ActiveWorkbook.Worksheets("Statistics").Activate
    Range("B46").Select
    ActiveCell.value = cvList.Size
    getAprovedTestCases cvList
    getReprovedTestCases cvList
    getNotTestedCases cvList
End Sub


'--------------------------------Get Aproved Test Cases------------------------------
'Sub Name:getAprovedTestCases
'Description: This sub is responsible get the approved tests from the received test list.
'Inputs: testList - TestCasesList type object with all the tests in the current workbok
'-------------------------------------------------------------------------------------
Sub getAprovedTestCases(testList)
    Dim cvList As New ArrayClass

    Range("B47").Select
    ActiveCell.value = testList.CountApprovedTests
End Sub


'--------------------------------Get Repproved Test Cases------------------------------
'Sub Name:getReprovedTestCases
'Description: This sub is responsible get the repproved tests from the received test list.
'Inputs: testList - TestCasesList type object with all the tests in the current workbok
'-------------------------------------------------------------------------------------
Sub getReprovedTestCases(testList)
    Dim cvList As New ArrayClass

    Range("B48").Select
    ActiveCell.value = testList.CountReprovedTests
End Sub


'--------------------------------Get Not Tested Cases------------------------------
'Sub Name:getNotTestedCases
'Description: This sub is responsible to get the test cases that weren't tested from the
'              received test list.
'Inputs: testList - TestCasesList type object with all the tests in the current workbok
'-------------------------------------------------------------------------------------
Sub getNotTestedCases(testList)
    Dim cvList As New ArrayClass

    Range("B49").Select
    ActiveCell.value = testList.CountNotTested
End Sub



