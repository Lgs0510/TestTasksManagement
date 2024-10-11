Attribute VB_Name = "TestCaseSheet"


'--------------------------------------------------------
'------------------ Public Functions -------------------
'--------------------------------------------------------


'------------------------------Update TestCases sheet - CVs Only------------------------------
'Function Name:updateTestCasesSheet_CvOnly_CvOnly
'Description: This function is responsible for check an remove all duplicated CVs from the new list, then it
'             shall select the first available cell in the TestCases sheet and insert the list of new CVs there.
'Inputs: newTestCasesList: list class containing all the new CVs to add to the TestCases sheet;
'Outputs: testCasesSheetList: string array containing all the old CVs in the TestCases sheet;
'--------------------------------------------------------------------------------------------
Public Function updateTestCasesSheet_CvOnly(newTestCasesList As list, testCasesSheetList() As String)
    Dim tCSLcopy As New list
    Dim cvSpaceless As String
    
    protectStatus = ActiveSheet.ProtectContents
    If IsNull(newTestCasesList) Or newTestCasesList.Size <= 0 Then
        Exit Function
    End If
    If Not IsNull(testCasesSheetList) Then
        tCSLcopy.letList = testCasesSheetList
    End If
    
    If newTestCasesList.Size > 0 Then
        If tCSLcopy.Size > 0 Then
            For Each cv In newTestCasesList.getList
                If tCSLcopy.Contains(cv) Then
                    index = newTestCasesList.Find(cv)
                    If Not IsNull(index) Then
                        newTestCasesList.Remove (index)
                    End If
                End If
                If newTestCasesList.Size < 1 Then
                    Exit Function
                End If
            Next
        End If
        newTestCasesList.Sort
        currentSheetName = ActiveSheet.Name
        ActiveWorkbook.Worksheets("TestCases").Activate
        lastCellAddress = "A" + CStr(tCSLcopy.Size + 2) + ":A" + CStr(tCSLcopy.Size + 1 + newTestCasesList.Size)
        Range(lastCellAddress).Select
        
        GenericFunctions.UnprotectSheet

        Range(lastCellAddress).value = Application.Transpose(newTestCasesList.getList)
        
        GenericFunctions.ProtectSheet(protectStatus)

        updateNewCVsFormulas
    End If
End Function



'------------------------------Update TestCases sheet------------------------------
'Function Name:updateTestCasesSheet
'Description: This function is responsible for check an remove all duplicated CVs from the new list, then it
'             shall select the first available cell in the TestCases sheet and insert the list of new CVs there.
'Inputs: newTestCasesList: list class containing all the new CVs to add to the TestCases sheet;
'Inputs: testCasesSheetList: string array containing all the old CVs in the TestCases sheet;
'-----------------------------------------------------------------------------------
Public Function updateTestCasesSheet(newTestCasesList As TestCasesList, testCasesSheetList() As String)
    Dim tCSLcopy As New list
    Dim cvSpaceless As String
        
    protectStatus = ActiveSheet.ProtectContents

    If IsNull(newTestCasesList) Or newTestCasesList.Size <= 0 Then
        Exit Function
    End If
    If Not IsNull(testCasesSheetList) Then
        tCSLcopy.letList = testCasesSheetList
    End If
    
    If newTestCasesList.Size > 0 Then
        If tCSLcopy.Size > 0 Then
            For Each cv In newTestCasesList.getArray
                If tCSLcopy.Contains(cv.cvNumber) Then
                    index = newTestCasesList.Find(cv.cvNumber)
                    If Not IsNull(index) Then
                        newTestCasesList.Remove (index)
                    End If
                End If
                If newTestCasesList.Size < 1 Then
                    Exit Function
                End If
            Next
        End If
        newTestCasesList.Sort
        currentSheetName = ActiveSheet.Name
        ActiveWorkbook.Worksheets("TestCases").Activate
        lastCellAddress = "A" + CStr(tCSLcopy.Size + 2)
        Range(lastCellAddress).Select
        
        GenericFunctions.UnprotectSheet
        For Each cv In newTestCasesList.getArray
            ActiveCell.value = cv.cvNumber
            ActiveCell.Offset(0, 1).Select
            ActiveCell.value = CV.testStatus
            ActiveCell.Offset(0, 1).Select
            ActiveCell.value = CV.cvOld
            ActiveCell.Offset(1, -2).Select
        Next
        
        GenericFunctions.ProtectSheet(protectStatus)

        updateNewCVsFormulas
    End If
End Function



'------------------------------Read TestCases sheet------------------------------
'Function Name:readTestCasesSheet
'Description: This function is responsible for read all the test cases from the TestCases sheet and return them
'             in a string array.
'Inputs: ---
'Output: string array with all the CVs in TestCases sheet.
'-----------------------------------------------------------------------------------
Public Function readTestCasesSheet() As String()
    Dim i As Integer
    Dim sheetTestCases() As String
    
    currentSheetName = ActiveSheet.Name
    ActiveWorkbook.Worksheets("TestCases").Activate
    LastRow = lastRowNumber
    If LastRow < 2 Then
        Exit Function
    End If
    wholeTestCasesList = Range("A2", "A" + CStr(LastRow)).value
    testCasesNumber = SizeOfArray(wholeTestCasesList)
    ReDim sheetTestCases(testCasesNumber - 1)
    i = 0
    If testCasesNumber > 1 Then
        For Each cv In wholeTestCasesList
            sheetTestCases(i) = cv
            i = i + 1
        Next
    Else
        sheetTestCases(0) = wholeTestCasesList
    End If
    ActiveWorkbook.Worksheets(currentSheetName).Activate
    readTestCasesSheet = sheetTestCases
End Function



'------------------------------Update Test Cases CV------------------------------
'Function Name:updateTestCasesCVs
'Description: This function is responsible for check in all sheets from current workbook for old CVs and overwrite
'             them with the new branched CV, accordingly to the received list.
'Inputs: newCVsList - list(TestCasesList) with all new cvs added to TestCases sheet
'-----------------------------------------------------------------------------------
Public Sub updateTestCasesCVs(newTestCvsList)
    Dim testList As New TestCasesList
    Dim isSingleUpdate As Boolean
    Dim index As Integer
    Dim cellCvNumber As String
    
     If StrComp(TypeName(newTestCvsList), "TestCaseObj", vbTextCompare) = 0 Then
        isSingleUpdate = True
    Else
        isSingleUpdate = False
     End If
    
    For Each curSheet In ActiveWorkbook.Sheets
        If StrComp(Left(curSheet.Name, 3), "CV-", vbTextCompare) = 0 Then
            LastRow = curSheet.Range("A" & curSheet.Rows.count).End(xlUp).Row
            For Each cell In curSheet.Range("B2", "B" + CStr(LastRow))
                cellCvNumber = cell
                If isSingleUpdate Then
                    If StrComp(newTestCvsList.cvOld, cellCvNumber, vbTextCompare) = 0 Then
                        curSheet.Range(cell.Address).value = newTestCvsList.cvNumber
                    End If
                Else
                    index = newTestCvsList.FindOldCv(cellCvNumber)  '------- testList is saved as CV-xxxx remove all CVs or insert them all
                    If index >= 0 Then
                       curSheet.Range(cell.Address).value = newTestCvsList.GetCV(index)
                    End If
                End If
            Next
        End If
    Next
End Sub



'------------------------------Delete Test Cases------------------------------
'Function Name:deleteTestCases
'Description: This function is responsible delete the selected test case from the TestCase sheet, deleting the whole line
'Inputs: --
'-----------------------------------------------------------------------------------
Public Sub deleteTestCases()
    Dim deleteAnswer As VbMsgBoxResult
    Dim listToDelete As New list
    Dim listOfDeletedCVs As New list


    protectStatus = ActiveSheet.ProtectContents
    calcPrevStatus = Application.Calculation
    If ActiveSheet.Name = "TestCases" Then
        deleteAnswer = MsgBox("Are you sure you want to delete the selected TestCases?", vbYesNo, "Delete Test Cases?")
        For Each selCell In Selection
            If Not listToDelete.Contains(selCell.Row) Then
                listToDelete.Add (selCell.Row)
                listOfDeletedCVs.Add (selCell.value)
            End If
        Next
        If listToDelete.Size > 0 Then
            listToDelete.SortUpSideDown
            GenericFunctions.UnprotectSheet
            GenericFunctions.uiDisable
            For Each selCell In listToDelete.getList
                If selCell > 0 Then
                    Range("A" + selCell).EntireRow.Delete
                End If
            Next
            GenericFunctions.ProtectSheet (protectStatus)

            updateNewCVsFormulas
            removeTestCasesCVs listOfDeletedCVs.getList
            GenericFunctions.uiEnable (calcPrevStatus)
        End If
    End If
End Sub



'------------------------------Remove Test Cases CVs------------------------------
'Function Name:removeTestCasesCVs
'Description: This function is responsible for look for all the Test Cases, from received list, 
'              present in all sheets and remove them.
'Inputs: testList - list(list) with all cvs that must get removed.
'-----------------------------------------------------------------------------------
Sub removeTestCasesCVs(testList)
    Dim testListToRemove As New list
    Dim isSingleUpdate As Boolean
    Dim index As Integer
    Dim cellCvNumber As String
      
    testListToRemove.letList = testList
    For Each curSheet In ActiveWorkbook.Sheets
        If StrComp(Left(curSheet.Name, 3), "CV-", vbTextCompare) = 0 Then
            LastRow = curSheet.Range("A" & curSheet.Rows.count).End(xlUp).Row
            For Each cell In curSheet.Range("B2", "B" + CStr(LastRow))
                If testListToRemove.Contains(cell.value) Then
                   curSheet.Range(cell.Address).EntireRow.Delete
                End If
            Next
        End If
    Next
End Sub

'------------------------------Update New CVs Formulas------------------------------
'Function Name:updateNewCVsFormulas
'Description: This function is responsible for keep the Nev CV collumn (in TestCases sheet) with the formula for find the New CV number.
'Inputs: --
'-----------------------------------------------------------------------------------
Sub updateNewCVsFormulas()
        protectStatus = ActiveSheet.ProtectContents
        GenericFunctions.UnprotectSheet
        Range("D2").Copy
        Range("D3:D" + CStr(lastRowNumber + 1000)).PasteSpecial
        
        GenericFunctions.ProtectSheet(protectStatus)
End Sub
