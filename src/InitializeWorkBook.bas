Attribute VB_Name = "InitializeWorkBook"
'-----------------------------------Initialize Workbook---------------------------------
'Function Name:InitializeWorkBook
'Description: This function is responsible for .
'Inputs:----
'-----------------------------------------------------------------------------------
Function InitializeWorkBook() As Boolean
    'Constants based on the Trace tab

    Dim WS_Count As Integer
    Dim curSheet As Integer
    Dim SheetsList As New list, sheetsToCreateList As New list
    Dim linkedReqs As Variant, linkedTests() As String
    Dim testCasesSheetCVs() As String
    Dim allTestsList As New Dictionary
    Dim calcPrevStatus As XlCalculation
    Dim newSheetsCreated As Boolean
    
    calcPrevStatus = Application.Calculation
    GenericFunctions.uiDisable
    ' Set WS_Count equal to the number of worksheets in the active
    ' workbook.
    WS_Count = ActiveWorkbook.Worksheets.count
    For curSheet = 1 To WS_Count
        If curSheet <= WS_Count Then
            If InStr(ActiveWorkbook.Worksheets(curSheet).Name, "Sheet") Then
                Application.DisplayAlerts = False
                ActiveWorkbook.Worksheets(curSheet).Delete
                Application.DisplayAlerts = True
                curSheet = curSheet - 1
                WS_Count = WS_Count - 1
            Else
                SheetsList.Add (ActiveWorkbook.Worksheets(curSheet).Name)
            End If
        End If
    Next
    
    
    For curRowNumber = 2 To 10000
        ActiveWorkbook.Worksheets("Trace").Activate
        If Not IsEmpty(Cells(curRowNumber, TRACE_CvNumberCN)) Then
            currentCV = Cells(curRowNumber, TRACE_WorkItemCN)
            If Not SheetsList.Contains(currentCV) Then
                SheetsList.Add (currentCV)
                sheetsToCreateList.Add (currentCV)
            End If
        Else
            Exit For
        End If
    Next
    SheetsList.RemoveStr ("Sample")
    SheetsList.RemoveStr ("Trace")
    SheetsList.RemoveStr ("TestCases")
    SheetsList.RemoveStr ("Statistics")
    newSheetsCreated = createNewSheets(sheetsToCreateList)
    For curRowNumber = 2 To 10000
        ActiveWorkbook.Worksheets("Trace").Activate
        If Not IsEmpty(Cells(curRowNumber, TRACE_CvNumberCN)) Then
            If Not IsEmpty(Cells(curRowNumber, TRACE_LinkedWorkItemsCN)) Then
                currentCV = Cells(curRowNumber, TRACE_WorkItemCN)
                If Not IsEmpty(currentCV) Then
                    If sheetsToCreateList.Contains(currentCV) Then
                        linkedReqsList = Cells(curRowNumber, TRACE_LinkedWorkItemsCN)
                        linkedTests = ReadLinkedTests(linkedReqsList)
                        linkedReqs = ReadLinkedReqs(linkedReqsList)
                        If SheetsList.Contains(currentCV) Then
                            ActiveWorkbook.Worksheets(currentCV).Activate
                            If arrayEmptyCheck(linkedTests) = 0 Then
                                fillTestCases (linkedTests)
                            End If
                            
                            If arrayEmptyCheck(linkedReqs) = 0 Then
                                fillSubRequirements (linkedReqs)
                            End If
                        End If
                        If arrayEmptyCheck(linkedTests) = 0 Then
                            For Each Test In linkedTests
                                allTestsList(Test) = ""
                            Next
                        End If
                    End If
                End If
            End If
        Else
            Exit For
        End If
    Next
    If allTestsList.count > 0 Then
        testCasesSheetCVs = readTestCasesSheet()
        A = updateTestCasesSheet_CvOnly(allTestsList, testCasesSheetCVs)
    End If
    GenericFunctions.uiEnable (calcPrevStatus)
    If newSheetsCreated Then
        MsgBox "Sheets created sucessfully"
    Else
        MsgBox "No sheet created"
    End If
    InitializeWorkBook = True
End Function

'--------------------------------------------------------
'------------------ Public Subs -------------------
'--------------------------------------------------------


'-----------------------------------Unhide Sheet---------------------------------
'Function Name:UnhideSheet
'Description: This function is responsible for make a given sheet visible.
'Inputs: sheetToUnhide: string with the name of the sheet to hide;
'-----------------------------------------------------------------------------------
Sub UnhideSheet(sheetToUnhide)
       Sheets(sheetToUnhide).Visible = True
End Sub


'-----------------------------------Very Hide Sheet---------------------------------
'Function Name:VeryHiddenSheet
'Description: This function is responsible for make a given sheet invisible.
'Inputs: sheetToHide: string with the name of the sheet to unhide;
'-----------------------------------------------------------------------------------
Sub VeryHiddenSheet(sheetToHide)
   Sheets(sheetToHide).Visible = xlVeryHidden
End Sub


'-----------------------------------Read Linked Requirements---------------------------------
'Function Name:ReadLinkedReqs
'Description: This function is responsible for split the received string in a list of CVs and return a list
'             with the ones linked as "is traced by"
'Inputs: celVal: string with the linked work items;
'Outputs: list with all CVs linked as "is traced by"
'-----------------------------------------------------------------------------------
Private Function ReadLinkedReqs(ByVal celVal As String) As String()
    Dim cvList() As String, reqsList() As String, auxArray() As String
    Dim j As Integer
    
    j = 0
    cvList = Strings.Split(celVal, ",")
    ReDim reqsList(SizeOfArray(cvList))
    For Each i In cvList
        auxArray() = Split(i, ":")
        If StrComp(Replace(auxArray(0), " ", ""), "istracedby", 1) = 0 Then
            cvNumberLenght = GLOBAL_cvMaxNumberLenght
            'While (Not IsNumeric(Mid(auxArray(i), cvLinePos + 2, cvNumberLenght))) And (cvNumberLenght > 0)
                'cvNumberLenght = cvNumberLenght - 1
            'Wend
            cvNumberLenght = cvNumberLenght - 1
            'reqsList(j) = Mid(auxArray(i), cvLinePos, cvNumberLenght + 3)
            j = j + 1
        End If
    Next i
    If j > 0 Then
        ReDim Preserve reqsList(j - 1)
        ReadLinkedReqs = reqsList
    End If
End Function
Private Function numericExtractor(strWithNumber As String, ByVal startPosition As Integer, ByVal maxLenght As Integer) As String
        
        While (Not IsNumeric(Mid(strWithNumber, startPosition + 2, maxLenght))) And (maxLenght > 0)
            maxLenght = maxLenght - 1
        Wend
        numericExtractor = Mid(strWithNumber, startPosition, maxLenght)
End Function

'-----------------------------------Read Linked Tests---------------------------------
'Function Name:ReadLinkedTests
'Description: This function is responsible for split the received string in a list of CVs and return a list
'             with the ones linked as "is tested by"
'Inputs: celVal: string with the linked work items;
'Outputs: list with all CVs linked as "is tested by"
'-----------------------------------------------------------------------------------
Private Function ReadLinkedTests(ByVal celVal As String) As String()
    Dim cvList() As String, testsList() As String, auxArray() As String
    Dim j As Integer
    
    j = 0
    cvList = Strings.Split(celVal, ",")
    ReDim testsList(SizeOfArray(cvList))
    For Each i In cvList
        auxArray() = Split(i, ":")
        If StrComp(Replace(auxArray(0), " ", ""), "istestedby", 1) = 0 Then
            cvLinePos = InStr(auxArray(1), "CV-")
            cvNumberLenght = GLOBAL_cvMaxNumberLenght
            testing = numericExtractor(auxArray(1), cvLinePos, 6)
            While (Not IsNumeric(Mid(auxArray(1), cvLinePos + 2 + cvNumberLenght, 1))) And (cvNumberLenght > 0)
                cvNumberLenght = cvNumberLenght - 1
            Wend
            testsList(j) = Replace(Mid(auxArray(1), cvLinePos, cvNumberLenght + 3), "CV-", "")
            j = j + 1
        End If
    Next i
    If j > 0 Then
        ReDim Preserve testsList(j - 1)
        ReadLinkedTests = testsList
    End If
End Function


'-----------------------------------Fill Tests Cases---------------------------------
'Function Name:fillTestCases
'Description: This function is responsible for insert the test cases numbers into the active sheet
'             and then copy the formulas of the second line into all other used lines.
'Inputs: TestCasesList: array with all test cases CVs
'-----------------------------------------------------------------------------------
Sub fillTestCases(TestCasesList)
    Dim cellCounter As Integer
    
    cellCounter = 0
    cellCounter = SizeOfArray(TestCasesList)
    rangeSelectionAddr = "B2:B" + CStr(cellCounter + 1)
    Range(rangeSelectionAddr).Select
    Range(rangeSelectionAddr).value = Application.Transpose(TestCasesList)
    Range("A2").Select
    Range("A2").Copy
    cell = ActiveSheet.Cells(cellCounter + 1, 1).Address(False, False)
    Range("A3", cell).Select
    ActiveSheet.Paste
    Range("B:B").NumberFormat = "CV-#"
    Range("D:E").NumberFormat = "CV-#"
    Range("C2:F2").Select
    Range("C2:F2").Copy
    cell = ActiveSheet.Cells(cellCounter + 2, 6).Address(False, False)
    Range("C3", cell).Select
    ActiveSheet.Paste
End Sub



'---------------------------------Fill Sub Requirements -------------------------------
'Function Name:fillSubRequirements
'Description: This function is responsible for insert the test cases numbers into the active sheet
'             and then copy the formulas of the second line into all other used lines.
'Inputs: TestCasesList: array with all test cases CVs
'-----------------------------------------------------------------------------------
Sub fillSubRequirements(subRequirementsList)
    Range("A2").Select
    While Not (IsEmpty(ActiveCell.value))
        ActiveCell.Offset(1, 0).Select
    Wend
    For Each cv In subRequirementsList
        ActiveCell.value = cv
        ActiveCell.Offset(1, 0).Select
    Next cv
End Sub
Sub prepareSheetTemplate()
    UnhideSheet ("Sample")
    ActiveWorkbook.Worksheets("Sample").Select
    Range("A1:K10").Select
    Range("A1:K10").Copy
End Sub
Sub applySheetTemplate(currentCVNumber As String)
    Sheets(currentCVNumber).Select
    Range("A1").Select
    ActiveSheet.Paste
    ActiveSheet.Columns("C").ColumnWidth = 20
    ActiveSheet.Columns("F").ColumnWidth = 20
    ActiveSheet.Columns("G").ColumnWidth = 100
End Sub
Sub closeSheetTemplate()
    VeryHiddenSheet ("Sample")
End Sub

Function createNewSheets(currentCVNumber As list) As Boolean
    On Error GoTo ErrorHandler
    
    createNewSheets = False
    If Not (currentCVNumber Is Nothing) Then
        If currentCVNumber.Size > 0 Then
            prepareSheetTemplate
            For Each cv In currentCVNumber.getList
                Set NewSheet = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.count))
                NewSheet.Name = cv
                applySheetTemplate (cv)
            Next
            closeSheetTemplate
            createNewSheets = True
        End If
    End If
    Exit Function
ErrorHandler:
        MsgBox ("There is some sheet already named: " + cv + " Please delete/rename it and try again!")
    Resume Next
End Function

Sub deleteAllSheets()
    Dim overwriteAnswer As VbMsgBoxResult
    
    overwriteAnswer = vbNo
    overwriteAnswer = MsgBox("Are you sure you want to delete ALL the sheets of this Workbook?", vbYesNo, "Delete Sheets!")
    If overwriteAnswer = vbYes Then
        Application.DisplayAlerts = False
        totalAmountOfSheets = ActiveWorkbook.Worksheets.count
        For curSheet = totalAmountOfSheets To 1 Step -1
            If Left(ActiveWorkbook.Worksheets(curSheet).Name, 3) = "CV-" Then
                ActiveWorkbook.Worksheets(curSheet).Delete
            End If
        Next
        Application.DisplayAlerts = True
    End If
End Sub
