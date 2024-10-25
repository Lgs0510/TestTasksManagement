Attribute VB_Name = "UpdateScriptsList"


'--------------------------------------------------------
'--------------------- Private Subs ---------------------
'--------------------------------------------------------

'------------------------------------Update Script List--------------------------------------
'Sub Name:UpdateScriptsList
'Description: This function is responsible for check an remove all duplicated CVs from the new list, then it
'             shall select the first available cell in the TestCases sheet and insert the list of new CVs there.
'Inputs: ---
'---------------------------------------------------------------------------------------------
Sub UpdateScriptsList()
    Dim scriptsTestCases As New Dictionary
    Dim calcPrevStatus As XlCalculation
    Dim debugList As New list
    
    If ActiveSheet.Name <> "TestCases" Then
        MsgBox ("Update Script List will only work at the TestCases sheet!")
        Exit Sub
    End If
    calcPrevStatus = Application.Calculation
    GenericFunctions.uiDisable
    
    If readScriptFolder(scriptsTestCases) Then
        protectionStatus = ActiveSheet.ProtectContents
        GenericFunctions.UnprotectSheet
        
        checkCurrentMappedTestCases testCasesDic:=scriptsTestCases

        insertNewTestCases testCasesDic:=scriptsTestCases

        ProgressLoadBarModule.closeProgressBar
        
        GenericFunctions.ProtectSheet (protectionStatus)
        MsgBox "Script List Updated!"
    End If
    
    GenericFunctions.uiEnable (calcPrevStatus)
    
End Sub

Private Function readScriptFolder(testCasesDic As Dictionary) As Boolean
    Dim scriptNameArray As New ScriptNameObj
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    SelectFolder.Show
    scriptFolder = DataComm.dataChannel.getData
    If scriptFolder = "" Then
        MsgBox ("Invalid path")
        readScriptFolder = False
        Exit Function
    ElseIf Not objFSO.FolderExists(scriptFolder) Then
        MsgBox ("Invalid path")
        readScriptFolder = False
        Exit Function
    End If
    
    Set objFolder = objFSO.GetFolder(scriptFolder)
    ProgressLoadBarModule.ProgressLoad curValue:=0, maxValue:=objFolder.Files.count, progressLabel:="Updating scripts list.... gathering scripts"
    For Each objFile In objFolder.Files
        If LCase(Right(objFile.Path, 4)) = ".txt" Then
            numberOfScripts = numberOfScripts + 1
            Set objCurScriptToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile(objFile, 1, True)
            Do Until (objCurScriptToRead.AtEndOfStream)
                strLine = objCurScriptToRead.ReadLine
                cvLinePos = InStr(strLine, "CV-")
                cvNumberLenght = GLOBAL_cvMaxNumberLenght
                If cvLinePos > 0 Then
                    While (Not IsNumeric(Mid(strLine, cvLinePos + 2 + cvNumberLenght, 1))) And (cvNumberLenght > 0)
                        cvNumberLenght = cvNumberLenght - 1
                    Wend
                    If cvNumberLenght > 0 Then
                        scriptNameArray.cvNumber = Replace(Mid(strLine, cvLinePos, cvNumberLenght + 3), "CV-", "")
                        scriptNameArray.ScriptName = objFile.Name
                        If Not testCasesDic.Exists(scriptNameArray.cvNumber) Then
                            testCasesDic(scriptNameArray.cvNumber) = scriptNameArray.ScriptName
                        End If
                    End If
                End If
            Loop
            ProgressLoadBarModule.ProgressLoad curValue:=numberOfScripts, maxValue:=objFolder.Files.count, progressLabel:="Updating scripts list.... gathering scripts"
        End If
    Next
    readScriptFolder = True
End Function

Private Sub checkCurrentMappedTestCases(ByRef testCasesDic As Dictionary)
    lastRowWithCVs = lastRowNumber
    totalTestCases = lastRowWithCVs
    
    For curRowNumber = 2 To lastRowWithCVs
        If Not IsEmpty(Cells(curRowNumber, TESTCASES_WorkItemCN)) Then
            curReqToSearch = CStr(Cells(curRowNumber, TESTCASES_WorkItemCN))
            If testCasesDic.Exists(curReqToSearch) Then
                Cells(curRowNumber, TESTCASES_ScriptNameCN) = testCasesDic(curReqToSearch)
                testCasesDic.Remove (curReqToSearch)
            End If
        End If
        ProgressLoadBarModule.ProgressLoad curValue:=curRowNumber, maxValue:=totalTestCases, progressLabel:="Updating scripts names"
    Next
End Sub

Private Sub insertNewTestCases(ByRef testCasesDic As Dictionary)
    Dim remainingAmountTestCases As Integer
    Dim addedAmountTestCases As Integer
        
    lastRowWithCVs = lastRowNumber
    If testCasesDic.count > 0 Then
        remainingAmountTestCases = testCasesDic.count
        For Each testCase In testCasesDic.Keys
            lastRowWithCVs = lastRowWithCVs + 1
            addedAmountTestCases = addedAmountTestCases + 1
            Cells(lastRowWithCVs, TESTCASES_WorkItemCN) = testCase
            Cells(lastRowWithCVs, TESTCASES_ScriptNameCN) = testCasesDic(testCase)
            ProgressLoadBarModule.ProgressLoad curValue:=addedAmountTestCases, maxValue:=remainingAmountTestCases, progressLabel:="Adding missing Test Cases"
        Next
    End If
End Sub
