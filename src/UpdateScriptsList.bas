Attribute VB_Name = "UpdateScriptsList"


Dim scriptNameArray As New ScriptNameObj
Dim testsScriptlist As New CvArray

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
    Dim calcPrevStatus As XlCalculation
    
    If readScriptFolder Then
        protectionStatus = ActiveSheet.ProtectContents
        calcPrevStatus = Application.Calculation
        GenericFunctions.UnprotectSheet
        testsScriptlist.RemoveDuplicates

        GenericFunctions.uiDisable
        
        checkCurrentMappedTestCases

        insertNewTestCases
        
        GenericFunctions.uiEnable(calcPrevStatus)

        GenericFunctions.ProtectSheet (protectionStatus)
        MsgBox "Script List Updated!"
    End If
    
End Sub

Private Function readScriptFolder() As Boolean
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
                        scriptNameArray.cvNumber = Mid(strLine, cvLinePos, cvNumberLenght + 3)
                        scriptNameArray.ScriptName = objFile.Name
                        'When objects are pased as paramenter, ther should be no parentheses
                        testsScriptlist.Add scriptNameArray
                    End If
                End If
            Loop
            ProgressLoadBarModule.ProgressLoad curValue:=numberOfScripts, maxValue:=objFolder.Files.count, progressLabel:="Updating scripts list.... gathering scripts"
        End If
    Next
    
    ProgressLoadBarModule.closeProgressBar
    readScriptFolder = True
End Function

Sub checkCurrentMappedTestCases()
    lastRowWithCVs = lastRowNumber
    totalTestCases = lastRowWithCVs
    
    For curRowNumber = 2 To lastRowWithCVs
        If Not IsEmpty(Cells(curRowNumber, TESTCASES_WorkItemCN)) Then
            curReqToSearch = Cells(curRowNumber, TESTCASES_WorkItemCN)
            reqIndex = testsScriptlist.Find(curReqToSearch)
            If reqIndex >= 0 Then
                Cells(curRowNumber, TESTCASES_ScriptNameCN) = testsScriptlist.GetScriptName(CInt(reqIndex))
                testsScriptlist.Remove (reqIndex)
            End If
        End If
        ProgressLoadBarModule.ProgressLoad curValue:=curRowNumber, maxValue:=totalTestCases, progressLabel:="Updating Test Cases From Scripts"
    Next
    ProgressLoadBarModule.closeProgressBar
End Sub

Sub insertNewTestCases()
    Dim remainingAmountTestCases As Integer
    Dim addedAmountTestCases As Integer

  If testsScriptlist.Size > 0 Then
        remainingAmountTestCases = testsScriptlist.Size
        For Each testCase In testsScriptlist.getArray
            lastRowWithCVs = lastRowWithCVs + 1
            addedAmountTestCases = addedAmountTestCases + 1
            Cells(lastRowWithCVs, TESTCASES_WorkItemCN) = testCase.cvNumber
            Cells(lastRowWithCVs, TESTCASES_ScriptNameCN) = testCase.ScriptName
            ProgressLoadBarModule.ProgressLoad curValue:=addedAmountTestCases, maxValue:=remainingAmountTestCases, progressLabel:="Updating Test Cases From Scripts"
        Next
    End If
    ProgressLoadBarModule.closeProgressBar
End Sub
