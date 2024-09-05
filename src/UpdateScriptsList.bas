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
    Dim bDoneStatus As Boolean
    Dim debugVar As Boolean
    Dim scriptNameArray As New ScriptNameObj
    Dim testsScriptlist As New CvArray
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    numberOfCvs = Range("I5").value
    
    SelectFolder.Show
    scriptFolder = DataComm.dataChannel.getData
    If scriptFolder = "" Then
        MsgBox ("Invalid path")
        Exit Sub
    ElseIf Not objFSO.FolderExists(scriptFolder) Then
        MsgBox ("Invalid path")
        Exit Sub
    End If
    
    Application.StatusBar = "Updating scripts list.... gathering data"
    Set objFolder = objFSO.GetFolder(scriptFolder)
    For Each objFile In objFolder.Files
        If LCase(Right(objFile.Path, 4)) = ".txt" Then
            Set objCurScriptToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile(objFile, 1, True)
            Do Until (objCurScriptToRead.AtEndOfStream Or bDoneStatus)
                strLine = objCurScriptToRead.ReadLine
                cvLinePos = InStr(strLine, "CV-")
                If cvLinePos > 0 Then
                    
                    While (Not IsNumeric(Mid(strLine, cvLinePos + 2 + cvNumberLenght, 1))) And (cvNumberLenght > 0)
                        cvNumberLenght = cvNumberLenght - 1
                    Wend
                    If cvNumberLenght > 0 Then
                        scriptNameArray.cvNumber = Mid(strLine, cvLinePos, cvNumberLenght + 3)
                        scriptNameArray.ScriptName = objFile.Name
                        'When objects are pased as paramenter, ther should be no parentheses
                        testsScriptlist.Add scriptNameArray
                        numberOfCvs = numberOfCvs + 1
                    End If
                End If
            Loop
        End If
    Next
    ActiveSheet.Unprotect (sheetsProtectionPassword)
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    g_vbaIsRunning = True
    For curRowNumber = 2 To numberOfCvs
        If Not IsEmpty(Cells(curRowNumber, TESTCASES_WorkItemCN)) Then
            curReqToSearch = Cells(curRowNumber, TESTCASES_WorkItemCN)
            reqIndex = testsScriptlist.Find(curReqToSearch)
            If reqIndex >= 0 Then
                Cells(curRowNumber, TESTCASES_ScriptNameCN) = testsScriptlist.GetScriptName(CInt(reqIndex))
            End If
        End If
        'Application.StatusBar = "Updating scripts list.... " + CStr(100 * curRowNumber / numberOfCvs) + "%"
    Next
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.Protect _
        Password:=sheetsProtectionPassword, _
        AllowFiltering:=True, _
        AllowSorting:=True
    Application.StatusBar = False
    g_vbaIsRunning = False
    MsgBox "Script List Updated!"
End Sub
