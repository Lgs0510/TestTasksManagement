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
    Dim WkItemCN, StatusCN, OldCvCN, NewCvCN, ScriptNameCN, ScriptPathCN, ScriptPathLN As Integer
    Dim scriptNameArray As New ScriptNameObj
    Dim testsScriptlist As New CvArray
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    WkItemCN = 1
    StatusCN = 2
    OldCvCN = 3
    NewCvCN = 4
    ScriptNameCN = 6

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
                    cvNumberLenght = 6
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
    ActiveSheet.Unprotect
    For curRowNumber = 2 To numberOfCvs
        If Not IsEmpty(Cells(curRowNumber, WkItemCN)) Then
            If IsEmpty(Cells(curRowNumber, ScriptNameCN)) Then
                curReqToSearch = Cells(curRowNumber, WkItemCN)
                reqIndex = testsScriptlist.Find(curReqToSearch)
                If reqIndex >= 0 Then
                    Cells(curRowNumber, ScriptNameCN) = testsScriptlist.GetScriptName(CInt(reqIndex))
                End If
            End If
        End If
        Application.StatusBar = "Updating scripts list.... " + CStr(100 * curRowNumber / numberOfCvs) + "%"
    Next
    ActiveSheet.Protect _
        AllowFiltering:=True, _
        AllowSorting:=True
    Application.StatusBar = False
    MsgBox "Script List Updated!"
End Sub
