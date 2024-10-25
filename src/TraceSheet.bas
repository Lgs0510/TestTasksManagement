Attribute VB_Name = "TraceSheet"
'------------------------------Read Trace sheet------------------------------
'Function Name:readReqsSheet
'Description: This function is responsible for read all the requirements from the Trace sheet and return them
'             in a string array.
'Inputs: ---
'Output: string array with all the CVs in requirements sheet.
'-----------------------------------------------------------------------------------
Public Function readTraceSheetReqs() As String()
    Dim i As Integer
    Dim sheetTraceReqs() As String
    
    currentSheetName = ActiveSheet.Name
    ActiveWorkbook.Worksheets("Trace").Activate
    LastRow = lastRowNumber
    If LastRow < 2 Then
        Exit Function
    End If
    wholeReqsList = Range("A2", "A" + CStr(LastRow)).value
    reqsNumber = SizeOfArray(wholeReqsList)
    ReDim sheetTraceReqs(reqsNumber - 1)
    i = 0
    If reqsNumber > 1 Then
        For Each cv In wholeReqsList
            sheetTraceReqs(i) = cv
            i = i + 1
        Next
    Else
        sheetTraceReqs(0) = wholeReqsList
    End If
    ActiveWorkbook.Worksheets(currentSheetName).Activate
    readTraceSheetReqs = sheetTraceReqs
End Function

Sub DeleteRequirement()
    Dim overwriteAnswer As VbMsgBoxResult
    Dim listToDelete As New list
    Dim listOfDeletedCVs As New list

    protectStatus = ActiveSheet.ProtectContents
    calcPrevStatus = Application.Calculation
    If ActiveSheet.Name <> "Trace" Then
        MsgBox ("Delete Requirements can only delete CVs at the TRACE sheet!")
        Exit Sub
    End If
    If ActiveSheet.Name = "Trace" Then
        GenericFunctions.UnprotectSheet
        GenericFunctions.uiDisable
        For Each selCell In Selection
            If Not listToDelete.Contains(selCell.Row) Then
                currentCV = Cells(selCell.Row, 2).value
                If deleteAllAnswer = 0 Then
                    deleteAllAnswer = MsgBox("Do you want to delete all the selected requirements?", vbYesNo, "Delete Requirement!")
                End If
                If deleteAllAnswer = vbNo Then
                    deleteAnswer = MsgBox("Are you sure you want to delete " + currentCV + "?", vbYesNo, "Delete Requirement!")
                End If
                
                If deleteAllAnswer = vbYes Or deleteAnswer = vbYes Then
                        listToDelete.Add (selCell.Row)
                        Debug.Print sheetExist(currentCV)
                        If sheetExist(currentCV) Then
                            Application.DisplayAlerts = False
                            ActiveWorkbook.Sheets(currentCV).Delete
                            Application.DisplayAlerts = True
                        End If
                End If
            End If
        Next
        listToDelete.SortUpSideDown
        For Each cellToDelete In listToDelete.getList
            ActiveWorkbook.Sheets("Trace").Rows(cellToDelete).EntireRow.Delete
        Next
        GenericFunctions.ProtectSheet (protectStatus)
        GenericFunctions.uiEnable (calcPrevStatus)
    End If
End Sub
