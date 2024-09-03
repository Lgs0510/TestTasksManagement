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
    
    If ActiveCell.Column = WorkItemCN Then
        If InStr(ActiveCell.Column, "CV-") Then
            overwriteAnswer = MsgBox("Are you sure you want to delete " + ActiveCell.value + "?", vbYesNo, "Delete Requirement!")
            If overwriteAnswer = vbYes Then
                ActiveWorkbook.Sheets(ActiveCell.value).Delete
                ActiveWorkbook.Sheets("Trace").Rows(ActiveCell.Row).EntireRow.Delete
            End If
        End If
    End If
End Sub
