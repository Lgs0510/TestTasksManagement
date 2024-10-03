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
    Dim SheetsList As New list
    Dim WS_Count As Integer
    
    protectStatus = ActiveSheet.ProtectContents
    WS_Count = ActiveWorkbook.Worksheets.count

    For curSheet = 1 To WS_Count
        SheetsList.Add (ActiveWorkbook.Worksheets(curSheet).Name)
    Next
    If ActiveCell.Row > 1 Then
        currentCV = "CV-" + CStr(Range(CVs_SHEETS_CvNumberCL + CStr(ActiveCell.Row)))
        If Not IsEmpty(currentCV) Then
            overwriteAnswer = MsgBox("Are you sure you want to delete " + currentCV + "?", vbYesNo, "Delete Requirement!")
            If overwriteAnswer = vbYes Then
                If SheetsList.Contains(currentCV) Then
                    Application.DisplayAlerts = False
                    ActiveWorkbook.Sheets(currentCV).Delete
                    Application.DisplayAlerts = True
                End If
                GenericFunctions.UnprotectSheet
                ActiveWorkbook.Sheets("Trace").Rows(ActiveCell.Row).EntireRow.Delete

                GenericFunctions.ProtectSheet (protectStatus)
            End If
        End If
    End If
End Sub
