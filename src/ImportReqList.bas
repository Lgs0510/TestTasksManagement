Attribute VB_Name = "ImportReqList"
Option Explicit

'--------------------------------------------------------
'------------------- Private Sub -------------------
'--------------------------------------------------------



'-------------------------------Imports Main Requirements----------------------------
'Sub Name:ImportMainReqs
'Description: This Function is responsible for start the import of requirements for the Trace sheet
'Inputs: ---
'-----------------------------------------------------------------------------------
Sub ImportMainReqs()
    Dim csvReqs As CsvClass
    Dim curTraceReqlist As New list
    Dim LastRow As Integer
    Dim curentRowNmb As Integer
    Dim req As Variant
    Dim protectStatus As Boolean
    Dim overwriteAnswer As VbMsgBoxResult
    Dim overwriteAllAnswer As VbMsgBoxResult
    Dim calcPrevStatus As XlCalculation

    If ActiveSheet.Name <> "Trace" Then
        MsgBox ("Import CSV will only work at the TRACE sheet!")
        Exit Sub
    End If
    GenericFunctions.uiDisable
    protectStatus = ActiveSheet.ProtectContents
    Set csvReqs = ImportCsvRequirements
    If csvReqs Is Nothing Then
        Exit Sub
    End If
    LastRow = lastRowNumber
    curTraceReqlist.letList = readTraceSheetReqs
    For Each req In csvReqs.getReqListNO
        calcPrevStatus = Application.Calculation
        GenericFunctions.UnprotectSheet
        If Not curTraceReqlist.Contains(Replace(req, "CV-", "")) Then
            LastRow = lastRowNumber
            Cells(LastRow + 1, TRACE_CvNumberCN) = req
            Cells(LastRow, TRACE_WorkItemCN).Formula2R1C1 = Trace_WorkItemFormula_00 & Trace_WorkItemFormula_01 & Trace_WorkItemFormula_02 & Trace_WorkItemFormula_03
            Cells(LastRow, TRACE_TestStatusCN).Formula2R1C1 = Trace_TestStatusFormula_00 & Trace_TestStatusFormula_01 & Trace_TestStatusFormula_02 & Trace_TestStatusFormula_03 & Trace_TestStatusFormula_04
            LastRow = LastRow + 1
            curentRowNmb = LastRow
        Else
             
             If overwriteAllAnswer = 0 Then
                overwriteAllAnswer = MsgBox("One or more requirements are already on the list!" + vbCrLf + "Do you want to update them all?", vbYesNo, "WorkItem already Exist!")
             End If
             If overwriteAllAnswer = vbNo Then
                overwriteAnswer = MsgBox("This requirement, CV-" + CStr(req) + ", is already on the list!" + vbCrLf + "Do you want to update it?", vbYesNo, "WorkItem already Exist!")
             ElseIf overwriteAllAnswer = vbYes Then
                overwriteAnswer = vbYes
             End If
             
            If overwriteAnswer = vbYes Then
                curentRowNmb = curTraceReqlist.Find(Replace(req, "CV-", "")) + 2
                If sheetExist("CV-" + CStr(Replace(req, "CV-", ""))) Then
                    Application.DisplayAlerts = False
                    ActiveWorkbook.Worksheets("CV-" + CStr(Replace(req, "CV-", ""))).Delete
                    Application.DisplayAlerts = True
                End If
            Else
                curentRowNmb = curTraceReqlist.Find(req) + 2
            End If
        End If
        ActiveWorkbook.Worksheets("Trace").Activate
        Cells(curentRowNmb, TRACE_LinkedWorkItemsCN) = csvReqs.getReqLikedWkItems("CV-" + req)
    Next
    GenericFunctions.ProtectSheet (protectStatus)
    GenericFunctions.uiEnable(calcPrevStatus)

    InitializeWorkBook.InitializeWorkBook
End Sub
