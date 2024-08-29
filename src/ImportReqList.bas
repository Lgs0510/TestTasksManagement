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
    
    protectStatus = ActiveSheet.ProtectContents
    UnprotectSheet (protectStatus)
    Set csvReqs = ImportCsvRequirements
    If csvReqs Is Nothing Then
        Exit Sub
    End If
    ActiveWorkbook.Worksheets("Trace").Activate
    LastRow = ActiveSheet.Range("A" & ActiveSheet.Rows.count).End(xlUp).Row
    curTraceReqlist.letList = readTraceSheetReqs
    For Each req In csvReqs.getReqListNO
        If Not curTraceReqlist.Contains(Replace(req, "CV-", "")) Then
            LastRow = Range("A" & Rows.count).End(xlUp).Row
            Cells(LastRow + 1, CvNumberCN) = req
            LastRow = LastRow + 1
            curentRowNmb = LastRow
        Else
             overwriteAnswer = MsgBox("This requirement is already on the list!" + vbCrLf + "Do you want to update it?", vbYesNo, "WorkItem already Exist!")
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
        Cells(curentRowNmb, LinkedWorkItemsCN) = csvReqs.getReqLikedWkItems("CV-" + req)
    Next
    ProtectSheet (protectStatus)
    InitializeWorkBook.InitializeWorkBook
End Sub
