Attribute VB_Name = "importingCsv"
Dim csvTable As CsvClass

'--------------------------------------------------------
'------------------- Private Function -------------------
'--------------------------------------------------------



'-------------------------------Imports Csv Requirements----------------------------
'Function Name:ImportCsvRequirements
'Description: This Function is responsible for import new software requirements from a Polarion CSV export file
'Inputs: ---
'Output: CsvClass - object with all the requirements inside the CSV object
'-----------------------------------------------------------------------------------
Function ImportCsvRequirements() As CsvClass
    Dim csvIds As New list
    Dim lineIndex As Integer
    Dim property As String
    Dim csvFile As textFile_t
    
    Set csvTable = New CsvClass
    csvFile = File_op.getTextFile
    If csvFile.numberOfLines > 0 Then
        Do Until (csvFile.textFile.AtEndOfStream Or bDoneStatus)
            strLine = csvFile.textFile.ReadLine
            strLine = Replace(strLine, """", "")
            If csvFile.textFile.Line = 5 Then
                csvIds.letList = Split(strLine, ";")
                csvIdsNumber = SizeOfArray(csvIds)
                csvTable.initCsv csvSize:=csvFile.numberOfLines - 5, listIDs:=csvIds
            ElseIf csvFile.textFile.Line > 5 Then
                csvTable.addLine (strLine)
            End If
        Loop
        Set ImportCsvRequirements = csvTable
    Else
        Exit Function
    End If
End Function

