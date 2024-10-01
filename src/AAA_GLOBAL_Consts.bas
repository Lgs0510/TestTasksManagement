Attribute VB_Name = "AAA_GLOBAL_Consts"
'Trace Sheet constants
Public Const TRACE_CvNumberCN = 1
Public Const TRACE_WorkItemCN = 2
Public Const TRACE_TestStatusCN = 3
Public Const TRACE_LinkedWorkItemsCN = 8



'TestCases Sheet constants
Public Const TESTCASES_WorkItemCN = 1
Public Const TESTCASES_StatusCN = 2
Public Const TESTCASES_OldCvCN = 3
Public Const TESTCASES_NewCvCN = 4
Public Const TESTCASES_ScriptNameCN = 6
Public Const TESTCASES_WorkItemCL = "A"
Public Const TESTCASES_StatusCL = "B"
Public Const TESTCASES_OldCvCL = "C"
Public Const TESTCASES_NewCvCL = "D"
Public Const TESTCASES_ScriptNameCL = "F"

'CVs Sheets constants
Public Const CVs_SHEETS_TestCvCN = 2
Public Const CVs_SHEETS_StatusCN = 3
Public Const CVs_SHEETS_OldCvCN = 4
Public Const CVs_SHEETS_NewCvCN = 5
Public Const CVs_SHEETS_TestCvCL = "B"
Public Const CVs_SHEETS_StatusCL = "C"
Public Const CVs_SHEETS_OldCvCL = "D"
Public Const CVs_SHEETS_NewCvCL = "E"



'GLOBAL Constants
Public Const sheetsProtectionPassword = "naotemsenha"
Public Const testCaseStatusToDELETE = "DELETED, DRAFT"
Public Const GLOBAL_cvMaxNumberLenght = 6
Public Const GLOBAL_cvMinNumberLenght = 1
'Formula of Work Item collumn in Trace sheet
Public Const Trace_WorkItemFormula_00 = "=IF(INDIRECT(CONCAT(""A"";ROW()))<>"""";"
Public Const Trace_WorkItemFormula_01 = "HYPERLINK(CONCAT(CONCAT(""["";TEXTBEFORE(TEXTAFTER(CELL(""filename"";"
Public Const Trace_WorkItemFormula_02 = "INDIRECT(""A1""));""["");""]"");""]'CV-"");INDIRECT(CONCAT(""A"";ROW()));""'!A1"");"
Public Const Trace_WorkItemFormula_03 = "CONCAT(""CV-"";INDIRECT(CONCAT(""A"";ROW()))));NA())"

'Formula of Test Status collumn in Trace sheet
Public Const Trace_TestStatusFormula_00 = "=IF(INDIRECT(CONCAT(""A"";ROW()))<>"""";IF((COUNTIF(INDIRECT(CONCAT(""'CV-"";"
Public Const Trace_TestStatusFormula_01 = " INDIRECT(CONCAT(""A"";ROW())); ""'!C:C""));""OK""))=(COUNTIF(INDIRECT(CONCAT(""'CV-"";"
Public Const Trace_TestStatusFormula_02 = "INDIRECT(CONCAT(""A"";ROW())); ""'!A:A""));""<>"")-1);""Tests OK""; IF((COUNTIF(INDIRECT"
Public Const Trace_TestStatusFormula_03 = "(CONCAT(""'CV-""; INDIRECT(CONCAT(""A"";ROW())); ""'!C:C""));""NOK""))>0; ""NOK""; "
Public Const Trace_TestStatusFormula_04 = """Pending....""));"""")"

