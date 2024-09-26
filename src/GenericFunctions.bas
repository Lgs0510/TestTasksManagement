Attribute VB_Name = "GenericFunctions"
'--------------------------------------------------------
'------------------ Public Functions -------------------
'--------------------------------------------------------


'-----------------------------------Size Of Array---------------------------------
'Function Name:SizeOfArray
'Description: This function is responsible for calculate the lenght of a given array, if valid, and return it.
'             In case the array is null, it shall return 0, if a non array variable is passed, it shall return 1.
'Inputs: arrayToMeasure: array to measure;
'Output: Length of the array, in positions
'-----------------------------------------------------------------------------------
Public Function SizeOfArray(arrayToMeasure As Variant) As Integer
    If Not IsNull(arrayToMeasure) Then
        If VarType(arrayToMeasure) > 8000 Then
            SizeOfArray = UBound(arrayToMeasure, 1) - LBound(arrayToMeasure, 1) + 1
        Else
            SizeOfArray = 1
        End If
    Else
        SizeOfArray = 0
    End If
End Function


'-----------------------------------Array Empty Check---------------------------------
'Function Name:arrayEmptyCheck
'Description: This function is responsible for check if an array variable is empty, in that case it returns TRUE.
'Inputs: arrayToTest: array to test for emtpy;
'Output: Boolean - True if the array is empty
'-----------------------------------------------------------------------------------
Public Function arrayEmptyCheck(arrayToTest As Variant) As Boolean
    On Error Resume Next
    
    intUpper = UBound(arrayToTest)
    
    If Err = 0 Then
        arrayEmptyCheck = False
    Else
        Err.Clear
        arrayEmptyCheck = True
    End If
End Function


'-----------------------------------Is In Array---------------------------------
'Function Name:IsInArray
'Description: This function is responsible for search an specified array for an determined string and return TRUE if finds it.
'Inputs: stringToBeFound: string to search for;
'Inputs: arr: array to search into;
'Output: True if the string exist inside the array; False otherwise
'-----------------------------------------------------------------------------------
Public Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    Dim i
    If Not IsNull(arr) Then
        For i = LBound(arr) To UBound(arr)
            If arr(i) = stringToBeFound Then
                IsInArray = True
                Exit Function
            End If
        Next i
    End If
    IsInArray = False

End Function


'-----------------------------------Unprotect Sheet---------------------------------
'Function Name:UnprotectSheet
'Description: This function is responsible for unprotec the active sheet if its parameter, protStatus is True.
'Inputs: protStatus: boolean flag to indicated if the active sheet is protected;
'-----------------------------------------------------------------------------------
Public Sub UnprotectSheet()
        ActiveSheet.Unprotect (sheetsProtectionPassword)
End Sub


'-----------------------------------Protect Sheet---------------------------------
'Function Name:ProtectSheet
'Description: This function is responsible for protec the active sheet if its parameter, protStatus is True.
'Inputs: protStatus: boolean flag to indicated if the active sheet should be protected;
'-----------------------------------------------------------------------------------
Public Sub ProtectSheet(protStatus)
        If protStatus Then
            ActiveSheet.Protect _
                Password:=sheetsProtectionPassword, _
                AllowFiltering:=True, _
                AllowSorting:=True
        End If
End Sub


'-----------------------------------Sheet Exist---------------------------------
'Function Name:sheetExist
'Description: This function is responsible for
'Inputs: prsheetName: string with the name of the Sheet in question
'-----------------------------------------------------------------------------------
Public Function sheetExist(sheetName As String) As Boolean
    On Error Resume Next
    
     ActiveWorkbook.Worksheets(sheetName).Select
    If Err = 0 Then
        sheetExist = True
    Else
        Err.Clear
        sheetExist = False
    End If
End Function


'-----------------------------------Last Row Number---------------------------------
'Function Name:lastRowNumber
'Description: This function is responsible for get the last row with data in the current active sheet
'Output: integer with the last row number
'-----------------------------------------------------------------------------------
Public Function lastRowNumber() As Integer
    lastRowNumber = ActiveSheet.Range("A" & ActiveSheet.Rows.count).End(xlUp).Row
End Function



'--------------------------------User Interface Disable-----------------------------
'Sub Name:uiDisable
'Description: This sub is responsible for disable the screen update, the events and the automatic calculation, in order to gain speed for some heavy processing
'Inputs:--
'-----------------------------------------------------------------------------------
Sub uiDisable()
        Application.Calculation = xlCalculationManual
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        g_vbaIsRunning = True
End Sub


'--------------------------------User Interface Enable-----------------------------
'Sub Name:uiEnable
'Description: This sub is responsible for enable the screen update, the events and the automatic calculation, in order to recovery the normal operation
'Inputs:--
'-----------------------------------------------------------------------------------
Sub uiEnable(calStatus)
        If calStatus Then
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            Application.EnableEvents = True
            g_vbaIsRunning = False
        End If
End Sub

