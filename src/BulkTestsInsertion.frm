VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BulkTestsInsertion 
   Caption         =   "New CV insertion"
   ClientHeight    =   5145
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9990.001
   OleObjectBlob   =   "BulkTestsInsertion.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BulkTestsInsertion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





'----------------- Public Local Variables----------------
Dim firstKeyPress As Boolean
Dim testsList() As String
Dim currentIndex As Integer
Dim curTest As New TestCaseObj
'--------------------------------------------------------
'----------------- Private Subs  EVENTs -----------------
'--------------------------------------------------------


'----------Bulk Test Insertion Text Box double Click(double click event)------------
'Sub Name:BulkTestsInsertion_txtBox_DblClick
'Description: This Sub is called when a double click is performed in the textBox of Test Cases (BulkTestsInsertion_txtBox).
'             It's pourpose is to clear the textbox from the example text.
'Input: Cancel
'-----------------------------------------------------------------------------------
Private Sub BulkTestsInsertion_txtBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    BulkTestsInsertion_txtBox.Text = ""
End Sub


'--------------Bulk Test Insertion Text Box Key Down (Key Down event)---------------
'Sub Name:BulkTestsInsertion_txtBox_KeyDown
'Description: This Sub is called when any key is pressed inside the textBox of Test Cases (BulkTestsInsertion_txtBox).
'             It's pourpose is to clear the textbox from the example text at first time.
'Inputs: KeyCode, Shift
'-----------------------------------------------------------------------------------
Private Sub BulkTestsInsertion_txtBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If firstKeyPress Then
        BulkTestsInsertion_txtBox.Text = ""
        firstKeyPress = False
    End If
End Sub


'--------------MultiPage object Change (Change event)---------------
'Sub Name:MultiPage1_Change
'Description: This Sub is called when the multipage object suffer any change, usualy the tab is changed in order to
'             update the ListBox with  all the inserted CVs.
'Inputs: ---
'-----------------------------------------------------------------------------------
Private Sub MultiPage1_Change()
    If MultiPage1.value = 1 Then
        testsListBox.list = testsList
        For cell = 1 To SizeOfArray(testsListBox.list) - 1
        Next
    End If
End Sub


'-----------------MultiPage page2 ADD Button click (click event)------------------
'Sub Name:Page2_Add_btn_Click
'Description: This Sub is called when the ADD button in page 1 of multipage object is clicked. It then will start the
'             process to add the new unique CVs into the TestCase sheet.
'Inputs: ---
'-----------------------------------------------------------------------------------
Private Sub Page2_Add_btn_Click()
    Dim allTestsList As New TestCasesList
    Dim testsArray() As New TestCaseObj
    Dim i As Integer
    Dim testCasesSheetCVs() As String
        
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    ReDim testsArray(SizeOfArray(testsList) - 1)
    For i = 1 To SizeOfArray(testsList) - 1
        testsArray(i - 1).cvNumber = testsList(i, 0)
        testsArray(i - 1).testStatus = testsList(i, 1)
        testsArray(i - 1).cvOld = testsList(i, 2)
    Next
    allTestsList.letArray = testsArray
    allTestsList.RemoveDuplicates
    If allTestsList.Size < 1 Then
        MsgBox "No new CV to add, perhaps they are already on the TestCases sheet."
        Exit Sub
    End If
    testCasesSheetCVs = readTestCasesSheet()
    A = updateTestCasesSheet(allTestsList, testCasesSheetCVs)
    updateTestCasesCVs allTestsList
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Unload Me
End Sub


'-----------------MultiPage page2 BACK Button click (click event)------------------
'Sub Name:Page2_Back_btn_Click
'Description: This Sub is called when the BACK button in page 1 of multipage object is clicked so the multipage returns
'             to page 1.
'Inputs: ---
'-----------------------------------------------------------------------------------
Private Sub Page2_Back_btn_Click()
    MultiPage1.value = 0
End Sub


'-----------------------Test Insertion OK Button Click(click event)-----------------
'Sub Name:TestInsertionOK_btn_Click
'Description: This Sub is called when the OK button in page 1 of multipage object is clicked. It will process the user input
'             in the text box to create a list with every inserted CV number and change the multipage object to page 1.
'Inputs: ---
'-----------------------------------------------------------------------------------
Private Sub TestInsertionOK_btn_Click()
    Dim cvLinePos As Long
     
    iputTextString = BulkTestsInsertion_txtBox.Text
    j = 0
    cvLinePos = 1
    While cvLinePos < Len(iputTextString)
        If (Mid(iputTextString, cvLinePos, 3) = "CV-") Then
            j = j + 1
            cvLinePos = cvLinePos + 3
        Else
            cvLinePos = cvLinePos + 1
        End If
    Wend
    ReDim testsList(j, 2)
    testsList(0, 0) = "TestCase"
    testsList(0, 1) = "Test Result"
    testsList(0, 2) = "Old CV"
    j = 1
    cvLinePos = 1
    While cvLinePos < Len(iputTextString)
        If (Mid(iputTextString, cvLinePos, 3) = "CV-") Then
            cvNumberLenght = GLOBAL_cvMinNumberLenght
            While (IsNumeric(Mid(iputTextString, cvLinePos + 2 + cvNumberLenght, 1))) And (cvNumberLenght < 7)
                cvNumberLenght = cvNumberLenght + 1
            Wend
            cvNumberLenght = cvNumberLenght - 1
            testsList(j, 0) = Mid(iputTextString, cvLinePos, cvNumberLenght + 3)
            j = j + 1
            cvLinePos = cvLinePos + cvNumberLenght + 3
        Else
            cvLinePos = cvLinePos + 1
        End If
    Wend
    
    MultiPage1.value = 1
End Sub


'-------------------------Test Listbox Click(click event)--------------------------
'Sub Name:testsListBox_Click
'Description: This Sub is called when a cick is performed anywhere inside the ListBox(testsListBox). It will load the
'             text boxs at the right side with the CV number and OLD CV number of the respective line, altogether with
'             the test result status.
'Inputs: ---
'-----------------------------------------------------------------------------------
Private Sub testsListBox_Click()
    If currentIndex <> testsListBox.ListIndex Then
        currentIndex = testsListBox.ListIndex
        curTest.letTestCase = loadListBoxLine()
        Application.EnableEvents = False
        txtBoxCvNumber = Mid(curTest.cvNumber, 4, Len(curTest.cvNumber) - 3)
        If Len(curTest.cvOld) > 3 Then
            txtBoxOldCvNumber = Mid(curTest.cvOld, 4, Len(curTest.cvOld) - 3)
        Else
            txtBoxOldCvNumber = ""
        End If
        TestsResultStatus_btn = convTestStatus(curTest.testStatus)
        Application.EnableEvents = True
    End If
End Sub


'--------------Test Result Button Change (change event)---------------
'Sub Name:TestsResultStatus_btn_Change
'Description: This Sub is called when a change happens in the Test Result Button(TestsResultStatus_btn). It will call the
'             button uptade SUB and clear the current test status variable and the List Box current line.
'Inputs: ---
'-----------------------------------------------------------------------------------
Private Sub TestsResultStatus_btn_Change()
    setTestResultButton
    curTest.testStatus = ""
    updateListBoxLine
End Sub


'-----------------------Test Result Status Button Click(click event)-----------------
'Sub Name:TestsResultStatus_btn_Click
'Description: This Sub is called when the Test Result button(TestsResultStatus_btn) is clicked. It will then call the
'             button update SUB and update the current test status variable and List Box.
'Inputs: ---
'-----------------------------------------------------------------------------------
Private Sub TestsResultStatus_btn_Click()
    If TestsResultStatus_btn = True Then
        curTest.testStatus = "OK"
    ElseIf TestsResultStatus_btn = False Then
        curTest.testStatus = "NOK"
    Else
        curTest.testStatus = ""
    End If
    updateListBoxLine
End Sub


'--------------Text Box Cv Number Change (change event)---------------
'Sub Name:txtBoxCvNumber_Change
'Description: This Sub is called when a change happens in the Text Box of the CV Number(txtBoxCvNumber) in page 1 of the
'             MultiPage object. It will update the local variable and the listBox.
'Inputs: ---
'-----------------------------------------------------------------------------------
Private Sub txtBoxCvNumber_Change()
    curTest.cvNumber = "CV-" + txtBoxCvNumber
    updateListBoxLine
End Sub


'--------------Text Box Old Cv Number Change (change event)---------------
'Sub Name:txtBoxOldCvNumber_Change
'Description: This Sub is called when a change happens in the Text Box of the Old CV Number(txtBoxCvNumber) in page 1 of the
'             MultiPage object. It will update the local variable and the listBox.
'Inputs: ---
'-----------------------------------------------------------------------------------
Private Sub txtBoxOldCvNumber_Change()
    If txtBoxOldCvNumber <> "" Then
        curTest.cvOld = "CV-" + txtBoxOldCvNumber
        updateListBoxLine
    End If
End Sub

'-----------------------User Form Initialize(initialize event)-----------------
'Sub Name:UserForm_Initialize
'Description: This Sub is called at the initialization process of this UserFoms, before it is shown. It will initialize
'             some local variables.
'Inputs: ---
'-----------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    firstKeyPress = True
    MultiPage1.value = 0
End Sub


'--------------------------------------------------------
'--------------------- Private Subs ---------------------
'--------------------------------------------------------

'------------------------------Set Test Result Button------------------------------
'Sub Name:setTestResultButton
'Description: This Sub is responsible for update(text and color) the Test Result Button(TestsResultStatus_btn) accordingly
'             with its current state or, if null(it is a tristate button) with an empty string.
'Inputs: ---
'-----------------------------------------------------------------------------------
Private Sub setTestResultButton()
    If TestsResultStatus_btn = True Then
        TestsResultStatus_btn.BackColor = &H80FF80
        TestsResultStatus_btn.Caption = "OK"
    ElseIf TestsResultStatus_btn = False Then
        TestsResultStatus_btn.BackColor = &H8080FF
        TestsResultStatus_btn.Caption = "NOK"
    Else
        If IsNull(TestsResultStatus_btn) Then
            TestsResultStatus_btn.Caption = ""
        Else
            TestsResultStatus_btn.Caption = TestsResultStatus_btn
        End If
        TestsResultStatus_btn.BackColor = &HE0E0E0
    End If
End Sub


'----------------------------Update List Box selected line----------------------------
'Sub Name:updateListBoxLine
'Description: This Sub is responsible for write the listBox current selected line with the latest data in correpondant
'             local variables
'Inputs: ---
'-----------------------------------------------------------------------------------
Private Sub updateListBoxLine()
    If curTest.cvNumber <> "" Then
        testsList(currentIndex, 0) = curTest.cvNumber
    End If
    If curTest.cvOld <> "" Then
        testsList(currentIndex, 2) = curTest.cvOld
    End If
    testsList(currentIndex, 1) = curTest.testStatus
    MultiPage1_Change
End Sub


'--------------------------------------------------------
'------------------ Private Functions -------------------
'--------------------------------------------------------

'------------------------------Convert Test Status------------------------------
'Function Name:convTestStatus
'Description: This function is responsible for return the boolean correspondant of the test case string OK/NOK. Or in case
'             it has something different, like DRAFT/"", return it.
'Inputs: testCaseStatus: string with the value set for the Test Result
'-----------------------------------------------------------------------------------

Private Function convTestStatus(testCaseStatus)
    If testCaseStatus = "OK" Then
        convTestStatus = True
    ElseIf testCaseStatus = "NOK" Then
        convTestStatus = False
    Else
        convTestStatus = testCaseStatus
    End If
End Function


'----------------------------Load List Box selected line----------------------------
'Function Name:loadListBoxLine
'Description: This Function is responsible for read the listBox current selected line and poppulate the correpondant local variables
'Inputs: ---
'Output: returns the testCase object witch contains the CV number, the test result and the old CV number.
'-----------------------------------------------------------------------------------
Private Function loadListBoxLine() As TestCaseObj
    
    curTest.cvNumber = testsListBox.list(currentIndex, 0)
    curTest.testStatus = testsListBox.list(currentIndex, 1)
    curTest.cvOld = testsListBox.list(currentIndex, 2)
    Set loadListBoxLine = curTest
End Function



