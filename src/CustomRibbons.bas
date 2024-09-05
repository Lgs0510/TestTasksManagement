Attribute VB_Name = "CustomRibbons"
Option Explicit


'--------------------------------------------------------
'--------------------- Private Subs ---------------------
'--------------------------------------------------------


'---------------------------------Back Up User Config-------------------------------
'Sub Name:backUpUserConfig
'Description: This sub has the porpouse to copy the current Excel UI user config file to a ".back" for backup.
'Inputs: ---;
'-----------------------------------------------------------------------------------
Sub backUpUserConfig()
 Dim sOfficeUIDir As String
 Dim backUpFile As String
 Dim sUIFile As String
 Dim sTest As String
 
 sOfficeUIDir = Environ("USERPROFILE") & "\AppData\Local\Microsoft\Office\"
 sUIFile = sOfficeUIDir & "Excel.officeUI"
 backUpFile = sOfficeUIDir & "Excel.officeUI.back"
 sTest = Dir(sUIFile)
 If Not sTest = "" Then
  FileCopy sUIFile, backUpFile
 End If
End Sub


'---------------------------------Check Bad C-losure-------------------------------
'Sub Name:checkBadClosure
'Description: This sub has the porpouse to check if there is a Excel UI user config file with ".back" witch indicate
'             that the last closure of the file did not recover the user config file correctly.
'Inputs: ---;
'-----------------------------------------------------------------------------------
Sub checkBadClosure()
    Dim sOfficeUIDir As String
    Dim backUpFile As String
    Dim sTest As String
    
    sOfficeUIDir = Environ("USERPROFILE") & "\AppData\Local\Microsoft\Office\"
    backUpFile = sOfficeUIDir & "Excel.officeUI.back"
    sTest = Dir(backUpFile)
    If Not sTest = "" Then
       Kill (backUpFile)
    End If
End Sub


'-------------------------------------Copy File-----------------------------------
'Sub Name:sbCopyFile
'Description: This sub has the porpouse to copy a custom config file into the folder used by
'             Excel overwriting the existing one.
'Inputs: ---;
'-----------------------------------------------------------------------------------
Sub sbCopyFile()
 Dim sOfficeUIDir As String
 Dim userDown As String
 Dim sHWFile As String
 Dim sUIFile As String
 Dim sTest As String
 
 xmlCreate (CVribbons_xml)
 sOfficeUIDir = Environ("USERPROFILE") & "\AppData\Local\Microsoft\Office\"
 userDown = Environ("USERPROFILE") & "\Downloads\"
 sHWFile = userDown & "customConfig.xml"
 sUIFile = sOfficeUIDir & "Excel.officeUI"
 sTest = Dir(sHWFile)
 backUpUserConfig
 If Not sTest = "" Then
  FileCopy sHWFile, sUIFile
 End If
 removeCreatedConfig
End Sub


'------------------------------remove Created Config--------------------------------
'Sub Name:removeCreatedConfig
'Description: This sub has the porpouse to delete the custom config file created for
'             overwrite the user config.
'Inputs: ---;
'-----------------------------------------------------------------------------------
Sub removeCreatedConfig()
 Dim sHWFile As String
 Dim sTest As String
 
 sHWFile = Environ("USERPROFILE") & "\Downloads\customConfig.xml"
 sTest = Dir(sHWFile)
 If Not sTest = "" Then
  Kill (sHWFile)
 End If
End Sub


'----------------------------------Delete File------------------------------------
'Sub Name:sbDeleteFile
'Description: This sub has the porpouse to delete the current Excel UI config file.
'Inputs: ---;
'-----------------------------------------------------------------------------------
Sub sbDeleteFile()
 Dim sOfficeUIDir As String
 Dim sUIFile As String
 Dim sTest As String
 sOfficeUIDir = Environ("USERPROFILE") & "\AppData\Local\Microsoft\Office\"
 sUIFile = sOfficeUIDir & "Excel.officeUI"
 sTest = Dir(sUIFile)
 If Not sTest = "" Then
  Kill (sUIFile)
 End If
End Sub


'------------------------------Restore Old User Config------------------------------
'Sub Name:restoreOldUserConfig
'Description: This sub has the porpouse restore the user config file from the ".back"
'Inputs: ---;
'-----------------------------------------------------------------------------------
Sub restoreOldUserConfig()
 Dim sOfficeUIDir As String
 Dim userDown As String
 Dim origFile As String
 Dim sUIFile As String
 Dim sHWFile As String
 Dim sTest As String

 sOfficeUIDir = Environ("USERPROFILE") & "\AppData\Local\Microsoft\Office\"
 origFile = sOfficeUIDir & "Excel.officeUI.back"
 sUIFile = sOfficeUIDir & "Excel.officeUI"
 sTest = Dir(origFile)
 If sTest = "" Then
    xmlCreate (defaultRibbons_xml)
    sOfficeUIDir = Environ("USERPROFILE") & "\AppData\Local\Microsoft\Office\"
    userDown = Environ("USERPROFILE") & "\Downloads\"
    sHWFile = userDown & "customConfig.xml"
    FileCopy sHWFile, origFile
 End If
 FileCopy origFile, sUIFile
 Kill (origFile)
End Sub
    

'--------------------------------------xml Create----------------------------------
'Sub Name:xmlCreate
'Description: This Sub has the porpouse to create a Excel UI xml config file
'Inputs: xmlText: string to use insert in the file;
'-----------------------------------------------------------------------------------
Sub xmlCreate(xmlText As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim oFile As Object
    
    Set oFile = fso.CreateTextFile(Environ("USERPROFILE") & "\Downloads\customConfig.xml")
    oFile.Write xmlText
    oFile.Close
    Set fso = Nothing
    Set oFile = Nothing
End Sub

'--------------------------------------------------------
'------------------- Private Functions ------------------
'--------------------------------------------------------


'------------------------------Clear Vision Ribbons------------------------------
'Function Name:CVribbons_xml
'Description: This function has the porpouse to return the string that compose the Excel UI config file for the Clear Vision
'Inputs: ---;
'Output: String with the text that will compose the file
'-----------------------------------------------------------------------------------
Function CVribbons_xml() As String
    Dim xmlText As String
    xmlText = "<mso:customUI xmlns:mso='http://schemas.microsoft.com/office/2009/07/customui'>"
    xmlText = xmlText & "<mso:ribbon>"
    xmlText = xmlText & "<mso:qat>"
    xmlText = xmlText & "<mso:sharedControls>"
    xmlText = xmlText & "<mso:control idQ=""mso:AutoSaveSwitch"" visible=""true""/>"
    xmlText = xmlText & "<mso:control idQ=""mso:FileNewDefault"" visible=""false""/>"
    xmlText = xmlText & "<mso:control idQ=""mso:FileOpenUsingBackstage"" visible=""false""/>"
    xmlText = xmlText & "<mso:control idQ=""mso:FileSave"" visible=""true""/>"
    xmlText = xmlText & "<mso:control idQ=""mso:FileSendAsAttachment"" visible=""false""/>"
    xmlText = xmlText & "<mso:control idQ=""mso:FilePrintQuick"" visible=""false""/>"
    xmlText = xmlText & "<mso:control idQ=""mso:PrintPreviewAndPrint"" visible=""false""/>"
    xmlText = xmlText & "<mso:control idQ=""mso:Spelling"" visible=""false""/>"
    xmlText = xmlText & "<mso:control idQ=""mso:Undo"" visible=""true""/>"
    xmlText = xmlText & "<mso:control idQ=""mso:Redo"" visible=""true""/>"
    xmlText = xmlText & "<mso:control idQ=""mso:SortAscendingExcel"" visible=""false""/>"
    xmlText = xmlText & "<mso:control idQ=""mso:SortDescendingExcel"" visible=""false""/>"
    xmlText = xmlText & "<mso:control idQ=""mso:PointerModeOptions"" visible=""false""/>"
    xmlText = xmlText & "<mso:control idQ=""mso:VisualBasic"" visible=""true""/>"
    xmlText = xmlText & "</mso:sharedControls>"
    xmlText = xmlText & "</mso:qat>"
    xmlText = xmlText & "<mso:tabs>"
    xmlText = xmlText & "<mso:tab idQ=""mso:TabDrawInk"" visible=""false""/>"
    xmlText = xmlText & "<mso:tab idQ=""mso:TabBackgroundRemoval"" visible=""false""/>"
    xmlText = xmlText & "<mso:tab idQ=""mso:TabHome"" visible=""false""/>"
    xmlText = xmlText & "<mso:tab idQ=""mso:TabInsert"" visible=""false""/>"
    xmlText = xmlText & "<mso:tab idQ=""mso:TabPageLayoutExcel"" visible=""false""/>"
    xmlText = xmlText & "<mso:tab idQ=""mso:TabFormulas"" visible=""false""/>"
    xmlText = xmlText & "<mso:tab idQ=""mso:TabData"" visible=""false""/>"
    xmlText = xmlText & "<mso:tab idQ=""mso:TabReview"" visible=""false""/>"
    xmlText = xmlText & "<mso:tab idQ=""mso:TabView"" visible=""false""/>"
    xmlText = xmlText & "<mso:tab idQ=""mso:TabAutomate"" visible=""false""/>"
    xmlText = xmlText & "<mso:tab idQ=""mso:HelpTab"" visible=""false""/>"
    xmlText = xmlText & "<mso:tab idQ=""mso:TabAddIns"" visible=""false""/>"
    xmlText = xmlText & "<mso:tab idQ=""mso:TabDeveloper"" visible=""false""/>"
    xmlText = xmlText & "<mso:tab id=""mso_c1.1F0DA1A6"" label=""CLEAR VISION"">"
    xmlText = xmlText & "<mso:group id=""mso_c2.1F0DA1B6"" label=""TRACE"" autoScale=""true"">"
    xmlText = xmlText & "<mso:button id=""InitializeWorkBook_Ribbon"" label=""Build Workbook"" imageMso=""AutoFormat"" onAction=""InitializeWorkBook.InitializeWorkBook"" visible=""true""/>"
    xmlText = xmlText & "<mso:button id=""InsertRequirement_Ribbon"" label=""Insert 1 Requirement"" imageMso=""LassoSelect"" onAction=""InsertRequirement.InsertRequirement"" visible=""true""/>"
    xmlText = xmlText & "<mso:button id=""ImportRequirementsFromCSV"" label=""Import From CSV"" imageMso=""HorizontalSpacingDecrease"" onAction=""ImportReqList.ImportMainReqs"" visible=""true""/>"
    xmlText = xmlText & "<mso:button id=""OverwriteOldCVsWithNewCVs"" label=""Overwrite Old CVs with New CVs"" imageMso=""HyperlinksVerify"" onAction=""TestCasesWorkItems.UpdateOldCvWithNewCv"" visible=""true""/>"
    xmlText = xmlText & "<mso:button id=""DeleteRequirement"" label=""Delete Selected Requirement"" imageMso=""CancelRequest"" onAction=""TraceSheet.DeleteRequirement"" visible=""true""/>"
    xmlText = xmlText & "</mso:group>"
    xmlText = xmlText & "<mso:group id=""mso_c1.1F151500"" label=""TestCases"" autoScale=""true"">"
    xmlText = xmlText & "<mso:button id=""UpdateScriptsList_Ribbon"" label=""Update Script List"" imageMso=""RecordsRefreshMenu"" onAction=""UpdateScriptsList.UpdateScriptsList"" visible=""true""/>"
    xmlText = xmlText & "<mso:button id=""TestCases_BulkAdd_Ribbon"" label=""Insert List of Test Cases"" imageMso=""Bullets"" onAction=""TestCases_BulkAdd.TestCases_BulkAdd"" visible=""true""/>"
    xmlText = xmlText & "<mso:button id=""Delete_Selected_CV"" label=""Delete Selected CV"" imageMso=""CancelRequest"" onAction=""TestCaseSheet.deleteTestCases"" visible=""true""/>"
    xmlText = xmlText & "</mso:group>"
    xmlText = xmlText & "<mso:group id=""mso_c2.1F15A2A9"" label=""Statistics"" autoScale=""true"">"
    xmlText = xmlText & "<mso:button id=""xlsm_getTestStatistict_Ribbon"" label=""Update Test Statistics"" imageMso=""DatabaseCopyDatabaseFile"" onAction=""getTestStatistict"" visible=""true""/>"
    xmlText = xmlText & "</mso:group>"
    xmlText = xmlText & "<mso:group id=""mso_c1.1F1B41C0"" label=""GENERAL"" autoScale=""true"">"
    xmlText = xmlText & "<mso:button id=""GoToEnd_Ribbon"" label=""End Of List"" imageMso=""_3DPerspectiveDecrease"" onAction=""GoToEnd.GoToEnd"" visible=""true""/>"
    xmlText = xmlText & "<mso:button id=""GoToBegining_Ribbon"" label=""Begining of the List"" imageMso=""_3DPerspectiveIncrease"" onAction=""GoToBegining.GoToBegining"" visible=""true""/>"
    xmlText = xmlText & "<mso:button id=""GoToTrace_Ribbon"" label=""Back to TRACE"" imageMso=""OutlinePromoteToHeading"" onAction=""GoToTrace.GoToTrace"" visible=""true""/>"
    xmlText = xmlText & "<mso:button id=""DeleteAllSheets"" label=""Delete All Sheets"" imageMso=""CancelRequest"" onAction=""InitializeWorkBook.deleteAllSheets"" visible=""true""/>"
    xmlText = xmlText & "</mso:group>"
    xmlText = xmlText & "<mso:group id=""mso_c1.4CB5718"" label=""DEVELOPMENT"" autoScale=""true"">"
    xmlText = xmlText & "<mso:button id=""ExportCode"" label=""Export Code"" imageMso=""SourceControlCheckOut"" onAction=""VersionControl.exportCode"" visible=""true""/>"
    xmlText = xmlText & "<mso:button id=""ImportCode"" label=""Import Code"" imageMso=""SourceControlCheckIn"" onAction=""VersionControl.importCode"" visible=""true""/>"
    xmlText = xmlText & "<mso:button id=""ShowSample"" label=""Show Sample"" imageMso=""SignatureShow"" onAction=""SampleVisibility.ShowSample"" visible=""true""/>"
    xmlText = xmlText & "<mso:button id=""HideSample"" label=""Hide Sample"" imageMso=""SlideMasterMediaPlaceholderInsert"" onAction=""SampleVisibility.HideSample"" visible=""true""/>"
    xmlText = xmlText & "<mso:button id=""ResetRibbons"" label=""Reset Menu Tabs"" imageMso=""ControlsGallery"" onAction=""CustomRibbons.ResetRibbons"" visible=""true""/>"
    xmlText = xmlText & "</mso:group>"
    xmlText = xmlText & "<mso:group id=""mso_c1.4CB5719"" label=""Help"" autoScale=""true"">"
    xmlText = xmlText & "<mso:button id=""Help"" label=""Help"" imageMso=""RmsInvokeBrowser"" onAction=""CustomRibbons.OpenOneNoteHelp"" visible=""true""/>"
    xmlText = xmlText & "</mso:group>"
    xmlText = xmlText & "</mso:tab>"
    xmlText = xmlText & "</mso:tabs>"
    xmlText = xmlText & "</mso:ribbon>"
    xmlText = xmlText & "</mso:customUI>"
    CVribbons_xml = xmlText
    
End Function
    

'------------------------------Default Ribbons------------------------------
'Function Name:defaultRibbons_xml
'Description: This function has the porpouse to return the string that compose the default Excel UI config file
'Inputs: ---;
'Output: String with the text that will compose the file
'-----------------------------------------------------------------------------------
Function defaultRibbons_xml() As String
    Dim xmlText As String
    xmlText = "<mso:customUI xmlns:mso='http://schemas.microsoft.com/office/2009/07/customui'>"
    xmlText = xmlText & "<mso:ribbon>"
    xmlText = xmlText & "<mso:qat>"
    xmlText = xmlText & "<mso:sharedControls>"
    xmlText = xmlText & "<mso:control idQ=""mso:AutoSaveSwitch"" visible=""true""/>"
    xmlText = xmlText & "<mso:control idQ=""mso:FileNewDefault"" visible=""false""/>"
    xmlText = xmlText & "<mso:control idQ=""mso:FileOpenUsingBackstage"" visible=""false""/>"
    xmlText = xmlText & "<mso:control idQ=""mso:FileSave"" visible=""true""/>"
    xmlText = xmlText & "<mso:control idQ=""mso:FileSendAsAttachment"" visible=""false""/>"
    xmlText = xmlText & "<mso:control idQ=""mso:FilePrintQuick"" visible=""false""/>"
    xmlText = xmlText & "<mso:control idQ=""mso:PrintPreviewAndPrint"" visible=""false""/>"
    xmlText = xmlText & "<mso:control idQ=""mso:Spelling"" visible=""false""/>"
    xmlText = xmlText & "<mso:control idQ=""mso:Undo"" visible=""true""/>"
    xmlText = xmlText & "<mso:control idQ=""mso:Redo"" visible=""true""/>"
    xmlText = xmlText & "<mso:control idQ=""mso:SortAscendingExcel"" visible=""false""/>"
    xmlText = xmlText & "<mso:control idQ=""mso:SortDescendingExcel"" visible=""false""/>"
    xmlText = xmlText & "<mso:control idQ=""mso:PointerModeOptions"" visible=""false""/>"
    xmlText = xmlText & "<mso:control idQ=""mso:VisualBasic"" visible=""true""/>"
    xmlText = xmlText & "</mso:sharedControls>"
    xmlText = xmlText & "</mso:qat>"
    xmlText = xmlText & "<mso:tabs>"
    xmlText = xmlText & "<mso:tab idQ=""mso:TabDrawInk"" visible=""true""/>"
    xmlText = xmlText & "<mso:tab idQ=""mso:TabBackgroundRemoval"" visible=""true""/>"
    xmlText = xmlText & "<mso:tab idQ=""mso:TabHome"" visible=""true""/>"
    xmlText = xmlText & "<mso:tab idQ=""mso:TabInsert"" visible=""true""/>"
    xmlText = xmlText & "<mso:tab idQ=""mso:TabPageLayoutExcel"" visible=""true""/>"
    xmlText = xmlText & "<mso:tab idQ=""mso:TabFormulas"" visible=""true""/>"
    xmlText = xmlText & "<mso:tab idQ=""mso:TabData"" visible=""true""/>"
    xmlText = xmlText & "<mso:tab idQ=""mso:TabReview"" visible=""true""/>"
    xmlText = xmlText & "<mso:tab idQ=""mso:TabView"" visible=""true""/>"
    xmlText = xmlText & "<mso:tab idQ=""mso:TabAutomate"" visible=""true""/>"
    xmlText = xmlText & "<mso:tab idQ=""mso:HelpTab"" visible=""true""/>"
    xmlText = xmlText & "<mso:tab idQ=""mso:TabAddIns"" visible=""true""/>"
    xmlText = xmlText & "<mso:tab idQ=""mso:TabDeveloper"" visible=""true""/>"
    xmlText = xmlText & "</mso:tabs>"
    xmlText = xmlText & "</mso:ribbon>"
    xmlText = xmlText & "</mso:customUI>"
    defaultRibbons_xml = xmlText
End Function

Sub ResetRibbons()
    sbDeleteFile
    checkBadClosure
End Sub

Sub OpenOneNoteHelp()
    'Keyboard Shortcut: Ctrl+Shift+W
    On Error GoTo ErrorHandler
    ActiveWorkbook.FollowHyperlink "https://uasc-my.sharepoint.com/personal/luis_schabarum_universalavionics_com/_layouts/OneNote.aspx?id=%2Fpersonal%2Fluis_schabarum_universalavionics_com%2FDocuments%2FAEL%20Wiki&wd=target%28EXCEL.one%7C1E98C166-2497-4C23-930A-0C299D81A9C3%2F%29onenote:https://uasc-my.sharepoint.com/personal/luis_schabarum_universalavionics_com/Documents/AEL%20Wiki/EXCEL.one#section-id={1E98C166-2497-4C23-930A-0C299D81A9C3}&end", NewWindow:=True
    Exit Sub
ErrorHandler:
    MsgBox "Can't open OneNote"
End Sub
