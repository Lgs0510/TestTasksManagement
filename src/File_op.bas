Attribute VB_Name = "File_op"
'--------------------------------------------------------
'-------------------- Public Types ----------------------
'--------------------------------------------------------
Public Type textFile_t
    textFile As TextStream
    numberOfLines As Integer
End Type

'--------------------------------------------------------
'------------------ Public Functions -------------------
'--------------------------------------------------------


'----------------------------------Open File Dialog---------------------------------
'Function Name:openFileDialog
'Description: This function is responsible open the file dialog so the user can select/enter
'             the path of a file
'Inputs: ---
'Output: String with the complete path of the selected file
'-----------------------------------------------------------------------------------
Public Function openFileDialog() As String
    Dim fd As Office.FileDialog

    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd

      .AllowMultiSelect = False

      ' Set the title of the dialog box.
      .Title = "Please select the file."

      ' Clear out the current filters, and add our own.
      .Filters.Clear
      .Filters.Add "All Files", "*.*"

      ' Show the dialog box. If the .Show method returns True, the
      ' user picked at least one file. If the .Show method returns
      ' False, the user clicked Cancel.
        If .Show = True Then
            txtFileName = .SelectedItems(1) 'replace txtFileName with your textbox
        End If
    End With
    If txtFileName <> "" Then
        openFileDialog = txtFileName
    End If
End Function
'----------------------------------Open Folder Dialog---------------------------------
'Function Name:openFolderDialog
'Description: This function is responsible open the folder dialog so the user can select/enter
'             the path of a folder
'Inputs: ---
'Output: String with the complete path of the selected folder
'-----------------------------------------------------------------------------------
Public Function openFolderDialog() As String
    Dim fd As Office.FileDialog

    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "Select a folder"
    fd.InitialFileName = "C:\"
    If fd.Show = True Then
        openFolderDialog = fd.SelectedItems(1) 'replace txtFolderName with your textbox
    End If
End Function
'----------------------------------Get File Path---------------------------------
'Function Name:getFilePath
'Description: This function is responsible open the folder dialog so the user can select/enter
'             the path of a folder
'Inputs: ---
'Output: String with the complete path of the selected file
'-----------------------------------------------------------------------------------
Private Function getFilePath() As String
    Dim FilePath As String
    
    ImportFile.Show
    FilePath = DataComm.dataChannel.getData
    getFilePath = FilePath

End Function
'----------------------------------Get Text File----------------------------------
'Function Name:getTextFile
'Description: This function is responsible open the folder dialog so the user can select/enter
'             the path of a folder
'Inputs: ---
'Output: textFile_t object with the TextStream and the number of lines
'-----------------------------------------------------------------------------------
Function getTextFile() As textFile_t
    Dim objTextFile As textFile_t
    FilePath = getFilePath
    If FilePath <> "" Then
        Set auxFile = CreateObject("Scripting.FileSystemObject").OpenTextFile(FilePath, 1, True)
        Do Until (auxFile.AtEndOfStream Or bDoneStatus)
            strLine = auxFile.ReadLine
        Loop
        objTextFile.numberOfLines = auxFile.Line
        Set objTextFile.textFile = CreateObject("Scripting.FileSystemObject").OpenTextFile(FilePath, 1, True)
        getTextFile = objTextFile
    End If
End Function

'----------------------------------Get File Path----------------------------------
'Function Name:getFilePath
'Description: This function is responsible open the folder dialog so the user can select/enter
'             the path of a folder
'Inputs: ---
'Output: String - path of File selected
'-----------------------------------------------------------------------------------
Function getGenFilePath() As String
    Dim FilePath As String
    
    FilePath = getFilePath
    getGenFilePath = FilePath
End Function

'----------------------------------Get Folder Path----------------------------------
'Function Name:getFolderPath
'Description: This function is responsible open the folder dialog so the user can select/enter
'             the path of a folder
'Inputs: ---
'Output: String - path of Folder selected
'-----------------------------------------------------------------------------------
Function getGenFolderPath() As String
    Dim FolderPath As String
    
    getGenFolderPath = openFolderDialog
End Function


