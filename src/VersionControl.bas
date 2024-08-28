Attribute VB_Name = "VersionControl"
Sub importCode()

    projName = ThisWorkbook.VBProject.Name
    importPath = File_op.getGenFolderPath
    teste = Application.Run("testImport", projName, importPath)
End Sub

Sub exportCode()

    projName = ThisWorkbook.VBProject.Name
    exportPath = getGenFolderPath
    teste = Application.Run("testExport", projName, exportPath)

End Sub
