Attribute VB_Name = "VersionControl"
Sub importCode()

    projName = ThisWorkbook.VBProject.Name
    importPath = File_op.getGenFolderPath
    If importPath <> "" Then
        teste = Application.Run("testImport", projName, importPath)
    End If
End Sub

Sub exportCode()

    projName = ThisWorkbook.VBProject.Name
    exportPath = getGenFolderPath
    If exportPath <> "" Then
        teste = Application.Run("testExport", projName, exportPath)
    End If
End Sub
Sub importExportImportCode()

    projName = "vbaDeveloper"
    importPath = File_op.getGenFolderPath
    If importPath <> "" Then
        teste = Application.Run("testImport", projName, importPath)
    End If
End Sub

Sub exportExportImportCode()

    projName = "vbaDeveloper"
    exportPath = getGenFolderPath
    If exportPath <> "" Then
        teste = Application.Run("testExport", projName, exportPath)
    End If
End Sub

