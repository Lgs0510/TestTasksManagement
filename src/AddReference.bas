Attribute VB_Name = "AddReference"
Sub AddRef(wbk As Workbook, sGuid As String, sRefName As String, Optional varMajor As Variant, Optional varMinor As Variant)
    Dim i As Integer
    On Error GoTo EH
    With wbk.VBProject.References
        If IsMissing(varMajor) Or IsMissing(varMinor) Then
           For i = 1 To .count
               If .item(i).Name = sRefName Then
                  Exit For
               End If
           Next i
           If i > .count Then
              .AddFromGuid sGuid, 0, 0 ' 0,0 should pick the latest version installed on the computer
           End If
        Else
           For i = 1 To .count
               If .item(i).GUID = sGuid Then
                  If .item(i).Major = varMajor And .item(i).Minor = varMinor Then
                     Exit For
                  Else
                     If vbYes = MsgBox(.item(i).Name & " v. " & .item(i).Major & "." & .item(i).Minor & " is currently installed," & vbCrLf & "do you want to replace it with v. " & varMajor & "." & varMinor, vbQuestion + vbYesNo, "Reference already exists") Then
                        DelRef wbk, sGuid
                     Else
                        i = 0
                        Exit For
                     End If
                  End If
               End If
           Next i
           If i > .count Then
              .AddFromGuid sGuid, varMajor, varMinor
           End If
        End If
    End With
EX: Exit Sub
EH: MsgBox "Error in 'AddRef' for guid:" & sGuid & " " & vbCrLf & vbCrLf & Err.Description
    Resume EX
    Resume ' debug code
End Sub

Public Sub DelRef(wbk As Workbook, sGuid As String)
    Dim oRef As Object
    For Each oRef In wbk.VBProject.References
        If oRef.GUID = sGuid Then
           Debug.Print "The reference to " & oRef.FullPath & " was removed."
           Call wbk.VBProject.References.Remove(oRef)
        End If
    Next
End Sub

Public Sub DebugPrintExistingRefsWithVersion()
    Dim i As Integer
    With Application.ThisWorkbook.VBProject.References
        For i = 1 To .count
            Debug.Print "   'AddRef wbk, """ & .item(i).GUID & """, """ & .item(i).Name & """" & Space(30 - Len("" & .item(i).Name)) & " ' install the latest version"
            Debug.Print "    AddRef wbk, """ & .item(i).GUID & """, """ & .item(i).Name & """, " & .item(i).Major & ", " & .item(i).Minor & Space(30 - Len(", " & .item(i).Major & ", " & .item(i).Minor) - Len("" & .item(i).Name)) & " ' install v. " & .item(i).Major & "." & .item(i).Minor
        Next i
    End With
End Sub

