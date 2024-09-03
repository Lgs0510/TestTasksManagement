VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ImportFile 
   Caption         =   "Import File"
   ClientHeight    =   3012
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11520
   OleObjectBlob   =   "ImportFile.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ImportFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub importCsv_CancelBtn_Click()
    DataComm.dataChannel.letArray = ""
    Unload Me
End Sub

Private Sub importCsv_OkBtn_Click()
    DataComm.dataChannel.letArray = importCsv_textbox
    Unload Me
End Sub


Private Sub importCsv_textbox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    importCsv_textbox = File_op.openFileDialog
End Sub

Private Sub Label1_Click()

End Sub
