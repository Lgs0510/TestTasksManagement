VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectFolder 
   Caption         =   "Select Folder"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11550
   OleObjectBlob   =   "SelectFolder.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SelectFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub selectFolder_textbox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    selectFolder_textbox = File_op.openFolderDialog
End Sub

Private Sub selectFolder_CancelBtn_Click()
    DataComm.dataChannel.letArray = ""
    Unload Me
End Sub

Private Sub selectFolder_OkBtn_Click()
    DataComm.dataChannel.letArray = selectFolder_textbox
    Unload Me
End Sub

