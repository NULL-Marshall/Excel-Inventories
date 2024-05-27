VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ModuleForm 
   Caption         =   "UserForm1"
   ClientHeight    =   1650
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4800
   OleObjectBlob   =   "ModuleForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ModuleForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnYes_Click()
    Me.Tag = "Yes"
    Me.Hide
End Sub

Private Sub btnNo_Click()
    Me.Tag = "Never"
    Me.Hide
End Sub

Private Sub btnLater_Click()
    Me.Tag = "Later"
    Me.Hide
End Sub

