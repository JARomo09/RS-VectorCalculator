VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "R-S Vector Calculator"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4515
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub CommandButton1_Click()
    Module1.R_add_item
End Sub

Private Sub CommandButton2_Click()
    Module1.R_remove_item
End Sub

Private Sub CommandButton3_Click()
    Module1.S_add_item
End Sub

Private Sub CommandButton4_Click()
    Module1.S_remove_item
End Sub

Private Sub CommandButton5_Click()
    Module1.clear_all
End Sub
