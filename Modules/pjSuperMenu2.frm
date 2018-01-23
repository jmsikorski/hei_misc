VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} pjSuperMenu2 
   Caption         =   "Superintendent Menu"
   ClientHeight    =   3255
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5565
   OleObjectBlob   =   "pjSuperMenu2.frx":0000
End
Attribute VB_Name = "pjSuperMenu2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    Me.Label2.Caption = job & vbNewLine & Format(week, "mm-dd-yy")
End Sub
