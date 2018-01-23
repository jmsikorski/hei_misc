VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} pjSuperPkt 
   Caption         =   "Add Leads"
   ClientHeight    =   6780
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5625
   OleObjectBlob   =   "pjSuperPkt2.frx":0000
End
Attribute VB_Name = "pjSuperPkt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lCnt As Integer

Private Sub UserForm_Initialize()
    Me.Label2.Caption = job & vbNewLine & Format(week, "mm-dd-yy")
    For i = 1 To 9
        Me.Controls("L" & i).Caption = "LEAD #0" & i
    Next i
    Me.Controls("L10").Caption = "LEAD #10"
    lCnt = 10
End Sub
