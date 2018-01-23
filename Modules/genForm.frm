VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} genForm 
   Caption         =   "Select Job Number"
   ClientHeight    =   4440
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8055
   OleObjectBlob   =   "genForm.frx":0000
End
Attribute VB_Name = "genForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton1_Click()

End Sub

Private Sub UserForm_Initialize()
    Dim cJob As Range
    Set cJob = Worksheets("JOBS").UsedRange
    For Each cJob In Worksheets("JOBS").Range("jobList")
      With Me.ComboBox1
        .AddItem cJob.Value
        .list(.ListCount - 1, 1) = cJob.Offset(0, 1).Value
      End With
    Next cJob
End Sub

