VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    On Error GoTo 0
    Dim xFile As String
    If DateDiff("h", Me.Worksheets("ADD NEW PHASE CODE").Range("updated"), Now()) > 1 Then
        xFile = Me.path & "\" & Me.name
        SetAttr xFile, vbNormal
        update_file
        Me.SaveAs xFile
    End If
    Me.Worksheets("Roster").Activate
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub
End Sub

Private Function resize_name_range(name As String, ws As Worksheet, c1 As Range, c2 As Range) As Integer
    On Error GoTo 10
    Dim wb As Workbook
    Dim nr As name
    Dim rng As Range
    Set wb = ThisWorkbook
    Set nr = wb.Names.Item(name)
    Set rng = ws.Range(c1, c2)
    nr.RefersTo = rng
    resize_name_range = 1
    Exit Function
10:
    resize_name_range = -1
    On Error GoTo 0
End Function

Private Sub update_file()
    On Error GoTo 0
    Dim phase_wb As Workbook
    Dim ws As Worksheet
    Dim xlFile As String
    Dim rng As Range
    On Error GoTo 10
    open_phase_code.update_phase_code

    With Me.Worksheets("ADD NEW PHASE CODE")
        .Unprotect
        .Range("updated") = Now()
        .Protect
    End With
    Exit Sub
10:
    MsgBox "Unable to update phase codes", vbExclamation, "ERROR"
    Me.Worksheets("Roster").Activate
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub
End Sub

