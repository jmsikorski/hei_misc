VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "update_emp_table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Enum state
    open_phase = 1
    close_phase = 2
    update_phase = 3
End Enum

Private Const pw = ""
    
Private Sub Workbook_Open()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Dim rng As Range
    Dim ws As Worksheet
    Dim xFile As String
    Set ws = ThisWorkbook.Worksheets("Open Phase Codes")
    ws.Unprotect
    If DateDiff("h", Me.Worksheets("instructions").Range("updated"), Now()) > 1 Then
        xFile = Me.path & "\" & Me.name
        SetAttr xFile, vbNormal
        update_phase_code
        Set rng = ws.ListObjects("emp_roster").DataBodyRange(ws.ListObjects("emp_roster").DataBodyRange.Rows.count, 1)
        resize_name_range "open_codes", ws, ws.ListObjects("emp_roster").DataBodyRange(1, 1).Offset(0, 2), rng.Offset(0, 2)
        Me.SaveAs xFile
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
    End If
    ws.Protect
End Sub

Public Sub close_phase_code()
Attribute close_phase_code.VB_Description = "Close phase code and delete it from the list\n"
Attribute close_phase_code.VB_ProcData.VB_Invoke_Func = "D\n14"
    'on error goto 10
    Dim new_code As Double
    Dim new_desc As String
    Dim rng As Range
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Open Phase Codes")
    Set rng = ws.ListObjects("emp_roster").Range(1, 1)
    new_code = get_code(close_phase)
    For Each rng In ws.ListObjects("emp_roster").ListColumns(1).DataBodyRange
        If rng.Value = new_code Then
            ws.Unprotect pw
            ws.Rows(rng.Row).EntireRow.Delete
            ws.Protect pw
            Exit Sub
        End If
    Next rng
    MsgBox "Phase Code does not exist", vbExclamation, "ERROR!"
    Exit Sub
10:
    MsgBox "Error: Unable to close Phase Code", vbExclamation, "ERROR!"
End Sub

Public Sub update_emp_table()
Attribute update_emp_table.VB_Description = "Update Open phase codes from Labor Report\n"
Attribute update_emp_table.VB_ProcData.VB_Invoke_Func = "U\n14"
    Application.ScreenUpdating = False
    'on error goto 10
    Dim new_emp As Range
    Dim rng As Range
    Dim ws As Worksheet
    Dim cnt As Integer
    cnt = 1
    Set ws = ThisWorkbook.Worksheets("Roster")
    ws.Unprotect pw
'    ws.Range("A2", ws.Range("A1").End(xlDown).Offset(0, 1)).Clear
    ws.ListObjects("emp_roster").DataBodyRange.Clear
'    Set rng = ws.Range("A2")
    For Each rng In ws.ListObjects("emp_roster").DataBodyRange
1:
        new_emp = get_emp(cnt)
        If new_emp = Nothing Then
            cnt = cnt + 1
            GoTo 1
        End If
        For Each rng In ws.ListObjects("emp_roster").ListColumns(1).DataBodyRange
            If rng.Value = vbNullString Then
                GoTo 5
            End If
            If rng.Value = new_emp Then
                cnt = cnt + 1
                GoTo 1
            End If
        Next rng
5:
        If insert_emp(new_emp) = -1 Then
            GoTo 20
        End If
        cnt = cnt + 1
    Loop
    On Error GoTo 0
    Set rng = ws.Range(ws.ListObjects("emp_roster").Range(1, 1), ws.ListObjects("emp_roster").Range(1, 2).End(xlDown))
    ws.ListObjects("emp_roster").Resize rng
    Workbooks("Attendance Tracking.xlsx").Close
    With Me.Worksheets("ROSTER")
        .Unprotect
        .Range("updated") = Now()
        .Protect
    End With
    Application.ScreenUpdating = True
    ws.Protect pw
    Exit Sub
10:
    Dim ans As Integer
    Dim xFile As String
    With Application.FileDialog(msoFileDialogOpen)
        .Title = "Find Attendance Tracking Roster"
        .Filters.Add "Excel Files", "*.xls*", 1
        .InitialFileName = Me.path & "\"
        ans = .Show
        If ans = 0 Then
            Exit Sub
        Else
            Workbooks.Open .SelectedItems(1)
            xFile = .SelectedItems(1)
        End If
    End With
    Set mb = Workbooks(xFile)
    Resume Next
    Exit Sub
20:
    MsgBox "ERROR: Unable to update roster", vbCritical, "ERROR!"
    On Error GoTo 0
    ws.Protect pw
    Application.ScreenUpdating = True
End Sub

Private Function get_emp(Optional cnt As Integer = 1) As Range
    Dim new_emp As Range
    Dim ans As Integer
1:
    Dim mb As Workbook
    Dim xlFile As String
    On Error GoTo 30
    Set mb = Workbooks("Attendance Tracking.xlsx")
    On Error GoTo 0
    If mb.Worksheets(1).ListObjects("emp_roster").Offset(cnt, 0).Interior.Color = 255 Then
        get_emp = Nothing
        Exit Function
    End If
    new_emp = mb.Worksheets(1).ListObjects("emp_roster").ListRows(cnt)
    Else
        get_emp = new_emp
    End If
Exit Function
30:
    xlFile = ThisWorkbook.path & "\Attendance Tracking.xlsx"
    Workbooks.Open xlFile
    Set mb = Workbooks("Attendance Tracking.xlsx")
    Resume Next
End Function


Private Function insert_emp(emp As Range, desc As String) As Integer
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rng As Range
    Set wb = ThisWorkbook
    Set ws = ActiveSheet
    For Each rng In ws.ListObjects("emp_roster").ListColumns(1).DataBodyRange 'ws.Range("A2", ws.Range("A1").End(xlDown))
        If rng.Value = emp.Cells(1, 1) Then
            GoTo 10
        End If
        If rng.Value > emp.Cells(1, 1) Then
1:
            With rng
                For i = 0 To emp.Columns.count
                .Value = emp.Cells(1, i).Value
                .Font.name = "Helvetica"
                .Font.Bold = False
                .Font.Size = 12
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Weight = xlThin
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).Weight = xlThin
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).Weight = xlThin
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlEdgeRight).Weight = xlThin
                Next i
            End With
            insert_code = 1
            Exit Function
        ElseIf rng.Value = vbNullString Then
            GoTo 1
        End If
    Next rng
    Set rng = ws.Range("A1").End(xlDown).Offset(1, 0)
    GoTo 1
10:
    insert_code = -1
    On Error GoTo 0
End Function


Private Function resize_name_range(name As String, ws As Worksheet, c1 As Range, c2 As Range) As Integer
    'on error goto 10
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
