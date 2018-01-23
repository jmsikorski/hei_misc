VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
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
    Dim rng As Range
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Open Phase Codes")
    'update_phase_code
    Set rng = ws.ListObjects("phase_list").DataBodyRange(ws.ListObjects("phase_list").DataBodyRange.Rows.Count, 1)
    resize_name_range "open_codes", ws, ws.ListObjects("phase_list").DataBodyRange(1, 1).Offset(0, 2), rng.Offset(5, 2)
End Sub

Public Sub close_phase_code()
Attribute close_phase_code.VB_Description = "Close phase code and delete it from the list\n"
Attribute close_phase_code.VB_ProcData.VB_Invoke_Func = "D\n14"
    On Error GoTo 10
    Dim new_code As Double
    Dim new_desc As String
    Dim rng As Range
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Open Phase Codes")
    Set rng = ws.ListObjects("phase_list").Range(1, 1)
    new_code = get_code(close_phase)
    For Each rng In ws.ListObjects("phase_list").ListColumns(1).DataBodyRange
        If rng.Value = new_code Then
            ws.Unprotect pw
            ws.Rows(rng.Row).EntireRow.Delete
            ws.Protect pw
            Exit Sub
        End If
    Next rng
    MsgBox "Phase Code does not exist", vbExclamation, "ERROR!"
10:
    MsgBox "Error: Unable to close Phase Code", vbExclamation, "ERROR!"
End Sub

Public Sub update_phase_code()
Attribute update_phase_code.VB_Description = "Update Open phase codes from Labor Report\n"
Attribute update_phase_code.VB_ProcData.VB_Invoke_Func = "U\n14"
    Application.ScreenUpdating = False
    On Error GoTo 10
    Dim new_code As Double
    Dim new_desc As String
    Dim rng As Range
    Dim ws As Worksheet
    Dim cnt As Integer
    cnt = 1
    Set ws = ThisWorkbook.Worksheets("Open Phase Codes")
    ws.Unprotect pw
'    ws.Range("A2", ws.Range("A1").End(xlDown).Offset(0, 1)).Clear
    ws.ListObjects("phase_list").DataBodyRange.Clear
'    Set rng = ws.Range("A2")
    Set rng = ws.ListObjects("phase_list").DataBodyRange(1, 1)
    new_code = 1
    Do While new_code <> 0
1:
        new_code = get_code(update_phase, cnt)
        If new_code = -1 Then
            GoTo 20
        ElseIf new_code = -2 Then
            cnt = cnt + 1
            GoTo 1
        ElseIf new_code = 0 Then
            Exit Do
        End If
        For Each rng In ws.ListObjects("phase_list").ListColumns(1).DataBodyRange
            If rng.Value = vbNullString Then
                GoTo 5
            End If
            If rng.Value = new_code Then
                GoTo 1
            End If
        Next rng
5:
        new_desc = get_description(update_phase, cnt)
        If new_desc = vbNullString Then
            GoTo 20
        Else
            If insert_code(new_code, new_desc) = -1 Then
                GoTo 20
            End If
            cnt = cnt + 1
        End If
    Loop
    On Error GoTo 0
'    ws.Range(ws.Cells(rng_end + 1, 1), ws.Cells(rng_end + 1000, 2)).EntireRow.Delete
    Set rng = ws.Range(ws.ListObjects("phase_list").Range(1, 1), ws.ListObjects("phase_list").Range(1, 2).End(xlDown))
    ws.ListObjects("phase_list").Resize rng
    For i = 0 To 4
        ws.ListObjects("phase_list").ListRows.Add
        For p = 1 To 2
            With ws.ListObjects("phase_list").Range(ws.ListObjects("phase_list").ListRows.Count + 1, p)
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Weight = xlThin
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).Weight = xlThin
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).Weight = xlThin
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlEdgeRight).Weight = xlThin
            End With
        Next p
    Next i
'    Dim c1 As Interior
'    Dim c2 As Interior
'    Set c1 = ws.Range("A2").Interior
'    Set c2 = ws.Range("A3").Interior
'
'    For i = 2 To rng_end
'        ws.Range("A" & i, "B" & i).Interior.Color = c1.ColorIndex
'        i = i + 1
'        ws.Range("A" & i, "B" & i).Interior.Color = c2.ColorIndex
'    Next i
'    resize_name_range "open_codes", ws, ws.Range("C2"), rng.Offset(0, 2)
    Workbooks("Labor Report.xlsx").Close
    Application.ScreenUpdating = True
    ws.Protect pw
    Exit Sub
10:
    Dim ans As Integer
    With Application.FileDialog(msoFileDialogOpen)
        .Title = "Select Labor Report"
        .Filters.Add "Excel Files", "*.xls*", 1
        .InitialFileName = Me.Path & "\"
        ans = .Show
        If ans = 0 Then
            Exit Sub
        Else
            Workbooks.Open .SelectedItems(1)
        End If
    End With
    Set mb = Workbooks("Labor Report.xlsx")
    Resume Next
20:
    MsgBox "ERROR: Unable to update phase Codes", vbCritical, "ERROR!"
    On Error GoTo 0
    ws.Protect pw
    Application.ScreenUpdating = True
End Sub

Public Sub open_phase_code()
Attribute open_phase_code.VB_Description = "Open new phase code and add it to the list"
Attribute open_phase_code.VB_ProcData.VB_Invoke_Func = "O\n14"
    On Error GoTo 10
    Dim new_code As Double
    Dim new_desc As String
    Dim rng As Range
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Open Phase Codes")
    ws.Unprotect pw
    new_code = get_code(open_phase)
    For Each rng In ws.ListObjects("phase_list").ListColumns(1).DataBodyRange
        If rng.Value = new_code Then
            MsgBox "Phase code already open!", vbExclamation, "ERROR!"
            Exit Sub
        End If
    Next rng
    If new_code = -1 Then
        GoTo 10
    Else
        new_desc = get_description(open_phase)
        If new_desc = vbNullString Then
            GoTo 10
        Else
            If insert_code(new_code, new_desc) = -1 Then
                GoTo 10
            End If
        End If
    End If
    Set rng = ws.Range("A1").End(xlDown)
    resize_name_range "open_codes", ws, ws.Range("C2"), rng.Offset(0, 2)
    ws.Protect pw
    Exit Sub
10:
    MsgBox "Unable to open phase code", vbExclamation
    On Error GoTo 0
    ws.Protect pw
End Sub

Private Function get_code(state As Integer, Optional cnt As Integer = 1) As Double
    Dim new_code As Double
    Dim ans As Integer
1:
    Select Case state
    Case open_phase
        new_code = InputBox("Enter Phase Code to Open", "Open Phase Code")
        If new_code < 0 Or new_code > 99999 Then
            MsgBox "Invalid Phase Code Entered!", vbCritical, "ERROR!"
            GoTo 1
        Else
            If new_code > 89999 Or new_code < 89000 Then
                ans = MsgBox("Unexpected Phase Code!" & vbNewLine & "Do you want to add " & new_code, vbYesNoCancel)
                If ans = vbYes Then
                    get_code = new_code
                    Exit Function
                ElseIf ans = vbCancel Then
                    get_code = -1
                Else
                    GoTo 1
                End If
            Else
                get_code = new_code
            End If
        End If
    Case close_phase
        new_code = InputBox("Enter Phase Code to Close", "Close Phase Code")
        If new_code < 0 Or new_code > 99999 Then
            MsgBox "Invalid Phase Code Entered!", vbCritical, "ERROR!"
            GoTo 1
        Else
            get_code = new_code
        End If
    Case update_phase
        Dim mb As Workbook
        Dim xlFile As String
        On Error GoTo 10
        Set mb = Workbooks("Labor Report.xlsx")
        On Error GoTo 0
        If mb.Worksheets(1).Range("C2").Offset(cnt, 0).Interior.Color = 255 Then
            get_code = -2
            Exit Function
        End If
        new_code = mb.Worksheets(1).Range("C2").Offset(cnt, 0)
        If new_code < 0 Or new_code > 99999 Then
            MsgBox "Invalid Phase Code Entered!", vbCritical, "ERROR!"
            get_code = -1
        Else
            get_code = new_code
        End If
    Case Else
        get_code = -1
    End Select
    Exit Function
10:
    xlFile = ThisWorkbook.Path & "\Labor Report.xlsx"
    Workbooks.Open xlFile
    Set mb = Workbooks("Labor Report.xlsx")
    Resume Next
End Function


Private Function get_description(state As Integer, Optional cnt As Integer = 1) As String
    Dim desc As String
    Dim ans As Integer
1:
    Select Case state
    Case open_phase
        desc = InputBox("Enter Phase Code Description", "Open Phase Code")
        If desc = vbNullString Then
            MsgBox "Description can not be empty!", vbCritical, "ERROR!"
            GoTo 1
        End If
        If Len(desc) > 50 Then
            ans = MsgBox("Description is too long!" & vbNewLine & "Do you want to add it anyway?", vbCritical + vbAbortRetryIgnore)
            If ans = vbIgnore Then
                get_description = desc
                Exit Function
            ElseIf ans = vbCancel Then
                get_description = vbNullString
            ElseIf ans = vbAbort Then
                get_description = vbNullString
            Else
                GoTo 1
            End If
        Else
            get_description = desc
        End If
    Case update_phase
        Dim mb As Workbook
        Dim xlFile As String
        On Error GoTo 10
        Set mb = Workbooks("Labor Report.xlsx")
        On Error GoTo 0
        desc = mb.Worksheets(1).Range("D2").Offset(cnt, 0)
        get_description = desc
    Case Else
        get_description = vbNullString
    End Select
    Exit Function
10:
    xlFile = ThisWorkbook.Path & "\Labor Report.xlsx"
    Workbooks.Open xlFile
    Set mb = Workbooks("Labor Report.xlsx")
    Resume Next
End Function


Private Function insert_code(code As Double, desc As String) As Integer
    On Error GoTo 10
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rng As Range
    Set wb = ThisWorkbook
    Set ws = ActiveSheet
'    Stop
    For Each rng In ws.ListObjects("phase_list").ListColumns(1).DataBodyRange 'ws.Range("A2", ws.Range("A1").End(xlDown))
        If rng.Value = code Then
            MsgBox "Phase code already open!", vbCritical, "ERROR!"
            GoTo 10
        End If
        If rng.Value > code Then
1:
            With rng
                .Value = code
                .Font.name = "Arial"
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
                .Offset(0, 1).Value = desc
                .Offset(0, 1).Font.name = "Arial"
                .Offset(0, 1).Font.Size = 12
                .Offset(0, 1).Font.Bold = False
                .Offset(0, 1).HorizontalAlignment = xlLeft
                .Offset(0, 1).Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Offset(0, 1).Borders(xlEdgeBottom).Weight = xlThin
                .Offset(0, 1).Borders(xlEdgeTop).LineStyle = xlContinuous
                .Offset(0, 1).Borders(xlEdgeTop).Weight = xlThin
                .Offset(0, 1).Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Offset(0, 1).Borders(xlEdgeLeft).Weight = xlThin
                .Offset(0, 1).Borders(xlEdgeRight).LineStyle = xlContinuous
                .Offset(0, 1).Borders(xlEdgeRight).Weight = xlThin
                .Offset(0, 2) = rng.Offset(-1, 2).Formula
                insert_code = 1
            End With
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
