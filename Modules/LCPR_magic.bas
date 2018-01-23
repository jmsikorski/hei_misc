Attribute VB_Name = "Module1"
Private Function open_file() As String
    On Error GoTo 10
    Dim strFileToOpen As Variant
    Dim f As Variant
    Dim wb As Workbook
1:
    strFileToOpen = Application.GetOpenFilename _
    (Title:="Please select file to import", _
    FileFilter:="Excel Files *.xls* (*.xls*),", _
    MultiSelect:=False)
    If Not IsArray(strFileToOpen) Then
        If strFileToOpen = False Then
            MsgBox "No file selected.", vbExclamation, "Sorry!"
            GoTo 1
        Else
            Set wb = Workbooks.Open(Filename:=strFileToOpen)
            open_file = wb.Name
            Exit Function
        End If
    Else
        For Each f In strFileToOpen
            Dim tmp() As String
            Set wb = Workbooks.Open(Filename:=strFileToOpen)
            open_file = wb.Name
            Exit Function
        Next f
    End If
    open_file = "ERROR"
    Exit Function
10:
    open_file = "ERROR"
    On Error GoTo 0
test:
    Workbooks.Open "C:\Users\jsikorski\Desktop\461705 LCPR.xls"
    open_file = "461705 LCPR.xls"
End Function

Public Sub new_report()
    Application.ScreenUpdating = False
    Dim ogwb As Workbook
    Set ogwb = ThisWorkbook
    Dim wb As Workbook
    Dim rng As Range
    Dim ws As Worksheet
    Set wb = Workbooks(open_file)
    Set ws = wb.Worksheets(1)
    Set rng = ws.Range("A1")
    Do While rng <> "Description"
        Set rng = rng.Offset(0, 1)
    Loop
    Set rng = ws.Range(ws.Range("A1"), rng.End(xlDown))
    rng.Delete
    Set rng = ws.Range("U1", ws.Range("U1").End(xlToRight))
    rng.Copy
    ws.Range("A1").End(xlDown).Offset(2, 0).PasteSpecial xlPasteAll
    Application.CutCopyMode = False
    ws.Range(rng, rng.End(xlToRight)).EntireColumn.Delete
    ws.Range("A1").EntireRow.Insert
    ws.Range("A1").EntireRow.Insert
    Set rng = ws.Range("C1")
    With ws.Range(rng, rng.Offset(0, 5))
        .Select
        .MergeCells = True
        .Value = "BUDGET"
    End With
    Set rng = ws.Range("I1")
    With ws.Range(rng, rng.Offset(0, 5))
        .Select
        .MergeCells = True
        .Value = "JOB TO DATE"
    End With
    Set rng = ws.Range("O1")
    With ws.Range(rng, rng.Offset(0, 5))
        .Select
        .MergeCells = True
        .Value = "PERIOD"
    End With
    Set rng = ws.Range("U1")
    With ws.Range(rng, rng.Offset(0, 5))
        .Select
        .MergeCells = True
        .Value = "REMAINING"
    End With
    Set rng = ws.Range("A2")
    With rng
        .Select
        .Value = "Phase"
        .Offset(0, 1).Value = "Description"
        .Offset(0, 2).Value = "Units"
        .Offset(0, 3).Value = "Hours"
        .Offset(0, 4).Value = "Cost"
        .Offset(0, 5).Value = "Hours/Unit"
        .Offset(0, 6).Value = "Unit Cost"
        .Offset(0, 7).Value = "Units/Hour"
    End With
        Set rng = rng.Offset(0, 7)
    With rng
        .Offset(0, 1).Value = "Units"
        .Offset(0, 2).Value = "Hours"
        .Offset(0, 3).Value = "Cost"
        .Offset(0, 4).Value = "Hours/Unit"
        .Offset(0, 5).Value = "Unit Cost"
        .Offset(0, 6).Value = "Units/Hour"
    End With
        Set rng = rng.Offset(0, 6)
    With rng
        .Offset(0, 1).Value = "Units"
        .Offset(0, 2).Value = "Hours"
        .Offset(0, 3).Value = "Cost"
        .Offset(0, 4).Value = "Hours/Unit"
        .Offset(0, 5).Value = "Unit Cost"
        .Offset(0, 6).Value = "Units/Hour"
    End With
        Set rng = rng.Offset(0, 6)
    With rng
        .Offset(0, 1).Value = "Units"
        .Offset(0, 2).Value = "Hours"
        .Offset(0, 3).Value = "Cost"
        .Offset(0, 4).Value = "Units/Hour"
        .Offset(0, 5).Value = "JTD Diff"
        .Offset(0, 6).Value = "EST CTC"
        .Offset(0, 7).Value = "BUD DIFF"
    End With
    For Each rng In ws.Range("A3", ws.Range("A3").End(xlDown))
        With rng
            .Offset(0, 20).Formula = "=IF(C" & .Row & "-I" & .Row & ">0,C" & .Row & "-I" & .Row & ",""OVER"")"
            .Offset(0, 21).Formula = "=IF(D" & .Row & "-J" & .Row & ">0,D" & .Row & "-J" & .Row & ",""OVER"")"
            .Offset(0, 22).Formula = "=IF(E" & .Row & "-K" & .Row & ">0,E" & .Row & "-K" & .Row & ",CONCATENATE(""+ $"",ROUND(K" & .Row & "-E" & .Row & ",2)))"
            .Offset(0, 23).Formula = "=IF(AND(U" & .Row & "<>""OVER"",V" & .Row & "<>""OVER""),ROUND(U" & .Row & "/V" & .Row & ",2),""N/A"")"
            .Offset(0, 24).Formula = "=IFERROR(X" & .Row & "-N" & .Row & ","""")"
            .Offset(0, 25).Formula = "=IFERROR(U" & .Row & "*M" & .Row & ","""")"
            .Offset(0, 26).Formula = "=IFERROR(E" & .Row & "-K" & .Row & "-Z" & .Row & ","""")"
        End With
    Next
    Application.ScreenUpdating = False
    ogwb.Close False
End Sub
