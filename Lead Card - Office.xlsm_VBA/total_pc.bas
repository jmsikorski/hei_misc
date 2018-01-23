Attribute VB_Name = "total_pc"
Private Enum state
    open_phase = 1
    close_phase = 2
    update_phase = 3
End Enum

Private Const pw = ""

Public Function total_phase_code() As Double
    Dim rng As Range
    Dim code As Double
    
    Set rng = ActiveSheet.Range(ActiveCell, ActiveCell)
    code = ActiveSheet.Cells(rng.Row, 1).Value
    total_phase_code = code
    
    
End Function

Public Function used_phase_code(code As Double) As Boolean
    For i = 1 To Worksheets("ROSTER").ListObjects("Monday").ListColumns(5).DataBodyRange.Rows.count
        If code = Left(Worksheets("ROSTER").ListObjects("Monday").DataBodyRange(i, 5), 6) Then
            used_phase_code = True
            Exit Function
        End If
    Next i
    used_phase_code = False
End Function

