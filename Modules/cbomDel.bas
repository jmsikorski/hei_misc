Attribute VB_Name = "Module11"
Sub del()
Attribute del.VB_ProcData.VB_Invoke_Func = "D\n14"
    Dim x As Integer
    x = 0
    For i = 0 To 300
        If IsEmpty(ActiveSheet.Range("B7").Offset(x, 0)) Then
            ActiveSheet.Rows(Range("B7").Offset(x, 0).Row).Delete
        Else
            x = x + 1
        End If
    Next i
End Sub
