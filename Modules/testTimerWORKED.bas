Attribute VB_Name = "testTimer"
'This module is used to display a timer value formated as "nn:ss"
'Value is updated every 1 second, does not update while Excel is
'processing other code or if user is actively working in an open
'Excel file but will update once Excel has available resources
'Written by: Jason Sikorski

Global testTime As Date 'Variable to store the amount of time for the timer
Global startTime As Date 'Variable to store the time timer started
Global endPrompt As String 'Variable for end message
Global dest As Range 'Range Variable for location of timer output

'Call beginTimer to
'Call endTimer in Workbook_close subroutine if testTime - Now > 0

Sub test()
    Dim d As Range
    Dim e As String
    e = "Times Up!"
    Set d = ThisWorkbook.Worksheets(1).Range("A3")
    Dim t As Date
    t = TimeValue("00:00:10")
    beginTimer t, d, e
End Sub

Sub timer()
    dest = Format(testTime - Now, " n:ss")
    If testTime - Now > 0 Then
        Application.OnTime Now + TimeValue("00:00:01"), "timer"
    Else
        endTimer
    End If
End Sub

Sub beginTimer(t As Date, d As Range, e As String)
    endPrompt = e
    startTime = Now
    Set dest = d
    testTime = startTime + t
    Application.OnTime Now + TimeValue("00:00:01"), "timer"
End Sub

Sub endTimer()
    dest = Format(TimeValue("00:00:00"), " n:ss")
    MsgBox (endPrompt)
End Sub
