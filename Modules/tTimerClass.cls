VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "tTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This Class is used to display a timer value formated as "nn:ss"
'Value is updated every 1 second, does not update while Excel is
'processing other code or if user is actively working in an open
'Excel file but will update once Excel has available resources
'REQIREMENTS: global instance of the created tTimer object
'REQIREMENTS: sub timer that calls global tTimer.runTimer
'Example Module
'Public t As New tTimer
'Sub test()
'    t.beginTimer "END", TimeValue("00:01:00"), ThisWorkbook.Worksheets(1).Range("A1")
'    End Sub
'
'Sub timer()
'    t.runTimer
'End Sub

'Written by: Jason Sikorski

Private xtestTime As Date 'Variable to store the amount of time for the timer
Private xendPrompt As String 'Variable for end message
Private xdest As Range 'Range Variable for location of timer output

Private Sub Class_Initialize()
    Set xdest = ThisWorkbook.Worksheets(1).Range("A1")
    xendPrompt = vbNullString
    xtestTime = TimeValue("00:10:00")
End Sub

Public Sub beginTimer(e As String, tt As Date, dt As Range) ' Call to begin Timer once variables have been set
'If variables not set default values are"
'dest Sheet1 Cell A1
'endPrompt = Empty
'testTime = 10 minutes
    xstartTime = Now
    xendPrompt = e
    Set xdest = dt
    xtestTime = tt + Now
    Application.OnTime Now + TimeValue("00:00:01"), "timer"
End Sub

Public Sub runTimer() 'Call in timer method in Module inside workbook
    If xtestTime - Now > 0 Then
        xdest = Format(xtestTime - Now, " n:ss")
        Application.OnTime Now + TimeValue("00:00:01"), "timer"
    Else
        endTimer
    End If
End Sub

Public Sub endTimer()
    If xendPrompt = vbNullString Then
    Else
        xdest = Format(TimeValue("00:00:00"), " n:ss")
        MsgBox (xendPrompt)
    End If
End Sub
