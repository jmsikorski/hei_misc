VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "board"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private h As Integer
Private w As Integer
Private bkgrd As Integer
Private bdy As Integer
Private snake() As Boolean

Public Sub build(height As Integer, wide As Integer, background As Integer, body As Integer)
    If height > 1 Then
        h = height
    Else
        h = 1
    If wide > 1 Then
        w = wide
    Else
        w = 0
    If background > 0 Then
        bkgrd = background
    Else
        bkgrd = 0
    End If
    If body > 0 Then
        bdy = body
    Else
        body = 0
    End If
    ReDim snake(1, w)
    Me.clr
End Sub

Public Property Get height() As Integer
    height = l
End Property

Public Property Get wide() As Integer
    wide = w
End Property

Public Property Get background() As Integer
    background = bkgrd
End Property

Public Property Get body() As Integer
    body = bdy
End Property

Public Property Get is_snake(x_pos, y_pos) As Boolean
    is_snake = snake(x_pos, y_pos)
End Property

Public Function add_snake(x_pos, y_pos) As Integer
    If snake(x_pos, y_pos) = True Then
        add_snake = -1
    Else
        snake(x_pos, y_pos) = True
        add_snake = 1
    End If
End Function

Public Property Get rm_snake(x_pos, y_pos) As Integer
    If snake(x_pos, y_pos) = False Then
        rm_snake = -1
    Else
        snake(x_pos, y_pos) = False
        rm_snake = 1
    End If
End Property

Public Sub clr(x_pos As Integer, y_pos As Integer)
    For i = 0 To Me.height
        For x = 0 To Me.wide
            If i = y_pos And x = x_pos Then
                If Me.add_snake(x, y) < 1 Then
                    MsgBox ("ERROR ADDING SNAKE IN FUNCTION CLR:1")
                End If
            Else
                If Me.rm_snake(x, y) < 1 Then
                    MsgBox ("ERROR ADDING SNAKE IN FUNCTION CLR:1")
                End If
            End If
        Next x
    Next i
End Sub
