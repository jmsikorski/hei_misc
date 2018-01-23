VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "snake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private s_x As Integer
Private s_y As Integer
Private s_heading As Integer
Private s_speed As Integer

Public Property Let x(x As Integer)
    s_x = x
End Property

Public Property Let y(y As Integer)
    s_y = y
End Property

Public Property Let heading(h As Integer)
    s_heading = h
End Property

Public Property Let speed(s As Integer)
    s_speed = s
End Property

Public Property Get x() As Integer
    x = s_x
End Property

Public Property Get y() As Integer
    y = s_y
End Property

Public Property Get heading() As Integer
    heading = s_heading
End Property

Public Property Get speed() As Integer
    speed = s_speed
End Property

