VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Employee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private fName As String
Private lName As String
Private num As Integer
Private eClass As String
Private shifts As Collection

Public Property Let efName(s As String)
    fName = s
End Property

Public Property Let elName(s As String)
    lName = s
End Property

Public Property Let emnum(s As Integer)
    num = s
End Property

Public Property Let emClass(s As String)
    eClass = s
End Property

Public Sub addShift(s As shift)
    shifts.Add s
End Sub

Public Property Get getFName() As String
    getFName = fName
End Property

Public Property Get getLName() As String
    getLName = lName
End Property

Public Property Get getNum() As Integer
    getNum = num
End Property

Public Property Get getClass() As String
    getClass = eClass
End Property

Public Property Get getShifts() As Collection
    Set getShifts = shifts
End Property

Private Sub Class_Initialize()
    fName = "NONE"
    lName = "NONE"
    num = -1
    eClass = "NONE"
    Set shifts = New Collection
End Sub




