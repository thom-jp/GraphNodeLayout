VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Node"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public KineticEnergy As Axes
Public Position As Axes
Public Velocity As Axes
Public ConnectedNode As Collection
Public ID As String

Private Sub Class_Initialize()
    Set ConnectedNode = New Collection
    Set KineticEnergy = New Axes
    Set Velocity = New Axes
    Set Position = New Axes
    Call Randomize
    Position.X = Rnd()
    Position.Y = Rnd()
End Sub

Public Sub Connect(n As Node)
    If n.ID = Me.ID Then Exit Sub
    On Error Resume Next
    ConnectedNode.Add n, n.ID
    On Error GoTo 0
End Sub

Public Function Connected(n As Node) As Boolean
    On Error GoTo Err_handler:
    Set X = ConnectedNode.Item(n.ID)
    On Error GoTo 0
    Connected = True
    Exit Function
Err_handler:
    Connected = False
    On Error GoTo 0
End Function
