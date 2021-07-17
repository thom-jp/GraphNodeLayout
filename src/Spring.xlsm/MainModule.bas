Attribute VB_Name = "MainModule"
Option Explicit
Sub Main()
    Const DELTA_T = 0.01
    Const LOSS_FACTOR = 0.99
    Const NODE_SIZE = 5
    
    Dim subjectNode As Node
    Dim objectNode As Node

    Dim Nodes As Collection
    Set Nodes = AnalyseNodes
    'Set Nodes = CreateNodes(5, 5)
    
    Do
        Dim kineticEnergyInField As Axes: Set kineticEnergyInField = New Axes
        For Each subjectNode In Nodes
            Dim receivedForceSum As Axes: Set receivedForceSum = New Axes
            
            For Each objectNode In Nodes
                If subjectNode Is objectNode Then GoTo Continue
                Dim force As Axes
                If subjectNode.Connected(objectNode) Then
                    Set force = HookeForce(subjectNode.Position, objectNode.Position)
                Else
                    Set force = CoulombForce(subjectNode.Position, objectNode.Position)
                End If
                receivedForceSum.X = receivedForceSum.X + force.X
                receivedForceSum.Y = receivedForceSum.Y + force.Y
Continue:
            Next
            subjectNode.Velocity.X = (subjectNode.Velocity.X + receivedForceSum.X * DELTA_T) * LOSS_FACTOR
            subjectNode.Velocity.Y = (subjectNode.Velocity.Y + receivedForceSum.Y * DELTA_T) * LOSS_FACTOR
            kineticEnergyInField.X = kineticEnergyInField.X + subjectNode.Velocity.X ^ 2
            kineticEnergyInField.Y = kineticEnergyInField.Y + subjectNode.Velocity.Y ^ 2
        Next
        
        For Each subjectNode In Nodes
            subjectNode.Position.X = subjectNode.Position.X + (subjectNode.Velocity.X * DELTA_T)
            subjectNode.Position.Y = subjectNode.Position.Y + (subjectNode.Velocity.Y * DELTA_T)
        Next
    Loop Until 0.0000005 > Round(Sqr(kineticEnergyInField.X ^ 2 + kineticEnergyInField.Y ^ 2), 7)
    
    'DrawSheet.ClearAllShapes

    'PlotOvals
    For Each subjectNode In Nodes
        Dim PosX As Single: PosX = subjectNode.Position.X * 1000
        Dim PosY As Single: PosY = subjectNode.Position.Y * 1000
        subjectNode.NodeShape.Left = PosX
        subjectNode.NodeShape.Top = PosY
    Next

    Call AdjustConnectors
    
End Sub

Sub AdjustConnectors()
    Dim sh As Shape, a As String, b As String
    For Each sh In Sheet1.Shapes
        If sh.Type = msoAutoShape And sh.AutoShapeType = -2 Then
            If sh.ConnectorFormat.BeginConnected And sh.ConnectorFormat.EndConnected Then
            Call ConnectOvals( _
                  sh.ConnectorFormat.BeginConnectedShape, _
                  sh.ConnectorFormat.EndConnectedShape, _
                  sh)
            End If
        End If
    Next
End Sub

Function AnalyseNodes() As Collection
    Dim ret As Collection: Set ret = New Collection
    Dim ov As Oval
    For Each ov In Sheet1.Ovals
        With New Node
            .ID = ov.Name
            Set .NodeShape = ov.ShapeRange(1)
            ret.Add .Self, .ID
        End With
    Next
    
    Dim sh As Shape, a As String, b As String
    For Each sh In Sheet1.Shapes
        If sh.Type = msoAutoShape And sh.AutoShapeType = -2 Then
            If sh.ConnectorFormat.BeginConnected And sh.ConnectorFormat.EndConnected Then
                a = sh.ConnectorFormat.BeginConnectedShape.Name
                b = sh.ConnectorFormat.EndConnectedShape.Name
                ret(a).Connect ret(b)
                ret(b).Connect ret(a)
            End If
        End If
    Next
    
    Set AnalyseNodes = ret
End Function

Function CreateNodes(number_of_rows As Integer, number_of_columns As Integer) As Collection
    Dim ret As Collection
    Set ret = New Collection
    
    Dim sbjIndex As Integer
    For sbjIndex = 0 To number_of_rows * number_of_columns - 1
        ret.Add New Node
        ret(ret.Count).ID = sbjIndex
    Next

    'Node Connection
    For sbjIndex = 0 To number_of_rows * number_of_columns - 1
        Dim sbjColumn: sbjColumn _
            = sbjIndex Mod number_of_columns
            
        Dim sbjRow: sbjRow _
            = sbjIndex \ number_of_columns
        
        Dim objIndex As Integer
        
        For objIndex = 0 To number_of_rows * number_of_columns - 1
            Dim objColumn: objColumn _
                = objIndex Mod number_of_columns
            
            Dim objRow: objRow _
                = objIndex \ number_of_columns
            
            
            Dim sameColumn As Boolean: sameColumn _
                = (sbjColumn = objColumn)
            
            Dim sameRow As Boolean: sameRow _
                = (sbjRow = objRow)
            
            Dim nextRow As Boolean: nextRow _
                = (sbjRow = objRow - 1 Or sbjRow = objRow + 1)
            
            Dim nextColumn As Boolean: nextColumn _
                = (sbjColumn = objColumn - 1 Or sbjColumn = objColumn + 1)

            If (sameColumn And nextRow) Or (sameRow And nextColumn) Then
                ret(sbjIndex + 1).Connect ret(objIndex + 1)
                ret(objIndex + 1).Connect ret(sbjIndex + 1)
            End If
        Next
        
    Next
    
    Set CreateNodes = ret
End Function

Function CoulombForce(sbjPos As Axes, objPos As Axes) As Axes
    Dim distanceX As Double, distanceY As Double
    distanceX = objPos.X - sbjPos.X
    distanceY = objPos.Y - sbjPos.Y
    
    Dim distance As Double
    distance = Sqr(distanceX ^ 2 + distanceY ^ 2)
    
    Dim qubedDistance As Double
    qubedDistance = distance ^ 3
    
    Dim constant As Double
    If qubedDistance = 0# Then
        constant = 0
    Else
        constant = 0.0001 / qubedDistance
    End If
    
    Set CoulombForce = New Axes
    CoulombForce.X = -constant * distanceX
    CoulombForce.Y = -constant * distanceY
End Function

Function HookeForce(sbjPos As Axes, objPos As Axes) As Axes
    Const SPRING_CONSTANT = 1#
    Dim distanceX As Double, distanceY As Double
    distanceX = objPos.X - sbjPos.X
    distanceY = objPos.Y - sbjPos.Y
    
    Dim distance As Double, dl As Double
    distance = Sqr(distanceX ^ 2 + distanceY ^ 2)
    
    dl = distance - 0.1
    
    Dim constant As Double
    constant = SPRING_CONSTANT * dl / distance
    
    Set HookeForce = New Axes
    HookeForce.X = constant * distanceX
    HookeForce.Y = constant * distanceY
End Function


Sub ConnectOvals(ov1 As Shape, ov2 As Shape, cn As Shape)
    Dim ov1_x, ov2_x, ov1_y, ov2_y
    ov1_x = ov1.Top + ov1.Height / 2
    ov2_x = ov2.Top + ov2.Height / 2
    ov1_y = ov1.Left + ov1.Width / 2
    ov2_y = ov2.Left + ov2.Width / 2
    
    Dim X As Double, Y As Double
    X = ov2_x - ov1_x
    Y = ov2_y - ov1_y
    
    Dim Cond_Vertical As Boolean
    Dim Cond_Horizontal As Boolean
    Dim Cond_Diagonal As Boolean
    Dim Cond_Below As Boolean
    Dim Cond_Above As Boolean
    Dim Cond_Right As Boolean
    Dim Cond_Left As Boolean
    
    Cond_Vertical = X = 0
    Cond_Horizontal = Y = 0
    Cond_Below = X > 0
    Cond_Above = Not Cond_Below
    Cond_Right = Y > 0
    Cond_Left = Not Cond_Right

    If Not (Cond_Vertical Or Cond_Horizontal) Then 'To avoid devided by zero error.
        Dim degree: degree = Abs(Math.Atn(Y / X) * (180 / (4 * Atn(1))))
        Cond_Vertical = degree < 22.5
        Cond_Horizontal = degree > 67.5
    End If
    
    Cond_Diagonal = Not (Cond_Horizontal Or Cond_Vertical)
    
    Select Case True
    Case Cond_Vertical And Cond_Below
        cn.ConnectorFormat.BeginConnect ov1, 5
        cn.ConnectorFormat.EndConnect ov2, 1
    Case Cond_Vertical And Cond_Above
        cn.ConnectorFormat.BeginConnect ov1, 1
        cn.ConnectorFormat.EndConnect ov2, 5
    Case Cond_Horizontal And Cond_Right
        cn.ConnectorFormat.BeginConnect ov1, 7
        cn.ConnectorFormat.EndConnect ov2, 3
    Case Cond_Horizontal And Cond_Left
        cn.ConnectorFormat.BeginConnect ov1, 3
        cn.ConnectorFormat.EndConnect ov2, 7
    Case Cond_Diagonal And Cond_Left And Cond_Below
        cn.ConnectorFormat.BeginConnect ov1, 4
        cn.ConnectorFormat.EndConnect ov2, 8
    Case Cond_Diagonal And Cond_Left And Cond_Above
        cn.ConnectorFormat.BeginConnect ov1, 2
        cn.ConnectorFormat.EndConnect ov2, 6
    Case Cond_Diagonal And Cond_Right And Cond_Below
        cn.ConnectorFormat.BeginConnect ov1, 6
        cn.ConnectorFormat.EndConnect ov2, 2
    Case Cond_Diagonal And Cond_Right And Cond_Above
        cn.ConnectorFormat.BeginConnect ov1, 8
        cn.ConnectorFormat.EndConnect ov2, 4
    End Select
    
End Sub

