Attribute VB_Name = "MainModule"
Option Explicit
Sub Main()
    Const DELTA_T = 0.01
    Const LOSS_FACTOR = 0.99
    Const NODE_SIZE = 5
    
    Dim subjectNode As Node
    Dim objectNode As Node

    Dim Nodes As Collection
    Set Nodes = CreateNodes(5, 5)
    
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
    
    DrawSheet.ClearAllShapes

    'PlotOvals
    For Each subjectNode In Nodes
        Dim PosX As Single: PosX = subjectNode.Position.X * 500
        Dim PosY As Single: PosY = subjectNode.Position.Y * 500
        Set subjectNode.NodeShape = DrawSheet.Shapes.AddShape(msoShapeOval, PosX, PosY, NODE_SIZE, NODE_SIZE)
    Next

    'Plot Connector
    For Each subjectNode In Nodes
        For Each objectNode In subjectNode.ConnectedNode
            Dim connector As Shape
            Set connector = DrawSheet.Shapes.AddConnector(msoConnectorStraight, 1, 1, 1, 1)
            connector.ConnectorFormat.BeginConnect subjectNode.NodeShape, 1
            connector.ConnectorFormat.EndConnect objectNode.NodeShape, 1
        Next
    Next
End Sub

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


