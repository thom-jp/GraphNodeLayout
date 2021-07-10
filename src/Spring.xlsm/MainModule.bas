Attribute VB_Name = "MainModule"
Sub Main()
    Dim d As Collection: Set d = New Collection
    Dim dr As Collection
    Const nrows = 5
    Const ncols = 5
    

    Const alpha = 1#
    Const beta = 0.0001
    Const k = 1#
    
    Const eta = 0.99
    Const delta_t = 0.01

    Dim Nodes As Collection
    Set Nodes = New Collection
    For i = 0 To nrows * ncols - 1
        Nodes.Add New Node
        Nodes(Nodes.Count).ID = i
    Next
    
    For i = 0 To nrows * ncols - 1
        ci = i Mod ncols
        ri = i \ ncols
        Set dr = New Collection
        For j = 0 To nrows * ncols - 1
            cj = j Mod ncols
            rj = j \ ncols
            If ((ci = cj) And (ri = rj - 1 Or ri = rj + 1) _
                Or (ri = rj And (ci = cj - 1 Or ci = cj + 1))) Then
                Nodes(i + 1).Connect Nodes(j + 1)
                Nodes(j + 1).Connect Nodes(i + 1)
            Else
                dr.Add 0#
                'Debug.Print "0 ";
            End If
        Next
        d.Add dr
        'Debug.Print
    Next
    
    m = d.Count
    
    Dim n As Node
    Dim nn As Node
    
    Do
        Dim KineticEnergyTotal As Axes: Set KineticEnergyTotal = New Axes
        For Each n In Nodes
            Dim F As Axes: Set F = New Axes
            For Each nn In Nodes
                If n Is nn Then GoTo Continue
                If n.Connected(nn) Then
                    'Debug.Print "Connected"
                    Set fij = Coulomb_force(n.Position, nn.Position)
                Else
                    'Debug.Print "UnConnected"
                    Set fij = Hooke_force(n.Position, nn.Position, 0.1)
                End If
                F.X = F.X + fij.X
                F.Y = F.Y + fij.Y
Continue:
            Next
            n.Velocity.X = (n.Velocity.X + alpha * F.X * delta_t) * eta
            n.Velocity.Y = (n.Velocity.Y + alpha * F.Y * delta_t) * eta
            KineticEnergyTotal.X = KineticEnergyTotal.X + alpha * (n.Velocity.X ^ 2)
            KineticEnergyTotal.Y = KineticEnergyTotal.Y + alpha * (n.Velocity.Y ^ 2)
        Next
        
        Debug.Print "Total Kinetic Energy: " & Round(Sqr(KineticEnergyTotal.X ^ 2 + KineticEnergyTotal.Y ^ 2), 2)
        
        For Each n In Nodes
            n.Position.X = n.Position.X + (n.Velocity.X * delta_t)
            n.Position.Y = n.Position.Y + (n.Velocity.Y * delta_t)
        Next
    Loop Until 0.01 > Round(Sqr(KineticEnergyTotal.X ^ 2 + KineticEnergyTotal.Y ^ 2), 2)
End Sub

Function Coulomb_force(A As Axes, B As Axes) As Axes
    Const beta = 0.0001
    dx = B.X - A.X
    dy = B.Y - A.Y
    ds2 = dx ^ 2 + dy ^ 2
    ds = Sqr(ds2)
    ds3 = ds2 * ds
    If ds3 = 0# Then
        con = 0
    Else
        con = beta / (ds2 * ds)
    End If
    Set Coulomb_force = New Axes
    Coulomb_force.X = -con * dx
    Coulomb_force.Y = -con * dy
End Function

Function Hooke_force(A As Axes, B As Axes, dij) As Axes
    Const k = 1#
    dx = B.X - A.X
    dy = B.Y - A.Y
    ds = Sqr(dx ^ 2 + dy ^ 2)
    dl = ds - dij
    con = k * dl / ds
    Set Hooke_force = New Axes
    Hooke_force.X = con * dx
    Hooke_force.Y = con * dy
End Function
